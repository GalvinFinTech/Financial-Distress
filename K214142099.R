packages <- c("readxl", "tidyr", "corrr", "fixest", "lmtest", "cowplot", "plm", 
              "ggcorrplot", "imputeTS", "gridExtra", "corrplot", "GGally", 
              "xtable", "knitr", "stargazer", "shiny", "ggplot2", "plotly", 
              "dplyr", "DT", "psych", "shinythemes", "zoo","openxlsx","readxl")

# Tạo vòng lặp để nạp các gói
for (pkg in packages) {
  if (!require(pkg, character.only = TRUE)) {
    install.packages(pkg, dependencies = TRUE)
    library(pkg, character.only = TRUE)
  }
}

df <- read_excel("K214142099.xlsx")

s_score <- 1.03 * (df$`Total Current Assets` - df$`Total Current Liabilities`) / df$`Total Assets` +
  3.07 * df$`Earnings before Interest & Taxes (EBIT)` / df$`Total Assets` +
  0.66 * df$`Income before Taxes` / df$`Total Current Liabilities` +
  0.4 * df$`Revenue from Business Activities - Total` / df$`Total Assets`

risk1 <- ifelse(s_score < 0.862, 1, 0)
z_score <- 6.56 * (df$`Total Current Assets` - df$`Total Current Liabilities`) / df$`Total Assets` +
  3.26 * df$`Retained Earnings - Total` / df$`Total Assets` +
  6.72 * df$`Earnings before Interest & Taxes (EBIT)` / df$`Total Assets` +
  1.05 * df$`Shareholders' Equity - Attributable to Parent ShHold - Total` / df$`Total Liabilities`

risk2 <- ifelse(z_score < 1.1, 1, 0)
SIZE <- log(df$`Total Assets`)
ROA <- df$`Net Income after Tax` / df$`Total Assets`
CUR <- df$`Total Current Assets` / df$`Total Current Liabilities`
LEV <- df$`Total Liabilities` / df$`Total Assets`
FAR <- df$`Total Fixed Assets - Net` / df$`Total Assets`

# Create a new dataframe
df_processed <- data.frame(
  code = df$code, 
  date = df$date, 
  s_score = s_score, 
  z_score = z_score, 
  SIZE = SIZE, 
  ROA = ROA, 
  CUR = CUR, 
  LEV = LEV, 
  FAR = FAR, 
  risk1 = risk1, 
  risk2 = risk2
)

colSums(is.na(df_processed))

# Loại bỏ các dòng chứa giá trị NaN từ df1
df_processed <- df_processed[complete.cases(df_processed), ]


# Hàm xử lý giá trị ngoại lai bằng IQR
remove_outliers <- function(x) {
  # Kiểm tra xem có giá trị NA không
  has_na <- any(is.na(x))
  
  # Tính quantile với na.rm = FALSE nếu có giá trị NA
  if (has_na) {
    Q1 <- quantile(x, 0.25, na.rm = FALSE)
    Q3 <- quantile(x, 0.75, na.rm = FALSE)
  } else {
    Q1 <- quantile(x, 0.25, na.rm = TRUE)
    Q3 <- quantile(x, 0.75, na.rm = TRUE)
  }
  
  IQR <- Q3 - Q1
  x[x < (Q1 - 1.5 * IQR) | x > (Q3 + 1.5 * IQR)] <- NA
  return(x)
}

# Xử lý giá trị ngoại lai cho df1
df_processed <- df_processed %>%
  mutate(
    CUR = remove_outliers(CUR),
    s_score = remove_outliers(s_score),
    z_score = remove_outliers(z_score) )


#PHÂN TÍCH THỔNG KẾ MÔ TẢ
summary(df_processed)
describe(df_processed)


# Chạy describe(df_1) và gán kết quả vào. một biến, Sử dụng kable để chuyển đổi kết quả thành LaTeX
df_1_desc <- describe(df_processed[,c('SIZE', 'ROA', 'CUR', 'LEV', 'FAR', 's_score', 'z_score')])
kable(df_1_desc)
kable(df_1_desc, format = "latex")

cor_matrix <- cor(df_processed[, c('SIZE', 'ROA', 'CUR', 'LEV', 'FAR', 's_score', 'z_score')], use = "complete.obs")
corrplot(cor_matrix, method = "circle", tl.cex = 0.8, number.cex = 0.8, addCoef.col = "black")

# Tạo bảng từ ma trận tương quan, # In bảng LaTeX
cor_table <- xtable(cor_matrix)
print(cor_table, include.rownames = TRUE)

# CHẠY MÔ HÌNH HỒI QUY 
# Kiểm tra loại dữ liệu của cột date
str(df_processed$date)
df_processed <- df_processed %>% mutate(date = as.integer(date))
# Hàm để chạy mô hình hồi quy và trả về kết quả
run_regression <- function(data, response_var, date_range, outcome_name) {
  df <- data %>% filter(date >= date_range[1] & date <= date_range[2])
  df <- df[!duplicated(df[c("code", "date")]), ]
  pdata <- plm::pdata.frame(df, index = c("code", "date"))
  model <- plm(as.formula(paste(response_var, "~ SIZE + ROA + CUR + LEV + FAR")), data = pdata, model = "within")
  cat("Summary of model for", outcome_name, ":\n")
  print(summary(model))
  
  # Tạo biểu đồ histogram của phần dư
  residuals <- residuals(model)
  hist(residuals, breaks = 30, main = paste("Residuals Histogram for", outcome_name), xlab = "Residuals", freq = FALSE)
  # Tính toán giá trị cho đường phân phối chuẩn
  x <- seq(min(residuals), max(residuals), length = 100)
  y <- dnorm(x, mean = mean(residuals), sd = sd(residuals))
  # Thêm đường phân phối chuẩn vào biểu đồ
  lines(x, y, col = "blue", lwd = 2)
  
  return(model)  # Trả về mô hình hồi quy
}

# Chạy mô hình hồi quy và gán kết quả vào các biến model1 đến model6
model1 <- run_regression(df_processed, "s_score", c(2014, 2019), "Pre-COVID (2014-2019)")
model2 <- run_regression(df_processed, "s_score", c(2020, 2021), "During COVID (2020-2021)")
model3 <- run_regression(df_processed, "s_score", c(2022, 2023), "Post-COVID (2022-2023)")
model4 <- run_regression(df_processed, "z_score", c(2014, 2019), "Pre-COVID (2014-2019)")
model5 <- run_regression(df_processed, "z_score", c(2020, 2021), "During COVID (2020-2021)")
model6 <- run_regression(df_processed, "z_score", c(2022, 2023), "Post-COVID (2022-2023)")

# Combine all models in one LaTeX table
stargazer(model1, model2, model3, model4, model5, model6,
          type = "latex",
          title = "Regression Results",
          dep.var.labels = c("s_score", "s_score", "s_score", "z_score", "z_score", "z_score"),
          covariate.labels = c("SIZE", "ROA", "CUR", "LEV", "FAR"),
          column.labels = c("Pre-COVID (2014-2019)", "During COVID (2020-2021)", "Post-COVID (2022-2023)",
                            "Pre-COVID (2014-2019)", "During COVID (2020-2021)", "Post-COVID (2022-2023)"),
          model.names = FALSE,
          out = "combined_regression_results.tex")

library(shiny)
library(DT)
library(viridis)
library(RColorBrewer)
# Define continuous variables
continuous_vars <- c("z_score", "s_score", "ROA", "CUR", "LEV", "FAR", "SIZE")
# Define colors for each variable
variable_colors <- c("z_score" = "blue", "s_score" = "green", "ROA" = "red", 
                     "CUR" = "purple", "LEV" = "orange", "FAR" = "pink", "SIZE" = "brown")
ui <- tagList(
  shinythemes::themeSelector(),
  navbarPage(
    "NHÂN TỐ TÁC ĐỘNG ĐẾN KIỆT QUỆ TÀI CHÍNH",
    
    tabPanel("TRỰC QUAN DỮ LIỆU",
             fluidRow(
               column(6, plotOutput(outputId = "density_z_score")),
               column(6, plotOutput(outputId = "density_s_score"))
             ),
             fluidRow(
               column(4,
                      lapply(continuous_vars, function(var) {
                        plotOutput(outputId = paste0("hist_", var))
                      })
               ),
               column(4,
                      lapply(continuous_vars, function(var) {
                        plotOutput(outputId = paste0("box_", var))
                      })
               ),
               column(4,
                      lapply(continuous_vars, function(var) {
                        plotOutput(outputId = paste0("violin_", var))
                      })
               )
             ),
             fluidRow(
               column(12, plotOutput("pairplot"))
             )
    ),
    tabPanel("KẾT QUẢ MÔ HÌNH",
             sidebarPanel(
               selectInput("model_selection", "Chọn mô hình:",
                           choices = c("Mô hình trước Covid", "Mô hình trong Covid", "Mô hình sau Covid"))
             ),
             mainPanel(
               fluidRow(
                 column(12,
                        h3("Kết quả hồi quy"),
                        tableOutput("regression_table")
                 )
               ),
               fluidRow(
                 column(12,
                        h3("Thống kê mô tả"),
                        tableOutput("summary_table")
                 )
               ),
               fluidRow(
                 column(6,
                        h3("Dự đoán kiệt quệ tài chính"),
                        plotOutput(outputId = "risk1_plot")
                 ),
                 column(6,
                        h3("Ma trận tương quan"),
                        plotOutput(outputId = "correlation_plot")
                 )
               )
             )
    ),
    tabPanel("PHÂN TÍCH DOANH NGHIỆP",
             sidebarLayout(
               sidebarPanel(
                 tabsetPanel(
                   id = "analysis_tabs",
                   tabPanel("Kiểm tra",
                            selectInput("stock_code", "Chọn cổ phiếu:", choices = unique(df_processed$code)),
                            selectInput("year", "Chọn năm:", choices = unique(df_processed$date)),
                            actionButton("submit_button", "Kết quả"),
                            textOutput("result_text")
                   ),
                   tabPanel("Top Chỉ số",
                            selectInput("metric", "Chọn chỉ số:", choices = c("ROA", "CUR", "LEV", "FAR", "SIZE")),
                            selectInput("top_n", "Chọn số lượng doanh nghiệp:", choices = 1:20, selected = 5)
                   ),
                   tabPanel("So sánh",
                            selectInput("compare_codes", "Chọn các cổ phiếu để so sánh:", choices = unique(df_processed$code), multiple = TRUE),
                            selectInput("compare_metric", "Chọn chỉ số để so sánh:", choices = c("ROA", "CUR", "LEV", "FAR", "SIZE"))
                   )
                 )
               ),
               mainPanel(
                 tabsetPanel(
                   tabPanel("Các chỉ số", plotlyOutput("chi_so")),
                   tabPanel("Top Chỉ số", plotlyOutput("top_metric_plot")),
                   tabPanel("So sánh", plotlyOutput("comparison_plot"))
                 )
               )
             )
    ),
    tabPanel("DỮ LIỆU TÀI CHÍNH",
             sidebarPanel(
               # Thêm các điều khiển để chọn dữ liệu
               selectInput("financial_data", "Chọn dữ liệu:",
                           choices = c("Dữ liệu hồi quy trước Covid", "Dữ liệu hồi quy trong Covid", "Dữ liệu hồi quy sau Covid"),
                           
               ),
               mainPanel(
                 # Thêm đối tượng để hiển thị dữ liệu
                 DTOutput("selected_data_table"),
                 downloadButton("download_data", "Tải dữ liệu về")
                 
               )
             )
    )
  )
)


# Define server function  
server <- function(input, output) {
  
  output$pairplot <- renderPlot({
    data_to_corr <- df_processed[, c("SIZE", "ROA", "CUR", "LEV", "FAR", "s_score", "z_score")]
    ggpairs(data_to_corr)
  })
  
  # Render risk1_plot
  output$risk1_plot <- renderPlot({
    # Tạo biểu đồ pie chart từ biến phân loại (risk2)
    ggplot(df_processed, aes(x = "", fill = factor(risk2, levels = c(0, 1)))) +
      geom_bar(width = 1) +  # Đặt width = 1 để tạo ra biểu đồ pie chart
      coord_polar(theta = "y") +  # Chuyển đổi thành biểu đồ pie chart
      labs(fill = "Risk", title = "Phân loại doanh nghiệp kiệt quệ tài chính") +
      scale_fill_manual(values = c("red", "blue")) +  # Tuỳ chỉnh màu sắc
      theme_void()  # Loại bỏ background và các yếu tố thêm
  })
  
  # Render correlation_plot
  output$correlation_plot <- renderPlot({
    cor_matrix <- cor(df_processed[, c('SIZE', 'ROA', 'CUR', 'LEV', 'FAR', 's_score','z_score')], use = "complete.obs")
    corrplot(cor_matrix, method = "circle", tl.cex = 0.8, number.cex = 0.8, addCoef.col = "black")
  })
  
  # Render summary_table
  output$summary_table <- renderTable({
    describe(df_processed[, c('SIZE', 'ROA', 'CUR', 'LEV', 'FAR', 's_score','z_score')])
  })
  
  output$regression_table <- renderTable({
    model_selection <- switch(input$model_selection,
                              "Mô hình trước Covid" = plm(s_score ~ SIZE +  ROA + CUR + LEV + FAR, data = df_processed[df_processed$date < 2020, ], model = "within"),
                              "Mô hình trong Covid" = plm(s_score ~ SIZE +  ROA + CUR + LEV + FAR, data = df_processed[df_processed$date >= 2020 & df_processed$date <= 2021, ], model = "within"),                         
                              "Mô hình sau Covid" = plm(s_score ~ SIZE +  ROA + CUR + LEV + FAR, data = df_processed[df_processed$date > 2021, ], model = "within"))
    
    model_summary <- summary(model_selection)
    model_table <- as.data.frame(model_summary$coefficients)
    
    # Print the names of the columns
    print(colnames(model_table))
    
    # Print the first few rows of the model_table to check the structure
    print(head(model_table))
    
    # Add a column for independent variables
    model_table <- tibble::rownames_to_column(model_table, "Biến độc lập")
    
    # Reorder columns (assuming these are the correct column names)
    colnames(model_table) <- c("Biến độc lập", "Hệ số", "Std. Error", "t-value", "Pr(>|t|)")
    model_table <- model_table[, c("Biến độc lập", "Hệ số", "Std. Error", "t-value", "Pr(>|t|)")]
    
    model_table
  })
  
  
  observeEvent(input$submit_button, {
    selected_stock <- input$stock_code
    selected_year <- as.numeric(input$year)
    
    # Lấy dữ liệu của doanh nghiệp được chọn
    selected_row <- df_processed[df_processed$code == selected_stock & df_processed$date == selected_year, ]
    stock_data <- df_processed[df_processed$code == selected_stock, ]
    
    # Tính toán tỉ lệ giữa ROA, CUR, LEV, FAR, SIZE
    if (nrow(stock_data) > 0) {
      # Xử lý log(Size)
      stock_data$SIZE <- exp(stock_data$SIZE)
      
      # Tạo biểu đồ
      p <- plot_ly(stock_data, x = ~date)
      for (metric in c("ROA", "CUR", "LEV", "FAR", "SIZE")) {
        p <- p %>%
          add_bars(y = as.formula(paste0("~", metric)), name = metric) %>%
          add_markers(y = as.formula(paste0("~", metric)), name = paste0(metric, " Points"), marker = list(color = 'black')) %>%
          add_lines(y = as.formula(paste0("~", metric)), name = paste0(metric, " Line"), line = list(color = brewer.pal(8, "Set2")[5])) 
      }
      
      p <- p %>%
        layout(title = paste("Các chỉ số của", selected_stock), xaxis = list(title = "Year"))
      
      output$chi_so <- renderPlotly({
        p
      })
    } else {
      output$result_text <- renderText("Không tìm thấy dữ liệu cho doanh nghiệp được chọn.")
    }
  })
  
  # Khai báo một reactive chứa dữ liệu đã chọn
  selected_data <- reactive({
    switch(input$financial_data,
           "Dữ liệu hồi quy trước Covid" = df_processed[df_processed$date < 2020, ],
           "Dữ liệu hồi quy trong Covid" = df_processed[df_processed$date >= 2020 & df_processed$date <= 2021, ],
           "Dữ liệu hồi quy sau Covid" = df_processed[df_processed$date > 2021, ]
    )
  })
  
  # Hiển thị dữ liệu đã chọn
  output$selected_data_table <- renderDT({
    datatable(selected_data())
  })
  
  output$download_data <- downloadHandler(
    filename = function() {
      paste("financial_data", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      if (input$financial_data == "Dữ liệu hồi quy trước Covid") {
        write.csv(df_processed[df_processed$date < 2020, ], file, row.names = FALSE)
      } else if (input$financial_data == "Dữ liệu hồi quy trong Covid") {
        write.csv(df_processed[df_processed$date >= 2020 & df_processed$date <= 2021, ], file, row.names = FALSE)
      } else if (input$financial_data == "Dữ liệu hồi quy sau Covid") {
        write.csv(df_processed[df_processed$date > 2021, ], file, row.names = FALSE)
      }
    })
  
  top_n <- function(df, var, n = 5) {
    df %>%
      group_by(code) %>%
      summarise(mean_val = mean({{var}}, na.rm = TRUE)) %>%
      arrange(desc(mean_val)) %>%
      slice_head(n = n)
  }
  
  output$top_metric_plot <- renderPlotly({
    metric <- input$metric
    top_n_value <- as.numeric(input$top_n)
    
    if (metric == "SIZE") {
      df_processed$SIZE_actual <- exp(df_processed$SIZE)
      top_data <- top_n(df_processed, SIZE_actual, top_n_value)
    } else {
      top_data <- top_n(df_processed, !!sym(metric), top_n_value)
    }
    
    colors <- colorRampPalette(brewer.pal(12, "Paired"))(top_n_value)
    
    p <- ggplot(top_data, aes(x = reorder(code, -mean_val), y = mean_val, text = paste0(metric, ": ", mean_val))) +
      geom_col(fill = colors) +
      geom_point(aes(label = code), color = "red") +
      geom_line(aes(group = 1), color = "red") +
      labs(title = paste("Top", top_n_value, "Doanh nghiệp có", metric, "cao nhất"),
           x = "Doanh nghiệp", y = metric) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, hjust = 1))  # Điều chỉnh góc và vị trí của nhãn trên trục x
    
    ggplotly(p, tooltip = "text")
  })
  
  observe({
    compare_codes <- input$compare_codes
    compare_metric <- input$compare_metric
    
    compare_data <- df_processed %>% filter(code %in% compare_codes)
    
    if (compare_metric == "SIZE") {
      compare_data <- compare_data %>% mutate(SIZE_actual = exp(SIZE))
      metric_col <- "SIZE_actual"
    } else {
      metric_col <- compare_metric
    }
    
    p <- ggplot(compare_data, aes_string(x = "date", y = metric_col, color = "code")) +
      geom_line() +
      geom_point(aes(text = code)) +
      labs(title = paste("So sánh", compare_metric, "giữa các doanh nghiệp"),
           x = "date", y = compare_metric, color = "Doanh nghiệp", text = "code") +
      theme_minimal()
    
    output$comparison_plot <- renderPlotly({
      ggplotly(p, tooltip = "text", verbose = FALSE)
    })
  })
  
  
  # Density plots for z_score and s_score
  output$density_z_score <- renderPlot({
    ggplot(df_processed, aes(x = z_score)) +
      geom_density(fill = variable_colors["z_score"], alpha = 0.5) +
      theme_minimal()
  })
  
  output$density_s_score <- renderPlot({
    ggplot(df_processed, aes(x = s_score)) +
      geom_density(fill = variable_colors["s_score"], alpha = 0.5) +
      theme_minimal()
  })
  
  # Histogram, Box Plot, and Violin Plot for each continuous variable
  for (var in continuous_vars) {
    local({
      var <- var
      output[[paste0("hist_", var)]] <- renderPlot({
        ggplot(df_processed, aes_string(x = var)) +
          geom_histogram(binwidth = 0.5, fill = variable_colors[[var]], color = "black") +
          labs(title = paste("Distribution of", var), x = var, y = "Frequency") +
          theme_minimal()
      })
      
      output[[paste0("box_", var)]] <- renderPlot({
        ggplot(df_processed, aes_string(y = var)) +
          geom_boxplot(fill = variable_colors[[var]], color = "black") +
          labs(title = paste("Box Plot of", var), x = "", y = var) +
          theme_minimal()
      })
      
      output[[paste0("violin_", var)]] <- renderPlot({
        ggplot(df_processed, aes(x = factor(1), y = !!sym(var))) +
          geom_violin(fill = variable_colors[[var]], color = "black") +
          labs(title = paste("Violin Plot of", var), x = "", y = var) +
          theme_minimal() +
          theme(axis.text.x = element_blank(), axis.ticks.x = element_blank())
      })
    })
  }
}

# Create Shiny app 
shinyApp(ui = ui, server = server)







