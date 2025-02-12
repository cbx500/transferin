library(shiny)
library(readxl)
library(openxlsx)
library(dplyr)
library(shinyjs)

server <- function(input, output, session) {
  
  shinyjs::useShinyjs()
  
  # Helper functions for European formatting (no scientific notation)
  format_number <- function(x) {
    if (is.numeric(x)) {
      format(round(x, 2), big.mark = ".", decimal.mark = ",", scientific = FALSE)
    } else {
      x
    }
  }
  
  format_date <- function(x) {
    dt <- as.Date(x)
    if (!is.na(dt)) {
      format(dt, "%d-%m-%Y")
    } else {
      x
    }
  }
  
  base_path <- getwd()
  data_path <- file.path(base_path, "data")
  output_path <- file.path(base_path, "output")
  
  # File paths for input files (always relative for deployment)
  personal_file <- file.path(data_path, "Personal_Data.xlsx")
  input_file <- file.path(data_path, "Input_Data.xlsx")
  static_file <- file.path(data_path, "Static_Data.xlsx")
  
  # Files used for processing and outputs
  results_file <- file.path(output_path, "TransferInCreditResults.xlsx")
  extracted_personal_file <- file.path(output_path, "Extracted_Parameters_TransferIn.xlsx")
  
  local_results_file <- results_file
  local_extracted_personal_file <- extracted_personal_file
  
  # ReactivePoll for Transfer-In Credit Results
  load_results <- reactivePoll(
    intervalMillis = 2000,
    session = session,
    checkFunc = function() {
      if (file.exists(local_results_file)) file.info(local_results_file)$mtime else Sys.time()
    },
    valueFunc = function() {
      cat("Reading results file at:", Sys.time(), "\n")
      flush.console()
      read_excel(local_results_file, sheet = "Results") %>%
        rename(Parameter = 1, Value = 2) %>%
        filter(!is.na(Parameter))
    }
  )
  
  # ReactivePoll for Extracted Personal Data
  load_personal <- reactivePoll(
    intervalMillis = 2000,
    session = session,
    checkFunc = function() {
      if (file.exists(local_extracted_personal_file)) file.info(local_extracted_personal_file)$mtime else Sys.time()
    },
    valueFunc = function() {
      cat("Reading personal data file at:", Sys.time(), "\n")
      flush.console()
      read_excel(local_extracted_personal_file, sheet = 1) %>%
        rename(Parameter = 1, Value = 2) %>%
        filter(!is.na(Parameter))
    }
  )
  
  # --- File Upload Observers ---
  observeEvent(input$upload_personal, {
    req(input$upload_personal)
    file.copy(input$upload_personal$datapath, personal_file, overwrite = TRUE)
    cat("Uploaded Personal Data file updated.\n")
  })
  
  observeEvent(input$upload_input, {
    req(input$upload_input)
    file.copy(input$upload_input$datapath, input_file, overwrite = TRUE)
    cat("Uploaded Input Data file updated.\n")
  })
  
  observeEvent(input$upload_static, {
    req(input$upload_static)
    file.copy(input$upload_static$datapath, static_file, overwrite = TRUE)
    cat("Uploaded Static Data file updated.\n")
  })
  
  # Enable run_calculation button only if both required uploads are available
  observe({
    if (!is.null(input$upload_personal) && !is.null(input$upload_input)) {
      shinyjs::enable("run_calculation")
    } else {
      shinyjs::disable("run_calculation")
    }
  })
  
  # Run Pipeline Script when "Run Transfer-In Calculation" is clicked
  observeEvent(input$run_calculation, {
    req(input$upload_personal, input$upload_input)
    
    script_path <- file.path(base_path, "scripts", "MasterPipeline.R")
    if (!file.exists(script_path)) {
      showNotification("Error: MasterPipeline.R not found!", type = "error")
      cat("❌ ERROR: MasterPipeline.R NOT found at:", script_path, "\n")
      return()
    }
    
    cat("✅ Running Transfer-In Calculation Pipeline...\n")
    system(paste("Rscript", shQuote(script_path)), wait = TRUE)
    showNotification("Pipeline script executed. Refreshing results...", type = "message")
    # Force a refresh by reading the files again
    load_results()
    load_personal()
  })
  
  # --- UI Outputs ---
  
  ## Personal Information
  output$staff_name <- renderText({
    val <- load_personal() %>% filter(Parameter == "Staff_Name") %>% pull(Value)
    paste("Staff Name:", val)
  })
  output$staff_dob <- renderText({
    val <- load_personal() %>% filter(Parameter == "DOB_Staff") %>% pull(Value)
    paste("Staff DOB:", format_date(val))
  })
  output$spouse_dob <- renderText({
    val <- load_personal() %>% filter(Parameter == "DOB_Spouse") %>% pull(Value)
    paste("Spouse DOB:", format_date(val))
  })
  output$retirement_date <- renderText({
    val <- load_personal() %>% filter(Parameter == "retirement_date") %>% pull(Value)
    paste("Retirement Date:", format_date(val))
  })
  output$report_date <- renderText({
    val <- load_personal() %>% filter(Parameter == "report_date") %>% pull(Value)
    paste("Report Date:", format_date(val))
  })
  output$transfer_in_amount <- renderText({
    val <- as.numeric(load_personal() %>% filter(Parameter == "transferred_value") %>% pull(Value))
    paste("Transfer-In Amount:", format_number(val))
  })
  
  ## Revalued Earnings Information
  # Annual value from the results file (renamed for display)
  output$annual_revalued_earnings_hr <- renderText({
    val <- as.numeric(load_results() %>% filter(Parameter == "Annual Revalued Earnings HR") %>% pull(Value))
    paste("Annual Revalued Earnings HR discounted at report date:", format_number(val))
  })
  # New monthly value: annual / 12
  output$monthly_revalued_earnings_hr <- renderText({
    annual_val <- as.numeric(load_results() %>% filter(Parameter == "Annual Revalued Earnings HR") %>% pull(Value))
    monthly_val <- annual_val / 12
    paste("Monthly revalued earnings HR discounted at report date:", format_number(monthly_val))
  })
  # Annual capitalised value (renamed)
  output$annual_revalued_earnings_hr_capitalised <- renderText({
    val <- as.numeric(load_results() %>% filter(Parameter == "Annual Revalued Earnings HR Capitalised") %>% pull(Value))
    paste("Annual Revalued Earnings HR at retirement age:", format_number(val))
  })
  # Calculated annual revalued earnings discounted at report date
  output$revalued_earnings_discounted <- renderText({
    val <- as.numeric(load_results() %>% filter(Parameter == "Revalued Earnings Discounted to Report Date") %>% pull(Value))
    paste("Calculated annual Revalued earnings discounted at report date:", format_number(val))
  })
  # Calculated annual revalued earnings at retirement age
  output$revalued_earnings_retirement_age65 <- renderText({
    val <- as.numeric(load_results() %>% filter(Parameter == "Revalued Earnings at Retirement Age 65") %>% pull(Value))
    paste("Calculated annual Revalued earnings at retirement age:", format_number(val))
  })
  
  ## Transfer-In Credit Results
  output$transfer_in_credit_calculated <- renderText({
    val <- as.numeric(load_results() %>% filter(Parameter == "Transfer-In Credit") %>% pull(Value))
    paste("Transfer-In Credit with Calculated RE:", format_number(val))
  })
  output$transfer_in_credit_provided <- renderText({
    val <- as.numeric(load_results() %>% filter(Parameter == "Transfer-In Credit HRRE") %>% pull(Value))
    paste("Transfer-In Credit with Provided RE:", format_number(val))
  })
  output$transfer_in_factor <- renderText({
    val <- load_results() %>% filter(Parameter == "Scheme Factor") %>% pull(Value)
    paste("Transfer-In Factor:", format_number(as.numeric(val)))
  })
  output$pcf_total <- renderText({
    val <- as.numeric(load_results() %>% filter(Parameter == "PCF Total") %>% pull(Value))
    paste("PCF Total:", format_number(val))
  })
  output$pcf_staff <- renderText({
    val <- as.numeric(load_results() %>% filter(Parameter == "PCF Staff") %>% pull(Value))
    paste("PCF Staff:", format_number(val))
  })
  output$pcf_spouse <- renderText({
    val <- as.numeric(load_results() %>% filter(Parameter == "PCF Spouse") %>% pull(Value))
    paste("PCF Spouse:", format_number(val))
  })
  output$time_to_retirement <- renderText({
    val <- load_results() %>% filter(Parameter == "Time to Retirement (TTR)") %>% pull(Value)
    paste("Time to Retirement (Years):", val)
  })
  output$discount_rate_nby <- renderText({
    val <- as.numeric(load_personal() %>% filter(Parameter == "Discount_Rate_NBY") %>% pull(Value))
    paste("Discount Rate NBY:", format_number(val * 100), "%", sep = " ")
  })
  output$preretirement_discountrate <- renderText({
    val <- as.numeric(load_personal() %>% filter(Parameter == "preretirement_discountrate") %>% pull(Value))
    paste("Pre-retirement Discount Rate:", format_number(val * 100), "%", sep = " ")
  })
  
  # --- Download Handlers (always active) ---
  output$download_personal <- downloadHandler(
    filename = "Personal_Data.xlsx",
    content = function(file) {
      if (file.exists(personal_file)) {
        file.copy(personal_file, file)
      } else {
        write.xlsx(data.frame(Message = "Personal Data file not available"), file)
      }
    }
  )
  
  output$download_input <- downloadHandler(
    filename = "Input_Data.xlsx",
    content = function(file) {
      if (file.exists(input_file)) {
        file.copy(input_file, file)
      } else {
        write.xlsx(data.frame(Message = "Input Data file not available"), file)
      }
    }
  )
  
  output$download_static <- downloadHandler(
    filename = "Static_Data.xlsx",
    content = function(file) {
      if (file.exists(static_file)) {
        file.copy(static_file, file)
      } else {
        write.xlsx(data.frame(Message = "Static Data file not available"), file)
      }
    }
  )
  
  output$download_transfer_info <- downloadHandler(
    filename = "TransferIn_Information.pdf",
    content = function(file) {
      info_file <- file.path(data_path, "transferin_information.pdf")
      if (file.exists(info_file)) {
        file.copy(info_file, file)
      } else {
        writeLines("Transfer-In Information file not available", con = file)
      }
    }
  )
  
  output$download_all_outputs <- downloadHandler(
    filename = "All_Output_Files.zip",
    content = function(file) {
      files_to_zip <- list.files(output_path, full.names = TRUE)
      zip(file, files = files_to_zip)
    }
  )
}
