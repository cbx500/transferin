# ------------------------------------------------------------------------------
# Title: Pension Calculation Script - Step 1: Data Loading and Preprocessing
# Description: 
#   Loads and processes:
#     - Personal Data
#     - Salary History
#     - Early Retirement Factors
#     - Salary Increases
#     - Revaluation Factors
#   and exports the consolidated data to an Excel file for verification.
# ------------------------------------------------------------------------------

# Step 1: Load Required Libraries
library(readxl)      # For reading Excel files
library(dplyr)       # For data manipulation
library(yaml)        # For reading YAML configuration files
library(lubridate)   # For date handling
library(openxlsx)    # For writing Excel output files
library(janitor)     # For cleaning column names

cat("Loading Configuration...\n")
# Load the unified YAML configuration file
config <- yaml.load_file("config/config.yaml")  # Use relative path to YAML file

# Construct paths using the unified YAML; note that we use config$root_path (not project_root)
personal_data_path <- file.path(config$root_path, config$personal_data_path)
static_data_path   <- file.path(config$root_path, config$static_data_path)
output_dir         <- file.path(config$root_path, config$output_dir)
personal_data_sheets <- config$personal_data_sheets
static_data_sheets   <- config$static_data_sheets

cat("Configuration loaded successfully!\n")
cat("Personal Data Path:", personal_data_path, "\n")
cat("Static Data Path:", static_data_path, "\n")
cat("Output Directory:", output_dir, "\n\n")

# ------------------------------------------------------------------------------
# Step 1.2: Load Static Data and Extract Scheme Start Date
# ------------------------------------------------------------------------------
cat("Loading Static Data...\n")
static_data <- read_excel(static_data_path, 
                          sheet = static_data_sheets$static_data, 
                          col_types = "text")

# Convert scheme_start_date to a proper date format
static_data <- static_data %>%
  mutate(
    Value = case_when(
      Parameter == "scheme_start_date" & grepl("^[0-9]+$", Value) ~ 
        suppressWarnings(as.character(as.Date(as.numeric(Value), origin = "1899-12-30"))),
      Parameter == "scheme_start_date" & grepl("^\\d{4}-\\d{2}-\\d{2}$", Value) ~ 
        suppressWarnings(as.character(ymd(Value))),
      TRUE ~ Value  # Keep other values as-is
    )
  )

scheme_start_date <- static_data %>% filter(Parameter == "scheme_start_date") %>% pull(Value)
inflation_rate <- as.numeric(static_data %>% filter(Parameter == "inflation_rate") %>% pull(Value))
discount_rate  <- as.numeric(static_data %>% filter(Parameter == "discount_rate") %>% pull(Value))
accrual_rate   <- as.numeric(static_data %>% filter(Parameter == "accrual_rate") %>% pull(Value))
preretirement_discountrate <- as.numeric(static_data %>% filter(Parameter == "preretirement_discountrate") %>% pull(Value))

cat("Static Data Loaded Successfully:\n")
cat("Scheme Start Date:", scheme_start_date, "\n")
cat("Inflation Rate:", inflation_rate, "\n")
cat("Discount Rate:", discount_rate, "\n")
cat("Accrual Rate:", accrual_rate, "\n")
cat("Pre-Retirement Discount Rate:", preretirement_discountrate, "\n\n")


cat("Static Data Loaded Successfully:\n")
cat("Scheme Start Date:", scheme_start_date, "\n")
cat("Inflation Rate:", inflation_rate, "\n")
cat("Discount Rate:", discount_rate, "\n")
cat("Accrual Rate:", accrual_rate, "\n\n")

# ------------------------------------------------------------------------------
# Step 1.3: Load Personal Data and Convert Relevant Fields
# ------------------------------------------------------------------------------
cat("Loading Personal Data...\n")
personal_data_file <- file.path(personal_data_path, "Personal_Data.xlsx")

personal_data <- read_excel(personal_data_file, 
                            sheet = personal_data_sheets$personal_data, 
                            col_types = "text") %>%
  select(Field, Value)  # Keep only necessary columns

# Convert date fields in Personal Data
personal_data <- personal_data %>%
  mutate(
    Value = case_when(
      Field %in% c("DOB", "ecb_joining_date", "retirement_date", "spouse_DOB", "report_date") & 
        grepl("^[0-9]+$", Value) ~ suppressWarnings(as.character(as.Date(as.numeric(Value), origin = "1899-12-30"))),
      Field %in% c("DOB", "ecb_joining_date", "retirement_date", "spouse_DOB", "report_date") & 
        grepl("^\\d{4}-\\d{2}-\\d{2}$", Value) ~ suppressWarnings(as.character(ymd(Value))),
      TRUE ~ Value  # Keep non-date values as is
    )
  )

cat("Personal Data Loaded Successfully:\n")
print(personal_data)

# ------------------------------------------------------------------------------
# Step 1.4: Load and Process Salary History
# ------------------------------------------------------------------------------
cat("Loading and Formatting Salary History...\n")
salary_history <- read_excel(personal_data_file, 
                             sheet = personal_data_sheets$salary_history, 
                             col_types = "text") %>%
  mutate(
    Date = as.Date(as.numeric(Date), origin = "1899-12-30"),
    Value = as.numeric(gsub(",", ".", Value)),
    worked = as.numeric(gsub(",", ".", worked))
  ) %>%
  filter(!is.na(worked))  # Capture all non-empty 'worked' values

num_rows_captured <- nrow(salary_history)
cat("Number of rows captured in the 'worked' column (non-empty):", num_rows_captured, "\n")

worked_sum <- sum(salary_history$worked, na.rm = TRUE)

salary_history <- bind_rows(
  salary_history,
  tibble(
    Field = "Total",
    Date = NA,
    Value = NA,
    worked = worked_sum,
    Worked_Salary = NA
  )
)

cat("Salary History Loaded Successfully:\n")
print(salary_history)

# ------------------------------------------------------------------------------
# Step 1.5: Load and Process Early Retirement Factors
# ------------------------------------------------------------------------------
cat("Loading Early Retirement Factors...\n")
early_retirement_factors <- read_excel(static_data_path, 
                                       sheet = static_data_sheets$early_retirement_factors, 
                                       col_types = "text") %>%
  mutate(
    Age = as.numeric(Age),
    Scheme_Prior_2009 = as.numeric(gsub(",", ".", Scheme_Prior_2009)),
    Scheme_Post_2009 = as.numeric(gsub(",", ".", Scheme_Post_2009))
  )

cat("Early Retirement Factors Loaded Successfully:\n")
print(early_retirement_factors)

# ------------------------------------------------------------------------------
# Step 1.6: Load Salary Increases
# ------------------------------------------------------------------------------
cat("Loading Salary Increases...\n")
salary_increases <- read_excel(static_data_path, 
                               sheet = static_data_sheets$salary_increases, 
                               col_types = "text") %>%
  clean_names() %>%
  mutate(
    from_age = as.numeric(from_age),
    increase_percentage = as.numeric(increase_percentage)
  )

cat("Salary Increases Loaded Successfully:\n")
print(salary_increases)

# ------------------------------------------------------------------------------
# Step 1.7: Load Revaluation Factors
# ------------------------------------------------------------------------------
cat("Loading Revaluation Factors...\n")
revaluation_factors <- read_excel(static_data_path, 
                                  sheet = static_data_sheets$revaluation_factors, 
                                  col_types = "text") %>%
  mutate(
    Year = as.numeric(Year),
    Revaluation_Factor = as.numeric(gsub(",", ".", Revaluation_Factor))
  )

cat("Revaluation Factors Loaded Successfully:\n")
print(revaluation_factors)

# ------------------------------------------------------------------------------
# Step 1.8: Export All Data to Excel with Proper Formatting
# ------------------------------------------------------------------------------
cat("Exporting All Data to Excel...\n")
wb <- createWorkbook()

# Function to apply date format to specific columns
apply_date_format <- function(wb, sheet, col_range, data) {
  addStyle(wb, sheet = sheet, style = createStyle(numFmt = "mm/dd/yyyy"), 
           cols = col_range, rows = 2:(nrow(data) + 1), gridExpand = TRUE)
}

# Function to apply numeric format to specific columns
apply_numeric_format <- function(wb, sheet, col_range, data) {
  addStyle(wb, sheet = sheet, style = createStyle(numFmt = "0.00"), 
           cols = col_range, rows = 2:(nrow(data) + 1), gridExpand = TRUE)
}

# Write Static Data
addWorksheet(wb, "Static Data")
writeData(wb, "Static Data", static_data)
apply_date_format(wb, "Static Data", which(static_data$Parameter == "scheme_start_date"), static_data)

# Write Personal Data
addWorksheet(wb, "Personal Data")
writeData(wb, "Personal Data", personal_data)
apply_date_format(wb, "Personal Data", which(personal_data$Field %in% c("DOB", "ecb_joining_date", "retirement_date", "spouse_DOB", "report_date")), personal_data)

# Write Salary History
addWorksheet(wb, "Salary History")
writeData(wb, "Salary History", salary_history)
apply_date_format(wb, "Salary History", which(colnames(salary_history) == "Date"), salary_history)
apply_numeric_format(wb, "Salary History", which(colnames(salary_history) %in% c("Value", "worked", "Worked_Salary")), salary_history)

# Write Early Retirement Factors
addWorksheet(wb, "Early Retirement Factors")
writeData(wb, "Early Retirement Factors", early_retirement_factors)

# Write Salary Increases
addWorksheet(wb, "Salary Increases")
writeData(wb, "Salary Increases", salary_increases)

# Write Revaluation Factors
addWorksheet(wb, "Revaluation Factors")
writeData(wb, "Revaluation Factors", revaluation_factors)

# Save the workbook to the output directory
saveWorkbook(wb, file.path(output_dir, "Step1_Loaded_Data.xlsx"), overwrite = TRUE)
cat("Step 1 Completed Successfully! All Data Saved to Excel.\n")
