# ------------------------------------------------------------------------------
# Title: Pension Report Generation Script (No Hardcoding)
# Description:
#   This script generates the pension report for each retirement age (from current age +1 to 65),
#   and calculates the prospective retirement year for each age. The prospective retirement year 
#   is calculated based on the staff's date of birth and retirement age, and the date is formatted 
#   as "YYYY-MM". It also ensures that all numeric values in the report are properly formatted 
#   without scientific notation. The report is then saved to an Excel file.
#
#   Update: Revalued earnings calculation now considers only the last 30 years 
#   relative to each retirement age.
# ------------------------------------------------------------------------------

library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)
library(yaml)

cat("Loading Data and Generating Report...\n")

# Load YAML Configuration
config <- yaml::read_yaml("config/config.yaml")

# Define paths dynamically based on YAML
static_data_path <- file.path(config$root_path, config$static_data_path)   # Path to Static_Data.xlsx
personal_data_path <- file.path(config$root_path, config$personal_data_path) # Folder containing Personal_Data.xlsx
final_revalued_tables_path <- file.path(config$root_path, config$output_dir, "Final_Revalued_Tables.xlsx")

# Define full paths to the necessary files
personal_data_file <- file.path(personal_data_path, "Personal_Data.xlsx")
static_data_file <- static_data_path  # Use as is

# ------------------------------------------------------------------------------
# Step 3: Load Static and Personal Data
# ------------------------------------------------------------------------------
# Load Static Data
static_data <- read_excel(static_data_file, sheet = "StaticData")
inflation_rate <- as.numeric(static_data %>% filter(Parameter == "inflation_rate") %>% pull(Value))
discount_rate <- as.numeric(static_data %>% filter(Parameter == "discount_rate") %>% pull(Value))
accrual_rate <- as.numeric(static_data %>% filter(Parameter == "accrual_rate") %>% pull(Value))

cat("Static Data Loaded:\n")
cat("  Inflation Rate:", inflation_rate, "\n")
cat("  Discount Rate:", discount_rate, "\n")
cat("  Accrual Rate:", accrual_rate, "\n")

# Load Personal Data (for Staff DOB, Joining Date, and Report Date)
personal_data <- read_excel(personal_data_file, sheet = "PersonalData")
staff_dob <- as.Date(as.numeric(personal_data %>% filter(Field == "DOB") %>% pull(Value)), origin = "1899-12-30")
joining_date <- as.Date(as.numeric(personal_data %>% filter(Field == "ecb_joining_date") %>% pull(Value)), origin = "1899-12-30")
report_date <- as.Date(as.numeric(personal_data %>% filter(Field == "report_date") %>% pull(Value)), origin = "1899-12-30")

cat("Personal Details Loaded:\n")
cat("  Staff DOB:", format(staff_dob, "%Y-%m-%d"), "\n")
cat("  Joining Date:", format(joining_date, "%Y-%m-%d"), "\n")
cat("  Report Date:", format(report_date, "%Y-%m-%d"), "\n")

# Calculate current age in years (for later use)
current_age <- year(report_date) - year(staff_dob)

# ------------------------------------------------------------------------------
# Step 4: Determine Retirement Ages to Process
# ------------------------------------------------------------------------------
# Use a fixed range if current age is less than 55; otherwise, use dynamic range from (current_age+1) to 65
if (current_age < 55) {
  retirement_ages <- 55:65
} else {
  retirement_ages <- (current_age + 1):65
}
cat("Retirement Ages to Process:", paste(retirement_ages, collapse = ", "), "\n")

# ------------------------------------------------------------------------------
# Step 5: Generate Pension Report for Each Retirement Age
# ------------------------------------------------------------------------------
# Initialize list to store report data for each retirement age
output_data <- list()

# Function to calculate retirement date at end of birthday month
calculate_retirement_date <- function(dob, retirement_age) {
  retirement_year <- year(dob) + retirement_age
  retirement_date <- make_date(year = retirement_year, month = month(dob), day = day(dob))
  ceiling_date(retirement_date, "month") - days(1)
}

# Loop through each retirement age to process its data from the revalued earnings table
for (retirement_age in retirement_ages) {
  cat("Processing Retirement Age:", retirement_age, "\n")
  
  # Build the expected sheet name in Final_Revalued_Tables.xlsx
  sheet_name <- paste0("Retirement Age ", retirement_age)
  
  # Check if the sheet exists; if not, skip this age
  available_sheets <- excel_sheets(final_revalued_tables_path)
  if (!(sheet_name %in% available_sheets)) {
    cat("  Skipping: Sheet", sheet_name, "not found.\n")
    next
  }
  
  # Read the revalued earnings data for this retirement age
  final_salaries_data <- read_excel(final_revalued_tables_path, sheet = sheet_name)
  
  # Clean the data by removing rows with missing age
  final_salaries_data <- final_salaries_data %>% filter(!is.na(age))
  
  # Ensure we have a record for the exact retirement age
  matching_index <- which(final_salaries_data$age == retirement_age)
  if (length(matching_index) == 0) {
    stop(paste("Error: Final salary not found for retirement age", retirement_age))
  }
  
  filtered_row <- final_salaries_data[matching_index, ]
  final_salary_at_retirement <- filtered_row$starting_salary
  
  cat("  Final Salary at Retirement Age", retirement_age, ":", final_salary_at_retirement, "\n")
  
  # Calculate average revalued earnings over the last 30 years (if available)
  revalued_earnings <- final_salaries_data %>%
    filter(age >= (retirement_age - 30) & age <= retirement_age) %>%
    summarise(avg_revalued_earnings = sum(revalued_salary, na.rm = TRUE) / sum(worked, na.rm = TRUE)) %>%
    pull(avg_revalued_earnings)
  
  if (is.na(revalued_earnings)) {
    stop(paste("Error: Revalued earnings calculation failed for retirement age", retirement_age))
  }
  
  # Calculate total pensionable service (sum of worked values up to this age)
  pensionable_service <- sum(final_salaries_data$worked[final_salaries_data$age <= retirement_age], na.rm = TRUE)
  cat("  Pensionable Service for Age", retirement_age, ":", pensionable_service, "\n")
  
  # Calculate the scheme pension before discounting
  scheme_pension_retirement <- (revalued_earnings * accrual_rate * pensionable_service) / 12
  
  # Determine retirement date and years to discount
  r_date <- calculate_retirement_date(staff_dob, retirement_age)
  years_to_discount <- as.numeric(difftime(r_date, report_date, units = "days")) / 365.25
  cat("  Retirement Date:", format(r_date, "%Y-%m-%d"), "Years to Discount:", round(years_to_discount, 6), "\n")
  
  # Discount final salary, revalued earnings, and scheme pension back to the report date
  final_salary_discounted <- final_salary_at_retirement / ((1 + inflation_rate) ^ years_to_discount)
  revalued_earnings_discounted <- revalued_earnings / ((1 + inflation_rate) ^ years_to_discount)
  scheme_pension_discounted <- scheme_pension_retirement / ((1 + inflation_rate) ^ years_to_discount)
  
  # Calculate the prospective retirement year (formatted as "YYYY-MM")
  prospective_year <- paste(retirement_age + year(staff_dob), format(staff_dob, "%m"), sep = "-")
  
  # Store all calculated data in the output list
  output_data[[as.character(retirement_age)]] <- tibble(
    staff_name = personal_data %>% filter(Field == "staff_name") %>% pull(Value),
    staff_dob = format(staff_dob, "%Y-%m-%d"),
    retirement_age = retirement_age,
    prospective_year = prospective_year,
    final_salary_at_retirement = round(final_salary_at_retirement, 2),
    final_salary_at_report_date = round(final_salary_discounted, 2),
    avg_revalued_earnings_at_retirement = round(revalued_earnings, 2),
    avg_revalued_earnings_at_report_date = round(revalued_earnings_discounted, 2),
    pensionable_service = pensionable_service,
    scheme_pension_at_retirement = round(scheme_pension_retirement, 2),
    scheme_pension_at_report_date = round(scheme_pension_discounted, 2),
    pcf1 = NA,  # Placeholders for PCF values (to be updated elsewhere if needed)
    pcf2 = NA
  )
}

# Combine all retirement age data into one final report
final_report <- bind_rows(output_data)

# ------------------------------------------------------------------------------
# Step 8: Export the Final Report to Excel
# ------------------------------------------------------------------------------
cat("Exporting the final pension report to Excel...\n")
wb <- createWorkbook()
addWorksheet(wb, "Pension Report")
writeData(wb, "Pension Report", final_report, rowNames = FALSE)

# Apply numeric formatting to ensure no scientific notation appears
num_format <- createStyle(numFmt = "0.00")
addStyle(wb, "Pension Report", num_format, rows = 2:(nrow(final_report) + 1), cols = 4:ncol(final_report), gridExpand = TRUE)

# Define output path for the pension report
pension_report_path <- file.path(config$root_path, config$output_dir, "Pension_Report.xlsx")

# Save the workbook
saveWorkbook(wb, pension_report_path, overwrite = TRUE)

cat("Report Generated Successfully! Saved to:", pension_report_path, "\n")
