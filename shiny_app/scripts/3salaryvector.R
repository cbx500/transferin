# ------------------------------------------------------------------------------
# Title: Historical and Future Salaries Table Generation
# Description:
#   Generates a consolidated table of starting salaries for all years
#   from employment start (or scheme start) until retirement.
#   Inputs (historical salary data and future salary data) are loaded dynamically.
# ------------------------------------------------------------------------------

# Load Required Libraries
library(dplyr)
library(lubridate)
library(openxlsx)
library(readxl)
library(yaml)
library(janitor)

# ------------------------------------------------------------------------------
# Step 1: Load Configuration and Input Files
# ------------------------------------------------------------------------------
cat("Loading configuration and input files...\n")
# Load the unified YAML configuration
config <- yaml::read_yaml("config/config.yaml")  # Use relative path

# Define paths dynamically using YAML settings (changed from config$project_root to config$root_path)
personal_data_path <- file.path(config$root_path, config$personal_data_path)
output_dir <- file.path(config$root_path, config$output_dir)
future_salaries_path <- file.path(output_dir, "Projected_Salaries.xlsx")  # Future salaries output from the previous block

# Sheet names as defined in the YAML for personal data
salary_history_sheet <- config$personal_data_sheets$salary_history
personal_data_sheet <- config$personal_data_sheets$personal_data
future_salaries_sheet <- "Projected Salaries"  # As produced by the projected salaries script

# ------------------------------------------------------------------------------
# Step 2: Load Personal Data with Correct Date Extraction
# ------------------------------------------------------------------------------
cat("Loading personal data and converting dates...\n")
# Define the full path to the Personal_Data.xlsx file
personal_data_file <- file.path(personal_data_path, "Personal_Data.xlsx")

# Read the personal data sheet (assumed to be named as specified in the YAML)
personal_data <- read_excel(personal_data_file, sheet = personal_data_sheet, col_types = "text")

# Convert date fields using a consistent method:
dob <- as.Date(as.numeric(personal_data$Value[personal_data$Field == "DOB"]), origin = "1899-12-30")
report_date <- as.Date(as.numeric(personal_data$Value[personal_data$Field == "report_date"]), origin = "1899-12-30")
joining_date <- as.Date(as.numeric(personal_data$Value[personal_data$Field == "ecb_joining_date"]), origin = "1899-12-30")

# Validate that key dates are present
if (is.na(dob) | is.na(report_date) | is.na(joining_date)) {
  stop("Error: Missing DOB, report_date, or joining_date in the PersonalData sheet.")
}

cat("Dates successfully converted:\n")
cat("DOB:", format(dob, "%Y-%m-%d"), "\n")
cat("Report Date:", format(report_date, "%Y-%m-%d"), "\n")
cat("Joining Date:", format(joining_date, "%Y-%m-%d"), "\n\n")

# ------------------------------------------------------------------------------
# Step 3: Load Historical Salaries
# ------------------------------------------------------------------------------
cat("Loading historical salary data...\n")
salary_history <- read_excel(personal_data_file, sheet = salary_history_sheet, col_types = "text") %>%
  mutate(
    Date = as.Date(as.numeric(Date), origin = "1899-12-30"),
    Value = as.numeric(Value)
  ) %>%
  filter(!is.na(Value))  # Remove any rows with missing salary values

# Determine the starting year from the salary history
scheme_start_year <- year(min(salary_history$Date, na.rm = TRUE))

# Calculate the employee's age for each salary record based on DOB
historical_salaries <- salary_history %>%
  mutate(
    Year = year(Date),
    Age = Year - year(dob),
    Starting_Salary = Value
  ) %>%
  select(Year, Age, Starting_Salary)

# ------------------------------------------------------------------------------
# Step 4: Load Future Salaries
# ------------------------------------------------------------------------------
cat("Loading future salaries...\n")
future_salaries <- read_excel(future_salaries_path, sheet = future_salaries_sheet) %>%
  select(Date, Age, Worked_Salary) %>%
  mutate(
    Year = year(Date),
    Starting_Salary = Worked_Salary
  ) %>%
  select(Year, Age, Starting_Salary)

# ------------------------------------------------------------------------------
# Step 5: Combine Historical and Future Salaries
# ------------------------------------------------------------------------------
cat("Combining historical and future salaries...\n")
combined_salaries <- bind_rows(historical_salaries, future_salaries) %>%
  arrange(Year)

# Check for duplicates in the Year column to ensure data consistency
if (any(duplicated(combined_salaries$Year))) {
  stop("Error: Duplicate years detected in the combined salary table.")
}

# ------------------------------------------------------------------------------
# Step 6: Export Results to Excel
# ------------------------------------------------------------------------------
cat("Exporting results to Excel...\n")
output_file <- file.path(output_dir, "Historical_and_Future_Salaries.xlsx")

wb <- createWorkbook()
addWorksheet(wb, "Salaries Table")
writeData(wb, "Salaries Table", combined_salaries)

# Save the workbook to the output directory
saveWorkbook(wb, output_file, overwrite = TRUE)

cat("Salaries table exported successfully to:", output_file, "\n")

# ------------------------------------------------------------------------------
# Step 7: Output Summary for Validation
cat("\nSummary of Combined Salaries Table:\n")
print(combined_salaries)
