# ------------------------------------------------------------------------------
# Title: Dynamic Salary Revaluation and Retirement Analysis
# Description:
# - Reads salary data from the "Final Salaries" sheet in "Final_Salaries_With_Worked.xlsx".
# - Calculates revalued salary tables dynamically for retirement ages (55 to 65).
# - Uses compounded revaluation factors from the matrix.
# - Dynamically adjusts based on:
#   - Current age (extracted from the staff DOB stored in "Step1_Loaded_Data.xlsx").
#   - Matrix of cumulative revaluation factors.
#   - Adjusts the `worked` column to ensure the last value for age 65 is used across all retirement ages.
# - Outputs a comprehensive Excel workbook with a separate sheet for each retirement age.
# ------------------------------------------------------------------------------

# Step 1: Load Required Libraries
library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)
library(yaml)

cat("Step 1: Loading Required Libraries...\n")

# ------------------------------------------------------------------------------
# Step 2: Define Input and Output Paths
# ------------------------------------------------------------------------------
cat("Step 2: Defining File Paths...\n")

# Load YAML Configuration
config <- yaml::read_yaml("config/config.yaml")

# Define paths using the unified YAML configuration
salary_file_path <- file.path(config$root_path, "output", "Final_Salaries_With_Worked.xlsx")
matrix_path      <- file.path(config$root_path, "output", "Compounded_Revaluation_Matrix.xlsx")
output_path      <- file.path(config$root_path, "output", "Final_Revalued_Tables.xlsx")

# Load Step 1 data to retrieve report_date and staff_dob from "Step1_Loaded_Data.xlsx" (sheet "Personal Data")
step1_data <- read_excel(file.path(config$root_path, "output", "Step1_Loaded_Data.xlsx"), sheet = "Personal Data") %>%
  clean_names()

# ------------------------------------------------------------------------------
# Step 3: Extract and Validate Staff DOB and Report Date
# ------------------------------------------------------------------------------
cat("Step 3: Extracting Staff DOB and Report Date...\n")

# Extract staff DOB (assumed stored in a field named "dob" in lowercase)
staff_dob_raw <- step1_data %>%
  filter(tolower(field) == "dob") %>%
  pull(value)
cat("Raw staff_dob Value: ", staff_dob_raw, "\n")

staff_dob <- if (all(grepl("^[0-9]+$", staff_dob_raw))) {
  as.numeric(staff_dob_raw) %>% as.Date(origin = "1899-12-30")
} else if (all(grepl("^\\d{4}-\\d{2}-\\d{2}$", staff_dob_raw))) {
  as.Date(staff_dob_raw, format = "%Y-%m-%d")
} else {
  stop("Invalid DOB value: Unrecognized format.")
}
cat("Converted staff_dob: ", format(staff_dob, "%Y-%m-%d"), "\n")

# Extract report date
report_date_raw <- step1_data %>%
  filter(tolower(field) == "report_date") %>%
  pull(value)
cat("Raw report_date Value: ", report_date_raw, "\n")

report_date <- if (all(grepl("^[0-9]+$", report_date_raw))) {
  as.numeric(report_date_raw) %>% as.Date(origin = "1899-12-30")
} else if (all(grepl("^\\d{4}-\\d{2}-\\d{2}$", report_date_raw))) {
  as.Date(report_date_raw, format = "%Y-%m-%d")
} else {
  stop("Invalid Report Date value: Unrecognized format.")
}
cat("Converted report_date: ", format(report_date, "%Y-%m-%d"), "\n")

if (is.na(staff_dob) | is.na(report_date)) {
  stop("Error: Missing or invalid DOB or Report Date.")
}

# Calculate the year staff turns 65 and current age
year_turn_65 <- year(staff_dob) + 65
current_age <- year(report_date) - year(staff_dob)

# Determine the required years (columns) for the revaluation matrix
if (current_age < 55) {
  required_years <- as.character((year(staff_dob) + 55):year_turn_65)
} else {
  required_years <- as.character((year(report_date) + 1):year_turn_65)
}
cat("Debug: Required Years for Revaluation Matrix: ", paste(required_years, collapse = ", "), "\n")

# ------------------------------------------------------------------------------
# Step 4: Load Salary Data and Revaluation Matrix
# ------------------------------------------------------------------------------
cat("Step 4: Loading Salary Data and Revaluation Matrix...\n")

# Load the salary data from the "Final Salaries" sheet
salary_data <- read_excel(salary_file_path, sheet = "Final Salaries") %>%
  clean_names()
if (!all(c("year", "age", "starting_salary", "worked", "worked_salary") %in% colnames(salary_data))) {
  stop("Error: Required columns are missing in the salary data!")
}
cat("Salary Data Loaded Successfully.\n")

# Load the revaluation matrix
revaluation_matrix <- read_excel(matrix_path)
# Remove any 'x' prefix from column names (if present)
colnames(revaluation_matrix) <- gsub("^x", "", colnames(revaluation_matrix))
cat("Revaluation Matrix Loaded Successfully and Column Names Cleaned.\n")

# ------------------------------------------------------------------------------
# Step 5: Retrieve the `worked` Value for Age 65
# ------------------------------------------------------------------------------
cat("Step 5: Retrieving `worked` Value for Age 65...\n")
worked_age_65 <- salary_data %>%
  filter(age == 65) %>%
  pull(worked)
if (length(worked_age_65) == 0) {
  stop("Error: No `worked` value found for age 65 in the salary data!")
}
cat("`worked` Value for Age 65 Retrieved: ", worked_age_65, "\n")

# ------------------------------------------------------------------------------
# Step 6: Process Retirement Ages and Adjust the `worked` Column
# ------------------------------------------------------------------------------
cat("Step 6: Processing Retirement Ages...\n")

# Determine retirement ages based on current age
if (current_age < 55) {
  retirement_ages <- 55:65
} else if (current_age < 65) {
  retirement_ages <- seq(current_age + 1, 65)
} else {
  retirement_ages <- 65
}
retirement_ages <- retirement_ages[retirement_ages <= 65]
cat("Retirement Ages to Process: ", paste(retirement_ages, collapse = ", "), "\n")

# Initialize a list to store processed tables for each retirement age
retirement_tables <- list()

for (ret_age in retirement_ages) {
  cat("Processing Retirement Age: ", ret_age, "\n")
  
  temp_table <- salary_data %>%
    filter(age <= ret_age) %>%
    mutate(
      target_year = as.character(pmin(year(report_date) + (ret_age - current_age), year_turn_65))
    ) %>%
    rowwise() %>%
    mutate(
      # Look up the revaluation factor from the matrix based on the target_year
      revaluation_factor = if (target_year %in% colnames(revaluation_matrix)) {
        revaluation_matrix %>%
          filter(origin_year == year) %>%
          pull(target_year)
      } else {
        NA
      },
      worked_salary = worked * starting_salary * 12,
      revalued_salary = worked_salary * revaluation_factor,
      retirement_age = ret_age
    ) %>%
    ungroup()
  
  # Adjust the last row's worked value to be the one for age 65
  temp_table <- temp_table %>%
    mutate(
      worked = ifelse(row_number() == n(), worked_age_65, worked),
      worked_salary = worked * starting_salary * 12,
      revalued_salary = worked_salary * revaluation_factor
    ) %>%
    mutate(year = as.character(year))
  
  cat("Debug: Worked column after adjustment for retirement age", ret_age, ":\n")
  print(temp_table$worked)
  
  total_row <- temp_table %>%
    summarize(
      year = "Total",
      age = NA,
      starting_salary = NA,
      worked = sum(worked, na.rm = TRUE),
      worked_salary = sum(worked_salary, na.rm = TRUE),
      cumulative_revaluation_factor = NA,
      revalued_salary = sum(revalued_salary, na.rm = TRUE),
      retirement_age = ret_age
    ) %>%
    mutate(year = as.character(year))
  
  temp_table <- bind_rows(temp_table, total_row)
  retirement_tables[[as.character(ret_age)]] <- temp_table
}

cat("All tables processed successfully.\n")

# ------------------------------------------------------------------------------
# Step 7: Export the Revalued Tables to Excel
# ------------------------------------------------------------------------------
cat("Step 7: Exporting Results to Excel...\n")
wb <- createWorkbook()
for (ret_age in retirement_ages) {
  cat("Adding data for Retirement Age: ", ret_age, "\n")
  sheet_name <- paste0("Retirement Age ", ret_age)
  addWorksheet(wb, sheet_name)
  retirement_data <- retirement_tables[[as.character(ret_age)]]
  writeData(wb, sheet_name, retirement_data)
}
saveWorkbook(wb, output_path, overwrite = TRUE)
cat("Results exported successfully to: ", output_path, "\n")
