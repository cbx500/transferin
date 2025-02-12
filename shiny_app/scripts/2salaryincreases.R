# ------------------------------------------------------------------------------
# Title: Projected Salaries Calculation Script
# Description:
#   This script processes the personal data and salary history from the Step1 output
#   and calculates projected salaries up to the retirement year. It includes:
#   - Extracting and validating dates from personal data.
#   - Loading static data and salary history.
#   - Calculating projected salaries based on age and salary increments.
#   - Exporting the projected salaries to an Excel file.
# ------------------------------------------------------------------------------

# Load Required Libraries
library(dplyr)
library(lubridate)
library(openxlsx)
library(readxl)
library(yaml)
library(janitor)

cat("Loading configuration...\n")
config <- yaml::read_yaml("config/config.yaml")

# Ensure that configuration fields are characters (in case they are not)
project_root <- as.character(config$root_path)  # changed from config$project_root to config$root_path
output_dir   <- as.character(config$output_dir)
static_data_path <- as.character(config$static_data_path)

# Define paths dynamically based on YAML
step1_output_path <- file.path(project_root, output_dir, "Step1_Loaded_Data.xlsx")
projected_salaries_path <- file.path(project_root, output_dir, "Projected_Salaries.xlsx")

cat("Step1 output file path:", step1_output_path, "\n")
cat("Projected salaries output file path:", projected_salaries_path, "\n")

# Check if the Step1 output file exists
if (!file.exists(step1_output_path)) {
  stop("Step1 output file not found at: ", step1_output_path)
}

# Step 1: Load Personal Data from Step1 Output
cat("Loading Personal Data from Step1 output...\n")
personal_data <- read_excel(step1_output_path, sheet = "Personal Data", col_types = "text") %>%
  clean_names()

# Extract and Convert Dates
cat("Extracting and converting dates...\n")
# Staff DOB
cat("Extracting Staff DOB...\n")
staff_dob_raw <- personal_data %>% filter(tolower(field) == "dob") %>% pull(value)
cat("Raw Staff DOB Value:", staff_dob_raw, "\n")
staff_dob <- if (all(grepl("^[0-9]+$", staff_dob_raw))) {
  as.numeric(staff_dob_raw) %>% as.Date(origin = "1899-12-30")
} else if (all(grepl("^\\d{4}-\\d{2}-\\d{2}$", staff_dob_raw))) {
  as.Date(staff_dob_raw, format = "%Y-%m-%d")
} else {
  stop("Invalid DOB value: Unrecognized format.")
}
cat("Converted Staff DOB:", format(staff_dob, "%Y-%m-%d"), "\n")

# Report Date
cat("Extracting Report Date...\n")
report_date_raw <- personal_data %>% filter(tolower(field) == "report_date") %>% pull(value)
cat("Raw Report Date Value:", report_date_raw, "\n")
report_date <- if (all(grepl("^[0-9]+$", report_date_raw))) {
  as.numeric(report_date_raw) %>% as.Date(origin = "1899-12-30")
} else if (all(grepl("^\\d{4}-\\d{2}-\\d{2}$", report_date_raw))) {
  as.Date(report_date_raw, format = "%Y-%m-%d")
} else {
  stop("Invalid Report Date value: Unrecognized format.")
}
cat("Converted Report Date:", format(report_date, "%Y-%m-%d"), "\n")

# Validate essential dates
if (is.na(staff_dob) | is.na(report_date)) {
  stop("Error: Missing or invalid Staff DOB or Report Date in Personal Data.")
}

# Step 2: Load Salary History from Step1 Output
cat("Loading Salary History...\n")
salary_history <- read_excel(step1_output_path, sheet = "Salary History") %>%
  clean_names()

# Extract the most recent salary record (assumes a 'date' column exists)
last_salary_record <- salary_history %>%
  filter(!is.na(value)) %>%
  arrange(desc(date)) %>%
  slice(1)
base_salary <- as.numeric(last_salary_record$value)
cat("Base Salary Loaded:", base_salary, "\n")

# Step 3: Load Static Data (Inflation Rate and Salary Increases)
cat("Loading Static Data...\n")
static_data <- read_excel(static_data_path, sheet = "StaticData", col_types = "text") %>%
  clean_names()
inflation_rate <- as.numeric(static_data %>% filter(parameter == "inflation_rate") %>% pull(value))
salary_increases <- read_excel(static_data_path, sheet = "SalaryIncreases", col_types = "text") %>%
  clean_names() %>%
  mutate(
    from_age = as.numeric(from_age),
    increase_percentage = as.numeric(increase_percentage)
  )
cat("Static Data Loaded:\n")
cat("Inflation Rate:", inflation_rate, "\n")
cat("Salary Increases:\n")
print(salary_increases)

# Step 4: Calculate Projected Salaries
current_year <- year(report_date)
retirement_age <- 65
retirement_year <- year(staff_dob) + retirement_age
current_age <- current_year - year(staff_dob)

projected_salaries <- tibble()

cat("Calculating Projected Salaries...\n")
for (yr in seq(current_year + 1, retirement_year)) {
  projected_age <- current_age + (yr - current_year)
  
  # Find the salary increment applicable at the projected age
  age_increment <- salary_increases %>%
    filter(from_age <= projected_age) %>%
    arrange(desc(from_age)) %>%
    slice(1) %>%
    pull(increase_percentage)
  
  # Apply inflation and the age increment; if projected_age >= 55, set increment to 0
  if (projected_age >= 55) {
    age_increment <- 0
  }
  
  revaluation_factor <- 1 + inflation_rate + age_increment
  
  # Update the base_salary for each year (cumulatively)
  base_salary <- base_salary * revaluation_factor
  
  projected_salaries <- bind_rows(projected_salaries, tibble(
    Field = "Projected_Salary",
    Date = as.Date(paste0(yr, "-01-01")),
    Value = NA,
    `%worked` = 1,
    Worked_Salary = base_salary,
    Age = projected_age
  ))
}

cat("Projected Salaries Until Retirement:\n")
print(projected_salaries)

# Step 5: Export Projected Salaries to Excel
cat("Exporting Projected Salaries to Excel...\n")
wb <- createWorkbook()
addWorksheet(wb, "Projected Salaries")
writeData(wb, "Projected Salaries", projected_salaries)
saveWorkbook(wb, projected_salaries_path, overwrite = TRUE)
cat("Projected Salaries have been saved to:", projected_salaries_path, "\n")
