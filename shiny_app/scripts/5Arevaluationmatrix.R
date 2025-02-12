# ------------------------------------------------------------------------------
# Title: Dynamic Compounded Revaluation Factors Matrix Generator
# Description:
#   - Generates a dynamic matrix of compounded revaluation factors.
#   - Dynamically adjusts based on the report year, person's date of birth (age),
#     starting year of the job, and revaluation/inflation data.
#   - Excludes values less than 1 (sets them to NA and later removes rows where all are NA).
#   - Outputs a clean dynamic matrix for future use in revaluation scripts.
# ------------------------------------------------------------------------------

# Step 1: Load Required Libraries
library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)
library(yaml)
library(janitor)  # (if needed for cleaning names)

cat("Step 1: Loading Required Libraries...\n")

# ------------------------------------------------------------------------------
# Step 2: Define Dynamic Inputs and Paths
# ------------------------------------------------------------------------------
cat("Step 2: Defining Inputs and Paths...\n")
# Load YAML configuration from the unified config file
config <- yaml::read_yaml("config/config.yaml")  # Use relative path

# Construct paths dynamically using YAML variables.
# Here we use config$root_path to build absolute paths.
step1_output_path <- file.path(config$root_path, config$output_dir, "Step1_Loaded_Data.xlsx")
output_matrix_path <- file.path(config$root_path, config$output_dir, "Compounded_Revaluation_Matrix.xlsx")

# ------------------------------------------------------------------------------
# Step 3: Load Configuration and Input Data
# ------------------------------------------------------------------------------
cat("Step 3: Loading Configuration and Input Data...\n")

# Load Revaluation Factors from the Step1 output Excel file
revaluation_factors <- read_excel(step1_output_path, sheet = "Revaluation Factors") %>%
  clean_names() %>%
  mutate(
    year = as.numeric(year),
    revaluation_factor = as.numeric(revaluation_factor)
  )

# Check that revaluation factors were loaded properly
if (any(is.na(revaluation_factors$revaluation_factor))) {
  stop("Error: Revaluation factors contain missing or invalid values.")
}

# Load Static Data from the Step1 output Excel file (to get inflation_rate)
static_data <- read_excel(step1_output_path, sheet = "Static Data") %>%
  clean_names()

inflation_rate <- static_data %>%
  filter(tolower(parameter) == "inflation_rate") %>%
  pull(value) %>%
  as.numeric()

if (is.na(inflation_rate)) {
  stop("Error: Inflation rate not found or invalid in Static Data.")
}

# Load Personal Data (for DOB and Report Date) from the Step1 output Excel file.
# (Assuming that the Step1 export already produced the Personal Data sheet.)
personal_data <- read_excel(step1_output_path, sheet = "Personal Data", col_types = "text")

# Extract and validate the DOB and Report Date from Personal Data.
dob <- personal_data %>%
  filter(tolower(Field) == "dob") %>%
  pull(Value)

cat("Raw DOB Value: ", dob, "\n")

dob <- if (all(grepl("^[0-9]+$", dob))) {
  as.numeric(dob) %>% as.Date(origin = "1899-12-30")
} else if (all(grepl("^\\d{4}-\\d{2}-\\d{2}$", dob))) {
  as.Date(dob, format = "%Y-%m-%d")
} else {
  stop("Invalid DOB value: Unrecognized format.")
}

cat("Converted DOB: ", format(dob, "%Y-%m-%d"), "\n")

report_date <- personal_data %>%
  filter(tolower(Field) == "report_date") %>%
  pull(Value)

report_date <- if (all(grepl("^[0-9]+$", report_date))) {
  as.numeric(report_date) %>% as.Date(origin = "1899-12-30")
} else if (all(grepl("^\\d{4}-\\d{2}-\\d{2}$", report_date))) {
  as.Date(report_date, format = "%Y-%m-%d")
} else {
  stop("Invalid Report Date value: Unrecognized format.")
}

cat("Converted Report Date: ", format(report_date, "%Y-%m-%d"), "\n")

if (is.na(dob) | is.na(report_date)) {
  stop("Error: Missing or invalid DOB or Report Date.")
}

# ------------------------------------------------------------------------------
# Step 4: Determine Origin and Retirement Years
# ------------------------------------------------------------------------------
cat("Step 4: Determining Origin and Retirement Years...\n")
# Starting Year of Job is taken as the minimum year in the revaluation factors table
starting_year <- min(revaluation_factors$year, na.rm = TRUE)

# Present Year is derived from the report date
present_year <- year(report_date)
cat("Present Year: ", present_year, "\n")

# Retirement Years: from the present year up to the year the person turns 65
retirement_years <- seq(present_year, year(dob) + 65)

# Origin Years for the matrix: from the starting year to the maximum retirement year
origin_years <- seq(starting_year, max(retirement_years))

cat("Starting Year: ", starting_year, "\n")
cat("Retirement Years: ", paste(retirement_years, collapse = ", "), "\n")

# ------------------------------------------------------------------------------
# Step 5: Generate Compounded Revaluation Factors Matrix
# ------------------------------------------------------------------------------
cat("Step 5: Generating Compounded Revaluation Factors Matrix...\n")
# Initialize the matrix with a column for origin_year
compounded_factors <- data.frame(origin_year = origin_years)

# For each retirement year, compute the cumulative revaluation factor for each origin year.
for (retirement_year in retirement_years) {
  compounded_factors[[as.character(retirement_year)]] <- sapply(origin_years, function(origin_year) {
    if (origin_year <= present_year) {
      # For past years, compound the revaluation factors from origin_year+1 to present_year
      past_factors <- prod(
        revaluation_factors$revaluation_factor[
          revaluation_factors$year > origin_year & revaluation_factors$year <= present_year
        ], na.rm = TRUE
      )
      
      # For future years, apply inflation from present_year to retirement_year
      future_factors <- (1 + inflation_rate)^(retirement_year - present_year)
      compounded_value <- past_factors * future_factors
    } else {
      # For future origin years, only apply inflation
      compounded_value <- (1 + inflation_rate)^(retirement_year - origin_year)
    }
    # If the compounded value is less than 1, return NA; otherwise, return the value
    ifelse(compounded_value < 1, NA, compounded_value)
  })
}

# Remove rows where all the computed values are NA
compounded_factors <- compounded_factors %>%
  filter(if_any(everything(), ~ !is.na(.)))

cat("Compounded Revaluation Factors Matrix Generated Successfully.\n")

# ------------------------------------------------------------------------------
# Step 6: Export Matrix to Excel
# ------------------------------------------------------------------------------
cat("Step 6: Exporting Compounded Revaluation Factors Matrix...\n")
wb <- createWorkbook()
addWorksheet(wb, "Revaluation Matrix")
writeData(wb, "Revaluation Matrix", compounded_factors)
saveWorkbook(wb, output_matrix_path, overwrite = TRUE)

cat("Compounded Revaluation Factors Matrix exported to:", output_matrix_path, "\n")
