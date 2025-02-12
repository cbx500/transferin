# Block 3: Cumulative Mortality Probability Calculator (Updated)
# Description:
# This script calculates cumulative survival and death probabilities for a pension plan member (staff)
# and their spouse for two scenarios: 
#   T  (baseline; e.g., the staff member is 65 in 2028 and ages dynamically from there) and 
#   T+1 (shifted; e.g., the staff member is assumed to be 66 in 2028 and then ages dynamically).
# These dual calculations are used later to interpolate for fractional retirement ages when computing
# the present value of life annuities.
#
# Additionally, the script selects the appropriate mortality table for each individual dynamically,
# based on the gender information (e.g., "Gender_Staff" and "Gender_Spouse") extracted from the input Excel.
#
# The calculated results (cumulative survival, cumulative death probabilities, and logged qx values)
# for both scenarios (T and T+1) are exported to an Excel workbook.

library(openxlsx)
library(yaml)

cat("Loading configuration...\n")
# Load the YAML configuration file
config_path <- file.path(getwd(), "config", "config.yaml")
if (!file.exists(config_path)) {
  stop("Configuration file not found. Please check the path:", config_path)
}
config <- yaml::read_yaml(config_path)
cat("Configuration loaded successfully from:", config_path, "\n")

# Use the unified root_path from YAML (assumed to be absolute)
root_path <- normalizePath(config$root_path, winslash = "/", mustWork = TRUE)

# Construct paths using the unified configuration
male_mortality_path   <- file.path(root_path, config$male_mortality_table_path)
female_mortality_path <- file.path(root_path, config$female_mortality_table_path)
input_data_path       <- file.path(root_path, config$input_data_path)
output_path           <- file.path(root_path, config$mortality_output_path)

cat("Using the following paths:\n")
cat(" - Male Mortality Table:", male_mortality_path, "\n")
cat(" - Female Mortality Table:", female_mortality_path, "\n")
cat(" - Input Data File:", input_data_path, "\n")
cat(" - Mortality Output File:", output_path, "\n")

# Load mortality data
male_mortality_data   <- read.xlsx(male_mortality_path, sheet = 1)
female_mortality_data <- read.xlsx(female_mortality_path, sheet = 1)

# Load input data from the Excel file (for parameters such as retirement age, gender, etc.)
input_data <- read.xlsx(input_data_path, sheet = 1)

# Extract required parameters
staff_retirement_age      <- as.numeric(input_data$Value[input_data$Parameter == "Retirement_Age_Staff"])
spouse_age_at_retirement  <- as.numeric(input_data$Value[input_data$Parameter == "Spouse_Age_whenstaffretirement"])
retirement_date           <- as.Date(as.numeric(input_data$Value[input_data$Parameter == "Retirement_Date_Staff"]),
                                     origin = "1899-12-30")
retirement_year           <- as.numeric(format(retirement_date, "%Y"))

# Extract and clean gender information
gender_staff  <- tolower(trimws(input_data$Value[input_data$Parameter == "Gender_Staff"]))
gender_spouse <- tolower(trimws(input_data$Value[input_data$Parameter == "Gender_Spouse"]))

# Select appropriate mortality tables based on gender
if (gender_staff == "male") {
  staff_mortality_table <- male_mortality_data
} else if (gender_staff == "female") {
  staff_mortality_table <- female_mortality_data
} else {
  stop("Invalid or missing Gender_Staff in the input file.")
}

if (gender_spouse == "male") {
  spouse_mortality_table <- male_mortality_data
} else if (gender_spouse == "female") {
  spouse_mortality_table <- female_mortality_data
} else {
  stop("Invalid or missing Gender_Spouse in the input file.")
}

# Debug: Print key extracted parameters
cat("Staff Retirement Age:", staff_retirement_age, "\n")
cat("Spouse Age at Retirement:", spouse_age_at_retirement, "\n")
cat("Retirement Year:", retirement_year, "\n")
cat("Staff Gender:", gender_staff, "\n")
cat("Spouse Gender:", gender_spouse, "\n")

# Initialize data frames to log qx values for both scenarios
staff_qx_log_T   <- data.frame(Year = integer(), Age = numeric(), q_x = numeric())
spouse_qx_log_T  <- data.frame(Year = integer(), Age = numeric(), q_x = numeric())
staff_qx_log_T1  <- data.frame(Year = integer(), Age = numeric(), q_x = numeric())
spouse_qx_log_T1 <- data.frame(Year = integer(), Age = numeric(), q_x = numeric())

# Function to calculate cumulative probabilities and log qx values
calculate_cumulative_probabilities_with_qx_log <- function(mortality_data, start_age, start_year, log_df) {
  cumulative_survival_prob <- 1  # Initialize with 100% survival
  results <- data.frame(Year = integer(), Age = numeric(), 
                        Cumulative_Death_Prob = numeric(), 
                        Cumulative_Survival_Prob = numeric())
  
  for (year_offset in 0:(max(mortality_data$Age) - floor(start_age))) {
    current_year <- start_year + year_offset
    current_age  <- round(start_age + year_offset)
    
    if (current_age %in% mortality_data$Age && as.character(current_year) %in% colnames(mortality_data)) {
      q_x <- mortality_data[mortality_data$Age == current_age, as.character(current_year)]
      if (length(q_x) != 1) stop("Unexpected number of qx values found for age ", current_age, " and year ", current_year)
      
      log_df <- rbind(log_df, data.frame(Year = current_year, Age = current_age, q_x = q_x))
      cumulative_survival_prob <- cumulative_survival_prob * (1 - q_x)
      cumulative_death_prob <- 1 - cumulative_survival_prob
      results <- rbind(results, data.frame(Year = current_year, Age = current_age,
                                           Cumulative_Death_Prob = cumulative_death_prob,
                                           Cumulative_Survival_Prob = cumulative_survival_prob))
    } else {
      break
    }
  }
  return(list(results = results, log_df = log_df))
}

# Calculate cumulative probabilities for Scenario T (baseline)
staff_output_T  <- calculate_cumulative_probabilities_with_qx_log(staff_mortality_table, staff_retirement_age, retirement_year, staff_qx_log_T)
spouse_output_T <- calculate_cumulative_probabilities_with_qx_log(spouse_mortality_table, spouse_age_at_retirement, retirement_year, spouse_qx_log_T)

# Calculate cumulative probabilities for Scenario T+1 (one-year shift)
staff_output_T1  <- calculate_cumulative_probabilities_with_qx_log(staff_mortality_table, staff_retirement_age + 1, retirement_year, staff_qx_log_T1)
spouse_output_T1 <- calculate_cumulative_probabilities_with_qx_log(spouse_mortality_table, spouse_age_at_retirement + 1, retirement_year, spouse_qx_log_T1)

# Extract the results and logs for both scenarios
staff_results_T  <- staff_output_T$results
staff_qx_log_T   <- staff_output_T$log_df
spouse_results_T <- spouse_output_T$results
spouse_qx_log_T  <- spouse_output_T$log_df

staff_results_T1  <- staff_output_T1$results
staff_qx_log_T1   <- staff_output_T1$log_df
spouse_results_T1 <- spouse_output_T1$results
spouse_qx_log_T1  <- spouse_output_T1$log_df

# Export the results to an Excel workbook with separate sheets for each scenario
wb <- createWorkbook()

# Scenario T sheets
addWorksheet(wb, "Staff_Cumulative_Survival_T")
addWorksheet(wb, "Staff_Cumulative_Death_T")
addWorksheet(wb, "Spouse_Cumulative_Survival_T")
addWorksheet(wb, "Spouse_Cumulative_Death_T")
addWorksheet(wb, "Staff_qx_Log_T")
addWorksheet(wb, "Spouse_qx_Log_T")

# Scenario T+1 sheets
addWorksheet(wb, "Staff_Cumulative_Survival_T1")
addWorksheet(wb, "Staff_Cumulative_Death_T1")
addWorksheet(wb, "Spouse_Cumulative_Survival_T1")
addWorksheet(wb, "Spouse_Cumulative_Death_T1")
addWorksheet(wb, "Staff_qx_Log_T1")
addWorksheet(wb, "Spouse_qx_Log_T1")

# Write Scenario T data
writeData(wb, "Staff_Cumulative_Survival_T", staff_results_T)
writeData(wb, "Staff_Cumulative_Death_T", staff_results_T[, c("Year", "Age", "Cumulative_Death_Prob")])
writeData(wb, "Spouse_Cumulative_Survival_T", spouse_results_T)
writeData(wb, "Spouse_Cumulative_Death_T", spouse_results_T[, c("Year", "Age", "Cumulative_Death_Prob")])
writeData(wb, "Staff_qx_Log_T", staff_qx_log_T)
writeData(wb, "Spouse_qx_Log_T", spouse_qx_log_T)

# Write Scenario T+1 data
writeData(wb, "Staff_Cumulative_Survival_T1", staff_results_T1)
writeData(wb, "Staff_Cumulative_Death_T1", staff_results_T1[, c("Year", "Age", "Cumulative_Death_Prob")])
writeData(wb, "Spouse_Cumulative_Survival_T1", spouse_results_T1)
writeData(wb, "Spouse_Cumulative_Death_T1", spouse_results_T1[, c("Year", "Age", "Cumulative_Death_Prob")])
writeData(wb, "Staff_qx_Log_T1", staff_qx_log_T1)
writeData(wb, "Spouse_qx_Log_T1", spouse_qx_log_T1)

# Save the workbook to the specified output path
saveWorkbook(wb, output_path, overwrite = TRUE)
cat("Cumulative probability calculations and qx logs exported to", output_path, "\n")
