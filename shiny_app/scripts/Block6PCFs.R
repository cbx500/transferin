# Block 6: Pension Calculation for Staff and Spouse
# Description:
# This script calculates pension metrics for the staff member and their spouse.
# It handles date and numeric formatting, loads the required cash flow data from previous blocks,
# performs calculations for both Scenario T (baseline) and Scenario T+1 (one-year shift),
# computes weighted pension conversion factors (PCF) based on fractional months,
# and saves the results directly to an Excel file.

library(openxlsx)  # For reading and writing Excel files
library(dplyr)     # For data manipulation
library(yaml)      # For reading configuration files

cat("Loading configuration...\n")
# Load YAML configuration file (using unified settings)
config_path <- file.path(getwd(), "config", "config.yaml")
if (!file.exists(config_path)) {
  stop("Configuration file not found. Please check the path:", config_path)
}
config <- yaml::read_yaml(config_path)
cat("Configuration loaded from:", config_path, "\n")

# Use the unified root_path from YAML
root_path <- normalizePath(config$root_path, winslash = "/", mustWork = TRUE)

# Construct paths dynamically using unified configuration
input_data_path <- file.path(root_path, config$input_data_path)
cash_flow_output_path <- file.path(root_path, config$cash_flow_output_path)
# For the validation output, where the PCF and other metrics will be stored:
validation_output_path <- file.path(root_path, "output", "MonthlyEquivalentAnnuityAndPCFCalculation_Validation.xlsx")

# ---------------------------
# Step 2: Load Input Data
# ---------------------------
cat("Loading input data from Excel...\n")
if (!file.exists(input_data_path)) stop("Input data file not found at:", input_data_path)
input_data <- read.xlsx(input_data_path, sheet = 1)

# ---------------------------
# Step 3: Extract Values Dynamically from Input Data
# ---------------------------
cat("Extracting parameters from input data...\n")
staff_name <- input_data$Value[input_data$Parameter == "Staff_Name"]
dob_staff <- as.Date(as.numeric(input_data$Value[input_data$Parameter == "DOB_Staff"]), origin = "1899-12-30")
gender_staff <- input_data$Value[input_data$Parameter == "Gender_Staff"]
dob_spouse <- as.Date(as.numeric(input_data$Value[input_data$Parameter == "DOB_Spouse"]), origin = "1899-12-30")
gender_spouse <- input_data$Value[input_data$Parameter == "Gender_Spouse"]
retirement_date_staff <- as.Date(as.numeric(input_data$Value[input_data$Parameter == "Retirement_Date_Staff"]), origin = "1899-12-30")
retirement_age_staff <- as.numeric(input_data$Value[input_data$Parameter == "Retirement_Age_Staff"])
staff_target_pension <- as.numeric(input_data$Value[input_data$Parameter == "staff_target_pension"])
spousetargetpension <- as.numeric(input_data$Value[input_data$Parameter == "spousetargetpension"])
staff_annuity_value <- as.numeric(input_data$Value[input_data$Parameter == "Staff_annuity"])
marriedprob <- as.numeric(input_data$Value[input_data$Parameter == "marriedprob"])

# NEW: Additional Numeric Parameters for fractional months
months_staff <- as.numeric(input_data$Value[input_data$Parameter == "months_staff"])
months_spouse <- as.numeric(input_data$Value[input_data$Parameter == "months_spouse"])

# Extract discount rate (NBY)
discount_rate <- as.numeric(input_data$Value[input_data$Parameter == "NBY"])

# Override parameters from sliders if present
if (exists("OVERRIDE_PARAMS", envir = .GlobalEnv)) {
  over <- get("OVERRIDE_PARAMS", envir = .GlobalEnv)
  cat("OVERRIDE_PARAMS loaded:\n")
  print(over)
  
  if (!is.null(over$NBY)) {
    discount_rate <- over$NBY
  }
  if (!is.null(over$SPOUSE_TARGET)) {
    spousetargetpension <- over$SPOUSE_TARGET
  }
  if (!is.null(over$STAFF_TARGET)) {
    staff_target_pension <- over$STAFF_TARGET
  }
} else {
  cat("OVERRIDE_PARAMS not found. Using default values.\n")
}

if (is.na(discount_rate)) stop("Invalid discount rate (NBY) provided in input data. Please check.")

# Store parameters in a list
parameters <- list(
  Staff_Name = staff_name,
  DOB_Staff = format(dob_staff, "%Y-%m-%d"),
  Gender_Staff = gender_staff,
  DOB_Spouse = format(dob_spouse, "%Y-%m-%d"),
  Gender_Spouse = gender_spouse,
  Retirement_Date_Staff = format(retirement_date_staff, "%Y-%m-%d"),
  Retirement_Age_Staff = retirement_age_staff,
  Spouse_Age_whenstaffretirement = as.numeric(input_data$Value[input_data$Parameter == "Spouse_Age_whenstaffretirement"]),
  months_staff = months_staff,
  months_spouse = months_spouse,
  Market_Value_VCS = as.numeric(input_data$Value[input_data$Parameter == "Market_Value_VCS"]),
  Market_Value_FBA = as.numeric(input_data$Value[input_data$Parameter == "Market_Value_FBA"]),
  spousetargetpension = spousetargetpension,
  staff_target_pension = staff_target_pension,
  Staff_annuity = staff_annuity_value,
  Spouse_Annuity_Percentage = as.numeric(input_data$Value[input_data$Parameter == "Spouse_Annuity_Percentage"]),
  Discount_Rate_NBY = discount_rate
)

cat("Extracted Parameters:\n")
print(parameters)

# ---------------------------
# Step 4: Load Cash Flow Data for Scenarios T and T+1
# ---------------------------
cat("Loading cash flow data from the Excel file...\n")
staff_cash_flow_T  <- read.xlsx(cash_flow_output_path, sheet = "Cumulative_Cash_Flow_Staff_T")
spouse_cash_flow_T <- read.xlsx(cash_flow_output_path, sheet = "Cumulative_Cash_Flow_Spouse_T")
staff_cash_flow_T1 <- read.xlsx(cash_flow_output_path, sheet = "Cumulative_Cash_Flow_Staff_T1")
spouse_cash_flow_T1<- read.xlsx(cash_flow_output_path, sheet = "Cumulative_Cash_Flow_Spouse_T1")

# ---------------------------
# Step 5: Perform Pension Calculations for Scenarios T and T+1
# ---------------------------
cat("Performing pension calculations...\n")
# Scenario T (baseline)
staff_annuity_postpaid_T <- sum(as.numeric(staff_cash_flow_T$Product[-nrow(staff_cash_flow_T)]), na.rm = TRUE) * staff_target_pension
staff_prepaid_annuity_T <- staff_annuity_postpaid_T + (staff_annuity_value * staff_target_pension)
spouse_annuity_postpaid_T <- sum(as.numeric(spouse_cash_flow_T$Product[-nrow(spouse_cash_flow_T)]), na.rm = TRUE)
spouse_total_postpaid_annuity_T <- spouse_annuity_postpaid_T * spousetargetpension * marriedprob
staff_monthly_equivalent_annuity_T <- max(staff_prepaid_annuity_T - (12 / 24), 0)
PCF_staff_T <- staff_monthly_equivalent_annuity_T * 12
PCF_spouse_T <- spouse_total_postpaid_annuity_T * 12
total_PCF_T <- PCF_staff_T + PCF_spouse_T

# Scenario T+1 (one-year shift)
staff_annuity_postpaid_T1 <- sum(as.numeric(staff_cash_flow_T1$Product[-nrow(staff_cash_flow_T1)]), na.rm = TRUE) * staff_target_pension
staff_prepaid_annuity_T1 <- staff_annuity_postpaid_T1 + (staff_annuity_value * staff_target_pension)
spouse_annuity_postpaid_T1 <- sum(as.numeric(spouse_cash_flow_T1$Product[-nrow(spouse_cash_flow_T1)]), na.rm = TRUE)
spouse_total_postpaid_annuity_T1 <- spouse_annuity_postpaid_T1 * spousetargetpension * marriedprob
staff_monthly_equivalent_annuity_T1 <- max(staff_prepaid_annuity_T1 - (12 / 24), 0)
PCF_staff_T1 <- staff_monthly_equivalent_annuity_T1 * 12
PCF_spouse_T1 <- spouse_total_postpaid_annuity_T1 * 12
total_PCF_T1 <- PCF_staff_T1 + PCF_spouse_T1

# Calculate the weighted PCF using the fractional month parameters
PCF_weighted_staff <- ((1 - (months_staff / 12)) * PCF_staff_T) + ((months_staff / 12) * PCF_staff_T1)
PCF_weighted_spouse <- ((1 - (months_spouse / 12)) * PCF_spouse_T) + ((months_spouse / 12) * PCF_spouse_T1)
Total_PCF <- PCF_weighted_staff + PCF_weighted_spouse

# ---------------------------
# Step 6: Prepare Results Table
# ---------------------------
cat("Preparing results table...\n")
results_table <- data.frame(
  Parameter = c(
    "Staff_Name", "DOB_Staff", "Gender_Staff", "DOB_Spouse", "Gender_Spouse",
    "Retirement_Date_Staff", "Retirement_Age_Staff", "Staff_Target_Pension",
    "Spouse_Target_Pension", "Married_Probability", 
    "Staff_Postpaid_Annuity_T", "Spouse_Postpaid_Annuity_T", "Staff_Prepaid_Annuity_T",
    "Spouse_Total_Postpaid_Annuity_T", "Staff_Monthly_Equivalent_Annuity_T",
    "PCF_Staff_T", "PCF_Spouse_T", "Total_PCF_T",
    "Staff_Postpaid_Annuity_T1", "Spouse_Postpaid_Annuity_T1", "Staff_Prepaid_Annuity_T1",
    "Spouse_Total_Postpaid_Annuity_T1", "Staff_Monthly_Equivalent_Annuity_T1",
    "PCF_Staff_T1", "PCF_Spouse_T1", "Total_PCF_T1",
    "PCF_weighted_staff", "PCF_weighted_spouse", "Total_Pension_Conversion_Factor (PCF)",
    "months_staff", "months_spouse"
  ),
  Value = c(
    staff_name, format(dob_staff, "%Y-%m-%d"), gender_staff,
    format(dob_spouse, "%Y-%m-%d"), gender_spouse,
    format(retirement_date_staff, "%Y-%m-%d"), retirement_age_staff,
    round(staff_target_pension, 3), round(spousetargetpension, 3), round(marriedprob, 3),
    # Scenario T (baseline)
    round(staff_annuity_postpaid_T, 3), round(spouse_annuity_postpaid_T, 3), round(staff_prepaid_annuity_T, 3),
    round(spouse_total_postpaid_annuity_T, 3), round(staff_monthly_equivalent_annuity_T, 3),
    round(PCF_staff_T, 3), round(PCF_spouse_T, 3), round(total_PCF_T, 3),
    # Scenario T+1 (one-year shift)
    round(staff_annuity_postpaid_T1, 3), round(spouse_annuity_postpaid_T1, 3), round(staff_prepaid_annuity_T1, 3),
    round(spouse_total_postpaid_annuity_T1, 3), round(staff_monthly_equivalent_annuity_T1, 3),
    round(PCF_staff_T1, 3), round(PCF_spouse_T1, 3), round(total_PCF_T1, 3),
    # Weighted PCF
    round(PCF_weighted_staff, 3), round(PCF_weighted_spouse, 3), round(Total_PCF, 3),
    round(months_staff, 3), round(months_spouse, 3)
  ),
  stringsAsFactors = FALSE
)

# ---------------------------
# Step 7: Save Results to Excel
# ---------------------------
cat("Saving results to Excel...\n")
write.xlsx(
  list(
    # Cash flow sheets from previous Block 4 (already generated)
    "Cumulative_Cash_Flow_Staff_T" = staff_cash_flow_T,
    "Cumulative_Cash_Flow_Spouse_T" = spouse_cash_flow_T,
    "Cumulative_Cash_Flow_Staff_T1" = staff_cash_flow_T1,
    "Cumulative_Cash_Flow_Spouse_T1" = spouse_cash_flow_T1,
    # Results Table
    "Validation_Results" = results_table
  ),
  file = validation_output_path,
  overwrite = TRUE
)

cat("Validation report saved successfully to:", validation_output_path, "\n")
