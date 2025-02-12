# Block 2: Load Input Data, Extract Parameters, and Export to Excel
# Description:
# This script reads input data from an Excel file (e.g., "C:/transferin/shiny_app/data/Input_Data.xlsx"),
# validates its structure, extracts key parameters required for subsequent calculations (including two new parameters:
# months_staff and months_spouse), and outputs them as an Excel file for verification and downstream use.

library(readxl)
library(openxlsx)

cat("Validating configuration and paths...\n")
# Assume that the YAML configuration was already loaded in Block 1 and the following variables are available:
#   - config, root_path

# Construct the full path to the input data file using the unified root_path
file_path <- file.path(root_path, config$input_data_path)

# Validate if the Excel file exists
if (!file.exists(file_path)) stop("Input data file not found at:", file_path)

# Use the default sheet name "Sheet1" (adjust if necessary)
sheet_name <- "Sheet1"

# Validate if the sheet exists in the Excel file
cat("Validating sheet...\n")
available_sheets <- excel_sheets(file_path)
if (!sheet_name %in% available_sheets) {
  stop(paste("Sheet", shQuote(sheet_name), "not found in the Excel file. Available sheets:", paste(available_sheets, collapse = ", ")))
}

# Load the Input Sheet
cat("Loading input data from Excel...\n")
input_data <- read_excel(file_path, sheet = sheet_name, col_names = c("Parameter", "Value", "Notes"))

# Verify the loaded data structure
cat("Loaded Data Structure:\n")
print(head(input_data))

cat("Extracting and converting parameters...\n")
# Date Parameters: Convert Excel serial numbers to R Date format
staff_dob <- as.Date(as.numeric(input_data$Value[input_data$Parameter == "DOB_Staff"]), origin = "1899-12-30")
spouse_dob <- as.Date(as.numeric(input_data$Value[input_data$Parameter == "DOB_Spouse"]), origin = "1899-12-30")
retirement_date <- as.Date(as.numeric(input_data$Value[input_data$Parameter == "Retirement_Date_Staff"]), origin = "1899-12-30")

# String Parameters
staff_name <- input_data$Value[input_data$Parameter == "Staff_Name"]
gender_staff <- input_data$Value[input_data$Parameter == "Gender_Staff"]
gender_spouse <- input_data$Value[input_data$Parameter == "Gender_Spouse"]

# Numeric Parameters
retirement_age <- as.numeric(input_data$Value[input_data$Parameter == "Retirement_Age_Staff"])
spouse_age_at_retirement <- as.numeric(input_data$Value[input_data$Parameter == "Spouse_Age_whenstaffretirement"])
market_value_vcs <- as.numeric(input_data$Value[input_data$Parameter == "Market_Value_VCS"])
market_value_fba <- as.numeric(input_data$Value[input_data$Parameter == "Market_Value_FBA"])
spouse_target_pension <- as.numeric(input_data$Value[input_data$Parameter == "spousetargetpension"])
staff_target_pension <- as.numeric(input_data$Value[input_data$Parameter == "staff_target_pension"])
staff_annuity <- as.numeric(input_data$Value[input_data$Parameter == "Staff_annuity"])
spouse_annuity_percentage <- as.numeric(input_data$Value[input_data$Parameter == "Spouse_Annuity_Percentage"])
discount_rate <- as.numeric(input_data$Value[input_data$Parameter == "NBY"])

# New: Additional numeric parameters for fractional months
months_staff <- as.numeric(input_data$Value[input_data$Parameter == "months_staff"])
months_spouse <- as.numeric(input_data$Value[input_data$Parameter == "months_spouse"])

# Override parameters from sliders if available (global OVERRIDE_PARAMS)
if (exists("OVERRIDE_PARAMS", envir = .GlobalEnv)) {
  over <- get("OVERRIDE_PARAMS", envir = .GlobalEnv)
  cat("OVERRIDE_PARAMS detected and applied:\n")
  print(over)
  if (!is.null(over$NBY)) {
    discount_rate <- over$NBY
  }
  if (!is.null(over$SPOUSE_TARGET)) {
    spouse_target_pension <- over$SPOUSE_TARGET
  }
  if (!is.null(over$STAFF_TARGET)) {
    staff_target_pension <- over$STAFF_TARGET
  }
} else {
  cat("OVERRIDE_PARAMS not found. Default values will be used.\n")
}

# Validate the discount rate
if (is.na(discount_rate)) stop("Invalid discount rate (NBY) provided in input data. Please check.")

# Store extracted parameters in a list
parameters <- list(
  Staff_Name = staff_name,
  DOB_Staff = format(staff_dob, "%Y-%m-%d"),
  Gender_Staff = gender_staff,
  DOB_Spouse = format(spouse_dob, "%Y-%m-%d"),
  Gender_Spouse = gender_spouse,
  Retirement_Date_Staff = format(retirement_date, "%Y-%m-%d"),
  Retirement_Age_Staff = retirement_age,
  Spouse_Age_whenstaffretirement = spouse_age_at_retirement,
  months_staff = months_staff,
  months_spouse = months_spouse,
  Market_Value_VCS = market_value_vcs,
  Market_Value_FBA = market_value_fba,
  spousetargetpension = spouse_target_pension,
  staff_target_pension = staff_target_pension,
  Staff_annuity = staff_annuity,
  Spouse_Annuity_Percentage = spouse_annuity_percentage,
  Discount_Rate_NBY = discount_rate
)

cat("\nExtracted Parameters:\n")
print(parameters)

cat("Exporting parameters to Excel...\n")
# Convert the parameters list to a data frame
parameters_df <- data.frame(
  Parameter = names(parameters),
  Value = unlist(parameters, use.names = FALSE),
  stringsAsFactors = FALSE
)

# Define the output path for the parameters Excel file using the unified output_dir
parameters_output_path <- file.path(root_path, config$output_dir, "Extracted_Parameters.xlsx")

# Write the parameters to an Excel file
write.xlsx(parameters_df, file = parameters_output_path, overwrite = TRUE)
cat("\nParameters exported successfully to:", parameters_output_path, "\n")

cat("\nBlock 2 completed successfully. Input data and parameters loaded and exported to Excel.\n")
