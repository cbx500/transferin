# ------------------------------------------------------------------------------
# Title: Transfer-In Credit Calculation Script (Final Version)
# Description:
#   - Extracts required values from generated Excel files.
#   - Computes Scheme Factor and Transfer-In Credit.
#   - Ensures correct extraction of transferred_value and report_date.
#   - Saves key parameters to a separate Transfer-In extracted parameters file.
#   - Ensures correct usage of revalued earnings at retirement.
#   - Exports all key variables to TransferInCreditResults.xlsx.
#   - Displays all parameters at the end using `cat()`.
# ------------------------------------------------------------------------------

# Load required libraries
library(readxl)    
library(openxlsx)  
library(dplyr)     
library(yaml)      

cat("\n🔄 Running Transfer-In Credit Calculation...\n")

# ------------------------------------------------------------------------------
# Step 1: Load Configuration and Define Paths
# ------------------------------------------------------------------------------
cat("\n📂 Loading configuration and setting file paths...\n")

# Load YAML configuration
config <- yaml::read_yaml(file.path(getwd(), "config", "config.yaml"))
root_path <- normalizePath(config$root_path, winslash = "/", mustWork = TRUE)

# 1) Define file paths for the files we MUST read:
files <- list(
  # This is the original Extracted_Parameters.xlsx (created by the first script).
  extracted_params = file.path(root_path, config$extracted_params_path),
  
  monthly_validation = file.path(root_path, config$output_dir, "MonthlyEquivalentAnnuityAndPCFCalculation_Validation.xlsx"),
  step1_data         = file.path(root_path, config$output_dir, "Step1_Loaded_Data.xlsx"),
  pension_report     = file.path(root_path, config$output_dir, "Pension_Report.xlsx")
)

# 2) Check existence of the files we read
for (name in names(files)) {
  if (!file.exists(files[[name]])) {
    stop(paste("❌ Error:", name, "file not found:", files[[name]]))
  }
}

# 3) Define a separate path for the NEW extracted parameters file 
#    (so we do NOT overwrite the original).
transfer_in_params <- file.path(root_path, config$transfer_in_extracted_params_path)

# Print paths
cat("✅ Original Extracted Parameters File (read):", files$extracted_params, "\n")
cat("✅ Monthly Validation File:", files$monthly_validation, "\n")
cat("✅ Step 1 Data File:", files$step1_data, "\n")
cat("✅ Pension Report File:", files$pension_report, "\n")
cat("✅ Transfer-In Extracted Parameters File (write):", transfer_in_params, "\n")

# ------------------------------------------------------------------------------
# Step 2: Load Required Data
# ------------------------------------------------------------------------------
cat("\n📖 Extracting required values from Excel files...\n")

# ------------------------------------------------------------------------------
# Step 2.1: Extract Pre-Retirement Discount Rate
# ------------------------------------------------------------------------------
cat("\n📖 Extracting Pre-Retirement Discount Rate from Step1_Loaded_Data.xlsx...\n")

static_data <- read.xlsx(files$step1_data, sheet = "Static Data")
pre_ret_rate <- as.numeric(static_data$Value[static_data$Parameter == "preretirement_discountrate"])
pre_ret_rate <- ifelse(is.na(pre_ret_rate), "MISSING", pre_ret_rate)

cat("✅ Pre-Retirement Discount Rate:", pre_ret_rate, "\n")

# ------------------------------------------------------------------------------
# Step 2.2: Extract NBY (Discount Rate)
# ------------------------------------------------------------------------------
cat("\n📖 Extracting Effective NBY from Extracted_Parameters.xlsx...\n")

params <- read.xlsx(files$extracted_params, sheet = 1)
effective_nby <- as.numeric(params$Value[params$Parameter == "Discount_Rate_NBY"])
effective_nby <- ifelse(is.na(effective_nby), "MISSING", effective_nby)

cat("✅ Effective NBY:", effective_nby, "\n")

# ------------------------------------------------------------------------------
# Step 2.3: Extract PCF Values
# ------------------------------------------------------------------------------
cat("\n📖 Extracting PCF values from MonthlyEquivalentAnnuityAndPCFCalculation_Validation.xlsx...\n")

pcf_df <- read.xlsx(files$monthly_validation, sheet = "Validation_Results")

total_pcf <- as.numeric(pcf_df$Value[pcf_df$Parameter == "Total_Pension_Conversion_Factor (PCF)"])
pcf_staff <- as.numeric(pcf_df$Value[pcf_df$Parameter == "PCF_weighted_staff"])
pcf_spouse <- as.numeric(pcf_df$Value[pcf_df$Parameter == "PCF_weighted_spouse"])

cat("✅ PCF Total:", total_pcf, "\n")
cat("✅ PCF Staff:", pcf_staff, "\n")
cat("✅ PCF Spouse:", pcf_spouse, "\n")

# ------------------------------------------------------------------------------
# Step 2.4: Extract Transferred Value and Report Date
# ------------------------------------------------------------------------------
cat("\n📖 Extracting Transferred Value and Report Date from Step1_Loaded_Data.xlsx...\n")

step1_df <- read.xlsx(files$step1_data, sheet = "Personal Data")

transferred_value <- as.numeric(step1_df$Value[step1_df$Field == "transferred_value"])
report_date <- as.character(step1_df$Value[step1_df$Field == "report_date"])
retirement_date <- as.character(step1_df$Value[step1_df$Field == "retirement_date"])

cat("✅ Transferred Value:", transferred_value, "\n")
cat("✅ Report Date:", report_date, "\n")
cat("✅ Retirement Date:", retirement_date, "\n")

# ------------------------------------------------------------------------------
# Step 2.5: Extract Revalued Earnings from Pension Report
# ------------------------------------------------------------------------------
cat("\n📖 Extracting Revalued Earnings from Pension_Report.xlsx...\n")

pr_df <- read.xlsx(files$pension_report, sheet = "Pension Report")
row65 <- pr_df[pr_df$retirement_age == 65, ]

revalued_earnings <- as.numeric(row65$avg_revalued_earnings_at_retirement)
revalued_earnings_report_date <- as.numeric(row65$avg_revalued_earnings_at_report_date)

cat("✅ Revalued Earnings at Retirement Age 65:", revalued_earnings, "\n")
cat("✅ Revalued Earnings Discounted to Report Date:", revalued_earnings_report_date, "\n")

# ------------------------------------------------------------------------------
# Step 3: Calculate Time to Retirement (TTR)
# ------------------------------------------------------------------------------
cat("\n📆 Calculating Time to Retirement (TTR)...\n")

ttr <- as.numeric(difftime(as.Date(retirement_date), as.Date(report_date), units = "days")) / 365.25
cat("✅ Time to Retirement (TTR):", round(ttr, 2), "years\n")

# ------------------------------------------------------------------------------
# Step 3.1: Fetch revalued_earningsHR and inflation_rate, then Compute Derived Values
# ------------------------------------------------------------------------------
cat("\n📖 Extracting revalued_earningsHR and inflation_rate from Step1_Loaded_Data.xlsx...\n")

# Load the Personal Data sheet dynamically from YAML
step1_df <- read.xlsx(file.path(root_path, config$step1_data_path), 
                      sheet = config$step1_data_sheets$personal_data)


# Extract revalued_earningsHR
revalued_earningsHR <- as.numeric(step1_df$Value[step1_df$Field == "revalued_earningsHR"])
if (is.na(revalued_earningsHR)) {
  stop("❌ Error: revalued_earningsHR value is missing!")
}

cat("✅ Revalued Earnings HR:", revalued_earningsHR, "\n")

# Extract inflation_rate (from static_data)
inflation_rate <- as.numeric(
  static_data %>% 
    filter(Parameter == "inflation_rate") %>% 
    pull(Value)
)

if (is.na(inflation_rate)) {
  stop("❌ Error: inflation_rate value is missing!")
}

cat("✅ Inflation Rate:", inflation_rate, "\n")

# Calculate derived values
annual_revalued_earnigsHr    <- revalued_earningsHR * 12
annual_revalued_earnigsHrRC  <- annual_revalued_earnigsHr * ((1 + inflation_rate) ^ ttr)

cat("✅ Annual Revalued Earnings HR:", annual_revalued_earnigsHr, "\n")
cat("✅ Annual Revalued Earnings HR Capitalised (RC):", annual_revalued_earnigsHrRC, "\n")

# ------------------------------------------------------------------------------
# Step 4: Compute Scheme Factor and Transfer-In Credit
# ------------------------------------------------------------------------------
cat("\n📊 Performing final calculations...\n")

scheme_factor <- ((1 + pre_ret_rate) ^ (-ttr)) * total_pcf * 0.02
transfer_in_credit <- transferred_value / (scheme_factor * revalued_earnings)

cat("✅ Scheme Factor:", round(scheme_factor, 4), "\n")
cat("✅ Transfer-In Credit:", round(transfer_in_credit, 2), "\n")

# ------------------------------------------------------------------------------
# Step 4.1: Perform Final Calculations Including HR RE
# ------------------------------------------------------------------------------
cat("\n📊 Performing final with HR RE calculations...\n")

scheme_factor <- ((1 + pre_ret_rate) ^ (-ttr)) * total_pcf * 0.02
transfer_in_creditHRRE <- transferred_value / (scheme_factor * annual_revalued_earnigsHrRC)

cat("✅ Scheme Factor:", round(scheme_factor, 4), "\n")
cat("✅ Transfer-In Credit HRRE:", round(transfer_in_creditHRRE, 2), "\n")

# ------------------------------------------------------------------------------
# Step 5: Export Results to a NEW Extracted Parameters file
# ------------------------------------------------------------------------------
params_df <- data.frame(
  Parameter = c("report_date", "retirement_date", "transferred_value",
                "preretirement_discountrate", "Discount_Rate_NBY",  
                "revalued_earnings", "revalued_earnings_report_date",
                "revalued_earningsHR", "annual_revalued_earnigsHr", 
                "annual_revalued_earnigsHrRC", "inflation_rate"),
  Value = c(
    report_date, retirement_date, transferred_value,
    pre_ret_rate, effective_nby,
    revalued_earnings, revalued_earnings_report_date,
    revalued_earningsHR, annual_revalued_earnigsHr,
    annual_revalued_earnigsHrRC, inflation_rate
  ),
  stringsAsFactors = FALSE
)

cat("\n🔍 Checking params_df Before Saving to Extracted_Parameters_TransferIn.xlsx:\n")
print(params_df)

# Append or update extracted parameters
params <- rbind(params[!params$Parameter %in% params_df$Parameter, ], params_df)

# Save to the new Transfer-In file (DO NOT overwrite the original)
write.xlsx(params, transfer_in_params, overwrite = TRUE)
cat("\n✅ Transfer-In Extracted Parameters File Updated Successfully!\n")

# ------------------------------------------------------------------------------
# Step 6: Export Final Results to Excel (TransferInCreditResults.xlsx)
# ------------------------------------------------------------------------------
output_path <- file.path(root_path, config$output_dir, "TransferInCreditResults.xlsx")
wb <- createWorkbook()
addWorksheet(wb, "Results")

# 🔸 We ADD two more lines to the Parameters vector:
results_df <- data.frame(
  Parameter = c("Time to Retirement (TTR)",
                "PCF Total",
                "PCF Staff",
                "PCF Spouse",
                "Scheme Factor",
                "Transfer-In Credit",
                "Transfer-In Credit HRRE",
                "Annual Revalued Earnings HR",
                "Annual Revalued Earnings HR Capitalised",
                # Add these two new lines:
                "Revalued Earnings at Retirement Age 65",
                "Revalued Earnings Discounted to Report Date"
  ),
  Value = c(round(ttr, 2),
            total_pcf,
            pcf_staff,
            pcf_spouse,
            round(scheme_factor, 4),
            round(transfer_in_credit, 2),
            round(transfer_in_creditHRRE, 2),
            round(annual_revalued_earnigsHr, 2),
            round(annual_revalued_earnigsHrRC, 2),
            # Add corresponding values:
            round(revalued_earnings, 2),
            round(revalued_earnings_report_date, 2)
  )
)

writeData(wb, "Results", results_df)
saveWorkbook(wb, output_path, overwrite = TRUE)

cat("\n✅ Results saved successfully to:", output_path, "\n")
cat("\n🎉 Transfer-In Credit Calculation Completed Successfully!\n")
