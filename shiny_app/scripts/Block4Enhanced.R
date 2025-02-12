# Block 4: Generate Cash Flow Sheets for Annuity Calculations
# Description:
# This script generates cash flow sheets for calculating the present value of post-paid annuities 
# for the staff member and their spouse. Two scenarios are calculated:
#   - Scenario T: The baseline scenario where the staff member retires at the given retirement age 
#     (e.g., 65 in 2028) and cash flows are generated starting at that age.
#   - Scenario T+1: A one-year shift scenario where the staff member is assumed to be one year older 
#     at retirement (e.g., 66 in 2028) and cash flows are generated starting at that shifted age.
# These dual calculations provide two sets of cash flow data that can later be weighted/interpolated 
# to accurately value annuities for individuals whose retirement ages are fractional.
#
# The script dynamically loads cumulative survival and death probability sheets (for both T and T+1)
# using gender-selected mortality tables, applies the preloaded parameters (including discount rate 
# and annuity amounts), and produces an Excel workbook with separate worksheets for staff and spouse 
# cash flows under both scenarios.

library(openxlsx)
library(dplyr)
library(yaml)

cat("Loading configuration...\n")
# Load the YAML configuration file
config_path <- file.path(getwd(), "config", "config.yaml")
if (!file.exists(config_path)) {
  stop("Configuration file not found at:", config_path)
}
config <- yaml::read_yaml(config_path)
cat("Configuration loaded successfully from:", config_path, "\n")

# Use the unified root_path from YAML (assumed to be absolute)
root_path <- normalizePath(config$root_path, winslash = "/", mustWork = TRUE)

# Construct necessary paths using the unified configuration
cumulative_probabilities_path <- file.path(root_path, config$cumulative_probabilities_path)
cash_flow_output_path <- file.path(root_path, config$cash_flow_output_path)

cat("Using the following paths:\n")
cat(" - Cumulative Probabilities Path:", cumulative_probabilities_path, "\n")
cat(" - Cash Flow Output Path:", cash_flow_output_path, "\n")

# ---------------------------
# 2. Load Cumulative Probability Data (for T and T+1)
# ---------------------------
cat("Loading cumulative survival and death probability data...\n")
# For the staff member (assumed male) – Scenario T and T+1
cumulative_survival_male_T  <- read.xlsx(cumulative_probabilities_path, sheet = "Staff_Cumulative_Survival_T")
cumulative_survival_male_T1 <- read.xlsx(cumulative_probabilities_path, sheet = "Staff_Cumulative_Survival_T1")
Staff_Cumulative_Death_T  <- read.xlsx(cumulative_probabilities_path, sheet = "Staff_Cumulative_Death_T")
Staff_Cumulative_Death_T1 <- read.xlsx(cumulative_probabilities_path, sheet = "Staff_Cumulative_Death_T1")

# For the spouse (assumed female) – Scenario T and T+1
cumulative_survival_female_T  <- read.xlsx(cumulative_probabilities_path, sheet = "Spouse_Cumulative_Survival_T")
cumulative_survival_female_T1 <- read.xlsx(cumulative_probabilities_path, sheet = "Spouse_Cumulative_Survival_T1")
Spouse_Cumulative_Death_T  <- read.xlsx(cumulative_probabilities_path, sheet = "Spouse_Cumulative_Death_T")
Spouse_Cumulative_Death_T1 <- read.xlsx(cumulative_probabilities_path, sheet = "Spouse_Cumulative_Death_T1")

# ---------------------------
# 3. Use Preloaded Parameters
# ---------------------------
cat("Using preloaded parameters...\n")
# (Assumes that the variable 'parameters' is available from Block 2)
staff_annuity <- as.numeric(parameters$Staff_annuity)
spouse_annuity_percentage <- as.numeric(parameters$Spouse_Annuity_Percentage)
spouse_dob <- as.Date(parameters$DOB_Spouse)
retirement_date <- as.Date(parameters$Retirement_Date_Staff)
discount_rate <- as.numeric(parameters$Discount_Rate_NBY)
if (exists("OVERRIDE_PARAMS", envir = .GlobalEnv)) {
  over <- get("OVERRIDE_PARAMS", envir = .GlobalEnv)
  if (!is.null(over$NBY)) {
    discount_rate <- over$NBY
  }
}

# ---------------------------
# 4. Define Helper Function for Discounting
# ---------------------------
calculate_discount_factor <- function(year_index, discount_rate) {
  # For post-paid annuities, discount one extra period (hence +1)
  (1 / (1 + discount_rate)) ^ (year_index + 1)
}

# ---------------------------
# 5. Generate Cash Flow Data Functions
# ---------------------------
# (a) For Staff Member
generate_cash_flow_staff_T <- function() {
  cat("Generating cash flow data for staff at T...\n")
  start_row <- which(cumulative_survival_male_T$Age == parameters$Retirement_Age_Staff)
  if (length(start_row) == 0) stop("Error: Retirement age not found in staff cumulative survival table for T.")
  num_years <- nrow(cumulative_survival_male_T) - start_row + 1
  
  cash_flows <- data.frame(
    Year = cumulative_survival_male_T$Year[start_row:(start_row + num_years - 1)],
    Age  = cumulative_survival_male_T$Age[start_row:(start_row + num_years - 1)],
    Cash_Flow = staff_annuity,
    Cumulative_Survival_Prob = cumulative_survival_male_T$Cumulative_Survival_Prob[start_row:(start_row + num_years - 1)],
    Discount_Factor = sapply(0:(num_years - 1), calculate_discount_factor, discount_rate = discount_rate)
  )
  cash_flows$Product <- cash_flows$Cash_Flow * cash_flows$Cumulative_Survival_Prob * cash_flows$Discount_Factor
  total_product <- sum(cash_flows$Product, na.rm = TRUE)
  cash_flows <- rbind(cash_flows, data.frame(Year = "Total", Age = NA, Cash_Flow = NA,
                                             Cumulative_Survival_Prob = NA, Discount_Factor = NA,
                                             Product = total_product))
  return(cash_flows)
}

generate_cash_flow_staff_T1 <- function() {
  cat("Generating cash flow data for staff at T+1...\n")
  start_row <- which(cumulative_survival_male_T1$Age == (parameters$Retirement_Age_Staff + 1))
  if (length(start_row) == 0) stop("Error: Staff's retirement age (T+1) not found in cumulative survival table for T+1.")
  num_years <- nrow(cumulative_survival_male_T1) - start_row + 1
  
  cash_flows <- data.frame(
    Year = cumulative_survival_male_T1$Year[start_row:(start_row + num_years - 1)],
    Age  = cumulative_survival_male_T1$Age[start_row:(start_row + num_years - 1)],
    Cash_Flow = staff_annuity,
    Cumulative_Survival_Prob = cumulative_survival_male_T1$Cumulative_Survival_Prob[start_row:(start_row + num_years - 1)],
    Discount_Factor = sapply(0:(num_years - 1), calculate_discount_factor, discount_rate = discount_rate)
  )
  cash_flows$Product <- cash_flows$Cash_Flow * cash_flows$Cumulative_Survival_Prob * cash_flows$Discount_Factor
  total_product <- sum(cash_flows$Product, na.rm = TRUE)
  cash_flows <- rbind(cash_flows, data.frame(Year = "Total", Age = NA, Cash_Flow = NA,
                                             Cumulative_Survival_Prob = NA, Discount_Factor = NA,
                                             Product = total_product))
  return(cash_flows)
}

# (b) For Spouse
generate_cash_flow_spouse_T <- function() {
  cat("Generating cash flow data for spouse at T...\n")
  spouse_ret_age <- as.integer(difftime(retirement_date, spouse_dob, units = "days") / 365.25)
  start_row <- which(cumulative_survival_female_T$Age == spouse_ret_age)
  if (length(start_row) == 0) stop("Error: Spouse's retirement age not found in cumulative survival table for T.")
  num_years <- nrow(cumulative_survival_female_T) - start_row + 1
  
  cash_flows <- data.frame(
    Year = cumulative_survival_female_T$Year[start_row:(start_row + num_years - 1)],
    Spouse_Age = cumulative_survival_female_T$Age[start_row:(start_row + num_years - 1)],
    Cash_Flow = staff_annuity * spouse_annuity_percentage,
    Spouse_Cumulative_Survival_Prob = cumulative_survival_female_T$Cumulative_Survival_Prob[start_row:(start_row + num_years - 1)],
    Staff_Cumulative_Mortality_Prob = Staff_Cumulative_Death_T$Cumulative_Death_Prob[start_row:(start_row + num_years - 1)],
    Discount_Factor = sapply(0:(num_years - 1), calculate_discount_factor, discount_rate = discount_rate)
  )
  cash_flows$Product <- cash_flows$Cash_Flow * cash_flows$Spouse_Cumulative_Survival_Prob *
    cash_flows$Staff_Cumulative_Mortality_Prob * cash_flows$Discount_Factor
  total_product <- sum(cash_flows$Product, na.rm = TRUE)
  cash_flows <- rbind(cash_flows, data.frame(Year = "Total", Spouse_Age = NA, Cash_Flow = NA,
                                             Spouse_Cumulative_Survival_Prob = NA, Staff_Cumulative_Mortality_Prob = NA,
                                             Discount_Factor = NA, Product = total_product))
  return(cash_flows)
}

generate_cash_flow_spouse_T1 <- function() {
  cat("Generating cash flow data for spouse at T+1...\n")
  spouse_ret_age_T1 <- as.integer(difftime(retirement_date, spouse_dob, units = "days") / 365.25) + 1
  start_row <- which(cumulative_survival_female_T1$Age == spouse_ret_age_T1)
  if (length(start_row) == 0) stop("Error: Spouse's retirement age (T+1) not found in cumulative survival table for T+1.")
  num_years <- nrow(cumulative_survival_female_T1) - start_row + 1
  
  cash_flows <- data.frame(
    Year = cumulative_survival_female_T1$Year[start_row:(start_row + num_years - 1)],
    Spouse_Age = cumulative_survival_female_T1$Age[start_row:(start_row + num_years - 1)],
    Cash_Flow = staff_annuity * spouse_annuity_percentage,
    Spouse_Cumulative_Survival_Prob = cumulative_survival_female_T1$Cumulative_Survival_Prob[start_row:(start_row + num_years - 1)],
    Staff_Cumulative_Mortality_Prob = Staff_Cumulative_Death_T1$Cumulative_Death_Prob[start_row:(start_row + num_years - 1)],
    Discount_Factor = sapply(0:(num_years - 1), calculate_discount_factor, discount_rate = discount_rate)
  )
  cash_flows$Product <- cash_flows$Cash_Flow * cash_flows$Spouse_Cumulative_Survival_Prob *
    cash_flows$Staff_Cumulative_Mortality_Prob * cash_flows$Discount_Factor
  total_product <- sum(cash_flows$Product, na.rm = TRUE)
  cash_flows <- rbind(cash_flows, data.frame(Year = "Total", Spouse_Age = NA, Cash_Flow = NA,
                                             Spouse_Cumulative_Survival_Prob = NA, Staff_Cumulative_Mortality_Prob = NA,
                                             Discount_Factor = NA, Product = total_product))
  return(cash_flows)
}

# ---------------------------
# 6. Generate Cash Flow Data for All Scenarios and Save to Excel
# ---------------------------
staff_cash_flows_T  <- generate_cash_flow_staff_T()
staff_cash_flows_T1 <- generate_cash_flow_staff_T1()
spouse_cash_flows_T  <- generate_cash_flow_spouse_T()
spouse_cash_flows_T1 <- generate_cash_flow_spouse_T1()

write.xlsx(list(
  "Cumulative_Cash_Flow_Staff_T"   = staff_cash_flows_T,
  "Cumulative_Cash_Flow_Staff_T1"  = staff_cash_flows_T1,
  "Cumulative_Cash_Flow_Spouse_T"  = spouse_cash_flows_T,
  "Cumulative_Cash_Flow_Spouse_T1" = spouse_cash_flows_T1
), file = cash_flow_output_path, overwrite = TRUE)

cat("Cash flow sheets for annuity calculation generated and saved successfully at", cash_flow_output_path, "\n")
