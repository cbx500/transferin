# Block 1: Setup and Package Initialization

# Description:
# This block initializes the environment by ensuring that all necessary packages are installed and loaded.
# It dynamically resolves paths using config.yaml and supports both local and Shiny deployments.

# List of required packages across all blocks
required_packages <- c("officer", "flextable", "yaml", "openxlsx", 
                       "dplyr", "readxl", "stringr", "ggplot2", "lifecycle")

# Function to install and load packages
install_and_load <- function(packages) {
  for (pkg in packages) {
    if (!requireNamespace(pkg, quietly = TRUE)) {
      install.packages(pkg)
      cat(paste("Installed missing package:", pkg, "\n"))
    }
    suppressPackageStartupMessages(library(pkg, character.only = TRUE))
    cat(paste("Loaded package:", pkg, "\n"))
  }
}

# Run the function to install and load all required packages
install_and_load(required_packages)

# Load configuration file
cat("Loading configuration...\n")
config_path <- file.path(getwd(), "config/config.yaml")  # Relative path to config.yaml in shiny_app
if (file.exists(config_path)) {
  config <- yaml::yaml.load_file(config_path)
  cat("Configuration loaded successfully from:", config_path, "\n")
} else {
  stop("Configuration file not found. Please check the path.")
}

# Determine root path dynamically
root_path <- if (!is.null(config$root_path)) {
  file.path(getwd(), config$root_path)  # Use root_path from config
} else {
  getwd()  # Use current working directory
}

# Ensure root_path exists
if (!dir.exists(root_path)) {
  stop("Root path does not exist: ", root_path)
} else {
  cat("Root path in use:", root_path, "\n")
}

# Dynamically construct paths relative to root_path
input_data_path <- file.path(root_path, config$input_data_path)
male_mortality_table_path <- file.path(root_path, config$male_mortality_table_path)
female_mortality_table_path <- file.path(root_path, config$female_mortality_table_path)
mortality_output_path <- file.path(root_path, config$mortality_output_path)
annuity_output_directory <- file.path(root_path, config$annuity_output_directory)
annuity_report_file <- file.path(root_path, config$annuity_report_file)
final_report_path <- file.path(root_path, config$final_report_path)
cumulative_probabilities_path <- file.path(root_path, config$cumulative_probabilities_path)
cash_flow_output_path <- file.path(root_path, config$cash_flow_output_path)
validation_results_path <- file.path(root_path, config$validation_results_path)

# Debug: Display constructed paths
cat("Paths loaded:\n")
cat(" - Input Data Path:", input_data_path, "\n")
cat(" - Male Mortality Table Path:", male_mortality_table_path, "\n")
cat(" - Female Mortality Table Path:", female_mortality_table_path, "\n")
cat(" - Mortality Output Path:", mortality_output_path, "\n")
cat(" - Annuity Output Directory:", annuity_output_directory, "\n")
cat(" - Annuity Report File:", annuity_report_file, "\n")
cat(" - Final Report Path:", final_report_path, "\n")
cat(" - Cumulative Probabilities Path:", cumulative_probabilities_path, "\n")
cat(" - Cash Flow Output Path:", cash_flow_output_path, "\n")
cat(" - Validation Results Path:", validation_results_path, "\n")

# Confirm setup completion
cat("Environment setup is complete. All required packages are installed and loaded.\n")
