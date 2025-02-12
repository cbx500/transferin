# ------------------------------------------------------------------------------
# Name: Master Pipeline Runner Script
# Description:
#   This script sequentially runs 12 individual R scripts that comprise the 
#   calculation pipeline for the Transfer-In Credit Calculator. 
#   It ensures that fresh data is processed each time the pipeline is run.
#   If any script encounters an error, the process stops and an error message is displayed.
#
# Usage:
#   Place this file in your "scripts" directory. When the user presses "Run Calculations"
#   in the Shiny app, this script will be sourced, executing all 12 scripts in order.
# ------------------------------------------------------------------------------
# List of scripts to run (adjust the names if necessary)
scripts_to_run <- c(
  "1.loading.R",
  "2salaryincreases.R",
  "3salaryvector.R",
  "4matrixrevearnings.R",
  "5Arevaluationmatrix.R",
  "6lastgoodeventually.R",
  "reportaaaa.R",  # 🟢 Moved up to ensure the pension report is available
  "Block1.R",
  "Block2.R",
  "Block3.R",
  "Block4Enhanced.R",
  "Block6PCFs.R",
  "calculate_transfer_in_credit.R"  # 🟢 Runs last to use all final values
)


# Get the scripts directory (assumes getwd() returns the app's root directory)
scripts_dir <- file.path(getwd(), "scripts")

# Loop over each script and source it
for (script_name in scripts_to_run) {
  script_path <- file.path(scripts_dir, script_name)
  
  if (!file.exists(script_path)) {
    stop("Script not found: ", script_path)
  }
  
  cat("Running: ", script_name, "\n")
  
  tryCatch({
    source(script_path, local = TRUE)
    cat(script_name, "completed successfully.\n")
  }, error = function(e) {
    stop("Error running ", script_name, ": ", e$message)
  })
}

cat("All 13 scripts executed successfully.\n")
