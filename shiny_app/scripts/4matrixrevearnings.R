# ------------------------------------------------------------------------------
# Title: Combine Worked Column with Historical and Future Salaries
# Description:
#   - Load the `Worked` column from Personal_Data.xlsx (SalaryHistory sheet)
#   - Merge the `Worked` data with the historical and future salaries
#   - Recalculate the Worked_Salary
#   - Export the final table with updated values
# ------------------------------------------------------------------------------

# Load Required Libraries
library(dplyr)
library(readxl)
library(openxlsx)
library(janitor)
library(yaml)

# ------------------------------------------------------------------------------
# Step 1: Define File Paths
# Description: Specify input and output file paths using the unified YAML configuration.
# ------------------------------------------------------------------------------
cat("Loading YAML configuration...\n")
config <- yaml::read_yaml("config/config.yaml")
personal_data_path <- file.path(config$root_path, config$personal_data_path)
historical_future_salaries_path <- file.path(config$root_path, config$output_dir, "Historical_and_Future_Salaries.xlsx")
output_path <- file.path(config$root_path, config$output_dir, "Final_Salaries_With_Worked.xlsx")

# ------------------------------------------------------------------------------
# Step 2: Load `Worked` Column from Personal_Data.xlsx
# ------------------------------------------------------------------------------
cat("Loading Worked column from Personal_Data.xlsx...\n")
personal_data_file <- file.path(personal_data_path, "Personal_Data.xlsx")
worked_data <- read_excel(personal_data_file, sheet = "SalaryHistory") %>%
  clean_names() %>%
  mutate(
    Year = year(as.Date(date, origin = "1899-12-30")),  # Extract year from the date column
    Worked = as.numeric(gsub(",", ".", worked))         # Convert Worked column to numeric
  ) %>%
  filter(!is.na(Worked)) %>%  # Remove rows where Worked is NA
  select(Year, Worked)        # Keep only Year and Worked columns

cat("Processed Worked data:\n")
print(tail(worked_data))  # Show last rows for verification

# ------------------------------------------------------------------------------
# Step 3: Load Historical and Future Salaries
# ------------------------------------------------------------------------------
cat("Loading Historical and Future Salaries data...\n")
historical_future_salaries <- read_excel(historical_future_salaries_path) %>%
  clean_names() %>%
  mutate(
    Year = as.integer(year),  # Ensure Year column is numeric
    Starting_Salary = as.numeric(gsub(",", ".", starting_salary))  # Convert salary to numeric
  ) %>%
  select(Year, age, Starting_Salary)

cat("Historical and Future Salaries loaded successfully:\n")
print(head(historical_future_salaries))

# ------------------------------------------------------------------------------
# Step 4: Merge Worked Column and Recalculate Worked_Salary
# ------------------------------------------------------------------------------
cat("Merging Worked column and recalculating Worked_Salary...\n")
final_salaries <- historical_future_salaries %>%
  left_join(worked_data, by = "Year") %>%  # Merge Worked data with salaries by Year
  mutate(
    Worked = ifelse(is.na(Worked), 1, Worked),        # Default Worked to 1 if missing
    Worked_Salary = Starting_Salary * 12 * Worked     # Calculate Worked_Salary
  )

cat("Final Salaries with Worked column:\n")
print(head(final_salaries))

# ------------------------------------------------------------------------------
# Step 5: Export the Final Table to Excel
# ------------------------------------------------------------------------------
cat("Exporting the final table to Excel...\n")
wb <- createWorkbook()
addWorksheet(wb, "Final Salaries")
writeData(wb, "Final Salaries", final_salaries)
saveWorkbook(wb, output_path, overwrite = TRUE)
cat("Final table exported successfully to:", output_path, "\n")
