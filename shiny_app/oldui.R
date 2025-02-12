library(shiny)
library(shinydashboard)
library(DT)
library(shinyjs)

ui <- dashboardPage(
  skin = "blue",
  
  dashboardHeader(title = "Transfer-In Credit Calculator", titleWidth = 450),
  
  dashboardSidebar(
    width = 300,
    sidebarMenu(
      menuItem("Instructions", tabName = "instructions", icon = icon("info-circle")),
      menuItem("Main Dashboard", tabName = "dashboard", icon = icon("dashboard"))
    )
  ),
  
  dashboardBody(
    useShinyjs(),  # Include shinyjs for enabling/disabling buttons dynamically
    tabItems(
      
      # Instructions Tab
      tabItem(tabName = "instructions",
              fluidRow(
                box(
                  title = "How to Use This Application", width = 12, status = "primary", solidHeader = TRUE,
                  p("This tool processes Transfer-In Credit calculations using provided Excel files."),
                  tags$ul(
                    tags$li("Download the Personal and Input Data Excel files."),
                    tags$li("Fill in the necessary details."),
                    tags$li("Upload the modified files (Personal and Input Data are mandatory)."),
                    tags$li("Click 'Run Transfer-In Calculation' to refresh results."),
                    tags$li("Download final reports and outputs as needed.")
                  )
                )
              )
      ),
      
      # Main Dashboard Tab
      tabItem(tabName = "dashboard",
              
              fluidRow(
                # Download Buttons (always active)
                box(title = "Download Files", status = "primary", solidHeader = TRUE, width = 12,
                    downloadButton("download_personal", "Download Personal Data"),
                    downloadButton("download_input", "Download Input Data"),
                    downloadButton("download_static", "Download Static Data (Optional)"),
                    br(), br(),
                    downloadButton("download_transfer_info", "Download Transfer-In Information"),
                    downloadButton("download_all_outputs", "Download All Output Files")
                )
              ),
              
              fluidRow(
                # Upload Buttons
                box(title = "Upload Your Data Files", status = "warning", solidHeader = TRUE, width = 12,
                    fileInput("upload_personal", "Upload Modified Personal Data", accept = ".xlsx"),
                    fileInput("upload_input", "Upload Modified Input Data", accept = ".xlsx"),
                    fileInput("upload_static", "Upload Modified Static Data (Optional)", accept = ".xlsx"),
                    shinyjs::disabled(actionButton("run_calculation", "Run Transfer-In Calculation", class = "btn-primary"))
                )
              ),
              
              fluidRow(
                # Personal Information
                box(title = "Personal Details", status = "info", solidHeader = TRUE, width = 6,
                    verbatimTextOutput("staff_name"),
                    verbatimTextOutput("staff_dob"),
                    verbatimTextOutput("spouse_dob"),
                    verbatimTextOutput("retirement_date"),
                    verbatimTextOutput("report_date"),
                    verbatimTextOutput("transfer_in_amount")
                ),
                
                # Revalued Earnings Information
                box(title = "Revalued Earnings Information", status = "info", solidHeader = TRUE, width = 6,
                    verbatimTextOutput("monthly_revalued_earnings_hr"),
                    verbatimTextOutput("annual_revalued_earnings_hr"),
                    verbatimTextOutput("annual_revalued_earnings_hr_capitalised"),
                    verbatimTextOutput("revalued_earnings_discounted"),
                    verbatimTextOutput("revalued_earnings_retirement_age65")
                )
              ),
              
              fluidRow(
                # Final Calculation Results
                box(title = "Transfer-In Credit Results", status = "success", solidHeader = TRUE, width = 12,
                    verbatimTextOutput("transfer_in_credit_calculated"),
                    verbatimTextOutput("transfer_in_credit_provided"),
                    verbatimTextOutput("transfer_in_factor"),
                    verbatimTextOutput("pcf_total"),
                    verbatimTextOutput("pcf_staff"),
                    verbatimTextOutput("pcf_spouse"),
                    verbatimTextOutput("time_to_retirement"),
                    verbatimTextOutput("discount_rate_nby"),
                    verbatimTextOutput("preretirement_discountrate")
                )
              )
      )
    )
  )
)
