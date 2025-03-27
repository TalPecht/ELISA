library(rmarkdown)
library(readxl)
library(tools)
library(htmltools)
library(shiny)
library(rstudioapi)

#### Read each excel file and perform analysis with the .rmd ##### 

# Get the directory where the script is saved
dir <- dirname(getActiveDocumentContext()$path)

# Get all Excel files in the 'Data' folder
files <- list.files(path = paste0(dir, "/Data/"), pattern = "\\.xlsx$", full.names = TRUE)

# Create the 'reports/' folder if it does not exist
if (!dir.exists(paste0(dir, "/reports"))) {
  dir.create(paste0(dir, "/reports"))
}

# Initialize an empty vector to store output file paths
html_files <- c()

# Loop through each Excel file
for (file in files) {
  
  # Extract the file name without extension
  file_name <- file_path_sans_ext(basename(file))
  
  # Define output file path based on Excel file name
  output_file <- paste0(dir, "/reports/", file_name, ".html")
  
  # Set dynamic title: "ELISA Analysis - [File Name]"
  dynamic_title <- paste("ELISA Analysis -", file_name)
  
  # Render the RMarkdown report, passing the dynamic title
  render("C:/Users/pachtt/Documents/Femmunity/ELISA/calculations/250307_calculate_ELISA_values.Rmd", 
         params = list(excel_file = file, title = dynamic_title), 
         output_file = output_file)
  
  # Store generated HTML report path
  html_files <- c(html_files, output_file)
}

# Save list of generated HTML files
writeLines(html_files, file.path(dir, "reports/generated_reports.txt"))

#####Merge Reports into One HTML File with Tabs #####
# Get the directory where the script is saved

# Read the list of generated HTML reports
html_files <- readLines(file.path(dir, "reports/generated_reports.txt"))

# Create a list of tabPanels with `tags$iframe()`
tabs <- lapply(html_files, function(file) {
  file_name <- file_path_sans_ext(basename(file))  # Extract file name
  tabPanel(
    title = file_name, 
    tags$iframe(src = file, width = "100%", height = "800px")  # Embed the full HTML report
  )
})

# Ensure tabs contain valid `tabPanel()` elements
if (length(tabs) == 0) {
  stop("Error: No valid tab panels were created. Check your file paths and ensure reports exist.")
}

# Wrap all tabs inside a dropdown menu
dropdown_menu <- navbarMenu("Reports", !!!tabs)

# Generate final HTML with tabs
final_report <- tagList(
  navbarPage("Combined ELISA Report Femmunity TruCulture LPS IL6", dropdown_menu)  # Use `!!!` to correctly insert the list
)


# Save the merged HTML
save_html(final_report, file.path(dir, "reports/",paste0(Sys.Date(),"IL6_final_report.html")))

message("Final merged report created at: ", file.path(dir, "reports/", paste0(Sys.Date(), "IL6_final_report.html")))


