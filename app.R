# TrichAnalytics Shiny App: Hair S Corrections Button
# Created as add-on for existing Shiny App 
# Created 8 Aug 2023 KMill 


# Load Libraries ----
library(shiny)
library(shinyFiles)
library(tidyverse)
library(openxlsx)


# Define UI for application ----
ui <- fluidPage(
    fileInput("file", label = NULL),
    downloadButton(outputId = "download_file",
                        label = "Download Data File",
                        icon = icon("table")), 
    numericInput("scorr", label = h5("Enter Sulphur Correction"), value = 40000)
)


# Define server logic ----
server <- function(input, output) {
  
  options(shiny.maxRequestSize = 30 * 1024^2) #increases max size of file that can be uploaded
  
  ## Download File button ----
  
  output$download_file <- downloadHandler(
    
    filename = function() {
      paste("S_Corrected_", input$file, sep = "")
    },
    
    content = function(file) {

      # Read workbook and name 'file_corr' 
      file_corr <- loadWorkbook(input$file$datapath)
      cloneWorksheet(file_corr, "Hair Results_S Corr", "Hair Results")
  
      # Create dataframes for intermediate steps; to allow processed data to be copied to original worksheet and retain original formatting. 
      res_orig <- readxl::read_excel(input$file$datapath, sheet = "Hair Results") 
      res_corr <- readxl::read_excel(input$file$datapath, sheet = "Hair Results")
      
      # Create objects corresponding to dataframe indices
      s_conc <- which(res_orig[,1] == "S") # Sulphate row number
      first_row <- which(res_orig[,4] == "mg/kg") + 1 # Dataframe first row 
      last_row <- nrow(res_orig[,4]) # Dataframe last row
      
      # Correct based on sulphur
      for (i in 4:ncol(res_orig)) {
        for (j in first_row:last_row) {
          res_corr[[i]] <- as.numeric(res_corr[[i]]) # Convert column to numeric
          res_orig[[i]] <- as.numeric(res_orig[[i]]) # Convert column to numeric
          res_corr[j,i] <- res_orig[j,i] * input$scorr/res_orig[s_conc, i] # Sulphur correction
        }
      }
      
      # Write processed data to duplicated hair result tab (retaining formatting)
      writeData(file_corr, 
                sheet = "Hair Results_S Corr", 
                res_corr[first_row:last_row, 4:ncol(res_orig)], 
                startCol = 4, startRow = first_row + 2, 
                colNames = FALSE)
      
      # Add footnotes for conditional formatting 
      writeData(file_corr, 
                sheet = "Hair Results_S Corr", 
                "Value is below reference range.",
                startCol = 2, last_row + 3, 
                colNames = FALSE)
      
      writeData(file_corr, 
                sheet = "Hair Results_S Corr", 
                "Value is above reference range.",
                startCol = 2, last_row + 4, 
                colNames = FALSE)
      
      writeData(file_corr, 
                sheet = "Hair Results_S Corr", 
                "Blue = ",
                startCol = 1, last_row + 3, 
                colNames = FALSE)
      
      writeData(file_corr, 
                sheet = "Hair Results_S Corr", 
                "Orange = ",
                startCol = 1, last_row + 4, 
                colNames = FALSE)
      
      addStyle(file_corr, 
               sheet = "Hair Results_S Corr",
               style = createStyle (fgFill = "#3c52b5", fontSize = 9), 
               rows = last_row + 3, 
               cols = 1)
      
      addStyle(file_corr, 
               sheet = "Hair Results_S Corr",
               style = createStyle (fgFill = "#c97814", fontSize = 9), 
               rows = last_row + 4, 
               cols = 1)
      
      addStyle(file_corr, 
               sheet = "Hair Results_S Corr",
               style = createStyle (fontSize = 9, halign = "left"), 
               rows = last_row + 3, 
               cols = 2)
      
      addStyle(file_corr, 
               sheet = "Hair Results_S Corr",
               style = createStyle (fontSize = 9, halign = "left"), 
               rows = last_row + 4, 
               cols = 2)
      
      # If below reference, cell = blue
      conditionalFormatting(file_corr,
                            sheet = "Hair Results_S Corr", 
                            cols = 4:ncol(res_orig), 
                            rows = (first_row + 2):last_row, 
                            rule = "<$B1", 
                            style = createStyle(bgFill = "#3c52b5"), 
                            type = "expression")
      
      # If above reference, cell = orange
      conditionalFormatting(file_corr,
                            sheet = "Hair Results_S Corr", 
                            cols = 4:ncol(res_orig), 
                            rows = (first_row+2):last_row, 
                            rule = ">$C1", 
                            style = createStyle(bgFill = "#c97814"),  
                            type = "expression")
      
      
      # If cell is blank, background is blank 
      conditionalFormatting(file_corr,
                            sheet = "Hair Results_S Corr", 
                            cols = 4:ncol(res_orig), 
                            rows = (first_row+2):last_row, 
                            rule = "=0", 
                            style = createStyle(bgFill = "#ffffff"),  
                            type = "expression")
    
      
      # Save workbook
      saveWorkbook(file_corr, file)
    })}

# Run the application ----
shinyApp(ui = ui, server = server)

