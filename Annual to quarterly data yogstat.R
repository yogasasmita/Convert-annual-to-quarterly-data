rm(list=ls())  # remove variables
graphics.off() # close figures

#install.packages("tempdisagg")
library(tempdisagg)
#Import dataset
#install.packages("readxl")
library(readxl)
data <- read_excel("E:/@Yoga Sasmita - ITS/Disertasi/Data/PDRB/PDB (2000-2022)/PDB Seri 2010 (Tahunan 2010 2022).xlsx")
#View(data)

# Making the function for processing data
process_data <- function(row_index) {
  pdb <- as.matrix(data[row_index, 2:14], nrow = 1, ncol = 13, bycol = FALSE)
  pdb <- t(pdb)
  pdb.ts <- ts(pdb, start = 2010, end = 2022)
  pdb.q <- td(pdb.ts ~ 1, to = "quarterly", method = "uniform")
  return(matrix(predict(pdb.q), nrow = 1))
}

# Processing data for all sektor
row_indices <- 1:65  # Define the row indices you want to process
pdb.adhb <- t(sapply(row_indices, function(row_index) process_data(row_index), simplify = "matrix"))

# Load the openxlsx library
#install.packages("openxlsx")
library(openxlsx)

# Specify the file path and name for the Excel file
file_path <- "E:/@Yoga Sasmita - ITS/Disertasi/Data/PDRB/PDB (2000-2022)/PDB quarterly (2010-2022).xlsx"

# Create a data frame from the output
output_df <- as.data.frame(pdb.adhb)

# Load the existing Excel file
wb <- loadWorkbook(file_path)

# Add a new sheet with the output data
addWorksheet(wb, sheetName = "uniform")  # Change "NewSheetName" to your desired sheet name

# Write the data frame to the new sheet
writeData(wb, sheet = "uniform", output_df)

# Save the updated Excel file
saveWorkbook(wb, file_path, overwrite = TRUE)