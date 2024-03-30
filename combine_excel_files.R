# combine data from excel files
# combine yield data from IBPA

# libraries
#install.packages("readxl")
#install.packages("dplyr")
library(readxl)
library(dplyr)
library(openxlsx) # tuk print dalam bentuk xlsx

# yang per tahun

rm(list=ls()) #clear the environment

#Seri <- "PBS032"
#tempat_nyimpen <- "C:/Users/andy.mustafa/OneDrive - Kemenkeu/2. Seksi Analisis Harga SBSN/#Kajian_analisis_seksiAH/2024/inverted_PBS2Y_2024/Kepemilikan_PBS_32"
#tempat_output <- "C:/Users/andy.mustafa/OneDrive - Kemenkeu/2. Seksi Analisis Harga SBSN/#Kajian_analisis_seksiAH/2024/inverted_PBS2Y_2024"

tempat_nyimpen <- "C:/Users/elva.novitasari/OneDrive - Kemenkeu/2. Seksi Analisis Harga SBSN/#Kajian_analisis_seksiAH/2024/inverted_PBS2Y_2024/Kepemilikan_PBS_32"
tempat_output <- "C:/Users/elva.novitasari/OneDrive - Kemenkeu/2. Seksi Analisis Harga SBSN/#Kajian_analisis_seksiAH/2024/inverted_PBS2Y_2024"

setwd(tempat_nyimpen)

# List the Excel files in the directory:
files <- list.files(pattern = "\\.xlsx$")

#Read and combine the data:
# Read and combine the data:
all_data <- lapply(files, function(file) {
  data <- read_excel(file)
  file_name <- substr(basename(file), nchar(basename(file)) - 14, nchar(basename(file)))  # Extract last 10 characters
  data$FileName <- file_name  # Add a new column with the file name
  return(data)
})  # Reads each Excel file into a list

combined_data <- bind_rows(all_data)  # Combines the list into a single data frame

# print the output
setwd(tempat_output)
write.xlsx(combined_data, paste0("Kepemilikan_tiap_seri.xlsx"))


#all_data <- lapply(files, function(file) read_excel(file))  # Reads each Excel file into a list
#combined_data <- bind_rows(all_data)  # Combines the list into a single data frame

