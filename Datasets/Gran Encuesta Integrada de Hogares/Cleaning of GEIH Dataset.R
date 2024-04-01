library(readxl)
library(dplyr)
library(tidyr)
library(readr)
library(writexl)
library(lubridate)

getwd()
#setwd()

library(readxl)
df <- read_excel("anex-GEIH-dic2023-(Gran Encuesta Integrada de Hogares).xlsx", 
                 sheet = "Ocupados 23 Ciudades_rama_Trim")
#View(df)

# Function to convert column index to Excel column name (e.g., 1 -> A, 28 -> AB)
excel_cols <- function(index) {
  letters <- c(LETTERS, sapply(LETTERS, function(l1) {
    sapply(LETTERS, function(l2) paste0(l1, l2))
  }))
  return(letters[index])
}

# Remove the first 6 rows
df <- df[-(1:5), ]

# Reset the column names to the excel ones.
col_names <- sapply(1:ncol(df), excel_cols)
colnames(df) <- col_names

# Define the city names and economic sectors
city_names <-c("Total 13 ciudades y áreas metropolitanas",
  "MEDELLÍN A.M.",
  "BARRANQUILLA A.M.",
  "BOGOTÁ D.C.",
  "CARTAGENA",
  "MANIZALES A.M.",
  "MONTERÍA",
  "VILLAVICENCIO",
  "PASTO",
  "CÚCUTA A.M.",
  "PEREIRA A.M.",
  "BUCARAMANGA A.M.",
  "IBAGUÉ",
  "CALI  A.M.",
  "TUNJA",
  "FLORENCIA",
  "POPAYÁN",
  "VALLEDUPAR",
  "QUIBDÓ",
  "NEIVA",
  "RIOHACHA",
  "SANTA MARTA",
  "ARMENIA",
  "SINCELEJO")
economic_sectors <- c("Ocupados",
                      "No informa",
                      "Agricultura, ganadería, caza, silvicultura y pesca",
                      "Explotación de minas y canteras",
                      "Industrias manufactureras",
                      "Suministro de electricidad gas, agua y gestión de desechos",
                      "Construcción",
                      "Comercio y reparación de vehículos",
                      "Alojamiento y servicios de comida",
                      "Transporte y almacenamiento",
                      "Información y comunicaciones",
                      "Actividades financieras y de seguros",
                      "Actividades inmobiliarias",
                      "Actividades profesionales, científicas, técnicas y servicios administrativos",
                      "Administración pública y defensa, educación y atención de la salud humana",
                      "Actividades artísticas, entretenimiento recreación y otras actividades de servicios")

# Initialize an empty list to store the tables
city_tables <- list()

# Loop through the city names to extract each table
for (city_name in city_names) {
  # Find the start row for the city
  start_row <- which(df$A == city_name)
  # Assume the table ends 19 rows down from the start row
  end_row <- start_row + 19
  
  # Extract the table for the city
  city_table <- df[start_row:end_row, ]
  
  # Add a column with the city name
  city_table <- city_table %>%
    mutate(city = city_name) %>%
    select(city, everything()) %>%
    slice(-1)
  
  # Extract the year as a string
  year_str <- as.character(city_table[1,3])
  # Concatenate to form a full date string "YYYY-01-01"
  full_date_str <- paste0(year_str, "-01-01")
  # Convert the string to a Date object
  start_date <- as.Date(full_date_str, format = "%Y-%m-%d")
  
  # Calculate the number of months to generate column names for
  num_months <- ncol(city_table) - 2  # Subtract 2 for the 'city' and 'Concepto' columns
  
  # Generate the sequence of month names
  month_names <- seq(from = start_date, by = "month", length.out = num_months)
  month_names <- format(month_names, "%Y-%m")  # Format as "YYYY-MM"
  
  # Assign the new names to the columns starting from the third column
  new_col_names <- c("city", "Concepto", month_names)
  colnames(city_table) <- new_col_names
  
  #Delete the first 3 rows
  city_table <- city_table %>% slice(-1:-3)
  
  # Add the table to the list using the city name as the key
  city_tables[[city_name]] <- city_table
}

# Combine all city tables into one joint dataframe
joint_df <- bind_rows(city_tables)

# Transform the joint dataframe into a long format
long_df <- joint_df %>%
  pivot_longer(
    cols = -c(city, Concepto),
    names_to = "date",
    values_to = "workers"
  ) %>%
  mutate(
    # Add "-01" to the date and then convert it to a Date format
    date = as.Date(paste0(date, "-01"), format = "%Y-%m-%d"),
    # Add 2 months to each date to represent the last month of the averaging period
    date = date + months(2),
    workers = round(as.double(workers) * 1000,0),
  )

summary(long_df)

#Determine the Minimum and Maximum dates in the GEIH File in order to build the historical record.
min_date <- min(long_df$date)
max_date <- max(long_df$date)

# Convert back to "YYYY-MM" format
min_date <- format(min_date, "%Y-%m")
max_date <- format(max_date, "%Y-%m")

# Create the filename using the min and max dates
filename <- paste0("GEIH_cleaned_", min_date, "-", max_date, ".xlsx")

# Write the dataframe to an Excel file
write_xlsx(long_df, filename)