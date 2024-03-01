library(readxl)
library(dplyr)
library(tidyr)
library(RColorBrewer)
library(plotly)
library(stringdist)

#Read Excel file and identify the working sheets.
file_path <- "1.2.4.IPC_Por ciudad_IQY.xlsx" 
sheet_names <- excel_sheets(file_path)

dfs <- list()
# Loop through the sheet names starting from the second sheet
for (i in 2:length(sheet_names)) {
  # Read each sheet into a data frame and store it in the list
  dfs[[i - 1]] <- read_excel(file_path, sheet = sheet_names[i])
}

# Function to transform dataset from wide to long format
transform_dataset <- function(df) {
  df_long <- df %>%
    pivot_longer(
      cols = -1, # Select all columns except the first one (city names)
      names_to = "date", 
      values_to = "value"
    ) %>%
    mutate(
      date = as.Date(as.numeric(date), origin = "1899-12-30"), # Convert Excel serial date number to Date
      date = format(date, "%Y-%m") # Format date as "YYYY-mm"
    ) %>%
    rename(city = 1)
  return(df_long)
}

#Store the datasets
IPC <- dfs[[1]]
year_to_date_var <- dfs[[2]]
yearly_var <- dfs[[3]]
montly_var <- dfs[[4]]

# Apply the transformation to each dataset
IPC_long <- transform_dataset(IPC)
year_to_date_var_long <- transform_dataset(year_to_date_var)
yearly_var_long <- transform_dataset(yearly_var)
montly_var_long <- transform_dataset(montly_var)

#Filter the cities of our interest
city_info <- read_excel("Ciudades_y_sus_Areas_Metropolitanas_en_GEIH.xlsx")

# Filter the dataset for rows where the "Tipo" column is "principal"
main_cities <- city_info %>%
  filter(Tipo == "principal")
main_cities


#Clean the IPC dataset to only include the information of the cities we are going to use in our analysis, from 2015 onward. 
IPC_model <- inner_join(IPC_long, main_cities, by = c("city" = "ciudad_am"))
IPC_model <- IPC_model %>%
  filter(date >= "2015-01") %>%
  select(city, date, value) %>%
  rename(IPC = value)

#Check the resulting dataset
summary(IPC_model)

# Create the interactive plot directly with plotly
plotly_obj <- plot_ly(data = IPC_model, x = ~date, y = ~IPC, 
                      color = ~city, colors = RColorBrewer::brewer.pal(n = 8, name = "Dark2"),
                      type = 'scatter', mode = 'lines+markers',
                      text = ~city, hoverinfo = 'text+x+y') %>%
  layout(title = "IPC Evolution by City from 2000 to 2023",
         xaxis = list(title = "date"),
         yaxis = list(title = "IPC", tickformat = ",d", 
                      range = c(0, max(IPC_model$IPC, na.rm = FALSE))),
         legend = list(orientation = "v", x = 1.05, y = 1))

# Display the interactive plot
plotly_obj

#Since there were no NA's, we'll clean the other datasets as well 
