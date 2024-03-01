library(readxl)
library(dplyr)
library(tidyr)
library(RColorBrewer)
library(plotly)
library(stringdist)
library(scales)

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
CPI <- dfs[[1]]
year_to_date_var <- dfs[[2]]
yearly_var <- dfs[[3]]
montly_var <- dfs[[4]]

# Apply the transformation to each dataset
CPI_long <- transform_dataset(CPI)
year_to_date_var_long <- transform_dataset(year_to_date_var)
yearly_var_long <- transform_dataset(yearly_var)
montly_var_long <- transform_dataset(montly_var)

#Filter the cities of our interest
city_info <- read_excel("Ciudades_y_sus_Areas_Metropolitanas_en_GEIH.xlsx")

# Filter the dataset for rows where the "Tipo" column is "principal"
main_cities <- city_info %>%
  filter(Tipo == "principal")
main_cities


#Clean the CPI dataset to only include the information of the cities we are going to use in our analysis, from 2015 onward. 
CPI_model <- inner_join(CPI_long, main_cities, by = c("city" = "ciudad_am"))
CPI_model <- CPI_model %>%
  filter(date >= "2015-01") %>%
  select(city, date, value) %>%
  rename(CPI = value)

#Check the resulting dataset
summary(CPI_model)

# Create the interactive plot directly with plotly
plotly_obj <- plot_ly(data = CPI_model, x = ~date, y = ~CPI, 
                      color = ~city, colors = RColorBrewer::brewer.pal(n = 8, name = "Dark2"),
                      type = 'scatter', mode = 'lines+markers',
                      text = ~city, hoverinfo = 'text+x+y') %>%
  layout(title = "CPI Evolution by City from 2015 to 2023",
         xaxis = list(title = "date"),
         yaxis = list(title = "CPI", tickformat = ",d", 
                      range = c(0, max(CPI_model$CPI, na.rm = FALSE))),
         legend = list(orientation = "v", x = 1.05, y = 1))

# Display the interactive plot
plotly_obj

#Since there were no NA's, we'll clean the other datasets as well 

# Year to date CPI Variation %
year_to_date_var_model <- inner_join(year_to_date_var_long, main_cities, by = c("city" = "ciudad_am"))
year_to_date_var_model <- year_to_date_var_model %>%
  filter(date >= "2015-01") %>%
  select(city, date, value) %>%
  rename(CPI_year_to_date_var = value)
year_to_date_var_model$CPI_year_to_date_var <- year_to_date_var_model$CPI_year_to_date_var/100
summary(year_to_date_var_model) #is good

plotly_obj <- plot_ly(data = year_to_date_var_model, x = ~date, y = ~CPI_year_to_date_var, 
                      color = ~city, colors = RColorBrewer::brewer.pal(n = 8, name = "Dark2"),
                      type = 'scatter', mode = 'lines+markers',
                      text = ~city, hoverinfo = 'text+x+y') %>%
  layout(title = "Year to Date CPI Variation Evolution by City from 2015 to 2023",
         xaxis = list(title = "date"),
         yaxis = list(title = "CPI year to date variation %",  
                      range = c(0, max(year_to_date_var_model$CPI_year_to_date_var, na.rm = FALSE)+0.01)),
         legend = list(orientation = "v", x = 1.05, y = 1))
# Display the interactive plot
plotly_obj

#Yearly CPI variation %
yearly_var_model <- inner_join(yearly_var_long, main_cities, by = c("city" = "ciudad_am"))
yearly_var_model <- yearly_var_model %>%
  filter(date >= "2015-01") %>%
  select(city, date, value) %>%
  rename(CPI_year_var = value)
yearly_var_model$CPI_year_var <- yearly_var_model$CPI_year_var/100
summary(yearly_var_model) #is good

plotly_obj <- plot_ly(data = yearly_var_model, x = ~date, y = ~CPI_year_var, 
                      color = ~city, colors = RColorBrewer::brewer.pal(n = 8, name = "Dark2"),
                      type = 'scatter', mode = 'lines+markers',
                      text = ~city, hoverinfo = 'text+x+y') %>%
  layout(title = "Yearly CPI Variation Evolution by City from 2015 to 2023",
         xaxis = list(title = "date"),
         yaxis = list(title = "CPI yearly variation %", 
                      range = c(0, max(yearly_var_model$CPI_year_var, na.rm = FALSE)+0.01)),
         legend = list(orientation = "v", x = 1.05, y = 1))
# Display the interactive plot
plotly_obj

#Monthly CPI Variation %
montly_var_model <- inner_join(montly_var_long, main_cities, by = c("city" = "ciudad_am"))
montly_var_model <- montly_var_model %>%
  filter(date >= "2015-01") %>%
  select(city, date, value) %>%
  rename(CPI_month_var = value)
montly_var_model$CPI_month_var <- montly_var_model$CPI_month_var/100
summary(montly_var_model) #is good

plotly_obj <- plot_ly(data = montly_var_model, x = ~date, y = ~CPI_month_var, 
                      color = ~city, colors = RColorBrewer::brewer.pal(n = 8, name = "Dark2"),
                      type = 'scatter', mode = 'lines+markers',
                      text = ~city, hoverinfo = 'text+x+y') %>%
  layout(title = "Monthly CPI Variation Evolution by City from 2015 to 2023",
         xaxis = list(title = "date"),
         yaxis = list(title = "CPI monthly variation %", 
                      range = c(0, max(montly_var_model$CPI_month_var, na.rm = FALSE)+0.01)),
         legend = list(orientation = "v", x = 1.05, y = 1))
# Display the interactive plot
plotly_obj

# Extract the minimum and maximum dates from the dataset
min_date <- min(CPI_model$date)
max_date <- max(CPI_model$date)

# Create the filename using the min and max dates
filename_CPI <- paste0("CPI_cleaned_", min_date, "-", max_date, ".xlsx")
filename_year_to_date_var <- paste0("CPI_Year_to_date_var_cleaned_", min_date, "-", max_date, ".xlsx")
filename_yearly_var <- paste0("CPI_Yearly_var_cleaned_", min_date, "-", max_date, ".xlsx")
filename_montly_var <- paste0("CPI_Montly_var_cleaned_", min_date, "-", max_date, ".xlsx")

# Write the dataframe to an Excel file
write_xlsx(CPI_model, filename_CPI)
write_xlsx(year_to_date_var_model, filename_year_to_date_var)
write_xlsx(yearly_var_model, filename_yearly_var)
write_xlsx(montly_var_model, filename_montly_var)



