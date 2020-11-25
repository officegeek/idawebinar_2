# IDA Webinar
# Tue Hellstern
# 25-11-2020

# CRTL + - = <- 
# CTRL + SHIFT + M = %>% 
#

# Pakker
# install.packages("readxl")
# install.packages("dplyr")
# install.packages("tidyr")
# install.packages("ggplot2")
# install.packages("officer")

library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(officer)

# *****************************************************************
# Data import
# *****************************************************************

# TourismeSpend 
TourismeSpend <- read_excel("tourism.xlsx", 
                            sheet = "Data",
                            skip = 3) 

# MetadataCountries
TourismeMetadata <- read_excel("tourism.xlsx",
                               sheet = "MetadataCountries",
                               range = cell_cols("A:C"))

# *****************************************************************
# Tilpas data
# *****************************************************************

# TourismeSpend 
# Kun 2008 til 2018
TourismeSpend <- TourismeSpend %>% 
  select(1:2, as.numeric(53:63)) %>% 
  rename(CountryName = "Country Name", CountryCode = "Country Code")

# gather()
TourismeSpendList <- TourismeSpend %>% 
  gather(Year, Spend, "2008":"2018") %>% 
  drop_na() 


# TourismeMetadata
# Fjern NA (tidy)
TourismeMetadata <- TourismeMetadata %>% 
  rename(CountryCode = "Country Code") %>% 
  drop_na()


# Join
# Spend by Region 2008 - 2018
TourismeSpendRegion <- TourismeSpendList %>% 
  left_join(TourismeMetadata, by = c("CountryCode" = "CountryCode")) %>% 
  select ("CountryName", "Region", "Year", "Spend") %>% 
  drop_na() %>% 
  group_by(Region, Year) %>% 
  summarise(Spend = sum(Spend)) %>% 
  arrange(desc(Spend))

TourismeSpendRegionYear <- TourismeSpendList %>% 
  left_join(TourismeMetadata, by = c("CountryCode" = "CountryCode")) %>% 
  select ("CountryName", "Region", "Year", "Spend") %>% 
  drop_na() %>% 
  group_by(Region, Year) %>% 
  summarise(Spend = sum(Spend)) %>% 
  arrange(desc(Spend))

# Udvikling Nordiske lande
NordicSpend <- TourismeSpend %>% 
  rename(CountryName = "CountryName") %>% 
  select(CountryName, "2008":"2018") %>% 
  filter(CountryName %in% c("Denmark", "Sweden", "Norway"))

# gather()
NordicSpendList <- NordicSpend %>% 
  gather(Year, Spend, "2008":"2018")
  


# *****************************************************************
# Plot - ggplot
# *****************************************************************

# Bar
TourismeSpendRegionYearPlot <- ggplot(data = TourismeSpendRegionYear, aes(x = Year, y = Spend, fill = Region)) +
  geom_bar(stat = "identity") +
  scale_y_continuous(name="USD", labels=function(x) format(x, big.mark = ".", decimal.mark = ",", scientific = FALSE)) +
  theme_minimal()

TourismeSpendRegionYearPlot

# Line
NordicPlot <- ggplot(data = NordicSpendList, aes(x = Year, y = Spend, group = CountryName)) +
  geom_line(aes(color=CountryName)) +
  geom_point(aes(color=CountryName)) +
  scale_y_continuous(name="In USD", labels=function(x) format(x, big.mark = ".", decimal.mark = ",", scientific = FALSE)) +
  ggtitle("Total spend by country") +
  theme_minimal()


NordicPlot


# *****************************************************************
# PowerPoint
# *****************************************************************

# Opret PowerPoint objekt
PowerPointTourisme <- read_pptx()

PowerPointTourisme <- PowerPointTourisme %>%
  # Titel slide
  add_slide(layout = "Title Only", master = "Office Theme") %>% 
  ph_with(value = "Tourisme spend", location = ph_location_type(type = "title")) %>% 
  ph_with(value = "Tue Hellstern", location = ph_location_type(type = "ftr")) %>% 
  
  # Slide Plot
  add_slide(layout = "Title and Content") %>% 
  ph_with(value = "Spend by region", location = ph_location_type(type = "title")) %>%
  ph_with(value = TourismeSpendRegionYearPlot, location = ph_location_type(type = "body")) %>% 
  
  # Slide Plot
  add_slide(layout = "Title and Content") %>% 
  ph_with(value = "Spend by nordic country", location = ph_location_type(type = "title")) %>%
  ph_with(value = NordicPlot, location = ph_location_type(type = "body"))

# Gem PowerPoint
print(PowerPointTourisme, target = paste("Tourisme", format(Sys.Date(), "%d-%m-%Y"), ".pptx"))

