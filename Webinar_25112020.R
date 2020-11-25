# IDA Webinar
# Tue Hellstern
# 25-11-2020

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

# *************************
# Data import
# *************************

TourismeSpend <- read_excel("tourism.xlsx",
                            sheet = "Data",
                            skip = 3)


TourismeMetaData <- read_excel("tourism.xlsx",
                               sheet = "MetadataCountries",
                               range = cell_cols("A:C"))

# *************************
# Data tilpas
# *************************

# Kun 2008 til 2018
TourismeSpend <- TourismeSpend %>% 
  select(1:2, as.numeric(53:63)) %>% 
  rename(CountryName = "Country Name", CountryCode = "Country Code")

TourismeSpendList <- TourismeSpend %>% 
  gather(Year, Spend, "2008":"2018") %>% 
  drop_na()

TourismeMetaData <- TourismeMetaData %>% 
  rename(CountryCode = "Country Code") %>% 
  drop_na()

# *************************
# Data Join
# *************************
TourismeSpendRegion <- TourismeSpendList %>% 
  left_join(TourismeMetaData, by = c("CountryCode" = "CountryCode")) %>% 
  select("CountryName", "Region", "Year", "Spend") %>% 
  drop_na() %>% 
  group_by(Region, Year) %>% 
  summarise(Spend = sum(Spend)) %>% 
  arrange(desc(Spend))

NordicSpend <- TourismeSpend %>% 
  filter(CountryName %in% c("Denmark", "Sweden", "Norway"))

NordicSpendList <- NordicSpend %>% 
  gather(Year, Spend, "2008":"2018")

TourismeSpendRegionYear <- TourismeSpendList %>% 
  left_join(TourismeMetaData, by = c("CountryCode" = "CountryCode")) %>% 
  select ("CountryName", "Region", "Year", "Spend") %>% 
  drop_na() %>% 
  group_by(Region, Year) %>% 
  summarise(Spend = sum(Spend)) %>% 
  arrange(desc(Spend))

# *************************
# Data Plot
# *************************

# Line
NordicPlot <- ggplot(data = NordicSpendList, aes(x=Year, y=Spend, group=CountryName)) +
  geom_line(aes(color=CountryName)) +
  geom_point(aes(color=CountryName)) +
  ggtitle("Total spend by Nordic country") +
  theme_minimal()

NordicPlot

TourismeSpendRegionYearPlot <- ggplot(data=TourismeSpendRegionYear, aes(x=Year, y=Spend, fill=Region)) +
  geom_bar(stat = "identity") +
  theme_minimal()

TourismeSpendRegionYearPlot


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


