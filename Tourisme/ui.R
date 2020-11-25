# IDA Webinar
# Tue Hellstern
# 25-11-2020

library(readxl)
library(shiny)
library(dplyr)
library(tidyr)

# Data
TourismeMetadata <- read_excel("tourism.xlsx",
    sheet = "MetadataCountries",
    range = cell_cols("B")) %>% 
    distinct() %>% 
    drop_na()

# Define UI for application
shinyUI(fluidPage(
    
    titlePanel("Salg efter Region"),
    
    # Opret sidebar
    sidebarLayout(
        sidebarPanel(helpText("Du skal vælge en Region"),
                     selectInput("ValgtRegion", h3("Vælg land"),
                                 choices = TourismeMetadata$Region,
                                 selected = 1)),
        
        # Placering af Plot
        mainPanel(
            plotOutput("TourismePlot")
        )
    )
))
