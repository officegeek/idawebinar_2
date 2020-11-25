# IDA Webinar
# Tue Hellstern
# 25-11-2020

library(shiny)
library(dplyr)
library(readxl)
library(ggthemes)
library(ggplot2)

# Data
TourismeSpend <- read_excel("tourism.xlsx", 
                            sheet = "Data",
                            skip = 3) 

TourismeMetadata <- read_excel("tourism.xlsx",
                               sheet = "MetadataCountries",
                               range = cell_cols("A:C"))

TourismeMetadata <- TourismeMetadata %>% 
    rename(CountryCode = "Country Code") %>% 
    drop_na()

TourismeSpend <- TourismeSpend %>% 
    select(1:2, as.numeric(53:63)) %>% 
    rename(CountryName = "Country Name", CountryCode = "Country Code") %>% 
    gather(Year, Spend, "2008":"2018") %>% 
    left_join(TourismeMetadata, by = c("CountryCode" = "CountryCode")) %>% 
    select ("Region", "Year", "Spend") %>% 
    drop_na()

# Define server logic
function(input, output) {
    
    output$selected_var <- renderText({
        paste("Valg af Region", input$ValgtRegion)
    })
    
    output$TourismePlot <- renderPlot({
        
        # Opret barplot
        filter(TourismeSpend, Region == input$ValgtRegion) %>%
            ggplot(aes(x = Year, y = Spend, fill = Region)) +
            geom_bar(stat = "identity", fill="lightblue") +
            scale_y_continuous(name="USD", labels=function(x) format(x, big.mark = ".", decimal.mark = ",", scientific = FALSE)) +
            theme(legend.position="none")
    })
}
