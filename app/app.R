library(shiny)
library(shinyWidgets)
library(shinythemes)
library(shinybusy)
library(dplyr)
library(openxlsx)
library(data.table)
library(reshape2)
library(tibble)
library(readxl)
library(lubridate)
library(zoo)
library(ggplot2)
library(downloader)
library(plotly)
library(tidyr)
options(shiny.maxRequestSize=1000*1024^2)


# Define UI for application 
ui <- fluidPage(
  theme = shinytheme("simplex"),
  img(height= 100, width=200, src= "logo.png"),
  titlePanel("Temperatuurscorrectie Gas"),
  tabsetPanel(
    tabPanel("Main", add_busy_spinner(spin = "fading-circle"),
             column(6, fileInput('HDD_normaal', 'Upload hier HDD normaal data',
                                 accept = c(".xlsx")), 
                    downloadButton('KNMI_data', "Download hier de KNMI data + HDD realisatie")), 
             sidebarPanel(
               fileInput('Allocatie', 'Upload Allocatie data met ook het vertaaltabel in een apart tabblad',
                         accept = c(".xlsx")),
               downloadButton('EAN_niveau', "Gecorrigeerde Allocatie Data (EAN)"),
               downloadButton('gecorrigeerd', "Gecorrigeerde Allocatie Data (Contract)")
             ), 
    pickerInput("winter", "Winter cluster", 
               choices = c("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"), 
               selected = c("1", "2", "3", "11", "12"),
               multiple = TRUE,
               options = list(`actions-box` = TRUE)),
    pickerInput("flank", "Flank cluster", 
                choices = c("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"), 
                selected = c("4", "10"),
                multiple = TRUE,
                options = list(`actions-box` = TRUE)),
    pickerInput("zomer", "Zomer cluster", 
                choices = c("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"), 
                selected = c("5", "6", "7", "8", "9"),
                multiple = TRUE,
                options = list(`actions-box` = TRUE))),
    tabPanel("HDD vergelijking", add_busy_spinner(spin = "fading-circle"), 
             actionButton("button", "Show Plot"), plotlyOutput("plot") 
    ))
)


# Define server logic 
server <- function(input, output, session) {
  
  ## Data read in from main
  # HDD normaal
  HDD_normaal_df <- reactive({ 
    req(input$HDD_normaal)
    read_excel(input$HDD_normaal$datapath, guess_max = 20000)
  })
  
  # Allocatie
  Allocatie_df <- reactive({ 
    req(input$Allocatie)
    read_excel(input$Allocatie$datapath, guess_max = 20000, sheet="Allocatie")%>%
      rename(date =`#`) %>% select(-CODE) 
  })
  
  # Vertaaltabel
  matching_EAN_Contract <- reactive({ 
    req(input$Allocatie)
    read_excel(input$Allocatie$datapath, guess_max = 20000, sheet="Mapping") %>%
      select(EAN, `Top contract`) %>% rename(EANs = EAN, Contract = `Top contract`)
  })
  
  ### Main section
  # HDD_realisatie webscraping & computation from KNMI
  df <- reactive({
    
    # haal de KNMI data op voor de berekening van de HDD realisatie 
    download("https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/daggegevens/etmgeg_240.zip", 
             dest="KNMI.zip", mode="wb") 
    unzip ("KNMI.zip")
    KNMI <- read.table(file = "etmgeg_240.txt", fill = TRUE, skip= 51, sep = ",")
    names(KNMI) <- c("STN","YYYYMMDD","DDVEC","FHVEC",   "FG",  "FHX", "FHXH",  
                     "FHN", "FHNH",  "FXX", "FXXH",   "TG",   "TN",  "TNH",   "TX",  "TXH", 
                     "T10N","T10NH",   "SQ",   "SP",    "Q",   "DR",   "RH",  "RHX", "RHXH",   "PG",   
                     "PX",  "PXH",   "PN",  "PNH",  "VVN", "VVNH",  "VVX", "VVXH",   "NG",   "UG",  
                     "UX",  "UXH",   "UN",  "UNH", "EV24")
    
    # berekening HDD_realisatie en gemiddelde nemen per maand "=GEMIDDELDE(T22;R22)/10-(0,3*(J22/10))+(0,002*Z22)"
    KNMI_data <- KNMI %>% mutate(HDD_realisatie = 
                                   ifelse((18-(TN/10 + TX/10)/2) > 0, 18-(TN/10 + TX/10)/2, 0)) %>%
      select(YYYYMMDD, HDD_realisatie, FG, TN, TX, Q)
    KNMI_data$YYYYMMDD <- as.Date(as.character(KNMI_data$YYYYMMDD),format="%Y%m%d") 
    
    KNMI_data <- KNMI_data %>% 
      rename(Date = YYYYMMDD)
    
    KNMI_data

  })
  
  
  ## computing corrected volume bericht via lineair regression analysis divided in seasonal clusters
  #EAN level 
  df2 <- reactive({
  
  # formatting allocation again 
    ## step 1: data wrangling allocation data 
    # import data
    Allocation <- Allocatie_df() 
    
    # remove EANs without data (NAs)
    Allocation <- Allocation[, colSums(is.na(Allocation)) != nrow(Allocation)] %>% as.data.frame()
    
    # cut off the names of the allocation df so it matches 
    names(Allocation)[2:length(Allocation)] <- substring(names(Allocation)[2:length(Allocation)], first = 3, last = 20)
    
    # format date variable in allocation df
    Allocation <- Allocation %>% mutate(date = ymd(date)) %>% rename(Date = date)
    
    # Divide the allocation data by 35.17 to go from joules to m^3
    Allocation[, -1] <- Allocation[, -1]/35.17
    
  # import hdd realisatie and hdd normaal 
  HDD_realisatie <- df()
  HDD_realisatie <- HDD_realisatie[,1:2]
  HDD_normaal <- HDD_normaal_df()
  HDD_normaal$Date <- as.Date(HDD_normaal$Date)
  
  ### Step 3: merge first the HDD realisation and the allocation 
  # add the werkelijke afname of all EAN to the table 
  df <- merge(HDD_realisatie, Allocation, by = "Date" ) 
  
  ### step 4: make segmentation 1 
  # add monthly dummy variable 
  df$month <- as.numeric(substring(df$Date, 6,7))
  
  # divide data into desired segmentation (need to be flexible) 
  winter <- df %>% filter(month %in% unlist(strsplit(input$winter, ", "))) %>% select(-month) 
  flank <- df %>% filter(month %in% unlist(strsplit(input$flank, ", "))) %>% select(-month) 
  zomer <- df %>% filter(month %in% unlist(strsplit(input$zomer, ", "))) %>% select(-month) 


  ### Step 5: For each separate contract, conduct weather correction 
  # for each EAN number, we conduct a separate regression analysis 
  # in order to see what the relationship is of the weather (X) compared to the offtake (Y)
  
  # regression for winter months for each ean number 
  w_regressions <- lapply(3:length(winter), function(x) lm(winter[,x] ~ winter$HDD_realisatie))
  
  # regression for winter months for each ean number 
  f_regressions <- lapply(3:length(flank), function(x) lm(flank[,x] ~ flank$HDD_realisatie))
  
  # regression for winter months for each ean number 
  s_regressions <- lapply(3:length(zomer), function(x) lm(zomer[,x] ~ zomer$HDD_realisatie))
  
  ## now we need HDD normaal which is on monthly basis so from here calculations are on monthly
  # Collapse all dataframes of the seasons on monthly basis
  winter <- winter %>%
    group_by(Date = strftime(Date, "%Y-%m")) %>%
    summarise_if(is.numeric, sum) %>% transform(Date = as.Date(as.yearmon(Date)))
  names(winter)[3:length(winter)] <- substring(names(winter)[3:length(winter)], first = 2, last = 20)
  
  flank <- flank %>%
    group_by(Date = strftime(Date, "%Y-%m")) %>%
    summarise_if(is.numeric, sum) %>% transform(Date = as.Date(as.yearmon(Date)))
  names(flank)[3:length(flank)] <- substring(names(flank)[3:length(flank)], first = 2, last = 20)
  
  zomer <- zomer %>%
    group_by(Date = strftime(Date, "%Y-%m")) %>%
    summarise_if(is.numeric, sum) %>% transform(Date = as.Date(as.yearmon(Date)))
  names(zomer)[3:length(zomer)] <- substring(names(zomer)[3:length(zomer)], first = 2, last = 20)
  
  # merging both HDD's
  winter <- merge(HDD_normaal, winter, by = "Date" )
  flank <- merge(HDD_normaal, flank, by = "Date" )
  zomer <- merge(HDD_normaal, zomer, by = "Date" )

  # compute the different between the HDDs
  w_HDD_afwijking <- winter$HDD_normaal - winter$HDD_realisatie
  f_HDD_afwijking <- flank$HDD_normaal - flank$HDD_realisatie
  s_HDD_afwijking <- zomer$HDD_normaal - zomer$HDD_realisatie
  
  # subset allocation data
  w_Allocation_monthly <- as.data.frame(winter[,4:length(winter)])
  f_Allocation_monthly <- as.data.frame(flank[,4:length(flank)])
  s_Allocation_monthly <- as.data.frame(zomer[,4:length(zomer)])
  
  # for each EAN, now we need to compute the corrected offtake based on the regression results, corrected with it's corr coefficient
  w_gecorrigeerde_afname <- lapply(1:length(w_Allocation_monthly), function (x) w_Allocation_monthly[,x] + w_HDD_afwijking * 
                                     w_regressions[[x]][["coefficients"]][["winter$HDD_realisatie"]])
  f_gecorrigeerde_afname <- lapply(1:length(f_Allocation_monthly), function (x) f_Allocation_monthly[,x] + f_HDD_afwijking * 
                                     f_regressions[[x]][["coefficients"]][["flank$HDD_realisatie"]])
  s_gecorrigeerde_afname <- lapply(1:length(s_Allocation_monthly), function (x) s_Allocation_monthly[,x] + s_HDD_afwijking * 
                                     s_regressions[[x]][["coefficients"]][["zomer$HDD_realisatie"]])
  
  # convert list to data frame 
  w_gecorrigeerde_afname <- as.data.frame(do.call(cbind, w_gecorrigeerde_afname))
  names(w_gecorrigeerde_afname) <- names(winter)[4:length(winter)]
  w_gecorrigeerde_afname <- cbind(Date = winter$Date, w_gecorrigeerde_afname)
  
  f_gecorrigeerde_afname <- as.data.frame(do.call(cbind, f_gecorrigeerde_afname))
  names(f_gecorrigeerde_afname) <- names(flank)[4:length(flank)]
  f_gecorrigeerde_afname <- cbind(Date = flank$Date, f_gecorrigeerde_afname)
  
  s_gecorrigeerde_afname <- as.data.frame(do.call(cbind, s_gecorrigeerde_afname))
  names(s_gecorrigeerde_afname) <- names(zomer)[4:length(zomer)]
  s_gecorrigeerde_afname <- cbind(Date = zomer$Date, s_gecorrigeerde_afname)
  
  # merge all seasonality dataframes together
  gecorrigeerde_afname <- rbind(w_gecorrigeerde_afname, f_gecorrigeerde_afname, s_gecorrigeerde_afname) %>%
    arrange(Date)
  
  # round it so it is in full m3
  gecorrigeerde_afname[,-1] <- round(gecorrigeerde_afname[,-1])
  
  gecorrigeerde_afname

  })
  
  # continue to get analysis on contract level 
  df3 <- reactive({
    
  gecorrigeerde_afname <- df2()
    
  ### Step 6: Match EAN numbers with contract 
  # import EAN <> Contract file 
  matching_EAN_Contract <- matching_EAN_Contract()
  
  ## Make dataframe with 3-dimensions, in which EAN and contract matches the allocation 
  # change date variable to eBase standard 
  gecorrigeerde_afname$Date <- format(as.Date(gecorrigeerde_afname$Date, "%Y-%m-%d"), "%Y%m")
  
  # transpose the corrected allocation df
  gecorrigeerde_afname <- gecorrigeerde_afname %>%
    gather("EANs", "value", 2:ncol(gecorrigeerde_afname)) %>%
    spread(Date, value)
  
  ## merge both df based on EANs and aggregate on contract number 
  gecorr_afname_contract <- merge(matching_EAN_Contract, gecorrigeerde_afname, by= "EANs") %>% 
    select(-EANs)
  
  # summarize the data on Contract number 
  gecorr_afname_contract <- gecorr_afname_contract %>% 
    group_by(Contract) %>%
    summarise_all(sum) %>% as.data.frame
  
  # set gecorr_afname_contract such that each element has its own column 
  volume_bericht_gecorrigeerd <- melt(setDT(gecorr_afname_contract), 
                                      id.vars = c("Contract"), 
                                      variable.name = "Levper") 
  volume_bericht_gecorrigeerd <- volume_bericht_gecorrigeerd %>% 
    rename(volume_bericht_gecorrigeerd = value) %>% as.data.frame %>% arrange(Contract, Levper) 
  
  volume_bericht_gecorrigeerd 
  
  })
  
  # plotting the HDDs against each other 
  hdd_comp_plot <- eventReactive(input$button, {

    # computing again HDD_realisation (otherwise plotly won't work) {
    
    # importeer de raw data of hdd_normaal and put date column correctly
    HDD_normaal <- HDD_normaal_df()
    HDD_normaal$Date <- as.Date(HDD_normaal$Date)
    
    # haal de KNMI data op voor de berekening van de HDD realisatie 
    download("https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/daggegevens/etmgeg_240.zip", dest="KNMI.zip", mode="wb") 
    unzip ("KNMI.zip")
    KNMI <- read.table(file = "etmgeg_240.txt", fill = TRUE, skip= 51, sep = ",")
    names(KNMI) <- c("STN","YYYYMMDD","DDVEC","FHVEC",   "FG",  "FHX", "FHXH",  
                     "FHN", "FHNH",  "FXX", "FXXH",   "TG",   "TN",  "TNH",   "TX",  "TXH", 
                     "T10N","T10NH",   "SQ",   "SP",    "Q",   "DR",   "RH",  "RHX", "RHXH",   "PG",   
                     "PX",  "PXH",   "PN",  "PNH",  "VVN", "VVNH",  "VVX", "VVXH",   "NG",   "UG",  
                     "UX",  "UXH",   "UN",  "UNH", "EV24")
    
    # berekening HDD_realisatie en gemiddelde nemen per maand "=GEMIDDELDE(T22;R22)/10-(0,3*(J22/10))+(0,002*Z22)"
    HDD_realisatie <- KNMI %>% mutate(HDD_realisatie = 
                                        ifelse((18-(TN/10 + TX/10)/2) > 0, 18-(TN/10 + TX/10)/2, 0)) %>%
      select(YYYYMMDD, HDD_realisatie)
    HDD_realisatie$YYYYMMDD <- as.Date(as.character(HDD_realisatie$YYYYMMDD),format="%Y%m%d") 
    
    HDD_realisatie <- HDD_realisatie %>% 
      rename(Date = YYYYMMDD)
    HDD_realisatie # }
    
    # comparing the change of HDDs via a plot 
    HDD_realisatie <- HDD_realisatie %>%
      group_by(Date = strftime(Date, "%Y-%m")) %>%
      summarise_if(is.numeric, sum) %>% transform(Date = as.Date(as.yearmon(Date)))
    
    hdd_comp <- merge(HDD_realisatie, HDD_normaal, by = "Date") 
    
    hdd_comp_plot <- plot_ly(x = hdd_comp[,1], y = hdd_comp[,2], 
                             type = 'scatter', mode = 'lines', name = "HDD Realisatie") %>% 
      add_trace(y=hdd_comp[,3], name = "HDD Normaal") %>%
      layout(title= "HDD comparison over time", 
             yaxis = list(title = 'HDD value'), 
             legend = list(orientation = "h",   # show entries horizontally
                           xanchor = "center",  
                           x = 0.5))      
  })
  
  
  # ------------------------------------------------------------------------------     
  
  ## Main
  # Output excel file HDD realisatie 
  output$KNMI_data <- downloadHandler(
    filename = function() { 
      paste("KNMI data en HDD Realisatie",".xlsx", sep="")
    },
    content = function(file) {
      write.xlsx(df(), file)
    })
  
  # Output excel file corrected allocation 
  output$EAN_niveau <- downloadHandler(
    filename = function() { 
      paste("Gecorrigeerd volume bericht (EAN)",".xlsx", sep="")
    },
    content = function(file) {
      write.xlsx(df2(), file)
    })
  
  # Output excel file corrected allocation 
  output$gecorrigeerd <- downloadHandler(
    filename = function() { 
      paste("Gecorrigeerd volume bericht (Contract)",".xlsx", sep="")
    },
    content = function(file) {
      write.xlsx(df3(), file)
    })
  
  # output of our plotly 
  output$plot <- renderPlotly({ 
    hdd_comp_plot()
  })
  
  
  session$onSessionEnded(function() {
    stopApp()
  })
  
}

# Run the application 
shinyApp(ui = ui, server = server)






















