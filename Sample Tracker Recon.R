# ---- Program Header ----
# Project: 
# Program Name: 
# Author: 
# Created: 
# Purpose: 
# Revision History:
# Date        Author        Revision
# 

# ---- Initialize Libraries ----
library(dplyr)
library(lubridate)
library(haven)
library(openxlsx)
library(tidyverse)

# ---- Load Functions ----
setwd("C:/Users/JordanMyer/Desktop/New OneDrive/Emanate Life Sciences/DM - Inflammatix - Documents/INF-04/11. Clinical Progamming/11.3 Production Reports/11.3.3 Input Files")
source("../11.3.1  R Production Programs/INF Global Functions.R")
cleaner()
source("../11.3.1  R Production Programs/INF Global Functions.R")

# ---- Load Raw Data ----
importDOV <- read_sas("EDC/SV.sas7bdat")
importLBCOLL <- read_sas("EDC/LB_COLL.sas7bdat")
importMBCL <- read_sas("EDC/MB_CL.sas7bdat")
importLBCL <- read_sas("EDC/LB_CL.sas7bdat")
trackerFileName <- MostRecentFile("Tricore/",".*INF-04 Tricore Sample Management Tracker_Data_Management.*xlsx$","ctime")

allSites <- read.xlsx(trackerFileName,sheet=2)
for (i in 3:length(excel_sheets(trackerFileName))){
  newSite <- read.xlsx(trackerFileName,sheet = i)
  allSites <- rbind(allSites,newSite) %>% filter(!is.na(Total.Number.of.Samples))
}

trimAllSites <- allSites %>% 
  select(-c(8,10,12)) 

finalAllSites <- trimAllSites %>% 
  mutate_at(c(colnames(trimAllSites)[c(2,5,7,8,9)]),~convertToDate(.)) %>% 
  mutate(Sample.ID = paste(substring(Sample.ID,1,2),"-",substring(Sample.ID,3,5),"-",substring(Sample.ID,6,8),sep = "")) %>% 
  select(1,2,6) %>% 
  set_names(c("Subject_Number","Tricore_Date","Test"))

sentSamples <- importLBCOLL %>% 
  filter(LBYN_CRP_PCT=="Yes") %>% 
  select(2,11)%>% 
  set_names(c("Subject_Number","EDC_Date"))
  
noRESPAN <- finalAllSites %>% 
  filter(Test!="RESPAN")

singleSampleSubjects <- Filter(function(elem) length(which(noRESPAN$Subject_Number == elem)) <= 1,noRESPAN$Subject_Number)

onlyOnce <- noRESPAN %>%
  filter(Subject_Number%in%singleSampleSubjects)

withBoth <- noRESPAN %>% 
  filter(!Subject_Number%in%singleSampleSubjects)

fjBoth <- full_join(sentSamples %>% filter(!Subject_Number%in%singleSampleSubjects), withBoth)



#saveWorkbook(wb,"../11.3.4 Output Files/Tricore Target Subject Statuses.xlsx",overwrite = TRUE)

