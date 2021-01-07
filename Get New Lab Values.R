
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


importLabResults <- read.xlsx("Lab Results/INF-04_Prod_Tricore_Laboratory_FULL.xlsx")
importLBCL <- read_sas("EDC/LB_CL.sas7bdat")

newResults <- read.xlsx("Lab Results/INF-04_Prod_Tricore_Laboratory_FULL.xlsx") %>% 
  mutate(Subject.Identifier = fixTricoreSubjectNumber(Subject.Identifier)) %>% 
  filter(!Subject.Identifier%in%unique(importLBCL$SubjectID)) %>% 
  mutate(PCTR.Results = ifelse(test=grepl("E",PCTR.Results),yes = as.numeric(substring(PCTR.Results,1,3))/100,ifelse(test = grepl("<",PCTR.Results),yes = PCTR.Results, no = round(as.numeric(PCTR.Results),2)))) %>% 
  mutate(CRP.Result = ifelse(test=grepl("E",CRP.Result),yes = as.numeric(substring(CRP.Result,1,3))/100,ifelse(test = grepl("<",CRP.Result),yes = CRP.Result, no = round(as.numeric(CRP.Result),2)))) %>%
  select(5,8,10,11,13)


fileName <- paste("../11.3.4 Output Files/TriCore/New Data to Enter",Sys.Date(),".xlsx")

write.xlsx(newResults,fileName)


