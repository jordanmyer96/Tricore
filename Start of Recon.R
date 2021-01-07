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

importDOV <- read_sas("EDC/SV.sas7bdat")
importLBCOLL <- read_sas("EDC/LB_COLL.sas7bdat")
importMBCL <- read_sas("EDC/MB_CL.sas7bdat")
importLBCL <- read_sas("EDC/LB_CL.sas7bdat")
importSD <- read_sas("EDC/DS_SD.sas7bdat")
importLabResults <- read.xlsx(MostRecentFile("Tricore/",".*INF-04_Prod_Tricore_Laboratory_FULL.*xlsx$","ctime"))
importMBResults <- read.xlsx(MostRecentFile("Tricore/",".*INF-04_Prod_Tricore_Microbiology_FULL.*xlsx$","ctime"))


trimDOV <- importDOV %>% 
  select(2,4)

trimLBCOLL <- importLBCOLL %>% 
  select(2,4,17,31)

trimSD <- importSD%>% 
  select(2,4,7,9)




trimLBCL <- importLBCL %>% 
  mutate(CRPEntered = ifelse(!is.na(LBORRES_CRP)|LBSTAT_CRP=="Not Done","Yes","No")) %>% 
  mutate(PCTEntered = ifelse(!is.na(LBORRES_PCT)|LBSTAT_PCT=="Not Done","Yes","No")) %>%  
  mutate(`Lab Entry Status` = ifelse(CRPEntered=="Yes"&PCTEntered=="Yes","Complete",
                                     ifelse(CRPEntered=="Yes"&PCTEntered=="No","Only CRP Entered",
                                            ifelse(CRPEntered=="No"&PCTEntered=="Yes","Only PCT Entered","Niether Entered")))) %>% 
  select(2,7,9,11,13,20) 
trimEDC <- left_join(trimDOV,left_join(trimSD,left_join(trimLBCOLL,trimLBCL)))

trimLabResults <- importLabResults %>%
  mutate(SubjectID= fixTricoreSubjectNumber(Subject.Identifier)) %>% 
  select(15,8,10,11,13,14)%>% 
  mutate(PCTR.Results = ifelse(test=grepl("E",PCTR.Results),yes = as.numeric(substring(PCTR.Results,1,3))/100,ifelse(test = grepl("<",PCTR.Results),yes = PCTR.Results, no = round(as.numeric(PCTR.Results),2)))) %>% 
  mutate(CRP.Result = ifelse(test=grepl("E",CRP.Result),yes = as.numeric(substring(CRP.Result,1,3))/100,ifelse(test = grepl("<",CRP.Result),yes = CRP.Result, no = round(as.numeric(CRP.Result),2))))
  

togetherForRecon <- left_join(trimEDC,trimLabResults)


fileName <- paste("../11.3.4 Output Files/TriCore/New Data to Enter",Sys.Date(),".xlsx")

write.xlsx(newResults,fileName)
