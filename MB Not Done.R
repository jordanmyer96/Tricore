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

importLBCOLL <- read_sas("EDC/LB_COLL.sas7bdat")
importMBCL <- read_sas("EDC/MB_CL.sas7bdat")
importLBCL <- read_sas("EDC/LB_CL.sas7bdat")
importSD <- read_sas("EDC/DS_SD.sas7bdat")
#importLabToENter <- read.xlsx("../11.3.4 Output Files/TriCore/New Lab Data to Enter.xlsx")

alreadyNotDone <- importMBCL %>% filter(MBSTAT_VIRAL=="Not Done") 
noRESPAN <- importLBCOLL %>% 
  filter(LBPERF2_NP!="Yes"&!SubjectID%in%alreadyNotDone$SubjectID) %>% 
  select(2,31) %>% 
  arrange(SubjectID)

doneWithStudy <- importSD %>% filter(DSDECOD!="") %>% 
  pull(SubjectID)

toMarkNotDone <- noRESPAN %>% 
  filter(SubjectID%in%doneWithStudy&SubjectID%in%importLBCL$SubjectID) %>% 
  arrange(SubjectID)




withNotDone <- left_join(importLabToENter,noRESPAN,by = c("Subject.Identifier"="SubjectID")) 
withNotDone$MBNotDone <- ""
withNotDone[!is.na(withNotDone$LBPERF2_NP),"MBNotDone"] <- "Mark MB As Not Done"

write.xlsx(toMarkNotDone,"../11.3.4 Output Files/TriCore/To Mark Not Done.xlsx")
write.xlsx(withNotDone,"../11.3.4 Output Files/TriCore/Mark Not Done.xlsx")

