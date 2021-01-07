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
importIC <- read_sas("EDC/DS_IC.sas7bdat")
importDem <- read_sas("EDC/DM.sas7bdat")
importSD <- read_sas("EDC/DS_SD.sas7bdat")
importEW <- read_sas("EDC/DS_EW.sas7bdat")
importIE <- read_sas("EDC/IE.sas7bdat")
importIM <- read_sas("EDC/IM.sas7bdat")
importLBCOLL <- read_sas("EDC/LB_COLL.sas7bdat")
importLBSITE<- read_sas("EDC/LB_SITE.sas7bdat")
importMBSITE<- read_sas("EDC/MB_SITE.sas7bdat")
importMH<- read_sas("EDC/MH.sas7bdat")
importSOFA<- read_sas("EDC/QS_SOFA.sas7bdat")
importTOE <- read_sas("EDC/QS_TOE.sas7bdat")
importVS <- read_sas("EDC/VS.sas7bdat")
importMBCL <- read_sas("EDC/MB_CL.sas7bdat")
importLBCL <- read_sas("EDC/LB_CL.sas7bdat")

importQueries <- read.xlsx(MostRecentFile("Medrio Reports/",".*Medrio_QueryExport_LIVE_Inflammatix_INF_04.*xlsx$","ctime"))
importVisitDT <- read_sas("EDC/SV.sas7bdat")
importSD <- read_sas("EDC/DS_SD.sas7bdat")


view(importLBCOLL %>% filter(LBYN_HOSTDX=="No"))

view(importLBCOLL %>% filter())
unique(importSD$DSYNSP_COV)

enrolledSubjects <- unique(importVisitDT  %>%  pull(SubjectID))


workingQueries <- importQueries%>%
  filter(SUBJECT.ID%in%enrolledSubjects)

allOpenQueries <- workingQueries%>%
  filter(STATUS=="Open")%>%
  filter(!is.na(CREATED.BY)) %>% 
  filter(QUERY.NAME!="releaseadjudReleasedForAdjudication:Missing Data")

openWithNoAnswer <- allOpenQueries %>% 
  filter(is.na(RESPONSES))

openWithAnswer <- allOpenQueries %>% 
  filter(!is.na(RESPONSES))

closedQueries <- workingQueries%>%
  filter(STATUS!="Open")%>%
  filter(QUERY.NAME!="releaseadjudReleasedForAdjudication:Missing Data")

completedStudy <- importSD %>% filter(DSDECOD=="28 day data collection complete")

saveWorkbook(wb,"../11.3.4 Output Files/Tricore Target Subject Statuses.xlsx",overwrite = TRUE)
