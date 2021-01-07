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

importMBCL <- read_sas("EDC/MB_CL.sas7bdat")
importLBCL <- read_sas("EDC/LB_CL.sas7bdat")
importLBCOLL <- read_sas("EDC/LB_COLL.sas7bdat")
importLabTransfer <- read.xlsx(MostRecentFile("Lab Results/",".*Tricore Results.*xlsx$","ctime")) %>% 
  mutate(Subject.Identifier = fixTricoreSubjectNumber(Subject.Identifier))

importOldLabTransfer <- read.xlsx("Lab Results/Tricore Results 3 Nov 2020.xlsx")%>% 
  mutate(Subject.Identifier = fixTricoreSubjectNumber(Subject.Identifier))


#newSubjects <- importLabTransfer$Subject.Identifier[!importLabTransfer$Subject.Identifier%in%importOldLabTransfer$Subject.Identifier]
#newSubjects[newSubjects%in%mainSubjects]

importFirstBatch <- read.xlsx("Target Subjects.xlsx")
mainSubjects <- importFirstBatch %>% 
  filter(Replacement=="Primary") %>% 
  pull(Subject.ID)


RESPANSubj <- importLBCOLL %>% 
  filter(LBPERF2_NP=="Yes"&SubjectID%in%mainSubjects)%>% 
  select(2) %>% 
  arrange(SubjectID)

noRespanSubj<- importLBCOLL %>% 
  filter(LBPERF2_NP!="Yes"&SubjectID%in%mainSubjects)%>% 
  select(2)%>% 
  arrange(SubjectID)


noLabResults <- importFirstBatch %>% 
  filter(Replacement=="Primary") %>% 
  filter(!Subject.ID%in%importLabTransfer$Subject.Identifier)%>% 
  select(2)%>% 
  arrange(Subject.ID)


ready <- noRespanSubj %>% 
  filter(SubjectID%in%importLabTransfer$Subject.Identifier)


write.xlsx(ready,"../11.3.4 Output Files/HasLabDoesntNeedRESPAN.xlsx")

needBoth <- RESPANSubj$SubjectID[RESPANSubj$SubjectID%in%noLabResults$Subject.ID]


readySubjects <- ready$SubjectID


onlyRESPAN <- RESPANSubj %>% 
  filter(!SubjectID%in%needBoth) 



onlyLab <- noLabResults %>% 
  filter(!Subject.ID%in%needBoth)


wb <- createWorkbook()
addWorksheet(wb,"Missing Lab and RESPAN")
writeDataTable(wb,1,DFNeedBoth,startRow = 2)
writeDataTable(wb,1,onlyRESPAN,startRow = 2,startCol = 3)
writeDataTable(wb,1,onlyLab,startRow = 2,startCol = 5)
bad <- createStyle(fontColour = "#9C0006" ,fgFill = "#FFC7CE")
yellow <- createStyle(fontColour = "#9C5700" ,fgFill = "#FFEB9C")
addStyle(wb,1,bad,rows = c(3:(nrow(DFNeedBoth)+2)),cols = 1)
addStyle(wb,1,yellow,rows=c(3:(nrow(onlyRESPAN)+2)),cols = 3)
setColWidths(wb,1,cols = c(1,3,5),widths = c(12,12,12))

saveWorkbook(wb,"../11.3.4 Output Files/Tricore Target Subject Statuses.xlsx",overwrite = TRUE)
# nrow(noLabResults)+
#nrow(noLabResults)+1+nrow(noMBResults)


