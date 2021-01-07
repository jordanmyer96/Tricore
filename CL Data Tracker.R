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
importFirstBatch <- read.xlsx("Target Subjects.xlsx")

importLabResults$Subject.Identifier[duplicated(importLabResults$Subject.Identifier)]
unique(importMBResults$Subject.Number)
#Make Combined EDC ----
trimDOV <- importDOV %>% 
  filter(Visit=="Day 0") %>% 
  select(2,4)

trimLBCOLL <- importLBCOLL %>% 
  select(2,4,17,31)

trimSD <- importSD%>% 
  select(2,4,7)

trimLBCL <- importLBCL %>% 
  mutate(CRPEntered = ifelse(!is.na(LBORRES_CRP)|LBSTAT_CRP=="Not Done","Yes","No")) %>% 
  mutate(PCTEntered = ifelse(!is.na(LBORRES_PCT)|LBSTAT_PCT=="Not Done","Yes","No")) %>%  
  mutate(`Lab Entry Status` = ifelse(CRPEntered=="Yes"&PCTEntered=="Yes","Complete",
                                     ifelse(CRPEntered=="Yes"&PCTEntered=="No","Only CRP Entered",
                                            ifelse(CRPEntered=="No"&PCTEntered=="Yes","Only PCT Entered","Niether Entered")))) %>% 
  select(2,20) 

trimMBCL <- importMBCL %>% 
  mutate(`MB Entry Status` = ifelse(test = MBSTAT_VIRAL=="Not Done","Marked Not Done",
                             ifelse(test = MBRESA_VIRAL=="Negative","Marked as Negative",
                                    ifelse(test = !is.na(MBRESB_VIRAL),"Positive Pathogens Entered","Form Triggered but blank")))) %>% 
  select(2,20)

trimEDC <- left_join(trimDOV,left_join(trimSD,left_join(trimLBCOLL,trimLBCL)))
trimEDC <- left_join(trimEDC,trimMBCL)



#Make Tricore Lab Results----
trimLabResults <- importLabResults %>%
  mutate(SubjectID= fixTricoreSubjectNumber(Subject.Identifier)) %>% 
  select(15,8,10,11,13,14)%>% 
  mutate(PCTR.Results = ifelse(test=grepl("E",PCTR.Results),yes = as.numeric(substring(PCTR.Results,1,3))/100,ifelse(test = grepl("<",PCTR.Results),yes = PCTR.Results, no = round(as.numeric(PCTR.Results),2)))) %>% 
  mutate(CRP.Result = ifelse(test=grepl("E",CRP.Result),yes = as.numeric(substring(CRP.Result,1,3))/100,ifelse(test = grepl("<",CRP.Result),yes = CRP.Result, no = round(as.numeric(CRP.Result),2)))) %>% 
  mutate(LabResultsReceived = "Results Received")


#Make Tricore MB Results----
positiveMBResults <- importMBResults %>% 
  mutate(SubjectID= fixTricoreSubjectNumber(Subject.Number)) %>% 
  filter(Result!="Negative")

subjectNumber <- c()
positivePathogens <- c()
for(i in 1:length(unique(positiveMBResults$Subject.Number))){
  subjectNumber[i] <- unique(positiveMBResults$Subject.Number)[i]
  oneSubjectPositive <- positiveMBResults %>% filter(Subject.Number==unique(positiveMBResults$Subject.Number)[i])
  positivePathogens[i] <- c(paste(oneSubjectPositive$Pathogen,collapse = ", "))
}

positivePathogenDF <- data.frame(SubjectID = subjectNumber,PositivePathogens=positivePathogens) 

trimMBResults <- importMBResults %>% 
  mutate(SubjectID= fixTricoreSubjectNumber(Subject.Number)) %>% 
  filter(Result!="Negative")

MBResultsDF <- left_join(data.frame(SubjectID = unique(importMBResults$Subject.Number)),positivePathogenDF) %>% 
  mutate(SubjectID= fixTricoreSubjectNumber(SubjectID)) 

MBResultsDF[is.na(MBResultsDF$PositivePathogens),"PositivePathogens"] <- "All Negative"
MBResultsDF$MBResultsReceived <- "MB Results Received"

#Combine everything----
trimLabResults$SubjectID[!trimLabResults$SubjectID%in%trimEDC$SubjectID]

together <- left_join(trimEDC,trimLabResults)
allTogether <- left_join(together,MBResultsDF)
final <- allTogether %>% 
  select(c(1,2,3,4,13,6,10,11,8,9,5,15,7,14,12)) %>% 
  set_names(c("Subject ID","Subject Status","Reason for Completion","CRP/PCT Collected","Lab Results Received",
              "Lab Entry Status","CRP Results","CRP Not Done","PCTR Results","PCTR Not Done","MB Swab Collected",
              "MB Results Received","MB Entry Status","Positive Pathogens","DM Comments")) %>% 
  arrange(`Subject ID`)

wb <- createWorkbook()
addWorksheet(wb,"CL Data Tracker")
writeDataTable(wb,1,final)
setColWidths(wb,1,1:ncol(final),widths = "auto")


mainSubjects <- importFirstBatch %>% 
  filter(Replacement=="Primary") %>% 
  pull(Subject.ID)
addWorksheet(wb,"Target 100 CL Data Tracker")
writeDataTable(wb,2,final %>% filter(`Subject ID`%in%mainSubjects))
setColWidths(wb,2,1:ncol(final %>% filter(`Subject ID`%in%mainSubjects)),widths = "auto")

fileName <- paste("../11.3.4 Output Files/TriCore/CL Data Tracker",Sys.Date(),".xlsx")
saveWorkbook(wb,fileName,overwrite = T)
