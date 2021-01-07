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

importLBSITE<- read_sas("EDC/LB_SITE.sas7bdat")
importFirstBatch <- read.xlsx("Target Subjects.xlsx")


withoutLactate <- importLBSITE %>% 
  filter(LBSTAT_LAC_MMOLL=="Not Done"&LBSTAT_LAC_MGDL=="Not Done"&
           LBSTAT2_LAC_MGDL=="Not Done"&LBSTAT2_LAC_MMOLL=="Not Done")%>% select(2,4) %>% arrange(SubjectID)


mainSubjects <- importFirstBatch 

targetWithoutLactate <- importLBSITE %>% 
  filter(LBSTAT_LAC_MMOLL=="Not Done"&LBSTAT_LAC_MGDL=="Not Done"&
           LBSTAT2_LAC_MGDL=="Not Done"&LBSTAT2_LAC_MMOLL=="Not Done") %>% 
  filter(SubjectID%in%importFirstBatch$Subject.ID)%>% select(2,4)%>% arrange(SubjectID)

wb <- createWorkbook()
addWorksheet(wb,"All Missing Lactate")
addWorksheet(wb,"Target Missing Lactate")
writeDataTable(wb,1,withoutLactate)
writeDataTable(wb,2,targetWithoutLactate)
setColWidths(wb,1,1:2,widths = "auto")
setColWidths(wb,2,1:2,widths = "auto")
saveWorkbook(wb,"../11.3.4 Output Files/Missing Lactate Samples.xlsx",overwrite = TRUE)