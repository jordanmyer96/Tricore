readyTracker <- cleanPatientTracker %>% filter(Subject%in%readySubjects) %>% arrange(Subject)
orderTable <- data.frame(Var1 = c("Pending Data Entry","Unanswered Queries","SDV Required",
                                  "DM Review Required","Pending Lab Data Entry","Ready For Adjudication"))
primaryFirstBatchStatus <- as.data.frame(table(primaryFirstBatch$Status))%>% rbind(data.frame(Var1 = "Ready For Adjudication",Freq = 0))
primaryFirstBatchStatus <- left_join(orderTable,primaryFirstBatchStatus)%>% 
  set_names(c("Subject Status","Count"))


readySDTable <- as.data.frame(table(readyTracker %>% filter(`Subject Disposition`!="") %>% pull(`Subject Disposition`))) %>% set_names(c("Subject Disposition","Count")) %>% arrange(-Count)

readyStatusTable <-as.data.frame(table(readyTracker$Status)) %>% 
  rbind(data.frame(Var1 = "Ready For Adjudication",Freq = 0)) 

readyStatusTable <- left_join(orderTable,readyStatusTable) %>% 
  set_names(c("Subject Status","Count"))

readWB <- createWorkbook()
addWorksheet(readWB,"Clean Patient Tracker")

writeDataTable(readWB,1,readyTracker)
writeDataTable(readWB,1,readySDTable,startCol = 14,withFilter = FALSE)
writeDataTable(readWB,1,readyStatusTable,startCol = 14,startRow = nrow(readySDTable)+4,withFilter = FALSE)
setColWidths(readWB,1,c(1:ncol(readyTracker),14,15),widths = "auto")
setColWidths(readWB,1,cols = c(4:11,15),widths = c(20,19,16,16,20,21,22,23,8))

saveWorkbook(readWB,"../11.3.4 Output Files/Outgoing Weekly Reports/Ready Clean Patient Tracker.xlsx",overwrite = TRUE)