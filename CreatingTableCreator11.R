Load_AB_from_excel <- function(name_of_file) {
  setwd("D:/Program Files/TableCreator by ForYou/input_documents/Accountant_balance")
  Otchet_vhod_dan <- loadWorkbook(name_of_file, create = TRUE)
  Readed_otchet_vhod_dan <- readWorksheet(Otchet_vhod_dan,2)
  Readed_otchet_vhod_dan_stroka <- sapply(Readed_otchet_vhod_dan, as.character)
  Searching_key_word_2and3 <- which(Readed_otchet_vhod_dan_stroka == "Код строки", arr.ind = TRUE)
  Inf_codes_2 <- Readed_otchet_vhod_dan_stroka[(Searching_key_word_2and3[1,1]):(length(Readed_otchet_vhod_dan_stroka[,1])),(Searching_key_word_2and3[1,2])]
  Inf_codes_3 <- Readed_otchet_vhod_dan_stroka[(Searching_key_word_2and3[1,1]):(length(Readed_otchet_vhod_dan_stroka[,1])),(Searching_key_word_2and3[1,2]+1)]
  Searching_key_word_4 <- which(Readed_otchet_vhod_dan_stroka == "4", arr.ind = TRUE)
  Inf_codes_4 <- Readed_otchet_vhod_dan_stroka[(Searching_key_word_4[1,1]):(length(Readed_otchet_vhod_dan_stroka[,1])),(Searching_key_word_4[1,2])]
  
  Table_edited_vector <- c(190,110,120,140,290,210,240,250,270,280,300)
  
  matcher2to3 <- function(x){
    return(Inf_codes_3[which(Inf_codes_2 == x,arr.ind = T)])
  }
  Table_edited_vector_first <- sapply(Table_edited_vector, FUN = matcher2to3) # Создание первой колонки таблицы
  #write.table(Table_edited_vector_first, file = "Table_edited_vector_first.txt", sep = "\t",row.names = FALSE,F)
  #fileVectortest<-file("Table_edited_vector_first.txt")  #Записываем в текст, чтобы потом обратно записать в data.frame
  #writeLines(Table_edited_vector_first, fileVectortest)
  #close(fileVectortest)
  #Table_super_edited_vector_first <- read.table("Table_edited_vector_first.txt", 
  #                 header=TRUE, sep="\t")
  #Избавление от лишних символов
  Table_edited_vector_first <- gsub("\\s","",Table_edited_vector_first)
  Table_edited_vector_first <- gsub("-","",Table_edited_vector_first)
  Table_edited_vector_first <- as.numeric(Table_edited_vector_first)
  return(Table_edited_vector_first)
}
#paste data
setwd("D:/Program Files/TableCreator by ForYou/input_documents/Accountant_balance")
Name_Inf <- dir()
Name_Inf_data <- gsub("\\D","",Name_Inf)

First_column <- Load_AB_from_excel(Name_Inf[1])
if ((as.numeric(Name_Inf_data[1])) <= 2015) {First_column <- First_column/10}
Second_column <- First_column/(First_column[length(First_column)])*100
Third_column <- Load_AB_from_excel(Name_Inf[2])
if ((as.numeric(Name_Inf_data[2])) <= 2015) {Third_column <- Third_column/10}
Fourth_column <- Third_column/(Third_column[length(Third_column)])*100
Fifth_column <- Load_AB_from_excel(Name_Inf[3])
if ((as.numeric(Name_Inf_data[3])) <= 2015) {Fifth_column <- Fifth_column/10}
Sixth_column <- Fifth_column/(Fifth_column[length(Fifth_column)])*100

#paste to excel table
setwd("D:/Program Files/TableCreator by ForYou/table_knowledge_template")
Otchet_vivod_dan <- loadWorkbook("Tables11-13.xlsx", create = TRUE)

#Стили
UpElement_style <- getCellStyle(Otchet_vivod_dan,"UpElement")
MiddleElement_style <- getCellStyle(Otchet_vivod_dan,"MiddleElement")
DownElement_style <- getCellStyle(Otchet_vivod_dan,"DownElement")
OneElement_style <- getCellStyle(Otchet_vivod_dan,"OneElement")

PerUpElement_style <- getCellStyle(Otchet_vivod_dan,"PerUpElement")
PerMiddleElement_style <- getCellStyle(Otchet_vivod_dan,"PerMiddleElement")
PerDownElement_style <- getCellStyle(Otchet_vivod_dan,"PerDownElement")
PerOneElement_style <- getCellStyle(Otchet_vivod_dan,"PerOneElement")
#function stylizing
stylizingStandart <- function(x){
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 4,col = x, cellstyle = UpElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 5,col = x, cellstyle = MiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 6,col = x, cellstyle = MiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 7,col = x, cellstyle = DownElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 8,col = x, cellstyle = UpElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 9,col = x, cellstyle = MiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 10,col = x, cellstyle = MiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 11,col = x, cellstyle = MiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 12,col = x, cellstyle = DownElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 13,col = x, cellstyle = OneElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 14,col = x, cellstyle = OneElement_style)
}

stylizingPercent <- function(x){
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 4,col = x, cellstyle = PerUpElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 5,col = x, cellstyle = PerMiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 6,col = x, cellstyle = PerMiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 7,col = x, cellstyle = PerDownElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 8,col = x, cellstyle = PerUpElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 9,col = x, cellstyle = PerMiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 10,col = x, cellstyle = PerMiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 11,col = x, cellstyle = PerMiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 12,col = x, cellstyle = PerDownElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 13,col = x, cellstyle = PerOneElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 14,col = x, cellstyle = PerOneElement_style)
}
#COMMON OUTPUT
writeWorksheet(Otchet_vivod_dan,First_column,sheet = "T11",4,2,F)
stylizingStandart(2)
writeWorksheet(Otchet_vivod_dan,Second_column,sheet = "T11",4,3,F)
stylizingPercent(3)
writeWorksheet(Otchet_vivod_dan,Third_column,sheet = "T11",4,4,F)
stylizingStandart(4)
writeWorksheet(Otchet_vivod_dan,Fourth_column,sheet = "T11",4,5,F)
stylizingPercent(5)
writeWorksheet(Otchet_vivod_dan,Fifth_column,sheet = "T11",4,6,F)
stylizingStandart(6)
writeWorksheet(Otchet_vivod_dan,Sixth_column,sheet = "T11",4,7,F)
stylizingPercent(7)

#COMMON OUTPUT DATE
writeWorksheet(Otchet_vivod_dan,Name_Inf_data[1],sheet = "T11",2,2,F)
setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 2,col = 2, cellstyle = OneElement_style)
writeWorksheet(Otchet_vivod_dan,Name_Inf_data[2],sheet = "T11",2,4,F)
setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 2,col = 4, cellstyle = OneElement_style)
writeWorksheet(Otchet_vivod_dan,Name_Inf_data[3],sheet = "T11",2,6,F)
setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 2,col = 6, cellstyle = OneElement_style)

#COMMON OUTPUT SUBTRACT TITLE
writeWorksheet(Otchet_vivod_dan,paste(Name_Inf_data[2],"г. от ",Name_Inf_data[1],"г.",sep=""),sheet = "T11",3,8,F)
setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 3,col = 8, cellstyle = OneElement_style)
writeWorksheet(Otchet_vivod_dan,paste(Name_Inf_data[3],"г. от ",Name_Inf_data[2],"г.",sep=""),sheet = "T11",3,9,F)
setCellStyle(Otchet_vivod_dan,sheet = "T11",row = 3,col = 9, cellstyle = OneElement_style)
#COMMON OUTPUT SUBTRACT
Seventh_column <- Third_column - First_column
Eigthth_column <- Fifth_column - Third_column
writeWorksheet(Otchet_vivod_dan,Seventh_column,sheet = "T11",4,8,F)
writeWorksheet(Otchet_vivod_dan,"-",sheet = "T11",14,8,F)
stylizingStandart(8)
writeWorksheet(Otchet_vivod_dan,Eigthth_column,sheet = "T11",4,9,F)
writeWorksheet(Otchet_vivod_dan,"-",sheet = "T11",14,9,F)
stylizingStandart(9)

saveWorkbook(Otchet_vivod_dan)