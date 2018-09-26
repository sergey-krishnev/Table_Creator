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
  
  Table_edited_vector <- c(170,190,260,270,290,490,510,590,690,700)
  
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
  #formula
  Kabsl <- (Table_edited_vector_first[3]+Table_edited_vector_first[4])/Table_edited_vector_first[9]
  Ktecl <- Table_edited_vector_first[5]/Table_edited_vector_first[9]
  Kfn <- Table_edited_vector_first[6]/Table_edited_vector_first[10]
  Kszs <- Table_edited_vector_first[6]/Table_edited_vector_first[8]
  Kosos <- (Table_edited_vector_first[6]+Table_edited_vector_first[8]-Table_edited_vector_first[2])/Table_edited_vector_first[5]
  Ksok <- Table_edited_vector_first[1]/Table_edited_vector_first[7]
  Table_edited_vector_formula_first <- c(0,Kabsl,Ktecl,0,Kfn,0,Kszs,Kosos,Ksok)
  return(Table_edited_vector_formula_first)
}
#paste data
setwd("D:/Program Files/TableCreator by ForYou/input_documents/Accountant_balance")
Name_Inf <- dir()
Name_Inf_data <- gsub("\\D","",Name_Inf)

First_column <- Load_AB_from_excel(Name_Inf[1])
#if (as.numeric(Name_Inf_data[1]) <= 2015) {First_column <- First_column/10} #учет деноминации

Third_column <- Load_AB_from_excel(Name_Inf[2])
#if (as.numeric(Name_Inf_data[2]) <= 2015) {Third_column <- Third_column/10}
Fifth_column <- Load_AB_from_excel(Name_Inf[3])
#if (as.numeric(Name_Inf_data[3]) <= 2015) {Fifth_column <- Fifth_column/10}

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
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 4,col = x, cellstyle = UpElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 5,col = x, cellstyle = MiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 6,col = x, cellstyle = DownElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 7,col = x, cellstyle = UpElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 8,col = x, cellstyle = MiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 9,col = x, cellstyle = MiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 10,col = x, cellstyle = MiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 11,col = x, cellstyle = MiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 12,col = x, cellstyle = DownElement_style)
}

stylizingPercent <- function(x){
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 4,col = x, cellstyle = PerUpElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 5,col = x, cellstyle = PerMiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 6,col = x, cellstyle = PerDownElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 7,col = x, cellstyle = PerUpElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 8,col = x, cellstyle = PerMiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 9,col = x, cellstyle = PerMiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 10,col = x, cellstyle = PerMiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 11,col = x, cellstyle = PerMiddleElement_style)
  setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 12,col = x, cellstyle = PerDownElement_style)
}
#COMMON OUTPUT
writeWorksheet(Otchet_vivod_dan,First_column,sheet = "T13",4,2,F)
writeWorksheet(Otchet_vivod_dan,"",sheet = "T13",4,2,F)
writeWorksheet(Otchet_vivod_dan,"",sheet = "T13",7,2,F)
stylizingPercent(2)
writeWorksheet(Otchet_vivod_dan,Third_column,sheet = "T13",4,3,F)
writeWorksheet(Otchet_vivod_dan,"",sheet = "T13",4,3,F)
writeWorksheet(Otchet_vivod_dan,"",sheet = "T13",7,3,F)
stylizingPercent(3)
writeWorksheet(Otchet_vivod_dan,Fifth_column,sheet = "T13",4,4,F)
writeWorksheet(Otchet_vivod_dan,"",sheet = "T13",4,4,F)
writeWorksheet(Otchet_vivod_dan,"",sheet = "T13",7,4,F)
stylizingPercent(4)

#COMMON OUTPUT DATE
writeWorksheet(Otchet_vivod_dan,Name_Inf_data[1],sheet = "T13",3,2,F)
setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 3,col = 2, cellstyle = OneElement_style)
writeWorksheet(Otchet_vivod_dan,Name_Inf_data[2],sheet = "T13",3,3,F)
setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 3,col = 3, cellstyle = OneElement_style)
writeWorksheet(Otchet_vivod_dan,Name_Inf_data[3],sheet = "T13",3,4,F)
setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 3,col = 4, cellstyle = OneElement_style)

#COMMON OUTPUT SUBTRACT TITLE
writeWorksheet(Otchet_vivod_dan,paste(Name_Inf_data[2],"г. от ",Name_Inf_data[1],"г.",sep=""),sheet = "T13",3,5,F)
setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 3,col = 5, cellstyle = OneElement_style)
writeWorksheet(Otchet_vivod_dan,paste(Name_Inf_data[3],"г. от ",Name_Inf_data[2],"г.",sep=""),sheet = "T13",3,6,F)
setCellStyle(Otchet_vivod_dan,sheet = "T13",row = 3,col = 6, cellstyle = OneElement_style)
#COMMON OUTPUT SUBTRACT
Seventh_column <- Third_column - First_column
Eigthth_column <- Fifth_column - Third_column
writeWorksheet(Otchet_vivod_dan,Seventh_column,sheet = "T13",4,5,F)
writeWorksheet(Otchet_vivod_dan,"",sheet = "T13",4,5,F)
writeWorksheet(Otchet_vivod_dan,"",sheet = "T13",7,5,F)
stylizingPercent(5)
writeWorksheet(Otchet_vivod_dan,Eigthth_column,sheet = "T13",4,6,F)
writeWorksheet(Otchet_vivod_dan,"",sheet = "T13",4,6,F)
writeWorksheet(Otchet_vivod_dan,"",sheet = "T13",7,6,F)
stylizingPercent(6)





saveWorkbook(Otchet_vivod_dan)