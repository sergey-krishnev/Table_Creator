@echo off
d:
cd D:\Program Files\TableCreator by ForYou\output_documents
del *txt
cd D:\Program Files\TableCreator by ForYou
Rscript CommonTableCreator.R
cd D:\Program Files\TableCreator by ForYou\table_knowledge_template
copy /Y *xlsx "D:\Program Files\TableCreator by ForYou\output_documents"
