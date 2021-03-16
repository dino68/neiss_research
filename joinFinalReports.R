library(openxlsx)

# set working directory
setwd("/Users/dayoorigunwa/Desktop/NEU/NEU MS Program Data Analytics/Spring 2021/IE 7275/semProject/NEISS-Data/data/NEISS_Final_Datasets/")

# import female 2019 injury report
f_report <- read.xlsx("NEISS_2019.XLSX")

# import male 2019 injury report
m_report <- read.xlsx("NEISS_2019 2.XLSX")

# bind datasets by row
final_report <- rbind(f_report, m_report)

# write datasets to new file & save to file
wb <- createWorkbook()
addWorksheet(wb, "Worksheet 1")
writeData(wb, sheet=1, final_report, colNames = TRUE)
saveWorkbook(wb, "Final NEISS Data Reports.xlsx", overwrite = TRUE)
