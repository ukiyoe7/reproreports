
library(DBI)
library(dplyr)
library(readr)
library(reshape2)
library(lubridate)
library(openxlsx)
library(ggplot2)

con2 <- dbConnect(odbc::odbc(), "repro",encoding="Latin1")


## PORTAS ATIVAS ==================================================


portas_ativas <- dbGetQuery(con2, statement = read_file('SQL/PORTAS_ATIVAS.sql'))

View(portas_ativas)

portas_ativas2 <- 
 portas_ativas %>% 
   group_by(MES=floor_date(PEDDTBAIXA,"month"),ANO=year(floor_date(PEDDTBAIXA,"year"))) %>% 
     summarize(PORTAS_ATIVAS=n_distinct(CLICODIGO)) 



## PCS ==================================================


qtd_pcs <- dbGetQuery(con2, statement = read_file('SQL/QTD_PCS.sql'))

View(qtd_pcs)


qtd_pcs2 <-  
qtd_pcs %>% 
  group_by(MES=floor_date(PEDDTBAIXA,"month"),ANO=year(floor_date(PEDDTBAIXA,"year"))) %>% 
  summarize(QTD=sum(QTD)) 



## GIRO ==================================================

giro <-
left_join(portas_ativas2,qtd_pcs2,by=c("MES","ANO")) %>% mutate(GIRO=round(QTD/PORTAS_ATIVAS,2))

View(giro)


## CHART ==================================================

chart <- 
giro %>% ggplot(.,aes(x=MES,y=QTD,fill=ANO)) + geom_bar(stat = "identity",position = "dodge")



ggsave("chart.png", plot = chart, device = "png")



wb <- createWorkbook()

addWorksheet(wb, "DADOS")

writeData(wb, "DADOS", giro, startCol = 1, startRow = 1)

addWorksheet(wb, "Chart")

insertImage(wb, "Chart", "chart.png", width = 12, height = 10, units = "in", startRow = 2, startCol = 2)

saveWorkbook(wb, "output.xlsx", overwrite = TRUE)



