
library(DBI)
library(dplyr)
library(readr)
library(reshape2)
library(lubridate)
library(openxlsx)
library(ggplot2)
library(reshape2)

con2 <- dbConnect(odbc::odbc(), "repro",encoding="Latin1")


## PORTAS ATIVAS ==================================================


portas_ativas <- dbGetQuery(con2, statement = read_file('SQL/PORTAS_ATIVAS.sql'))

View(portas_ativas)

portas_ativas2 <- 
 portas_ativas %>% 
   group_by(MES = format(floor_date(PEDDTBAIXA, "month"), "%b"),ANO=year(floor_date(PEDDTBAIXA,"year"))) %>% 
     summarize(PORTAS_ATIVAS=n_distinct(CLICODIGO)) %>% 
    mutate(MES=factor(MES, levels = c("jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"), ordered = TRUE))

portas_ativas3 <-
portas_ativas2 %>%  dcast(MES ~ ANO)

portas_ativas4 <- "PORTAS ATIVAS"

## PCS ==================================================


qtd_pcs <- dbGetQuery(con2, statement = read_file('SQL/QTD_PCS.sql'))

View(qtd_pcs)


qtd_pcs2 <-  
qtd_pcs %>% 
  group_by(MES = format(floor_date(PEDDTBAIXA, "month"), "%b"),ANO=year(floor_date(PEDDTBAIXA,"year"))) %>% 
  summarize(QTD=sum(QTD))  %>% 
  mutate(MES=factor(MES, levels = c("jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"), ordered = TRUE))

qtd_pcs3 <-
qtd_pcs2 %>%  dcast(MES ~ ANO)


qtd_pcs4 <- "QTD PEÇAS"

## GIRO ==================================================

giro <-
left_join(portas_ativas2,qtd_pcs2,by=c("MES","ANO")) %>% mutate(GIRO=round(QTD/PORTAS_ATIVAS,2))

View(giro)

giro3 <-
giro %>% select(MES,ANO,GIRO) %>% dcast(MES ~ ANO)

giro4 <- "GIRO"


## CHART ==================================================

chart1 <-
portas_ativas2 %>% 
  ggplot(.,aes(x=MES,y=PORTAS_ATIVAS,color=as.factor(ANO),group = as.factor(ANO))) + geom_line() +
  scale_color_manual(values = c("2022" = "#8297b0", "2023" = "#faf31e"))  +
  geom_text(aes(label = PORTAS_ATIVAS), vjust = -0.5, hjust = 1) +
  theme_minimal() +
  theme(
    plot.background = element_rect(fill = "#324254"),
    panel.background = element_rect(fill = "#324254"),
    legend.background = element_rect(fill = "#324254"),
    legend.text = element_text(color = "white"),
    legend.title = element_text(color = "white"),
    axis.text = element_text(color = "white"),
    axis.title = element_text(color = "white"),
    plot.title = element_text(color = "white"),
    panel.grid.major.y = element_blank(),
    panel.grid.minor.y = element_blank(),
    panel.grid.major.x = element_line(colour = "#425266"),
    legend.position = "top"
  ) +
  guides(color = guide_legend(title = NULL))


ggsave("chart1.jpg", plot = chart1, device = "jpg",width = 7, height =5)

## EXCEL ==================================================


wb <- createWorkbook()

## DADOS

addWorksheet(wb, "DADOS")


writeData(wb, "DADOS", portas_ativas4, startCol = 1, startRow = 2)
writeData(wb, "DADOS", portas_ativas3, startCol = 1, startRow = 3)

writeData(wb, "DADOS", qtd_pcs4, startCol = 5, startRow = 2)
writeData(wb, "DADOS", qtd_pcs3, startCol = 5, startRow = 3)

writeData(wb, "DADOS", giro4, startCol = 9, startRow = 2)
writeData(wb, "DADOS", giro3, startCol = 9, startRow = 3)

## CHART

addWorksheet(wb, "GRAFICO")

insertImage(wb, "GRAFICO", "chart1.jpg", width = 12, height = 5,units = "in", startRow = 1, startCol = 1)

saveWorkbook(wb, "PORTAS_ABERTAS.xlsx", overwrite = TRUE)



