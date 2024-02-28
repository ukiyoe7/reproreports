
library(DBI)
library(dplyr)
library(readr)
library(reshape2)
library(lubridate)
library(openxlsx)
library(ggplot2)
library(reshape2)
library(bindata)

con2 <- dbConnect(odbc::odbc(), "repro",encoding="Latin1")

## HEAD ==================================================

data_atual <- paste0("DATA: ",format(Sys.Date(),"%d/%m/%Y"))

## PORTAS ATIVAS ==================================================


portas_ativas <- dbGetQuery(con2, statement = read_file('SQL/PORTAS_ATIVAS.sql'))


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


giro3 <-
giro %>% select(MES,ANO,GIRO) %>% dcast(MES ~ ANO)

giro4 <- "GIRO"


## CLIENTES AUSENTES ================================================


# missing 2 months ago compared to 3 months ago
ausentes_penultimo_mes <-
  anti_join(
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date() - months(3), "month"))) %>% 
      distinct(CLICODIGO) 
    ,
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date() - months(2), "month"))) %>% 
      distinct(CLICODIGO)
    ,by="CLICODIGO") 


# missing 1 months ago compared to 2 months ago
ausentes_ultimo_mes <-
  anti_join(
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date() - months(2), "month"))) %>% 
      distinct(CLICODIGO) 
    ,
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date() - months(1), "month"))) %>% 
      distinct(CLICODIGO)
    ,by="CLICODIGO") 


#  missing current month compared to 1 month ago

ausentes_mes_atual <-
  anti_join(
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date() - months(1), "month"))) %>% 
      distinct(CLICODIGO) 
    ,
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date(), "month"))) %>% 
      distinct(CLICODIGO)
    ,by="CLICODIGO")
  
  
ausentes_mes_atual2 <-  
  anti_join(ausentes_mes_atual,inner_join(ausentes_penultimo_mes,ausentes_mes_atual,by="CLICODIGO"),by="CLICODIGO")


ausentes_penultimo_mes2 <-  
  anti_join(ausentes_penultimo_mes,inner_join(ausentes_penultimo_mes,portas_ativas %>% 
                                                filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date(), "month"))) %>% 
                                                distinct(CLICODIGO),by="CLICODIGO"),by="CLICODIGO")


ausentes_penultimo_mes3 <-  
  anti_join(ausentes_penultimo_mes,inner_join(ausentes_penultimo_mes2,portas_ativas %>% 
                                                filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date()- months(1), "month"))) %>% 
                                                distinct(CLICODIGO),by="CLICODIGO"),by="CLICODIGO")

# max value df
max_rows <- max(nrow(ausentes_penultimo_mes3), nrow(ausentes_ultimo_mes), nrow(ausentes_mes_atual2))


# complete missing values and bind columns
ausentes <-
cbind(
ausentes_penultimo_mes3 %>%  rbind(data.frame(CLICODIGO = rep(NA, max_rows - nrow(.)))),
ausentes_ultimo_mes %>%  rbind(data.frame(CLICODIGO= rep(NA, max_rows - nrow(.)))),
ausentes_mes_atual2 %>%  rbind(data.frame(CLICODIGO = rep(NA, max_rows - nrow(.))))) %>% as.data.frame() 


ausentes2 <-
ausentes %>% 
  setNames(c(format(Sys.Date() - months(2), "%B"), 
             format(Sys.Date() - months(1), "%B"), 
             format(Sys.Date(), "%B"))) 



## CHART PORTAS ATIVAS ==================================================

portas_ativas_plot <-
portas_ativas2 %>% 
  ggplot(.,aes(x=MES,y=PORTAS_ATIVAS,color=as.factor(ANO),group = as.factor(ANO))) + geom_line() +
  scale_color_manual(values = c("2023" = "#8297b0", "2024" = "#faf31e"))  +
  geom_text(aes(label = PORTAS_ATIVAS), vjust = -0.5, hjust = 1,size=5) +
  geom_point(size = 4.5) +
  theme_minimal() +
  theme(
    plot.background = element_rect(fill = "#324254"),
    panel.background = element_rect(fill = "#324254"),
    legend.background = element_rect(fill = "#324254"),
    legend.title = element_text(color = "white"),
    axis.title.x = element_blank(),
    legend.text = element_text(color = "white", size = 13), 
    axis.text = element_text(color = "white"),
    axis.title = element_text(color = "white"),
    axis.title.y.left = element_text(size = 15),
    plot.title = element_text(color = "white"),
    panel.grid.major.y = element_blank(),
    panel.grid.minor.y = element_blank(),
    panel.grid.major.x = element_line(colour = "#425266"),
    legend.position = "top",
    axis.text.x = element_text(size = 15),
    legend.key.size = unit(2, "lines"),
    axis.text.y = element_blank()
  ) +
  guides(color = guide_legend(title = NULL))


ggsave("portas_ativas_plot.jpg", plot = portas_ativas_plot, device = "jpg",width = 11, height =6)


## CHART GIRO ==================================================

giro_plot <-
  giro %>% 
  ggplot(.,aes(x=MES,y=GIRO,color=as.factor(ANO),group = as.factor(ANO))) + 
  geom_line() +
  geom_point(size = 4.5) +
  scale_color_manual(values = c("2023" = "#8297b0", "2024" = "#00ffff"))  +
  geom_text(aes(label = GIRO), vjust = -0.5, hjust = 1,size=5) +
  theme_minimal() +
  theme(
    plot.background = element_rect(fill = "#324254"),
    panel.background = element_rect(fill = "#324254"),
    legend.background = element_rect(fill = "#324254"),
    legend.title = element_text(color = "white"),
    axis.title.x = element_blank(),
    legend.text = element_text(color = "white", size = 13), 
    axis.text = element_text(color = "white"),
    axis.title = element_text(color = "white"),
    axis.title.y.left = element_text(size = 15),
    plot.title = element_text(color = "white"),
    panel.grid.major.y = element_blank(),
    panel.grid.minor.y = element_blank(),
    panel.grid.major.x = element_line(colour = "#425266"),
    legend.position = "top",
    axis.text.x = element_text(size = 15),
    legend.key.size = unit(2, "lines"),
    axis.text.y = element_blank()
  ) +
  guides(color = guide_legend(title = NULL))


ggsave("giro_plot.jpg", plot = giro_plot, device = "jpg",width = 11, height =6)



## EXCEL ==================================================


wb <- createWorkbook()


## DADOS

addWorksheet(wb, "DADOS")

# DATA ATUAL

writeData(wb, "DADOS", data_atual, startCol = 2, startRow = 1)


writeData(wb, "DADOS", portas_ativas4, startCol = 2, startRow = 3)

writeDataTable(wb, "DADOS", portas_ativas3, startRow = 4, startCol = 2,tableStyle = "TableStyleMedium20")


writeData(wb, "DADOS", qtd_pcs4, startCol = 6, startRow = 3)

writeDataTable(wb, "DADOS", qtd_pcs3, startRow = 4, startCol = 6, tableStyle = "TableStyleMedium21")


writeData(wb, "DADOS", giro4, startCol = 10, startRow = 3)

writeDataTable(wb, "DADOS", giro3, startRow = 4, startCol = 10, tableStyle = "TableStyleMedium19")


## CLIENTES AUSENTES ===========================================


addWorksheet(wb, "CLIENTES AUSENTES")


writeDataTable(wb, "CLIENTES AUSENTES", ausentes2, startRow = 4, startCol = 2,tableStyle = "TableStyleMedium20")



## PORTAS ATIVAS =========================================


addWorksheet(wb, "PORTAS ATIVAS")

insertImage(wb, "PORTAS ATIVAS", "portas_ativas_plot.jpg",units = "in", width = 9, height =5,startRow = 1, startCol = 1)


## GIRO =================================================


addWorksheet(wb, "GIRO")

insertImage(wb, "GIRO", "giro_plot.jpg",units = "in", width = 9, height =5,startRow = 1, startCol = 1)



## SAVE

saveWorkbook(wb, "PORTAS_ATIVAS.xlsx", overwrite = TRUE)



