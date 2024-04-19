## VARILUX VENDAS E VOUCHERS
## LAST UPDATE 19.04.2024
## SANDRO JAKOSKA

## LOAD =======================================================================

library(tidyverse)
library(lubridate)
library(reshape2)
library(DBI)
library(openxlsx)

con2 <- dbConnect(odbc::odbc(), "repro",encoding="Latin1")


## CLIENTS ==========================================================================

clientes <- dbGetQuery(con2,"
SELECT DISTINCT C.CLICODIGO,
CLINOMEFANT NOMEFANTASIA, 
C.GCLCODIGO CODGRUPO,
IIF(C.GCLCODIGO IS NULL,C.CLICODIGO || ' ' || CLINOMEFANT,'G' || C.GCLCODIGO || ' ' || GCLNOME) CLIENTE,
GCLNOME GRUPO
FROM CLIEN C
LEFT JOIN GRUPOCLI ON C.GCLCODIGO=GRUPOCLI.GCLCODIGO
")


## SQL ===============================


vlx_2024 <- dbGetQuery(con2,"
WITH FIS AS (SELECT FISCODIGO FROM TBFIS WHERE (FISTPNATOP IN ('V','SR','R') OR FISCODIGO IN ('5.91V','6.91V'))),
    
           CLI AS (SELECT DISTINCT C.CLICODIGO,
                           CLINOMEFANT,
                            ENDCODIGO,
                             SETOR
                              FROM CLIEN C
                               LEFT JOIN (SELECT CLICODIGO,E.ZOCODIGO,ZODESCRICAO SETOR,ENDCODIGO FROM ENDCLI E
                                LEFT JOIN (SELECT ZOCODIGO,ZODESCRICAO FROM ZONA WHERE ZOCODIGO IN (20,21,22,23,24,25,26,27,28))Z ON E.ZOCODIGO=Z.ZOCODIGO WHERE ENDFAT='S')A ON C.CLICODIGO=A.CLICODIGO
                                 WHERE CLICLIENTE='S'),
    
             PED AS (SELECT ID_PEDIDO,
                             P.CLICODIGO,
                              SETOR,
                               PEDDTBAIXA
                                FROM PEDID P
                                  LEFT JOIN CLI C ON P.CLICODIGO=C.CLICODIGO AND P.ENDCODIGO=C.ENDCODIGO
                                   WHERE PEDDTBAIXA BETWEEN '01.01.2024' AND 'TODAY' AND PEDSITPED<>'C' AND PEDLCFINANC IN ('S', 'L','N')),
            
            VLX AS  (SELECT PROCODIGO FROM PRODU WHERE MARCODIGO=57)
            
            SELECT PEDDTBAIXA,
                    CLICODIGO,
                    CASE
 WHEN PD.FISCODIGO IN ('5.91V','6.91V') THEN 1
  ELSE 0 END VOUCHER,
                     SUM(PDPQTDADE)QTD,
                      SUM(PDPUNITLIQUIDO*PDPQTDADE)VRVENDA 
                                     FROM PDPRD PD
                                      INNER JOIN PED P ON PD.ID_PEDIDO=P.ID_PEDIDO
                                       INNER JOIN FIS F ON PD.FISCODIGO=F.FISCODIGO
                                        INNER JOIN VLX VX ON PD.PROCODIGO=VX.PROCODIGO
                                          GROUP BY 1,2,3") 

vlx_2023 <- get(load("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\BASES\\vlx_2023.RData"))

## SUMMARISE DATA ==============================

vlx_2023_2024 <- union_all(vlx_2023,vlx_2024)

vlx_vendas <- 
vlx_2023_2024  %>% inner_join(.,clientes %>% select(CLICODIGO,CLIENTE),by="CLICODIGO") %>%  
group_by(CLIENTE) %>% 
  summarize(
    
    LASTMONTHTHISYEAR=sum(VRVENDA[VOUCHER==0 & floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    THISMONTHLASTYEAR=sum(VRVENDA[VOUCHER==0 &floor_date(PEDDTBAIXA,"day")>=floor_date((Sys.Date()-years(1)), 'month') & floor_date(PEDDTBAIXA,"day")<=floor_date(Sys.Date()-years(1), "day")-1],na.rm = TRUE),
    
    CURRENTMONTH=sum(VRVENDA[VOUCHER==0 & floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date(),'month') %m-% days(1)],na.rm = TRUE),
    
    YTD23=sum(VRVENDA[VOUCHER==0 & floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date() - years(1),"day")-1],na.rm = TRUE),
    
    YTD24=sum(VRVENDA[VOUCHER==0 & floor_date(PEDDTBAIXA,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1],na.rm = TRUE),
    
    MEDIA23=sum(VRVENDA[VOUCHER==0 & floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= 
                          floor_date(Sys.Date()-years(1), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date()-years(1),"year"),
                                        floor_date(Sys.Date()-years(1),"month"),by="month")))),
    
    MEDIA24=sum(VRVENDA[VOUCHER==0 & floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date(), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date(),"year"),
                                        floor_date(Sys.Date(),"day"),by="month")))),
    
    PAST12=sum(VRVENDA[VOUCHER==0 & floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(VRVENDA[VOUCHER==0 & floor_date(PEDDTBAIXA,"day") >  ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)]),
    
    YTD23_VOUCHER=sum(VRVENDA[VOUCHER==1 & floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date() - years(1),"day")-1],na.rm = TRUE) ,
    YTD24_VOUCHER=sum(VRVENDA[VOUCHER==1 & floor_date(PEDDTBAIXA,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1],na.rm = TRUE) 
    
    ) %>% 
   mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) %>% 
  
  mutate(VAR24_VLR=ifelse(is.finite(YTD24-YTD23),YTD24-YTD23,0)) %>% 
  
  mutate(VAR24_VOUCHER=ifelse(is.finite(YTD24_VOUCHER/YTD23_VOUCHER-1),YTD24_VOUCHER/YTD23_VOUCHER-1,0)) %>% 
  
  mutate(VAR24_VLR_VOUCHER=ifelse(is.finite(YTD24_VOUCHER-YTD23_VOUCHER),YTD24_VOUCHER-YTD23_VOUCHER,0)) %>% 
  
  mutate(STATUS=case_when(
    VAR24>0 ~ 'CRESCIMENTO',
    YTD24!=0 & VAR24<0 ~ 'QUEDA',
    YTD23==0 & YTD24>0 ~ 'RECUPERADO',
    YTD23>0 & YTD24==0 ~ 'PERDIDO',
    YTD23==0 & YTD24==0 ~ 'SEM RECEITA',
    TRUE ~ ''
  )) %>% 
  
  mutate(STATUS_V=case_when(
    VAR24_VLR_VOUCHER >0 ~ 'AUMENTO',
    VAR24_VLR_VOUCHER <0 ~ 'QUEDA',
    TRUE ~ ''
  )) %>% 
  
  arrange(desc(.$YTD24)) 

LASTMONTHTHISYEAR <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))

THISMONTHLASTYEAR <- toupper(format(floor_date(Sys.Date()-years(1), "month"),"%b%/%Y"))

corder <- c("CLIENTE",LASTMONTHTHISYEAR,THISMONTHLASTYEAR,CURRENTMONTH,"YTD23","YTD24","VAR24","VAR24_VLR","STATUS","MEDIA23","MEDIA24","YTD23_VOUCHER","YTD24_VOUCHER","VAR24_VOUCHER","VAR24_VLR_VOUCHER","STATUS_V")

vlx_vendas2 <- vlx_vendas %>% arrange(desc(.$YTD24)) %>% as.data.frame() %>% 
  rename_at(2:4,~ c(LASTMONTHTHISYEAR,THISMONTHLASTYEAR,CURRENTMONTH)) %>% .[,corder]



## CHART ====================================

vlx_plot <-
vlx_2023_2024 %>%
  filter((floor_date(PEDDTBAIXA, "day") >= floor_date(Sys.Date() - years(1), "year") & 
            floor_date(PEDDTBAIXA, "day") <= floor_date(Sys.Date() - years(1), "day") - 1) |
           (floor_date(PEDDTBAIXA, "day") >= floor_date(Sys.Date(), "year") & 
              floor_date(PEDDTBAIXA, "day") <= floor_date(Sys.Date(), "day") - 1)) %>%
  group_by(ANO = format(floor_date(PEDDTBAIXA, "year"), "%Y"), VOUCHER) %>%
  summarize(VALOR = round(sum(VRVENDA), 0)) %>% mutate(VOUCHER=if_else(VOUCHER==1,"VOUCHER","VENDAS")) %>% 
  mutate(VOUCHER = as.factor(VOUCHER)) %>%  # Convert VOUCHER to factor

ggplot(aes(x = as.factor(ANO), y = VALOR, fill = VOUCHER)) +
  geom_bar(stat = "identity", position = "dodge") +  # Use position = "dodge" for side-by-side bars
  geom_text(aes(label = scales::number_format(big.mark = ".", decimal.mark = ",")(VALOR)),
            position = position_dodge(width = 0.9), vjust = -0.5, hjust = 0.5, size = 8, color = "white") +
  labs(x = "Ano", title = "VARILUX YTD 23 X 24 | VENDAS E VOUCHERS  ") +
  ylim(0, 13000000) +
  theme(axis.text.y = element_blank(),
        plot.background = element_rect(fill = "#324254"),
        panel.background = element_rect(fill = "#324254"),
        legend.background = element_rect(fill = "#324254"),
        legend.text = element_text(color = "white", size = 20),
        legend.title = element_text(color = "transparent"),
        axis.title.x = element_blank(),
        axis.text = element_text(color = "white"),
        axis.title = element_text(color = "white"),
        axis.title.y.left = element_blank(),
        plot.title = element_text(color = "white", size = 20),
        panel.grid.major.x = element_blank(),  
        panel.grid.minor.x = element_blank(),  
        panel.grid.major.y = element_blank(),  
        panel.grid.minor.y = element_blank(),  
        legend.position = "bottom",
        axis.text.x = element_text(size = 20),
        legend.spacing.x = unit(0.5, "cm")
  ) +
  scale_fill_manual(values = c("#ffd39b", "#115193"))

width <- 11
height <- width * (1.7/3)  # Assuming a 3:2 aspect ratio, adjust as needed

ggsave("vlx_plot.jpg", plot = vlx_plot, device = "jpg", width = width, height = height)


## WRITE EXCEL  ==========================================================================


wkb_vlx_voucher <- createWorkbook()


## sheet ranking

addWorksheet(wkb_vlx_voucher, "RESUMO")

addWorksheet(wkb_vlx_voucher, "GRAFICO")


## format cols

setColWidths(wkb_vlx_voucher, sheet = "RESUMO", cols = 1, widths = 27)

setColWidths(wkb_vlx_voucher, sheet = "RESUMO", cols = 2:8, widths = 12)

setColWidths(wkb_vlx_voucher, sheet = "RESUMO", cols = 9, widths = 15)

setColWidths(wkb_vlx_voucher, sheet = "RESUMO", cols = 10:11, widths = 12)

setColWidths(wkb_vlx_voucher, sheet = "RESUMO", cols = 12:16, widths = 19)


crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

recuperado <- createStyle(fontColour = "#132782", bgFill = "#6074cf")

perdido <- createStyle(fontColour = "#a60000", bgFill = "#f34c4c")


## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wkb_vlx_voucher, sheet = "RESUMO", style = estiloNumerico, cols = c(2:6,8,10:15), rows = 1:3000, gridExpand = TRUE)

addStyle(wkb_vlx_voucher, sheet = "RESUMO", style = estiloPorcentagem, cols = c(7,14), rows = 1:3000, gridExpand = TRUE)


## cond format class

conditionalFormatting(wkb_vlx_voucher, sheet = "RESUMO", cols = 9, rows = 1:1000, rule = 'I1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wkb_vlx_voucher, sheet = "RESUMO", cols = 9, rows = 1:1000, rule = 'I1 == "QUEDA"', style = queda)

conditionalFormatting(wkb_vlx_voucher, sheet = "RESUMO", cols = 9, rows = 1:1000, rule = 'I1 == "RECUPERADO"', style = recuperado)

conditionalFormatting(wkb_vlx_voucher, sheet = "RESUMO", cols = 9, rows = 1:1000, rule = 'I1 == "PERDIDO"', style = perdido)

conditionalFormatting(wkb_vlx_voucher, sheet = "RESUMO", cols = 7, rows = 1:1000, rule = 'G1 < 0', style = queda2)

conditionalFormatting(wkb_vlx_voucher, sheet = "RESUMO", cols = 8, rows = 1:1000, rule = 'H1 < 0', style = queda2)

conditionalFormatting(wkb_vlx_voucher, sheet = "RESUMO", cols = 16, rows = 1:1000, rule = 'P1 == "AUMENTO"', style = queda)

conditionalFormatting(wkb_vlx_voucher, sheet = "RESUMO", cols = 16, rows = 1:1000, rule = 'P1 == "QUEDA"', style = crescimento)


writeDataTable(wkb_vlx_voucher, "RESUMO", vlx_vendas2, startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

insertImage(wkb_vlx_voucher, "GRAFICO", "vlx_plot.jpg",units = "in", width = 7, height =4,startRow = 1, startCol = 1)

Styling_object <- createStyle(bgFill = "#f9fcff")

conditionalFormatting(wkb_vlx_voucher, sheet = "GRAFICO", cols = c(1:20), rows = 1:1000, rule = 'I1 == ""', style = Styling_object )


## save workbook

saveWorkbook(wkb_vlx_voucher, file = "VENDAS_VARILUX_E_VLX_VOUCHER.xlsx", overwrite = TRUE)


