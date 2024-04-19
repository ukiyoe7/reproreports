## RELATORIO RESULTADO ESPERADO POR CLIENTE
## CLIENT: CRISTIANO REGIS
## LAST UPDATE 18.03.2024

## LOAD =======================================================================

library(tidyverse)
library(lubridate)
library(reshape2)
library(DBI)
library(openxlsx)

con2 <- dbConnect(odbc::odbc(), "repro",encoding = "latin1")


sales_2023_esperado<-
  get(load("BASES/sales_2023_esperado.RData"))


##  SQL ==========================================================================

sales_2024_esperado <- dbGetQuery(con2,"
WITH CLI AS (
SELECT DISTINCT C.CLICODIGO,
                  C.GCLCODIGO,
                   GCLNOME,
                    IIF(C.GCLCODIGO IS NULL,C.CLICODIGO || ' ' || CLINOMEFANT,'G' || C.GCLCODIGO || ' ' || GCLNOME) CLIENTE,
                                            SETOR
                                             FROM CLIEN C
                                              INNER JOIN (SELECT CLICODIGO,E.ZOCODIGO,CIDNOME,ENDBAIRRO,ZODESCRICAO SETOR,ENDCODIGO FROM ENDCLI E
                                               LEFT JOIN CIDADE CID ON E.CIDCODIGO=CID.CIDCODIGO
                                                LEFT JOIN (SELECT ZOCODIGO,ZODESCRICAO FROM ZONA WHERE ZOCODIGO IN (20,21,22,23,24,25,26,27,28))Z ON 
                                                 E.ZOCODIGO=Z.ZOCODIGO WHERE ENDFAT='S')A ON C.CLICODIGO=A.CLICODIGO
                                                  LEFT JOIN GRUPOCLI GR ON C.GCLCODIGO=GR.GCLCODIGO
                                                   WHERE CLICLIENTE='S'),
  
FIS AS (SELECT FISCODIGO FROM TBFIS WHERE FISTPNATOP IN ('V','SR','R')),
  
PED AS (SELECT ID_PEDIDO,
                CLIENTE,
                 SETOR,
                PEDDTEMIS,
                 P.CLICODIGO FROM PEDID P
                  INNER JOIN CLI C ON P.CLICODIGO=C.CLICODIGO
                   WHERE PEDSITPED<>'C' AND PEDDTEMIS BETWEEN '01.01.2024' AND 'TODAY')


SELECT PEDDTEMIS,
CLIENTE,
 SETOR,
 SUM(PDPUNITLIQUIDO*PDPQTDADE)VRVENDA  
  FROM PDPRD PD
   INNER JOIN PED P ON PD.ID_PEDIDO=P.ID_PEDIDO
    INNER JOIN FIS F ON PD.FISCODIGO=F.FISCODIGO
     GROUP BY 1,2,3") 


sales_esperado <- union_all(sales_2023_esperado,sales_2024_esperado)



## CALCULATE WORKING DAYS ============================================================

first_day_month <- floor_date(Sys.Date(), unit = "month")

num_days_month <- as.numeric(days_in_month(first_day_month))

last_day <- floor_date(Sys.Date() %m-% days(1), unit = "day")

last_day_of_month <- ceiling_date(Sys.Date() , unit = "month")-1

num_weekends_month <- sum(wday(seq(first_day_month,ceiling_date(Sys.Date() , unit = "month")-1, by = "day")) %in% c(1,7))

num_weekends_last_day <- sum(wday(seq(first_day_month,last_day, by = "day")) %in% c(1,7))

num_weekedays_last_day <- sum(!wday(seq(first_day_month,last_day, by = "day")) %in% c(1,7))

holidays <- data.frame(DATES=c(as.Date('2024-01-01'),
                               as.Date('2024-02-13'),
                               as.Date('2024-03-29'),
                               as.Date('2024-05-01'),
                               as.Date('2024-05-30'),
                               as.Date('2024-11-15'),
                               as.Date('2024-12-25')
))

num_days_working_days_month = num_days_month - num_weekends_month

##  SUMMARY 1 ==========================================================================

sales_sm_1_1 <-
  sales_esperado %>% group_by(CLIENTE) %>% 
  summarize(
    LASTMONTH=sum(VRVENDA[floor_date(PEDDTEMIS,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTEMIS,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    CURRENTMONTH=sum(VRVENDA[floor_date(PEDDTEMIS,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTEMIS,"day")<=floor_date(Sys.Date(),'day') %m-% days(1)],na.rm = TRUE),
    
    MEDIA_2024=sum(VRVENDA[floor_date(PEDDTEMIS,"day")>= floor_date(Sys.Date(), "year") & floor_date(PEDDTEMIS,"day") < floor_date(Sys.Date(), "month")]/as.numeric(length(seq(floor_date(Sys.Date(),"year"),floor_date(Sys.Date(),"month"),by="month"))-1)),
    
    MEDIA_DIAS=sum(VRVENDA[floor_date(PEDDTEMIS,"day")>= floor_date(Sys.Date(), "year") & floor_date(PEDDTEMIS,"day") < floor_date(Sys.Date(), "month")]/as.numeric(length(seq(floor_date(Sys.Date(),"year"),floor_date(Sys.Date(),"month"),by="month"))-1))/num_days_working_days_month,
    
    ESPERADO=sum(VRVENDA[floor_date(PEDDTEMIS,"day")>= floor_date(Sys.Date(), "year") & floor_date(PEDDTEMIS,"day") < floor_date(Sys.Date(), "month")]/as.numeric(length(seq(floor_date(Sys.Date(),"year"),floor_date(Sys.Date(),"month"),by="month"))-1))/num_days_working_days_month*num_weekedays_last_day,
    
    YTD23=sum(VRVENDA[floor_date(PEDDTEMIS,"day") >= floor_date(Sys.Date() %m-% years(1), "year") & floor_date(PEDDTEMIS,"day") <= floor_date(Sys.Date() %m-% years(1), "day")-1],na.rm = TRUE),
    
    YTD24=sum(VRVENDA[floor_date(PEDDTEMIS,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTEMIS,"day") <= floor_date(Sys.Date(), "day")-1],na.rm = TRUE)) %>% as.data.frame() %>% 
  
   mutate(VAR_ATUAL_X_ESPERADO_PERC=CURRENTMONTH/ESPERADO-1) %>%
  
  mutate(VAR_ATUAL_X_ESPERADO_PERC=replace(VAR_ATUAL_X_ESPERADO_PERC,is.infinite(VAR_ATUAL_X_ESPERADO_PERC),0)) %>%
  
  mutate(VAR_ATUAL_X_ESPERADO_VAL=CURRENTMONTH-ESPERADO) %>%
  
  mutate(VAR_ATUAL_X_ESPERADO_VAL=replace(VAR_ATUAL_X_ESPERADO_VAL,is.infinite(VAR_ATUAL_X_ESPERADO_VAL),0)) %>% 
  
  mutate_if(is.numeric, ~ round(.,2)) 

LASTMONTH <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))

sales_sm_1_2 <- sales_sm_1_1 %>% arrange(desc(.$YTD24)) %>% as.data.frame() %>% 
  rename_at(2:3,~ c(LASTMONTH,CURRENTMONTH)) 

corder1 <- c("CLIENTE",LASTMONTH,CURRENTMONTH,"ESPERADO","VAR_ATUAL_X_ESPERADO_PERC","VAR_ATUAL_X_ESPERADO_VAL","MEDIA_2024","MEDIA_DIAS","YTD23","YTD24")

sales_sm_1_3 <- 
  sales_sm_1_2 %>% .[,corder1]




##  SUMMARY 2 ==========================================================================

sales_sm_2_1 <-
  sales_esperado %>% group_by(CLIENTE,SETOR) %>% 
  summarize(
    LASTMONTH=sum(VRVENDA[floor_date(PEDDTEMIS,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTEMIS,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    CURRENTMONTH=sum(VRVENDA[floor_date(PEDDTEMIS,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTEMIS,"day")<=floor_date(Sys.Date(),'day') %m-% days(1)],na.rm = TRUE),
    
    MEDIA_2024=sum(VRVENDA[floor_date(PEDDTEMIS,"day")>= floor_date(Sys.Date(), "year") & floor_date(PEDDTEMIS,"day") < floor_date(Sys.Date(), "month")]/as.numeric(length(seq(floor_date(Sys.Date(),"year"),floor_date(Sys.Date(),"month"),by="month"))-1)),
    
    MEDIA_DIAS=sum(VRVENDA[floor_date(PEDDTEMIS,"day")>= floor_date(Sys.Date(), "year") & floor_date(PEDDTEMIS,"day") < floor_date(Sys.Date(), "month")]/as.numeric(length(seq(floor_date(Sys.Date(),"year"),floor_date(Sys.Date(),"month"),by="month"))-1))/num_days_working_days_month,
    
    ESPERADO=sum(VRVENDA[floor_date(PEDDTEMIS,"day")>= floor_date(Sys.Date(), "year") & floor_date(PEDDTEMIS,"day") < floor_date(Sys.Date(), "month")]/as.numeric(length(seq(floor_date(Sys.Date(),"year"),floor_date(Sys.Date(),"month"),by="month"))-1))/num_days_working_days_month*num_weekedays_last_day,
    
    YTD23=sum(VRVENDA[floor_date(PEDDTEMIS,"day") >= floor_date(Sys.Date() %m-% years(1), "year") & floor_date(PEDDTEMIS,"day") <= floor_date(Sys.Date() %m-% years(1), "day")-1],na.rm = TRUE),
    
    YTD24=sum(VRVENDA[floor_date(PEDDTEMIS,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTEMIS,"day") <= floor_date(Sys.Date(), "day")-1],na.rm = TRUE)) %>% as.data.frame() %>% 
  
  mutate(VAR_ATUAL_X_ESPERADO_PERC=CURRENTMONTH/ESPERADO-1) %>%
  
  mutate(VAR_ATUAL_X_ESPERADO_PERC=replace(VAR_ATUAL_X_ESPERADO_PERC,is.infinite(VAR_ATUAL_X_ESPERADO_PERC),0)) %>%
  
  mutate(VAR_ATUAL_X_ESPERADO_VAL=CURRENTMONTH-ESPERADO) %>%
  
  mutate(VAR_ATUAL_X_ESPERADO_VAL=replace(VAR_ATUAL_X_ESPERADO_VAL,is.infinite(VAR_ATUAL_X_ESPERADO_VAL),0)) %>% 
  
  mutate_if(is.numeric, ~ round(.,2)) 

LASTMONTH <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))

sales_sm_2_2 <- sales_sm_2_1 %>% arrange(desc(.$YTD24)) %>% as.data.frame() %>% 
  rename_at(3:4,~ c(LASTMONTH,CURRENTMONTH)) 

corder2 <- c("CLIENTE","SETOR",LASTMONTH,CURRENTMONTH,"ESPERADO","VAR_ATUAL_X_ESPERADO_PERC","VAR_ATUAL_X_ESPERADO_VAL","MEDIA_2024","MEDIA_DIAS","YTD23","YTD24")

sales_sm_2_3 <- 
  sales_sm_2_2 %>% .[,corder2]

## EXCEL ===========================================================================

wbk <- createWorkbook()


## sheet GERAL

addWorksheet(wbk, "GERAL")

writeData(wbk, sheet = "GERAL", x = sales_sm_1_3)

writeDataTable(wbk, "GERAL", sales_sm_1_3, startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

setColWidths(wbk, sheet = "GERAL", cols = 1, widths = 50)

setColWidths(wbk, sheet = "GERAL", cols = 5, widths = 30)

setColWidths(wbk, sheet = "GERAL", cols = 6, widths = 27)

setColWidths(wbk, sheet = "GERAL", cols = c(2,3,4,7,8,9,10), widths = 12)



## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wbk, sheet = "GERAL", style = estiloNumerico, cols = c(2,3,4,6,7,8,9,10), rows = 1:3000, gridExpand = TRUE)

addStyle(wbk, sheet = "GERAL", style = estiloPorcentagem, cols = 5, rows = 1:3000, gridExpand = TRUE)


## sheet SETORES

addWorksheet(wbk, "SETORES")

writeData(wbk, sheet = "SETORES", x = sales_sm_2_3)

writeDataTable(wbk, "SETORES", sales_sm_2_3, startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

setColWidths(wbk, sheet = "SETORES", cols = 1, widths = 55)

setColWidths(wbk, sheet = "SETORES", cols = 2, widths = 28)

setColWidths(wbk, sheet = "SETORES", cols = 6, widths =30)

setColWidths(wbk, sheet = "SETORES", cols = 7, widths =27)

setColWidths(wbk, sheet = "SETORES", cols = c(3,4,5,8,9), widths = 12)

addStyle(wbk, sheet = "SETORES", style = estiloNumerico, cols = c(3,4,5,7,8,9,10,11), rows = 1:3000, gridExpand = TRUE)

addStyle(wbk, sheet = "SETORES", style = estiloPorcentagem, cols = 6, rows = 1:3000, gridExpand = TRUE)

## cond format class

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

conditionalFormatting(wbk, sheet = "GERAL", cols = 5, rows = 2:1000, rule = 'E2>0', style = crescimento)

conditionalFormatting(wbk, sheet = "GERAL", cols = 5, rows = 2:1000, rule = 'E2<0', style = queda)

conditionalFormatting(wbk, sheet = "GERAL", cols = 6, rows = 2:1000, rule = 'F2>0', style = crescimento)

conditionalFormatting(wbk, sheet = "GERAL", cols = 6, rows = 2:1000, rule = 'F2<0', style = queda)


conditionalFormatting(wbk, sheet = "SETORES", cols = 6, rows = 2:1000, rule = 'F2>0', style = crescimento)

conditionalFormatting(wbk, sheet = "SETORES", cols = 6, rows = 2:1000, rule = 'F2<0', style = queda)

conditionalFormatting(wbk, sheet = "SETORES", cols = 7, rows = 2:1000, rule = 'G2>0', style = crescimento)

conditionalFormatting(wbk, sheet = "SETORES", cols = 7, rows = 2:1000, rule = 'G2<0', style = queda)



saveWorkbook(wbk, file = "RANKING_RESULT_ESPERADO.xlsx", overwrite = TRUE)


