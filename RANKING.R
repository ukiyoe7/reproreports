## RANKING CLIENTS
## LAST UPDATE 15.03.2024
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
SELECT DISTINCT CLIEN.CLICODIGO,
CLINOMEFANT NOMEFANTASIA, 
CLIEN.GCLCODIGO CODGRUPO,
GCLNOME GRUPO,
ZODESCRICAO SETOR,
CIDNOME CIDADE,
CLIPCDESCPRODU DESCTGERAL,
CLIDTCAD DATACADASTRO
FROM CLIEN
LEFT JOIN ENDCLI ON CLIEN.CLICODIGO=ENDCLI.CLICODIGO
LEFT JOIN CIDADE ON ENDCLI.CIDCODIGO=CIDADE.CIDCODIGO
LEFT JOIN ZONA ON ENDCLI.ZOCODIGO=ZONA.ZOCODIGO 
LEFT JOIN GRUPOCLI ON CLIEN.GCLCODIGO=GRUPOCLI.GCLCODIGO
WHERE CLICLIENTE='S' AND CLIFORNEC<>'S' AND ENDFAT='S'
AND ENDCLI.ZOCODIGO IN (20,21,22,23,24,25,26,27,28)
")

inativos <- dbGetQuery(con2,"SELECT DISTINCT SITCLI.CLICODIGO FROM SITCLI
INNER JOIN (SELECT DISTINCT SITCLI.CLICODIGO,MAX(SITDATA)ULTIMA FROM SITCLI
GROUP BY 1)A ON SITCLI.CLICODIGO=A.CLICODIGO AND A.ULTIMA=SITCLI.SITDATA 
INNER JOIN (SELECT DISTINCT SITCLI.CLICODIGO,SITDATA,MAX(SITSEQ)USEQ FROM SITCLI
GROUP BY 1,2)MSEQ ON A.CLICODIGO=MSEQ.CLICODIGO AND MSEQ.SITDATA=A.ULTIMA AND MSEQ.USEQ=SITCLI.SITSEQ
WHERE SITCODIGO=4")


##  SALES ==========================================================================

sales2024 <- dbGetQuery(con2,"
WITH CLI AS (SELECT DISTINCT CLIEN.CLICODIGO,CLINOMEFANT
  FROM CLIEN
   WHERE CLICLIENTE='S'),
  
FIS AS (SELECT FISCODIGO FROM TBFIS WHERE FISTPNATOP IN ('V','SR','R')),
  
PED AS (SELECT ID_PEDIDO,PEDDTBAIXA,P.CLICODIGO FROM PEDID P
   
    INNER JOIN CLI C ON P.CLICODIGO=C.CLICODIGO
     WHERE PEDSITPED<>'C' AND PEDDTBAIXA BETWEEN '01.01.2024' AND 'YESTERDAY'),
     
AUX AS (SELECT PROCODIGO,PROTIPO FROM PRODU)     

SELECT PEDDTBAIXA,
P.CLICODIGO,
CASE 
WHEN PROTIPO='E' THEN 1
WHEN PROTIPO='P' THEN 1
WHEN PROTIPO='F' THEN 1
ELSE 0 
END LENTES,
 SUM(PDPQTDADE) QTD,SUM(PDPUNITLIQUIDO*PDPQTDADE)VRVENDA  
  FROM PDPRD PD
   INNER JOIN PED P ON PD.ID_PEDIDO=P.ID_PEDIDO
    INNER JOIN FIS F ON PD.FISCODIGO=F.FISCODIGO
    INNER JOIN AUX ON PD.PROCODIGO= AUX.PROCODIGO
     GROUP BY 1,2,3") 

sales2023 <- get(load("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\BASES\\sales2023.RData"))

sales <- union_all(sales2023,sales2024)


##  DISCOUNTS ==========================================================================

desct <- dbGetQuery(con2,"
SELECT CLICODIGO,
 TBPCODIGO,
  TBPDESC2,
   CASE 
    WHEN TBPCODIGO=100 THEN 'VARILUX'
    WHEN TBPCODIGO=101 THEN 'VARILUX' 
    WHEN TBPCODIGO=102 THEN 'VARILUX'
    WHEN TBPCODIGO=103 THEN 'LA CRIZAL'
    WHEN TBPCODIGO=104 THEN 'LA KODAK' 
    WHEN TBPCODIGO=105 THEN 'VS ESSILOR'
    WHEN TBPCODIGO=201 THEN 'KODAK TRAD'
    WHEN TBPCODIGO=202 THEN 'KODAK DIGITAL'
    WHEN TBPCODIGO=301 THEN 'IMAGEM TRAD'
    WHEN TBPCODIGO=302 THEN 'IMAGEM DIGITAL'
    WHEN TBPCODIGO=303 THEN 'UZ+'
    WHEN TBPCODIGO=304 THEN 'ACTUALITE'
    WHEN TBPCODIGO=305 THEN 'AVANCE'
    WHEN TBPCODIGO=308 THEN 'INSIGNE'
    WHEN TBPCODIGO=309 THEN 'EYEZEN'
    WHEN TBPCODIGO=3660 THEN 'PROSING'
    WHEN TBPCODIGO=3661 THEN 'PROSING POLY'
   ELSE '' END LINHA
  FROM CLITBP
  WHERE TBPCODIGO IN (100,101,102,103,104,105,201,202,301,302,303,304,305,308,309,3660,3661)
  ") 

mean_no_zero <- function(x) {
  x <- x[x != 0] 
  if (length(x) == 0) return(as.double(NA))  
  mean(x, na.rm = TRUE)  
}

mdesc <- desct %>% dcast(.,CLICODIGO ~ LINHA,value.var = "TBPDESC2",fun.aggregate = mean_no_zero) %>% as.data.frame() 

gdesct <- dbGetQuery(con2,"
SELECT C.CLICODIGO,
CODGRUPO,
 TBPCODIGO,
  TBPDESC2,
   CASE 
   WHEN TBPCODIGO=100 THEN 'VARILUX'
    WHEN TBPCODIGO=101 THEN 'VARILUX' 
    WHEN TBPCODIGO=102 THEN 'VARILUX'
    WHEN TBPCODIGO=103 THEN 'LA CRIZAL'
    WHEN TBPCODIGO=104 THEN 'LA KODAK' 
    WHEN TBPCODIGO=105 THEN 'VS ESSILOR'
    WHEN TBPCODIGO=201 THEN 'KODAK TRAD'
    WHEN TBPCODIGO=202 THEN 'KODAK DIGITAL'
    WHEN TBPCODIGO=301 THEN 'IMAGEM TRAD'
    WHEN TBPCODIGO=302 THEN 'IMAGEM DIGITAL'
    WHEN TBPCODIGO=303 THEN 'UZ+'
    WHEN TBPCODIGO=304 THEN 'ACTUALITE'
    WHEN TBPCODIGO=305 THEN 'AVANCE'
    WHEN TBPCODIGO=308 THEN 'INSIGNE'
    WHEN TBPCODIGO=309 THEN 'EYEZEN'
    WHEN TBPCODIGO=3660 THEN 'PROSING'
    WHEN TBPCODIGO=3661 THEN 'PROSING POLY'
   ELSE '' END LINHA
  FROM CLITBP C
  LEFT JOIN (SELECT CLICODIGO,GCLCODIGO CODGRUPO
  FROM CLIEN WHERE CLICLIENTE='S')A 
  ON C.CLICODIGO=A.CLICODIGO
  WHERE TBPCODIGO IN (100,101,102,103,104,105,201,202,301,302,303,304,305,308,309,3660,3661)
  ") %>% filter(CODGRUPO!='') %>% group_by(CODGRUPO,LINHA) %>% summarise(m=mean(TBPDESC2)) %>% 
  dcast(.,CODGRUPO ~ LINHA,mean) %>% replace(is.na(.),0) %>% apply(.,2,function(x) round(x,0)) %>% as.data.frame()


##  RANKING CLIENTS ==========================================================================

nsales <- sales %>% group_by(CLICODIGO) %>% 
  summarize(
    LASTMONTHLASTYEAR=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date((Sys.Date()-years(1)) %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    LASTMONTHTHISYEAR=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    CURRENTMONTH=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date(),'month') %m-% days(1)],na.rm = TRUE),
    
    YTD23=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date() - years(1),"day")-1],na.rm = TRUE),
    
    YTD24=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1],na.rm = TRUE),
    
    MEDIA23=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= 
                          floor_date(Sys.Date()-years(1), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date()-years(1),"year"),
                                        floor_date(Sys.Date()-years(1),"month"),by="month")))),
    
    MEDIA24=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date(), "year") & 
                          floor_date(PEDDTBAIXA,"day") < floor_date(Sys.Date(), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date(),"year"),
                                        floor_date(Sys.Date(),"day"),by="month")))),
    
    PAST12=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") >  ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)])
  ) %>% mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0))

nsales <- apply(nsales,2,function(x) round(x,2)) %>% as.data.frame()

data <- left_join(anti_join(clientes,inativos,by="CLICODIGO"),nsales,by="CLICODIGO") %>% 
  left_join(.,mdesc,by="CLICODIGO") %>% as.data.frame()

data <- data %>%  mutate(STATUS=case_when(
  DATACADASTRO>=floor_date(floor_date(Sys.Date() %m-% months(1), 'month')-years(1), "month") ~ 'CLIENTE NOVO',
  SEMRECEITA==0 ~ 'SEM RECEITA',
  VAR24>=0 ~ 'CRESCIMENTO',
  VAR24<0 ~ 'QUEDA',
  PAST12==0  ~ 'PERDIDO',
  YTD23==0 & YTD24>0 ~ 'RECUPERADO',
  TRUE ~ ''
))

LASTMONTHLASTYEAR1 <- toupper(format(floor_date(Sys.Date()-years(1), "month")-1,"%b%/%Y"))

LASTMONTHTHISYEAR1 <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH1 <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))

data <- data %>% arrange(desc(.$YTD23)) %>% as.data.frame() %>% 
  rename_at(9:11,~ c(LASTMONTHLASTYEAR1,LASTMONTHTHISYEAR1,CURRENTMONTH1)) %>% .[,c(1:6,9:13,18,34,14:15,7,19:33)] %>% mutate_all(~replace(., is.nan(.), NA))



##  RANKING GROUPS ==========================================================================


dt <- left_join(anti_join(clientes,inativos,by="CLICODIGO"),sales,by="CLICODIGO") %>% as.data.frame()

gsales <- dt %>% group_by(SETOR,CODGRUPO,GRUPO) %>% 
  summarize(
    LASTMONTHLASTYEAR2=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date((Sys.Date()-years(1)) %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    LASTMONTHTHISYEAR2=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    CURRENTMONTH2=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date(),'month') %m-% days(1)],na.rm = TRUE),
    
    YTD23=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date() - years(1),"day")-1],na.rm = TRUE),
    
    YTD24=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1],na.rm = TRUE),
    
    MEDIA23=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= 
                          floor_date(Sys.Date()-years(1), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date()-years(1),"year"),
                                        floor_date(Sys.Date()-years(1),"month"),by="month")))),
    
    MEDIA24=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date(), "year") & 
                          floor_date(PEDDTBAIXA,"day") < floor_date(Sys.Date(), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date(),"year"),
                                        floor_date(Sys.Date(),"day"),by="month")))),
    
    PAST12=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") >  ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)])
  ) %>% mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) 

gsales <- gsales %>%  mutate(STATUS=case_when(
  SEMRECEITA==0 ~ 'SEM RECEITA',
  VAR24>0 ~ 'CRESCIMENTO',
  VAR24<0 ~ 'QUEDA',
  PAST12==0  ~ 'PERDIDO',
  YTD23==0 & YTD24>0 ~ 'RECUPERADO',
  TRUE ~ ''
))

gdata <-apply(gsales[,4:13],2,function(x) round(x,2)) %>% as.data.frame() %>% 
  cbind(gsales[,1:3],.) %>% cbind(.,gsales[,14])


gdata <- gdata %>% left_join(.,gdesct,by="CODGRUPO") %>% as.data.frame()

gdata <- gdata %>% filter(CODGRUPO!='')

LASTMONTHLASTYEAR1 <- toupper(format(floor_date(Sys.Date()-years(1), "month")-1,"%b%/%Y"))

LASTMONTHTHISYEAR1 <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH1 <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))


gdata <- gdata %>% arrange(desc(.$YTD24)) %>% as.data.frame() %>% 
  rename_at(4:6,~ c(LASTMONTHLASTYEAR1,LASTMONTHTHISYEAR1,CURRENTMONTH1)) %>% .[,c(2,3,1,4:8,13,14,9,10,15:29)] %>% mutate_all(~replace(., is.nan(.), NA))


## CREATE EXCEL ===============================


wb <- createWorkbook()


## sheet ranking

addWorksheet(wb, "RANKING")

writeData(wb, sheet = "RANKING", x = data)

setColWidths(wb, sheet = "RANKING", cols = 1, widths = 12)

setColWidths(wb, sheet = "RANKING", cols = 2, widths = 40)

setColWidths(wb, sheet = "RANKING", cols = 3, widths = 12)

setColWidths(wb, sheet = "RANKING", cols = 4, widths = 30)

setColWidths(wb, sheet = "RANKING", cols = 5, widths = 30)

setColWidths(wb, sheet = "RANKING", cols = 6, widths = 20)

setColWidths(wb, sheet = "RANKING", cols = 7:12, widths = 12)

setColWidths(wb, sheet = "RANKING", cols = 13, widths = 17)

setColWidths(wb, sheet = "RANKING", cols = 14:15, widths = 15)

setColWidths(wb, sheet = "RANKING", cols = 16:31, widths = 15)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb, sheet = "RANKING", style = estiloNumerico, cols = c(7:11,14:31), rows = 1:3000, gridExpand = TRUE)

addStyle(wb, sheet = "RANKING", style = estiloPorcentagem, cols = 12, rows = 1:3000, gridExpand = TRUE)

## cond format class

conditionalFormatting(wb, sheet = "RANKING", cols = 13, rows = 1:1000, rule = 'M1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb, sheet = "RANKING", cols = 13, rows = 1:1000, rule = 'M1 == "QUEDA"', style = queda)

writeDataTable(wb, "RANKING", data, startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## sheet ranking groups

addWorksheet(wb, "RANKING GRUPOS")

writeData(wb, sheet = "RANKING GRUPOS", x = gdata)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 1, widths = 12)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 2, widths = 40)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 3, widths = 40)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 4:9, widths = 10)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 10, widths = 17)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 11:12, widths = 10)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 13:27, widths = 15)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")


addStyle(wb, sheet = "RANKING GRUPOS", style = estiloNumerico, cols = c(4:8,11:27), rows = 1:3000, gridExpand = TRUE)

addStyle(wb, sheet = "RANKING GRUPOS", style = estiloPorcentagem, cols = 9, rows = 1:3000, gridExpand = TRUE)


conditionalFormatting(wb, sheet = "RANKING GRUPOS", cols = 10, rows = 1:1000, rule = 'J1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb, sheet = "RANKING GRUPOS", cols = 10, rows = 1:1000, rule = 'J1 == "QUEDA"', style = queda)

writeDataTable(wb, "RANKING GRUPOS", gdata, startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


saveWorkbook(wb, file = "RANKING_CLIENTES.xlsx", overwrite = TRUE)

