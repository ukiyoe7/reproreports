## RANKING CLIENTES ATENDIMENTO
## CRIADO POR SANDRO JAKOSKA

## LOAD LIBS =======================================================================

library(tidyverse)
library(lubridate)
library(reshape2)
library(DBI)
library(openxlsx)
library(readxl)

## sql connect

con2 <- dbConnect(odbc::odbc(), "repro",encoding="Latin1")


## SQLS ==========================================================================

clientes <- dbGetQuery(con2,"
SELECT DISTINCT C.CLICODIGO,
CLINOMEFANT NOMEFANTASIA, 
C.GCLCODIGO CODGRUPO,
IIF(C.GCLCODIGO IS NULL,C.CLICODIGO || ' ' || CLINOMEFANT,'G' || C.GCLCODIGO || ' ' || GCLNOME) CLIENTE,
GCLNOME GRUPO,
ZODESCRICAO SETOR,
CIDNOME CIDADE,
CLIPCDESCPRODU DESCTGERAL,
CLIDTCAD DATACADASTRO
FROM CLIEN C
LEFT JOIN ENDCLI ON C.CLICODIGO=ENDCLI.CLICODIGO
LEFT JOIN CIDADE ON ENDCLI.CIDCODIGO=CIDADE.CIDCODIGO
LEFT JOIN ZONA ON ENDCLI.ZOCODIGO=ZONA.ZOCODIGO 
LEFT JOIN GRUPOCLI ON C.GCLCODIGO=GRUPOCLI.GCLCODIGO
WHERE CLICLIENTE='S' AND CLIFORNEC<>'S' AND ENDFAT='S'
AND ENDCLI.ZOCODIGO IN (20,21,22,23,24,25,26,27,28)
")

inativos <- dbGetQuery(con2,"SELECT DISTINCT SITCLI.CLICODIGO FROM SITCLI
INNER JOIN (SELECT DISTINCT SITCLI.CLICODIGO,MAX(SITDATA)ULTIMA FROM SITCLI
GROUP BY 1)A ON SITCLI.CLICODIGO=A.CLICODIGO AND A.ULTIMA=SITCLI.SITDATA 
INNER JOIN (SELECT DISTINCT SITCLI.CLICODIGO,SITDATA,MAX(SITSEQ)USEQ FROM SITCLI
GROUP BY 1,2)MSEQ ON A.CLICODIGO=MSEQ.CLICODIGO AND MSEQ.SITDATA=A.ULTIMA AND MSEQ.USEQ=SITCLI.SITSEQ
WHERE SITCODIGO=4")

atendimento <- dbGetQuery(con2,"SELECT CLICODIGO,FUNNOME ATENDIMENTO FROM CLIEN C
                    LEFT JOIN FUNCIO F ON C.FUNCODIGO2=F.FUNCODIGO
                     WHERE CLICLIENTE='S'")

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

## data stored 

sales2023 <- get(load("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\BASES\\sales2023.RData"))

## union all sales data

sales <- union_all(sales2023,sales2024)


desct <- dbGetQuery(con2,"
SELECT CLICODIGO,
 TBPCODIGO,
  TBPDESC2,
   CASE 
    WHEN TBPCODIGO=100 THEN 'VARILUX XR'
    WHEN TBPCODIGO=101 THEN 'VARILUX TRAD' 
    WHEN TBPCODIGO=102 THEN 'VARILUX DIGI'
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
    WHEN TBPCODIGO=307 THEN 'MARCAS PROPRIAS'
    WHEN TBPCODIGO=308 THEN 'INSIGNE'
    WHEN TBPCODIGO=309 THEN 'EYEZEN'
    WHEN TBPCODIGO=3660 THEN 'PROSING'
    WHEN TBPCODIGO=3661 THEN 'PROSING POLY'
   ELSE '' END LINHA
  FROM CLITBP
  WHERE TBPCODIGO IN (100,101,102,103,104,105,201,202,301,302,303,304,305,307,308,309,3660,3661)
  ") %>% mutate(LINHA=str_trim(LINHA))

mean_no_zero <- function(x) {
  x <- x[x != 0] 
  if (length(x) == 0) return(as.double(NA))  
  mean(x, na.rm = TRUE)  
}

mdesc <- desct %>% dcast(.,CLICODIGO ~ LINHA,value.var = "TBPDESC2",fun.aggregate = sum) %>% as.data.frame() 

gdesct <- dbGetQuery(con2,"
SELECT C.CLICODIGO,
CODGRUPO,
 TBPCODIGO,
  TBPDESC2,
   CASE 
   WHEN TBPCODIGO=100 THEN 'VARILUX XR'
    WHEN TBPCODIGO=101 THEN 'VARILUX TRAD' 
    WHEN TBPCODIGO=102 THEN 'VARILUX DIGI'
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
    WHEN TBPCODIGO=307 THEN 'MARCAS PROPRIAS'
    WHEN TBPCODIGO=308 THEN 'INSIGNE'
    WHEN TBPCODIGO=309 THEN 'EYEZEN'
    WHEN TBPCODIGO=3660 THEN 'PROSING'
    WHEN TBPCODIGO=3661 THEN 'PROSING POLY'
   ELSE '' END LINHA
  FROM CLITBP C
  LEFT JOIN (SELECT CLICODIGO,GCLCODIGO CODGRUPO
  FROM CLIEN WHERE CLICLIENTE='S')A 
  ON C.CLICODIGO=A.CLICODIGO
  WHERE TBPCODIGO IN (100,101,102,103,104,105,201,202,301,302,303,304,305,307,308,309,3660,3661) 
  ") %>%  mutate(LINHA=str_trim(LINHA)) %>% 
  
  filter(CODGRUPO!='') %>% group_by(CODGRUPO,LINHA) %>% summarise(m=mean(TBPDESC2)) %>% 
  dcast(.,CODGRUPO ~ LINHA,mean_no_zero) %>% replace(is.na(.),0) %>% apply(.,2,function(x) round(x,0)) %>% as.data.frame()


##  SUMMARIZE ==========================================================================

nsales <- sales %>% group_by(CLICODIGO) %>% 
  summarize(
    
    LASTMONTHTHISYEAR=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    THISMONTHLASTYEAR=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date((Sys.Date()-years(1)), 'month') & floor_date(PEDDTBAIXA,"day")<=floor_date(Sys.Date()-years(1), "day")-1],na.rm = TRUE),
    
    CURRENTMONTH=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date(),'month') %m-% days(1)],na.rm = TRUE),
    
    YTD23=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date() - years(1),"day")-1],na.rm = TRUE),
    
    YTD24=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1],na.rm = TRUE),
    
    MEDIA23=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= 
                          floor_date(Sys.Date()-years(1), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date()-years(1),"year"),
                                        floor_date(Sys.Date()-years(1),"month"),by="month")))),
    
    MEDIA24=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date(), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date(),"year"),
                                        floor_date(Sys.Date(),"day"),by="month")))),
    
    PAST12=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") >  ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)])
  ) %>% mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) %>% 
  
  mutate(VAR24_VLR=ifelse(is.finite(YTD24-YTD23),YTD24-YTD23,0))

nsales <- apply(nsales,2,function(x) round(x,2)) %>% as.data.frame()

data <- left_join(anti_join(clientes,inativos,by="CLICODIGO"),nsales,by="CLICODIGO") %>% 
  left_join(.,mdesc,by="CLICODIGO") %>% as.data.frame()

data <- data %>%  mutate(STATUS=case_when(
  DATACADASTRO>=floor_date(floor_date(Sys.Date() %m-% months(1), 'month')-years(1), "month") ~ 'CLIENTE NOVO',
  
  VAR24>=0 ~ 'CRESCIMENTO',
  VAR24<0 ~ 'QUEDA',
  PAST12==0  ~ 'PERDIDO',
  YTD23==0 & YTD24>0 ~ 'RECUPERADO',
  SEMRECEITA==0 ~ 'SEM RECEITA',
  TRUE ~ ''
))

LASTMONTHTHISYEAR <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))

THISMONTHLASTYEAR <- toupper(format(floor_date(Sys.Date()-years(1), "month"),"%b%/%Y"))

data <- data %>% arrange(desc(.$YTD24)) %>% as.data.frame() %>% 
  rename_at(10:12,~ c(LASTMONTHTHISYEAR,THISMONTHLASTYEAR,CURRENTMONTH)) %>% left_join(.,atendimento,by="CLICODIGO")

corder <- c("ATENDIMENTO","CLICODIGO","NOMEFANTASIA","CODGRUPO","GRUPO","SETOR","CIDADE","VAR24","STATUS","MEDIA23","MEDIA24","VARILUX DIGI","VARILUX TRAD","VARILUX XR","KODAK DIGITAL","KODAK TRAD","ACTUALITE","AVANCE","UZ+","IMAGEM DIGITAL","IMAGEM TRAD","INSIGNE","MARCAS PROPRIAS","EYEZEN","LA CRIZAL","LA KODAK","PROSING","PROSING POLY","VS ESSILOR","DESCTGERAL")

data <- data %>% .[,corder] %>% mutate_all(~replace(., is.nan(.), NA)) 


## get names to run loop

NOMES_ATENDIMENTOS <- read_excel("C:\\Users\\REPRO SANDRO\\OneDrive - Luxottica Group S.p.A (1)\\ATENDIMENTO\\NOMES_ATENDIMENTO.xlsx") %>% 
   data.frame(.) %>% rename(ATENDIMENTO=NOMES)


## EXCEL LOOP =====================================================

# Function to create individual workbooks
create_client_workbook <- function(client_name) {
  # Filter data for the specific client
  client_data <- data[data$ATENDIMENTO == client_name, ]
  
  # Create a new workbook
  wb <- createWorkbook()
  
  # Add worksheet
  addWorksheet(wb, "RANKING")
  
  # Write data to worksheet
  writeData(wb, sheet = "RANKING", x = client_data)
  
  # Set column widths
  setColWidths(wb, sheet = "RANKING", cols = 1, widths = 40)
  
  setColWidths(wb, sheet = "RANKING", cols = 2, widths = 15)
  
  setColWidths(wb, sheet = "RANKING", cols = 3, widths = 40)
  
  setColWidths(wb, sheet = "RANKING", cols = 4, widths = 15)
  
  setColWidths(wb, sheet = "RANKING", cols = 5, widths = 30)
  
  setColWidths(wb, sheet = "RANKING", cols = 6, widths = 40)
  
  setColWidths(wb, sheet = "RANKING", cols = 7, widths = 30)
  
  setColWidths(wb, sheet = "RANKING", cols = 8, widths = 12)
  
  setColWidths(wb, sheet = "RANKING", cols = 9, widths = 20)
  
  setColWidths(wb, sheet = "RANKING", cols = 10:30, widths = 15)
  
  # Create styles
  crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")
  queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")
  queda2 <- createStyle(fontColour = "#FF0800")
  estiloNumerico <- createStyle(numFmt = "#,##0")
  estiloPorcentagem <- createStyle(numFmt = "0%")
  
  # Apply styles
  addStyle(wb, sheet = "RANKING", style = estiloNumerico, cols = c(10:30), rows = 1:3000, gridExpand = TRUE)
  addStyle(wb, sheet = "RANKING", style = estiloPorcentagem, cols = 8, rows = 1:3000, gridExpand = TRUE)
  
  # Conditional formatting
  conditionalFormatting(wb, sheet = "RANKING", cols = 9, rows = 1:1000, rule = 'I1 == "CRESCIMENTO"', style = crescimento)
  conditionalFormatting(wb, sheet = "RANKING", cols = 9, rows = 1:1000, rule = 'I1 == "QUEDA"', style = queda)
  conditionalFormatting(wb, sheet = "RANKING", cols = 8, rows = 1:1000, rule = 'H1 < 0', style = queda2)
  
  # Write data table with styles
  writeDataTable(wb, "RANKING", client_data, startCol = 1, startRow = 1, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", withFilter = TRUE)
  
  # Save workbook
  saveWorkbook(wb, file = paste0("C:\\Users\\REPRO SANDRO\\OneDrive - Luxottica Group S.p.A (1)\\ATENDIMENTO\\RANKING_CLIENTES_ATENDIMENTO_", gsub(" .*", "", client_name), ".xlsx"), overwrite = TRUE)
}

# Iterate over each client and create workbooks
lapply(NOMES_ATENDIMENTOS$ATENDIMENTO, create_client_workbook)


