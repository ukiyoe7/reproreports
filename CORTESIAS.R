## CORTESIAS
## LAST UPDATE 14.06.2024
## SANDRO JAKOSKA

## LOAD =======================================================================

library(tidyverse)
library(lubridate)
library(reshape2)
library(DBI)
library(openxlsx)

con2 <- dbConnect(odbc::odbc(), "repro",encoding="Latin1")


cortesias_sql <- 
dbGetQuery(con2,"WITH CLI AS (SELECT DISTINCT C.CLICODIGO,
                           CLINOMEFANT,
                            ENDCODIGO,
                             GCLCODIGO,
                             SETOR
                              FROM CLIEN C
                               INNER JOIN (SELECT CLICODIGO,E.ZOCODIGO,ZODESCRICAO SETOR,ENDCODIGO FROM ENDCLI E
                                INNER JOIN (SELECT ZOCODIGO,ZODESCRICAO FROM ZONA 
                                 WHERE ZOCODIGO IN (20,21,22,23,24,25,26,27,28))Z ON 
                                  E.ZOCODIGO=Z.ZOCODIGO WHERE ENDFAT='S')A ON C.CLICODIGO=A.CLICODIGO
                                   WHERE CLICLIENTE='S'  ),
                                 
                                 
     FIS AS (SELECT FISCODIGO FROM TBFIS WHERE FISCODIGO IN ('5.910','6.910','5.91O','6.91O')),
      
      PED AS (SELECT ID_PEDIDO,
                      EMPCODIGO,
                      TPCODIGO,
                       PEDDTEMIS,
                        PEDDTBAIXA,
                        P.CLICODIGO,
                         GCLCODIGO,
                          SETOR,
                           
                            CLINOMEFANT,
                             PEDORIGEM
                              FROM PEDID P
                               
                                INNER JOIN CLI C ON P.CLICODIGO=C.CLICODIGO AND P.ENDCODIGO=C.ENDCODIGO
                                 WHERE (PEDDTBAIXA BETWEEN '01.01.2024' AND 'YESTERDAY' OR 
                                 PEDDTBAIXA BETWEEN '01.01.2023' AND DATEADD(YEAR, -1, CURRENT_DATE))
                                 AND PEDSITPED<>'C'),
                                 
        PROD AS (SELECT PROCODIGO,IIF(PROCODIGO2 IS NULL,PROCODIGO,PROCODIGO2)CHAVE,PROTIPO,GR2CODIGO 
        FROM PRODU )
  
      
        SELECT PD.ID_PEDIDO,
                PEDDTBAIXA,
                  CLICODIGO,
                   CLINOMEFANT,
                    GCLCODIGO,
                     SETOR,
                      PD.FISCODIGO,
                        CHAVE,
                         (SELECT PRODESCRICAO FROM PRODU WHERE PROCODIGO=CHAVE) DESCRICAO,
                          SUM(PDPQTDADE)QTD,
                           SUM(PDPUNITLIQUIDO*PDPQTDADE)VRVENDA
                            FROM PDPRD PD
                             INNER JOIN PED P ON PD.ID_PEDIDO=P.ID_PEDIDO
                              INNER JOIN FIS F ON PD.FISCODIGO=F.FISCODIGO
                               LEFT JOIN PROD PR ON PD.PROCODIGO=PR.PROCODIGO
                                GROUP BY 1,2,3,4,5,6,7,8,9") 


View(cortesias_sql)

## CURRENT YEAR

cortesias<- 
cortesias_sql %>% 
  filter(floor_date(PEDDTBAIXA, "year")==floor_date(Sys.Date(), "year")) %>% 
  group_by(MES = toupper(format(floor_date(PEDDTBAIXA, "month"), "%b")),SETOR) %>% 
  mutate(SETOR=str_extract(SETOR, "SETOR \\d+")) %>% 
  summarize(v=sum(VRVENDA)) %>% 
  mutate(MES=factor(MES, levels = c("JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"), ordered = TRUE)) %>% 
  
  
  dcast(SETOR ~ MES) 

cortesias_with_totals <- addmargins(as.matrix(cortesias[-1]), 1:2)
cortesias_with_totals <- as.data.frame(cortesias_with_totals)

# Adding the SETOR column back and renaming the last column
cortesias_with_totals$SETOR <- c(as.character(cortesias$SETOR), "TOTAL")
names(cortesias_with_totals)[ncol(cortesias_with_totals) - 1] <- "TOTAL" # ncol(cortesias_with_totals) - 1 refers to the last column before adding TOTAL

# Reorder columns to ensure SETOR is first
cortesias_with_totals <- cortesias_with_totals %>% select(SETOR, everything())

# Print the result
View(cortesias_with_totals)


## LAST YEAR

cortesias_ly<- 
  cortesias_sql %>% 
  filter(floor_date(PEDDTBAIXA, "year")==floor_date(Sys.Date()-years(1), "year")) %>% 
  group_by(MES = toupper(format(floor_date(PEDDTBAIXA, "month"), "%b")),SETOR) %>% 
  mutate(SETOR=str_extract(SETOR, "SETOR \\d+")) %>% 
  summarize(v=sum(VRVENDA)) %>% 
  mutate(MES=factor(MES, levels = c("JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"), ordered = TRUE)) %>% 
  
  
  dcast(SETOR ~ MES) 

cortesias_with_totals_ly <- addmargins(as.matrix(cortesias_ly[-1]), 1:2)
cortesias_with_totals_ly <- as.data.frame(cortesias_with_totals_ly)

# Adding the SETOR column back and renaming the last column
cortesias_with_totals_ly$SETOR <- c(as.character(cortesias_ly$SETOR), "TOTAL")
names(cortesias_with_totals_ly)[ncol(cortesias_with_totals_ly) - 1] <- "TOTAL" # ncol(cortesias_with_totals) - 1 refers to the last column before adding TOTAL

# Reorder columns to ensure SETOR is first
cortesias_with_totals_ly <- cortesias_with_totals_ly %>% select(SETOR, everything())

# Print the result
View(cortesias_with_totals_ly)


titulo1 <- "RESUMO CORTESIAS - CFOPS 5.910,6.910,5.91O,6.91O"

titulo2 <- "2024"

titulo3 <- "2023"

## WRITE EXCEL  ==========================================================================


wb_cortesias <- createWorkbook()


## sheet

addWorksheet(wb_cortesias, "RESUMO")



## format cols

setColWidths(wb_cortesias, sheet = "RESUMO", cols = 1, widths = 2)



crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

size <- createStyle(fontSize = 17)

size2 <- createStyle(fontSize = 17)

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")
  
addStyle(wb_cortesias, sheet = "RESUMO", style = estiloNumerico, cols = c(2:9), rows = 4:27, gridExpand = TRUE)

addStyle(wb_cortesias, sheet = "RESUMO", style = size, cols = 2, rows = 2, gridExpand = TRUE)

addStyle(wb_cortesias, sheet = "RESUMO", style = size2, cols = 2, rows = 4, gridExpand = TRUE)

addStyle(wb_cortesias, sheet = "RESUMO", style = size2, cols = 2, rows = 17, gridExpand = TRUE)



writeData(wb_cortesias, "RESUMO", titulo1, startCol = 2, startRow = 2)

writeData(wb_cortesias, "RESUMO", titulo2, startCol = 2, startRow = 4)

writeData(wb_cortesias, "RESUMO", titulo3, startCol = 2, startRow = 17)


writeDataTable(wb_cortesias, "RESUMO", cortesias_with_totals, startCol = 2, startRow = 5, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

writeDataTable(wb_cortesias, "RESUMO", cortesias_with_totals_ly, startCol = 2, startRow = 18, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## WRITE WORKSHEET

saveWorkbook(wb_cortesias, file = "CORTESIAS.xlsx", overwrite = TRUE)




