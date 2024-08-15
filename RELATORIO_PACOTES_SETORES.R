  ## RELATORIO PACOTES
  ## SANDRO JAKOSKA
  
  ## load =====================================
  
  library(DBI)
  library(tidyverse)
  library(openxlsx)
  library(readr)
  library(lubridate)
  library(reshape2)
  library(stringr)
  
  con2 <- dbConnect(odbc::odbc(), "repro",encoding="Latin1")
  
  ## sql ================================================
  
  base_pacotes <- dbGetQuery(con2,"
  WITH SETOR AS (SELECT ZOCODIGO,ZODESCRICAO FROM ZONA),
  
  ENDE AS (SELECT CLICODIGO, ZODESCRICAO FROM ENDCLI INNER JOIN SETOR ON ENDCLI.ZOCODIGO=SETOR.ZOCODIGO WHERE ENDFAT='S'),
  
  CLI AS (SELECT  CLIEN.CLICODIGO,CLIRAZSOCIAL,ZODESCRICAO FROM CLIEN LEFT JOIN ENDE ON CLIEN.CLICODIGO=ENDE.CLICODIGO ),
  
  PROD AS (SELECT PROCODIGO,PRODESCRICAO FROM PRODU),
  
  CLIPCT AS (SELECT PCTCLI.CLICODIGO,CLIRAZSOCIAL,ZODESCRICAO SETOR_COMERCIAL,PCTNUMERO,PCTDTCAD,PCTDTFIM,PCTDTSUSP,PCTOBSER,PCTSUSPOBS FROM PCTCLI
             LEFT JOIN CLI ON PCTCLI.CLICODIGO=CLI.CLICODIGO WHERE PCTSITUACAO<>'C')
  
  SELECT CLICODIGO,CLIRAZSOCIAL RAZAO_SOCIAL,SETOR_COMERCIAL, 
  PCTPRO.PCTNUMERO,PCTDTCAD PCT_DATA_CADASTRO,PCTDTFIM TERMINO,PCTPRO.PROCODIGO,
  PRODESCRICAO DESCRICAO,CAST (PCPPCOUNIT AS NUMERIC(15,5)) PRECO_UNIT,PCPQTDADE QTD,PCPSALDO SALDO_ATUAL,
  PCTDTSUSP DATA_SUSPENSAO,CAST(PCTOBSER AS VARCHAR(500)) OBS,PCTSUSPOBS
  FROM PCTPRO 
  INNER JOIN CLIPCT ON PCTPRO.PCTNUMERO=CLIPCT.PCTNUMERO
  LEFT JOIN PROD ON PCTPRO.PROCODIGO=PROD.PROCODIGO") %>% 
  
  mutate(`DIAS DE PACOTE ATIVO`= as.numeric((Sys.Date()-PCT_DATA_CADASTRO))) %>% 
    
    mutate(`MEDIA CONSUMO`= round(as.numeric((QTD-SALDO_ATUAL)/`DIAS DE PACOTE ATIVO`),4)) %>% 
    
    mutate(PREVISÃO_TERMINO = if_else(
      is_valid_date <- !is.na(Sys.Date() + (SALDO_ATUAL/`MEDIA CONSUMO`)),
      Sys.Date() + (SALDO_ATUAL/`MEDIA CONSUMO`),TERMINO)) %>% 
    
    mutate(PREVISÃO_TERMINO=if_else(`MEDIA CONSUMO`==0,NA,PREVISÃO_TERMINO)) %>% 
    
    mutate(STATUS = if_else(TERMINO < Sys.Date(), "VENCIDO",
                          if_else(TERMINO > Sys.Date(), "ATIVO","INDEFINIDO"))) %>% 
    
    mutate(STATUS = if_else(SALDO_ATUAL ==0, "ZERADO",STATUS)) %>% 
    mutate(STATUS = if_else(!is.na(DATA_SUSPENSAO), "SUSPENSO",STATUS)) %>% 
    
    mutate(PREVISÃO_TERMINO=if_else(STATUS=="VENCIDO" | STATUS=="ZERADO" | STATUS=="SUSPENSO",NA,PREVISÃO_TERMINO)) %>% 
    
    mutate(VALOR_TOTAL=PRECO_UNIT*QTD) %>% 
    
    arrange(desc(PCT_DATA_CADASTRO))
  
  
  dev_pacotes <- dbGetQuery(con2,"WITH CLI AS (SELECT DISTINCT C.CLICODIGO,
                                CLINOMEFANT,
                                 SETOR
                                  FROM CLIEN C
  LEFT JOIN (SELECT CLICODIGO,
                     E.ZOCODIGO,
                      ZODESCRICAO SETOR FROM ENDCLI E
  LEFT JOIN (SELECT ZOCODIGO,
                     ZODESCRICAO FROM 
                      ZONA WHERE ZOCODIGO IN (20,21,22,23,24,25,26,27,28)
  )Z ON E.ZOCODIGO=Z.ZOCODIGO WHERE ENDFAT='S')A ON C.CLICODIGO=A.CLICODIGO
   WHERE CLICLIENTE='S')
  
  SELECT P.CLICODIGO,
          SETOR,
           PEFDTEMIS EMISSAO,
            FISCODIGO,
             ID_PEDIDO,
              PEFVRTOTAL VRTOTAL FROM PEDFO P
               INNER JOIN CLI C ON P.CLICODIGO=C.CLICODIGO
                WHERE PEFSIT<>'C' AND 
                 PEFDTEMIS BETWEEN '01.01.2023' AND 
                  'TODAY' AND 
                    FISCODIGO IN ('1.94C','2.94C','1.92C','2.92C') AND PEFSIT<>'C'")
  
  
  consumo_pacotes <- dbGetQuery(con2,"  WITH CLI AS (SELECT DISTINCT C.CLICODIGO,
                             CLINOMEFANT,
                              ENDCODIGO,
                               GCLCODIGO,
                               SETOR
                                FROM CLIEN C
                                 INNER JOIN (SELECT CLICODIGO,E.ZOCODIGO,ZODESCRICAO SETOR,ENDCODIGO FROM ENDCLI E
                                  INNER JOIN (SELECT ZOCODIGO,ZODESCRICAO FROM ZONA 
                                  WHERE ZOCODIGO IN (20,21,22,23,24,25,26,27,28))Z ON 
                                   E.ZOCODIGO=Z.ZOCODIGO WHERE ENDFAT='S')A ON C.CLICODIGO=A.CLICODIGO
                                    WHERE CLICLIENTE='S'),
                                   
       FIS AS (SELECT FISCODIGO FROM TBFIS WHERE FISTPNATOP IN ('V','R','SR')),
        
        PED AS (SELECT ID_PEDIDO,
                         PEDDTEMIS,
                          P.CLICODIGO,
                           GCLCODIGO,
                            SETOR,
                             CLINOMEFANT
                              FROM PEDID P
                                  INNER JOIN CLI C ON P.CLICODIGO=C.CLICODIGO AND P.ENDCODIGO=C.ENDCODIGO
                                   WHERE 
                                   (PEDDTEMIS BETWEEN DATEADD(-60 DAY TO CURRENT_DATE) AND 'YESTERDAY')
                                   AND PEDSITPED<>'C' ),
                                   
          PROD AS (SELECT PROCODIGO,IIF(PROCODIGO2 IS NULL,PROCODIGO,PROCODIGO2)CHAVE,PROTIPO,GR2CODIGO 
          FROM PRODU )
        
          SELECT PD.ID_PEDIDO,
                   PEDDTEMIS DATA_EMISSAO,
                    CLICODIGO,
                     CLINOMEFANT NOME_FANTASIA,
                      GCLCODIGO GRUPO,
                       SETOR,
                         CHAVE COD_LENTE,
                          PDPDESCRICAO DESCRICAO,
                           PCTNUMERO,
                            SUM(PDPQTDADE)QTD,
                             SUM(PDPUNITLIQUIDO*PDPQTDADE)VRVENDA
                              FROM PDPRD PD
                               INNER JOIN PED P ON PD.ID_PEDIDO=P.ID_PEDIDO
                                INNER JOIN FIS F ON PD.FISCODIGO=F.FISCODIGO
                                 LEFT JOIN PROD PR ON PD.PROCODIGO=PR.PROCODIGO
                                  WHERE PCTNUMERO IS NOT NULL
                                  GROUP BY 1,2,3,4,5,6,7,8,9 ORDER BY ID_PEDIDO DESC")
  
  
  # summary current year ============================================
  
  # Create a sequence of months for the current year
  all_months <- toupper(format(seq(from = as.Date(floor_date(Sys.Date(), "year")), 
                                   to = as.Date(floor_date(Sys.Date(), "year") + years(1) - days(1)), 
                                   by = "month"), "%b"))
  
  # Process the data to ensure all months are included
  pacotes_current_year <-
    base_pacotes %>% 
    
  ## add one row  
    rbind(base_pacotes %>% filter(SETOR_COMERCIAL=='SETOR 9 - ZONA NEUTRA') %>%   
            filter(floor_date(PCT_DATA_CADASTRO, "year") == floor_date(Sys.Date()-years(1), "year")) %>% 
            .[1,] %>%   mutate(
              PCT_DATA_CADASTRO = floor_date(Sys.Date(), "year"),
              VALOR_TOTAL=0,
              across(-c(SETOR_COMERCIAL, PCT_DATA_CADASTRO,VALOR_TOTAL), ~ NA)))  %>% 
    
  
    filter(floor_date(PCT_DATA_CADASTRO, "year") == floor_date(Sys.Date(), "year")) %>% 
    mutate(MES = toupper(format(floor_date(PCT_DATA_CADASTRO, "month"), "%b")),
           SETOR_COMERCIAL = str_extract(SETOR_COMERCIAL, "SETOR \\d+")) %>% 
    group_by(MES, SETOR_COMERCIAL) %>% 
    summarize(v = sum(VALOR_TOTAL, na.rm = TRUE), .groups = 'drop') %>% 
    complete(MES = all_months, SETOR_COMERCIAL = unique(SETOR_COMERCIAL), fill = list(v = 0)) %>% 
    mutate(MES = factor(MES, levels = all_months, ordered = TRUE)) %>% 
    dcast(SETOR_COMERCIAL ~ MES, value.var = "v") %>% 
    replace_na(list(JAN = 0, FEV = 0, MAR = 0, ABR = 0, MAI = 0, JUN = 0, JUL = 0, AGO = 0, SET = 0, OUT = 0, NOV = 0, DEZ = 0))
  
  # Add totals row and column
  pacotes_current_year_with_totals <- addmargins(as.matrix(pacotes_current_year[-1]), 1:2)
  pacotes_current_year_with_totals <- as.data.frame(pacotes_current_year_with_totals)
  
  # Adding the SETOR column back and renaming the last column
  pacotes_current_year_with_totals$SETOR <- c(as.character(pacotes_current_year$SETOR_COMERCIAL), "TOTAL")
  names(pacotes_current_year_with_totals)[ncol(pacotes_current_year_with_totals) - 1] <- "TOTAL" # ncol(cortesias_with_totals) - 1 refers to the last column before adding TOTAL
  
  # Reorder columns to ensure SETOR is first
  pacotes_current_year_with_totals <- pacotes_current_year_with_totals %>% select(SETOR, everything())
  
  
  ## summary last year ========================================
  
  pacotes_last_year <-
    base_pacotes  %>% 
    filter(floor_date(PCT_DATA_CADASTRO, "year")==floor_date(Sys.Date()-years(1), "year")) %>% 
    mutate(MES = toupper(format(floor_date(PCT_DATA_CADASTRO, "month"), "%b")),
           SETOR_COMERCIAL = str_extract(SETOR_COMERCIAL, "SETOR \\d+")) %>% 
    group_by(MES, SETOR_COMERCIAL) %>% 
    summarize(v = sum(VALOR_TOTAL, na.rm = TRUE), .groups = 'drop') %>% 
    complete(MES = all_months, SETOR_COMERCIAL = unique(SETOR_COMERCIAL), fill = list(v = 0)) %>% 
    mutate(MES = factor(MES, levels = all_months, ordered = TRUE)) %>% 
    dcast(SETOR_COMERCIAL ~ MES, value.var = "v") %>% 
    replace_na(list(JAN = 0, FEV = 0, MAR = 0, ABR = 0, MAI = 0, JUN = 0, JUL = 0, AGO = 0, SET = 0, OUT = 0, NOV = 0, DEZ = 0))
  
  
  pacotes_last_year_with_totals <- addmargins(as.matrix(pacotes_last_year[-1]), 1:2)
  
  pacotes_last_year_with_totals <- as.data.frame(pacotes_last_year_with_totals)
  
  # Adding the SETOR column back and renaming the last column
  pacotes_last_year_with_totals$SETOR <- c(as.character(pacotes_last_year$SETOR), "TOTAL")
  names(pacotes_last_year_with_totals)[ncol(pacotes_last_year_with_totals) - 1] <- "TOTAL" # ncol(cortesias_with_totals) - 1 refers to the last column before adding TOTAL
  
  # Reorder columns to ensure SETOR is first
  pacotes_last_year_with_totals <- pacotes_last_year_with_totals %>% select(SETOR, everything())
  
  
  # Align dataframes by SETOR
  merged <- full_join(pacotes_current_year_with_totals, pacotes_last_year_with_totals, by = "SETOR", suffix = c("_current", "_last"))
  
  
  ## resumo setores ====================================
  
  pacotes_year <-
    rbind(
      pacotes_current_year %>% mutate(ANO=format(floor_date(Sys.Date(), "year"),"%Y")),
      pacotes_last_year %>% mutate(ANO=format(floor_date(Sys.Date()-years(1), "year"),"%Y"))) %>% .[,c(1,14,2:13)] %>% 
    mutate(TOTAL = rowSums(select(., 3:14)))
  
  
  ## matrix operation ===============================================
  
  varYY <- merged %>%
    mutate(across(all_of(c(paste0(all_months, "_current"), "TOTAL_current")),  # added "TOTAL_current"
                  ~ (.x / merged[[sub("_current", "_last", cur_column())]] - 1),
                  .names = "{sub('_current', '', .col)}"))
  
  # Select and reorder the necessary columns
  varYY <- varYY %>% select(SETOR, all_of(c("TOTAL", all_months))) %>%   # added "TOTAL"
    mutate(across(everything(), ~ ifelse(is.nan(.), NA, .))) %>% 
    mutate(ANO = paste0(format(Sys.Date() - years(1), "%Y"), " X ", format(Sys.Date(), "%Y"))) %>% 
    select(SETOR, ANO, all_of(c(all_months,"TOTAL")))# added "TOTAL"
  
  
  ## devolução pacotes ===============================================
  
    
    # Process the data to ensure all months are included
  dev_pacotes_current_year <-
    dev_pacotes %>% 
    
    filter(floor_date(EMISSAO, "year") == floor_date(Sys.Date(), "year")) %>% 
    mutate(MES = toupper(format(floor_date(EMISSAO, "month"), "%b")),
           SETOR = str_extract(SETOR, "SETOR \\d+")) %>% 
    group_by(MES, SETOR) %>% 
    summarize(v = sum(VRTOTAL, na.rm = TRUE), .groups = 'drop') %>% 
    complete(MES = all_months, SETOR = unique(SETOR), fill = list(v = 0)) %>% 
    mutate(MES = factor(MES, levels = all_months, ordered = TRUE)) %>% 
    dcast(SETOR ~ MES, value.var = "v") %>% 
    replace_na(list(JAN = 0, FEV = 0, MAR = 0, ABR = 0, MAI = 0, JUN = 0, JUL = 0, AGO = 0, SET = 0, OUT = 0, NOV = 0, DEZ = 0))
  
  # Add totals row and column
  dev_pacotes_current_year_with_totals <- addmargins(as.matrix(dev_pacotes_current_year[-1]), 1:2)
  dev_pacotes_current_year_with_totals <- as.data.frame(dev_pacotes_current_year_with_totals)
  
  # Adding the SETOR column back and renaming the last column
  dev_pacotes_current_year_with_totals$SETOR <- c(as.character(dev_pacotes_current_year$SETOR), "TOTAL")
  names(dev_pacotes_current_year_with_totals)[ncol(dev_pacotes_current_year_with_totals) - 1] <- "TOTAL" # ncol(cortesias_with_totals) - 1 refers to the last column before adding TOTAL
  
  # Reorder columns to ensure SETOR is first
  dev_pacotes_current_year_with_totals <- dev_pacotes_current_year_with_totals %>% mutate(ANO = paste0(format(Sys.Date(), "%Y"))) %>% 
    select(SETOR, ANO, all_of(c(all_months,"TOTAL")))# added "TOTAL"
  
  
  ## METAS ================================================
  
  pacotes_last_year_with_totals_metas <-
  pacotes_last_year_with_totals[,-1] %>% as.matrix() *1.1 
  
  pacotes_last_year_with_totals_metas <- as.data.frame(pacotes_last_year_with_totals_metas)
  
  # Adding the SETOR column back and renaming the last column
  pacotes_last_year_with_totals_metas $SETOR <- as.character(pacotes_last_year_with_totals$SETOR)
  
  # Reorder columns to ensure SETOR is first
  pacotes_last_year_with_totals_metas <- pacotes_last_year_with_totals_metas %>% mutate(ANO = paste0(format(Sys.Date(), "%Y"))) %>%
    select(SETOR,ANO, everything())
  
  # dif metas
  
  pacotes_last_year_with_totals_metas_dif <-
   ((pacotes_current_year_with_totals[,-1])/(pacotes_last_year_with_totals[,-1] %>% as.matrix() *1.1))-1 
  
  pacotes_last_year_with_totals_metas_dif <- as.data.frame(pacotes_last_year_with_totals_metas_dif)
  
  pacotes_last_year_with_totals_metas_dif$SETOR <- as.character( pacotes_current_year_with_totals$SETOR)
  
  pacotes_last_year_with_totals_metas_dif <- pacotes_last_year_with_totals_metas_dif %>% select(SETOR, everything())
  
  numeric_cols <- sapply(pacotes_last_year_with_totals_metas_dif, is.numeric)
  
  
  pacotes_last_year_with_totals_metas_dif[, numeric_cols] <- lapply(
    pacotes_last_year_with_totals_metas_dif[, numeric_cols], 
    function(x) {
      x[is.nan(x)] <- 0      # Replace NaN with 0
      x[is.infinite(x)] <- NA  # Replace Inf and -Inf with NA (or another appropriate value)
      return(x)
    }
  )
  
  pacotes_last_year_with_totals_metas_dif[is.na(pacotes_last_year_with_totals_metas_dif)] <- 0
  
  pacotes_last_year_with_totals_metas_dif <-
  pacotes_last_year_with_totals_metas_dif %>%  mutate(ANO = paste0(format(Sys.Date(), "%Y"))) %>%
    select(SETOR,ANO, everything())
  
  
  ## create excel setores ==================================================
  
  
  unique_setores <- unique(base_pacotes$SETOR_COMERCIAL)
  
  # Function to create individual workbooks
     create_setores_workbook <- function(setor) {
    
    # Filter data for the specific client
    setores_data <- base_pacotes[base_pacotes$SETOR_COMERCIAL == setor, ]
    
    setores_data2 <- pacotes_year[pacotes_year$SETOR == str_extract(setor, "SETOR \\d+"), ]
    
    setores_data3 <- varYY[varYY$SETOR == str_extract(setor, "SETOR \\d+"), ]
    
    setores_data4 <- dev_pacotes_current_year_with_totals[dev_pacotes_current_year_with_totals$SETOR == str_extract(setor, "SETOR \\d+"), ]
    
    setores_data5 <- pacotes_last_year_with_totals_metas[pacotes_last_year_with_totals_metas$SETOR == str_extract(setor, "SETOR \\d+"), ]
  
    setores_data6 <- pacotes_last_year_with_totals_metas_dif[pacotes_last_year_with_totals_metas_dif$SETOR == str_extract(setor, "SETOR \\d+"), ]
    
    setores_data7 <- consumo_pacotes[consumo_pacotes$SETOR == setor, ]
      
    pacotes <- createWorkbook()
    
    addWorksheet(pacotes, "RESUMO")
    addWorksheet(pacotes, "LANCAMENTO PACOTES")
    addWorksheet(pacotes, "CONSUMO PACOTES")
    
    titulo1 <- "RESUMO LANÇAMENTO PACOTES 2023/2024 "
    titulo2 <- "DEVOLUCOES PACOTES"
    titulo3 <- "METAS 10%"
    titulo4 <- "DIF METAS x LANÇAMENTOS 2024"
    
    # Set column widths
    setColWidths(pacotes, sheet = "RESUMO", cols = 1, widths = 1)
    
    setColWidths(pacotes, sheet = "RESUMO", cols = 2, widths =20)
    
    setColWidths(pacotes, sheet = "RESUMO", cols = 3, widths = 10)
    
    setColWidths(pacotes, sheet = "RESUMO", cols = c(4:16), widths = 9)
    
  ## styles
    numstyle <- createStyle(numFmt = "#,##0")
    
    percstyle <- createStyle(numFmt = "0.00%")
    
    datesty <- createStyle(numFmt = "dd/MM/yyyy")
    
    crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")
    
    queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")
    
    
    writeData(pacotes, "RESUMO", titulo1, startCol = 2, startRow = 2)
    writeData(pacotes, "RESUMO", titulo2, startCol = 2, startRow = 14)
    writeData(pacotes, "RESUMO", titulo3, startCol = 2, startRow = 19)
    writeData(pacotes, "RESUMO", titulo4, startCol = 2, startRow = 24)
    
    writeDataTable(pacotes, "RESUMO", setores_data2 %>% mutate(ANO=as.numeric(ANO)), startCol = 2, startRow = 5, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
      
    writeDataTable(pacotes, "RESUMO", setores_data3, startCol = 2, startRow = 10, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
  
    writeDataTable(pacotes, "RESUMO", setores_data4 %>% mutate(ANO=as.numeric(ANO)), startCol = 2, startRow = 15, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
    
    writeDataTable(pacotes, "RESUMO", setores_data5 %>% mutate(ANO=as.numeric(ANO)), startCol = 2, startRow = 20, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
    
    writeDataTable(pacotes, "RESUMO", setores_data6 %>% mutate(ANO=as.numeric(ANO)), startCol = 2, startRow = 25, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
    
      
    # Apply styles
    addStyle(pacotes, sheet = "RESUMO", style = numstyle, cols = 4:16, rows = c(6:7), gridExpand = TRUE)
    addStyle(pacotes, sheet = "RESUMO", style = numstyle, cols = 4:16, rows = 16, gridExpand = TRUE)
    addStyle(pacotes, sheet = "RESUMO", style = numstyle, cols = 4:16, rows = 21, gridExpand = TRUE)
    addStyle(pacotes, sheet = "RESUMO", style = percstyle, cols = 4:16, rows = 11, gridExpand = TRUE)
    addStyle(pacotes, sheet = "RESUMO", style = percstyle, cols = 4:16, rows = 26, gridExpand = TRUE)
    
    
    crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")
    
    queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")
    
    conditionalFormatting(pacotes, sheet = "RESUMO", cols = 4:16, rows = 11, rule = '> 0', style = crescimento)
    conditionalFormatting(pacotes, sheet = "RESUMO", cols = 4:16, rows = 11, rule = '< 0', style = queda)
    
    conditionalFormatting(pacotes, sheet = "RESUMO", cols = 4:16, rows = 26, rule = '> 0', style = crescimento)
    conditionalFormatting(pacotes, sheet = "RESUMO", cols = 4:16, rows = 26, rule = '< 0', style = queda)
    
    
    # Set column widths
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 1, widths = 12)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 2, widths =40)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 3, widths = 40)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 4, widths = 15)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 5, widths = 15)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 6, widths = 15)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 7, widths = 20)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 8, widths = 40)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 9, widths = 20)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 10, widths = 20)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 11, widths = 20)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 12, widths = 20)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 13, widths = 20)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 14, widths = 20)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 15, widths = 20)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 16, widths = 20)
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 17, widths = 20) 
    
    setColWidths(pacotes, sheet = "LANCAMENTO PACOTES", cols = 19, widths = 20) 
    
    
    writeDataTable(pacotes, "LANCAMENTO PACOTES", setores_data, startCol = 1, startRow = 1, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
    
    addStyle(pacotes, sheet = "LANCAMENTO PACOTES", style = numstyle, cols = c(9,19), rows = 1:nrow(setores_data)+1, gridExpand = TRUE)
    
    addStyle(pacotes, sheet = "LANCAMENTO PACOTES", style = datesty, cols = c(5,6,17), rows = 1:nrow(setores_data)+1, gridExpand = TRUE)
    
    
    setColWidths(pacotes, sheet = "CONSUMO PACOTES", cols = 1, widths = 12)
    
    setColWidths(pacotes, sheet = "CONSUMO PACOTES", cols = 2, widths =17)
    
    setColWidths(pacotes, sheet = "CONSUMO PACOTES", cols = 3, widths = 15)
    
    setColWidths(pacotes, sheet = "CONSUMO PACOTES", cols = 4, widths = 30)
    
    setColWidths(pacotes, sheet = "CONSUMO PACOTES", cols = 5, widths = 17)
    
    setColWidths(pacotes, sheet = "CONSUMO PACOTES", cols = 6, widths = 35)
    
    setColWidths(pacotes, sheet = "CONSUMO PACOTES", cols = 7, widths = 17)
    
    setColWidths(pacotes, sheet = "CONSUMO PACOTES", cols = 8, widths = 40)
    
    setColWidths(pacotes, sheet = "CONSUMO PACOTES", cols = 9, widths = 15)
    
    setColWidths(pacotes, sheet = "CONSUMO PACOTES", cols = 10, widths = 13)
    
    setColWidths(pacotes, sheet = "CONSUMO PACOTES", cols = 11, widths = 13)
    
    
    writeDataTable(pacotes, "CONSUMO PACOTES", setores_data7, startCol = 1, startRow = 1, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
    
    addStyle(pacotes, sheet = "CONSUMO PACOTES", style = numstyle, cols = c(9,19), rows = 1:nrow(setores_data7)+1, gridExpand = TRUE)
    
    addStyle(pacotes, sheet = "CONSUMO PACOTES", style = datesty, cols = 2, rows = 1:nrow(setores_data7)+1, gridExpand = TRUE)
    
    
    setor_prefix <- sub("^(\\D*)(\\d+).*", "\\1_\\2", setor)
    
    file_path <- paste0("C:\\Users\\REPRO SANDRO\\OneDrive - Luxottica Group S.p.A (1)\\PACOTES\\PACOTES_", setor_prefix, ".xlsx")

    saveWorkbook(pacotes, file = file_path, overwrite = TRUE)
    
    }
  
  lapply(unique_setores, create_setores_workbook)
  
  
  ## create excel geral ==================================================
  
  pacotes_geral <- createWorkbook()
  
  addWorksheet(pacotes_geral, "RESUMO")
  addWorksheet(pacotes_geral, "LANCAMENTO PACOTES")
  addWorksheet(pacotes_geral, "CONSUMO PACOTES")
  
  titulo1 <- "RESUMO LANCAMENTO PACOTES 2023/2024 "
  titulo2 <- "DEVOLUCOES PACOTES"
  titulo3 <- "METAS 10%"
  titulo4 <- "DIF METAS x LANÇAMENTOS 2024"
  
  # Set column widths
  setColWidths(pacotes_geral, sheet = "RESUMO", cols = 1, widths = 1)
  
  setColWidths(pacotes_geral, sheet = "RESUMO", cols = 2, widths =20)
  
  setColWidths(pacotes_geral, sheet = "RESUMO", cols = 3, widths = 10)
  
  setColWidths(pacotes_geral, sheet = "RESUMO", cols = c(4:16), widths = 9)
  
  ## styles
  numstyle <- createStyle(numFmt = "#,##0")
  
  percstyle <- createStyle(numFmt = "0.00%")
  
  datesty <- createStyle(numFmt = "dd/MM/yyyy")
  
  crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")
  
  queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")
  
  
  # Apply styles
  addStyle(pacotes_geral, sheet = "RESUMO", style = numstyle, cols = 4:16, rows = 1:24, gridExpand = TRUE)
  addStyle(pacotes_geral, sheet = "RESUMO", style = percstyle, cols = 4:16, rows = 28:36, gridExpand = TRUE)
  addStyle(pacotes_geral, sheet = "RESUMO", style = numstyle, cols = 4:16, rows = 41:48, gridExpand = TRUE)
  addStyle(pacotes_geral, sheet = "RESUMO", style = numstyle, cols = 4:16, rows = 53:61, gridExpand = TRUE)
  addStyle(pacotes_geral, sheet = "RESUMO", style = percstyle, cols = 4:16, rows = 66:74, gridExpand = TRUE)
  
  
  crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")
  
  
  queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")
  
  
  conditionalFormatting(pacotes_geral, sheet = "RESUMO", cols = 4:16, rows = 28:36, rule = '> 0', style = crescimento)
  conditionalFormatting(pacotes_geral, sheet = "RESUMO", cols = 4:16, rows = 28:36, rule = '< 0', style = queda)
  conditionalFormatting(pacotes_geral, sheet = "RESUMO", cols = 4:16, rows = 66:74, rule = '> 0', style = crescimento)
  conditionalFormatting(pacotes_geral, sheet = "RESUMO", cols = 4:16, rows = 66:74, rule = '< 0', style = queda)
  
  
   
  writeData(pacotes_geral, "RESUMO", titulo1, startCol = 2, startRow = 1)
  writeData(pacotes_geral, "RESUMO", titulo2, startCol = 2, startRow = 39)
  writeData(pacotes_geral, "RESUMO", titulo3, startCol = 2, startRow = 51)
  writeData(pacotes_geral, "RESUMO", titulo4, startCol = 2, startRow = 64)
  
  writeDataTable(pacotes_geral, "RESUMO", pacotes_current_year_with_totals %>% mutate(ANO=as.numeric(format(floor_date(Sys.Date(), "year"),"%Y"))) %>%  select(1, ANO, everything()), startCol = 2, startRow = 3, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
  writeDataTable(pacotes_geral, "RESUMO", pacotes_last_year_with_totals %>% mutate(ANO=as.numeric(format(floor_date(Sys.Date()-years(1), "year"),"%Y"))) %>%  select(1, ANO, everything()), startCol = 2, startRow = 15, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
  writeDataTable(pacotes_geral, "RESUMO", varYY, startCol = 2, startRow = 27, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
  writeDataTable(pacotes_geral, "RESUMO", dev_pacotes_current_year_with_totals %>%  mutate(ANO=as.numeric(format(floor_date(Sys.Date(), "year"),"%Y"))), startCol = 2, startRow = 40, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
  writeDataTable(pacotes_geral, "RESUMO", pacotes_last_year_with_totals_metas %>%  mutate(ANO=as.numeric(format(floor_date(Sys.Date(), "year"),"%Y"))), startCol = 2, startRow = 52, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
  writeDataTable(pacotes_geral, "RESUMO", pacotes_last_year_with_totals_metas_dif %>%  mutate(ANO=as.numeric(format(floor_date(Sys.Date(), "year"),"%Y"))), startCol = 2, startRow = 65, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
  
  
  # Set column widths
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 1, widths = 12)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 2, widths =40)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 3, widths = 40)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 4, widths = 15)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 5, widths = 15)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 6, widths = 15)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 7, widths = 20)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 8, widths = 40)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 9, widths = 20)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 10, widths = 20)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 11, widths = 20)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 12, widths = 20)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 13, widths = 20)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 14, widths = 20)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 15, widths = 20)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 16, widths = 20)
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 17, widths = 20)  
  
  setColWidths(pacotes_geral, sheet = "LANCAMENTO PACOTES", cols = 19, widths = 20) 
  
  
  writeDataTable(pacotes_geral, "LANCAMENTO PACOTES", base_pacotes, startCol = 1, startRow = 1, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
  
  addStyle(pacotes_geral, sheet = "LANCAMENTO PACOTES", style = numstyle, cols = c(9,19), rows = 1:nrow(base_pacotes)+1, gridExpand = TRUE)
  
  addStyle(pacotes_geral, sheet = "LANCAMENTO PACOTES", style = datesty, cols = c(5,6,17), rows = 1:nrow(base_pacotes)+1, gridExpand = TRUE)
  

  setColWidths(pacotes_geral, sheet = "CONSUMO PACOTES", cols = 1, widths = 12)
  
  setColWidths(pacotes_geral, sheet = "CONSUMO PACOTES", cols = 2, widths =17)
  
  setColWidths(pacotes_geral, sheet = "CONSUMO PACOTES", cols = 3, widths = 15)
  
  setColWidths(pacotes_geral, sheet = "CONSUMO PACOTES", cols = 4, widths = 30)
  
  setColWidths(pacotes_geral, sheet = "CONSUMO PACOTES", cols = 5, widths = 17)
  
  setColWidths(pacotes_geral, sheet = "CONSUMO PACOTES", cols = 6, widths = 35)
  
  setColWidths(pacotes_geral, sheet = "CONSUMO PACOTES", cols = 7, widths = 17)
  
  setColWidths(pacotes_geral, sheet = "CONSUMO PACOTES", cols = 8, widths = 40)
  
  setColWidths(pacotes_geral, sheet = "CONSUMO PACOTES", cols = 9, widths = 15)
  
  setColWidths(pacotes_geral, sheet = "CONSUMO PACOTES", cols = 10, widths = 13)
  
  setColWidths(pacotes_geral, sheet = "CONSUMO PACOTES", cols = 11, widths = 13)
  
  writeDataTable(pacotes_geral, "CONSUMO PACOTES", consumo_pacotes, startCol = 1, startRow = 1, tableStyle = "TableStyleMedium2",tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)
  
  addStyle(pacotes_geral, sheet = "CONSUMO PACOTES", style = numstyle, cols = c(9,19), rows = 1:nrow(consumo_pacotes)+1, gridExpand = TRUE)
  
  addStyle(pacotes_geral, sheet = "CONSUMO PACOTES", style = datesty, cols = 2, rows = 1:nrow(consumo_pacotes)+1, gridExpand = TRUE)
  
  
  file_path2 <- paste0("C:\\Users\\REPRO SANDRO\\OneDrive - Luxottica Group S.p.A (1)\\PACOTES\\PACOTES_GERAL.xlsx")
  
  saveWorkbook(pacotes_geral, file = file_path2, overwrite = TRUE)
  
  
  
  
  
  
