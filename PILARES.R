## PILARES
## ULTIMA ATUALIZACAO 26.03.2024
## SANDRO JAKOSKA


## LOAD  =======================================================================

library(tidyverse)
library(lubridate)
library(reshape2)
library(DBI)
library(openxlsx)

con2 <- dbConnect(odbc::odbc(), "repro")


## SQL =================================================

  pilares_ytd <- dbGetQuery(con2,"
    
    WITH FIS AS (SELECT FISCODIGO FROM TBFIS WHERE FISTPNATOP IN ('V','R','SR')),
    
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
                                   WHERE PEDDTBAIXA BETWEEN '01.01.2024' AND 'YESTERDAY' AND PEDSITPED<>'C' AND PEDLCFINANC IN ('S', 'L','N')),
            
            VLX AS  (SELECT PROCODIGO FROM PRODU WHERE MARCODIGO=57),
            
            KDK_MF AS  (SELECT PROCODIGO FROM PRODU WHERE MARCODIGO=24 AND GR2CODIGO=1),
            
            KDK_VS AS  (SELECT PROCODIGO FROM PRODU WHERE MARCODIGO=24 AND GR2CODIGO=3),
            
            LA AS  (SELECT PROCODIGO FROM PRODU WHERE GR1CODIGO=2),
            
            TCRZ AS  (SELECT PROCODIGO FROM PRODU WHERE PRODESCRICAO LIKE '%CRIZAL%' AND PROTIPO='T'),
            
            LACRZ AS  (SELECT PROCODIGO FROM PRODU WHERE (PRODESCRICAO LIKE '%CRIZAL%' OR PRODESCRICAO LIKE '%C.FORTE%')  AND GR1CODIGO=2),
            
            MPR AS  (SELECT PROCODIGO FROM PRODU
                            WHERE MARCODIGO IN (128,189,135,106,158,159,264)  AND PROTIPO<>'T'),
                            
            MCLIENTE AS (SELECT PROCODIGO FROM NGRUPOS WHERE GRCODIGO=162),                
            
            TRANS AS  (SELECT PROCODIGO FROM PRODU WHERE (PRODESCRICAO LIKE '%TGEN8%' OR PRODESCRICAO LIKE '%TRANS%'))  
    
            SELECT PEDDTBAIXA,
                    CLICODIGO,
                     PD.PROCODIGO,
                      PDPDESCRICAO,
                       SETOR,
                        CASE 
                         WHEN VX.PROCODIGO IS NOT NULL THEN 'VARILUX'
                          WHEN KDK_MF.PROCODIGO IS NOT NULL THEN 'KODAK MF'
                           WHEN KDK_VS.PROCODIGO IS NOT NULL THEN 'KODAK VS'
                            WHEN MP.PROCODIGO IS NOT NULL THEN 'MARCA REPRO'
                             WHEN LZ.PROCODIGO IS NOT NULL THEN 'LA CRIZAL'
                              WHEN TZ.PROCODIGO IS NOT NULL THEN 'TRAT CRIZAL'
                               WHEN MC.PROCODIGO IS NOT NULL THEN 'MARCA REPRO'
                                ELSE 'OUTROS' END MARCA,
                               IIF (T.PROCODIGO IS NOT NULL,'TRANSITIONS','') TRANSITIONS,
                                CASE 
                                 WHEN L.PROCODIGO IS NOT NULL THEN 'LA'
                                  ELSE '' END TIPO,
                                   SUM(PDPQTDADE)QTD,
                                    SUM(PDPUNITLIQUIDO*PDPQTDADE)VRVENDA 
                                     FROM PDPRD PD
                                      INNER JOIN PED P ON PD.ID_PEDIDO=P.ID_PEDIDO
                                       INNER JOIN FIS F ON PD.FISCODIGO=F.FISCODIGO
                                        LEFT JOIN VLX VX ON PD.PROCODIGO=VX.PROCODIGO
                                         LEFT JOIN KDK_MF  ON PD.PROCODIGO=KDK_MF.PROCODIGO
                                          LEFT JOIN KDK_VS  ON PD.PROCODIGO=KDK_VS.PROCODIGO
                                          LEFT JOIN MPR MP ON PD.PROCODIGO=MP.PROCODIGO
                                           LEFT JOIN LACRZ LZ ON PD.PROCODIGO=LZ.PROCODIGO
                                            LEFT JOIN TCRZ TZ ON PD.PROCODIGO=TZ.PROCODIGO
                                             LEFT JOIN TRANS T ON PD.PROCODIGO=T.PROCODIGO
                                              LEFT JOIN MCLIENTE MC ON PD.PROCODIGO=MC.PROCODIGO
                                               LEFT JOIN LA L ON PD.PROCODIGO=L.PROCODIGO
                                                GROUP BY 1,2,3,4,5,6,7,8")


pilares_ytd_2023 <- get(load("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\BASES\\pilares_ytd_2023.RData"))

pilares <- union_all(pilares_ytd,pilares_ytd_2023) 



##  VALORES ==========================================================================

npilares_setores_marcas <- pilares %>% group_by(SETOR,MARCA) %>% 
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
                          floor_date(PEDDTBAIXA,"day") < floor_date(Sys.Date(), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date(),"year"),
                                        floor_date(Sys.Date(),"day"),by="month")))),
    
    PAST12=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") > ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)])
  ) %>%  mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) %>% mutate(VAR24_VLR=ifelse(is.finite(YTD24-YTD23),YTD24-YTD23,0))


npilares_setores_trans <- pilares %>% filter(TRANSITIONS=='TRANSITIONS') %>% group_by(SETOR,TRANSITIONS) %>% 
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
    
    SEMRECEITA=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") > ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)])
  ) %>%  
  mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) %>% 
  
  mutate(VAR24_VLR=ifelse(is.finite(YTD24-YTD23),YTD24-YTD23,0)) %>% 
  
  mutate(MARCA='TRANSITIONS') %>% .[,-2] %>% .[,c(1,13,3:ncol(.)-1)]


npilares_setores <- union_all(npilares_setores_marcas,npilares_setores_trans) %>% arrange(SETOR)

npilares_setores <- npilares_setores %>%  mutate(STATUS=case_when(
  SEMRECEITA==0 ~ 'SEM RECEITA',
  VAR24>=0 ~ 'CRESCIMENTO',
  VAR24<0 ~ 'QUEDA',
  TRUE ~ ''
))


THISMONTHLASTYEAR1 <- toupper(format(floor_date(Sys.Date()-years(1), "month"),"%b%/%Y"))

LASTMONTHTHISYEAR1 <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH1 <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))


npilares_setores <- npilares_setores  %>% 
  rename_at(3:5,~ c(LASTMONTHTHISYEAR1,THISMONTHLASTYEAR1,CURRENTMONTH1)) %>% 
  .[,c(1:7,12,13,14,8,9)]



##  QUANTIDADE ==========================================================================

npilares_setores_marcas_qtd <- pilares %>% group_by(SETOR,MARCA) %>% 
  summarize(
    
    LASTMONTHTHISYEAR=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    THISMONTHLASTYEAR=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date((Sys.Date()-years(1)), 'month') & floor_date(PEDDTBAIXA,"day")<=floor_date(Sys.Date()-years(1), "day")-1],na.rm = TRUE),
    
    CURRENTMONTH=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date(),'month') %m-% days(1)],na.rm = TRUE),
    
    YTD23=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date() - years(1),"day")-1],na.rm = TRUE),
    
    YTD24=sum(QTD[floor_date(PEDDTBAIXA,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1],na.rm = TRUE),
    
    MEDIA23=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= 
                          floor_date(Sys.Date()-years(1), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date()-years(1),"year"),
                                        floor_date(Sys.Date()-years(1),"month"),by="month")))),
    
    MEDIA24=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date(), "year") & 
                          floor_date(PEDDTBAIXA,"day") < floor_date(Sys.Date(), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date(),"year"),
                                        floor_date(Sys.Date(),"day"),by="month")))),
    
    PAST12=sum(QTD[floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(QTD[floor_date(PEDDTBAIXA,"day") > ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)])
  ) %>%  mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) %>% mutate(VAR24_VLR=ifelse(is.finite(YTD24-YTD23),YTD24-YTD23,0))


npilares_setores_trans_qtd <- pilares %>% filter(TRANSITIONS=='TRANSITIONS') %>% group_by(SETOR,TRANSITIONS) %>% 
  summarize(
    
    LASTMONTHTHISYEAR=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    THISMONTHLASTYEAR=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date((Sys.Date()-years(1)), 'month') & floor_date(PEDDTBAIXA,"day")<=floor_date(Sys.Date()-years(1), "day")-1],na.rm = TRUE),
    
    CURRENTMONTH=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date(),'month') %m-% days(1)],na.rm = TRUE),
    
    YTD23=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date() - years(1),"day")-1],na.rm = TRUE),
    
    YTD24=sum(QTD[floor_date(PEDDTBAIXA,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1],na.rm = TRUE),
    
    MEDIA23=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= 
                          floor_date(Sys.Date()-years(1), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date()-years(1),"year"),
                                        floor_date(Sys.Date()-years(1),"month"),by="month")))),
    
    MEDIA24=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date(), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date(),"year"),
                                        floor_date(Sys.Date(),"day"),by="month")))),
    
    PAST12=sum(QTD[floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(QTD[floor_date(PEDDTBAIXA,"day") > ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)])
  ) %>%  
  mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) %>% 
  
  mutate(VAR24_VLR=ifelse(is.finite(YTD24-YTD23),YTD24-YTD23,0)) %>% 
  
  mutate(MARCA='TRANSITIONS') %>% .[,-2] %>% .[,c(1,13,3:ncol(.)-1)]


npilares_setores_qtd <- union_all(npilares_setores_marcas_qtd,npilares_setores_trans_qtd) %>% arrange(SETOR)

npilares_setores_qtd <- npilares_setores_qtd %>%  mutate(STATUS=case_when(
  SEMRECEITA==0 ~ 'SEM RECEITA',
  VAR24>=0 ~ 'CRESCIMENTO',
  VAR24<0 ~ 'QUEDA',
  TRUE ~ ''
))


THISMONTHLASTYEAR1 <- toupper(format(floor_date(Sys.Date()-years(1), "month"),"%b%/%Y"))

LASTMONTHTHISYEAR1 <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH1 <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))


npilares_setores_qtd <- npilares_setores_qtd  %>% 
  rename_at(3:5,~ c(LASTMONTHTHISYEAR1,THISMONTHLASTYEAR1,CURRENTMONTH1)) %>% 
  .[,c(1:7,12,13,14,8,9)]



## RESUMO VALORES ======================================================================================================  


npilares_setores_marcas_resumo <- pilares %>% group_by(MARCA) %>% 
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
                          floor_date(PEDDTBAIXA,"day") < floor_date(Sys.Date(), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date(),"year"),
                                        floor_date(Sys.Date(),"day"),by="month")))),
    
    PAST12=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") > ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)])
  ) %>%  mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) %>% mutate(VAR24_VLR=ifelse(is.finite(YTD24-YTD23),YTD24-YTD23,0))


npilares_setores_trans_resumo <- pilares %>% filter(TRANSITIONS=='TRANSITIONS') %>% group_by(TRANSITIONS) %>% 
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
    
    SEMRECEITA=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") > ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)])
  ) %>%  
  mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) %>% 
  
  mutate(VAR24_VLR=ifelse(is.finite(YTD24-YTD23),YTD24-YTD23,0)) %>% 
  
  mutate(MARCA='TRANSITIONS') %>% .[,-1] %>% .[,c(12,2:ncol(.)-1)]


npilares_setores_resumo <- union_all(npilares_setores_marcas_resumo,npilares_setores_trans_resumo) 

npilares_setores_resumo <- npilares_setores_resumo %>%  mutate(STATUS=case_when(
  SEMRECEITA==0 ~ 'SEM RECEITA',
  VAR24>=0 ~ 'CRESCIMENTO',
  VAR24<0 ~ 'QUEDA',
  TRUE ~ ''
))


THISMONTHLASTYEAR1 <- toupper(format(floor_date(Sys.Date()-years(1), "month"),"%b%/%Y"))

LASTMONTHTHISYEAR1 <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH1 <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))


npilares_setores_resumo <- npilares_setores_resumo  %>% 
  rename_at(2:4,~ c(LASTMONTHTHISYEAR1,THISMONTHLASTYEAR1,CURRENTMONTH1)) %>% 
  .[,c(1:6,11,12,13,7,8)]

## QTD ==========================================================================

npilares_setores_marcas_qtd_resumo <- pilares %>% group_by(MARCA) %>% 
  summarize(
    
    LASTMONTHTHISYEAR=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    THISMONTHLASTYEAR=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date((Sys.Date()-years(1)), 'month') & floor_date(PEDDTBAIXA,"day")<=floor_date(Sys.Date()-years(1), "day")-1],na.rm = TRUE),
    
    CURRENTMONTH=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date(),'month') %m-% days(1)],na.rm = TRUE),
    
    YTD23=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date() - years(1),"day")-1],na.rm = TRUE),
    
    YTD24=sum(QTD[floor_date(PEDDTBAIXA,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1],na.rm = TRUE),
    
    MEDIA23=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= 
                          floor_date(Sys.Date()-years(1), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date()-years(1),"year"),
                                        floor_date(Sys.Date()-years(1),"month"),by="month")))),
    
    MEDIA24=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date(), "year") & 
                          floor_date(PEDDTBAIXA,"day") < floor_date(Sys.Date(), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date(),"year"),
                                        floor_date(Sys.Date(),"day"),by="month")))),
    
    PAST12=sum(QTD[floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(QTD[floor_date(PEDDTBAIXA,"day") > ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)])
  ) %>%  mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) %>% mutate(VAR24_VLR=ifelse(is.finite(YTD24-YTD23),YTD24-YTD23,0))


npilares_setores_trans_qtd_resumo <- pilares %>% filter(TRANSITIONS=='TRANSITIONS') %>% group_by(TRANSITIONS) %>% 
  summarize(
    
    LASTMONTHTHISYEAR=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    THISMONTHLASTYEAR=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date((Sys.Date()-years(1)), 'month') & floor_date(PEDDTBAIXA,"day")<=floor_date(Sys.Date()-years(1), "day")-1],na.rm = TRUE),
    
    CURRENTMONTH=sum(QTD[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date(),'month') %m-% days(1)],na.rm = TRUE),
    
    YTD23=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date() - years(1),"day")-1],na.rm = TRUE),
    
    YTD24=sum(QTD[floor_date(PEDDTBAIXA,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1],na.rm = TRUE),
    
    MEDIA23=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= 
                          floor_date(Sys.Date()-years(1), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date()-years(1),"year"),
                                        floor_date(Sys.Date()-years(1),"month"),by="month")))),
    
    MEDIA24=sum(QTD[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date(), "year") & 
                          floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "day")-1]/
                  as.numeric(length(seq(floor_date(Sys.Date(),"year"),
                                        floor_date(Sys.Date(),"day"),by="month")))),
    
    PAST12=sum(QTD[floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(QTD[floor_date(PEDDTBAIXA,"day") > ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)])
  ) %>%  
  mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) %>% 
  
  mutate(VAR24_VLR=ifelse(is.finite(YTD24-YTD23),YTD24-YTD23,0)) %>% 
  
  mutate(MARCA='TRANSITIONS') %>% .[,-1] %>% .[,c(12,2:ncol(.)-1)]


npilares_setores_qtd_resumo <- union_all(npilares_setores_marcas_qtd_resumo,npilares_setores_trans_qtd_resumo) 

npilares_setores_qtd_resumo <- npilares_setores_qtd_resumo %>%  mutate(STATUS=case_when(
  SEMRECEITA==0 ~ 'SEM RECEITA',
  VAR24>=0 ~ 'CRESCIMENTO',
  VAR24<0 ~ 'QUEDA',
  TRUE ~ ''
))


THISMONTHLASTYEAR1 <- toupper(format(floor_date(Sys.Date()-years(1), "month"),"%b%/%Y"))

LASTMONTHTHISYEAR1 <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH1 <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))


npilares_setores_qtd_resumo <- npilares_setores_qtd_resumo  %>% 
  rename_at(2:4,~ c(LASTMONTHTHISYEAR1,THISMONTHLASTYEAR1,CURRENTMONTH1)) %>% 
  .[,c(1:6,11,12,13,7,8)]



## WRITE EXCEL  ==========================================================================


wb_pilares <- createWorkbook()


## sheet ranking

addWorksheet(wb_pilares, "RESUMO")



## format cols

setColWidths(wb_pilares, sheet = "RESUMO", cols = 1, widths = 15)

setColWidths(wb_pilares, sheet = "RESUMO", cols = 2:8, widths = 12)

setColWidths(wb_pilares, sheet = "RESUMO", cols = 9, widths = 15)


crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb_pilares, sheet = "RESUMO", style = estiloNumerico, cols = c(2:6,8,10:11), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_pilares, sheet = "RESUMO", style = estiloPorcentagem, cols = 7, rows = 1:3000, gridExpand = TRUE)


## cond format class

conditionalFormatting(wb_pilares, sheet = "RESUMO", cols = 9, rows = 1:1000, rule = 'I1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_pilares, sheet = "RESUMO", cols = 9, rows = 1:1000, rule = 'I1 == "QUEDA"', style = queda)

conditionalFormatting(wb_pilares, sheet = "RESUMO", cols = 7, rows = 1:1000, rule = 'G1 < 0', style = queda2)

conditionalFormatting(wb_pilares, sheet = "RESUMO", cols = 8, rows = 1:1000, rule = 'H1 < 0', style = queda2)

writeData(wb_pilares, "RESUMO", "VALOR", startCol = 1, startRow = 2)

writeDataTable(wb_pilares, "RESUMO", npilares_setores_resumo, startCol = 1, startRow = 3, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

writeData(wb_pilares, "RESUMO", "QTD", startCol = 1, startRow = 13)

writeDataTable(wb_pilares, "RESUMO", npilares_setores_qtd_resumo, startCol = 1, startRow = 14, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)




## SETORES  VALOR ===================


addWorksheet(wb_pilares, "SETORES VALOR")


## format cols

setColWidths(wb_pilares, sheet = "SETORES VALOR", cols = 1, widths = 30)

setColWidths(wb_pilares, sheet = "SETORES VALOR", cols = 2, widths = 15)

setColWidths(wb_pilares, sheet = "SETORES VALOR", cols = 3:9, widths = 12)

setColWidths(wb_pilares, sheet = "SETORES VALOR", cols = 10, widths = 15)

setColWidths(wb_pilares, sheet = "SETORES VALOR", cols = 11:12, widths = 15)


crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb_pilares, sheet = "SETORES VALOR", style = estiloNumerico, cols = c(3:7,9,11:12), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_pilares, sheet = "SETORES VALOR", style = estiloPorcentagem, cols = 8, rows = 1:3000, gridExpand = TRUE)


## cond format class

conditionalFormatting(wb_pilares, sheet = "SETORES VALOR", cols = 10, rows = 1:1000, rule = 'J1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_pilares, sheet = "SETORES VALOR", cols = 10, rows = 1:1000, rule = 'J1 == "QUEDA"', style = queda)

conditionalFormatting(wb_pilares, sheet = "SETORES VALOR", cols = 8, rows = 1:1000, rule = 'H1 < 0', style = queda2)

conditionalFormatting(wb_pilares, sheet = "SETORES VALOR", cols = 9, rows = 1:1000, rule = 'I1 < 0', style = queda2)



writeDataTable(wb_pilares, "SETORES VALOR", npilares_setores, startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

## SETORES  QTD  ===================


addWorksheet(wb_pilares, "SETORES QTD")


## format cols

setColWidths(wb_pilares, sheet = "SETORES QTD", cols = 1, widths = 30)

setColWidths(wb_pilares, sheet = "SETORES QTD", cols = 2, widths = 15)

setColWidths(wb_pilares, sheet = "SETORES QTD", cols = 3:9, widths = 12)

setColWidths(wb_pilares, sheet = "SETORES QTD", cols = 10, widths = 15)

setColWidths(wb_pilares, sheet = "SETORES QTD", cols = 11:12, widths = 15)


crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb_pilares, sheet = "SETORES QTD", style = estiloNumerico, cols = c(3:7,9,11:12), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_pilares, sheet = "SETORES QTD", style = estiloPorcentagem, cols = 8, rows = 1:3000, gridExpand = TRUE)


## cond format class

conditionalFormatting(wb_pilares, sheet = "SETORES QTD", cols = 10, rows = 1:1000, rule = 'J1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_pilares, sheet = "SETORES QTD", cols = 10, rows = 1:1000, rule = 'J1 == "QUEDA"', style = queda)

conditionalFormatting(wb_pilares, sheet = "SETORES QTD", cols = 8, rows = 1:1000, rule = 'H1 < 0', style = queda2)

conditionalFormatting(wb_pilares, sheet = "SETORES QTD", cols = 9, rows = 1:1000, rule = 'I1 < 0', style = queda2)



writeDataTable(wb_pilares, "SETORES QTD", npilares_setores_qtd, startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = FALSE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## save workbook

saveWorkbook(wb_pilares, file = "PILARES.xlsx", overwrite = TRUE)



