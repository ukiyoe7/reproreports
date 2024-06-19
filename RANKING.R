## RANKING CLIENTS
## LAST UPDATE 12.04.2024
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


##  RANKING CLIENTS ==========================================================================

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
  rename_at(10:12,~ c(LASTMONTHTHISYEAR,THISMONTHLASTYEAR,CURRENTMONTH))

corder <- c("CLICODIGO","NOMEFANTASIA","CODGRUPO","CLIENTE","GRUPO","SETOR","CIDADE",LASTMONTHTHISYEAR,THISMONTHLASTYEAR,CURRENTMONTH,"YTD23","YTD24","VAR24","VAR24_VLR","STATUS","MEDIA23","MEDIA24","VARILUX DIGI","VARILUX TRAD","VARILUX XR","KODAK DIGITAL","KODAK TRAD","ACTUALITE","AVANCE","UZ+","IMAGEM DIGITAL","IMAGEM TRAD","INSIGNE","MARCAS PROPRIAS","EYEZEN","LA CRIZAL","LA KODAK","PROSING","PROSING POLY","VS ESSILOR","DESCTGERAL")

data <- data %>% .[,corder] %>% mutate_all(~replace(., is.nan(.), NA))


##  RANKING GRUPOS ==========================================================================


dt <- left_join(anti_join(clientes,inativos,by="CLICODIGO"),sales,by="CLICODIGO") %>% as.data.frame()

gsales <- dt %>% group_by(SETOR,CODGRUPO,GRUPO) %>% 
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
    
    SEMRECEITA=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") >  ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)])
  ) %>% mutate(VAR24=ifelse(is.finite(YTD24/YTD23-1),YTD24/YTD23-1,0)) %>% 
  
  mutate(VAR24_VLR=ifelse(is.finite(YTD24-YTD23),YTD24-YTD23,0)) 

gsales <- gsales %>%  mutate(STATUS=case_when(
  SEMRECEITA==0 ~ 'SEM RECEITA',
  VAR24>0 ~ 'CRESCIMENTO',
  VAR24<0 ~ 'QUEDA',
  PAST12==0  ~ 'PERDIDO',
  YTD23==0 & YTD24>0 ~ 'RECUPERADO',
  TRUE ~ ''
))

gdata <-apply(gsales[,4:14],2,function(x) round(x,2)) %>% as.data.frame() %>% 
  cbind(gsales[,1:3],.) %>% cbind(.,gsales[,15])


gdata <- gdata %>% left_join(.,gdesct,by="CODGRUPO") %>% as.data.frame()

gdata <- gdata %>% filter(CODGRUPO!='')


gdata <- gdata %>% arrange(desc(.$YTD24)) %>% as.data.frame() %>% 
  rename_at(4:6,~ c(LASTMONTHTHISYEAR1,THISMONTHLASTYEAR1,CURRENTMONTH1)) 

gcorder <- c("CODGRUPO","SETOR","GRUPO",LASTMONTHTHISYEAR,THISMONTHLASTYEAR,CURRENTMONTH,"YTD23","YTD24","VAR24","VAR24_VLR","STATUS","MEDIA23","MEDIA24","VARILUX DIGI","VARILUX TRAD","VARILUX XR","KODAK DIGITAL","KODAK TRAD","ACTUALITE","AVANCE","UZ+","IMAGEM DIGITAL","IMAGEM TRAD","INSIGNE","MARCAS PROPRIAS","EYEZEN","LA CRIZAL","LA KODAK","PROSING","PROSING POLY","VS ESSILOR")

gdata <- gdata %>% .[,gcorder] %>% mutate_all(~replace(., is.nan(.), NA))


## ESPERADO =====================================

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

## CREATE EXCEL ===============================

## RANKING CLIENTES ===========================

## GERAL

wb <- createWorkbook()


addWorksheet(wb, "RANKING")

writeData(wb, sheet = "RANKING", x = data)

setColWidths(wb, sheet = "RANKING", cols = 1, widths = 12)

setColWidths(wb, sheet = "RANKING", cols = 2, widths = 40)

setColWidths(wb, sheet = "RANKING", cols = 3, widths = 12)

setColWidths(wb, sheet = "RANKING", cols = 4, widths = 33)

setColWidths(wb, sheet = "RANKING", cols = 5, widths = 30)

setColWidths(wb, sheet = "RANKING", cols = 6, widths = 33)

setColWidths(wb, sheet = "RANKING", cols = 7, widths = 36)

setColWidths(wb, sheet = "RANKING", cols = 7:12, widths = 12)

setColWidths(wb, sheet = "RANKING", cols = 13, widths = 13)

setColWidths(wb, sheet = "RANKING", cols = 14, widths = 15)

setColWidths(wb, sheet = "RANKING", cols = 15, widths = 17)

setColWidths(wb, sheet = "RANKING", cols = 16:17, widths = 15)

setColWidths(wb, sheet = "RANKING", cols = 18:36, widths = 18)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb, sheet = "RANKING", style = estiloNumerico, cols = c(8:12,14,16:36), rows = 1:3000, gridExpand = TRUE)

addStyle(wb, sheet = "RANKING", style = estiloPorcentagem, cols = 13, rows = 1:3000, gridExpand = TRUE)

## cond format class

conditionalFormatting(wb, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "QUEDA"', style = queda)

conditionalFormatting(wb, sheet = "RANKING", cols = 13, rows = 1:1000, rule = 'M1 < 0', style = queda2)

conditionalFormatting(wb, sheet = "RANKING", cols = 14, rows = 1:1000, rule = 'N1 < 0', style = queda2)


writeDataTable(wb, "RANKING", data, startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## RANKING GRUPOS ===========================


addWorksheet(wb, "RANKING GRUPOS")

writeData(wb, sheet = "RANKING GRUPOS", x = gdata)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 1, widths = 12)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 2, widths = 35)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 3, widths = 35)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 4:9, widths = 15)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 10, widths = 13)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 11, widths = 17)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 12:13, widths = 15)

setColWidths(wb, sheet = "RANKING GRUPOS", cols = 14:31, widths = 17)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")


addStyle(wb, sheet = "RANKING GRUPOS", style = estiloNumerico, cols = c(4:8,10,12:30), rows = 1:3000, gridExpand = TRUE)

addStyle(wb, sheet = "RANKING GRUPOS", style = estiloPorcentagem, cols = 9, rows = 1:3000, gridExpand = TRUE)


conditionalFormatting(wb, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "QUEDA"', style = queda)

conditionalFormatting(wb, sheet = "RANKING GRUPOS", cols = 9, rows = 1:1000, rule = 'I1 < 0', style = queda2)

conditionalFormatting(wb, sheet = "RANKING GRUPOS", cols = 10, rows = 1:1000, rule = 'J1 < 0', style = queda2)

writeDataTable(wb, "RANKING GRUPOS", gdata, startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## ESPERADO ===========================================================


## GERAL

addWorksheet(wb, "GERAL ESPERADO")

writeData(wb, sheet = "GERAL ESPERADO", x = sales_sm_1_3)

writeDataTable(wb, "GERAL ESPERADO", sales_sm_1_3, startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

setColWidths(wb, sheet = "GERAL ESPERADO", cols = 1, widths = 40)

setColWidths(wb, sheet = "GERAL ESPERADO", cols = c(5,6), widths = 30)

setColWidths(wb, sheet = "GERAL ESPERADO", cols = c(2,3,4,7,8,9,10), widths = 15)



## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb, sheet = "GERAL ESPERADO", style = estiloNumerico, cols = c(2,3,4,6,7,8,9,10), rows = 1:3000, gridExpand = TRUE)

addStyle(wb, sheet = "GERAL ESPERADO", style = estiloPorcentagem, cols = 5, rows = 1:3000, gridExpand = TRUE)


## SETORES

addWorksheet(wb, "SETORES ESPERADO")

writeData(wb, sheet = "SETORES ESPERADO", x = sales_sm_2_3)

writeDataTable(wb, "SETORES ESPERADO", sales_sm_2_3, startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

setColWidths(wb, sheet = "SETORES ESPERADO", cols = 1, widths = 50)

setColWidths(wb, sheet = "SETORES ESPERADO", cols = 2, widths = 30)

setColWidths(wb, sheet = "SETORES ESPERADO", cols = c(3,4,5), widths = 15)

setColWidths(wb, sheet = "SETORES ESPERADO", cols = 6, widths =30)

setColWidths(wb, sheet = "SETORES ESPERADO", cols = 7, widths =30)

setColWidths(wb, sheet = "SETORES ESPERADO", cols = c(8:11), widths = 15)

addStyle(wb, sheet = "SETORES ESPERADO", style = estiloNumerico, cols = c(3,4,5,7,8,9,10,11), rows = 1:3000, gridExpand = TRUE)

addStyle(wb, sheet = "SETORES ESPERADO", style = estiloPorcentagem, cols = 6, rows = 1:3000, gridExpand = TRUE)

## cond format class

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")


conditionalFormatting(wb, sheet = "GERAL ESPERADO", cols = 5, rows = 1:1000, rule = 'E1 < 0', style = queda2)

conditionalFormatting(wb, sheet = "GERAL ESPERADO", cols = 5, rows = 1:1000, rule = 'E1 < 0', style = queda2)

conditionalFormatting(wb, sheet = "GERAL ESPERADO", cols = 6, rows = 2:1000, rule = 'F2>0', style = crescimento)

conditionalFormatting(wb, sheet = "GERAL ESPERADO", cols = 6, rows = 2:1000, rule = 'F2<0', style = queda)




conditionalFormatting(wb, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2>0', style = crescimento)

conditionalFormatting(wb, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2<0', style = queda)



saveWorkbook(wb, file = "RANKING_CLIENTES.xlsx", overwrite = TRUE)

## SETOR 1 ===========================================================================


wb_st1 <- createWorkbook()


addWorksheet(wb_st1, "RANKING")

writeData(wb_st1, sheet = "RANKING", x = data %>% filter(SETOR=="SETOR 1 - FLORIANOPOLIS REDES"))

setColWidths(wb_st1, sheet = "RANKING", cols = 1, widths = 12)

setColWidths(wb_st1, sheet = "RANKING", cols = 2, widths = 40)

setColWidths(wb_st1, sheet = "RANKING", cols = 3, widths = 12)

setColWidths(wb_st1, sheet = "RANKING", cols = 4, widths = 33)

setColWidths(wb_st1, sheet = "RANKING", cols = 5, widths = 30)

setColWidths(wb_st1, sheet = "RANKING", cols = 6, widths = 33)

setColWidths(wb_st1, sheet = "RANKING", cols = 7, widths = 36)

setColWidths(wb_st1, sheet = "RANKING", cols = 7:12, widths = 12)

setColWidths(wb_st1, sheet = "RANKING", cols = 13, widths = 13)

setColWidths(wb_st1, sheet = "RANKING", cols = 14, widths = 15)

setColWidths(wb_st1, sheet = "RANKING", cols = 15, widths = 17)

setColWidths(wb_st1, sheet = "RANKING", cols = 16:17, widths = 15)

setColWidths(wb_st1, sheet = "RANKING", cols = 18:36, widths = 18)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb_st1, sheet = "RANKING", style = estiloNumerico, cols = c(8:12,14,16:36), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st1, sheet = "RANKING", style = estiloPorcentagem, cols = 13, rows = 1:3000, gridExpand = TRUE)

## cond format class

conditionalFormatting(wb_st1, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st1, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st1, sheet = "RANKING", cols = 13, rows = 1:1000, rule = 'M1 < 0', style = queda2)

conditionalFormatting(wb_st1, sheet = "RANKING", cols = 14, rows = 1:1000, rule = 'N1 < 0', style = queda2)


writeDataTable(wb_st1, "RANKING", x = data %>% filter(SETOR=="SETOR 1 - FLORIANOPOLIS REDES"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## RANKING GRUPOS ===========================


addWorksheet(wb_st1, "RANKING GRUPOS")

writeData(wb_st1, sheet = "RANKING GRUPOS", x = gdata)

setColWidths(wb_st1, sheet = "RANKING GRUPOS", cols = 1, widths = 12)

setColWidths(wb_st1, sheet = "RANKING GRUPOS", cols = 2, widths = 35)

setColWidths(wb_st1, sheet = "RANKING GRUPOS", cols = 3, widths = 35)

setColWidths(wb_st1, sheet = "RANKING GRUPOS", cols = 4:9, widths = 15)

setColWidths(wb_st1, sheet = "RANKING GRUPOS", cols = 10, widths = 13)

setColWidths(wb_st1, sheet = "RANKING GRUPOS", cols = 11, widths = 17)

setColWidths(wb_st1, sheet = "RANKING GRUPOS", cols = 12:13, widths = 15)

setColWidths(wb_st1, sheet = "RANKING GRUPOS", cols = 14:31, widths = 17)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")


addStyle(wb_st1, sheet = "RANKING GRUPOS", style = estiloNumerico, cols = c(4:8,10,12:30), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st1, sheet = "RANKING GRUPOS", style = estiloPorcentagem, cols = 9, rows = 1:3000, gridExpand = TRUE)


conditionalFormatting(wb_st1, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st1, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st1, sheet = "RANKING GRUPOS", cols = 9, rows = 1:1000, rule = 'I1 < 0', style = queda2)

conditionalFormatting(wb_st1, sheet = "RANKING GRUPOS", cols = 10, rows = 1:1000, rule = 'J1 < 0', style = queda2)

writeDataTable(wb_st1, "RANKING GRUPOS", gdata %>% filter(SETOR=="SETOR 1 - FLORIANOPOLIS REDES"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## ESPERADO ===========================================================


## SETORES

addWorksheet(wb_st1, "SETORES ESPERADO")

writeData(wb_st1, sheet = "SETORES ESPERADO", x = sales_sm_2_3)

writeDataTable(wb_st1, "SETORES ESPERADO", sales_sm_2_3 %>% filter(SETOR=="SETOR 1 - FLORIANOPOLIS REDES"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

setColWidths(wb_st1, sheet = "SETORES ESPERADO", cols = 1, widths = 50)

setColWidths(wb_st1, sheet = "SETORES ESPERADO", cols = 2, widths = 30)

setColWidths(wb_st1, sheet = "SETORES ESPERADO", cols = c(3,4,5), widths = 15)

setColWidths(wb_st1, sheet = "SETORES ESPERADO", cols = 6, widths =30)

setColWidths(wb_st1, sheet = "SETORES ESPERADO", cols = 7, widths =30)

setColWidths(wb_st1, sheet = "SETORES ESPERADO", cols = c(8:11), widths = 15)

addStyle(wb_st1, sheet = "SETORES ESPERADO", style = estiloNumerico, cols = c(3,4,5,7,8,9,10,11), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st1, sheet = "SETORES ESPERADO", style = estiloPorcentagem, cols = 6, rows = 1:3000, gridExpand = TRUE)

## cond format class

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")



conditionalFormatting(wb_st1, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st1, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st1, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2>0', style = crescimento)

conditionalFormatting(wb_st1, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2<0', style = queda)



saveWorkbook(wb_st1, file = "RANKING_CLIENTES_SETOR1.xlsx", overwrite = TRUE)

## SETOR 2 ===========================================================================


wb_st2 <- createWorkbook()


addWorksheet(wb_st2, "RANKING")

writeData(wb_st2, sheet = "RANKING", x = data %>% filter(SETOR=="SETOR 2 - CRICIUMA - SUL"))

setColWidths(wb_st2, sheet = "RANKING", cols = 1, widths = 12)

setColWidths(wb_st2, sheet = "RANKING", cols = 2, widths = 40)

setColWidths(wb_st2, sheet = "RANKING", cols = 3, widths = 12)

setColWidths(wb_st2, sheet = "RANKING", cols = 4, widths = 33)

setColWidths(wb_st2, sheet = "RANKING", cols = 5, widths = 30)

setColWidths(wb_st2, sheet = "RANKING", cols = 6, widths = 33)

setColWidths(wb_st2, sheet = "RANKING", cols = 7, widths = 36)

setColWidths(wb_st2, sheet = "RANKING", cols = 7:12, widths = 12)

setColWidths(wb_st2, sheet = "RANKING", cols = 13, widths = 13)

setColWidths(wb_st2, sheet = "RANKING", cols = 14, widths = 15)

setColWidths(wb_st2, sheet = "RANKING", cols = 15, widths = 17)

setColWidths(wb_st2, sheet = "RANKING", cols = 16:17, widths = 15)

setColWidths(wb_st2, sheet = "RANKING", cols = 18:36, widths = 18)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb_st2, sheet = "RANKING", style = estiloNumerico, cols = c(8:12,14,16:36), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st2, sheet = "RANKING", style = estiloPorcentagem, cols = 13, rows = 1:3000, gridExpand = TRUE)

## cond format class

conditionalFormatting(wb_st2, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st2, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st2, sheet = "RANKING", cols = 13, rows = 1:1000, rule = 'M1 < 0', style = queda2)

conditionalFormatting(wb_st2, sheet = "RANKING", cols = 14, rows = 1:1000, rule = 'N1 < 0', style = queda2)


writeDataTable(wb_st2, "RANKING", x = data %>% filter(SETOR=="SETOR 2 - CRICIUMA - SUL"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## RANKING GRUPOS ===========================


addWorksheet(wb_st2, "RANKING GRUPOS")

writeData(wb_st2, sheet = "RANKING GRUPOS", x = gdata)

setColWidths(wb_st2, sheet = "RANKING GRUPOS", cols = 1, widths = 12)

setColWidths(wb_st2, sheet = "RANKING GRUPOS", cols = 2, widths = 35)

setColWidths(wb_st2, sheet = "RANKING GRUPOS", cols = 3, widths = 35)

setColWidths(wb_st2, sheet = "RANKING GRUPOS", cols = 4:9, widths = 15)

setColWidths(wb_st2, sheet = "RANKING GRUPOS", cols = 10, widths = 13)

setColWidths(wb_st2, sheet = "RANKING GRUPOS", cols = 11, widths = 17)

setColWidths(wb_st2, sheet = "RANKING GRUPOS", cols = 12:13, widths = 15)

setColWidths(wb_st2, sheet = "RANKING GRUPOS", cols = 14:31, widths = 17)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")


addStyle(wb_st2, sheet = "RANKING GRUPOS", style = estiloNumerico, cols = c(4:8,10,12:30), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st2, sheet = "RANKING GRUPOS", style = estiloPorcentagem, cols = 9, rows = 1:3000, gridExpand = TRUE)


conditionalFormatting(wb_st2, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st2, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st2, sheet = "RANKING GRUPOS", cols = 9, rows = 1:1000, rule = 'I1 < 0', style = queda2)

conditionalFormatting(wb_st2, sheet = "RANKING GRUPOS", cols = 10, rows = 1:1000, rule = 'J1 < 0', style = queda2)

writeDataTable(wb_st2, "RANKING GRUPOS", gdata %>% filter(SETOR=="SETOR 2 - CRICIUMA - SUL"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## ESPERADO ===========================================================


## SETORES

addWorksheet(wb_st2, "SETORES ESPERADO")

writeData(wb_st2, sheet = "SETORES ESPERADO", x = sales_sm_2_3)

writeDataTable(wb_st2, "SETORES ESPERADO", sales_sm_2_3 %>% filter(SETOR=="SETOR 2 - CRICIUMA - SUL"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

setColWidths(wb_st2, sheet = "SETORES ESPERADO", cols = 1, widths = 50)

setColWidths(wb_st2, sheet = "SETORES ESPERADO", cols = 2, widths = 30)

setColWidths(wb_st2, sheet = "SETORES ESPERADO", cols = c(3,4,5), widths = 15)

setColWidths(wb_st2, sheet = "SETORES ESPERADO", cols = 6, widths =30)

setColWidths(wb_st2, sheet = "SETORES ESPERADO", cols = 7, widths =30)

setColWidths(wb_st2, sheet = "SETORES ESPERADO", cols = c(8:11), widths = 15)

addStyle(wb_st2, sheet = "SETORES ESPERADO", style = estiloNumerico, cols = c(3,4,5,7,8,9,10,11), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st2, sheet = "SETORES ESPERADO", style = estiloPorcentagem, cols = 6, rows = 1:3000, gridExpand = TRUE)

## cond format class

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")



conditionalFormatting(wb_st2, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st2, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st2, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2>0', style = crescimento)

conditionalFormatting(wb_st2, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2<0', style = queda)



saveWorkbook(wb_st2, file = "RANKING_CLIENTES_SETOR2.xlsx", overwrite = TRUE)

## SETOR 3 ===========================================================================


wb_st3 <- createWorkbook()


addWorksheet(wb_st3, "RANKING")

writeData(wb_st3, sheet = "RANKING", x = data %>% filter(SETOR=="SETOR 3 - CHAPECO - OESTE - RS"))

setColWidths(wb_st3, sheet = "RANKING", cols = 1, widths = 12)

setColWidths(wb_st3, sheet = "RANKING", cols = 2, widths = 40)

setColWidths(wb_st3, sheet = "RANKING", cols = 3, widths = 12)

setColWidths(wb_st3, sheet = "RANKING", cols = 4, widths = 33)

setColWidths(wb_st3, sheet = "RANKING", cols = 5, widths = 30)

setColWidths(wb_st3, sheet = "RANKING", cols = 6, widths = 33)

setColWidths(wb_st3, sheet = "RANKING", cols = 7, widths = 36)

setColWidths(wb_st3, sheet = "RANKING", cols = 7:12, widths = 12)

setColWidths(wb_st3, sheet = "RANKING", cols = 13, widths = 13)

setColWidths(wb_st3, sheet = "RANKING", cols = 14, widths = 15)

setColWidths(wb_st3, sheet = "RANKING", cols = 15, widths = 17)

setColWidths(wb_st3, sheet = "RANKING", cols = 16:17, widths = 15)

setColWidths(wb_st3, sheet = "RANKING", cols = 18:36, widths = 18)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb_st3, sheet = "RANKING", style = estiloNumerico, cols = c(8:12,14,16:36), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st3, sheet = "RANKING", style = estiloPorcentagem, cols = 13, rows = 1:3000, gridExpand = TRUE)

## cond format class

conditionalFormatting(wb_st3, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st3, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st3, sheet = "RANKING", cols = 13, rows = 1:1000, rule = 'M1 < 0', style = queda2)

conditionalFormatting(wb_st3, sheet = "RANKING", cols = 14, rows = 1:1000, rule = 'N1 < 0', style = queda2)


writeDataTable(wb_st3, "RANKING", x = data %>% filter(SETOR=="SETOR 3 - CHAPECO - OESTE - RS"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## RANKING GRUPOS ===========================


addWorksheet(wb_st3, "RANKING GRUPOS")

writeData(wb_st3, sheet = "RANKING GRUPOS", x = gdata)

setColWidths(wb_st3, sheet = "RANKING GRUPOS", cols = 1, widths = 12)

setColWidths(wb_st3, sheet = "RANKING GRUPOS", cols = 2, widths = 35)

setColWidths(wb_st3, sheet = "RANKING GRUPOS", cols = 3, widths = 35)

setColWidths(wb_st3, sheet = "RANKING GRUPOS", cols = 4:9, widths = 15)

setColWidths(wb_st3, sheet = "RANKING GRUPOS", cols = 10, widths = 13)

setColWidths(wb_st3, sheet = "RANKING GRUPOS", cols = 11, widths = 17)

setColWidths(wb_st3, sheet = "RANKING GRUPOS", cols = 12:13, widths = 15)

setColWidths(wb_st3, sheet = "RANKING GRUPOS", cols = 14:31, widths = 17)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")


addStyle(wb_st3, sheet = "RANKING GRUPOS", style = estiloNumerico, cols = c(4:8,10,12:30), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st3, sheet = "RANKING GRUPOS", style = estiloPorcentagem, cols = 9, rows = 1:3000, gridExpand = TRUE)


conditionalFormatting(wb_st3, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st3, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st3, sheet = "RANKING GRUPOS", cols = 9, rows = 1:1000, rule = 'I1 < 0', style = queda2)

conditionalFormatting(wb_st3, sheet = "RANKING GRUPOS", cols = 10, rows = 1:1000, rule = 'J1 < 0', style = queda2)

writeDataTable(wb_st3, "RANKING GRUPOS", gdata %>% filter(SETOR=="SETOR 3 - CHAPECO - OESTE - RS"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## ESPERADO ===========================================================


## SETORES

addWorksheet(wb_st3, "SETORES ESPERADO")

writeData(wb_st3, sheet = "SETORES ESPERADO", x = sales_sm_2_3)

writeDataTable(wb_st3, "SETORES ESPERADO", sales_sm_2_3 %>% filter(SETOR=="SETOR 3 - CHAPECO - OESTE - RS"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

setColWidths(wb_st3, sheet = "SETORES ESPERADO", cols = 1, widths = 50)

setColWidths(wb_st3, sheet = "SETORES ESPERADO", cols = 2, widths = 30)

setColWidths(wb_st3, sheet = "SETORES ESPERADO", cols = c(3,4,5), widths = 15)

setColWidths(wb_st3, sheet = "SETORES ESPERADO", cols = 6, widths =30)

setColWidths(wb_st3, sheet = "SETORES ESPERADO", cols = 7, widths =30)

setColWidths(wb_st3, sheet = "SETORES ESPERADO", cols = c(8:11), widths = 15)

addStyle(wb_st3, sheet = "SETORES ESPERADO", style = estiloNumerico, cols = c(3,4,5,7,8,9,10,11), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st3, sheet = "SETORES ESPERADO", style = estiloPorcentagem, cols = 6, rows = 1:3000, gridExpand = TRUE)

## cond format class

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")



conditionalFormatting(wb_st3, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st3, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st3, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2>0', style = crescimento)

conditionalFormatting(wb_st3, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2<0', style = queda)



saveWorkbook(wb_st3, file = "RANKING_CLIENTES_SETOR3.xlsx", overwrite = TRUE)

## SETOR 4 ===========================================================================


wb_st4 <- createWorkbook()


addWorksheet(wb_st4, "RANKING")

writeData(wb_st4, sheet = "RANKING", x = data %>% filter(SETOR=="SETOR 4 - JOINVILLE - NORTE"))

setColWidths(wb_st4, sheet = "RANKING", cols = 1, widths = 12)

setColWidths(wb_st4, sheet = "RANKING", cols = 2, widths = 40)

setColWidths(wb_st4, sheet = "RANKING", cols = 3, widths = 12)

setColWidths(wb_st4, sheet = "RANKING", cols = 4, widths = 33)

setColWidths(wb_st4, sheet = "RANKING", cols = 5, widths = 30)

setColWidths(wb_st4, sheet = "RANKING", cols = 6, widths = 33)

setColWidths(wb_st4, sheet = "RANKING", cols = 7, widths = 36)

setColWidths(wb_st4, sheet = "RANKING", cols = 7:12, widths = 12)

setColWidths(wb_st4, sheet = "RANKING", cols = 13, widths = 13)

setColWidths(wb_st4, sheet = "RANKING", cols = 14, widths = 15)

setColWidths(wb_st4, sheet = "RANKING", cols = 15, widths = 17)

setColWidths(wb_st4, sheet = "RANKING", cols = 16:17, widths = 15)

setColWidths(wb_st4, sheet = "RANKING", cols = 18:36, widths = 18)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb_st4, sheet = "RANKING", style = estiloNumerico, cols = c(8:12,14,16:36), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st4, sheet = "RANKING", style = estiloPorcentagem, cols = 13, rows = 1:3000, gridExpand = TRUE)

## cond format class

conditionalFormatting(wb_st4, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st4, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st4, sheet = "RANKING", cols = 13, rows = 1:1000, rule = 'M1 < 0', style = queda2)

conditionalFormatting(wb_st4, sheet = "RANKING", cols = 14, rows = 1:1000, rule = 'N1 < 0', style = queda2)


writeDataTable(wb_st4, "RANKING", x = data %>% filter(SETOR=="SETOR 4 - JOINVILLE - NORTE"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## RANKING GRUPOS ===========================


addWorksheet(wb_st4, "RANKING GRUPOS")

writeData(wb_st4, sheet = "RANKING GRUPOS", x = gdata)

setColWidths(wb_st4, sheet = "RANKING GRUPOS", cols = 1, widths = 12)

setColWidths(wb_st4, sheet = "RANKING GRUPOS", cols = 2, widths = 35)

setColWidths(wb_st4, sheet = "RANKING GRUPOS", cols = 3, widths = 35)

setColWidths(wb_st4, sheet = "RANKING GRUPOS", cols = 4:9, widths = 15)

setColWidths(wb_st4, sheet = "RANKING GRUPOS", cols = 10, widths = 13)

setColWidths(wb_st4, sheet = "RANKING GRUPOS", cols = 11, widths = 17)

setColWidths(wb_st4, sheet = "RANKING GRUPOS", cols = 12:13, widths = 15)

setColWidths(wb_st4, sheet = "RANKING GRUPOS", cols = 14:31, widths = 17)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")


addStyle(wb_st4, sheet = "RANKING GRUPOS", style = estiloNumerico, cols = c(4:8,10,12:30), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st4, sheet = "RANKING GRUPOS", style = estiloPorcentagem, cols = 9, rows = 1:3000, gridExpand = TRUE)


conditionalFormatting(wb_st4, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st4, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st4, sheet = "RANKING GRUPOS", cols = 9, rows = 1:1000, rule = 'I1 < 0', style = queda2)

conditionalFormatting(wb_st4, sheet = "RANKING GRUPOS", cols = 10, rows = 1:1000, rule = 'J1 < 0', style = queda2)

writeDataTable(wb_st4, "RANKING GRUPOS", gdata %>% filter(SETOR=="SETOR 4 - JOINVILLE - NORTE"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## ESPERADO ===========================================================


## SETORES

addWorksheet(wb_st4, "SETORES ESPERADO")

writeData(wb_st4, sheet = "SETORES ESPERADO", x = sales_sm_2_3)

writeDataTable(wb_st4, "SETORES ESPERADO", sales_sm_2_3 %>% filter(SETOR=="SETOR 4 - JOINVILLE - NORTE"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

setColWidths(wb_st4, sheet = "SETORES ESPERADO", cols = 1, widths = 50)

setColWidths(wb_st4, sheet = "SETORES ESPERADO", cols = 2, widths = 30)

setColWidths(wb_st4, sheet = "SETORES ESPERADO", cols = c(3,4,5), widths = 15)

setColWidths(wb_st4, sheet = "SETORES ESPERADO", cols = 6, widths =30)

setColWidths(wb_st4, sheet = "SETORES ESPERADO", cols = 7, widths =30)

setColWidths(wb_st4, sheet = "SETORES ESPERADO", cols = c(8:11), widths = 15)

addStyle(wb_st4, sheet = "SETORES ESPERADO", style = estiloNumerico, cols = c(3,4,5,7,8,9,10,11), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st4, sheet = "SETORES ESPERADO", style = estiloPorcentagem, cols = 6, rows = 1:3000, gridExpand = TRUE)

## cond format class

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")



conditionalFormatting(wb_st4, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st4, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st4, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2>0', style = crescimento)

conditionalFormatting(wb_st4, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2<0', style = queda)



saveWorkbook(wb_st4, file = "RANKING_CLIENTES_SETOR4.xlsx", overwrite = TRUE)

## SETOR 5 ===========================================================================


wb_st5 <- createWorkbook()


addWorksheet(wb_st5, "RANKING")

writeData(wb_st5, sheet = "RANKING", x = data %>% filter(SETOR=="SETOR 5 - BLUMENAU - VALE"))

setColWidths(wb_st5, sheet = "RANKING", cols = 1, widths = 12)

setColWidths(wb_st5, sheet = "RANKING", cols = 2, widths = 40)

setColWidths(wb_st5, sheet = "RANKING", cols = 3, widths = 12)

setColWidths(wb_st5, sheet = "RANKING", cols = 4, widths = 33)

setColWidths(wb_st5, sheet = "RANKING", cols = 5, widths = 30)

setColWidths(wb_st5, sheet = "RANKING", cols = 6, widths = 33)

setColWidths(wb_st5, sheet = "RANKING", cols = 7, widths = 36)

setColWidths(wb_st5, sheet = "RANKING", cols = 7:12, widths = 12)

setColWidths(wb_st5, sheet = "RANKING", cols = 13, widths = 13)

setColWidths(wb_st5, sheet = "RANKING", cols = 14, widths = 15)

setColWidths(wb_st5, sheet = "RANKING", cols = 15, widths = 17)

setColWidths(wb_st5, sheet = "RANKING", cols = 16:17, widths = 15)

setColWidths(wb_st5, sheet = "RANKING", cols = 18:36, widths = 18)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb_st5, sheet = "RANKING", style = estiloNumerico, cols = c(8:12,14,16:36), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st5, sheet = "RANKING", style = estiloPorcentagem, cols = 13, rows = 1:3000, gridExpand = TRUE)

## cond format class

conditionalFormatting(wb_st5, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st5, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st5, sheet = "RANKING", cols = 13, rows = 1:1000, rule = 'M1 < 0', style = queda2)

conditionalFormatting(wb_st5, sheet = "RANKING", cols = 14, rows = 1:1000, rule = 'N1 < 0', style = queda2)


writeDataTable(wb_st5, "RANKING", x = data %>% filter(SETOR=="SETOR 5 - BLUMENAU - VALE"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## RANKING GRUPOS ===========================


addWorksheet(wb_st5, "RANKING GRUPOS")

writeData(wb_st5, sheet = "RANKING GRUPOS", x = gdata)

setColWidths(wb_st5, sheet = "RANKING GRUPOS", cols = 1, widths = 12)

setColWidths(wb_st5, sheet = "RANKING GRUPOS", cols = 2, widths = 35)

setColWidths(wb_st5, sheet = "RANKING GRUPOS", cols = 3, widths = 35)

setColWidths(wb_st5, sheet = "RANKING GRUPOS", cols = 4:9, widths = 15)

setColWidths(wb_st5, sheet = "RANKING GRUPOS", cols = 10, widths = 13)

setColWidths(wb_st5, sheet = "RANKING GRUPOS", cols = 11, widths = 17)

setColWidths(wb_st5, sheet = "RANKING GRUPOS", cols = 12:13, widths = 15)

setColWidths(wb_st5, sheet = "RANKING GRUPOS", cols = 14:31, widths = 17)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")


addStyle(wb_st5, sheet = "RANKING GRUPOS", style = estiloNumerico, cols = c(4:8,10,12:30), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st5, sheet = "RANKING GRUPOS", style = estiloPorcentagem, cols = 9, rows = 1:3000, gridExpand = TRUE)


conditionalFormatting(wb_st5, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st5, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st5, sheet = "RANKING GRUPOS", cols = 9, rows = 1:1000, rule = 'I1 < 0', style = queda2)

conditionalFormatting(wb_st5, sheet = "RANKING GRUPOS", cols = 10, rows = 1:1000, rule = 'J1 < 0', style = queda2)

writeDataTable(wb_st5, "RANKING GRUPOS", gdata %>% filter(SETOR=="SETOR 5 - BLUMENAU - VALE"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## ESPERADO ===========================================================


## SETORES

addWorksheet(wb_st5, "SETORES ESPERADO")

writeData(wb_st5, sheet = "SETORES ESPERADO", x = sales_sm_2_3)

writeDataTable(wb_st5, "SETORES ESPERADO", sales_sm_2_3 %>% filter(SETOR=="SETOR 5 - BLUMENAU - VALE"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

setColWidths(wb_st5, sheet = "SETORES ESPERADO", cols = 1, widths = 50)

setColWidths(wb_st5, sheet = "SETORES ESPERADO", cols = 2, widths = 30)

setColWidths(wb_st5, sheet = "SETORES ESPERADO", cols = c(3,4,5), widths = 15)

setColWidths(wb_st5, sheet = "SETORES ESPERADO", cols = 6, widths =30)

setColWidths(wb_st5, sheet = "SETORES ESPERADO", cols = 7, widths =30)

setColWidths(wb_st5, sheet = "SETORES ESPERADO", cols = c(8:11), widths = 15)

addStyle(wb_st5, sheet = "SETORES ESPERADO", style = estiloNumerico, cols = c(3,4,5,7,8,9,10,11), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st5, sheet = "SETORES ESPERADO", style = estiloPorcentagem, cols = 6, rows = 1:3000, gridExpand = TRUE)

## cond format class

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")



conditionalFormatting(wb_st5, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st5, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st5, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2>0', style = crescimento)

conditionalFormatting(wb_st5, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2<0', style = queda)



saveWorkbook(wb_st5, file = "RANKING_CLIENTES_SETOR5.xlsx", overwrite = TRUE)

## SETOR 6 ===========================================================================


wb_st6 <- createWorkbook()


addWorksheet(wb_st6, "RANKING")

writeData(wb_st6, sheet = "RANKING", x = data %>% filter(SETOR=="SETOR 7 - FLORIANOPOLIS  LOJAS"))

setColWidths(wb_st6, sheet = "RANKING", cols = 1, widths = 12)

setColWidths(wb_st6, sheet = "RANKING", cols = 2, widths = 40)

setColWidths(wb_st6, sheet = "RANKING", cols = 3, widths = 12)

setColWidths(wb_st6, sheet = "RANKING", cols = 4, widths = 33)

setColWidths(wb_st6, sheet = "RANKING", cols = 5, widths = 30)

setColWidths(wb_st6, sheet = "RANKING", cols = 6, widths = 33)

setColWidths(wb_st6, sheet = "RANKING", cols = 7, widths = 36)

setColWidths(wb_st6, sheet = "RANKING", cols = 7:12, widths = 12)

setColWidths(wb_st6, sheet = "RANKING", cols = 13, widths = 13)

setColWidths(wb_st6, sheet = "RANKING", cols = 14, widths = 15)

setColWidths(wb_st6, sheet = "RANKING", cols = 15, widths = 17)

setColWidths(wb_st6, sheet = "RANKING", cols = 16:17, widths = 15)

setColWidths(wb_st6, sheet = "RANKING", cols = 18:36, widths = 18)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb_st6, sheet = "RANKING", style = estiloNumerico, cols = c(8:12,14,16:36), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st6, sheet = "RANKING", style = estiloPorcentagem, cols = 13, rows = 1:3000, gridExpand = TRUE)

## cond format class

conditionalFormatting(wb_st6, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st6, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st6, sheet = "RANKING", cols = 13, rows = 1:1000, rule = 'M1 < 0', style = queda2)

conditionalFormatting(wb_st6, sheet = "RANKING", cols = 14, rows = 1:1000, rule = 'N1 < 0', style = queda2)


writeDataTable(wb_st6, "RANKING", x = data %>% filter(SETOR=="SETOR 7 - FLORIANOPOLIS  LOJAS"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## RANKING GRUPOS ===========================


addWorksheet(wb_st6, "RANKING GRUPOS")

writeData(wb_st6, sheet = "RANKING GRUPOS", x = gdata)

setColWidths(wb_st6, sheet = "RANKING GRUPOS", cols = 1, widths = 12)

setColWidths(wb_st6, sheet = "RANKING GRUPOS", cols = 2, widths = 35)

setColWidths(wb_st6, sheet = "RANKING GRUPOS", cols = 3, widths = 35)

setColWidths(wb_st6, sheet = "RANKING GRUPOS", cols = 4:9, widths = 15)

setColWidths(wb_st6, sheet = "RANKING GRUPOS", cols = 10, widths = 13)

setColWidths(wb_st6, sheet = "RANKING GRUPOS", cols = 11, widths = 17)

setColWidths(wb_st6, sheet = "RANKING GRUPOS", cols = 12:13, widths = 15)

setColWidths(wb_st6, sheet = "RANKING GRUPOS", cols = 14:31, widths = 17)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")


addStyle(wb_st6, sheet = "RANKING GRUPOS", style = estiloNumerico, cols = c(4:8,10,12:30), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st6, sheet = "RANKING GRUPOS", style = estiloPorcentagem, cols = 9, rows = 1:3000, gridExpand = TRUE)


conditionalFormatting(wb_st6, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st6, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st6, sheet = "RANKING GRUPOS", cols = 9, rows = 1:1000, rule = 'I1 < 0', style = queda2)

conditionalFormatting(wb_st6, sheet = "RANKING GRUPOS", cols = 10, rows = 1:1000, rule = 'J1 < 0', style = queda2)

writeDataTable(wb_st6, "RANKING GRUPOS", gdata %>% filter(SETOR=="SETOR 7 - FLORIANOPOLIS  LOJAS"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## ESPERADO ===========================================================


## SETORES

addWorksheet(wb_st6, "SETORES ESPERADO")

writeData(wb_st6, sheet = "SETORES ESPERADO", x = sales_sm_2_3)

writeDataTable(wb_st6, "SETORES ESPERADO", sales_sm_2_3 %>% filter(SETOR=="SETOR 7 - FLORIANOPOLIS  LOJAS"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

setColWidths(wb_st6, sheet = "SETORES ESPERADO", cols = 1, widths = 50)

setColWidths(wb_st6, sheet = "SETORES ESPERADO", cols = 2, widths = 30)

setColWidths(wb_st6, sheet = "SETORES ESPERADO", cols = c(3,4,5), widths = 15)

setColWidths(wb_st6, sheet = "SETORES ESPERADO", cols = 6, widths =30)

setColWidths(wb_st6, sheet = "SETORES ESPERADO", cols = 7, widths =30)

setColWidths(wb_st6, sheet = "SETORES ESPERADO", cols = c(8:11), widths = 15)

addStyle(wb_st6, sheet = "SETORES ESPERADO", style = estiloNumerico, cols = c(3,4,5,7,8,9,10,11), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st6, sheet = "SETORES ESPERADO", style = estiloPorcentagem, cols = 6, rows = 1:3000, gridExpand = TRUE)

## cond format class

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")



conditionalFormatting(wb_st6, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st6, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st6, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2>0', style = crescimento)

conditionalFormatting(wb_st6, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2<0', style = queda)



saveWorkbook(wb_st6, file = "RANKING_CLIENTES_SETOR6.xlsx", overwrite = TRUE)

## SETOR 7 ===========================================================================


wb_st7 <- createWorkbook()


addWorksheet(wb_st7, "RANKING")

writeData(wb_st7, sheet = "RANKING", x = data %>% filter(SETOR=="SETOR 7 - FLORIANOPOLIS  LOJAS"))

setColWidths(wb_st7, sheet = "RANKING", cols = 1, widths = 12)

setColWidths(wb_st7, sheet = "RANKING", cols = 2, widths = 40)

setColWidths(wb_st7, sheet = "RANKING", cols = 3, widths = 12)

setColWidths(wb_st7, sheet = "RANKING", cols = 4, widths = 33)

setColWidths(wb_st7, sheet = "RANKING", cols = 5, widths = 30)

setColWidths(wb_st7, sheet = "RANKING", cols = 6, widths = 33)

setColWidths(wb_st7, sheet = "RANKING", cols = 7, widths = 36)

setColWidths(wb_st7, sheet = "RANKING", cols = 7:12, widths = 12)

setColWidths(wb_st7, sheet = "RANKING", cols = 13, widths = 13)

setColWidths(wb_st7, sheet = "RANKING", cols = 14, widths = 15)

setColWidths(wb_st7, sheet = "RANKING", cols = 15, widths = 17)

setColWidths(wb_st7, sheet = "RANKING", cols = 16:17, widths = 15)

setColWidths(wb_st7, sheet = "RANKING", cols = 18:36, widths = 18)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")

## format numbers

estiloNumerico <- createStyle(numFmt = "#,##0")

estiloPorcentagem <- createStyle(numFmt = "0%")

addStyle(wb_st7, sheet = "RANKING", style = estiloNumerico, cols = c(8:12,14,16:36), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st7, sheet = "RANKING", style = estiloPorcentagem, cols = 13, rows = 1:3000, gridExpand = TRUE)

## cond format class

conditionalFormatting(wb_st7, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st7, sheet = "RANKING", cols = 15, rows = 1:1000, rule = 'O1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st7, sheet = "RANKING", cols = 13, rows = 1:1000, rule = 'M1 < 0', style = queda2)

conditionalFormatting(wb_st7, sheet = "RANKING", cols = 14, rows = 1:1000, rule = 'N1 < 0', style = queda2)


writeDataTable(wb_st7, "RANKING", x = data %>% filter(SETOR=="SETOR 7 - FLORIANOPOLIS  LOJAS"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## RANKING GRUPOS ===========================


addWorksheet(wb_st7, "RANKING GRUPOS")

writeData(wb_st7, sheet = "RANKING GRUPOS", x = gdata)

setColWidths(wb_st7, sheet = "RANKING GRUPOS", cols = 1, widths = 12)

setColWidths(wb_st7, sheet = "RANKING GRUPOS", cols = 2, widths = 35)

setColWidths(wb_st7, sheet = "RANKING GRUPOS", cols = 3, widths = 35)

setColWidths(wb_st7, sheet = "RANKING GRUPOS", cols = 4:9, widths = 15)

setColWidths(wb_st7, sheet = "RANKING GRUPOS", cols = 10, widths = 13)

setColWidths(wb_st7, sheet = "RANKING GRUPOS", cols = 11, widths = 17)

setColWidths(wb_st7, sheet = "RANKING GRUPOS", cols = 12:13, widths = 15)

setColWidths(wb_st7, sheet = "RANKING GRUPOS", cols = 14:31, widths = 17)

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")

queda2 <- createStyle(fontColour = "#FF0800")


addStyle(wb_st7, sheet = "RANKING GRUPOS", style = estiloNumerico, cols = c(4:8,10,12:30), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st7, sheet = "RANKING GRUPOS", style = estiloPorcentagem, cols = 9, rows = 1:3000, gridExpand = TRUE)


conditionalFormatting(wb_st7, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "CRESCIMENTO"', style = crescimento)

conditionalFormatting(wb_st7, sheet = "RANKING GRUPOS", cols = 11, rows = 1:1000, rule = 'K1 == "QUEDA"', style = queda)

conditionalFormatting(wb_st7, sheet = "RANKING GRUPOS", cols = 9, rows = 1:1000, rule = 'I1 < 0', style = queda2)

conditionalFormatting(wb_st7, sheet = "RANKING GRUPOS", cols = 10, rows = 1:1000, rule = 'J1 < 0', style = queda2)

writeDataTable(wb_st7, "RANKING GRUPOS", gdata %>% filter(SETOR=="SETOR 7 - FLORIANOPOLIS  LOJAS"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)


## ESPERADO ===========================================================


## SETORES

addWorksheet(wb_st7, "SETORES ESPERADO")

writeData(wb_st7, sheet = "SETORES ESPERADO", x = sales_sm_2_3)

writeDataTable(wb_st7, "SETORES ESPERADO", sales_sm_2_3 %>% filter(SETOR=="SETOR 7 - FLORIANOPOLIS  LOJAS"), startCol = 1, startRow = 1, xy = NULL, colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2", tableName = NULL, headerStyle = NULL, withFilter = TRUE, keepNA = FALSE, na.string = NULL, sep = ", ", stack = FALSE, firstColumn = FALSE, lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE)

setColWidths(wb_st7, sheet = "SETORES ESPERADO", cols = 1, widths = 50)

setColWidths(wb_st7, sheet = "SETORES ESPERADO", cols = 2, widths = 30)

setColWidths(wb_st7, sheet = "SETORES ESPERADO", cols = c(3,4,5), widths = 15)

setColWidths(wb_st7, sheet = "SETORES ESPERADO", cols = 6, widths =30)

setColWidths(wb_st7, sheet = "SETORES ESPERADO", cols = 7, widths =30)

setColWidths(wb_st7, sheet = "SETORES ESPERADO", cols = c(8:11), widths = 15)

addStyle(wb_st7, sheet = "SETORES ESPERADO", style = estiloNumerico, cols = c(3,4,5,7,8,9,10,11), rows = 1:3000, gridExpand = TRUE)

addStyle(wb_st7, sheet = "SETORES ESPERADO", style = estiloPorcentagem, cols = 6, rows = 1:3000, gridExpand = TRUE)

## cond format class

crescimento <- createStyle(fontColour = "#02862a", bgFill = "#ccf2d8")

queda <- createStyle(fontColour = "#7b1e1e", bgFill = "#e69999")



conditionalFormatting(wb_st7, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st7, sheet = "SETORES ESPERADO", cols = 6, rows = 1:1000, rule = 'F1 < 0', style = queda2)

conditionalFormatting(wb_st7, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2>0', style = crescimento)

conditionalFormatting(wb_st7, sheet = "SETORES ESPERADO", cols = 7, rows = 2:1000, rule = 'G2<0', style = queda)



saveWorkbook(wb_st7, file = "RANKING_CLIENTES_SETOR7.xlsx", overwrite = TRUE)







