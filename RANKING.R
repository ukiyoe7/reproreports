library(tidyverse)
library(lubridate)
library(reshape2)
library(DBI)
library(googlesheets4)
con2 <- dbConnect(odbc::odbc(), "reproreplica")

sales2021 <- get(load("C:\\Users\\Repro\\Documents\\R\\ADM\\REPORT\\BASES\\sales2021.RData"))

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
WHERE CLICLIENTE='S' AND ENDFAT='S'
AND ENDCLI.ZOCODIGO IN (20,21,22,23,24,25,28)
")

inativos <- dbGetQuery(con2,"SELECT DISTINCT SITCLI.CLICODIGO FROM SITCLI
INNER JOIN (SELECT DISTINCT SITCLI.CLICODIGO,MAX(SITDATA)ULTIMA FROM SITCLI
GROUP BY 1)A ON SITCLI.CLICODIGO=A.CLICODIGO AND A.ULTIMA=SITCLI.SITDATA 
INNER JOIN (SELECT DISTINCT SITCLI.CLICODIGO,SITDATA,MAX(SITSEQ)USEQ FROM SITCLI
GROUP BY 1,2)MSEQ ON A.CLICODIGO=MSEQ.CLICODIGO AND MSEQ.SITDATA=A.ULTIMA AND MSEQ.USEQ=SITCLI.SITSEQ
WHERE SITCODIGO=4")


sales2022 <- dbGetQuery(con2,"
WITH CLI AS (SELECT DISTINCT CLIEN.CLICODIGO,CLINOMEFANT
  FROM CLIEN
   WHERE CLICLIENTE='S'),
  
FIS AS (SELECT FISCODIGO FROM TBFIS WHERE FISTPNATOP IN ('V','SR','R')),
  
PED AS (SELECT ID_PEDIDO,PEDDTBAIXA,P.CLICODIGO FROM PEDID P
   INNER JOIN FIS F ON P.FISCODIGO1=F.FISCODIGO
    INNER JOIN CLI C ON P.CLICODIGO=C.CLICODIGO
     WHERE PEDSITPED<>'C' AND PEDDTBAIXA BETWEEN '01.01.2022' AND 'TODAY'),
     
AUX AS (SELECT PROCODIGO,PROTIPO FROM PRODU )     

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
    INNER JOIN AUX ON PD.PROCODIGO= AUX.PROCODIGO
     GROUP BY 1,2,3") 

##union of sales

sales <- union_all(sales2021,sales2022)

##discounts

desct <- dbGetQuery(con2,"
SELECT CLICODIGO,
 TBPCODIGO,
  TBPDESC2,
   CASE 
    WHEN TBPCODIGO=101 THEN 'VARILUX'
    WHEN TBPCODIGO=102 THEN 'VARILUX' 
    WHEN TBPCODIGO=103 THEN 'VARILUX' 
    WHEN TBPCODIGO=104 THEN 'VARILUX' 
    WHEN TBPCODIGO=105 THEN 'VARILUX'
    WHEN TBPCODIGO=202 THEN 'KODAK DIGITAL'
    WHEN TBPCODIGO=309 THEN 'EYEZEN'
    WHEN TBPCODIGO=302 THEN 'IMAGEM DIGITAL'
    WHEN TBPCODIGO=303 THEN 'UZ+'
    WHEN TBPCODIGO=304 THEN 'ACTUALITE'
    WHEN TBPCODIGO=305 THEN 'AVANCE'
    WHEN TBPCODIGO=308 THEN 'INSIGNE'
    WHEN TBPCODIGO=312 THEN 'ADVANS'
   ELSE '' END LINHA
  FROM CLITBP
  WHERE TBPCODIGO IN (101,102,103,104,105,202,309,302,303,304,305,308,312)
  ") 

mdesc <- desct %>% dcast(.,CLICODIGO ~ LINHA,value.var = "TBPDESC2",mean)%>% as.data.frame()  %>% replace(.,is.na(.),0)

gdesct <- dbGetQuery(con2,"
SELECT C.CLICODIGO,
CODGRUPO,
 TBPCODIGO,
  TBPDESC2,
   CASE 
    WHEN TBPCODIGO=101 THEN 'VARILUX'
    WHEN TBPCODIGO=102 THEN 'VARILUX' 
    WHEN TBPCODIGO=103 THEN 'VARILUX' 
    WHEN TBPCODIGO=104 THEN 'VARILUX' 
    WHEN TBPCODIGO=105 THEN 'VARILUX'
    WHEN TBPCODIGO=202 THEN 'KODAK DIGITAL'
    WHEN TBPCODIGO=309 THEN 'EYEZEN'
    WHEN TBPCODIGO=302 THEN 'IMAGEM DIGITAL'
    WHEN TBPCODIGO=303 THEN 'UZ+'
    WHEN TBPCODIGO=304 THEN 'ACTUALITE'
    WHEN TBPCODIGO=305 THEN 'AVANCE'
    WHEN TBPCODIGO=308 THEN 'INSIGNE'
    WHEN TBPCODIGO=312 THEN 'ADVANS'
   ELSE '' END LINHA
  FROM CLITBP C
  LEFT JOIN (SELECT CLICODIGO,GCLCODIGO CODGRUPO
  FROM CLIEN WHERE CLICLIENTE='S')A 
  ON C.CLICODIGO=A.CLICODIGO
  WHERE TBPCODIGO IN (101,102,103,104,105,202,309,302,303,304,305,308,312)
  ") %>% filter(CODGRUPO!='') %>% group_by(CODGRUPO,LINHA) %>% summarise(m=mean(TBPDESC2)) %>% 
  dcast(.,CODGRUPO ~ LINHA,mean) %>% replace(is.na(.),0) %>% apply(.,2,function(x) round(x,0)) %>% as.data.frame()


nsales <- sales %>% group_by(CLICODIGO) %>% 
  summarize(
    LASTMONTHLASTYEAR=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date((Sys.Date()-years(1)) %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    LASTMONTHTHISYEAR=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    CURRENTMONTH=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date(),'month') %m-% days(1)],na.rm = TRUE),
    
    YTD21=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date()-years(1), "month")-1],na.rm = TRUE),
    
    YTD22=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "month")-1],na.rm = TRUE),
    
    MEDIA21=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date()-years(1), "month")-1]/as.numeric(length(seq(floor_date(Sys.Date(),"year"),floor_date(Sys.Date(),"month"),by="month"))-1)),
    
    MEDIA22=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") < floor_date(Sys.Date(), "month")]/as.numeric(length(seq(floor_date(Sys.Date(),"year"),floor_date(Sys.Date(),"month"),by="month"))-1)),
    
    PAST12=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") >  ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)])
  ) %>% mutate(VAR2022=ifelse(is.finite(YTD22/YTD21-1),YTD22/YTD21-1,0))

nsales <- apply(nsales,2,function(x) round(x,2)) %>% as.data.frame()

data <- left_join(anti_join(clientes,inativos,by="CLICODIGO"),nsales,by="CLICODIGO") %>% 
  left_join(.,mdesc,by="CLICODIGO") %>% as.data.frame()

data <- data %>%  mutate(STATUS=case_when(
  DATACADASTRO>=floor_date(floor_date(Sys.Date() %m-% months(1), 'month')-years(1), "month") ~ 'CLIENTE NOVO',
  SEMRECEITA==0 ~ 'SEM RECEITA',
  VAR2022>=0 ~ 'CRESCIMENTO',
  VAR2022<0 ~ 'QUEDA',
  PAST12==0  ~ 'PERDIDO',
  YTD21==0 & YTD22>0 ~ 'RECUPERADO',
  TRUE ~ ''
))

LASTMONTHLASTYEAR1 <- toupper(format(floor_date(Sys.Date()-years(1), "month")-1,"%b%/%Y"))

LASTMONTHTHISYEAR1 <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH1 <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))

data <- data %>% arrange(desc(.$YTD22)) %>% as.data.frame() %>% 
  rename_at(9:11,~ c(LASTMONTHLASTYEAR1,LASTMONTHTHISYEAR1,CURRENTMONTH1)) %>% .[,c(1:3,5,6,9:13,18,28,14:15,27,25,19:24,26,7)]



##GRUPOS


dt <- left_join(anti_join(clientes,inativos,by="CLICODIGO"),sales,by="CLICODIGO") %>% as.data.frame()

gsales <- dt %>% group_by(SETOR,CODGRUPO,GRUPO) %>% 
  summarize(
    LASTMONTHLASTYEAR2=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date((Sys.Date()-years(1)) %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    LASTMONTHTHISYEAR2=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date() %m-% months(1), 'month') & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date() %m-% months(1), 'month') %m-% days(1)],na.rm = TRUE),
    
    CURRENTMONTH2=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>=floor_date(Sys.Date(), "month") & floor_date(PEDDTBAIXA,"day")<=ceiling_date(Sys.Date(),'month') %m-% days(1)],na.rm = TRUE),
    
    YTD21=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date()-years(1), "month")-1],na.rm = TRUE),
    
    YTD22=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") >= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") <= floor_date(Sys.Date(), "month")-1],na.rm = TRUE),
    
    MEDIA21=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date()-years(1), "year") & floor_date(PEDDTBAIXA,"day") < floor_date(Sys.Date(), "year")-1]/as.numeric(length(seq(floor_date(Sys.Date(),"year"),floor_date(Sys.Date(),"month"),by="month"))-1),na.rm = TRUE),
    
    MEDIA22=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")>= floor_date(Sys.Date(), "year") & floor_date(PEDDTBAIXA,"day") < floor_date(Sys.Date(), "month")]/as.numeric(length(seq(floor_date(Sys.Date(),"year"),floor_date(Sys.Date(),"month"),by="month"))-1),na.rm = TRUE),
    
    PAST12=sum(VRVENDA[floor_date(PEDDTBAIXA,"day")< Sys.Date()],na.rm = TRUE),
    
    SEMRECEITA=sum(VRVENDA[floor_date(PEDDTBAIXA,"day") > ceiling_date((Sys.Date()-years(1)) %m-% months(1), 'month') %m-% days(1) ],na.rm = TRUE),
    
    DESCTO_GERAL=mean(DESCTGERAL,na.omit=TRUE)
    
  ) %>% mutate(VAR2022=ifelse(is.finite(YTD22/YTD21-1),YTD22/YTD21-1,0))

gsales <- gsales %>%  mutate(STATUS=case_when(
  SEMRECEITA==0 ~ 'SEM RECEITA',
  VAR2022>0 ~ 'CRESCIMENTO',
  VAR2022<0 ~ 'QUEDA',
  PAST12==0  ~ 'PERDIDO',
  YTD20==0 & YTD21>0 ~ 'RECUPERADO',
  TRUE ~ ''
))

gdata <-apply(gsales[,4:14],2,function(x) round(x,2)) %>% as.data.frame() %>% 
  cbind(gsales[,1:3],.) %>% cbind(.,gsales[,15])


gdata <- gdata %>% left_join(.,gdesct,by="CODGRUPO") %>% as.data.frame()

gdata <- gdata %>% filter(CODGRUPO!='')

LASTMONTHLASTYEAR3 <- toupper(format(floor_date(Sys.Date()-years(1), "month")-1,"%b%/%Y"))

LASTMONTHTHISYEAR3 <- toupper(format(floor_date(Sys.Date(), "month")-1,"%b%/%Y"))

CURRENTMONTH3 <- toupper(format(floor_date(Sys.Date(), "month"),"%b%/%Y"))

gdata <- gdata %>% arrange(desc(.$YTD22)) %>% as.data.frame() %>% 
  rename_at(4:6,~ c(LASTMONTHLASTYEAR3,LASTMONTHTHISYEAR3,CURRENTMONTH3)) %>% 
  .[,c(3,2,1,4:8,14,15,9:10,23,21,16:20,22,13)]

sheet_write(data,ss = "1GpUPX7RQWL-TDrujKNhDYrKSZ5VzmumZPE8a3TXwaek",sheet = "DADOS")
sheet_write(gdata,ss = "1GpUPX7RQWL-TDrujKNhDYrKSZ5VzmumZPE8a3TXwaek",sheet = "DADOS2") 




