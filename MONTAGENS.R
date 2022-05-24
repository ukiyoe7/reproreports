## RELATÓRIO DE MONTAGENS
## 05.2022
## SANDRO JAKOSKA


## LIBRARIES =======================================================================

library(dplyr)
library(lubridate)
library(DBI)
library(googlesheets4)
library(reshape2)
library(gmailr)

## DB CONNECTION ====================================================================

con2 <- dbConnect(odbc::odbc(), "reproreplica")

## SQL  ====================================================================


montagens_2022 <- dbGetQuery(con2,"
WITH CLI AS (SELECT DISTINCT C.CLICODIGO,CLINOMEFANT,
IIF(C.GCLCODIGO IS NULL,C.CLICODIGO||' '||CLINOMEFANT,'G'||C.GCLCODIGO||' '||GCLNOME) CLIENTE,
SETOR
FROM CLIEN C
LEFT JOIN GRUPOCLI GC ON  C.GCLCODIGO=GC.GCLCODIGO
LEFT JOIN (SELECT CLICODIGO,E.ZOCODIGO,ZODESCRICAO SETOR FROM ENDCLI E
LEFT JOIN (SELECT ZOCODIGO,ZODESCRICAO FROM ZONA WHERE ZOCODIGO IN (20,21,22,23,24,25,28)
)Z ON E.ZOCODIGO=Z.ZOCODIGO WHERE ENDFAT='S')A ON C.CLICODIGO=A.CLICODIGO
 WHERE CLICLIENTE='S'),
FIS AS (SELECT FISCODIGO FROM TBFIS WHERE FISTPNATOP IN ('V','SR','R')),
PED AS (SELECT ID_PEDIDO,PEDDTBAIXA,CLIENTE,SETOR FROM PEDID P
 INNER JOIN FIS F ON P.FISCODIGO1=F.FISCODIGO
  INNER JOIN CLI C ON P.CLICODIGO=C.CLICODIGO
   WHERE PEDSITPED<>'C' AND PEDDTBAIXA BETWEEN '01/01/2022' AND 'TODAY'),
   
AUX AS (SELECT PROCODIGO,PROTIPO FROM PRODU
WHERE PROTIPO IN ('M'))     
SELECT 
PEDDTBAIXA,
CLIENTE,
SETOR,
PD.PROCODIGO MONT,
PDPDESCRICAO,
SUM(PDPQTDADE) QTD,
SUM(PDPUNITLIQUIDO*PDPQTDADE)VRVENDA
FROM PDPRD PD
 INNER JOIN PED P ON PD.ID_PEDIDO=P.ID_PEDIDO
  INNER JOIN AUX A ON PD.PROCODIGO= A.PROCODIGO
   GROUP BY 1,2,3,4,5")


## RESHAPE DATA ====================================================================

vmontagens_2022  <- montagens_2022 %>% 
  group_by(CLIENTE,SETOR,MONT,MES=toupper(format(floor_date(PEDDTBAIXA,"month"),"%b"))) %>% 
  summarize(VLR=round(sum(VRVENDA)/sum(QTD),1)) %>% 
  dcast(CLIENTE + SETOR + MONT ~ factor(MES, levels =c("JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO",
                                                       "SET","OUT","NOV","DEZ"))) 

qmontagens_2022  <- montagens_2022 %>% 
  group_by(CLIENTE,SETOR,MONT,MES=toupper(format(floor_date(PEDDTBAIXA,"month"),"%b"))) %>% 
  summarize(QTD=round(sum(QTD),1)) %>% 
  dcast(CLIENTE + SETOR + MONT ~ factor(MES, levels =c("JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO",
                                                       "SET","OUT","NOV","DEZ"))) 

## WRIET ON GOOGLE ====================================================================

m1 <- vmontagens_2022 %>%  filter(SETOR=='SETOR 1 - GRANDE FLORIANOPOLIS') %>% mutate(SETOR='SETOR 1') 
range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 1",data = m1, range = "A2",reformat = FALSE)

m1_2 <- left_join(vmontagens_2022 %>% 
                    filter(SETOR=='SETOR 1 - GRANDE FLORIANOPOLIS') %>% 
                    select(CLIENTE,SETOR,MONT),qmontagens_2022,by=c("CLIENTE","SETOR","MONT")) %>% .[,c(-1,-2,-3)]

range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 1",data = m1_2, range = "O2",reformat = FALSE)

## SETOR2

m2 <- vmontagens_2022 %>%  filter(SETOR=='SETOR 2 - CRICIUMA - SUL') %>% mutate(SETOR='SETOR 2') 

range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 2",data = m2, range = "A2",reformat = FALSE)

m2_2 <- left_join(vmontagens_2022 %>% 
                    filter(SETOR=='SETOR 2 - CRICIUMA - SUL') %>% 
                    select(CLIENTE,SETOR,MONT),qmontagens_2022,by=c("CLIENTE","SETOR","MONT")) %>% .[,c(-1,-2,-3)]

range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 2",data = m2_2, range = "O2",reformat = FALSE)


## SETOR3

m3 <- vmontagens_2022 %>%  filter(SETOR=='SETOR 3 - CHAPECO-PLANAL-OESTE') %>% mutate(SETOR='SETOR 3') 

range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 3",data = m3, range = "A2",reformat = FALSE)

m3_2 <- left_join(vmontagens_2022 %>% 
                    filter(SETOR=='SETOR 3 - CHAPECO-PLANAL-OESTE') %>% 
                    select(CLIENTE,SETOR,MONT),qmontagens_2022,by=c("CLIENTE","SETOR","MONT")) %>% .[,c(-1,-2,-3)]

range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 3",data = m3_2, range = "O2",reformat = FALSE)


## SETOR4

m4 <- vmontagens_2022 %>%  filter(SETOR=='SETOR 4 - JOINVILLE - NORTESC') %>% mutate(SETOR='SETOR 4') 

range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 4",data = m4, range = "A2",reformat = FALSE)

m4_2 <- left_join(vmontagens_2022 %>% 
                    filter(SETOR=='SETOR 4 - JOINVILLE - NORTESC') %>% 
                    select(CLIENTE,SETOR,MONT),qmontagens_2022,by=c("CLIENTE","SETOR","MONT")) %>% .[,c(-1,-2,-3)]

range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 4",data = m4_2, range = "O2",reformat = FALSE)


## SETOR5

m5 <- vmontagens_2022 %>%  filter(SETOR=='SETOR 5 - BLUMENAU - VALE') %>% mutate(SETOR='SETOR 5') 

range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 5",data = m5, range = "A2",reformat = FALSE)

m5_2 <- left_join(vmontagens_2022 %>% 
                    filter(SETOR=='SETOR 5 - BLUMENAU - VALE') %>% 
                    select(CLIENTE,SETOR,MONT),qmontagens_2022,by=c("CLIENTE","SETOR","MONT")) %>% .[,c(-1,-2,-3)]

range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 5",data = m5_2, range = "O2",reformat = FALSE)

## SETOR6

m6 <- vmontagens_2022 %>%  filter(SETOR=='SETOR 6 - BALNEARIO-LITORAL') %>% mutate(SETOR='SETOR 6') 

range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 6",data = m6, range = "A2",reformat = FALSE)

m6_2 <- left_join(vmontagens_2022 %>% 
                    filter(SETOR=='SETOR 6 - BALNEARIO-LITORAL') %>% 
                    select(CLIENTE,SETOR,MONT),qmontagens_2022,by=c("CLIENTE","SETOR","MONT")) %>% .[,c(-1,-2,-3)]

range_write("1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU",
            sheet = "SETOR 6",data = m6_2, range = "O2",reformat = FALSE)


## SEND MAIL ====================================================================


gm_auth_configure(path = "C:\\Users\\Repro\\Documents\\R\\ADM\\REPORT\\sendmail.json")

mymail <- gm_mime() %>% 
  gm_to("sandro.jakoska@repro.com.br,cristiano.regis@repro.com.br") %>% 
  gm_from ("comunicacao@repro.com.br") %>%
  gm_subject("RELATORIO MONTAGENS") %>%
  gm_text_body("RELATÓRIO ATUALIZADO.
ACESSE O LINK https://docs.google.com/spreadsheets/d/1FtDQ9RTFxTveHa3HPTQKzLjzVADnJa-K1UT780nEKdU/edit?usp=sharing.

ESSE É UM EMAIL AUTOMÁTICO.") 
gm_send_message(mymail)




