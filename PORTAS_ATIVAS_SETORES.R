## REPORT PORTAS ATIVAS

library(DBI)
library(dplyr)
library(readr)
library(reshape2)
library(lubridate)
library(openxlsx)
library(ggplot2)
library(reshape2)
library(bindata)

con2 <- dbConnect(odbc::odbc(), "repro",encoding="Latin1")

## HEAD ==================================================

data_atual <- paste0("DATA: ",format(Sys.Date(),"%d/%m/%Y"))

## SQL ========================


portas_ativas_2023 <- get(load("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\BASES\\portas_ativas_2023.RData"))

portas_ativas_2024 <- dbGetQuery(con2, statement = read_file('C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\SQL\\PORTAS_ATIVAS.sql'))

portas_ativas <-
union_all(portas_ativas_2023,portas_ativas_2024)


qtd_pcs_2023 <- get(load("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\BASES\\qtd_pcs_2023.RData"))

qtd_pcs_2024 <- dbGetQuery(con2, statement = read_file('C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\SQL\\QTD_PCS.sql'))

qtd_pcs <-
union_all(qtd_pcs_2023,qtd_pcs_2024)



qtd_pcs2 <-  
  qtd_pcs %>% 
  group_by(MES = format(floor_date(PEDDTBAIXA, "month"), "%b"), ANO = year(floor_date(PEDDTBAIXA, "year"))) %>% 
  summarize(QTD = sum(QTD, na.rm = TRUE), .groups = "drop") %>% 
  mutate(MES = factor(MES, levels = c("jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"), ordered = TRUE))


qtd_pcs_setores <-  
  qtd_pcs %>% 
  group_by(SETOR,MES = format(floor_date(PEDDTBAIXA, "month"), "%b"), ANO = year(floor_date(PEDDTBAIXA, "year"))) %>% 
  summarize(QTD = sum(QTD, na.rm = TRUE), .groups = "drop") %>% 
  mutate(MES = factor(MES, levels = c("jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"), ordered = TRUE))


## PORTAS ATIVAS ==================================================


## geral

portas_ativas2 <- 
  portas_ativas %>% 
  group_by(MES = format(floor_date(PEDDTBAIXA, "month"), "%b"),ANO=year(floor_date(PEDDTBAIXA,"year"))) %>% 
  summarize(PORTAS_ATIVAS=n_distinct(CLICODIGO),.groups = "drop") %>% 
  mutate(MES=factor(MES, levels = c("jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"), ordered = TRUE))

portas_ativas3 <-
  portas_ativas2 %>%  dcast(MES ~ ANO)

portas_ativas4 <- "PORTAS ATIVAS"

## setores


portas_ativas_setores <- 
  portas_ativas %>% 
  group_by(SETOR,MES = format(floor_date(PEDDTBAIXA, "month"), "%b"),ANO=year(floor_date(PEDDTBAIXA,"year"))) %>% 
  summarize(PORTAS_ATIVAS=n_distinct(CLICODIGO)) %>% 
  mutate(MES=factor(MES, levels = c("jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"), ordered = TRUE))

portas_ativas_setores2 <-
  portas_ativas_setores %>%  dcast(MES ~ ANO)

portas_ativas4 <- "PORTAS ATIVAS"



qtd_pcs3 <-
  qtd_pcs2 %>%  dcast(MES ~ ANO)


qtd_pcs4 <- "QTD PEÇAS"


## GIRO ==================================================

giro <-
  left_join(portas_ativas2,qtd_pcs2,by=c("MES","ANO")) %>% mutate(GIRO=round(QTD/PORTAS_ATIVAS,2))


giro2 <-
  giro %>% select(MES,ANO,GIRO) %>% dcast(MES ~ ANO,fun.aggregate = sum)

giro3 <- "GIRO"



## GIRO SETORES ==================================================

giro_setores <-
  left_join(portas_ativas_setores,qtd_pcs_setores,by=c("SETOR","MES","ANO")) %>% mutate(GIRO=round(QTD/PORTAS_ATIVAS,2))


giro_setores2 <-
  giro_setores %>% select(MES,ANO,GIRO) %>% dcast(MES ~ ANO,fun.aggregate = sum)


## CLIENTES AUSENTES ================================================


# missing 2 months ago compared to 3 months ago
ausentes_penultimo_mes <-
  anti_join(
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date() - months(3), "month"))) %>% 
      distinct(CLICODIGO) 
    ,
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date() - months(2), "month"))) %>% 
      distinct(CLICODIGO)
    ,by="CLICODIGO") 




# missing 1 months ago compared to 2 months ago
ausentes_ultimo_mes <-
  anti_join(
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date() - months(2), "month"))) %>% 
      distinct(CLICODIGO) 
    ,
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date() - months(1), "month"))) %>% 
      distinct(CLICODIGO)
    ,by="CLICODIGO") 


#  missing current month compared to 1 month ago

ausentes_mes_atual <-
  anti_join(
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date() - months(1), "month"))) %>% 
      distinct(CLICODIGO) 
    ,
    portas_ativas %>% 
      filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date(), "month"))) %>% 
      distinct(CLICODIGO)
    ,by="CLICODIGO")


ausentes_mes_atual2 <-  
  anti_join(ausentes_mes_atual,inner_join(ausentes_penultimo_mes,ausentes_mes_atual,by="CLICODIGO"),by="CLICODIGO")


ausentes_penultimo_mes2 <-  
  anti_join(ausentes_penultimo_mes,inner_join(ausentes_penultimo_mes,portas_ativas %>% 
                                                filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date(), "month"))) %>% 
                                                distinct(CLICODIGO),by="CLICODIGO"),by="CLICODIGO")


ausentes_penultimo_mes3 <-  
  anti_join(ausentes_penultimo_mes,inner_join(ausentes_penultimo_mes2,portas_ativas %>% 
                                                filter(floor_date(PEDDTBAIXA, "month")== as.Date(floor_date(Sys.Date()- months(1), "month"))) %>% 
                                                distinct(CLICODIGO),by="CLICODIGO"),by="CLICODIGO")

# max value df
max_rows <- max(nrow(ausentes_penultimo_mes3), nrow(ausentes_ultimo_mes), nrow(ausentes_mes_atual2))


# complete missing values and bind columns
ausentes <-
  cbind(
    ausentes_penultimo_mes3 %>%  rbind(data.frame(CLICODIGO = rep(NA, max_rows - nrow(.)))),
    ausentes_ultimo_mes %>%  rbind(data.frame(CLICODIGO= rep(NA, max_rows - nrow(.)))),
    ausentes_mes_atual2 %>%  rbind(data.frame(CLICODIGO = rep(NA, max_rows - nrow(.))))) %>% as.data.frame() 


ausentes2 <-
  ausentes %>% 
  setNames(c(format(Sys.Date() - months(2), "%B"), 
             format(Sys.Date() - months(1), "%B"), 
             format(Sys.Date(), "%B"))) 


## CHART PORTAS ATIVAS ==================================================

## geral

portas_ativas_plot <-
  portas_ativas2 %>% 
  ggplot(.,aes(x=MES,y=PORTAS_ATIVAS,color=as.factor(ANO),group = as.factor(ANO))) + geom_line() +
  scale_color_manual(values = c("2023" = "#8297b0", "2024" = "#faf31e"))  +
  geom_text(aes(label = PORTAS_ATIVAS), vjust = -0.5, hjust = 1,size=5) +
  geom_point(size = 4.5) +
  theme_minimal() +
  theme(
    plot.background = element_rect(fill = "#324254"),
    panel.background = element_rect(fill = "#324254"),
    legend.background = element_rect(fill = "#324254"),
    legend.title = element_text(color = "white"),
    axis.title.x = element_blank(),
    legend.text = element_text(color = "white", size = 13), 
    axis.text = element_text(color = "white"),
    axis.title = element_text(color = "white"),
    axis.title.y.left = element_text(size = 15),
    plot.title = element_text(color = "white"),
    panel.grid.major.y = element_blank(),
    panel.grid.minor.y = element_blank(),
    panel.grid.major.x = element_line(colour = "#425266"),
    legend.position = "top",
    axis.text.x = element_text(size = 15),
    legend.key.size = unit(2, "lines"),
    axis.text.y = element_blank()
  ) +
  guides(color = guide_legend(title = NULL))


ggsave("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\portas_ativas_plot.jpg", plot = portas_ativas_plot, device = "jpg",width = 11, height =6)

## setores


# Iterate over each unique 'SETOR' and generate a plot
unique_setores <- unique(portas_ativas_setores$SETOR)

for (setor in unique_setores) {
  # Filter data for the current 'SETOR'
  setor_data <- portas_ativas_setores %>% filter(SETOR == setor)
  
  # Generate the plot
  portas_ativas_setores_plot <- setor_data %>%
    ggplot(aes(x = MES, y = PORTAS_ATIVAS, color = as.factor(ANO), group = as.factor(ANO))) +
    geom_line() +
    scale_color_manual(values = c("2023" = "#8297b0", "2024" = "#faf31e")) +
    geom_text(aes(label = PORTAS_ATIVAS), vjust = -0.5, hjust = 1, size = 5) +
    geom_point(size = 4.5) +
    theme_minimal() +
    theme(
      plot.background = element_rect(fill = "#324254"),
      panel.background = element_rect(fill = "#324254"),
      legend.background = element_rect(fill = "#324254"),
      legend.title = element_text(color = "white"),
      axis.title.x = element_blank(),
      legend.text = element_text(color = "white", size = 13),
      axis.text = element_text(color = "white"),
      axis.title = element_text(color = "white"),
      axis.title.y.left = element_text(size = 15),
      plot.title = element_text(color = "white"),
      panel.grid.major.y = element_blank(),
      panel.grid.minor.y = element_blank(),
      panel.grid.major.x = element_line(colour = "#425266"),
      legend.position = "top",
      axis.text.x = element_text(size = 15),
      legend.key.size = unit(2, "lines"),
      axis.text.y = element_blank()
    ) +
    guides(color = guide_legend(title = NULL)) +
    ggtitle(paste("SETOR:", setor))  # Optional: Add setor name to the title
  
  # Extract the name of the SETOR up to the first number
  setor_name <- sub("^(\\D*\\d+).*", "\\1", setor)
  
  # Save the plot as a JPG file with the appropriate name
  image_filename <- paste0("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\portas_ativas_", gsub(" ", "_", setor_name), ".jpg")
  ggsave(image_filename, plot = portas_ativas_setores_plot, device = "jpg", width = 11, height = 6)
}


## CHART GIRO ==================================================

## geral

giro_plot <-
  giro %>% 
  ggplot(.,aes(x=MES,y=GIRO,color=as.factor(ANO),group = as.factor(ANO))) + 
  geom_line() +
  geom_point(size = 4.5) +
  scale_color_manual(values = c("2023" = "#8297b0", "2024" = "#00ffff"))  +
  geom_text(aes(label = GIRO), vjust = -0.5, hjust = 1,size=5) +
  theme_minimal() +
  theme(
    plot.background = element_rect(fill = "#324254"),
    panel.background = element_rect(fill = "#324254"),
    legend.background = element_rect(fill = "#324254"),
    legend.title = element_text(color = "white"),
    axis.title.x = element_blank(),
    legend.text = element_text(color = "white", size = 13), 
    axis.text = element_text(color = "white"),
    axis.title = element_text(color = "white"),
    axis.title.y.left = element_text(size = 15),
    plot.title = element_text(color = "white"),
    panel.grid.major.y = element_blank(),
    panel.grid.minor.y = element_blank(),
    panel.grid.major.x = element_line(colour = "#425266"),
    legend.position = "top",
    axis.text.x = element_text(size = 15),
    legend.key.size = unit(2, "lines"),
    axis.text.y = element_blank()
  ) +
  guides(color = guide_legend(title = NULL))


ggsave("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\giro_plot.jpg", plot = giro_plot, device = "jpg",width = 11, height =6)


## setores

# Iterate over each SETOR in giro_setores and generate individual plots
unique_giro_setores <- unique(giro_setores$SETOR)

for (setor in unique_giro_setores) {
  # Filter data for the current SETOR
  setor_data <- giro_setores %>% filter(SETOR == setor)
  
  # Generate the plot for the current SETOR
  giro_setores_plot <- setor_data %>% 
    ggplot(aes(x = MES, y = GIRO, color = as.factor(ANO), group = as.factor(ANO))) +
    geom_line() +
    geom_point(size = 4.5) +
    scale_color_manual(values = c("2023" = "#8297b0", "2024" = "#00ffff")) +
    geom_text(aes(label = GIRO), vjust = -0.5, hjust = 1, size = 5) +
    theme_minimal() +
    theme(
      plot.background = element_rect(fill = "#324254"),
      panel.background = element_rect(fill = "#324254"),
      legend.background = element_rect(fill = "#324254"),
      legend.title = element_text(color = "white"),
      axis.title.x = element_blank(),
      legend.text = element_text(color = "white", size = 13), 
      axis.text = element_text(color = "white"),
      axis.title = element_text(color = "white"),
      axis.title.y.left = element_text(size = 15),
      plot.title = element_text(color = "white"),
      panel.grid.major.y = element_blank(),
      panel.grid.minor.y = element_blank(),
      panel.grid.major.x = element_line(colour = "#425266"),
      legend.position = "top",
      axis.text.x = element_text(size = 15),
      legend.key.size = unit(2, "lines"),
      axis.text.y = element_blank()
    ) +
    guides(color = guide_legend(title = NULL)) +
    ggtitle(paste("SETOR:", setor))  # Optional: Add setor name to the title
  
  # Extract the name of the SETOR up to the first number
  setor_name <- sub("^(\\D*\\d+).*", "\\1", setor)
  
  # Save the plot as a JPG file with the appropriate name
  image_filename <- paste0("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\giro_", gsub(" ", "_", setor_name), ".jpg")
  ggsave(image_filename, plot = giro_setores_plot, device = "jpg", width = 11, height = 6)
}



## EXCEL ==================================================


wb <- createWorkbook()


## DADOS

addWorksheet(wb, "DADOS")

# DATA ATUAL

writeData(wb, "DADOS", data_atual, startCol = 2, startRow = 1)


writeData(wb, "DADOS", portas_ativas4, startCol = 2, startRow = 3)

writeDataTable(wb, "DADOS", portas_ativas3, startRow = 4, startCol = 2,tableStyle = "TableStyleMedium20")


writeData(wb, "DADOS", qtd_pcs4, startCol = 6, startRow = 3)

writeDataTable(wb, "DADOS", qtd_pcs3, startRow = 4, startCol = 6, tableStyle = "TableStyleMedium21")


writeData(wb, "DADOS", giro3, startCol = 10, startRow = 3)

writeDataTable(wb, "DADOS", giro2, startRow = 4, startCol = 10, tableStyle = "TableStyleMedium19")


## CLIENTES AUSENTES ===========================================


addWorksheet(wb, "CLIENTES AUSENTES")


writeDataTable(wb, "CLIENTES AUSENTES", ausentes2, startRow = 4, startCol = 2,tableStyle = "TableStyleMedium20")

## PORTAS ATIVAS =========================================

addWorksheet(wb, "PORTAS ATIVAS")

insertImage(wb, "PORTAS ATIVAS", "C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\portas_ativas_plot.jpg",units = "in", width = 9, height =5,startRow = 1, startCol = 1)



## GIRO =================================================


addWorksheet(wb, "GIRO")

insertImage(wb, "GIRO", "C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\giro_plot.jpg",units = "in", width = 9, height =5,startRow = 1, startCol = 1)




unique_setores <- unique(portas_ativas_setores$SETOR)

for (setor in unique_setores) {
  # Extract a consistent name pattern until the first number
  setor_pattern_name <- sub("^(\\D*\\d+).*", "\\1", setor)
  
  # Create the worksheet with the extracted name
  addWorksheet(wb, setor_pattern_name)
  
  # Generate filenames using the same consistent naming pattern
  portas_image_filename <- paste0("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\portas_ativas_", gsub(" ", "_", setor_pattern_name), ".jpg")
  giro_image_filename <- paste0("C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\giro_", gsub(" ", "_", setor_pattern_name), ".jpg")
  
  # Check and insert the 'portas_ativas' image
  if (file.exists(portas_image_filename)) {
    insertImage(wb, setor_pattern_name, portas_image_filename, units = "in", 
                width = 9, height = 5, startRow = 1, startCol = 1)
  } else {
    message(paste("Warning: Image not found -", portas_image_filename))
  }
  
  # Check and insert the 'giro' image
  if (file.exists(giro_image_filename)) {
    insertImage(wb, setor_pattern_name, giro_image_filename, units = "in", 
                width = 9, height = 5, startRow = 1, startCol = 12)
  } else {
    message(paste("Warning: Image not found -", giro_image_filename))
  }
}


## SAVE =====

saveWorkbook(wb, "C:\\Users\\REPRO SANDRO\\Documents\\R\\REPORTS\\PORTAS_ATIVAS.xlsx", overwrite = TRUE)

