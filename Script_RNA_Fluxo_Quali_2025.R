#*******************************************************************************************************
# Limpa os dados do ambiente
rm(list = ls())

# 0 Importacao das bibliotecas (verificar se ja estao instaladas)

library(readxl)
library(openxlsx)
library(neuralnet)
library(NeuralNetTools)
library(fastDummies)
#*******************************************************************************************************
#*******************************************************************************************************

#*******************************************************************************************************
#*******************************************************************************************************
#*******************************************************************************************************
# ETAPA APENAS PARA ORGANIZAR A BASE DE DADOS INICIAL
# 1 Importacao de dados (esta lendo o arquivo na pasta do projeto)
dados <- read_excel("Dados_IFC_CUB_Renato.xlsx", sheet="IFC")

# 2 Preparacao dos dados
# 2.1 Filtra apenas dados com altura > 0
dados_a <- subset.data.frame(dados, dados$HT > 0)

# 2.2 Define os dados que serao utilizados para treinamento e validacao
dados_a$Aleat <- runif(nrow(dados_a), 0, 1)
dados_a$Tipo  <- ifelse(dados_a$Aleat < 0.7,'T', 'V')
dados_a$id    <- seq_len(nrow(dados_a))
dados_a$Dap   <- dados_a$CAP/3.1416
dados_a$Grupo <- format(dados_a$DATAMEDICAO, "%Y")

# 2.3 Filtra as colunas que serao utilizadas (verificar no dataframe)
# FAZENDA, CLONE, ESPACAMENTO, ID, DAP, ALTURA, ALEAT, TIPO
dados_b <- dados_a[, c(1,4,5,24, 25, 19, 22, 23, 26)]


#*******************************************************************************************************
# 3 Etapa para organizacao do processamento
dados_c <- subset.data.frame(dados_b, dados_b$Grupo <= 2021)

# 3.1 Normalizando as variaveis continuas
min_max          = c(min(dados_c$Dap),max(dados_c$Dap),min(dados_c$HT),max(dados_c$HT))
dados_c$ALTURAn <- (dados_c$HT-min_max[3])/(min_max[4]-min_max[3])
dados_c$DAPn    <- (dados_c$Dap-min_max[1])/(min_max[2]-min_max[1])

# 3.2 Cria variaveis dummy para as variaveis categoricas
dados_c <- dummy_cols(dados_c, select_columns = c('CLONE', 'ESP'))

# 3.3 Total de neuronios na entrada: clones + espacamentos + quantitativas + bias
n_Entrada = c(length(unique(dados_c$CLONE)) + length(unique(dados_c$ESP)) + 1 + 1)

# 3.4 Filtra apenas dados para treinamento
dados_d <- subset.data.frame(dados_c, dados_c$Tipo == 'T')

# 3.5 Obtem os nomes das variaveis de entrada (continuas + dummy)
colunas  <- data.frame(names(dados_d))
linhas   <- c(1:10)
colunasX <- colunas[-linhas, ]

# 3.6 Cria a formula da relacao entre variavel de saida e variaveis de entrada
modelo <- as.formula(paste("ALTURAn ~ ", paste(colunasX, collapse= " + ")))
modelo

# 3.7 Treina o modelo de rede neural com pesos aleatorios
nneurons = 30
n_hidden = c(nneurons)
nn <- neuralnet(modelo, 
                data          = dados_d, 
                hidden        = n_hidden, 
                linear.output = F, 
                rep           = 1, 
                act.fct       = "logistic", 
                err.fct       = "sse", 
                algorithm     = 'rprop+',
                stepmax       = 100000)

# 3.8 Pega a lista com os pesos da rna treinada
pesos <- nn$weights
pesos

# 3.9 Estima os valores para dados de treinamento e validacao
dados_c$HTn_est <- predict(nn, dados_c)
dados_c$HT_est  <- dados_c$HTn_est*(min_max[4]-min_max[3])+min_max[3]


# 3.10 Calcula a correlacao
grupo_Cor <- data.frame("Grupo"  = dados_c$Grupo[1],
                        "Correl" = cor(dados_c$HT_est,dados_c$HT),
                        "Tipo"   = 'T1')

# 3.11 Limpa os dataframes para utilizar posteriormente
dados_f   <- dados_c[0,]
grupo_Cor <- grupo_Cor[0,]

# 3.12 Substitui os pesos por valores aleatorios entre -0.5 e 0.5
nn$weights
pesos_ini  = nn$weights
lin        = c(1:n_Entrada)
col        = c(1:n_hidden)
col2       = c(1:n_hidden+1)
pesos_new <- pesos_ini
for(a in lin){
  for(b in col){
    valor = runif(1, -0.5, 0.5)
    pesos_new[[1]][[1]][a, b] = valor
  }
}

for(c in col2){
  valor = runif(1, -0.5, 0.5)
  pesos_new[[1]][[2]][c, 1] = valor
}

# 3.13 Armazena os pesos iniciais que serao utilizados nas próximas etapas
pesos_base <- pesos_new

#*******************************************************************************************************
#*******************************************************************************************************
#*******************************************************************************************************






#*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~**
#*******************************************************************************************************
#*******************************************************************************************************
#*
# TESTE 1: Recebe novos dados, treina a rede do zero para cada grupo de dados, salva os pesos
#   

pesos_save <- pesos_base
pesos_save[0] <- pesos_base
i = 1
valor = unique(dados_b$Grupo)

for(v in valor){
  
  # 3.1 Obtem os dados do grupo e normaliza as variaveis contínuas
  dados_c <- subset.data.frame(dados_b, dados_b$Grupo == v)
  min_max = c(min(dados_c$Dap),max(dados_c$Dap),min(dados_c$HT),max(dados_c$HT))
  dados_c$DAPn<- (dados_c$Dap-min_max[1])/(min_max[2]-min_max[1])
  dados_c$ALTURAn<- (dados_c$HT-min_max[3])/(min_max[4]-min_max[3])
  
  # 3.2 Cria variaveis dummy para as variaveis categoricas
  dados_c <- dummy_cols(dados_c,select_columns = c('CLONE', 'ESP'))
  
  # 3.3 Total de neuronios na entrada: clones + espacamentos + quantitativas + bias
  n_Entrada = c(length(unique(dados_c$CLONE)) + length(unique(dados_c$ESP)) + 1 + 1)
  
  # 3.4 Filtra apenas dados para treinamento
  dados_d <- subset.data.frame(dados_c, dados_c$Tipo == 'T')
  
  # 3.5 Obtem os nomes das variaveis de entrada (continuas + dummy)
  colunas<-data.frame(names(dados_d))
  linhas <- c(1:10)
  colunasX<-colunas[-linhas, ]
  
  # 3.6 Cria a formula da relacao entre variavel de saida e variaveis de entrada
  modelo <- as.formula(paste("ALTURAn ~ ", paste(colunasX, collapse= " + ")))
  n_hidden = c(nneurons)
  
  # 3.7 Treina a rede neural MLP
  nn <- neuralnet(modelo,
                  data          = dados_d, 
                  hidden        = n_hidden, 
                  startweights  = pesos_new, 
                  linear.output = F, 
                  rep           = 1, 
                  act.fct       = "logistic", 
                  err.fct       = "sse", 
                  algorithm     = 'rprop+',
                  stepmax       = 100000)
  
  # 3.8 Recebe os pesos calculados
  pesos           <- nn$weights
  
  # 3.9 Estima os valores para a variável de saída para dados de treino e validacao
  dados_c$HTnest <- predict(nn, dados_c)
  dados_c$HTest <- dados_c$HTnest*(min_max[4]-min_max[3])+min_max[3]
  
  #grupo_Cora      <- data.frame("Grupo" = dados_c$Grupo[1],"Correl" = cor(dados_c$HTest, dados_c$HT),"Tipo" = 'T1')
  #grupo_Cor       <- rbind(grupo_Cor, grupo_Cora)
  colnames(dados_f)[11] <- "CLONE"
  colnames(dados_f)[12] <- "ESP"
  colnames(dados_c)[11] <- "CLONE"
  colnames(dados_c)[12] <- "ESP"

  
  dados_f <- rbind(dados_f, dados_c)
  print(v)
  flush.console()
  i = i + 1
  pesos_save[i] <- pesos

}

# 3.10 Calcula as estatisticas de análise

statdf <- data.frame(Teste = character(),
                 Grupo = character(),
                 Tipo = character(),
                 Corr = numeric(),
                 Bias = numeric(),
                 EMP = numeric(),
                 EQM = numeric(),
                 RMSE = numeric())


valorA = unique(dados_f$Grupo)
valorB = unique(dados_f$Tipo)
print(paste("va" , " " , "vb" , " " , "corr" , " " , "bias" , " " , "emp" , " " , "eqm" , " " , "rmse"))
for(va in valorA){
  for(vb in valorB){
    # 3.10.1 Seleciona os dados
    dados_g <- subset.data.frame(dados_f[,c(6, 18)], dados_f$Grupo == va & dados_f$Tipo == vb)
    
    # 3.10.2 Calcula os erros
    dados_g$erro  <- dados_g$HTest-dados_g$HT
    dados_g$errop <- (dados_g$HTest-dados_g$HT)/dados_g$HT
    dados_g$erro2 <- dados_g$erro^2
    
    # 3.10.3 Calcula as estatisticas
    bias = mean(dados_g$erro)
    emp  = mean(dados_g$errop)
    eqm  = mean(dados_g$erro2)
    rmse = eqm^(1/2)
    corr = cor(dados_g$HTest, dados_g$HT)
    
    # 3.10.4 Imprime as estatisticas
    print(paste(va , " " , vb , " " , corr , " " , bias , " " , emp , " " , eqm , " " , rmse))
    flush.console()
    
    newLine <- data.frame(Teste = "1", Grupo = va, Tipo = vb, Corr = corr, Bias = bias, EMP = emp, EQM = eqm, RMSE = rmse)
    statdf <- rbind(statdf, newLine)
  }
}

# Organiza os dados de saída
dados_ft1 <- dados_f
dados_ft1 <- dados_f[,c(4,6,8,9,18)]



#*******************************************************************************************************
#*
# teste 2: Recebe novos dados, treina a rede a partir dos pesos do ano anterior, salva os pesos
#   

pesos_new <- pesos_base

dados_f <- dados_c[0,]
pesos_save2 <- pesos_new
pesos_save2[0] <- pesos_new
i = 1

valor = unique(dados_b$Grupo)
#valor = c(1:188)
for(v in valor){
  dados_c <- subset.data.frame(dados_b, dados_b$Grupo == v)
  min_max = c(min(dados_c$Dap),max(dados_c$Dap),min(dados_c$HT),max(dados_c$HT))
  dados_c$DAPn<- (dados_c$Dap-min_max[1])/(min_max[2]-min_max[1])
  dados_c$ALTURAn<- (dados_c$HT-min_max[3])/(min_max[4]-min_max[3])
  
  # 3.2 Cria variaveis dummy para as variaveis categoricas
  dados_c <- dummy_cols(dados_c,select_columns = c('CLONE', 'ESP'))
  
  # 3.3 Total de neuronios na entrada: clones + espacamentos + quantitativas + bias
  n_Entrada = c(length(unique(dados_c$CLONE)) + length(unique(dados_c$ESP)) + 1 + 1)
  
  # 3.4 Filtra apenas dados para treinamento
  dados_d <- subset.data.frame(dados_c, dados_c$Tipo == 'T')
  
  # 3.5 Obtem os nomes das variaveis de entrada (continuas + dummy)
  colunas<-data.frame(names(dados_d))
  linhas <- c(1:10)
  colunasX<-colunas[-linhas, ]
  
  
  # 3.6 Cria a formula da relacao entre variavel de saida e variaveis de entrada
  modelo <- as.formula(paste("ALTURAn ~ ", paste(colunasX, collapse= " + ")))
  n_hidden = c(nneurons)
  nn <- neuralnet(modelo,
                  data          = dados_d, 
                  hidden        = n_hidden, 
                  startweights  = pesos_new, 
                  linear.output = F, 
                  rep           = 1, 
                  act.fct       = "logistic", 
                  err.fct       = "sse", 
                  algorithm     = 'rprop+',
                  stepmax       = 100000)
  pesos           <- nn$weights
  dados_c$HTnest <- predict(nn, dados_c)
  dados_c$HTest <- dados_c$HTnest*(min_max[4]-min_max[3])+min_max[3]
  
  grupo_Cora      <- data.frame("Grupo" = dados_c$Grupo[1],"Correl" = cor(dados_c$HTest,dados_c$HT),"Tipo" = 'T1')
  grupo_Cor       <- rbind(grupo_Cor, grupo_Cora)
  colnames(dados_f)[11] <- "CLONE"
  colnames(dados_f)[12] <- "ESP"
  colnames(dados_c)[11] <- "CLONE"
  colnames(dados_c)[12] <- "ESP"
  dados_f <- rbind(dados_f, dados_c)
  print(v)
  #print(pesos)
  flush.console()
  i = i + 1
  pesos_save2[i] <- pesos
  
  # 3.12 Substitui os pesos por valores aleatorios
  pesos_new = nn$weights

}

# Organiza os dados de saída
dados_ft2 <- dados_f
dados_ft2 <- dados_f[,c(4,6,8,9,18)]

# 3.10 Calcula as estatisticas de análise
valorA = unique(dados_f$Grupo)
valorB = unique(dados_f$Tipo)
print(paste("va" , " " , "vb" , " " , "corr" , " " , "bias" , " " , "emp" , " " , "eqm" , " " , "rmse"))
for(va in valorA){
  for(vb in valorB){
    # 3.10.1 Seleciona os dados
    dados_g <- subset.data.frame(dados_f[,c(6, 18)], dados_f$Grupo == va & dados_f$Tipo == vb)
    
    # 3.10.2 Calcula os erros
    dados_g$erro  <- dados_g$HTest-dados_g$HT
    dados_g$errop <- (dados_g$HTest-dados_g$HT)/dados_g$HT
    dados_g$erro2 <- dados_g$erro^2
    
    # 3.10.3 Calcula as estatisticas
    bias = mean(dados_g$erro)
    emp  = mean(dados_g$errop)
    eqm  = mean(dados_g$erro2)
    rmse = eqm^(1/2)
    corr = cor(dados_g$HTest, dados_g$HT)
    print(paste(va , " " , vb , " " , corr , " " , bias , " " , emp , " " , eqm , " " , rmse))
    flush.console()
    
    newLine <- data.frame(Teste = "2", Grupo = va, Tipo = vb, Corr = corr, Bias = bias, EMP = emp, EQM = eqm, RMSE = rmse)
    statdf <- rbind(statdf, newLine)
  }
}


# Publicando os resultados em Excel

wb <- createWorkbook()

# Adicionar primeira planilha
addWorksheet(wb, "Res_Padrao")
writeData(wb, "Res_Padrao", dados_ft1)

# Adicionar segunda planilha
addWorksheet(wb, "Res_Proposta")
writeData(wb, "Res_Proposta", dados_ft2)

# Adicionar terceira planilha
addWorksheet(wb, "statdf")
writeData(wb, "statdf", statdf)

# Salvar o arquivo
excel_name = paste("Res_",nneurons,"neurons.xlsx", sep = "")
saveWorkbook(wb, file = excel_name, overwrite = TRUE)
