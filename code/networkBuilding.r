# ---------------------------------------------------------------------------------
# Construir as rede de interesse
# Build interest networks
#
# Autores: Hernane Borges de Barros Pereira, Ludmilla Monfort Oliveira Sousa, 
# Maíra Lima de Souza, and Marcelo A. Moret 
# Authors: Hernane Borges de Barros Pereira, Ludmilla Monfort Oliveira Sousa, 
# Maíra Lima de Souza, and Marcelo A. Moret                                                                   |
#
# Última atualização: 08/03/2023 
# Last update: March 08, 2023                            
# ---------------------------------------------------------------------------------

install.packages(c("data.table","tidyverse","tibble","stringr","sjmisc","readxl",
                   "openxlsx","dplyr","Xmisc","rapportools","abjutils","gtools","bibliometrix", "regexPipes"))
install.packages("Xmisc")
# Bibliotecas para manipular base de dados
# Libraries to manipulate database
library(data.table)
library(tidyverse)
library(tibble)

# Bibliotecas para para manipular strings
# Libraries for manipulating strings
library(stringr)
library(sjmisc)

# Bibliotecas para excel
# Libraries for excel
library(readxl)
library(openxlsx)
library(dplyr)

# Bibiloteca para manipular dados e executa comandos is.empty, is.character, is.number etc.
# Library to manipulate data and execute commands is.empty, is.character, is.number, etc
require(rapportools)

require(Xmisc)
library(regexPipes)
require(abjutils)

# Carregar os pacotes/bibliotecas para realizar analise combinatoria
# Load the packages/libraries to perform combinatorial analysis
#install.packages("gtools")
require(gtools)

# Bibiloteca para rede citação
# Library for citation network
require(bibliometrix)

# Carregar a base de dados original no diretório de trabalho
# Load the original database into the working directory

# Definir o diretório de trabalho
# Set working directory
# ../database/
setwd("/Users/hbbpereira/Downloads/OverviewCOVID19pandemic/database/")
path = "/Users/hbbpereira/Downloads/OverviewCOVID19pandemic/database/WOS_PBL2_1006.xlsx"

# Caso a planilha tenha mais de uma sheet, percorrer todas as sheets
# If the worksheet has more than one sheet, go through all the sheets
#
# Indicar as colunas que comporão o dataset
# Indicate the columns that will compose the dataset
sheet <- loadWorkbook(path)
sheetNames <- sheets(sheet)
data<-as.data.frame(matrix(,ncol=0,nrow=0))
var_norteadora=c('Authors','Author.Full.Name','Document.Title','Publication.Name','Document.Type','Author.Keywords','Keywords.Plus.','Abstract',
'Author.Address','Reprint.Address','Cited.References','Cited.Reference.Count','Web.of.Science.Core.Collection.Times.Cited.Count',
'Total.Times.Cited.Count..Web.of.Science.Core.Collection..BIOSIS.Citation.Index..Chinese.Science.Citation.Database..Data.Citation.Index..Russian.Science.Citation.Index..SciELO.Citation.Index.',
'Usage.Count..Last.180.Days.','Usage.Count..Since.2013.','Publisher','International.Standard.Serial.Number..ISSN.',
'Electronic.International.Standard.Serial.Number..eISSN.','International.Standard.Book.Number..ISBN.',
'X29.Character.Source.Abbreviation','ISO.Source.Abbreviation','Publication.Date','Year.Published','Volume','Issue',
'Beginning.Page','Ending.Page','Article.Number','Digital.Object.Identifier..DOI.','Book.Digital.Object.Identifier..DOI.',
'Page.Count','Web.of.Science.Categories','Research.Areas','PubMed.ID','Open.Access.Indicator')
for(i in 1:length(sheetNames))
{
    tmp = openxlsx::read.xlsx(path, i, detectDates= TRUE)%>%data.frame()
    data<- rbind(data, tmp[var_norteadora])
}

# Obter a quantidade de linhas e colunas do dataset
# Get the number of rows and columns in the dataset
dim(data)

# Percorrer todo o dataset para inserir a coluna "COD_PROD"
# Go through the entire dataset to insert the column "COD_PROD"
for (j in 1:nrow(data)){
   data[j,"COD_PROD"]<-str_c("PROD_00",j)
}

# Reorganizar as colunas para que a nova coluna seja a primeira
# Rearrange the columns so that the new column is first
data<- data %>% select(COD_PROD, everything())

# Excluir obras sem autor (i.e. Anonymous)
# Exclude works without an author (i.e. Anonymous) 
Data_excluidos<-subset(data,(is.empty(Authors)|Authors=='[Anonymous]'))

# Criar um novo dataset sem os autores com nome em branco 
# Create a new dataset without authors with blank name
data_v2<-subset(data,!is.empty(Authors))

# Criar um novo dataset sem os autores Anonymous
# Create a new dataset without Anonymous authors
data_v2<-subset(data_v2,Authors!='[Anonymous]')

# Criar uma nova versão do dataset
# Create a new version of the dataset

data_v2 <- data_v2 %>% 
  select(Document.Title, Year.Published,Abstract,Author.Keywords,Publication.Name,
         Reprint.Address,Digital.Object.Identifier..DOI.,Cited.References,Author.Full.Name,Authors)

# Percorrer todo o dataset para inserir a coluna "NEWCOD_PROD"
# Go through the entire dataset to insert the column "NEWCOD_PROD"
for (j in 1:nrow(data_v2)){
   data_v2[j,"NEWCOD_PROD"]<-str_c("PROD_00",j)
}
# Reorganizar as colunas para que a nova coluna seja a primeira
# Rearrange the columns so that the new column is first
data_v2<- data_v2 %>% select(NEWCOD_PROD, everything())

# Remover os acentos e pontuações diferentes do DOI
# Remove accents and different punctuation from DOI
data_v2$Digital.Object.Identifier..DOI.<-data_v2$Digital.Object.Identifier..DOI.%>%abjutils::rm_accent()

# Idenitificar se há mais de uma ocorrência de um mesmo DOI
# Identify if there is more than one occurrence of the same DOI
df<- data_v2%>%group_by(Digital.Object.Identifier..DOI.) %>%count(Digital.Object.Identifier..DOI.)%>% ungroup()%>%filter(n>1) # seleciono as obras duplicadas

lista<-df$Digital.Object.Identifier..DOI. 
#listadrop<-c("1") 
# ou/or 
listadrop<-c("10") 
lista<-lista[-as.integer(listadrop)] 

# Guardar as linhas que contêm o mesmo DOI
# Save lines that contain the same DOI
# Manter no dataset original apenas a primeira ocorrência do DOI, apagando as demais ocorrências
# Keep only the first occurrence of the DOI in the original dataset, deleting the other occurrences
    var_select_cod=list()
    data_doi_data_select<- data.frame( 
                 DOI=character(),
                 r =list(),
                 COD=list(),
                 stringsAsFactors=FALSE) 
     for (j in seq(1,length(lista),1)){
        r <-which(lista[j]==data_v2$Digital.Object.Identifier..DOI.)
        if(length(r)>1) 
        {
        data_doi_data_select[j,"DOI"]<-lista[j]
        data_doi_data_select$r[j]<-list(which(lista[j]==data_v2$Digital.Object.Identifier..DOI.)) 
        data_doi_data_select$COD[j]<-list(data_v2[r,1])  
        linha<-r[2]
        data_v2 <- data_v2[-linha,]
        }
        else{
         if(length(r)==1) break()    
        } 
    }

# ---------------------------------------------------------------------------------
# Função para rede coautoria
# Function for co-authoring network
# ---------------------------------------------------------------------------------

# Função que abrevia nome de autor
# Function to abbreviate author name
# De/From: DE SOUZA, MAIRA LIMA -> Para/To: DE SOUZA, M.L.

abreviar_nome_autor <- function(autor) {
 
    lista_preposicoes = c('DA', 'DO', 'DAS', 'DOS', 'DE')
    lista_abreviacao = list()
    i=autor
    partes_nome<- str_split(i,"",simplify = TRUE) 
    
    if(',' %in% partes_nome){
        
        abreviacao<-str_trim(unlist(str_extract_all(i, "^.+?,|[A-Z]+|\\s[A-Z]+"))) 
        if(nchar(abreviacao[2])>1){ # se nome tiver mais de uma letra eg: ML
           partes_nome2<- str_split(abreviacao[2],"",simplify = TRUE)
           abreviacao[2]<-str_c(paste0(partes_nome2,collapse="."),".")
        }       
        else{
           abreviacao[2]<-str_c(paste0(abreviacao[2],collapse="."),".")
        } 
        abreviacao[1]<-ifelse(str_contains(abreviacao[1], "-"),str_trim(str_replace(abreviacao[1], "-", "")), str_trim(abreviacao[1])) 
        
        lista_abreviacao<-str_c(abreviacao[1],abreviacao[2]) 
    
    } 
        
    else{
                # Caso o nome do autor não esteja marcado com vírgula, pegar o último sobrenome e abreviar o resto
                # If the author's name is not marked with a comma, take the last name and abbreviate the rest
                #lista_abreviacao.append(autor)
                
                # Gerar uma lista de partes do nome do autor
                # Generate a list of author name parts
                lista_partes_nome_autor = str_split(i,' ', simplify =TRUE) 
                
                
                # Obter o último nome para usá-lo como parte não abreviada
                # Get last name to use as unabbreviated part
                tam_nome<-length(lista_partes_nome_autor)
                
                if (tam_nome>1){ 
                    
                    abreviacao=str_c(lista_partes_nome_autor[tam_nome-1]) 
                    if (abreviacao %in% lista_preposicoes){
                        abreviacao<-str_trim(str_c(abreviacao,lista_partes_nome_autor[tam_nome],","))
                        parte2<-lista_partes_nome_autor[1:tam_nome-2]
                        
                    }
                
                    else{
                        abreviacao=str_c(lista_partes_nome_autor[tam_nome])
                        abreviacao<-str_c(abreviacao,"," )
                        parte2<-lista_partes_nome_autor[1:tam_nome-1]
                        
                    }
                   
                
                    # Colocar o nome na lista "parte2" e no formato SOBRENOME,N1.N2.
                    # Put the name in the list "part2" and in the format SURNAME,N1.N2.
                    for (k in seq(1,length(parte2),1)){
                
                        parte = parte2[k]
                        parte_aux=str_split(parte,"", simplify =TRUE)
                        parte=str_c(parte_aux[1],'.')#abreviando 
                        abreviacao=str_c(abreviacao,parte)
                                
                    } 
                 
                 lista_abreviacao<-cbind(lista_abreviacao,abreviacao) 
                 
                } 
                else # Recuperar autores sem sobrenome
                     # Retrieve authors without last name
                    
                {
                 abreviacao=lista_partes_nome_autor
                 lista_abreviacao<-cbind(lista_abreviacao,abreviacao)  
                }           
   } 
    return(lista_abreviacao)
}

# ---------------------------------------------------------------------------------
# Funções para tratar os título e construir redes semânticas
# Functions to handle titles and build semantic networks
# Função 1 - ajustar_nomes: Remover espaços em branco, acentos e caracteres especiais
# Function 1 - ajustar_nomes: Remove whitespace, accents and special characters
# Função 2 - ajustar_nomes2: Tratar as palavras ligadas a covid19
# Function 2 - ajustar_nomes2: Treat words linked to covid19 
# Função 3 - int_to_words: Transformar o dígito numérico em texto, e.g. 1 -> one
# Function 3 - int_to_words: Transform numeric digit into text, e.g. 1 -> one
# Função 4 - .simpleCap: Colocar a primeira letra do título maiúscula e o restante em minúsculo
# Function 4 - .simpleCap: Put the first letter of the title in uppercase and the rest in lowercase
# ---------------------------------------------------------------------------------

#http://www.botanicaamazonica.wiki.br/labotam/doku.php?id=bot89:precurso:1textfun:inicio
#https://gomesfellipe.github.io/post/2017-12-17-string/string/
#install.packages("abjutils")

require(abjutils)
ajustar_nomes=function(titulo){
  novo<-titulo%>%
  stringr::str_trim() %>% 
  stringr::str_to_lower() %>% 
  abjutils::rm_accent() %>% 
  stringr::str_replace_all("[,*:&/' '!.();?]", " ") %>% 
  stringr::str_replace_all("_+", " ") 
 return(novo)  
}

ajustar_nomes_2=function(x){
    
    partes_nome<- str_split(toupper(x)," ")[[1]] 
    partes_nome
    vc_covid<-c("2019-NOVEL","2019-NOVEL-CORONAVIRUS","COVID-19","COVID19","2019-NCOV","SARS-NCOV-2","SARS-COV-2","SARSCoV-2","CORONAVIRUS-2") 

    for (k in seq(1,length(partes_nome),1)){
        parte = partes_nome [k]
        
        if(nchar(parte)==1 && str_detect(partes_nome [k],"[a-z|A-Z]") ){ 
           partes_nome [k]=""
           next  
        }
        
        if(parte=="CORONAVIRUS" && k<length(partes_nome)) {
            if ( partes_nome [k+1]=="2"){
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
            }
        }
                  
        if(parte=="NON-COVID-19") { 
            partes_nome [k]="NON COVIDONENINE"
            next
        }
        
        if(parte=="POST-COVID-19") { 
            partes_nome [k]="POST COVIDONENINE"
            next
        }
        
        if(parte=="2019-NOVEL" && k<length(partes_nome) && partes_nome [k+1]=="CORONAVIRUS") { 
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
        }
        
        if(parte=="2019" && k<length(partes_nome) && partes_nome [k+1]=="NOVEL" && partes_nome [k+2]=="CORONAVIRUS") { 
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            partes_nome [k+2]=""
            next
        }
       if(parte=="SARS" && k<length(partes_nome) && partes_nome [k+1]=="CoV-2") { 
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
        } 
         
        if(parte=="CORONAVIRUS" && k<length(partes_nome) && partes_nome [k+1]=="DISEASE" && partes_nome [k+2]=="2019") { 
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]="DISEASE"
            partes_nome [k+2]=""
            next
        }
                        
        if(parte=="CORONA" && k<length(partes_nome) && partes_nome [k+1]=="DISEASE") { 
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
        }
        if(parte=="CORONAVIRUS" && k<length(partes_nome) && partes_nome [k+1]=="DISEASE-19") { 
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]="DISEASE"
            next
        }
        if(parte=="COVID" && k<length(partes_nome) && partes_nome [k+1]=="DISEASE") { 
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
        }
        if(parte=="COVID" && k<length(partes_nome) && partes_nome [k+1]=="19") { 
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
        }
        if(parte=="COVID-19-ASSOCIATED"|parte== "SARS-COV-2-ASSOCIATED") { 
            partes_nome [k]="COVIDONENINE ASSOCIATED"
            next
        }
        if(parte=="COVID-19-MOTIVATED") { 
            partes_nome [k]="COVIDONENINE MOTIVATED"
            next
        }
        if(parte=="COVID-19-RELATED") { 
            partes_nome [k]="COVIDONENINE RELATED"
            next
        } 
        if(parte=="COVID-19-RECOMMENDATIONS") { 
            partes_nome [k]="COVIDONENINE RECOMMENDATIONS"
            next
        }
              
        if(parte %in%vc_covid){ 
            partes_nome [k]="COVIDONENINE"
            next
        } 
        
        # Colocar os números por extenso
        # Put the numbers in words
        if(str_detect(partes_nome [k],"[0-9]"))
        {
            num<-gsub("[^0-9.]+", "", partes_nome [k])
            digito= int_to_words(as.numeric(gsub("[^0-9.]+", "", partes_nome [k])))
            if(str_detect(partes_nome [k],"-") && !str_detect(partes_nome [k],"[a-z|A-Z]")){
              consoante=gsub ("[^a-z|A-Z.]+","",partes_nome [k]) 
              partes_nome [k] = gsub("[[:space:]]", "", paste(toupper(digito),consoante,collapse=""))  
            }
            else{
                frase=str_replace(partes_nome[k],num,digito)
                partes_nome [k] = gsub("[[:space:]]", "", paste(toupper(frase),collapse=""))
            }
                 
        }
    } 
     
    x<-paste(partes_nome, collapse=" ")
    
    # Eliminar o excesso de espaço entre as palavras
    # Eliminate excess space between words
    novo_nome<-gsub("[[:space:]]","_",x) 
    novo_nome<-gsub('[\"]', '', novo_nome) 
    
    titulofinal<-novo_nome%>%
      stringr::str_replace_all("_+", " ") 
    
    return(titulofinal)
}

.simpleCap <- function(x) { 
    s <- strsplit(x, " ")[[1]]
    paste(toupper(substr(x,1,1)),substring(x, 2), sep = "", collapse = " ") 
}

int_to_words <- function(x) {
    
    numero<-list()
    digits <- str_split(as.character(x), "",simplify = TRUE)
    nDigits <- length(digits)
    
    words <- c('zero', 'one', 'two', 'three', 'four',
                'five', 'six', 'seven', 'eight', 'nine',
                 'ten')
    
    for (k in seq(1,nDigits,1)){
    parte = digits [k]
    index <- as.integer(digits) + 1
    numero<-words[index]   
       
    } 
    numero= paste(numero,collapse = "")  
    return(numero)
   
}

# ---------------------------------------------------------------------------------
# 1. Separar os nomes da coluna AUTHOR
# 1. Separate AUTHOR column names
# 2. Abreviar o nome dos autores
# 2. Abbreviate authors' names
# 3. Criar o dataset que terá o vocabulário de controle do nome dos autores tratados
# 3. Create the dataset that will have the vocabulary to control the names of the treated authors
# 4. Criar o dataset que terá a lista de autores com nome tratado (abreviado) quee será a base para a criação da rede coautoria
# 4. Create the dataset that will have the list of authors with treated name (abbreviated) that will be the basis for the creation of the co-authorship network
# ---------------------------------------------------------------------------------

j=0

DataDePara<-as.data.frame(matrix(,ncol=0,nrow=0)) 

DataComb<-data.frame(AUTHOR=list(), 
                 YEAR=character(), 
                 NEWCOD_PROD=character(), 
                 stringsAsFactors=FALSE) 
DataPBL<-as.data.frame(matrix(,ncol=0,nrow=0)) 
DataExcluido<-as.data.frame(matrix(,ncol=0,nrow=0)) 

linhaautor_drop<-list() 

for (j in 1:nrow(data_v2)){
      
    autoria<-c('') 
    autoria_aux<-c('')
    
    nomecompleto<-c('') 
    nomecompleto_aux<-c('')
    
    autoria_aux<-str_split(unlist(as.list(data_v2[j,11])), fixed(';')) 
    autoria<-str_split(unlist(autoria_aux), fixed('""')) 
    
    nomecompleto_aux<-str_split(unlist(as.list(data_v2[j,10])), fixed(';')) 
    nomecompleto<-str_split(unlist(nomecompleto_aux), fixed('""'))  
     
    count<-1
    totalautoresobra<-length(autoria)
    listadatacomb<-list()
    cod_prod<-data_v2[j,1]
    DataPBL[j,"NEWCOD_PROD"]<-cod_prod
    DataPBL[j,"TITULO"]<-data_v2[j,2]
    DataPBL[j,"ANO"]<-data_v2[j,3]
    DataPBL[j,"RESUMO"]<-data_v2[j,4]
    DataPBL[j,"PALAVRAS_CHAVES"]<-data_v2[j,5]
    DataPBL[j,"PUBLICACAO"]<-data_v2[j,6]
    DataPBL[j,"AFILIACAO"]<-data_v2[j,7]
    DataPBL[j,"DOI"]<-data_v2[j,8]
    DataPBL[j,"CR"]<-data_v2[j,9]
       
    for (i in autoria){ 
      
        if (is.na(i)| i==" "| i=="...") break 

        nomeabreviado<-toupper(abreviar_nome_autor(i))
        nomeabreviado<-gsub("[[:space:]]", "", nomeabreviado) 
        
        coluna<-str_c("AUTOR_", count)
        coluna_para<-str_c("VC_AUTOR_",count)    
              
        DataPBL[j,coluna]<-nomeabreviado                  
    
        DataDePara[j,"NEWCOD_PROD"]<-cod_prod
        DataDePara[j,coluna]<-lstrip(toupper(nomecompleto[count]), char = " ")
        DataDePara[j,coluna_para]<-nomeabreviado
        
        
        listadatacomb=cbind(listadatacomb,nomeabreviado) 
        
        count<-count+1
   }
   
    DataComb[j,"AUTHOR"]<-""
    DataComb$AUTHOR[j]<-list(listadatacomb) 
    DataComb[j,"YEAR"]<-data_v2[j,3]
    DataComb[j,"NEWCOD_PROD"]<-cod_prod

}

# ---------------------------------------------------------------------------------
# Aplicação da função de combinação 
# Application of the combination function
# ---------------------------------------------------------------------------------

dados_combina_aut <- DataComb %>% 
  select(AUTHOR,YEAR,NEWCOD_PROD)%>% as.data.table()

# Carregar os pacotes/bibliotecas para realizar análise combinatória
# Load packages/libraries to perform combinatorial analysis
#install.packages("gtools")
require(gtools)

matriz<-NULL
m<-matrix()
data_comb <- data.frame()
ano<-""

for (n in 1:nrow(dados_combina_aut)){
    listaautoria<-list()
    
    listaautoria<-toupper(str_split(unlist(dados_combina_aut[n,1]), fixed('""')))
    listaautoria<-unique(sort(listaautoria))
    totalautoresobra<-length(listaautoria)
        
    
    if(totalautoresobra>=2)
    {
        matriz=gtools::combinations(totalautoresobra,2,listaautoria,repeats=FALSE)
        v1<-matriz[,1]
        v2<-matriz[,2]
        ano<-as.character(dados_combina_aut[n,2])
        cod<-as.character(dados_combina_aut[n,3])
        v3<-rep(ano, times= length(v1))
        v4<-rep(cod, times= length(v1))
        m<-cbind(v1,v2,v3,v4)
        
    }
    else 
        {
        matriz<-listaautoria
        v1<-matriz
        v2<-matriz
        ano<-as.character(dados_combina_aut[n,2])
        cod<-as.character(dados_combina_aut[n,3])
        v3<-rep(ano, times= length(v1))
        v4<-rep(cod, times= length(v1))
        m<-cbind(v1,v2,v3,v4)
        
        
    }
    data_comb<-rbind(data_comb,m)
}

names(data_comb)<-c("DE","PARA","ANO","NEWCOD_PROD")

# ---------------------------------------------------------------------------------
# Tratar o título e gerar o arquivo para redes semânticas
# Treat the title and generate the file for semantic networks
# ---------------------------------------------------------------------------------

j=0

DataTituloVC<-as.data.frame(matrix(,ncol=0,nrow=0)) 

for (j in 1:nrow(DataPBL)){
 
  titulos<-list()
  titulos<-DataPBL[j,2]
  DataTituloVC[j,"NEWCOD_PROD"]<-DataPBL[j,1]  
  DataTituloVC[j,"TITULO_B"]<-lstrip(DataPBL[j,2],char = " ")
  
  titulo=ajustar_nomes(titulos) 
  titulo=ajustar_nomes_2(titulo) 
  titulo=paste(str_replace_all(titulo,"-",""), collapse="") 
  
  titulo=.simpleCap(tolower(titulo)) 
    
  titulos=cbind(titulos,titulo) 
    
  DataPBL[j,2]<-lstrip(titulo, char = " ") 
  DataTituloVC[j,"TITULO_A"]<-lstrip(titulo, char = " ") 
 
}

# ---------------------------------------------------------------------------------
# Criar o dataset que coleta as citações
# Create the dataset that collects the quotes
# ---------------------------------------------------------------------------------

start_time <- Sys.time()

novodoi=NULL
autor=NULL
ano=NULL

Data_CR<-data.frame( 
                 COD_PROD_OR=character(),
                 COD_PROD_CIT=character(), 
                 AUTHOR=character(),
                 ANO=character(),
                 DOI=character(),          
                 stringsAsFactors=FALSE) 

listadoi=list()

i=1

for (j in seq(1,nrow(DataPBL),1)){

    autoria_aux=list()
    autoria=list()
    cat("'linha data \n\t", j, "\n")
    
    autoria_aux<-str_split(unlist(DataPBL[j,"CR"]), fixed(';')) 
    autoria<-str_split(unlist(autoria_aux), fixed(',')) 
    
    codigoproducao<-DataPBL[j,"NEWCOD_PROD"]
    
    if(is.empty(DataPBL[j,"CR"])){
     next
    }
    
    for (aut in autoria){
      
        Data_CR[i,"COD_PROD_OR"]<-codigoproducao
                 
        autor=aut[1] 
        autor=trimws(trimES(base::gsub("[[:punct:]]"," ",autor))) 
        ano=aut[2] 
        
        if(length(k <- aut %>% regexPipes::grep("DOI"))){
            novodoi<-trim.leading(aut[k])
            
           if(length(novodoi)>1){
             novodoi<-trim.leading(novodoi[2])
           }
        novodoi<-novodoi %>% regexPipes::gsub("([DOI])", "_")
        novodoi<-novodoi%>%stringr::str_replace_all("_+", " ") 
        novodoi<-gsub("\\[|\\]", "", novodoi) 
        novodoi<-trim.leading(novodoi)

        }
        else{
          novodoi=str_c("sem doi","_CIT_",i)  
        }
       Data_CR[i,"COD_PROD_CIT"] <-str_c(codigoproducao,"_CIT_",i)
        
       Data_CR[i,"AUTHOR"]<-autor
        
       if (!is.empty(ano) && str_detect(ano,"[0-9]")&& !str_detect(ano,"[a-zA-Z]")){ 
           ano<-ano
       }
       else{
           ano<-""
       } 
        
       Data_CR[i,"ANO"]<-ano
       Data_CR[i,"DOI"]<-novodoi 
     
        i<-i+1 
  }
    
}

end_time <- Sys.time()
end_time - start_time

# ---------------------------------------------------------------------------------
# Verificar a ocorrência de DOIs duplicados em Data_CR
# Check for duplicate DOIs in Data_CR
# ---------------------------------------------------------------------------------

count<-1
listadoi<-unique(Data_CR$DOI)
data_doi_select<- data.frame( 
                 DOI=character(),
                 r =list(),                   
                 stringsAsFactors=FALSE) 
    
  
     for (j in seq(1,length(listadoi),1)){

        r <-which(listadoi[j]==Data_CR$DOI)
        
        data_doi_select[j,"DOI"]<-listadoi[j]
        data_doi_select$r[j]<-list(which(listadoi[j]==Data_CR$DOI))
         
    }

# ---------------------------------------------------------------------------------
# Criar dataset com a estrutura/informação "Obra é citada por"
# Create dataset with structure/information "Work is cited by"
# ---------------------------------------------------------------------------------
 
start_time <- Sys.time()
count<-1
var_select_doi=NULL
var_select_cod_or=NULL
var_select_codcit=NULL
var_select_autor=NULL
var_select_ano=NULL

data_var_select<- data.frame( 
                 DOI=character(),
                 AUTHOR=character(),
                 ANO=character(),   
                 COD_OR=list(),
                 COD_CIR=list(),          
                 stringsAsFactors=FALSE) 
    
        
       for (j in 1:nrow(data_doi_select)){
         
        r <-unlist(data_doi_select[j,2])  
                      
             var_select_doi<-data_doi_select[j,1] 
             var_select_autor<-as.list(Data_CR[r,3])
             var_select_codcit<-as.list(Data_CR[r,2])
             var_select_cod_or<-as.list(Data_CR[r,1])
             var_select_ano<-as.list(Data_CR[r,4])
             
             data_var_select[count,"DOI"]<-var_select_doi
             data_var_select[count,"AUTHOR"]<-var_select_autor[1]
             data_var_select[count,"ANO"]<-var_select_ano[1]
           
             for (d in 1:length(r)){
                coluna<-str_c("COD_OR_", d)
                coluna_para<-str_c("COD_CIR_",d)    
                data_var_select[count,coluna]<-var_select_cod_or[d]                      
            }

           count= count+1  

    } 

end_time <- Sys.time()
end_time - start_time

# ---------------------------------------------------------------------------------
# Importar o "COD_PROD" dos DOIs existentes no dataset do vocabulário de controle de DOIs
# Import the "COD_PROD" of existing DOIs in the DOI control vocabulary dataset
# ---------------------------------------------------------------------------------

DataCitacaovc<-data.frame( 
                 COD_PROD=character(),
                 AUTHOR=list(),
                 DOI=character(),          
                 stringsAsFactors=FALSE) 

for (j in 1:nrow(data_v2)){
     
    DataCitacaovc[j,"COD_PROD"]<-data_v2[j,"NEWCOD_PROD"]
    DataCitacaovc[j,"AUTHOR"]<-data_v2[j,"Authors"]
    DataCitacaovc[j,"DOI"]<-trim.leading(as.character(data_v2[j,'Digital.Object.Identifier..DOI.']))
       
}

# ---------------------------------------------------------------------------------
# Inserir as obras no vocabulario de controle
# Insert the works in the control vocabulary
# ---------------------------------------------------------------------------------

data_var_select_v2<-merge(data_var_select, DataCitacaovc, by.x="DOI", by.y="DOI", all.x=TRUE, sort = FALSE)

data_var_select_v2<-data_var_select_v2 %>% select(COD_PROD,everything())

# ---------------------------------------------------------------------------------
# Verificar a existência de DOIs nos datasets
# Check for DOIs in datasets
# ---------------------------------------------------------------------------------

listadoi<-unique(DataCitacaovc$DOI)

data_doi_citacao_comum<- data.frame( 
                 DOI=character(),
                 r =list(),                   
                 stringsAsFactors=FALSE) 
count=1
for (j in 1:length(listadoi)){
    
    if(!is.empty(listadoi[j])){
        
        r <-which(listadoi[j]==data_var_select_v2$DOI)
    
        if (length(r)==0){ 
        }
        else{
           if (length(r)>=1){
            data_doi_citacao_comum[count,"DOI"]<-listadoi[j]
            data_doi_citacao_comum$r[count]<-list(which(listadoi[j]==data_var_select_v2$DOI))
            count=count+1
           } 
        }   
    }   
} 

# ---------------------------------------------------------------------------------
# Criar tabela com as citações
# Create table with citations
# ---------------------------------------------------------------------------------

tam=1750 # Último código + 1 (Last code + 1)

listacontroladoi<-DataCitacaovc$DOI

for (j in seq(1,nrow(data_var_select_v2),1)){ 
    
    if(data_var_select_v2[j,"DOI"] %in% listacontroladoi){
    }
    else{
        
      DataCitacaovc[tam,"COD_PROD"]<-str_c("PROD_00",tam)
      DataCitacaovc[tam,"AUTHOR"]<-data_var_select_v2[j,"AUTHOR.x"]
      DataCitacaovc[tam,"ANO"]<-data_var_select_v2[j,"ANO"]
      DataCitacaovc[tam,"DOI"]<-data_var_select_v2[j,"DOI"]
      data_var_select_v2[j,"COD_PROD"]<-str_c("PROD_00",tam) 
       
      tam=tam+1
        
    }

}  
data_var_select_v2<- data_var_select_v2 %>% select(COD_PROD, everything())    

# ---------------------------------------------------------------------------------
# Importar "COD_PROD" gerado para o dataset
# Import generated "COD_PROD" into the dataset
# ---------------------------------------------------------------------------------

Data_CR_COD_ANTI<-inner_join(Data_CR,data_var_select_v2,by="DOI")
Data_CR_COD_ANTI<-Data_CR_COD_ANTI %>% select(COD_PROD,DOI,COD_PROD_OR,COD_PROD_CIT,AUTHOR, ANO.x)

# ---------------------------------------------------------------------------------
# Construir a rede de citação
# Build the citation network
# ---------------------------------------------------------------------------------

lista_cod_or=unique(Data_CR_COD_ANTI$COD_PROD_OR)
data_cod_citacao_comum<- data.frame( 
                 COD_PROD_OR=character(),
                 r =list(),                   
                 stringsAsFactors=FALSE) 
count=1

for (j in seq(1,length(lista_cod_or),1)){
    
    if(!is.empty(lista_cod_or[j])){
        
        r <-which(lista_cod_or[j]==Data_CR_COD_ANTI$COD_PROD_OR)
    
        if (length(r)==0){ 
        }
        else{
           if (length(r)>=1){
            data_cod_citacao_comum[count,"COD_PROD_OR"]<-lista_cod_or[j]
            data_cod_citacao_comum$r[count]<-list(which(lista_cod_or[j]==Data_CR_COD_ANTI$COD_PROD_OR))
            count=count+1
           } 
        }    
    }
}

# ---------------------------------------------------------------------------------
# Rede na versão coluna
# Network in column version
# ---------------------------------------------------------------------------------

start_time <- Sys.time()
count<-1

var_select_doi=NULL
var_select_codprod=NULL
#var_select_cod_or=NULL
var_select_codcit=NULL
var_select_autor=NULL
var_select_ano=NULL

data_cocitacao<- data.frame( 
                 COD_PROD=character(),
                 stringsAsFactors=FALSE) 
        
       for (j in 1:nrow(data_cod_citacao_comum)){
         
        r <-unlist(data_cod_citacao_comum[j,2])  
                      
             var_select_codprod<-data_cod_citacao_comum[j,1]
             var_select_doi<-as.list(Data_CR_COD_ANTI[r,2])
             var_select_autor<-as.list(Data_CR_COD_ANTI[r,5])
             var_select_ano<-as.list(Data_CR_COD_ANTI[r,6])
             var_select_cod_or<-as.list(Data_CR_COD_ANTI[r,1])
             var_select_codcit<-as.list(Data_CR_COD_ANTI[r,4])
           
             data_cocitacao[count,"COD_PROD"]<-var_select_codprod  
                          
             for (d in 1:length(r)){
                coluna<-str_c("TAG_", d)
                coluna_para<-str_c("COD_CIT_",d)
                data_cocitacao[count,coluna_para]<-var_select_cod_or[d]
                data_cocitacao[count,coluna]<-str_c(var_select_autor[d],var_select_ano[d])                
            } 

           count= count+1  
    } 

end_time <- Sys.time()
end_time - start_time

# ---------------------------------------------------------------------------------
# Rede na versão linha
# Network in line version
# ---------------------------------------------------------------------------------

start_time <- Sys.time()
count<-1

var_select_doi=NULL
var_select_codprod=NULL
var_select_codcit=NULL
var_select_autor=NULL
var_select_ano=NULL

data_coc<- data.frame( 
                 COD_PROD=character(),
                 stringsAsFactors=FALSE) 
    
       for (j in 1:nrow(data_cod_citacao_comum)){
         
        r <-unlist(data_cod_citacao_comum[j,2])  
                      
             var_select_codprod<-data_cod_citacao_comum[j,1]
             var_select_cod_or<-as.list(Data_CR_COD_ANTI[r,1])
             var_select_codcit<-as.list(Data_CR_COD_ANTI[r,4])
             
             var_select_doi<-as.list(Data_CR_COD_ANTI[r,2])
             var_select_autor<-as.list(Data_CR_COD_ANTI[r,5])
             var_select_ano<-as.list(Data_CR_COD_ANTI[r,6])
                        
             for (d in 1:length(r)){

                data_coc[count,"COD_PROD"]<-var_select_codprod  
                data_coc[count,"PARA"]<-var_select_cod_or[d]
                data_coc[count,"TAG"]<-str_c(var_select_autor[d],var_select_ano[d]) 
                count= count+1                
            } 
    }

end_time <- Sys.time()
end_time - start_time

# ---------------------------------------------------------------------------------
# Salvar os dados
# Save the data
# ---------------------------------------------------------------------------------

# Os arquivos são salvos a partir da data/hora
# Files are saved from date/time
library(lubridate) 
tempo<- Sys.time()
tempo<-ymd_hms(tempo)
hora<-hour(tempo)
min<-minute(tempo)
dia<- Sys.Date()

nomearqxls<-paste0("/Users/hbbpereira/Downloads/OverviewCOVID19pandemic/database/biblioData_",dia,"_",hora,"_",min,".xlsx") 

wb <- createWorkbook()
addWorksheet(wb, sheetName = "planilha_original")
addWorksheet(wb, sheetName = "lista_artigos_excluidos")
addWorksheet(wb, sheetName = "lista_artigos_doiduplicado")
addWorksheet(wb, sheetName = "planilha_analisada")
addWorksheet(wb, sheetName = "vc_autor")
addWorksheet(wb, sheetName = "vc_titulo")
addWorksheet(wb, sheetName = "rede_coautoria")

writeData(wb, sheet = "planilha_original", x = data)
writeData(wb, sheet = "lista_artigos_excluidos", x = Data_excluidos)
writeData(wb, sheet = "lista_artigos_doiduplicado", x = data_doi_data_select)
writeData(wb, sheet = "planilha_analisada", x = DataPBL) 
writeData(wb, sheet = "vc_autor", x = DataDePara)
writeData(wb, sheet = "vc_titulo", x = DataTituloVC)
writeData(wb, sheet = "rede_coautoria", x = data_comb)
saveWorkbook(wb, file = nomearqxls)

nomearqxls<-paste0("/Users/hbbpereira/Downloads/OverviewCOVID19pandemic/database/citacao.xlsx") 

wb <- createWorkbook()
addWorksheet(wb, sheetName = "artigo_rede")
addWorksheet(wb, sheetName = "citacao_vc")
addWorksheet(wb, sheetName = "rede_citacao_completa")
addWorksheet(wb, sheetName = "rede_citacao_criarnet")

writeData(wb, sheet = "artigo_rede", x = data)
writeData(wb, sheet = "citacao_vc", x = DataCitacaovc)
writeData(wb, sheet = "rede_citacao_completa", x = data_cocitacao)
writeData(wb, sheet = "rede_citacao_criarnet", x = data_coc)
saveWorkbook(wb, file = nomearqxls)
