# ---------------------------------------------------------------------------------
# Calcular as eficiências global e local da rede de interesse
# Calculate the global and local efficiencies of the network of interest
#
# Autores: Hernane Borges de Barros Pereira, Ludmilla Monfort Oliveira Sousa, 
# Maíra Lima de Souza, and Marcelo A. Moret 
# Authors: Hernane Borges de Barros Pereira, Ludmilla Monfort Oliveira Sousa, 
# Maíra Lima de Souza, and Marcelo A. Moret                                                                   |
#
# Última atualização: 08/03/2023 
# Last update: March 08, 2023                            
# ---------------------------------------------------------------------------------

# ---------------------------------------------------------------------------------
# Instalação de pacotes para paralelizar a execução do script no caso de rede grande 
# Installation of packages to parallelize script execution in case of large networks
# https://github.com/cwatson/brainGraph/issues/15
# ---------------------------------------------------------------------------------
# instalar os pacotes apenas uma vez
# install packages only once
# caso os pacotes já tenham sido instalados instalado, comentar as linhas
# if packages have already been installed, comment the lines

install.packages('snow')
install.packages('doSNOW')

# ---------------------------------------------------------------------------------
# Instalacao de pacotes para trabalhar com redes 
# Installing packages to work with networks
# ---------------------------------------------------------------------------------

install.packages("igraph")
library(igraph)

install.packages("brainGraph")
library(brainGraph)

# ---------------------------------------------------------------------------------
# Executar esta parte para paralelizar
# Run this part of the script to parallelize
# ---------------------------------------------------------------------------------

OS <- .Platform$OS.type
if (OS == 'windows') {
  library(snow)
  library(doSNOW)
  num.cores <- as.numeric(Sys.getenv('NUMBER_OF_PROCESSORS'))
  cl <- makeCluster(num.cores, type='SOCK')
  clusterExport(cl, 'sim.rand.graph.par')   # Or whatever function you will use
  registerDoSNOW(cl)
} else {
  library(doMC)
  num.cores <- detectCores()
  registerDoMC(num.cores)
}  

# ---------------------------------------------------------------------------------
# Abrir arquivos de redes no formato Pajek 
# Open network files in Pajek format
# ---------------------------------------------------------------------------------

g <- read.graph(file.choose(), format = "pajek")

# ---------------------------------------------------------------------------------
# Calcular a eficiência global da rede de interesse
# Calculate the global efficiency of the network of interest
# ---------------------------------------------------------------------------------

# efficiency(g, type = "global", weights = NULL, use.parallel = FALSE, A = NULL)
# ou/or
# efficiency(g, type = "nodal", weights = NULL, use.parallel = FALSE, A = NULL)

enodal <- efficiency(g, type = "nodal", weights = NULL, use.parallel = TRUE, A = NULL)

mediaglobal<-mean(enodal)
mediaglobal

# ---------------------------------------------------------------------------------
# Calcular a eficiência local da rede de interesse
# Calculate the local efficiency of the network of interest
#
# Para grandes redes deve-se indicar como TRUE o comando use.parallel
# For large networks the use.parallel command must be set to TRUE
# ---------------------------------------------------------------------------------

elocal <- efficiency(g, type = "local", weights = NULL, use.parallel = TRUE, A = NULL)
#elocal <- efficiency(g, type = "local", weights = NULL, use.parallel = FALSE, A = NULL)

medialocal<-mean(elocal)
medialocal
