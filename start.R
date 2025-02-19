#Start

#Carregar pacotes

library(tidyverse)
library(openxlsx)

#Pará----

ri_para <- tibble(
  tematica = "-",
  indicador = "-",
  regiao = c("Pará", "Araguaia", "Baixo Amazonas", "Carajás", "Guajará", "Guamá", "Lago de Tucuruí",
             "Marajó", "Rio Caeté", "Rio Capim", "Tapajós", "Tocantins", "Xingu"),
  localidade = c("Pará", "Araguaia", "Baixo Amazonas", "Carajás", "Guajará", "Guamá", "Lago de Tucuruí",
                 "Marajó", "Rio Caeté", "Rio Capim", "Tapajós", "Tocantins", "Xingu"),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-",
  categoria4 = "-",
  categoria5 = "-"
)


#RI Araguaia----
ri_araguaia <- 
  tibble(
    tematica = "-",
    indicador = "-",
    regiao = "Araguaia",
    localidade = c(
      "Água Azul do Norte",
      "Bannach",
      "Conceição do Araguaia",
      "Cumaru do Norte",
      "Floresta do Araguaia",
      "Ourilândia do Norte",
      "Pau D'Arco",
      "Redenção",
      "Rio Maria",
      "Santa Maria das Barreiras",
      "Santana do Araguaia",
      "São Félix do Xingu",
      "Sapucaia",
      "Tucumã",
      "Xinguara"
    ),
    categoria1 = "-",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
#RI Baixo Amazonas----
ri_baixo_amazonas <- 
  tibble(
    tematica = "-",
    indicador = "-",
    regiao = "Baixo Amazonas",
    localidade = c(
      "Alenquer",
      "Almeirim",
      "Belterra",
      "Curuá",
      "Faro",
      "Juruti",
      "Mojuí dos Campos",
      "Monte Alegre",
      "Óbidos",
      "Oriximiná",
      "Prainha",
      "Santarém",
      "Terra Santa"
    ),
    categoria1 = "-",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )

#RI Carajás----
ri_carajas <- tibble(
  tematica = "-",
  indicador = "-",
  regiao = "Carajás",
  localidade = c(
    "Bom Jesus do Tocantins",
    "Brejo Grande do Araguaia",
    "Canaã dos Carajás",
    "Curionópolis",
    "Eldorado do Carajás",
    "Marabá",
    "Palestina do Pará",
    "Parauapebas",
    "Piçarra",
    "São Domingos do Araguaia",
    "São Geraldo do Araguaia",
    "São João do Araguaia"
  ),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-",
  categoria4 = "-",
  categoria5 = "-"
)

#RI Guajará----
ri_guajara <- tibble(
  tematica = "-",
  indicador = "-",
  regiao = "Guajará",
  localidade = c(
    "Ananindeua",
    "Belém",
    "Benevides",
    "Marituba",
    "Santa Bárbara do Pará"
  ),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-",
  categoria4 = "-",
  categoria5 = "-"
)

#RI Guamá----
ri_guama <- tibble(
  tematica = "-",
  indicador = "-",
  regiao = "Guamá",
  localidade = c(
    "Castanhal",
    "Colares",
    "Curuçá",
    "Igarapé-Açu",
    "Inhangapi",
    "Magalhães Barata",
    "Maracanã",
    "Marapanim",
    "Santa Izabel do Pará",
    "Santa Maria do Pará",
    "Santo Antônio do Tauá",
    "São Caetano de Odivelas",
    "São Domingos do Capim",
    "São Francisco do Pará",
    "São João da Ponta",
    "São Miguel do Guamá",
    "Terra Alta",
    "Vigia"
  ),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-",
  categoria4 = "-",
  categoria5 = "-"
)

#RI Lago de Tucuruí----
ri_lago_de_tucurui <- tibble(
  tematica = "-",
  indicador = "-",
  regiao = "Lago de Tucuruí",
  localidade = c(
    "Breu Branco",
    "Goianésia do Pará",
    "Itupiranga",
    "Jacundá",
    "Nova Ipixuna",
    "Novo Repartimento",
    "Tucuruí"
  ),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-",
  categoria4 = "-",
  categoria5 = "-"
)

#RI Marajó----
ri_marajo <- tibble(
  tematica = "-",
  indicador = "-",
  regiao = "Marajó",
  localidade = c(
    "Afuá",
    "Anajás",
    "Bagre",
    "Breves",
    "Cachoeira do Arari",
    "Chaves",
    "Curralinho",
    "Gurupá",
    "Melgaço",
    "Muaná",
    "Oeiras do Pará",
    "Ponta de Pedras",
    "Portel",
    "Salvaterra",
    "Santa Cruz do Arari",
    "São Sebastião da Boa Vista",
    "Soure"
  ),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-",
  categoria4 = "-",
  categoria5 = "-"
)

#RI Rio Caeté----
ri_rio_caete <- tibble(
  tematica = "-",
  indicador = "-",
  regiao = "Rio Caeté",
  localidade = c(
    "Augusto Corrêa",
    "Bonito",
    "Bragança",
    "Cachoeira do Piriá",
    "Capanema",
    "Nova Timboteua",
    "Peixe-Boi",
    "Primavera",
    "Quatipuru",
    "Salinópolis",
    "Santa Luzia do Pará",
    "Santarém Novo",
    "São João de Pirabas",
    "Tracuateua",
    "Viseu"
  ),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-",
  categoria4 = "-",
  categoria5 = "-"
)

#RI Rio Capim----
ri_rio_capim <- tibble(
  tematica = "-",
  indicador = "-",
  regiao = "Rio Capim",
  localidade = c(
    "Abel Figueiredo",
    "Aurora do Pará",
    "Bujaru",
    "Capitão Poço",
    "Concórdia do Pará",
    "Dom Eliseu",
    "Garrafão do Norte",
    "Ipixuna do Pará",
    "Irituia",
    "Mãe do Rio",
    "Nova Esperança do Piriá",
    "Ourém",
    "Paragominas",
    "Rondon do Pará",
    "Tomé-Açu",
    "Ulianópolis"
  ),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-",
  categoria4 = "-",
  categoria5 = "-"
)

#RI Tapajós----
ri_tapajos <- tibble(
  tematica = "-",
  indicador = "-",
  regiao = "Tapajós",
  localidade = c(
    "Aveiro",
    "Itaituba",
    "Jacareacanga",
    "Novo Progresso",
    "Rurópolis",
    "Trairão"
  ),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-",
  categoria4 = "-",
  categoria5 = "-"
)

#RI Tocantins----
ri_tocantins <- tibble(
  tematica = "-",
  indicador = "-",
  regiao = "Tocantins",
  localidade = c(
    "Abaetetuba",
    "Acará",
    "Baião",
    "Barcarena",
    "Cametá",
    "Igarapé-Miri",
    "Limoeiro do Ajuru",
    "Mocajuba",
    "Moju",
    "Tailândia"
  ),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-",
  categoria4 = "-",
  categoria5 = "-"
)

#RI Xingu----
ri_xingu <- tibble(
  tematica = "-",
  indicador = "-",
  regiao = "Xingu",
  localidade = c(
    "Altamira",
    "Anapu",
    "Brasil Novo",
    "Medicilândia",
    "Pacajá",
    "Placas",
    "Porto de Moz",
    "Senador José Porfírio",
    "Uruará",
    "Vitória do Xingu"
  ),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-",
  categoria4 = "-",
  categoria5 = "-"
)
