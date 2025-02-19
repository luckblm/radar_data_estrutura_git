# PARÁ----
bases <- ri_para
# DEMOGRAFIA----
## Tabela 1 - População, Área Territorial (km²) e Densidade Demográfica----
d1 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População, Área Territorial (km²) e Densidade Demográfica",
    categoria1 = "População Estimada Total",
    categoria2 = "Estimada Total",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População, Área Territorial (km²) e Densidade Demográfica",
    categoria1 = "Área Territorial  em km²",
    categoria2 = "Área",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População, Área Territorial (km²) e Densidade Demográfica",
    categoria1 = "Densidade Demográfica",
    categoria2 = "Densidade",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 1 - População, Área Territorial km2 e Densidade Demográfica.xlsx",
  overwrite = TRUE
)
## Tabela 2 - População por Sexo e Razão entre os Sexos----
d1 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Sexo e Razão entre os Sexos",
    categoria1 = "Masculino",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Sexo e Razão entre os Sexos",
    categoria1 = "Feminino",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Sexo e Razão entre os Sexos",
    categoria1 = "Razão de Sexos",
    categoria2 = "Razão",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb,
  "para/Tabela 2 - População por Sexo e Razão entre os Sexos.xlsx",
  overwrite = TRUE
)
## Tabela 3 - População por Faixa Etária----
d1 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Faixa Etária",
    categoria1 = "0 a 4 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Faixa Etária",
    categoria1 = "5 a 9 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Faixa Etária",
    categoria1 = "10 a 14 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Faixa Etária",
    categoria1 = "15 a 19 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Faixa Etária",
    categoria1 = "20 a 29 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Faixa Etária",
    categoria1 = "30 a 39 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Faixa Etária",
    categoria1 = "40 a 49 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Faixa Etária",
    categoria1 = "50 a 59 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Faixa Etária",
    categoria1 = "60 a 69 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Faixa Etária",
    categoria1 = "70 a 79 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d11 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População por Faixa Etária",
    categoria1 = "80 anos e mais",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb, "para/Tabela 3 - População por Faixa Etária.xlsx", overwrite = TRUE)
## Tabela 4 - Proporção de Idosos, Razão de Dependência e Índice de Envelhecimento----
d1 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População Total e Estrutura Etária usada no cálculo dos Indicadores: Proporção de Idosos, Razão de Dependência e Índice de Envelhecimento",
    categoria1 = "População Total",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )

d2 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População Total e Estrutura Etária usada no cálculo dos Indicadores: Proporção de Idosos, Razão de Dependência e Índice de Envelhecimento",
    categoria1 = "População menor que 15 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População Total e Estrutura Etária usada no cálculo dos Indicadores: Proporção de Idosos, Razão de Dependência e Índice de Envelhecimento",
    categoria1 = "População entre 15 e 64 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População Total e Estrutura Etária usada no cálculo dos Indicadores: Proporção de Idosos, Razão de Dependência e Índice de Envelhecimento",
    categoria1 = "População maior que 64 anos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População Total e Estrutura Etária usada no cálculo dos Indicadores: Proporção de Idosos, Razão de Dependência e Índice de Envelhecimento",
    categoria1 = "Proporção de Idosos",
    categoria2 = "População",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População Total e Estrutura Etária usada no cálculo dos Indicadores: Proporção de Idosos, Razão de Dependência e Índice de Envelhecimento",
    categoria1 = "Razão de Dependência",
    categoria2 = "Razão",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Demografia",
    indicador = "População Total e Estrutura Etária usada no cálculo dos Indicadores: Proporção de Idosos, Razão de Dependência e Índice de Envelhecimento",
    categoria1 = "Índice de Envelhecimento",
    categoria2 = "Índice",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 4 - Proporção de Idosos, Razão de Dependência e Índice de Envelhecimento.xlsx",
  overwrite = TRUE
)
# EDUCAÇÃO----
## Tabela 5 - Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa----
d1 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Creche",
    categoria2 = "Federal",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Creche",
    categoria2 = "Estadual",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Creche",
    categoria2 = "Municipal",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Creche",
    categoria2 = "Privado",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Creche",
    categoria2 = "Total",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Pré-Escola",
    categoria2 = "Federal",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Pré-Escola",
    categoria2 = "Estadual",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Pré-Escola",
    categoria2 = "Municipal",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Pré-Escola",
    categoria2 = "Privado",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Pré-Escola",
    categoria2 = "Total",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d11 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Federal",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d12 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Estadual",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d13 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Municipal",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d14 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Privado",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d15 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Total",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d16 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Federal",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d17 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Estadual",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d18 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Municipal",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d19 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Privado",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )
d20 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Total",
    categoria3 = "Número de Matrículas",
    categoria4 = "-",
    categoria5 = "-"
  )

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12,
  "d13" = d13,
  "d14" = d14,
  "d15" = d15,
  "d16" = d16,
  "d17" = d17,
  "d18" = d18,
  "d19" = d19,
  "d20" = d20
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 5 - Número de Matrículas nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa.xlsx",
  overwrite = TRUE
)
## Tabela 6 - Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa----
d1 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Creche",
    categoria2 = "Federal",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Creche",
    categoria2 = "Estadual",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Creche",
    categoria2 = "Municipal",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Creche",
    categoria2 = "Privado",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Creche",
    categoria2 = "Total",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino Pré-Escolar",
    categoria2 = "Federal",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino Pré-Escolar",
    categoria2 = "Estadual",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino Pré-Escolar",
    categoria2 = "Municipal",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino Pré-Escolar",
    categoria2 = "Privado",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino Pré-Escolar",
    categoria2 = "Total",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d11 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Federal",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d12 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Estadual",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d13 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Municipal",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d14 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Privado",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d15 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Total",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d16 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Federal",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d17 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Estadual",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d18 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Municipal",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d19 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Privado",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
d20 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Total",
    categoria3 = "Número de Estabelecimentos",
    categoria4 = "-",
    categoria5 = "-"
  )
# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12,
  "d13" = d13,
  "d14" = d14,
  "d15" = d15,
  "d16" = d16,
  "d17" = d17,
  "d18" = d18,
  "d19" = d19,
  "d20" = d20
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 6 - Estabelecimentos de Ensino Fundamental e Médio por Dependência Administrativa.xlsx",
  overwrite = TRUE
)
## Tabela 7 - Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa----
d1 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Creche",
    categoria2 = "Federal",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Creche",
    categoria2 = "Estadual",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Creche",
    categoria2 = "Municipal",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Creche",
    categoria2 = "Privado",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Creche",
    categoria2 = "Total",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Pré-Escola",
    categoria2 = "Federal",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Pré-Escola",
    categoria2 = "Estadual",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Pré-Escola",
    categoria2 = "Municipal",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Pré-Escola",
    categoria2 = "Privado",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Pré-Escola",
    categoria2 = "Total",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d11 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Federal",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d12 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Estadual",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d13 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Municipal",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d14 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Privado",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d15 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Total",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d16 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Federal",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d17 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Estadual",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d18 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Municipal",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d19 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Privado",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )
d20 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa",
    categoria1 = "Ensino médio",
    categoria2 = "Total",
    categoria3 = "Número de Docentes",
    categoria4 = "-",
    categoria5 = "-"
  )

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12,
  "d13" = d13,
  "d14" = d14,
  "d15" = d15,
  "d16" = d16,
  "d17" = d17,
  "d18" = d18,
  "d19" = d19,
  "d20" = d20
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 7 - Número de Docentes nos Ensinos Pré-escolar, Fundamental e Médio por Esfera Administrativa.xlsx",
  overwrite = TRUE
)
## Tabela 8 - Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa----
d1 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Aprovação",
    categoria2 = "Federal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Aprovação",
    categoria2 = "Estadual",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Aprovação",
    categoria2 = "Municipal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Aprovação",
    categoria2 = "Privada",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Aprovação",
    categoria2 = "Aprovação Total",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Reprovação",
    categoria2 = "Federal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Reprovação",
    categoria2 = "Estadual",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Reprovação",
    categoria2 = "Municipal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Reprovação",
    categoria2 = "Privada",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Reprovação",
    categoria2 = "Aprovação Total",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d11 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Abandono",
    categoria2 = "Federal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d12 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Abandono",
    categoria2 = "Estadual",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d13 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Abandono",
    categoria2 = "Municipal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d14 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Abandono",
    categoria2 = "Privada",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d15 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa",
    categoria1 = "Abandono",
    categoria2 = "Aprovação Total",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12,
  "d13" = d13,
  "d14" = d14,
  "d15" = d15
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 8 - Taxas de Aprovação, Reprovação e Evasão no Ensino Fundamental por Esfera Administrativa.xlsx",
  overwrite = TRUE
)
## Tabela 9 - Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa----
d1 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Aprovação",
    categoria2 = "Federal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Aprovação",
    categoria2 = "Estadual",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Aprovação",
    categoria2 = "Municipal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Aprovação",
    categoria2 = "Privada",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Aprovação",
    categoria2 = "Aprovação Total",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Reprovação",
    categoria2 = "Federal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Reprovação",
    categoria2 = "Estadual",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Reprovação",
    categoria2 = "Municipal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Reprovação",
    categoria2 = "Privada",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Reprovação",
    categoria2 = "Aprovação Total",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d11 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Abandono",
    categoria2 = "Federal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d12 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Abandono",
    categoria2 = "Estadual",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d13 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Abandono",
    categoria2 = "Municipal",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d14 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Abandono",
    categoria2 = "Privada",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )
d15 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa",
    categoria1 = "Abandono",
    categoria2 = "Aprovação Total",
    categoria3 = "Taxas",
    categoria4 = "-",
    categoria5 = "-"
  )



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12,
  "d13" = d13,
  "d14" = d14,
  "d15" = d15
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 9 - Taxas de Aprovação, Reprovação e Evasão no Ensino Médio por Esfera Administrativa.xlsx",
  overwrite = TRUE
)
## Tabela 10 - Distorção Idade-Série Total por Nível de Ensino----
d1 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Distorção idade-série total por nível de Ensino",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Distorção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Distorção idade-série total por nível de Ensino",
    categoria1 = "Ensino Médio",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb,
  "para/Tabela 10 - Distorção Idade-Série Total por Nível de Ensino.xlsx",
  overwrite = TRUE
)
## Tabela 11 - Índice de Desenvolvimento da Educação Básica - IDEB (Escola Pública)----
d1 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Índice de Desenvolvimento da Educação Básica - IDEB (Escola Pública)",
    categoria1 = "Nota IDEB (Escola Pública)",
    categoria2 = "Séries Iniciais 5º Ano",
    categoria3 = "Índice de Desenvonvimento",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Índice de Desenvolvimento da Educação Básica - IDEB (Escola Pública)",
    categoria1 = "Nota IDEB (Escola Pública)",
    categoria2 = "Séries Finais 9º Ano",
    categoria3 = "Índice de Desenvonvimento",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 11 - Índice de Desenvolvimento da Educação Básica - IDEB_Escola Pública.xlsx",
  overwrite = TRUE
)
## Tabela 12 - Média de Alunos por Nível de Ensino----
d1 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Média de Alunos por Nível de Ensino",
    categoria1 = "Creche",
    categoria2 = "Média de Alunos",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Média de Alunos por Nível de Ensino",
    categoria1 = "Pré-Escola",
    categoria2 = "Média de Alunos",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Média de Alunos por Nível de Ensino",
    categoria1 = "Ensino Fundamental",
    categoria2 = "Média de Alunos",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Educação",
    indicador = "Média de Alunos por Nível de Ensino",
    categoria1 = "Ensino Médio",
    categoria2 = "Média de Alunos",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb,
  "para/Tabela 12 - Média de Alunos por Nível de Ensino.xlsx",
  overwrite = TRUE
)
# SAÚDE----
## Tabela 13 - Taxas de Mortalidade Infantil, Mortalidade em Menores que 05 Anos e Mortalidade Materna----
d1 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxas de Mortalidade Infantil, Mortalidade em Menores que 05 Anos e Mortalidade Materna",
    categoria1 = "Taxa de Mortalidade Infantil",
    categoria2 = "Taxas",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxas de Mortalidade Infantil, Mortalidade em Menores que 05 Anos e Mortalidade Materna",
    categoria1 = "Taxa de Mortalidade em Menores que 05 Anos",
    categoria2 = "Taxas",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxas de Mortalidade Infantil, Mortalidade em Menores que 05 Anos e Mortalidade Materna",
    categoria1 = "Taxa de Mortalidade Materna",
    categoria2 = "Taxas",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 13 - Taxas de Mortalidade Infantil, Mortalidade em Menores que 05 Anos e Mortalidade Materna.xlsx",
  overwrite = TRUE
)
## Tabela 14 - Taxas de Mortalidade Geral e Percentual de Mortes por sexo----
d1 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxas de Mortalidade Geral e Percentual de Mortes por sexo",
    categoria1 = "Total de Óbitos",
    categoria2 = "Total",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxas de Mortalidade Geral e Percentual de Mortes por sexo",
    categoria1 = "Taxa de Mortalidade Geral",
    categoria2 = "Taxa",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxas de Mortalidade Geral e Percentual de Mortes por sexo",
    categoria1 = "% Morte por Sexo",
    categoria2 = "Masculino",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxas de Mortalidade Geral e Percentual de Mortes por sexo",
    categoria1 = "% Morte por Sexo",
    categoria2 = "Feminino",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxas de Mortalidade Geral e Percentual de Mortes por sexo",
    categoria1 = "% Morte por Sexo",
    categoria2 = "Ignorado",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 14 - Taxas de Mortalidade Geral e Percentual de Mortes por sexo.xlsx",
  overwrite = TRUE
)
## Tabela 15 - Taxa de Natalidade e Percentual de Nascidos Vivos por tipo de Parto----
d1 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxa de Natalidade e Percentual de Nascidos Vivos por tipo de Parto",
    categoria1 = "Taxa de Natalidade	",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxa de Natalidade e Percentual de Nascidos Vivos por tipo de Parto",
    categoria1 = "% Tipo de Parto",
    categoria2 = "Vaginal",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxa de Natalidade e Percentual de Nascidos Vivos por tipo de Parto",
    categoria1 = "% Tipo de Parto",
    categoria2 = "Cesáreo",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Taxa de Natalidade e Percentual de Nascidos Vivos por tipo de Parto",
    categoria1 = "% Tipo de Parto",
    categoria2 = "Ignorado",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 15 - Taxa de Natalidade e Percentual de Nascidos Vivos por tipo de Parto.xlsx",
  overwrite = TRUE
)
## Tabela 16 - Percentual de Nascidos Vivos Conforme o Número de Consultas Pré-Natal----
d1 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme o Número de Consultas Pré-Natal",
    categoria1 = "Taxa de Natalidade",
    categoria2 = "Taxa",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme o Número de Consultas Pré-Natal",
    categoria1 = "% Tipo de Parto",
    categoria2 = "Vaginal",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme o Número de Consultas Pré-Natal",
    categoria1 = "% Tipo de Parto",
    categoria2 = "Cesáreo",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme o Número de Consultas Pré-Natal",
    categoria1 = "% Tipo de Parto",
    categoria2 = "Ignorado",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 16 - Percentual de Nascidos Vivos Conforme o Número de Consultas Pré-Natal.xlsx",
  overwrite = TRUE
)
## Tabela 17 - Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos----
d1 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos",
    categoria1 = "Nascidos Vivos",
    categoria2 = "Quantidade",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos",
    categoria1 = "Percentual de Nascidos Vivos por Idade da Mãe",
    categoria2 = " 10 a 14 anos",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos",
    categoria1 = "Percentual de Nascidos Vivos por Idade da Mãe",
    categoria2 = "15 a 19 anos",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos",
    categoria1 = "Percentual de Nascidos Vivos por Idade da Mãe",
    categoria2 = "20 a 24 anos",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos",
    categoria1 = "Percentual de Nascidos Vivos por Idade da Mãe",
    categoria2 = "25 a 29 anos",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos",
    categoria1 = "Percentual de Nascidos Vivos por Idade da Mãe",
    categoria2 = "30 a 34 anos",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos",
    categoria1 = "Percentual de Nascidos Vivos por Idade da Mãe",
    categoria2 = "35 a 39 anos",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos",
    categoria1 = "Percentual de Nascidos Vivos por Idade da Mãe",
    categoria2 = "40 ou mais",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos",
    categoria1 = "Percentual de Nascidos Vivos por Idade da Mãe",
    categoria2 = "Idade ignorada",
    categoria3 = "Percentual",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos",
    categoria1 = "Proporção de Mulheres de 25 a 64 Anos que Realizaram Exame Citopatológico",
    categoria2 = "Proporção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 17 - Percentual de Nascidos Vivos Conforme a Faixa Etária da Mãe e Razão de Exames Citopatológicos do Colo do Útero em Mulheres de 25 a 64 anos.xlsx",
  overwrite = TRUE
)
## Tabela 18 - Óbitos segundo Principais Causas (CID-10)----
d1 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Óbitos segundo Principais Causas (CID-10)",
    categoria1 = "Óbitos Totais",
    categoria2 = "Quantidade",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Óbitos segundo Principais Causas (CID-10)",
    categoria1 = "Principais Causas de Óbitos",
    categoria2 = "Doenças Infeciocciosas e Parasitárias",
    categoria3 = "Quantidade",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Óbitos segundo Principais Causas (CID-10)",
    categoria1 = "Principais Causas de Óbitos",
    categoria2 = "Neoplasias",
    categoria3 = "Quantidade",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Óbitos segundo Principais Causas (CID-10)",
    categoria1 = "Principais Causas de Óbitos",
    categoria2 = "Sistema Nervoso",
    categoria3 = "Quantidade",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Óbitos segundo Principais Causas (CID-10)",
    categoria1 = "Principais Causas de Óbitos",
    categoria2 = "Aparelho Circulatório",
    categoria3 = "Quantidade",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Óbitos segundo Principais Causas (CID-10)",
    categoria1 = "Principais Causas de Óbitos",
    categoria2 = "Aparelho Respiratório",
    categoria3 = "Quantidade",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Óbitos segundo Principais Causas (CID-10)",
    categoria1 = "Principais Causas de Óbitos",
    categoria2 = "Aparelho Digestivo",
    categoria3 = "Quantidade",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Óbitos segundo Principais Causas (CID-10)",
    categoria1 = "Principais Causas de Óbitos",
    categoria2 = "Total",
    categoria3 = "Quantidade",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb,
  "para/Tabela 18 - Óbitos segundo Principais Causas CID_10.xlsx",
  overwrite = TRUE
)
## Tabela 19 - Caracterização Hospitalar por Tipo----
d1 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = " Caracterização Hospitalar por  Tipo",
    categoria1 = "Nome do Hospital",
    categoria2 = "Tipo",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb,
  "para/Tabela 19 - Caracterização Hospitalar por Tipo.xlsx",
  overwrite = TRUE
)
## Tabela 20 - Médicos e Profissionais de Ensino Superior na Área da Saúde----
d1 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Médicos",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Assistente Social",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Bioquímico/ farmacêutico",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Cirurgião Geral",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Clínico Geral",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Enfermeiro",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Fisioterapeuta",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Fonoaudiólogo",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Gineco Obstetra",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Médico de Família",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d11 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Nutricionista",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d12 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Odontólogo",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d13 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Pediatra",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d14 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Psicólogo",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d15 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Psiquiatra",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d16 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Radiologista",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d17 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Sanitarista",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d18 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Outras especialidades médicas",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d19 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Médicos e Profissionais de Ensino Superior na Área da Saúde",
    categoria1 = "Profissionais de Ensino Superior na Área da Saúde",
    categoria2 = "Outras ocupações de nível superior relac à Saúde",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d8,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12,
  "d13" = d13,
  "d14" = d14,
  "d15" = d15,
  "d16" = d16,
  "d17" = d17,
  "d18" = d18,
  "d19" = d19
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 20 - Médicos e Profissionais de Ensino Superior na Área da Saúde.xlsx",
  overwrite = TRUE
)
## Tabela 21 – Caracterização Leitos Existentes----
d1 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização dos Leitos Existentes",
    categoria1 = "Leitos Hospitalares de Internação	",
    categoria2 = " Quantidade SUS",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização dos Leitos Existentes",
    categoria1 = "Leitos Hospitalares de Internação	",
    categoria2 = " Quantidade Não SUS",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização dos Leitos Existentes",
    categoria1 = "Leitos Hospitalares de Internação	",
    categoria2 = "Total",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização dos Leitos Existentes",
    categoria1 = "Leitos Hospitalares Complementares",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb,
  "para/Tabela 21 – Caracterização Leitos Existentes.xlsx",
  overwrite = TRUE
)
## Tabela 22 – Caracterização Hospitalar – Equipamentos de diagnósticos por imagem e infra estrutura----
d1 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Equipamentos de Manutenção da Vida",
    categoria2 = "Desfibrilador",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Equipamentos de Manutenção da Vida",
    categoria2 = "Desfibrilador",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Equipamentos de Manutenção da Vida",
    categoria2 = "Monitor de ECG",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Equipamentos de Manutenção da Vida",
    categoria2 = "Monitor de ECG",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Equipamentos de Manutenção da Vida",
    categoria2 = "Respirador/ Ventilador",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Equipamentos de Manutenção da Vida",
    categoria2 = "Respirador/ Ventilador",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Imagem",
    categoria2 = "Mamógrafo com Comando Simples",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Imagem",
    categoria2 = "Mamógrafo com Comando Simples",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Imagem",
    categoria2 = "Raio X de 100 a 500 MA",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Imagem",
    categoria2 = "Raio X de 100 a 500 MA",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d11 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Imagem",
    categoria2 = "Tomógrafo Computadorizado",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d12 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Imagem",
    categoria2 = "Tomógrafo Computadorizado",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d13 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Imagem",
    categoria2 = "Ultrassom Convencional",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d14 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Imagem",
    categoria2 = "Ultrassom Convencional",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d15 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Método Gráfico",
    categoria2 = "Eletrocardiografo",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d16 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Método Gráfico",
    categoria2 = "Eletrocardiografo",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d17 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Método Gráfico",
    categoria2 = "Eletroencefalografo",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d18 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Diagnóstico por Método Gráfico",
    categoria2 = "Eletroencefalografo",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d19 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Infraestrutura",
    categoria2 = "Controlo Ambiental/ Ar-Condicionado Central",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d20 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Infraestrutura",
    categoria2 = "Controlo Ambiental/ Ar-Condicionado Central",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d21 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Infraestrutura",
    categoria2 = "Grupo Gerador",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d22 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Infraestrutura",
    categoria2 = "Grupo Gerador",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d23 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Infraestrutura",
    categoria2 = "Usina de Oxigênio",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d24 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Infraestrutura",
    categoria2 = "Usina de Oxigênio",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )
d25 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Total",
    categoria2 = "-",
    categoria3 = "Existente",
    categoria4 = "-",
    categoria5 = "-"
  )
d26 <- bases %>%
  transform(
    tematica = "Saúde",
    indicador = "Caracterização Hospitalar – Equipamentos de Manutenção da Vida, Equipamentos de Diagnóstico e Infraestrutura",
    categoria1 = "Total",
    categoria2 = "-",
    categoria3 = "Em Uso",
    categoria4 = "-",
    categoria5 = "-"
  )



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12,
  "d13" = d13,
  "d14" = d14,
  "d15" = d15,
  "d16" = d16,
  "d17" = d17,
  "d18" = d18,
  "d19" = d19,
  "d20" = d20,
  "d21" = d21,
  "d22" = d22,
  "d23" = d23,
  "d24" = d24,
  "d25" = d25,
  "d26" = d26
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 22 – Caracterização Hospitalar – Equipamentos de diagnósticos por imagem e infra estrutura.xlsx",
  overwrite = TRUE
)
# MERCADO DE TRABALHO----
## Tabela 23 - Vínculos Empregatícios Total e por Sexo no Emprego Formal----
d1 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios Total e por Sexo no Emprego Formal",
    categoria1 = "Total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios Total e por Sexo no Emprego Formal",
    categoria1 = "Sexo",
    categoria2 = "Masculino",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios Total e por Sexo no Emprego Formal",
    categoria1 = "Sexo",
    categoria2 = "Feminino",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 23 - Vínculos Empregatícios Total e por Sexo no Emprego Formal.xlsx",
  overwrite = TRUE
)
## Tabela 24 - Vínculos Empregatícios no Emprego Formal por Grande Setor (IBGE)----
d1 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Grande Setor (IBGE)",
    categoria1 = "Total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Grande Setor (IBGE)",
    categoria1 = "Grande Setor (IBGE)",
    categoria2 = "Indústria",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Grande Setor (IBGE)",
    categoria1 = "Grande Setor (IBGE)",
    categoria2 = "Construção Civil",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Grande Setor (IBGE)",
    categoria1 = "Grande Setor (IBGE)",
    categoria2 = "Comércio",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Grande Setor (IBGE)",
    categoria1 = "Grande Setor (IBGE)",
    categoria2 = "Serviços",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Grande Setor (IBGE)",
    categoria1 = "Grande Setor (IBGE)",
    categoria2 = "Agropecuária",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 24 - Vínculos Empregatícios no Emprego Formal por Grande Setor_IBGE.xlsx",
  overwrite = TRUE
)
## Tabela 25 - Vínculos Empregatícios no Emprego Formal por Setor Econômico----
d1 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Setor Econômico",
    categoria1 = "Total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Setor Econômico",
    categoria1 = "Setor Econômico",
    categoria2 = "Extrativa mineral",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Setor Econômico",
    categoria1 = "Setor Econômico",
    categoria2 = "Indústria de transformação",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Setor Econômico",
    categoria1 = "Setor Econômico",
    categoria2 = "Servicos industriais de utilidade pública",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Setor Econômico",
    categoria1 = "Setor Econômico",
    categoria2 = "Construção Civil",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Setor Econômico",
    categoria1 = "Setor Econômico",
    categoria2 = "Comércio",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Setor Econômico",
    categoria1 = "Setor Econômico",
    categoria2 = "Serviços",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Setor Econômico",
    categoria1 = "Setor Econômico",
    categoria2 = "Administração Pública",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios no Emprego Formal por Setor Econômico",
    categoria1 = "Setor Econômico",
    categoria2 = "Agropecuária, extração vegetal, caça e pesca",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 25 - Vínculos Empregatícios no Emprego Formal por Setor Econômico.xlsx",
  overwrite = TRUE
)
## Tabela 26 - Vínculos Empregatícios por Escolaridade do Trabalhador Formal----
d1 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios por Escolaridade do Trabalhador Formal",
    categoria1 = "Total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios por Escolaridade do Trabalhador Formal",
    categoria1 = "Escolaridade",
    categoria2 = "Analfabeto",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios por Escolaridade do Trabalhador Formal",
    categoria1 = "Escolaridade",
    categoria2 = "Até 5ª Incompleto",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios por Escolaridade do Trabalhador Formal",
    categoria1 = "Escolaridade",
    categoria2 = "5ª Completo Fundamental",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios por Escolaridade do Trabalhador Formal",
    categoria1 = "Escolaridade",
    categoria2 = "6ª a 9ª Fundamental",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios por Escolaridade do Trabalhador Formal",
    categoria1 = "Escolaridade",
    categoria2 = "Fundamental Completo",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios por Escolaridade do Trabalhador Formal",
    categoria1 = "Escolaridade",
    categoria2 = "Médio Incompleto",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios por Escolaridade do Trabalhador Formal",
    categoria1 = "Escolaridade",
    categoria2 = "Médio Completo",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios por Escolaridade do Trabalhador Formal",
    categoria1 = "Escolaridade",
    categoria2 = "Superior Incompleto",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Vínculos Empregatícios por Escolaridade do Trabalhador Formal",
    categoria1 = "Escolaridade",
    categoria2 = "Superior Completo",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 26 - Vínculos Empregatícios por Escolaridade do Trabalhador Formal.xlsx",
  overwrite = TRUE
)
## Tabela 27 - Remuneração Média (R$) Total e por Sexo do Trabalhador Formal----
d1 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Remuneração Média (R$) Total e por Sexo do Trabalhador Formal",
    categoria1 = "Total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Remuneração Média (R$) Total e por Sexo do Trabalhador Formal",
    categoria1 = "Sexo",
    categoria2 = "Masculino",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Mercado de Trabalho",
    indicador = "Remuneração Média (R$) Total e por Sexo do Trabalhador Formal",
    categoria1 = "Sexo",
    categoria2 = "Feminino",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 27 - Remuneração Média (R$) Total e por Sexo do Trabalhador Formal.xlsx",
  overwrite = TRUE
)
# ASSISTÊNCIA E PREVIDÊNCIA SOCIAL----
## Tabela 28 - Informações de Famílias Inscritas no Cadastro Único para Programas Sociais (CadÚnico) e Bolsa Família----
d1 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Informações de Famílias Inscritas no Cadastro Único para Programas Sociais (CadÚnico) e Bolsa Família",
    categoria1 = "CadÚnico*",
    categoria2 = "Famílias inscritas	",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Informações de Famílias Inscritas no Cadastro Único para Programas Sociais (CadÚnico) e Bolsa Família",
    categoria1 = "CadÚnico*",
    categoria2 = "Famílias inscritas com rendimento familiar per capita de até 1/2 salário mínimo",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Informações de Famílias Inscritas no Cadastro Único para Programas Sociais (CadÚnico) e Bolsa Família",
    categoria1 = "Auxílio Brasil**",
    categoria2 = "Famílias",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Informações de Famílias Inscritas no Cadastro Único para Programas Sociais (CadÚnico) e Bolsa Família",
    categoria1 = "Auxílio Brasil**",
    categoria2 = "Valor Total (R$)",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Informações de Famílias Inscritas no Cadastro Único para Programas Sociais (CadÚnico) e Bolsa Família",
    categoria1 = "CadÚnico***",
    categoria2 = "Famílias inscritas	",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Informações de Famílias Inscritas no Cadastro Único para Programas Sociais (CadÚnico) e Bolsa Família",
    categoria1 = "CadÚnico*",
    categoria2 = "Famílias inscritas com rendimento familiar per capita de até 1/2 salário mínimo",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Informações de Famílias Inscritas no Cadastro Único para Programas Sociais (CadÚnico) e Bolsa Família",
    categoria1 = "Bolsa Família****",
    categoria2 = "Famílias",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Informações de Famílias Inscritas no Cadastro Único para Programas Sociais (CadÚnico) e Bolsa Família",
    categoria1 = "Bolsa Família**",
    categoria2 = "Valor Total (R$)",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 28 - Informações de Famílias Inscritas no Cadastro Único para Programas Sociais_CadÚnico e Bolsa Família.xlsx",
  overwrite = TRUE
)
## Tabela 29 - Arrecadação e Benefícios Emitidos pela Previdência Social----
d1 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Benefícios Emitidos pela Previdência Social",
    categoria1 = "Quantidade de benefícios emitidos no mês de dezembro",
    categoria2 = "Total",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Benefícios Emitidos pela Previdência Social",
    categoria1 = "Quantidade de benefícios emitidos no mês de dezembro",
    categoria2 = "Urbano",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Benefícios Emitidos pela Previdência Social",
    categoria1 = "Quantidade de benefícios emitidos no mês de dezembro",
    categoria2 = "Rural",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Benefícios Emitidos pela Previdência Social",
    categoria1 = "Valor dos benefícios emitidos no mês de dezembro",
    categoria2 = "Total",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Benefícios Emitidos pela Previdência Social",
    categoria1 = "Valor dos benefícios emitidos no mês de dezembro",
    categoria2 = "Urbano",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Benefícios Emitidos pela Previdência Social",
    localidade = "Valor dos benefícios emitidos no ano",
    categoria1 = "Valor dos benefícios emitidos no mês de dezembro",
    categoria2 = "Rural",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Benefícios Emitidos pela Previdência Social",
    categoria1 = "Valor dos benefícios emitidos no ano",
    categoria2 = "Total",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Benefícios Emitidos pela Previdência Social",
    categoria1 = "Valor dos benefícios emitidos no ano",
    categoria2 = "Urbano",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Assistência e Previdência Social",
    indicador = "Benefícios Emitidos pela Previdência Social",
    categoria1 = "Valor dos benefícios emitidos no ano",
    categoria2 = "Rural",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 29 - Arrecadação e Benefícios Emitidos pela Previdência Social.xlsx",
  overwrite = TRUE
)
# SEGURANÇA----
## Tabela 30 - Número de Óbitos por Agressão, População e Taxa de Homicídio Total----
d1 <- bases %>%
  transform(
    tematica = "Segurança",
    indicador = "Número de Óbitos por Agressão, População e Taxa de Homicídio Total",
    categoria1 = "Óbitos por Agressões",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Segurança",
    indicador = "Número de Óbitos por Agressão, População e Taxa de Homicídio Total",
    categoria1 = "População Total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Segurança",
    indicador = "Número de Óbitos por Agressão, População e Taxa de Homicídio Total",
    categoria1 = "Taxa de Homicídio (100 Mil habitantes)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 30 - Número de Óbitos por Agressão, População e Taxa de Homicídio Total.xlsx",
  overwrite = TRUE
)
## Tabela 31 - Número de Óbitos de Jovens de 15 a 29 anos por Agressão, População Jovem e Taxa de Homicídio de Jovens----
d1 <- bases %>%
  transform(
    tematica = "Segurança",
    indicador = "Número de Óbitos de Jovens de 15 a 29 anos por Agressão, População Jovem e Taxa de Homicídio de Jovens",
    categoria1 = "Óbitos de Jovens por Agressão",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Segurança",
    indicador = "Número de Óbitos de Jovens de 15 a 29 anos por Agressão, População Jovem e Taxa de Homicídio de Jovens",
    categoria1 = "População Jovem de 15 a 29 anos",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Segurança",
    indicador = "Número de Óbitos de Jovens de 15 a 29 anos por Agressão, População Jovem e Taxa de Homicídio de Jovens",
    categoria1 = "Taxa de Homicídio de Jovens (100.000 hab)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 31 - Número de Óbitos de Jovens de 15 a 29 anos por Agressão, População Jovem e Taxa de Homicídio de Jovens.xlsx",
  overwrite = TRUE
)
## Tabela 32 - Número de Óbitos por Acidente de Trânsito, População e Taxa de Homicídio no Trânsito----
d1 <- bases %>%
  transform(
    tematica = "Segurança",
    indicador = "Número de Óbitos por Acidente de Trânsito, População e Taxa de Homicídio no Trânsito",
    categoria1 = "Óbitos por Acidentes de Trânsito",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Segurança",
    indicador = "Número de Óbitos por Acidente de Trânsito, População e Taxa de Homicídio no Trânsito",
    categoria1 = "População Total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Segurança",
    indicador = "Número de Óbitos por Acidente de Trânsito, População e Taxa de Homicídio no Trânsito",
    categoria1 = "Taxa de Mortes por Acidentes de Trânsito (100.000 hab)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 32 - Número de Óbitos por Acidente de Trânsito, População e Taxa de Homicídio no Trânsito.xlsx",
  overwrite = TRUE
)
# MEIO AMBIENTE----
## Tabela 33 - Desflorestamento Acumulado (km²), Incremento do Desflorestamento (km²) e Focos de Calor----
d1 <- bases %>%
  transform(
    tematica = "Meio Ambiente",
    indicador = "Desflorestamento Acumulado (km²), Incremento do Desflorestamento (km²) e Focos de Calor",
    categoria1 = " Desflorestamento Acumulado (km²)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Meio Ambiente",
    indicador = "Desflorestamento Acumulado (km²), Incremento do Desflorestamento (km²) e Focos de Calor",
    categoria1 = " Incremento do Desflorestamento (km²)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Meio Ambiente",
    indicador = "Desflorestamento Acumulado (km²), Incremento do Desflorestamento (km²) e Focos de Calor",
    categoria1 = " Focos de Calor",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 33 - Desflorestamento Acumulado _km²_, Incremento do Desflorestamento _km²_ e Focos de Calor.xlsx",
  overwrite = TRUE
)
## Tabela 34 -Área de Floresta (km²) e Hidrografia (km²)----
d1 <- bases %>%
  transform(
    tematica = "Meio Ambiente",
    indicador = "Área de Floresta (km²) e Hidrografia (km²)",
    categoria1 = "Área de Floresta (km²)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Meio Ambiente",
    indicador = "Área de Floresta (km²) e Hidrografia (km²)",
    categoria1 = " Hidrografia (km²)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb,
  "para/Tabela 34 - Área de Floresta _km²_ e Hidrografia _km²_.xlsx",
  overwrite = TRUE
)
## Tabela 35 - Área Territorial, Área Cadastrável e Área Cadastrada no Cadastro Ambiental Rural (CAR)----
d1 <- bases %>%
  transform(
    tematica = "Meio Ambiente",
    indicador = "Área Territorial, Área Cadastrável e Área Cadastrada no Cadastro Ambiental Rural (CAR)",
    categoria1 = "Área Territorial (IBGE/km²)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Meio Ambiente",
    indicador = "Área Territorial, Área Cadastrável e Área Cadastrada no Cadastro Ambiental Rural (CAR)",
    categoria1 = "Área Cadastrável (km²)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Meio Ambiente",
    indicador = "Área Territorial, Área Cadastrável e Área Cadastrada no Cadastro Ambiental Rural (CAR)",
    categoria1 = "% de Área Cadastrável",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Meio Ambiente",
    indicador = "Área Territorial, Área Cadastrável e Área Cadastrada no Cadastro Ambiental Rural (CAR)",
    categoria1 = "Área de CAR (km²)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Meio Ambiente",
    indicador = "Área Territorial, Área Cadastrável e Área Cadastrada no Cadastro Ambiental Rural (CAR)",
    categoria1 = "% de Área de CAR",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 35 - Área Territorial, Área Cadastrável e Área Cadastrada no Cadastro Ambiental Rural_CAR.xlsx",
  overwrite = TRUE
)
# ECONOMIA----
## Tabela 36 - Produto Interno Bruto Total, Valor Adicionado e Impostos (R$ 1.000)----
d1 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Produto Interno Bruto Total, Valor Adicionado e Impostos (R$ 1.000)",
    categoria1 = " Produto Interno Bruto (PIB)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Produto Interno Bruto Total, Valor Adicionado e Impostos (R$ 1.000)",
    categoria1 = "Valor Adicionado Bruto",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Produto Interno Bruto Total, Valor Adicionado e Impostos (R$ 1.000)",
    categoria1 = " Impostos, Líquidos de Subsídios, sobre Produtos",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 36 - Produto Interno Bruto Total, Valor Adicionado e Impostos (R$ 1.000).xlsx",
  overwrite = TRUE
)
## Tabela 37 - Valor Adicionado Total e por Setores com Administração Pública (R$ 1.000)----
d1 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Valor Adicionado Total e por Setores com Administração Pública (R$ 1.000)",
    categoria1 = "VA Agropecuária",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Valor Adicionado Total e por Setores com Administração Pública (R$ 1.000)",
    categoria1 = "VA Indústria",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Valor Adicionado Total e por Setores com Administração Pública (R$ 1.000)",
    categoria1 = "VA Serviços, exclusive Administração Pública",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Valor Adicionado Total e por Setores com Administração Pública (R$ 1.000)",
    categoria1 = "VA Administração Pública",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Valor Adicionado Total e por Setores com Administração Pública (R$ 1.000)",
    categoria1 = "VA Total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 37 - Valor Adicionado Total e por Setores com Administração Pública.xlsx",
  overwrite = TRUE
)
## Tabela 38 -Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Total do Estado----
d1 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Total do Estado",
    categoria1 = "% Participação PIB Estadual",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Total do Estado",
    categoria1 = "% Participação VA Agropecuária Estadual",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Total do Estado",
    categoria1 = "% Participação VA Industrial Estadual",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Total do Estado",
    categoria1 = "% Participação VA Serviços Estadual, Exclusive Administração Pública",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Total do Estado",
    categoria1 = "% Participação VA Administração Pública Estadual",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Total do Estado",
    categoria1 = "% Participação VA Total Estadual",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 38 -Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Total do Estado.xlsx",
  overwrite = TRUE
)
## Tabela 39 - Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Município----
d1 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Município",
    categoria1 = "% Participação VA Agropecuária",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Município",
    categoria1 = "% Participação VA Indústria",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Município",
    categoria1 = "% Participação VA Serviços, exclusive Administração Pública",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Município",
    categoria1 = "% Participação VA Administração Pública",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Município",
    categoria1 = "% Participação VA Total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 39 - Participação do Valor Adicionado dos Setores e da Administração Pública em relação ao Município.xlsx",
  overwrite = TRUE
)
## Tabela 40 - PIB Total (R$ 1.000), Ranking Estadual e Participação em Relação ao Estado e a Região de Integração----
d1 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "PIB Total (R$ 1.000), Ranking Estadual e Participação em Relação ao Estado e a Região de Integração",
    categoria1 = "PIB (R$ 1.000)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "PIB Total (R$ 1.000), Ranking Estadual e Participação em Relação ao Estado e a Região de Integração",
    categoria1 = "Ranking Estadual",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "PIB Total (R$ 1.000), Ranking Estadual e Participação em Relação ao Estado e a Região de Integração",
    categoria1 = "Ranking Regional",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "PIB Total (R$ 1.000), Ranking Estadual e Participação em Relação ao Estado e a Região de Integração",
    categoria1 = "Partipação no Pará (%)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "PIB Total (R$ 1.000), Ranking Estadual e Participação em Relação ao Estado e a Região de Integração",
    categoria1 = "Partipação na Região (%)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 40 - PIB Total_ Ranking Estadual e Participação em Relação ao Estado e a Região de Integração.xlsx",
  overwrite = TRUE
)
## Tabela 41 - PIB Total (R$ 1.000), População, PIB per capita (R$ 1,00) e Razão entre PIB per capita do Município e do Estado----
d1 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "PIB Total (R$ 1.000), População, PIB per capita (R$ 1,00) e Razão entre PIB per capita do Município e do Estado",
    categoria1 = "PIB (R$ 1.000)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "PIB Total (R$ 1.000), População, PIB per capita (R$ 1,00) e Razão entre PIB per capita do Município e do Estado",
    categoria1 = "População",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "PIB Total (R$ 1.000), População, PIB per capita (R$ 1,00) e Razão entre PIB per capita do Município e do Estado",
    categoria1 = "PIB Per capita",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "PIB Total (R$ 1.000), População, PIB per capita (R$ 1,00) e Razão entre PIB per capita do Município e do Estado",
    categoria1 = "Razão PIB Per capita entre RIs e Pará (R$ 1,00)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 41 - PIB Total População, PIB per capita e Razão entre PIB per capita do Município e do Estado.xlsx",
  overwrite = TRUE
)
## Tabela 42 - Balança Comercial - Exportação, Importação e Saldo (US$ 1.000)----
d1 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Balança Comercial - Exportação, Importação e Saldo (US$ 1.000)",
    categoria1 = "Balança Comercial (US$ 1.000)",
    categoria2 = "Exportação",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Balança Comercial - Exportação, Importação e Saldo (US$ 1.000)",
    categoria1 = "Balança Comercial (US$ 1.000)",
    categoria2 = "Importação",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Balança Comercial - Exportação, Importação e Saldo (US$ 1.000)",
    categoria1 = "Balança Comercial (US$ 1.000)",
    categoria2 = "Saldo",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 42 - Balança Comercial - Exportação, Importação e Saldo.xlsx",
  overwrite = TRUE
)
## Tabela 43 - Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente----
d1 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Banana (cacho)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Banana (cacho)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Banana (cacho)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Cacau (em amêndoa)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Cacau (em amêndoa)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Cacau (em amêndoa)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Café (em grão) Total",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Café (em grão) Total",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Café (em grão) Total",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Coco-da-baía (Mil Frutos)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d11 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Coco-da-baía (Mil Frutos)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d12 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Coco-da-baía (Mil Frutos)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d13 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Dendê (cacho de coco)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d14 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Dendê (cacho de coco)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d15 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Dendê (cacho de coco)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d16 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Laranja",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d17 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Laranja",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d18 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Laranja",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d19 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Limão",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d20 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Limão",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d21 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Limão",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d22 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Mamão",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d23 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Mamão",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d24 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Mamão",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d25 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Maracujá",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d26 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Maracujá",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d27 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Maracujá",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d28 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Pimenta-do-reino",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d29 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Pimenta-do-reino",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d30 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Permanente",
    categoria1 = "Pimenta-do-reino",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12,
  "d13" = d13,
  "d14" = d14,
  "d15" = d15,
  "d16" = d16,
  "d17" = d17,
  "d18" = d18,
  "d19" = d19,
  "d20" = d20,
  "d21" = d21,
  "d22" = d22,
  "d23" = d23,
  "d24" = d24,
  "d25" = d25,
  "d26" = d26,
  "d27" = d27,
  "d28" = d28,
  "d29" = d29,
  "d30" = d30
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 43 - Área colhida_ Quantidade Produzida e Valor da Produção por Tipo de Lavoura Permanente.xlsx",
  overwrite = TRUE
)
## Tabela 44 - Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária----
d1 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Abacaxi (Mil frutos)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Abacaxi (Mil frutos)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Abacaxi (Mil frutos)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Arroz (em casca) (Toneladas)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Arroz (em casca) (Toneladas)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Arroz (em casca) (Toneladas)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Cana-de-açúcar (Toneladas)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Cana-de-açúcar (Toneladas)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Cana-de-açúcar (Toneladas)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Feijão (em grão) (Toneladas)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d11 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Feijão (em grão) (Toneladas)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d12 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Feijão (em grão) (Toneladas)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d13 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Mandioca (Toneladas)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d14 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Mandioca (Toneladas)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d15 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Mandioca (Toneladas)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d16 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Melancia (Toneladas)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d17 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Melancia (Toneladas)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d18 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Melancia (Toneladas)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d19 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Milho (em grão) (Toneladas)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d20 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Milho (em grão) (Toneladas)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d21 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Milho (em grão) (Toneladas)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d22 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Soja (em grão) (Toneladas)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d23 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Soja (em grão) (Toneladas)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d24 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Soja (em grão) (Toneladas)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d25 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Tomate (Toneladas)",
    categoria2 = "Área Colhida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d26 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Tomate (Toneladas)",
    categoria2 = "Qtde Produzida",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d27 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Área colhida (Hectares), Quantidade Produzida e Valor (R$) da Produção por Tipo de Lavoura Temporária",
    categoria1 = "Tomate (Toneladas)",
    categoria2 = "Valor da Produção",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12,
  "d13" = d13,
  "d14" = d14,
  "d15" = d15,
  "d16" = d16,
  "d17" = d17,
  "d18" = d18,
  "d19" = d19,
  "d20" = d20,
  "d21" = d21,
  "d22" = d22,
  "d23" = d23,
  "d24" = d24,
  "d25" = d25,
  "d26" = d26,
  "d27" = d27
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 44 - Área colhida Quantidade Produzida e Valor da Produção por Tipo de Lavoura Temporária.xlsx",
  overwrite = TRUE
)
## Tabela 45 - Efetivo de Rebanho por Tipo----
d1 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Efetivo de Rebanho por Tipo",
    categoria1 = "Bovino",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Efetivo de Rebanho por Tipo",
    categoria1 = "Bubalino",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Efetivo de Rebanho por Tipo",
    categoria1 = "Equino",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Efetivo de Rebanho por Tipo",
    categoria1 = "Suíno - total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Efetivo de Rebanho por Tipo",
    categoria1 = "Suíno - matrizes de suínos",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Efetivo de Rebanho por Tipo",
    categoria1 = "Caprino",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Efetivo de Rebanho por Tipo",
    categoria1 = "Ovino",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Efetivo de Rebanho por Tipo",
    categoria1 = "Galináceos - Total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Efetivo de Rebanho por Tipo",
    categoria1 = "Galináceos - galinhas",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Efetivo de Rebanho por Tipo",
    categoria1 = "Codornas",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb, "para/Tabela 45 - Efetivo de Rebanho por Tipo.xlsx", overwrite = TRUE)
## Tabela 46 - Produção de Origem Animal por Tipo----
d1 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Produção de Origem Animal por Tipo",
    categoria1 = "Leite (Mil litros)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Produção de Origem Animal por Tipo",
    categoria1 = "Ovos de galinha (Mil dúzias)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Produção de Origem Animal por Tipo",
    categoria1 = "Ovos de codorna (Mil dúzias)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Economia",
    indicador = "Produção de Origem Animal por Tipo",
    categoria1 = "Mel de abelha (Quilogramas)",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb,
  "para/Tabela 46 - Produção de Origem Animal por Tipo.xlsx",
  overwrite = TRUE
)
# FINANÇAS PÚBLICAS----
## Tabela 47 - Repasse e Participação ICMS, IPI E IPVA----
d1 <- bases %>%
  transform(
    tematica = "Finanças Públicas",
    indicador = "Repasse e Participação ICMS, IPI E IPVA",
    categoria1 = "ICMS",
    categoria2 = "Repasse",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Finanças Públicas",
    indicador = "Repasse e Participação ICMS, IPI E IPVA",
    categoria1 = "ICMS",
    categoria2 = "Participação (%)",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Finanças Públicas",
    indicador = "Repasse e Participação ICMS, IPI E IPVA",
    categoria1 = "IPI",
    categoria2 = "Repasse",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Finanças Públicas",
    indicador = "Repasse e Participação ICMS, IPI E IPVA",
    categoria1 = "IPI",
    categoria2 = "Participação (%)",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Finanças Públicas",
    indicador = "Repasse e Participação ICMS, IPI E IPVA",
    categoria1 = "IPVA",
    categoria2 = "Repasse",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Finanças Públicas",
    indicador = "Repasse e Participação ICMS, IPI E IPVA",
    categoria1 = "IPVA",
    categoria2 = "Participação (%)",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb,
  "para/Tabela 47 - Repasse e Participação ICMS, IPI E IPVA.xlsx",
  overwrite = TRUE
)
## Tabela 48 - Receitas Orçamentária, Corrente, Transferidas e Impostos----
d1 <- bases %>%
  transform(
    tematica = "Finanças Públicas",
    indicador = "Receitas Orçamentária, Corrente, Transferidas e Impostos",
    categoria1 = "Receita Orçamentária",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Finanças Públicas",
    indicador = "Receitas Orçamentária, Corrente, Transferidas e Impostos",
    categoria1 = "Receita Corrente",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Finanças Públicas",
    indicador = "Receitas Orçamentária, Corrente, Transferidas e Impostos",
    categoria1 = "Receita de Transferências Correntes",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Finanças Públicas",
    indicador = "Receitas Orçamentária, Corrente, Transferidas e Impostos",
    categoria1 = "Impostos",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 48 - Receitas Orçamentária, Corrente, Transferidas e Impostos.xlsx",
  overwrite = TRUE
)
# INFRAESTRUTURA----
## Tabela 49 - Consumo de Energia Elétrica Total (kWH) e Consumidores de Energia Elétrica por Tipo----
d1 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "_",
    categoria1 = "Consumo de Energia Elétrica Total (kWH) e Consumidores de Energia Elétrica por Tipo",
    categoria2 = "Consumo de energia Elétrica (kWH)",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "_",
    categoria1 = "Consumo de Energia Elétrica Total (kWH) e Consumidores de Energia Elétrica por Tipo",
    categoria2 = "Consumidores por Tipo",
    categoria3 = "Residencial",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "_",
    categoria1 = "Consumo de Energia Elétrica Total (kWH) e Consumidores de Energia Elétrica por Tipo",
    categoria2 = "Consumidores por Tipo",
    categoria3 = "Industrial",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "_",
    categoria1 = "Consumo de Energia Elétrica Total (kWH) e Consumidores de Energia Elétrica por Tipo",
    categoria2 = "Consumidores por Tipo",
    categoria3 = "Comercial",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "_",
    categoria1 = "Consumo de Energia Elétrica Total (kWH) e Consumidores de Energia Elétrica por Tipo",
    categoria2 = "Outros",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "_",
    categoria1 = "Consumo de Energia Elétrica Total (kWH) e Consumidores de Energia Elétrica por Tipo",
    categoria2 = "Total",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 49 - Consumo de Energia Elétrica Total_kWH e Consumidores de Energia Elétrica por Tipo.xlsx",
  overwrite = TRUE
)
## Tabela 50 - Total da Frota de Veículos subdivididos em Licenciados e Não Licenciados----
d1 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos subdivididos em Licenciados e Não Licenciados",
    categoria1 = "Frota",
    categoria2 = "Licenciados",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos subdivididos em Licenciados e Não Licenciados",
    categoria1 = "Frota",
    categoria2 = "Não Licenciados",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos subdivididos em Licenciados e Não Licenciados",
    categoria1 = "Frota",
    categoria2 = "Total",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "para/Tabela 50 - Total da Frota de Veículos subdivididos em Licenciados e Não Licenciados.xlsx",
  overwrite = TRUE
)
## Tabela 51 - Total da Frota de Veículos por Tipo----
d1 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Total",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d2 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Automóvel",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d3 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Caminhão",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d4 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Caminhão trator",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d5 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Caminhonete",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d6 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Camioneta",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d7 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Ciclomotor",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d8 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Micro-ônibus",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d9 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Motocicleta",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d10 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Motoneta",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d11 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Ônibus",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d12 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Reboque",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d13 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Semi-reboque",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d14 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Triciclo",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d15 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Utilitário",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )
d16 <- bases %>%
  transform(
    tematica = "Infraestrutura",
    indicador = "Total da Frota de Veículos por Tipo",
    categoria1 = "Outros",
    categoria2 = "-",
    categoria3 = "-",
    categoria4 = "-",
    categoria5 = "-"
  )


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12,
  "d13" = d13,
  "d14" = d14,
  "d15" = d15,
  "d16" = d16
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome) # Criar aba
  writeData(wb, nome, dados_lista[[nome]]) # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(wb,
  "para/Tabela 51 - Total da Frota de Veículos por Tipo.xlsx",
  overwrite = TRUE
)
