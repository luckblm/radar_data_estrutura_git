bases

d1 <- bases %>% 
  transform(
    tematica = "_",
    indicador = "_",
    regiao = "-",
    localidade = "-",
    categoria1 = "_",
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
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
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
  "/.xlsx",
  overwrite = TRUE
)