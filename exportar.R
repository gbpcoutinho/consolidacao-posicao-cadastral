

# ============================================================
# 6. EXPORTAÇÃO FORMATADA PARA EXCEL
# ============================================================

wb <- createWorkbook()

addWorksheet(
  wb,
  sheetName = "Base Tratada",
  gridLines = FALSE
)

estilo_titulo <- createStyle(
  fontSize = 14,
  textDecoration = "bold",
  fontColour = "#FFFFFF",
  fgFill = "#1F4E78",
  halign = "center",
  valign = "center"
)

estilo_cabecalho <- createStyle(
  fontColour = "#FFFFFF",
  fgFill = "#4472C4",
  textDecoration = "bold",
  halign = "center",
  valign = "center",
  wrapText = TRUE,
  border = "Bottom"
)

estilo_corpo <- createStyle(
  valign = "center"
)

estilo_data <- createStyle(
  numFmt = "dd/mm/yyyy",
  valign = "center"
)


writeData(
  wb,
  sheet = "Base Tratada",
  x = "BASE TRATADA — BENEFICIÁRIOS",
  startRow = 1,
  startCol = 1
)

addStyle(
  wb,
  sheet = "Base Tratada",
  style = estilo_titulo,
  rows = 1,
  cols = 1:ncol(final),
  gridExpand = TRUE
)

# Tabela com cabeçalhos e filtros
writeDataTable(
  wb,
  sheet = "Base Tratada",
  x = final,
  startRow = 1,
  startCol = 1,
  tableStyle = "TableStyleMedium2",
  withFilter = TRUE
)

# Estilo do cabeçalho
addStyle(
  wb,
  sheet = "Base Tratada",
  style = estilo_cabecalho,
  rows = 1,
  cols = 1:ncol(final),
  gridExpand = TRUE
)

# Estilo das linhas de dados
if (nrow(final) > 0) {
  
  linhas_dados <- seq(2, nrow(final) + 1)
  
  addStyle(
    wb,
    sheet = "Base Tratada",
    style = estilo_corpo,
    rows = linhas_dados,
    cols = 1:ncol(final),
    gridExpand = TRUE
  )
  
  # Formata as colunas cujo nome contém "DATA".
  colunas_data <- which(
    grepl("DATA", names(final), ignore.case = TRUE)
  )
  
  if (length(colunas_data) > 0) {
    addStyle(
      wb,
      sheet = "Base Tratada",
      style = estilo_data,
      rows = linhas_dados,
      cols = colunas_data,
      gridExpand = TRUE,
      stack = TRUE
    )
  }
}

# Ajusta colunas e mantém título/cabeçalho visíveis ao rolar a planilha.
setColWidths(
  wb,
  sheet = "Base Tratada",
  cols = 1:ncol(final),
  widths = "auto"
)

setColWidths(
  wb,
  sheet = "Base Tratada",
  cols = match(
    c("TITULARIDADE","SubF","SEXO","NUMERO DO CPF","PLANO","REGIAO"),
    names(final)
  ),
  widths = c(18,11,11,18,11,11)
)

setRowHeights(
  wb,
  sheet = "Base Tratada",
  rows = 1,
  heights = 26
)

freezePane(
  wb,
  sheet = "Base Tratada",
  firstActiveRow = 3
)



estilo_centralizado <- createStyle(
  halign = "center",
  valign = "center"
)




# FORMATANDO DATAS ===========
estilo_data_centralizada <- createStyle(
  numFmt = "dd/mm/yyyy",
  halign = "center",
  valign = "center"
)

colunas_data <- which(
  grepl("DATA", names(final), ignore.case = TRUE)
)

addStyle(
  wb,
  sheet = "Base Tratada",
  style = estilo_data_centralizada,
  rows = 2:(nrow(final) + 1),
  cols = colunas_data,
  gridExpand = TRUE,
  stack = TRUE
)
























# FORMATANDO COLUNAS ==========
#====================
# final |>
#   select(
#     SubF,
#     `NUMERO DO CERTIFICADO`,
#     TITULARIDADE,
#     PLANO,
#     REGIAO,
#     `CODIGO DO DEPENDENTE`,
#     `GRAU PARENTESCO`,
#     `NUMERO DO CARTAO DO DEPENDENTE`,
#     SEXO,
#     `ESTADO CIVIL`,
#     `NUMERO DO CPF`,
#     `CARGO DE OCUPACAO`
#   )
# 


colunas_centralizadas <- match(
  c(
    "SubF",
    "NUMERO DO CERTIFICADO",
    "TITULARIDADE",
    "PLANO",
    "REGIAO",
    #"CODIGO DO DEPENDENTE",
    "GRAU PARENTESCO",
    "NUMERO DO CARTAO DO DEPENDENTE",
    "SEXO",
    "ESTADO CIVIL",
    "NUMERO DO CPF",
    "CARGO DE OCUPACAO",
    "NUMERO DA MATRICULA"
  ),
  names(final)
)

colunas_centralizadas <- colunas_centralizadas[!is.na(colunas_centralizadas)]

addStyle(
  wb,
  sheet = "Base Tratada",
  style = estilo_centralizado,
  rows = 2:(nrow(final) + 1),
  cols = colunas_centralizadas,
  gridExpand = TRUE,
  stack = TRUE
)






saveWorkbook(
  wb,
  file = arquivo_saida,
  overwrite = TRUE
)

message("Arquivo exportado com sucesso: ", arquivo_saida)
