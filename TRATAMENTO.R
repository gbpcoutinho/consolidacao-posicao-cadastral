# ============================================================================
# CONSOLIDACAO DE POSICAO CADASTRAL - TITULARES + DEPENDENTES
#
# O script le os arquivos-fonte .xlsx de CFG$pasta e gera um unico arquivo
# Excel com as abas Base Tratada, Pendencias e Log.
#
# Regras de seguranca:
# - arquivos temporarios do Excel e arquivos de saida sao ignorados;
# - identificadores sao lidos como texto; zeros ausentes na origem nao sao
#   reconstruidos;
# - datas validas sao mantidas como Date; datas invalidas ficam vazias e sao
#   descritas no log;
# - arquivos com problemas estruturais sao registrados no log e nao entram
#   parcialmente no consolidado;
# - os dados CARTAO e NOME DA MAE da aba suplementar sao associados ao titular
#   pelo NUMERO DO CERTIFICADO antes de acrescentar as linhas de dependentes.
#
# Pacotes: readxl, dplyr, purrr, openxlsx, lubridate
# ============================================================================

# Usa a localidade padrao do Windows para que nomes com acentos sejam
# preservados na leitura do CSV e na gravacao dos cabecalhos do Excel.
if (.Platform$OS.type == "windows") {
  .locale_sistema <- suppressWarnings(Sys.setlocale("LC_CTYPE", ""))
}

library(readxl)
library(dplyr)
library(purrr)
library(openxlsx)
library(lubridate)

# Monta o caminho com escapes Unicode para funcionar tambem em sessoes R
# do Windows que estejam com uma codificacao de terminal diferente de UTF-8.
# PASTA_RAIZ <- paste0(
#   here::here(),
#   "\u00c7\u00c3O CADASTRAL"
# )

PASTA_RAIZ <- here::here()

# ----------------------------------------------------------------------------
# 1. CONFIGURACAO
# ----------------------------------------------------------------------------

CFG <- list(
  raiz = here::here(),
  pasta = paste0(
    PASTA_RAIZ, "/PC_CONSOLIDADO_", month(today()), "_", year(today())
  ),
  arquivo_mapa_empresas = file.path(PASTA_RAIZ, "COD_EMPRESAS.csv"),
  nome_saida = paste0(
    "PC_CONSOLIDADO_", month(today()), "_", year(today()), ".xlsx"
  ),
  nomes_saida_legados = c(
    "Consolidado - tratado.xlsx",
    "Consolidado - pendencias.xlsx",
    "Consolidado - log de processamento.xlsx"
  ),
  ordem_inicial = c(
    "REFERENCIA", "EMPRESA", "SubF", "NUMERO DO CERTIFICADO",
    "C\u00d3DIGO UNICO","TITULARIDADE","CODIGO DO DEPENDENTE","NOME DO SEGURADO",
    "NOME DEPENDENTE","NOME DA MAE","NUMERO DO CARTAO", "GRAU PARENTESCO","PLANO"
  ),
  colunas_obrigatorias_titular = c(
    "NUMERO DO CERTIFICADO", "NUMERO DA SUBFATURA",
    "NOME DO SEGURADO", "PLANO"
  ),
  colunas_obrigatorias_dependente = c(
    "NUMERO DO CERTIFICADO", "NOME DEPENDENTE",
    "CODIGO DO DEPENDENTE", "GRAU PARENTESCO",
    "NOME DA MAE DO DEPENDENTE", "NUMERO DO CARTAO DO DEPENDENTE"
  ),
  colunas_obrigatorias_suplementar = c(
    "NUMERO DO CERTIFICADO", "CARTAO", "NOME DA MAE"
  ),
  colunas_opcionais_titular = c(
    "DATA DE REATIVACAO", "REGIAO", "CARGO DE OCUPACAO",
    "SEXO", "NUMERO DO CPF", "NUMERO DA MATRICULA"
  ),
  colunas_opcionais_dependente = c(
    "DATA DE REATIVACAO", "SEXO", "ESTADO CIVIL",
    "NUMERO DA SUBFATURA"
  ),
  largura_manual = c(
    "REFERENCIA" = 18,
    "EMPRESA" = 28,
    "TITULARIDADE" = 18,
    "SubF" = 11,
    "NUMERO DO CERTIFICADO" = 22,
    "C\u00d3DIGO UNICO" = 26,
    "NOME DO SEGURADO" = 44,
    "NUMERO DO CARTAO" = 22,
    "NOME DA MAE" = 34,
    "NOME DEPENDENTE" = 40,
    "NOME DA MAE" = 34,
    "CARGO DE OCUPACAO" = 28,
    "AVISOS_ESTRUTURA" = 60,
    "OBSERVACOES_DATAS" = 60,
    "DATAS_INVALIDAS" = 14,
    "DIVERGENCIAS_DATAS" = 60,
    "QTD_DIVERGENCIAS_DATAS" = 18,
    "MENSAGEM" = 60,
    "SEXO" = 11,
    "PLANO" = 11,
    "NUMERO DO CPF"= 19,
    "DATA DE INICIO DE VIGENCIA Y2K" =20,
    "REGIAO" = 11,
    "DATA DE CANCELAMENTO Y2K" = 20,
    "DATA DE INICIO DE VIGENCIA" = 20,
    "DATA DE NASCIMENTO" = 16,
    "DATA DE NASCIMENTO Y2K" = 20,
    "DATA DE ADMISSAO" = 16
  )
)

PADRAO_IDENTIFICADOR <- paste0(
  "NUMERO|CODIGO|CARTAO|CERTIFICADO|SUBFATURA|CPF|MATRICULA|INSCRICAO"
)


# ----------------------------------------------------------------------------
# 2. CABECALHOS E VALIDACOES DE ESTRUTURA
# ----------------------------------------------------------------------------

normalizar_nome <- function(x) {
  x <- toupper(trimws(as.character(x)))
  x <- chartr(
    "\u00c1\u00c0\u00c2\u00c3\u00c4\u00c9\u00c8\u00ca\u00cb\u00cd\u00cc\u00ce\u00cf\u00d3\u00d2\u00d4\u00d5\u00d6\u00da\u00d9\u00db\u00dc\u00c7",
    "AAAAAEEEEIIIIOOOOOUUUUC",
    x
  )
  x <- gsub("[^A-Z0-9]+", "_", x)
  x <- gsub("^_+|_+$", "", x)
  x[is.na(x)] <- ""
  x
}


renomear_e_validar_cabecalhos <- function(df, regras, contexto) {
  nomes_originais <- names(df)
  nomes_norm <- normalizar_nome(nomes_originais)
  
  idx_tipo <- which(nomes_norm == "TIPO_DE_REGISTRO")
  if (length(idx_tipo) > 0) {
    df <- df[, -idx_tipo, drop = FALSE]
    nomes_originais <- names(df)
    nomes_norm <- normalizar_nome(nomes_originais)
  }
  
  if (any(nomes_norm == "")) {
    stop(sprintf("[%s] Existe coluna sem nome; layout nao seguro.", contexto),
         call. = FALSE)
  }
  
  nomes_finais <- gsub("_+", " ", nomes_norm)
  nomes_finais <- trimws(nomes_finais)
  
  for (padrao in names(regras)) {
    idx <- grep(padrao, nomes_norm, perl = TRUE)
    if (length(idx) > 1) {
      stop(sprintf(
        "[%s] O padrao '%s' corresponde a varias colunas: %s",
        contexto, padrao, paste(nomes_originais[idx], collapse = ", ")
      ), call. = FALSE)
    }
    if (length(idx) == 1) {
      nomes_finais[idx] <- unname(regras[[padrao]])
    }
  }
  
  repetidos <- unique(nomes_finais[duplicated(nomes_finais)])
  if (length(repetidos) > 0) {
    stop(sprintf(
      "[%s] A padronizacao gerou cabecalhos duplicados: %s",
      contexto, paste(repetidos, collapse = ", ")
    ), call. = FALSE)
  }
  
  names(df) <- nomes_finais
  df
}


validar_colunas <- function(df, obrigatorias, contexto) {
  faltantes <- setdiff(obrigatorias, names(df))
  if (length(faltantes) > 0) {
    stop(sprintf(
      "[%s] Coluna(s) obrigatoria(s) ausente(s): %s",
      contexto, paste(faltantes, collapse = ", ")
    ), call. = FALSE)
  }
  invisible(TRUE)
}


# ----------------------------------------------------------------------------
# 3. LEITURA CONTROLADA E DATAS
# ----------------------------------------------------------------------------

tipos_por_cabecalho <- function(nomes) {
  nomes_norm <- normalizar_nome(nomes)
  tipos <- rep("text", length(nomes_norm))
  tipos[grepl("DATA", nomes_norm)] <- "list"
  tipos
}


ler_aba <- function(caminho, aba, skip = 0) {
  cabecalho <- read_excel(
    caminho, sheet = aba, skip = skip, n_max = 0,
    .name_repair = "minimal"
  )
  if (length(names(cabecalho)) == 0) {
    stop(sprintf("A aba '%s' nao tem cabecalhos identificaveis.", aba),
         call. = FALSE)
  }
  read_excel(
    caminho,
    sheet = aba,
    skip = skip,
    col_types = tipos_por_cabecalho(names(cabecalho)),
    .name_repair = "minimal"
  )
}


texto_celula <- function(x) {
  if (length(x) == 0 || all(is.na(x))) return(NA_character_)
  if (inherits(x, "Date")) return(format(x[1], "%Y-%m-%d"))
  if (inherits(x, "POSIXt")) return(format(x[1], "%Y-%m-%d %H:%M:%S", tz = "UTC"))
  if (is.numeric(x)) return(format(x[1], scientific = FALSE, trim = TRUE))
  as.character(x[1])
}


converter_serial_excel <- function(serial) {
  dia <- floor(serial)
  if (!is.finite(serial) || dia < 1 || dia > 2958465 || dia == 60) {
    return(as.Date(NA))
  }
  # O Excel tem um dia ficticio (serial 60 = 29/02/1900). Para dias anteriores
  # a ele, o ajuste de origem e diferente; para 60, a data e invalida.
  origem <- if (dia < 60) "1899-12-31" else "1899-12-30"
  as.Date(dia, origin = origem)
}


converter_data_com_ano_curto <- function(valor) {
  # Regra explicita para datas como 05/05/71 e 05/05/02:
  # 00-29 => 2000-2029; 30-99 => 1930-1999.
  # Retorna NULL quando o texto nao tem ano com dois digitos e Date NA quando
  # tem o formato, mas a combinacao dia/mes/ano e invalida.
  partes <- regmatches(
    valor,
    regexec(
      "^([0-9]{1,2})[/.-]([0-9]{1,2})[/.-]([0-9]{2})(?:[ T].*)?$",
      valor,
      perl = TRUE
    )
  )[[1]]
  if (length(partes) == 0) return(NULL)
  
  dia <- as.integer(partes[2])
  mes <- as.integer(partes[3])
  ano_curto <- as.integer(partes[4])
  ano <- if (ano_curto <= 29L) 2000L + ano_curto else 1900L + ano_curto
  iso <- sprintf("%04d-%02d-%02d", ano, mes, dia)
  data <- suppressWarnings(as.Date(iso, format = "%Y-%m-%d"))
  
  if (is.na(data) || format(data, "%Y-%m-%d") != iso) {
    return(as.Date(NA))
  }
  data
}


converter_celula_data <- function(x, nome_coluna, linha, contexto) {
  vazio <- list(data = as.Date(NA), observacao = NULL)
  if (length(x) == 0 || all(is.na(x))) return(vazio)
  
  if (inherits(x, "Date")) {
    return(list(data = as.Date(x[1]), observacao = NULL))
  }
  if (inherits(x, "POSIXt")) {
    return(list(data = as.Date(x[1], tz = "UTC"), observacao = NULL))
  }
  
  if (is.numeric(x)) {
    serial <- as.numeric(x[1])
    data_serial <- converter_serial_excel(serial)
    if (!is.na(data_serial)) {
      return(list(
        data = data_serial,
        observacao = NULL
      ))
    }
    valor_original <- texto_celula(x)
  } else {
    valor_original <- trimws(texto_celula(x))
    if (is.na(valor_original) || valor_original == "") return(vazio)
    
    data_ano_curto <- converter_data_com_ano_curto(valor_original)
    if (!is.null(data_ano_curto)) {
      if (!is.na(data_ano_curto)) {
        return(list(data = data_ano_curto, observacao = NULL))
      }
      # O formato era de ano curto, mas a data nao existe no calendario.
      # Nao tentamos reinterpretar o ano de dois digitos como ano do seculo I.
      observacao <- sprintf(
        "%s | coluna '%s' | registro %d | valor invalido '%s' (convertido para vazio)",
        contexto, nome_coluna, linha, valor_original
      )
      return(list(data = as.Date(NA), observacao = observacao))
    }
    
    formatos <- c(
      "%d/%m/%Y", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M",
      "%d-%m-%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S",
      "%Y-%m-%dT%H:%M:%S", "%Y%m%d", "%d%m%Y"
    )
    for (formato in formatos) {
      convertido <- suppressWarnings(as.Date(valor_original, format = formato))
      if (!is.na(convertido)) {
        return(list(data = convertido, observacao = NULL))
      }
    }
    
    # Algumas origens armazenam a data como serial Excel, mas como texto.
    if (grepl("^[0-9]+([.,][0-9]+)?$", valor_original)) {
      serial <- suppressWarnings(as.numeric(sub(",", ".", valor_original, fixed = TRUE)))
      data_serial <- converter_serial_excel(serial)
      if (!is.na(data_serial)) {
        return(list(
          data = data_serial,
          observacao = NULL
        ))
      }
    }
  }
  
  observacao <- sprintf(
    "%s | coluna '%s' | registro %d | valor invalido '%s' (convertido para vazio)",
    contexto, nome_coluna, linha, valor_original
  )
  list(data = as.Date(NA), observacao = observacao)
}


validar_colunas_data <- function(df, contexto) {
  observacoes <- character()
  divergencias <- character()
  idx_datas <- which(grepl("DATA", normalizar_nome(names(df))))
  
  for (i in idx_datas) {
    coluna <- df[[i]]
    if (inherits(coluna, "Date")) next
    if (inherits(coluna, "POSIXt")) {
      df[[i]] <- as.Date(coluna, tz = "UTC")
      next
    }
    
    celulas <- if (is.list(coluna)) coluna else as.list(coluna)
    resultados <- lapply(seq_along(celulas), function(j) {
      converter_celula_data(
        celulas[[j]], names(df)[i], j, contexto
      )
    })
    
    observacoes <- c(
      observacoes,
      unlist(lapply(resultados, `[[`, "observacao"), use.names = FALSE)
    )
    numeros <- vapply(
      resultados,
      function(x) as.numeric(x$data),
      numeric(1)
    )
    df[[i]] <- as.Date(numeros, origin = "1970-01-01")
  }
  
  # Quando a data principal e a correspondente Y2K representam o mesmo dia e
  # mes, a Y2K fornece o ano completo autorizado pelo usuario. Se dia/mes nao
  # coincidirem, preservamos a data principal e registramos a divergencia.
  nomes_norm <- normalizar_nome(names(df))
  idx_y2k <- which(grepl("_Y2K$", nomes_norm))
  for (idx_y in idx_y2k) {
    nome_base_norm <- sub("_Y2K$", "", nomes_norm[idx_y])
    idx_base <- which(
      nomes_norm == nome_base_norm & grepl("DATA", nomes_norm) &
        seq_along(nomes_norm) != idx_y
    )
    if (length(idx_base) > 1) {
      stop(sprintf(
        "[%s] Mais de uma coluna de data corresponde a '%s'.",
        contexto, names(df)[idx_y]
      ), call. = FALSE)
    }
    if (length(idx_base) == 0) next
    
    datas_base <- as.Date(df[[idx_base]])
    datas_y2k <- as.Date(df[[idx_y]])
    ambas_preenchidas <- !is.na(datas_base) & !is.na(datas_y2k)
    mesmo_dia_mes <- rep(FALSE, length(datas_base))
    mesmo_dia_mes[ambas_preenchidas] <-
      format(datas_base[ambas_preenchidas], "%m-%d") ==
      format(datas_y2k[ambas_preenchidas], "%m-%d")
    
    usar_y2k <- ambas_preenchidas & mesmo_dia_mes
    if (any(usar_y2k)) {
      datas_base[usar_y2k] <- datas_y2k[usar_y2k]
      df[[idx_base]] <- datas_base
    }
    
    divergentes <- which(ambas_preenchidas & !mesmo_dia_mes)
    if (length(divergentes) > 0) {
      divergencias <- c(divergencias, vapply(divergentes, function(j) {
        sprintf(
          "%s | registro %d | divergencia entre '%s' (%s) e '%s' (%s); mantida a data principal",
          contexto, j, names(df)[idx_base],
          format(datas_base[j], "%d/%m/%Y"), names(df)[idx_y],
          format(datas_y2k[j], "%d/%m/%Y")
        )
      }, character(1)))
    }
  }
  
  list(dados = df, observacoes = observacoes, divergencias = divergencias)
}


resumir_observacoes_datas <- function(observacoes, max_amostras = 4L) {
  if (length(observacoes) == 0) return(NA_character_)
  
  # O Excel limita uma celula a 32.767 caracteres. O log mantem a contagem
  # completa por aba/coluna e exemplos, sem tentar gravar milhares de celulas
  # invalidas em uma unica celula do workbook.
  grupos_chave <- sub(" \\| registro [0-9]+ \\| .*$", "", observacoes)
  grupos <- split(observacoes, grupos_chave)
  resumos <- vapply(names(grupos), function(chave) {
    itens <- grupos[[chave]]
    exemplos <- head(itens, max_amostras)
    exemplos <- sub(
      "^.* \\| (registro [0-9]+ \\| .*)$",
      "\\1",
      exemplos
    )
    complemento <- if (length(itens) > max_amostras) {
      paste0("; +", length(itens) - max_amostras, " caso(s)")
    } else {
      ""
    }
    paste0(chave, ": ", length(itens), " valor(es); amostras: ",
           paste(exemplos, collapse = "; "), complemento)
  }, character(1))
  paste(resumos, collapse = " || ")
}


validar_identificadores_como_texto <- function(df, contexto) {
  idx <- which(grepl(PADRAO_IDENTIFICADOR, normalizar_nome(names(df)), perl = TRUE))
  for (i in idx) {
    if (!is.character(df[[i]])) {
      stop(sprintf(
        "[%s] A coluna identificadora '%s' nao foi lida como texto.",
        contexto, names(df)[i]
      ), call. = FALSE)
    }
  }
  df
}


chave_certificado <- function(x) {
  # A chave auxiliar e aparada apenas para comparacao. O valor original da
  # coluna NUMERO DO CERTIFICADO permanece intacto na base exportada.
  x <- as.character(x)
  x <- trimws(x)
  x[is.na(x) | x == ""] <- NA_character_
  x
}


unificar_colunas <- function(df, nome_final, fontes, contexto) {
  # Mantem o primeiro valor preenchido na ordem: coluna final e depois aliases.
  # Valores divergentes em colunas equivalentes interrompem o arquivo para nao
  # descartar uma das informacoes silenciosamente.
  origens <- unique(c(nome_final, fontes))
  presentes <- intersect(origens, names(df))
  valores_finais <- rep(NA_character_, nrow(df))
  
  for (origem in presentes) {
    valores <- as.character(df[[origem]])
    preenchidos <- !is.na(valores) & trimws(valores) != ""
    ja_preenchido <- !is.na(valores_finais) & trimws(valores_finais) != ""
    conflitos <- ja_preenchido & preenchidos &
      trimws(valores_finais) != trimws(valores)
    if (any(conflitos)) {
      linha <- which(conflitos)[1]
      stop(sprintf(
        "[%s] Valores conflitantes nas colunas equivalentes usadas para '%s' (registro %d).",
        contexto, nome_final, linha
      ), call. = FALSE)
    }
    usar <- !ja_preenchido & preenchidos
    valores_finais[usar] <- valores[usar]
  }
  
  df[[nome_final]] <- valores_finais
  remover <- setdiff(presentes, nome_final)
  if (length(remover) > 0) {
    df <- df[, setdiff(names(df), remover), drop = FALSE]
  }
  df
}


# ----------------------------------------------------------------------------
# 4. MAPA CODIGO -> EMPRESA
# ----------------------------------------------------------------------------

carregar_mapa_empresas <- function(caminho) {
  if (!file.exists(caminho)) {
    stop("Arquivo de empresas nao encontrado: ", caminho, call. = FALSE)
  }
  
  # Le a entrada como linhas UTF-8 antes de separar os campos. Isso evita que
  # sessoes R com locale C descartem linhas que contenham nomes acentuados.
  linhas_csv <- readLines(caminho, encoding = "UTF-8", warn = FALSE)
  if (length(linhas_csv) == 0) {
    stop("O arquivo COD_EMPRESAS.csv esta vazio.", call. = FALSE)
  }
  mapa <- read.csv2(
    text = paste(linhas_csv, collapse = "\n"),
    colClasses = "character",
    check.names = FALSE,
    stringsAsFactors = FALSE,
    # O BOM eventual e removido do primeiro nome de coluna logo abaixo.
    fileEncoding = ""
  )
  if (length(names(mapa)) > 0 && startsWith(names(mapa)[1], "\ufeff")) {
    names(mapa)[1] <- substring(names(mapa)[1], 2)
  }
  nomes_norm <- normalizar_nome(names(mapa))
  idx_empresa <- match("EMPRESA", nomes_norm)
  idx_codigo <- match("COD_SAUDE", nomes_norm)
  
  if (is.na(idx_empresa) || is.na(idx_codigo)) {
    stop(
      "O COD_EMPRESAS.csv precisa conter as colunas EMPRESA e COD_SAUDE.",
      call. = FALSE
    )
  }
  
  mapa <- data.frame(
    EMPRESA = mapa[[idx_empresa]],
    COD_SAUDE = mapa[[idx_codigo]],
    stringsAsFactors = FALSE
  )
  mapa$EMPRESA <- trimws(mapa$EMPRESA)
  mapa$COD_SAUDE <- trimws(mapa$COD_SAUDE)
  mapa <- mapa[
    !is.na(mapa$EMPRESA) & nzchar(mapa$EMPRESA) &
      !is.na(mapa$COD_SAUDE) & nzchar(mapa$COD_SAUDE),
    , drop = FALSE
  ]
  
  if (nrow(mapa) == 0) {
    stop("O COD_EMPRESAS.csv nao tem pares EMPRESA/COD_SAUDE validos.",
         call. = FALSE)
  }
  
  codigos <- unique(mapa$COD_SAUDE)
  mapa_unico <- lapply(codigos, function(codigo) {
    linhas <- mapa[mapa$COD_SAUDE == codigo, , drop = FALSE]
    empresas <- unique(linhas$EMPRESA)
    if (length(empresas) > 1) {
      stop(sprintf(
        "COD_SAUDE '%s' esta associado a mais de uma EMPRESA: %s",
        codigo, paste(empresas, collapse = ", ")
      ), call. = FALSE)
    }
    data.frame(EMPRESA = empresas[1], COD_SAUDE = codigo,
               stringsAsFactors = FALSE)
  })
  bind_rows(mapa_unico)
}


empresa_da_referencia <- function(referencia, mapa) {
  partes <- regmatches(
    referencia,
    regexec("^PC([0-9]+)", referencia, ignore.case = TRUE)
  )[[1]]
  if (length(partes) < 2) {
    stop(sprintf(
      "Nao foi possivel extrair o codigo apos 'PC' da referencia '%s'.",
      referencia
    ), call. = FALSE)
  }
  codigo <- partes[2]
  idx <- match(codigo, mapa$COD_SAUDE)
  if (is.na(idx)) {
    stop(sprintf(
      "Codigo '%s' extraido de '%s' nao encontrado em COD_SAUDE no COD_EMPRESAS.csv.",
      codigo, referencia
    ), call. = FALSE)
  }
  list(codigo = codigo, empresa = mapa$EMPRESA[idx])
}


# ----------------------------------------------------------------------------
# 5. LEITURA E TRATAMENTO DE CADA ABA
# ----------------------------------------------------------------------------

ler_titular <- function(caminho, aba, referencia) {
  contexto <- paste("TITULAR:", referencia)
  df <- ler_aba(caminho, aba, skip = 2)
  regras <- c(
    "REATIVA" = "DATA DE REATIVACAO",
    "REGI.O" = "REGIAO",
    "OCUPA" = "CARGO DE OCUPACAO",
    "SEXO.*SEGURADO" = "SEXO"
  )
  df <- renomear_e_validar_cabecalhos(df, regras, contexto)
  df <- validar_identificadores_como_texto(df, contexto)
  datas <- validar_colunas_data(df, contexto)
  df <- datas$dados
  validar_colunas(df, CFG$colunas_obrigatorias_titular, contexto)
  
  chaves <- chave_certificado(df[["NUMERO DO CERTIFICADO"]])
  if (anyNA(chaves)) {
    stop(sprintf(
      "[%s] Ha titular sem NUMERO DO CERTIFICADO; nao e possivel relacionar os dados.",
      contexto
    ), call. = FALSE)
  }
  
  # Regra confirmada: titular recebe CODIGO DO DEPENDENTE = "00".
  df$TITULARIDADE <- "TITULAR"
  df[["CODIGO DO DEPENDENTE"]] <- "00"
  df$REFERENCIA <- referencia
  df$.CHAVE_CERTIFICADO <- chaves
  
  list(
    dados = df,
    avisos = setdiff(CFG$colunas_opcionais_titular, names(df)),
    observacoes_datas = datas$observacoes,
    divergencias_datas = datas$divergencias
  )
}


ler_dependente <- function(caminho, aba, referencia) {
  contexto <- paste("DEPENDENTE:", referencia)
  bruto <- ler_aba(caminho, aba, skip = 0)
  
  if (nrow(bruto) <= 2) {
    stop(sprintf(
      "[%s] Ha %d linha(s); nao existem dados apos remover as duas linhas finais de resumo.",
      contexto, nrow(bruto)
    ), call. = FALSE)
  }
  
  # As duas linhas finais sao sintese/resumo, nao registros nem erros de layout.
  df <- bruto[seq_len(nrow(bruto) - 2), , drop = FALSE]
  regras <- c(
    "REATIVA" = "DATA DE REATIVACAO",
    "SEXO.*DEPENDENTE" = "SEXO",
    "ESTADO.*CIVIL" = "ESTADO CIVIL",
    "CART.*DEPENDENTE" = "NUMERO DO CARTAO DO DEPENDENTE",
    "NOME.*M.E.*DEPENDENTE" = "NOME DA MAE DO DEPENDENTE"
  )
  df <- renomear_e_validar_cabecalhos(df, regras, contexto)
  df <- validar_identificadores_como_texto(df, contexto)
  datas <- validar_colunas_data(df, contexto)
  df <- datas$dados
  validar_colunas(df, CFG$colunas_obrigatorias_dependente, contexto)
  
  sobrepostas <- intersect(c("NOME DO SEGURADO", "PLANO"), names(df))
  if (length(sobrepostas) > 0) {
    stop(sprintf(
      "[%s] Campo(s) que devem vir do titular tambem aparecem na aba dependente: %s",
      contexto, paste(sobrepostas, collapse = ", ")
    ), call. = FALSE)
  }
  
  if ("NUMERO DA SUBFATURA" %in% names(df)) {
    df <- dplyr::rename(df, .SUBF_DEPENDENTE = `NUMERO DA SUBFATURA`)
  } else {
    df$.SUBF_DEPENDENTE <- NA_character_
  }
  
  df$TITULARIDADE <- "DEPENDENTE"
  df$REFERENCIA <- referencia
  df$.CHAVE_CERTIFICADO <- chave_certificado(df[["NUMERO DO CERTIFICADO"]])
  
  list(
    dados = df,
    avisos = setdiff(CFG$colunas_opcionais_dependente, names(df)),
    observacoes_datas = datas$observacoes,
    divergencias_datas = datas$divergencias
  )
}


linhas_vazias <- function(df) {
  if (nrow(df) == 0) return(logical())
  vapply(seq_len(nrow(df)), function(i) {
    valores <- lapply(df, function(coluna) {
      if (is.list(coluna)) coluna[[i]] else coluna[i]
    })
    all(vapply(valores, function(valor) {
      txt <- texto_celula(valor)
      is.na(txt) || trimws(txt) == ""
    }, logical(1)))
  }, logical(1))
}


agregar_aba_suplementar <- function(df, contexto) {
  df <- df[!linhas_vazias(df), , drop = FALSE]
  if (nrow(df) == 0) {
    stop(sprintf("[%s] A aba suplementar nao contem registros.", contexto),
         call. = FALSE)
  }
  
  validar_colunas(df, CFG$colunas_obrigatorias_suplementar, contexto)
  df <- validar_identificadores_como_texto(df, contexto)
  chaves <- chave_certificado(df[["NUMERO DO CERTIFICADO"]])
  if (anyNA(chaves)) {
    stop(sprintf(
      "[%s] Ha registro suplementar sem NUMERO DO CERTIFICADO.", contexto
    ), call. = FALSE)
  }
  df$.CHAVE_CERTIFICADO <- chaves
  
  # Repeticoes identicas sao agregadas em um registro. Se um certificado tiver
  # mais de um CARTAO ou NOME DA MAE diferente, o arquivo falha: nao se escolhe
  # um valor arbitrariamente nem se multiplica a linha do titular.
  grupos <- split(df, df$.CHAVE_CERTIFICADO)
  agregados <- lapply(names(grupos), function(chave) {
    grupo <- grupos[[chave]]
    obter_valor_unico <- function(nome_coluna) {
      valores <- as.character(grupo[[nome_coluna]])
      preenchidos <- !is.na(valores) & nzchar(trimws(valores))
      distintos <- unique(trimws(valores[preenchidos]))
      if (length(distintos) > 1) {
        stop(sprintf(
          "[%s] O certificado '%s' tem valores conflitantes em '%s': %s",
          contexto, chave, nome_coluna, paste(distintos, collapse = " | ")
        ), call. = FALSE)
      }
      if (!any(preenchidos)) NA_character_ else valores[which(preenchidos)[1]]
    }
    data.frame(
      .CHAVE_CERTIFICADO = chave,
      .CARTAO_SUP = obter_valor_unico("CARTAO"),
      .NOME_MAE_SUP = obter_valor_unico("NOME DA MAE"),
      stringsAsFactors = FALSE
    )
  })
  bind_rows(agregados)
}


ler_suplementar <- function(caminho, aba, referencia) {
  contexto <- paste("POS. CADASTRAL (NOME M E):", referencia)
  df <- ler_aba(caminho, aba, skip = 0)
  regras <- c(
    "^CART.?O$" = "CARTAO",
    "^NOME_DA_M.?E$" = "NOME DA MAE"
  )
  df <- renomear_e_validar_cabecalhos(df, regras, contexto)
  datas <- validar_colunas_data(df, contexto)
  df <- datas$dados
  agregado <- agregar_aba_suplementar(df, contexto)
  list(
    dados = agregado,
    observacoes_datas = datas$observacoes,
    divergencias_datas = datas$divergencias
  )
}


# ----------------------------------------------------------------------------
# 6. PROCESSAMENTO DE UM ARQUIVO COMPLETO
# ----------------------------------------------------------------------------

processar_arquivo <- function(caminho, mapa_empresas) {
  arquivo <- basename(caminho)
  referencia <- tools::file_path_sans_ext(arquivo)
  observacoes_datas <- character()
  divergencias_datas <- character()
  
  resultado <- tryCatch({
    empresa_info <- empresa_da_referencia(referencia, mapa_empresas)
    # No Windows, o caminho 8.3 evita falhas de unzip/readxl em pastas com
    # acentos quando a sessao R esta usando uma code page diferente de UTF-8.
    caminho_leitura <- caminho
    if (.Platform$OS.type == "windows") {
      caminho_curto <- tryCatch(utils::shortPathName(caminho), error = function(e) "")
      if (length(caminho_curto) == 1 && nzchar(caminho_curto)) {
        caminho_leitura <- caminho_curto
      }
    }
    
    abas <- excel_sheets(caminho_leitura)
    abas_norm <- normalizar_nome(abas)
    
    aba_titular <- abas[grepl("TITULAR", abas_norm)]
    aba_dependente <- abas[grepl("DEPENDENT", abas_norm)]
    aba_suplementar <- abas[grepl("POS.*CADASTRAL.*NOME.*M.*E", abas_norm)]
    
    if (length(aba_titular) != 1) {
      stop(sprintf(
        "Esperada exatamente uma aba de TITULAR; encontrada(s) %d.",
        length(aba_titular)
      ), call. = FALSE)
    }
    if (length(aba_dependente) != 1) {
      stop(sprintf(
        "Esperada exatamente uma aba de DEPENDENTE; encontrada(s) %d.",
        length(aba_dependente)
      ), call. = FALSE)
    }
    if (length(aba_suplementar) != 1) {
      stop(sprintf(
        "Esperada exatamente uma aba POS. CADASTRAL (NOME M E); encontrada(s) %d.",
        length(aba_suplementar)
      ), call. = FALSE)
    }
    
    titular_info <- ler_titular(caminho_leitura, aba_titular, referencia)
    observacoes_datas <- c(observacoes_datas, titular_info$observacoes_datas)
    divergencias_datas <- c(divergencias_datas, titular_info$divergencias_datas)
    dependente_info <- ler_dependente(caminho_leitura, aba_dependente, referencia)
    observacoes_datas <- c(observacoes_datas, dependente_info$observacoes_datas)
    divergencias_datas <- c(divergencias_datas, dependente_info$divergencias_datas)
    suplementar_info <- ler_suplementar(caminho_leitura, aba_suplementar, referencia)
    observacoes_datas <- c(observacoes_datas, suplementar_info$observacoes_datas)
    divergencias_datas <- c(divergencias_datas, suplementar_info$divergencias_datas)
    
    titular <- titular_info$dados
    dependente <- dependente_info$dados
    suplementar <- suplementar_info$dados
    
    # Reune os nomes equivalentes da aba dependente nas colunas finais comuns.
    dependente <- unificar_colunas(
      dependente,
      "NUMERO DO CARTAO",
      "NUMERO DO CARTAO DO DEPENDENTE",
      contexto = paste("DEPENDENTE:", referencia)
    )
    dependente <- unificar_colunas(
      dependente,
      "NOME DA MAE",
      "NOME DA MAE DO DEPENDENTE",
      contexto = paste("DEPENDENTE:", referencia)
    )
    
    certificados_titular <- titular$.CHAVE_CERTIFICADO
    duplicados <- unique(certificados_titular[duplicated(certificados_titular)])
    if (length(duplicados) > 0) {
      stop(sprintf(
        "Certificado(s) duplicado(s) na aba TITULAR: %s",
        paste(duplicados, collapse = ", ")
      ), call. = FALSE)
    }
    
    antes_join <- nrow(titular)
    titular <- left_join(titular, suplementar, by = ".CHAVE_CERTIFICADO")
    if (nrow(titular) != antes_join) {
      stop(
        "A junção da aba suplementar alterou a quantidade de titulares; consolidacao interrompida.",
        call. = FALSE
      )
    }
    
    # Mantem valores ja presentes no titular e completa com a aba suplementar.
    titular <- unificar_colunas(
      titular,
      "NUMERO DO CARTAO",
      c("CARTAO", ".CARTAO_SUP"),
      contexto = paste("TITULAR:", referencia)
    )
    titular <- unificar_colunas(
      titular,
      "NOME DA MAE",
      ".NOME_MAE_SUP",
      contexto = paste("TITULAR:", referencia)
    )
    
    # Empresa e codigo da referencia sao atribuidos antes da combinacao das abas.
    titular$EMPRESA <- empresa_info$empresa
    dependente$EMPRESA <- empresa_info$empresa
    lookup_titular <- titular |>
      transmute(
        .CHAVE_CERTIFICADO,
        .NOME_TITULAR = `NOME DO SEGURADO`,
        .PLANO_TITULAR = PLANO,
        .SUBF_TITULAR = `NUMERO DA SUBFATURA`,
        .TITULAR_ENCONTRADO = TRUE
      )
    
    n_dependentes_antes <- nrow(dependente)
    dependente_enriquecido <- left_join(
      dependente,
      lookup_titular,
      by = ".CHAVE_CERTIFICADO",
      na_matches = "never"
    )
    if (nrow(dependente_enriquecido) != n_dependentes_antes) {
      stop(
        "A juncao dos dependentes multiplicou ou removeu registros.",
        call. = FALSE
      )
    }
    
    dependente_enriquecido <- dependente_enriquecido |>
      mutate(
        `NOME DO SEGURADO` = .NOME_TITULAR,
        PLANO = .PLANO_TITULAR,
        `NUMERO DA SUBFATURA` = coalesce(.SUBF_TITULAR, .SUBF_DEPENDENTE)
      )
    
    pendencias <- dependente_enriquecido |>
      filter(is.na(.TITULAR_ENCONTRADO)) |>
      transmute(
        REFERENCIA,
        EMPRESA,
        `NUMERO DA SUBFATURA`,
        `NUMERO DO CERTIFICADO`,
        `CODIGO DO DEPENDENTE`,
        `NOME DEPENDENTE`,
        `NOME DO SEGURADO`,
        PLANO,
        MOTIVO = "CERTIFICADO SEM TITULAR CORRESPONDENTE"
      )
    
    dependente_final <- dependente_enriquecido |>
      select(-all_of(c(
        ".SUBF_DEPENDENTE", ".NOME_TITULAR", ".PLANO_TITULAR",
        ".SUBF_TITULAR", ".TITULAR_ENCONTRADO"
      )))
    
    # Campos especificos do titular permanecem vazios nas linhas dependentes.
    final_arquivo <- bind_rows(titular, dependente_final) |>
      rename(SubF = `NUMERO DA SUBFATURA`)
    
    certificado <- as.character(final_arquivo[["NUMERO DO CERTIFICADO"]])
    codigo_dep <- as.character(final_arquivo[["CODIGO DO DEPENDENTE"]])
    valido <- !is.na(certificado) & trimws(certificado) != "" &
      !is.na(codigo_dep) & trimws(codigo_dep) != ""
    final_arquivo[["C\u00d3DIGO UNICO"]] <- NA_character_
    final_arquivo[["C\u00d3DIGO UNICO"]][valido] <- paste0(
      certificado[valido], codigo_dep[valido]
    )
    
    chaves_suplementares <- suplementar$.CHAVE_CERTIFICADO
    titulares_sem_suplemento <- sum(
      !titular$.CHAVE_CERTIFICADO %in% chaves_suplementares
    )
    suplementares_sem_titular <- sum(
      !chaves_suplementares %in% titular$.CHAVE_CERTIFICADO
    )
    titulares_sem_cartao <- sum(
      is.na(titular$`NUMERO DO CARTAO`) |
        trimws(titular$`NUMERO DO CARTAO`) == ""
    )
    titulares_sem_nome_mae <- sum(
      is.na(titular$`NOME DA MAE`) | trimws(titular$`NOME DA MAE`) == ""
    )
    avisos <- c(
      if (length(titular_info$avisos) > 0) {
        paste0("TITULAR ausente: ", paste(titular_info$avisos, collapse = ", "))
      },
      if (length(dependente_info$avisos) > 0) {
        paste0("DEPENDENTE ausente: ", paste(dependente_info$avisos, collapse = ", "))
      },
      if (titulares_sem_suplemento > 0) {
        paste0("Titulares sem registro na aba suplementar: ", titulares_sem_suplemento)
      },
      if (suplementares_sem_titular > 0) {
        paste0("Certificados suplementares sem titular correspondente: ",
               suplementares_sem_titular)
      },
      if (titulares_sem_cartao > 0) {
        paste0("Titulares sem CARTAO apos a juncao: ", titulares_sem_cartao)
      },
      if (titulares_sem_nome_mae > 0) {
        paste0("Titulares sem NOME DA MAE apos a juncao: ", titulares_sem_nome_mae)
      }
    )
    
    list(
      status = "OK",
      mensagem = "Processado com sucesso.",
      avisos = if (length(avisos) > 0) paste(avisos, collapse = " | ") else NA_character_,
      observacoes_datas = observacoes_datas,
      divergencias_datas = divergencias_datas,
      final = final_arquivo,
      pendencias = pendencias,
      n_titulares = nrow(titular),
      n_dependentes = nrow(dependente),
      codigo_empresa = empresa_info$codigo,
      empresa = empresa_info$empresa
    )
  }, error = function(e) {
    list(
      status = "ERRO",
      mensagem = conditionMessage(e),
      avisos = NA_character_,
      observacoes_datas = observacoes_datas,
      divergencias_datas = divergencias_datas,
      final = NULL,
      pendencias = NULL,
      n_titulares = NA_integer_,
      n_dependentes = NA_integer_,
      codigo_empresa = NA_character_,
      empresa = NA_character_
    )
  })
  
  resultado$arquivo <- arquivo
  resultado$referencia <- referencia
  resultado
}


# ----------------------------------------------------------------------------
# 7. PROCESSAMENTO DA PASTA E CONSOLIDACAO
# ----------------------------------------------------------------------------

base_final_vazia <- function() {
  base <- tibble::tibble(
    REFERENCIA = character(),
    EMPRESA = character(),
    SubF = character(),
    `NUMERO DO CERTIFICADO` = character(),
    cod_unico = character(),
    TITULARIDADE = character(),
    `NOME DO SEGURADO` = character(),
    `NUMERO DO CARTAO` = character(),
    `NOME DA MAE` = character(),
    `NOME DEPENDENTE` = character(),
    `CODIGO DO DEPENDENTE` = character(),
    `GRAU PARENTESCO` = character(),
    PLANO = character()
  )
  names(base)[names(base) == "cod_unico"] <- "C\u00d3DIGO UNICO"
  base
}


pendencias_vazias <- function() {
  tibble::tibble(
    REFERENCIA = character(),
    EMPRESA = character(),
    `NUMERO DA SUBFATURA` = character(),
    `NUMERO DO CERTIFICADO` = character(),
    `CODIGO DO DEPENDENTE` = character(),
    `NOME DEPENDENTE` = character(),
    `NOME DO SEGURADO` = character(),
    PLANO = character(),
    MOTIVO = character()
  )
}


processar_pasta <- function(pasta, mapa_empresas) {
  if (!dir.exists(pasta)) {
    stop("A pasta de entrada nao existe: ", pasta, call. = FALSE)
  }
  
  # Sys.glob funciona melhor que list.files em algumas sessoes Windows com
  # pastas cujos nomes contem acentos. As duas extensoes cobrem maiusculas.
  arquivos <- unique(c(
    Sys.glob(file.path(pasta, "*.xlsx")),
    Sys.glob(file.path(pasta, "*.XLSX"))
  ))
  info_arquivos <- file.info(arquivos)
  arquivos <- arquivos[!is.na(info_arquivos$isdir) & !info_arquivos$isdir]
  arquivos <- arquivos[!grepl("^~\\$", basename(arquivos))]
  nomes_excluir <- tolower(c(CFG$nome_saida, CFG$nomes_saida_legados))
  arquivos <- arquivos[!(tolower(basename(arquivos)) %in% nomes_excluir)]
  # Ignora consolidacoes mensais anteriores, inclusive se o nome de saida mudar.
  arquivos <- arquivos[!grepl("^PC_CONSOLIDADO_.*\\.xlsx$", basename(arquivos),
                              ignore.case = TRUE)]
  arquivos <- sort(arquivos)
  
  if (length(arquivos) == 0) {
    stop("Nenhum arquivo .xlsx de entrada encontrado em: ", pasta, call. = FALSE)
  }
  
  cat("Arquivos de entrada encontrados:", length(arquivos), "\n\n")
  resultados <- purrr::map(arquivos, function(arq) {
    cat("Processando:", basename(arq), "\n")
    processar_arquivo(arq, mapa_empresas)
  })
  
  # Adiciona um arquivo por vez para registrar no log uma eventual incompatibilidade.
  aceitos <- rep(FALSE, length(resultados))
  consolidado <- NULL
  for (i in seq_along(resultados)) {
    r <- resultados[[i]]
    if (r$status != "OK") next
    
    tentativa <- tryCatch(
      list(ok = TRUE, dados = bind_rows(consolidado, r$final)),
      error = function(e) list(ok = FALSE, mensagem = conditionMessage(e))
    )
    if (tentativa$ok) {
      consolidado <- tentativa$dados
      aceitos[i] <- TRUE
    } else {
      r$status <- "ERRO"
      r$mensagem <- paste(
        "Incompatibilidade ao consolidar este arquivo com os anteriores:",
        tentativa$mensagem
      )
      r$final <- NULL
      r$pendencias <- NULL
      r$n_titulares <- NA_integer_
      r$n_dependentes <- NA_integer_
      resultados[[i]] <- r
    }
  }
  
  if (is.null(consolidado)) consolidado <- base_final_vazia()
  consolidado <- consolidado |>
    relocate(any_of(CFG$ordem_inicial)) |>
    arrange(REFERENCIA, SubF, `NUMERO DO CERTIFICADO`)
  
  pendencias <- if (any(aceitos)) {
    bind_rows(purrr::map(resultados[aceitos], "pendencias"))
  } else {
    pendencias_vazias()
  }
  
  log_df <- purrr::map_dfr(resultados, function(r) {
    observacoes <- resumir_observacoes_datas(r$observacoes_datas)
    divergencias <- resumir_observacoes_datas(r$divergencias_datas)
    tibble::tibble(
      ARQUIVO = r$arquivo,
      REFERENCIA = r$referencia,
      CODIGO_EMPRESA = r$codigo_empresa,
      EMPRESA = r$empresa,
      STATUS = r$status,
      TITULARES = r$n_titulares,
      DEPENDENTES = r$n_dependentes,
      AVISOS_ESTRUTURA = r$avisos,
      DATAS_INVALIDAS = length(r$observacoes_datas),
      OBSERVACOES_DATAS = observacoes,
      QTD_DIVERGENCIAS_DATAS = length(r$divergencias_datas),
      DIVERGENCIAS_DATAS = divergencias,
      MENSAGEM = r$mensagem,
      OBSERVACAO_IDENTIFICADORES =
        "Identificadores lidos como texto; zeros ausentes na origem nao foram reconstruidos."
    )
  })
  
  list(consolidado = consolidado, pendencias = pendencias, log = log_df)
}


# ----------------------------------------------------------------------------
# 8. FORMATACAO E EXPORTACAO DO EXCEL
# ----------------------------------------------------------------------------

adicionar_aba_formatada <- function(wb, df, nome_aba, nome_tabela) {
  if (ncol(df) == 0) {
    stop(sprintf("A aba '%s' nao tem colunas para exportar.", nome_aba),
         call. = FALSE)
  }
  addWorksheet(wb, sheetName = nome_aba, gridLines = FALSE)
  
  estilo_direita <- createStyle(
    halign = "right",
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
  estilo_corpo <- createStyle(valign = "center")
  estilo_data <- createStyle(
    numFmt = "dd/mm/yyyy", halign = "center", valign = "center"
  )
  estilo_centralizado <- createStyle(halign = "center", valign = "center")
  
  # Em todas as abas, os cabecalhos ficam na linha 1 e os dados comecam na 2.
  if (nrow(df) > 0) {
    writeDataTable(
      wb, sheet = nome_aba, x = df,
      startRow = 1, startCol = 1,
      tableName = nome_tabela,
      tableStyle = "TableStyleMedium2",
      withFilter = TRUE
    )
  } else {
    writeData(wb, sheet = nome_aba, x = df, startRow = 1, startCol = 1)
    addFilter(wb, sheet = nome_aba, rows = 1, cols = seq_len(ncol(df)))
  }
  
  addStyle(
    wb, sheet = nome_aba, style = estilo_cabecalho,
    rows = 1, cols = seq_len(ncol(df)), gridExpand = TRUE, stack = TRUE
  )
  setColWidths(wb, sheet = nome_aba, cols = seq_len(ncol(df)), widths = "auto")
  
  for (nome_coluna in names(CFG$largura_manual)) {
    idx <- match(nome_coluna, names(df))
    if (!is.na(idx)) {
      setColWidths(
        wb, sheet = nome_aba, cols = idx,
        widths = unname(CFG$largura_manual[[nome_coluna]])
      )
    }
  }
  
  # A mascara dd/mm/yyyy precisa de largura minima para nao aparecer como ####.
  colunas_data_cabecalho <- which(grepl("DATA", names(df), ignore.case = TRUE))
  if (length(colunas_data_cabecalho) > 0) {
    setColWidths(
      wb, sheet = nome_aba, cols = colunas_data_cabecalho, widths = 13
    )
  }
  
  setRowHeights(wb, sheet = nome_aba, rows = 1, heights = 32)
  
  if (nrow(df) > 0) {
    linhas_dados <- seq.int(2, nrow(df) + 1)
    addStyle(
      wb, sheet = nome_aba, style = estilo_corpo,
      rows = linhas_dados, cols = seq_len(ncol(df)),
      gridExpand = TRUE, stack = TRUE
    )
    
    colunas_data <- which(
      grepl("DATA", names(df), ignore.case = TRUE) &
        vapply(df, function(x) inherits(x, "Date") || inherits(x, "POSIXt"),
               logical(1))
    )
    if (length(colunas_data) > 0) {
      addStyle(
        wb, sheet = nome_aba, style = estilo_data,
        rows = linhas_dados, cols = colunas_data,
        gridExpand = TRUE, stack = TRUE
      )
    }
    
    
    col_codigo_dependente <- match("CODIGO DO DEPENDENTE", names(df))
    
    if (!is.na(col_codigo_dependente)) {
      addStyle(
        wb,
        sheet = nome_aba,
        style = estilo_direita,
        rows = linhas_dados,
        cols = col_codigo_dependente,
        gridExpand = TRUE,
        stack = TRUE
      )
    }
    
    
    
    
    colunas_centralizadas <- match(
      c(
        "EMPRESA", "SubF", "NUMERO DO CERTIFICADO", "C\u00d3DIGO UNICO",
        "TITULARIDADE", "NUMERO DO CARTAO", "PLANO", "REGIAO",
        "GRAU PARENTESCO",
        "SEXO", "ESTADO CIVIL", "NUMERO DO CPF", "CARGO DE OCUPACAO",
        "NUMERO DA MATRICULA"
      ),
      names(df)
    )
    colunas_centralizadas <- colunas_centralizadas[!is.na(colunas_centralizadas)]
    if (length(colunas_centralizadas) > 0) {
      addStyle(
        wb, sheet = nome_aba, style = estilo_centralizado,
        rows = linhas_dados, cols = colunas_centralizadas,
        gridExpand = TRUE, stack = TRUE
      )
    }
  }
  
  freezePane(wb, sheet = nome_aba, firstRow = TRUE)
}


# ----------------------------------------------------------------------------
# 9. EXECUCAO
# ----------------------------------------------------------------------------

mapa_empresas <- carregar_mapa_empresas(CFG$arquivo_mapa_empresas)
resultado <- processar_pasta(CFG$pasta, mapa_empresas)

arquivo_saida <- file.path(CFG$pasta, CFG$nome_saida)
wb <- createWorkbook()

adicionar_aba_formatada(
  wb, resultado$consolidado,
  nome_aba = "Base Tratada",
  nome_tabela = "TabelaBaseTratada"
)
adicionar_aba_formatada(
  wb, resultado$pendencias,
  nome_aba = "Pendencias",
  nome_tabela = "TabelaPendencias"
)
adicionar_aba_formatada(
  wb, resultado$log,
  nome_aba = "Log",
  nome_tabela = "TabelaLog"
)

saveWorkbook(wb, arquivo_saida, overwrite = TRUE)

arquivos_ok <- sum(resultado$log$STATUS == "OK")
arquivos_erro <- sum(resultado$log$STATUS == "ERRO")
cat("\n====================================\n")
cat("RESUMO DO PROCESSAMENTO\n")
cat("====================================\n")
cat("Arquivos processados com sucesso:", arquivos_ok, "\n")
cat("Arquivos com erro:", arquivos_erro, "\n")
cat("Registros consolidados:", nrow(resultado$consolidado), "\n")
cat("Pendencias encontradas:", nrow(resultado$pendencias), "\n")
cat("Valores de data invalidos registrados:",
    sum(resultado$log$DATAS_INVALIDAS, na.rm = TRUE), "\n")
cat("Divergencias entre datas principais e Y2K:",
    sum(resultado$log$QTD_DIVERGENCIAS_DATAS, na.rm = TRUE), "\n")
if (arquivos_erro > 0) {
  cat("\nArquivos com erro:\n")
  erros <- resultado$log |>
    filter(STATUS == "ERRO") |>
    select(ARQUIVO, MENSAGEM)
  for (i in seq_len(nrow(erros))) {
    cat(" -", erros$ARQUIVO[i], ":", erros$MENSAGEM[i], "\n")
  }
}
cat("====================================\n")
cat("Arquivo gerado:", arquivo_saida, "\n")
cat("Abas: Base Tratada, Pendencias, Log\n")

