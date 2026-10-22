# Cabecalhos e validacoes de estrutura (secao 2 do TRATAMENTO.R).
import re
import unicodedata

import pandas as pd

from .config import PADRAO_IDENTIFICADOR
from .erros import ErroValidacao

# Mesma tabela do chartr() do R; caracteres fora dela viram "_" adiante.
_TABELA_ACENTOS = str.maketrans(
    "ÁÀÂÃÄÉÈÊËÍÌÎÏ"
    "ÓÒÔÕÖÚÙÛÜÇ",
    "AAAAAEEEEIIIIOOOOOUUUUC",
)


def normalizar_nome(x) -> str:
    if x is None:
        return ""
    x = str(x).upper().strip()
    x = x.translate(_TABELA_ACENTOS)
    x = re.sub(r"[^A-Z0-9]+", "_", x)
    return x.strip("_")


def chave_locale(texto: str) -> tuple:
    """Aproxima sort()/split() do R no Windows: letras e digitos decidem
    primeiro, sem acento nem caixa; espaco e pontuacao so desempatam."""
    sem_acento = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in sem_acento if not unicodedata.combining(c))
    principal = re.sub(r"[^0-9a-z]", "", sem_acento.casefold())
    return principal, sem_acento.casefold(), texto


def _preenchido(v) -> bool:
    return v is not None and str(v).strip() != ""


def renomear_e_validar_cabecalhos(df: pd.DataFrame, regras: dict, contexto: str) -> pd.DataFrame:
    nomes_originais = list(df.columns)
    nomes_norm = [normalizar_nome(n) for n in nomes_originais]

    idx_tipo = [i for i, n in enumerate(nomes_norm) if n == "TIPO_DE_REGISTRO"]
    if idx_tipo:
        df = df.drop(columns=[df.columns[i] for i in idx_tipo])
        nomes_originais = list(df.columns)
        nomes_norm = [normalizar_nome(n) for n in nomes_originais]

    if any(n == "" for n in nomes_norm):
        raise ErroValidacao(f"[{contexto}] Existe coluna sem nome; layout nao seguro.")

    nomes_finais = [re.sub(r"_+", " ", n).strip() for n in nomes_norm]

    for padrao, destino in regras.items():
        idx = [i for i, n in enumerate(nomes_norm) if re.search(padrao, n)]
        if len(idx) > 1:
            raise ErroValidacao(
                f"[{contexto}] O padrao '{padrao}' corresponde a varias colunas: "
                + ", ".join(str(nomes_originais[i]) for i in idx)
            )
        if len(idx) == 1:
            nomes_finais[idx[0]] = destino

    repetidos = []
    vistos = set()
    for n in nomes_finais:
        if n in vistos and n not in repetidos:
            repetidos.append(n)
        vistos.add(n)
    if repetidos:
        raise ErroValidacao(
            f"[{contexto}] A padronizacao gerou cabecalhos duplicados: " + ", ".join(repetidos)
        )

    df = df.copy()
    df.columns = nomes_finais
    return df


def validar_colunas(df: pd.DataFrame, obrigatorias: list, contexto: str) -> None:
    faltantes = [c for c in obrigatorias if c not in df.columns]
    if faltantes:
        raise ErroValidacao(
            f"[{contexto}] Coluna(s) obrigatoria(s) ausente(s): " + ", ".join(faltantes)
        )


def validar_identificadores_como_texto(df: pd.DataFrame, contexto: str) -> pd.DataFrame:
    # No R, colunas com DATA no nome nunca sao character (sao lidas como list
    # e convertidas em Date), entao um identificador com DATA sempre falha.
    for nome in df.columns:
        norm = normalizar_nome(nome)
        if not re.search(PADRAO_IDENTIFICADOR, norm):
            continue
        texto = "DATA" not in norm and all(
            v is None or isinstance(v, str) for v in df[nome].tolist()
        )
        if not texto:
            raise ErroValidacao(
                f"[{contexto}] A coluna identificadora '{nome}' nao foi lida como texto."
            )
    return df


def chave_certificado(valores: pd.Series) -> pd.Series:
    # A chave auxiliar e aparada apenas para comparacao; o valor original da
    # coluna NUMERO DO CERTIFICADO permanece intacto na base exportada.
    def chave(v):
        if v is None:
            return None
        v = str(v).strip()
        return v if v != "" else None

    return pd.Series([chave(v) for v in valores.tolist()], index=valores.index, dtype=object)


def unificar_colunas(df: pd.DataFrame, nome_final: str, fontes: list, contexto: str) -> pd.DataFrame:
    # Mantem o primeiro valor preenchido na ordem: coluna final e depois aliases.
    # Valores divergentes em colunas equivalentes interrompem o arquivo para nao
    # descartar uma das informacoes silenciosamente.
    origens = list(dict.fromkeys([nome_final, *fontes]))
    presentes = [o for o in origens if o in df.columns]
    finais = [None] * len(df)

    for origem in presentes:
        valores = [None if v is None else str(v) for v in df[origem].tolist()]
        for i, v in enumerate(valores):
            if not _preenchido(v):
                continue
            if _preenchido(finais[i]):
                if finais[i].strip() != v.strip():
                    raise ErroValidacao(
                        f"[{contexto}] Valores conflitantes nas colunas equivalentes "
                        f"usadas para '{nome_final}' (registro {i + 1})."
                    )
            else:
                finais[i] = v

    df = df.copy()
    df[nome_final] = pd.Series(finais, index=df.index, dtype=object)
    remover = [p for p in presentes if p != nome_final]
    return df.drop(columns=remover)
