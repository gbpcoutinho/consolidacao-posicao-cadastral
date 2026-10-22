# Leitura e tratamento de cada aba (secoes 3 e 5 do TRATAMENTO.R).
from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.utils.datetime import to_excel

from .cabecalhos import (
    chave_certificado,
    chave_locale,
    normalizar_nome,
    renomear_e_validar_cabecalhos,
    validar_colunas,
    validar_identificadores_como_texto,
)
from .config import (
    COLUNAS_OBRIGATORIAS_DEPENDENTE,
    COLUNAS_OBRIGATORIAS_SUPLEMENTAR,
    COLUNAS_OBRIGATORIAS_TITULAR,
    COLUNAS_OPCIONAIS_DEPENDENTE,
    COLUNAS_OPCIONAIS_TITULAR,
)
from .datas import texto_celula, validar_colunas_data
from .erros import ErroValidacao


@dataclass
class ResultadoAba:
    dados: pd.DataFrame
    avisos: list = field(default_factory=list)
    observacoes_datas: list = field(default_factory=list)
    divergencias_datas: list = field(default_factory=list)


def _valor_bruto(celula):
    # Celulas com erro do Excel (#N/A etc.) sao NA no readxl.
    if getattr(celula, "data_type", None) == "e":
        return None
    v = celula.value
    if isinstance(v, str):
        # readxl: trim_ws = TRUE e na = "".
        v = v.strip()
        return v if v != "" else None
    return v


def _numero_como_texto(x) -> str:
    if float(x).is_integer():
        return str(int(x))
    return format(float(x), ".15g")


def _como_texto(v):
    # Equivalente a col_types = "text" do readxl.
    if v is None or isinstance(v, str):
        return v
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return _numero_como_texto(v)
    if isinstance(v, (datetime, date, time)):
        return _numero_como_texto(to_excel(v))
    return str(v)


def ler_aba(caminho: Path, aba: str, skip: int = 0) -> pd.DataFrame:
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    try:
        linhas = [
            [_valor_bruto(c) for c in linha]
            for linha in wb[aba].iter_rows()
        ]
    finally:
        wb.close()

    linhas = linhas[skip:]
    # Linhas vazias no inicio e no fim sao ignoradas, como no readxl.
    nao_vazias = [i for i, l in enumerate(linhas) if any(v is not None for v in l)]
    if not nao_vazias:
        raise ErroValidacao(f"A aba '{aba}' nao tem cabecalhos identificaveis.")
    linhas = linhas[nao_vazias[0]: nao_vazias[-1] + 1]

    colunas_usadas = [
        j for l in linhas for j, v in enumerate(l) if v is not None
    ]
    inicio, fim = min(colunas_usadas), max(colunas_usadas) + 1
    linhas = [(l + [None] * fim)[inicio:fim] for l in linhas]

    cabecalho = ["" if v is None else _como_texto(v) for v in linhas[0]]
    colunas = {}
    for j, nome in enumerate(cabecalho):
        valores = [l[j] for l in linhas[1:]]
        if "DATA" not in normalizar_nome(nome):
            valores = [_como_texto(v) for v in valores]
        colunas[j] = valores

    df = pd.DataFrame({j: pd.Series(v, dtype=object) for j, v in colunas.items()})
    if df.empty and len(df.columns) == 0:
        df = pd.DataFrame(columns=range(len(cabecalho)), dtype=object)
    df.columns = cabecalho
    return df


def _adicionar_coluna(df: pd.DataFrame, nome: str, valores) -> None:
    if not isinstance(valores, (list, pd.Series)):
        valores = [valores] * len(df)
    df[nome] = pd.Series(list(valores), index=df.index, dtype=object)


def ler_titular(caminho: Path, aba: str, referencia: str) -> ResultadoAba:
    contexto = f"TITULAR: {referencia}"
    df = ler_aba(caminho, aba, skip=2)
    regras = {
        "REATIVA": "DATA DE REATIVACAO",
        "REGI.O": "REGIAO",
        "OCUPA": "CARGO DE OCUPACAO",
        "SEXO.*SEGURADO": "SEXO",
    }
    df = renomear_e_validar_cabecalhos(df, regras, contexto)
    df = validar_identificadores_como_texto(df, contexto)
    df, observacoes, divergencias = validar_colunas_data(df, contexto)
    validar_colunas(df, COLUNAS_OBRIGATORIAS_TITULAR, contexto)

    chaves = chave_certificado(df["NUMERO DO CERTIFICADO"])
    if any(c is None for c in chaves):
        raise ErroValidacao(
            f"[{contexto}] Ha titular sem NUMERO DO CERTIFICADO; "
            "nao e possivel relacionar os dados."
        )

    # Regra confirmada: titular recebe CODIGO DO DEPENDENTE = "00".
    _adicionar_coluna(df, "TITULARIDADE", "TITULAR")
    _adicionar_coluna(df, "CODIGO DO DEPENDENTE", "00")
    _adicionar_coluna(df, "REFERENCIA", referencia)
    _adicionar_coluna(df, ".CHAVE_CERTIFICADO", chaves)

    return ResultadoAba(
        dados=df,
        avisos=[c for c in COLUNAS_OPCIONAIS_TITULAR if c not in df.columns],
        observacoes_datas=observacoes,
        divergencias_datas=divergencias,
    )


def ler_dependente(caminho: Path, aba: str, referencia: str) -> ResultadoAba:
    contexto = f"DEPENDENTE: {referencia}"
    bruto = ler_aba(caminho, aba, skip=0)

    if len(bruto) <= 2:
        raise ErroValidacao(
            f"[{contexto}] Ha {len(bruto)} linha(s); nao existem dados apos remover "
            "as duas linhas finais de resumo."
        )

    # As duas linhas finais sao sintese/resumo, nao registros nem erros de layout.
    df = bruto.iloc[: len(bruto) - 2].copy()
    regras = {
        "REATIVA": "DATA DE REATIVACAO",
        "SEXO.*DEPENDENTE": "SEXO",
        "ESTADO.*CIVIL": "ESTADO CIVIL",
        "CART.*DEPENDENTE": "NUMERO DO CARTAO DO DEPENDENTE",
        "NOME.*M.E.*DEPENDENTE": "NOME DA MAE DO DEPENDENTE",
    }
    df = renomear_e_validar_cabecalhos(df, regras, contexto)
    df = validar_identificadores_como_texto(df, contexto)
    df, observacoes, divergencias = validar_colunas_data(df, contexto)
    validar_colunas(df, COLUNAS_OBRIGATORIAS_DEPENDENTE, contexto)

    sobrepostas = [c for c in ("NOME DO SEGURADO", "PLANO") if c in df.columns]
    if sobrepostas:
        raise ErroValidacao(
            f"[{contexto}] Campo(s) que devem vir do titular tambem aparecem na aba "
            "dependente: " + ", ".join(sobrepostas)
        )

    if "NUMERO DA SUBFATURA" in df.columns:
        df = df.rename(columns={"NUMERO DA SUBFATURA": ".SUBF_DEPENDENTE"})
    else:
        _adicionar_coluna(df, ".SUBF_DEPENDENTE", None)

    _adicionar_coluna(df, "TITULARIDADE", "DEPENDENTE")
    _adicionar_coluna(df, "REFERENCIA", referencia)
    _adicionar_coluna(df, ".CHAVE_CERTIFICADO", chave_certificado(df["NUMERO DO CERTIFICADO"]))

    return ResultadoAba(
        dados=df,
        avisos=[c for c in COLUNAS_OPCIONAIS_DEPENDENTE if c not in df.columns],
        observacoes_datas=observacoes,
        divergencias_datas=divergencias,
    )


def linhas_vazias(df: pd.DataFrame) -> list:
    def vazia(v):
        txt = texto_celula(v)
        return txt is None or txt.strip() == ""

    return [all(vazia(v) for v in linha) for linha in df.itertuples(index=False)]


def agregar_aba_suplementar(df: pd.DataFrame, contexto: str) -> pd.DataFrame:
    df = df[[not v for v in linhas_vazias(df)]]
    if len(df) == 0:
        raise ErroValidacao(f"[{contexto}] A aba suplementar nao contem registros.")

    validar_colunas(df, COLUNAS_OBRIGATORIAS_SUPLEMENTAR, contexto)
    df = validar_identificadores_como_texto(df, contexto)
    chaves = chave_certificado(df["NUMERO DO CERTIFICADO"])
    if any(c is None for c in chaves):
        raise ErroValidacao(f"[{contexto}] Ha registro suplementar sem NUMERO DO CERTIFICADO.")

    # Repeticoes identicas sao agregadas em um registro. Se um certificado tiver
    # mais de um CARTAO ou NOME DA MAE diferente, o arquivo falha.
    grupos = {}
    for chave, cartao, mae in zip(chaves, df["CARTAO"].tolist(), df["NOME DA MAE"].tolist()):
        grupo = grupos.setdefault(chave, {"CARTAO": [], "NOME DA MAE": []})
        grupo["CARTAO"].append(cartao)
        grupo["NOME DA MAE"].append(mae)

    def valor_unico(chave, nome_coluna, valores):
        valores = [None if v is None else texto_celula(v) for v in valores]
        preenchidos = [v for v in valores if v is not None and v.strip() != ""]
        distintos = list(dict.fromkeys(v.strip() for v in preenchidos))
        if len(distintos) > 1:
            raise ErroValidacao(
                f"[{contexto}] O certificado '{chave}' tem valores conflitantes em "
                f"'{nome_coluna}': " + " | ".join(distintos)
            )
        return preenchidos[0] if preenchidos else None

    linhas = []
    # split() do R ordena os grupos pela chave.
    for chave in sorted(grupos, key=chave_locale):
        g = grupos[chave]
        linhas.append(
            {
                ".CHAVE_CERTIFICADO": chave,
                ".CARTAO_SUP": valor_unico(chave, "CARTAO", g["CARTAO"]),
                ".NOME_MAE_SUP": valor_unico(chave, "NOME DA MAE", g["NOME DA MAE"]),
            }
        )
    return pd.DataFrame(linhas, columns=[".CHAVE_CERTIFICADO", ".CARTAO_SUP", ".NOME_MAE_SUP"], dtype=object)


def ler_suplementar(caminho: Path, aba: str, referencia: str) -> ResultadoAba:
    contexto = f"POS. CADASTRAL (NOME M E): {referencia}"
    df = ler_aba(caminho, aba, skip=0)
    regras = {"^CART.?O$": "CARTAO", "^NOME_DA_M.?E$": "NOME DA MAE"}
    df = renomear_e_validar_cabecalhos(df, regras, contexto)
    df, observacoes, divergencias = validar_colunas_data(df, contexto)
    agregado = agregar_aba_suplementar(df, contexto)
    return ResultadoAba(
        dados=agregado,
        avisos=[],
        observacoes_datas=observacoes,
        divergencias_datas=divergencias,
    )
