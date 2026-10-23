# Leitura e tratamento de cada aba (secoes 3 e 5 do TRATAMENTO.R).
import posixpath
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path
from xml.etree import ElementTree

import openpyxl
import pandas as pd
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.datetime import from_excel, to_excel

from .cabecalhos import (
    chave_certificado,
    chave_locale,
    normalizar_nome,
    renomear_e_validar_cabecalhos,
    trimws,
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


# Celula que existe na planilha mas e NA (erro do Excel, texto vazio ou so
# espaco). Conta para a faixa lida, como no readxl, e depois vira None.
_PRESENTE_NA = object()


def _valor_bruto(celula, epoch):
    # Celulas com erro do Excel (#N/A etc.) sao NA no readxl.
    if getattr(celula, "data_type", None) == "e":
        return _PRESENTE_NA
    v = celula.value
    if isinstance(v, (int, float)) and not isinstance(v, bool) and is_date_format(celula.number_format):
        # Numero fora da faixa de datas (ex.: CPF em celula com formato de data)
        # fica como numero, como no readxl, em vez de virar erro.
        try:
            return from_excel(v, epoch)
        except (ValueError, OverflowError):
            return v
    if isinstance(v, str):
        # readxl: trim_ws = TRUE (so espaco e tab) e na = "". Celula so com
        # espacos e quebras de linha tambem e NA no readxl.
        if v.strip(" \t\r\n") == "":
            return _PRESENTE_NA
        return v.strip(" \t")
    return v


def _sem_marcador(v):
    return None if v is _PRESENTE_NA else v


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


def _nome_local(tag: str) -> str:
    # Ignora o namespace: ha geradores com prefixo (x:c) e o padrao "strict".
    return tag.rsplit("}", 1)[-1]


def _alvo(origem: str, alvo: str) -> str:
    if alvo.startswith("/"):
        return alvo.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(origem), alvo))


def _relacoes(zf: zipfile.ZipFile, origem: str) -> dict:
    rels = posixpath.join(posixpath.dirname(origem), "_rels", posixpath.basename(origem) + ".rels")
    raiz = ElementTree.fromstring(zf.read(rels))
    return {r.get("Id"): r for r in raiz if _nome_local(r.tag) == "Relationship"}


def _xml_da_aba(zf: zipfile.ZipFile, aba: str) -> str:
    """Caminho do XML da aba dentro do .xlsx, pelo nome que aparece no Excel."""
    pacote = _relacoes(zf, "")
    livro = next(
        _alvo("", r.get("Target")) for r in pacote.values()
        if r.get("Type", "").endswith("/officeDocument")
    )
    rels = _relacoes(zf, livro)
    for elem in ElementTree.fromstring(zf.read(livro)).iter():
        if _nome_local(elem.tag) == "sheet" and elem.get("name") == aba:
            rid = next(v for k, v in elem.attrib.items() if _nome_local(k) == "id")
            return _alvo(livro, rels[rid].get("Target"))
    raise KeyError(aba)


def _linha_coluna(ref: str) -> tuple:
    letras = ref.rstrip("0123456789")
    coluna = 0
    for c in letras.upper():
        coluna = coluna * 26 + ord(c) - 64
    return int(ref[len(letras):]), coluna


def varrer_celulas(caminho: Path, aba: str):
    """Percorre o XML da aba e devolve (linha, coluna, tipo, <v> bruto, tem <f>)
    de cada celula. <v> ausente vira None; <v/> vira "". Linhas e celulas sem
    o atributo r seguem a anterior, como no openpyxl e no readxl."""
    with zipfile.ZipFile(caminho) as zf, zf.open(_xml_da_aba(zf, aba)) as xml:
        linha, coluna = 0, 0
        celula = None
        for evento, elem in ElementTree.iterparse(xml, events=("start", "end")):
            nome = _nome_local(elem.tag)
            if evento == "start":
                if nome == "row":
                    linha = int(elem.get("r")) if elem.get("r") else linha + 1
                    coluna = 0
                elif nome == "c":
                    ref = elem.get("r")
                    if ref:
                        linha, coluna = _linha_coluna(ref)
                    else:
                        coluna += 1
                    celula = [elem.get("t", "n"), None, False]
                continue
            if celula is not None:
                if nome == "v":
                    celula[1] = elem.text or ""
                elif nome == "f":
                    celula[2] = True
                elif nome == "c":
                    yield (linha, coluna, *celula)
                    celula = None
            if nome in ("c", "row"):
                elem.clear()


def _presentes_sem_valor(caminho: Path, aba: str) -> list:
    # O readxl conta na faixa toda celula com <f> ou <v>, mesmo sem valor
    # (formula sem cache, <v/>). O openpyxl le essas celulas como None.
    return [
        (linha, coluna)
        for linha, coluna, _, valor, formula in varrer_celulas(caminho, aba)
        if formula or valor == ""
    ]


def _marcar_presentes(linhas: list, presentes: list) -> None:
    for linha, coluna in presentes:
        while len(linhas) < linha:
            linhas.append([])
        valores = linhas[linha - 1]
        if len(valores) < coluna:
            valores.extend([None] * (coluna - len(valores)))
        if valores[coluna - 1] is None:
            valores[coluna - 1] = _PRESENTE_NA


def abrir_planilha(caminho: Path):
    """Abre o arquivo uma vez para ler todas as abas (feche com .close())."""
    return openpyxl.load_workbook(caminho, read_only=True, data_only=True)


def ler_aba(caminho: Path, aba: str, skip: int = 0, wb=None) -> pd.DataFrame:
    proprio = wb is None
    if proprio:
        wb = abrir_planilha(caminho)
    try:
        ws = wb[aba]
        # O readxl nao usa o <dimension> gravado no arquivo; ele pode estar errado.
        ws.reset_dimensions()
        # A conversao de datas e duracoes e feita em _valor_bruto: a do openpyxl
        # transforma numeros fora da faixa em erro e duracoes em timedelta.
        wb._date_formats = set()
        wb._timedelta_formats = set()
        linhas = [[_valor_bruto(c, wb.epoch) for c in linha] for linha in ws.iter_rows()]
    finally:
        if proprio:
            wb.close()
    _marcar_presentes(linhas, _presentes_sem_valor(caminho, aba))

    # Como no readxl: o cabecalho fica na linha max(skip, 1a linha com conteudo
    # da aba inteira). Se ha conteudo antes do skip e a linha do skip esta
    # vazia, o R para aqui. Linhas vazias no fim sao ignoradas.
    nao_vazias = [i for i, l in enumerate(linhas) if any(v is not None for v in l)]
    inicio = max(skip, nao_vazias[0]) if nao_vazias else skip
    if not nao_vazias or inicio > nao_vazias[-1] or all(v is None for v in linhas[inicio]):
        raise ErroValidacao(f"A aba '{aba}' nao tem cabecalhos identificaveis.")
    linhas = linhas[inicio: nao_vazias[-1] + 1]

    colunas_usadas = [
        j for l in linhas for j, v in enumerate(l) if v is not None
    ]
    inicio, fim = min(colunas_usadas), max(colunas_usadas) + 1
    linhas = [[_sem_marcador(v) for v in (l + [None] * fim)[inicio:fim]] for l in linhas]

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


def ler_titular(caminho: Path, aba: str, referencia: str, wb=None) -> ResultadoAba:
    contexto = f"TITULAR: {referencia}"
    df = ler_aba(caminho, aba, skip=2, wb=wb)
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


def ler_dependente(caminho: Path, aba: str, referencia: str, wb=None) -> ResultadoAba:
    contexto = f"DEPENDENTE: {referencia}"
    bruto = ler_aba(caminho, aba, skip=0, wb=wb)

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
        return txt is None or trimws(txt) == ""

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
        preenchidos = [v for v in valores if v is not None and trimws(v) != ""]
        distintos = list(dict.fromkeys(trimws(v) for v in preenchidos))
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


def ler_suplementar(caminho: Path, aba: str, referencia: str, wb=None) -> ResultadoAba:
    contexto = f"POS. CADASTRAL (NOME M E): {referencia}"
    df = ler_aba(caminho, aba, skip=0, wb=wb)
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
