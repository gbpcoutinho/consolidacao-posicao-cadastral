# Mapa codigo -> empresa (secao 4 do TRATAMENTO.R).
import csv
import io
import re
from pathlib import Path

from .cabecalhos import normalizar_nome, trimws
from .erros import ErroValidacao


def _ler_texto(caminho: Path) -> str:
    try:
        dados = caminho.read_bytes()
    except OSError as e:
        # Pasta no lugar do arquivo, sem permissao ou travado por outro programa.
        raise ErroValidacao(
            f"Nao foi possivel ler o arquivo de empresas '{caminho}': confira se e um "
            "arquivo .csv e se nao esta em uso ou sem permissao de leitura."
        ) from e
    try:
        return dados.decode("utf-8-sig")
    except UnicodeDecodeError:
        pass
    try:
        # CSV salvo pelo Excel em ANSI. O R leria com acentos corrompidos.
        return dados.decode("cp1252")
    except UnicodeDecodeError as e:
        raise ErroValidacao(
            f"O arquivo de empresas '{caminho}' esta em uma codificacao nao suportada; "
            "salve-o novamente como CSV (UTF-8 ou ANSI)."
        ) from e


def carregar_mapa_empresas(caminho: Path) -> dict:
    caminho = Path(caminho)
    if not caminho.exists():
        raise ErroValidacao(f"Arquivo de empresas nao encontrado: {caminho}")

    linhas = [l for l in _ler_texto(caminho).splitlines() if l.strip() != ""]
    if not linhas:
        raise ErroValidacao("O COD_EMPRESAS.csv esta vazio.")

    registros = list(csv.reader(io.StringIO("\n".join(linhas)), delimiter=";", quotechar='"'))
    nomes_norm = [normalizar_nome(n) for n in registros[0]]
    if "EMPRESA" not in nomes_norm or "COD_SAUDE" not in nomes_norm:
        raise ErroValidacao("O COD_EMPRESAS.csv precisa conter as colunas EMPRESA e COD_SAUDE.")
    idx_empresa = nomes_norm.index("EMPRESA")
    idx_codigo = nomes_norm.index("COD_SAUDE")

    def campo(registro, i):
        valor = registro[i] if i < len(registro) else ""
        # read.csv2 trata o texto "NA" como ausente.
        return None if valor == "NA" else trimws(valor)

    pares = []
    for registro in registros[1:]:
        empresa, codigo = campo(registro, idx_empresa), campo(registro, idx_codigo)
        if empresa and codigo:
            pares.append((empresa, codigo))

    if not pares:
        raise ErroValidacao("O COD_EMPRESAS.csv nao tem pares EMPRESA/COD_SAUDE validos.")

    empresas_por_codigo = {}
    for empresa, codigo in pares:
        lista = empresas_por_codigo.setdefault(codigo, [])
        if empresa not in lista:
            lista.append(empresa)

    mapa = {}
    for codigo, empresas in empresas_por_codigo.items():
        if len(empresas) > 1:
            raise ErroValidacao(
                f"COD_SAUDE '{codigo}' esta associado a mais de uma EMPRESA: "
                + ", ".join(empresas)
            )
        mapa[codigo] = empresas[0]
    return mapa


def empresa_da_referencia(referencia: str, mapa: dict) -> tuple:
    m = re.match(r"PC([0-9]+)", referencia, flags=re.IGNORECASE)
    if not m:
        raise ErroValidacao(
            f"Nao foi possivel extrair o codigo apos 'PC' da referencia '{referencia}'."
        )
    codigo = m.group(1)
    if codigo not in mapa:
        raise ErroValidacao(
            f"Codigo '{codigo}' extraido de '{referencia}' nao encontrado em "
            "COD_SAUDE no COD_EMPRESAS.csv."
        )
    return codigo, mapa[codigo]
