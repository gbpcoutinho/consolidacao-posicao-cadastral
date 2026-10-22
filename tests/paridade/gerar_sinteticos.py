"""Gera um conjunto de planilhas sinteticas que cobre as regras do tratamento."""
from datetime import datetime
from pathlib import Path

import openpyxl

from tests.fabrica import (
    DEPENDENTE_OK,
    SUPLEMENTAR_OK,
    TITULAR_OK,
    criar_arquivo_pc,
    escrever_csv_empresas,
    forcar_dimensao,
)


def _com(registro: dict, **valores) -> dict:
    return {**registro, **valores}


def gerar(pasta: Path) -> Path:
    pasta = Path(pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    escrever_csv_empresas(pasta / "COD_EMPRESAS.csv", [("ACME SAÚDE", "123"), ("BETA", "456")])

    criar_arquivo_pc(pasta / "PC123 OK.xlsx", TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    criar_arquivo_pc(
        pasta / "PC123 ABA FALTANDO.xlsx", TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK,
        abas=("TITULAR", "DEPENDENTE", "OUTRA"),
    )
    criar_arquivo_pc(
        pasta / "PC123 DUPLICADO.xlsx", [TITULAR_OK[0], TITULAR_OK[0]], DEPENDENTE_OK, SUPLEMENTAR_OK
    )

    titulares_datas = [
        _com(TITULAR_OK[0], **{"DATA DE NASCIMENTO": "05/05/20",
                               "DATA DE NASCIMENTO Y2K": datetime(1920, 5, 5)}),
        _com(TITULAR_OK[1], **{"DATA DE NASCIMENTO": "01/01/90",
                               "DATA DE NASCIMENTO Y2K": "02/01/1990"}),
        _com(TITULAR_OK[1], **{"NUMERO DO CERTIFICADO": "003", "DATA DE NASCIMENTO": "xx",
                               "DATA DE NASCIMENTO Y2K": None}),
        _com(TITULAR_OK[1], **{"NUMERO DO CERTIFICADO": "004", "DATA DE NASCIMENTO": 45000,
                               "DATA DE NASCIMENTO Y2K": "31/02/71"}),
    ]
    dependentes_datas = [
        _com(DEPENDENTE_OK[0], **{"DATA DE NASCIMENTO": "2010-10-10"}),
        _com(DEPENDENTE_OK[1], **{"DATA DE NASCIMENTO": "20120201"}),
    ]
    criar_arquivo_pc(pasta / "PC456 DATAS.xlsx", titulares_datas, dependentes_datas, SUPLEMENTAR_OK)

    criar_arquivo_pc(
        pasta / "PC456 CONFLITO SUP.xlsx", TITULAR_OK, DEPENDENTE_OK,
        SUPLEMENTAR_OK + [_com(SUPLEMENTAR_OK[0], **{"CARTÃO": "OUTRO"})],
    )
    criar_arquivo_pc(pasta / "PC777 SEM EMPRESA.xlsx", TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)

    titulares_numericos = [
        _com(TITULAR_OK[0], **{"NUMERO DO CPF": 12345678901, "NUMERO DA SUBFATURA": 7}),
        _com(TITULAR_OK[1], **{"NUMERO DO CPF": 98765432100, "NUMERO DA SUBFATURA": 7.0,
                               "NUMERO DO CERTIFICADO": 2}),
    ]
    dependentes_numericos = [
        _com(DEPENDENTE_OK[0], **{"NUMERO DO CERTIFICADO": "2", "CODIGO DO DEPENDENTE": 1}),
    ]
    criar_arquivo_pc(
        pasta / "PC456 NUMERICO.xlsx", titulares_numericos, dependentes_numericos, SUPLEMENTAR_OK
    )

    criar_arquivo_pc(
        pasta / "PC123 LINHAS FORMATADAS.xlsx", TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK,
        linhas_vazias_formatadas_no_fim=5,
    )

    # Achados da revisao final: <dimension> errado, NBSP e celula so com espaco.
    criar_arquivo_pc(pasta / "PC123 DIMENSAO.xlsx", TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    forcar_dimensao(pasta / "PC123 DIMENSAO.xlsx", 2, "A1")
    dependentes_espacos = [
        _com(DEPENDENTE_OK[0], **{"NUMERO DO CERTIFICADO": "001\xa0"}),
        DEPENDENTE_OK[1],
    ]
    p = pasta / "PC456 ESPACOS.xlsx"
    criar_arquivo_pc(p, TITULAR_OK, dependentes_espacos, SUPLEMENTAR_OK)
    wb = openpyxl.load_workbook(p)
    wb["DEPENDENTE"].append([" "])
    wb.save(p)
    return pasta
