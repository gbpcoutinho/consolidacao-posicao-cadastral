# Formatacao e exportacao do Excel (secao 8 do TRATAMENTO.R).
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from .config import CODIGO_UNICO, LARGURA_MANUAL

COLUNAS_CENTRALIZADAS = [
    "EMPRESA", "SubF", "NUMERO DO CERTIFICADO", CODIGO_UNICO,
    "TITULARIDADE", "NUMERO DO CARTAO", "PLANO", "REGIAO",
    "GRAU PARENTESCO",
    "SEXO", "ESTADO CIVIL", "NUMERO DO CPF", "CARGO DE OCUPACAO",
    "NUMERO DA MATRICULA",
]

_ESTILO_CABECALHO = {
    "font": Font(color="FFFFFF", bold=True),
    "fill": PatternFill("solid", fgColor="4472C4"),
    "alignment": Alignment(horizontal="center", vertical="center", wrap_text=True),
    "border": Border(bottom=Side(style="thin")),
}
_ALINHAMENTO_CORPO = Alignment(vertical="center")
_ALINHAMENTO_CENTRO = Alignment(horizontal="center", vertical="center")
_ALINHAMENTO_DIREITA = Alignment(horizontal="right", vertical="center")


def _texto_exibido(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (date, datetime)):
        return "00/00/0000"
    return str(v)


def _coluna_de_datas(nome: str, valores: list) -> bool:
    # No R, o estilo de data vale para colunas com DATA no nome e classe Date.
    return "DATA" in nome.upper() and any(isinstance(v, date) for v in valores) and all(
        v is None or isinstance(v, date) for v in valores
    )


def adicionar_aba_formatada(wb, df: pd.DataFrame, nome_aba: str, nome_tabela: str) -> None:
    if len(df.columns) == 0:
        raise ValueError(f"A aba '{nome_aba}' nao tem colunas para exportar.")
    ws = wb.create_sheet(nome_aba)
    ws.sheet_view.showGridLines = False

    colunas = [str(c) for c in df.columns]
    n_linhas, n_colunas = len(df), len(colunas)
    valores = {c: df[c].tolist() for c in df.columns}

    # Em todas as abas, os cabecalhos ficam na linha 1 e os dados comecam na 2.
    ws.append(colunas)
    for i in range(n_linhas):
        ws.append([valores[c][i] for c in df.columns])

    ultima = get_column_letter(n_colunas)
    if n_linhas > 0:
        tabela = Table(displayName=nome_tabela, ref=f"A1:{ultima}{n_linhas + 1}")
        tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tabela)
    else:
        ws.auto_filter.ref = f"A1:{ultima}1"

    for celula in ws[1]:
        for atributo, estilo in _ESTILO_CABECALHO.items():
            setattr(celula, atributo, estilo)

    for j, nome in enumerate(colunas, start=1):
        largura = max(len(_texto_exibido(v)) for v in [nome, *valores[df.columns[j - 1]]]) + 2
        if nome in LARGURA_MANUAL:
            largura = LARGURA_MANUAL[nome]
        # A mascara dd/mm/yyyy precisa de largura minima para nao aparecer como ####.
        if "DATA" in nome.upper():
            largura = 13
        ws.column_dimensions[get_column_letter(j)].width = largura

    ws.row_dimensions[1].height = 32

    if n_linhas > 0:
        for j, coluna in enumerate(df.columns, start=1):
            nome = colunas[j - 1]
            data = _coluna_de_datas(nome, valores[coluna])
            if data:
                alinhamento = _ALINHAMENTO_CENTRO
            elif nome == "CODIGO DO DEPENDENTE":
                alinhamento = _ALINHAMENTO_DIREITA
            elif nome in COLUNAS_CENTRALIZADAS:
                alinhamento = _ALINHAMENTO_CENTRO
            else:
                alinhamento = _ALINHAMENTO_CORPO
            for (celula,) in ws.iter_rows(min_row=2, max_row=n_linhas + 1, min_col=j, max_col=j):
                celula.alignment = alinhamento
                if data:
                    celula.number_format = "dd/mm/yyyy"

    ws.freeze_panes = "A2"


def escrever_excel(c, destino: Path) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    adicionar_aba_formatada(wb, c.consolidado, "Base Tratada", "TabelaBaseTratada")
    adicionar_aba_formatada(wb, c.pendencias, "Pendencias", "TabelaPendencias")
    adicionar_aba_formatada(wb, c.log, "Log", "TabelaLog")
    wb.save(Path(destino))
