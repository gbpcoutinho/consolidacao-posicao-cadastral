"""Copia congelada da leitura de celulas anterior a fase 2 (openpyxl read_only
+ segunda varredura do XML), usada como oraculo: a leitura em uma passada tem
de produzir exatamente as mesmas linhas brutas."""
from pathlib import Path

from openpyxl import load_workbook as _carregar
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.datetime import from_excel

from consolidacao.leitura import _PRESENTE_NA, _NumeroBruto, varrer_celulas


def _valor_bruto(celula, epoch):
    if getattr(celula, "data_type", None) == "e":
        return _PRESENTE_NA
    v = celula.value
    if isinstance(v, (int, float)) and not isinstance(v, bool) and is_date_format(celula.number_format):
        try:
            return from_excel(v, epoch)
        except (ValueError, OverflowError):
            return v
    if isinstance(v, str):
        if v.strip(" \t\r\n") == "":
            return _PRESENTE_NA
        return v.strip(" \t")
    return v


def _aplicar_xml(linhas: list, caminho: Path, aba: str) -> None:
    for linha, coluna, tipo, bruto, formula in varrer_celulas(caminho, aba):
        if formula and not bruto or bruto == "":
            while len(linhas) < linha:
                linhas.append([])
            valores = linhas[linha - 1]
            if len(valores) < coluna:
                valores.extend([None] * (coluna - len(valores)))
            if valores[coluna - 1] is None:
                valores[coluna - 1] = _PRESENTE_NA
        elif bruto is not None and tipo in ("n", "d"):
            try:
                v = linhas[linha - 1][coluna - 1]
            except IndexError:
                continue
            if v is not None and v is not _PRESENTE_NA:
                linhas[linha - 1][coluna - 1] = _NumeroBruto(v, bruto)


def linhas_brutas(caminho: Path, aba: str) -> list:
    # Referencia propria: testes que contam aberturas nao veem o oraculo.
    wb = _carregar(caminho, read_only=True, data_only=True)
    try:
        ws = wb[aba]
        ws.reset_dimensions()
        wb._date_formats = set()
        wb._timedelta_formats = set()
        epoch = wb.epoch
        linhas = [[_valor_bruto(c, epoch) for c in linha] for linha in ws.iter_rows()]
    finally:
        wb.close()
    _aplicar_xml(linhas, caminho, aba)
    return linhas


def normalizar(linhas: list) -> list:
    """Forma comparavel: marcador e numero bruto viram tuplas com o tipo."""
    def um(v):
        if v is _PRESENTE_NA:
            return ("NA",)
        if isinstance(v, _NumeroBruto):
            # Celula repetida no XML pode embrulhar duas vezes: compara por dentro.
            return ("bruto", um(v.valor), v.texto)
        return (type(v).__name__, v)

    return [[um(v) for v in linha] for linha in linhas]
