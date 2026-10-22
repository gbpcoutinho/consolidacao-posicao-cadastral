"""Faixa de leitura e espacos: casos achados na revisao final (comparados com o R)."""
import openpyxl
import pandas as pd
import pytest

from consolidacao.cabecalhos import chave_certificado
from consolidacao.leitura import ler_aba, ler_dependente
from tests.fabrica import (
    DEPENDENTE_OK,
    SUPLEMENTAR_OK,
    TITULAR_OK,
    criar_arquivo_pc,
    forcar_dimensao,
)


def _planilha(caminho, linhas):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "A"
    for linha in linhas:
        ws.append(linha)
    wb.save(caminho)


@pytest.mark.parametrize("ref", ["A1:H4", "A1"])
def test_dependente_ignora_dimension_errada(tmp_path, ref):
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    forcar_dimensao(p, 2, ref)
    assert len(ler_dependente(p, "DEPENDENTE", "PC1").dados) == 2


def test_celula_so_com_espaco_conta_na_faixa(tmp_path):
    p = tmp_path / "x.xlsx"
    _planilha(p, [["A", "B"], ["1", "2"], ["3", "4"], [" "]])
    df = ler_aba(p, "A")
    assert len(df) == 3
    assert df["A"].tolist() == ["1", "3", None]


def test_cabecalho_so_com_espaco_e_coluna_sem_nome(tmp_path):
    p = tmp_path / "x.xlsx"
    _planilha(p, [["A", "B", " "], ["1", "2", None]])
    df = ler_aba(p, "A")
    assert list(df.columns) == ["A", "B", ""]


def test_dependente_com_espaco_abaixo_do_resumo(tmp_path):
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    wb = openpyxl.load_workbook(p)
    wb["DEPENDENTE"].append([" "])
    wb.save(p)
    # Como no R: saem a linha do espaco e a da contagem; "TOTAL" vira registro.
    assert len(ler_dependente(p, "DEPENDENTE", "PC1").dados) == 3


def test_leitura_apara_so_espaco_e_tab(tmp_path):
    p = tmp_path / "x.xlsx"
    _planilha(p, [["A", "B"], ["5\xa0", " \t7\t "], ["\xa0", "x"]])
    df = ler_aba(p, "A")
    assert df["A"].tolist() == ["5\xa0", "\xa0"]
    assert df["B"].tolist() == ["7", "x"]


def test_chave_certificado_usa_trimws_do_r():
    s = pd.Series(["001\xa0", " 002\r\n", "\xa0"], dtype=object)
    assert chave_certificado(s).tolist() == ["001\xa0", "002", "\xa0"]


def test_certificado_so_nbsp_nao_e_vazio(tmp_path):
    from consolidacao.leitura import ler_titular

    tit = [TITULAR_OK[0], dict(TITULAR_OK[1], **{"NUMERO DO CERTIFICADO": "\xa0"})]
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, tit, DEPENDENTE_OK, SUPLEMENTAR_OK)
    assert ler_titular(p, "TITULAR", "PC1").dados[".CHAVE_CERTIFICADO"].tolist() == ["001", "\xa0"]
