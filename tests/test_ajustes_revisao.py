"""Achados do debate dos 3 revisores (2026-10-22)."""
from datetime import date

import openpyxl

from consolidacao.leitura import ler_aba
from consolidacao.datas import validar_colunas_data


def _planilha_formatada(caminho, cabecalho, valores, formato):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "A"
    ws.append(cabecalho)
    ws.append(valores)
    for celula in ws[2]:
        celula.number_format = formato
    wb.save(caminho)


def test_a1_numero_grande_em_celula_de_data_nao_some(tmp_path):
    p = tmp_path / "x.xlsx"
    _planilha_formatada(p, ["NUMERO DO CPF", "DATA DE NASCIMENTO"], [12345678901, 19900512], "dd/mm/yyyy")
    df = ler_aba(p, "A")
    assert df["NUMERO DO CPF"].tolist() == ["12345678901"]
    _, obs, _ = validar_colunas_data(df, "C")
    assert len(obs) == 1 and "19900512" in obs[0]


def test_a1_data_normal_continua_data(tmp_path):
    p = tmp_path / "x.xlsx"
    _planilha_formatada(p, ["DATA DE NASCIMENTO"], [45000], "dd/mm/yyyy")
    df, obs, _ = validar_colunas_data(ler_aba(p, "A"), "C")
    assert df["DATA DE NASCIMENTO"].tolist() == [date(2023, 3, 15)] and obs == []


def test_a2_duracao_em_coluna_de_texto_vira_numero(tmp_path):
    p = tmp_path / "x.xlsx"
    _planilha_formatada(p, ["HORAS"], [1.5], "[h]:mm:ss")
    assert ler_aba(p, "A")["HORAS"].tolist() == ["1.5"]
