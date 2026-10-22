from datetime import date, datetime

import pandas as pd

from consolidacao.datas import (
    SEM_FORMATO,
    converter_celula_data,
    converter_data_com_ano_curto,
    converter_serial_excel,
    parse_formato_r,
    resumir_observacoes_datas,
    validar_colunas_data,
)


def test_serial_excel():
    assert converter_serial_excel(45000) == date(2023, 3, 15)
    assert converter_serial_excel(59) == date(1900, 2, 28)
    assert converter_serial_excel(61) == date(1900, 3, 1)
    assert converter_serial_excel(60) is None
    assert converter_serial_excel(0) is None
    assert converter_serial_excel(2958466) is None
    assert converter_serial_excel(45000.9) == date(2023, 3, 15)


def test_ano_curto():
    assert converter_data_com_ano_curto("05/05/71") == date(1971, 5, 5)
    assert converter_data_com_ano_curto("5-5-29") == date(2029, 5, 5)
    assert converter_data_com_ano_curto("05.05.30 10:00") == date(1930, 5, 5)
    assert converter_data_com_ano_curto("31/02/71") is None
    assert converter_data_com_ano_curto("05/05/1971") is SEM_FORMATO


def test_parse_formato_r_leniente():
    assert parse_formato_r("05/05/2020 10:00", "%d/%m/%Y") == date(2020, 5, 5)
    assert parse_formato_r("2020-1-5", "%Y-%m-%d") == date(2020, 1, 5)
    assert parse_formato_r("20200105", "%Y%m%d") == date(2020, 1, 5)
    assert parse_formato_r("31/02/2020", "%d/%m/%Y") is None


def test_converter_celula_data_tipos():
    assert converter_celula_data(datetime(2020, 1, 2, 13, 0), "D", 1, "C") == (date(2020, 1, 2), None)
    assert converter_celula_data(45000, "D", 1, "C") == (date(2023, 3, 15), None)
    assert converter_celula_data(" 45000 ", "D", 1, "C") == (date(2023, 3, 15), None)
    assert converter_celula_data("45000,0", "D", 1, "C") == (date(2023, 3, 15), None)
    assert converter_celula_data(None, "D", 1, "C") == (None, None)
    assert converter_celula_data("  ", "D", 1, "C") == (None, None)
    assert converter_celula_data("abc", "D", 3, "CTX") == (
        None,
        "CTX | coluna 'D' | registro 3 | valor invalido 'abc' (convertido para vazio)",
    )
    assert converter_celula_data(-5, "D", 1, "C")[1].endswith(
        "valor invalido '-5' (convertido para vazio)"
    )


def test_y2k_substitui_quando_dia_mes_coincidem_e_registra_divergencia():
    df = pd.DataFrame(
        {
            "DATA DE NASCIMENTO": ["05/05/20", "01/01/90"],
            "DATA DE NASCIMENTO Y2K": ["05/05/1920", "02/01/1990"],
        }
    )
    out, obs, div = validar_colunas_data(df, "CTX")
    assert out["DATA DE NASCIMENTO"].tolist() == [date(1920, 5, 5), date(1990, 1, 1)]
    assert obs == []
    assert div == [
        "CTX | registro 2 | divergencia entre 'DATA DE NASCIMENTO' (01/01/1990) e "
        "'DATA DE NASCIMENTO Y2K' (02/01/1990); mantida a data principal"
    ]


def test_resumir_observacoes():
    obs = [
        f"C | coluna 'D' | registro {i} | valor invalido 'x' (convertido para vazio)"
        for i in range(1, 7)
    ]
    r = resumir_observacoes_datas(obs)
    assert r.startswith("C | coluna 'D': 6 valor(es); amostras: registro 1 |")
    assert r.endswith("; +2 caso(s)")
    assert resumir_observacoes_datas([]) is None


def test_resumir_ordena_grupos_como_locale_do_r():
    # split() do R usa a collation do Windows: espaco/pontuacao nao decidem.
    obs = [
        "C | coluna 'DATA Y2K' | registro 1 | valor invalido 'x' (convertido para vazio)",
        "C | coluna 'DATA' | registro 2 | valor invalido 'y' (convertido para vazio)",
    ]
    assert resumir_observacoes_datas(obs).startswith("C | coluna 'DATA': 1 valor(es)")
