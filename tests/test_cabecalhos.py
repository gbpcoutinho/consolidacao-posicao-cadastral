import pandas as pd
import pytest

from consolidacao.cabecalhos import (
    normalizar_nome,
    renomear_e_validar_cabecalhos,
    unificar_colunas,
    validar_colunas,
    validar_identificadores_como_texto,
)
from consolidacao.erros import ErroValidacao


def test_normalizar_nome_casos_sujos():
    assert normalizar_nome("  Região ") == "REGIAO"
    assert normalizar_nome("Nº  do Cartão") == "N_DO_CARTAO"
    assert normalizar_nome("DATA DE NASCIMENTO Y2K") == "DATA_DE_NASCIMENTO_Y2K"
    assert normalizar_nome("__x__") == "X"
    assert normalizar_nome(None) == ""


def test_renomear_remove_tipo_de_registro_e_aplica_regras():
    df = pd.DataFrame(columns=["TIPO DE REGISTRO", "Região", "Sexo do Segurado"])
    out = renomear_e_validar_cabecalhos(
        df, {"REGI.O": "REGIAO", "SEXO.*SEGURADO": "SEXO"}, "T"
    )
    assert list(out.columns) == ["REGIAO", "SEXO"]


def test_renomear_coluna_sem_nome():
    with pytest.raises(
        ErroValidacao, match=r"^\[T\] Existe coluna sem nome; layout nao seguro\.$"
    ):
        renomear_e_validar_cabecalhos(pd.DataFrame(columns=["A", ""]), {}, "T")


def test_renomear_padrao_em_varias_colunas():
    df = pd.DataFrame(columns=["DATA REATIVACAO", "REATIVADO"])
    with pytest.raises(
        ErroValidacao,
        match=r"O padrao 'REATIVA' corresponde a varias colunas: DATA REATIVACAO, REATIVADO",
    ):
        renomear_e_validar_cabecalhos(df, {"REATIVA": "DATA DE REATIVACAO"}, "T")


def test_renomear_duplicado_apos_padronizar():
    with pytest.raises(
        ErroValidacao, match="A padronizacao gerou cabecalhos duplicados: A B"
    ):
        renomear_e_validar_cabecalhos(pd.DataFrame(columns=["A-B", "a b"]), {}, "T")


def test_validar_colunas_faltantes():
    with pytest.raises(
        ErroValidacao, match=r"Coluna\(s\) obrigatoria\(s\) ausente\(s\): X, Y"
    ):
        validar_colunas(pd.DataFrame(columns=["A"]), ["A", "X", "Y"], "T")


def test_identificador_com_data_no_nome_falha():
    df = pd.DataFrame({"DATA DO CARTAO": [None]})
    with pytest.raises(
        ErroValidacao,
        match="A coluna identificadora 'DATA DO CARTAO' nao foi lida como texto",
    ):
        validar_identificadores_como_texto(df, "T")


def test_unificar_colunas_preenche_e_remove_alias():
    df = pd.DataFrame(
        {
            "X": ["1", None, None],
            "NUMERO DO CARTAO": [None, "2", " "],
            "ALIAS": ["1", None, "4"],
        }
    )
    out = unificar_colunas(df, "NUMERO DO CARTAO", ["ALIAS"], "T")
    assert out["NUMERO DO CARTAO"].tolist() == ["1", "2", "4"]
    assert list(out.columns) == ["X", "NUMERO DO CARTAO"]


def test_unificar_colunas_conflito():
    df = pd.DataFrame({"F": ["1"], "A": ["2"]})
    with pytest.raises(
        ErroValidacao,
        match=r"Valores conflitantes nas colunas equivalentes usadas para 'F' \(registro 1\)",
    ):
        unificar_colunas(df, "F", ["A"], "T")
