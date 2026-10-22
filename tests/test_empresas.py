import pytest

from consolidacao.empresas import carregar_mapa_empresas, empresa_da_referencia
from consolidacao.erros import ErroValidacao
from tests.fabrica import escrever_csv_empresas


def test_mapa_basico(tmp_path):
    p = tmp_path / "e.csv"
    escrever_csv_empresas(p, [("ACME", "123"), (" BETA ", " 9 "), ("", "7"), ("NA", "8")])
    assert carregar_mapa_empresas(p) == {"123": "ACME", "9": "BETA"}


def test_mapa_cp1252(tmp_path):
    p = tmp_path / "e.csv"
    escrever_csv_empresas(p, [("JOÃO SAÚDE", "1")], encoding="cp1252")
    assert carregar_mapa_empresas(p) == {"1": "JOÃO SAÚDE"}


def test_mapa_codigo_com_duas_empresas(tmp_path):
    p = tmp_path / "e.csv"
    escrever_csv_empresas(p, [("A", "1"), ("B", "1")])
    with pytest.raises(
        ErroValidacao, match="COD_SAUDE '1' esta associado a mais de uma EMPRESA: A, B"
    ):
        carregar_mapa_empresas(p)


def test_mapa_sem_colunas(tmp_path):
    p = tmp_path / "e.csv"
    p.write_text("X;Y\n1;2\n", encoding="utf-8")
    with pytest.raises(
        ErroValidacao,
        match="O COD_EMPRESAS.csv precisa conter as colunas EMPRESA e COD_SAUDE.",
    ):
        carregar_mapa_empresas(p)


def test_mapa_inexistente_e_vazio(tmp_path):
    with pytest.raises(ErroValidacao, match="Arquivo de empresas nao encontrado: "):
        carregar_mapa_empresas(tmp_path / "x.csv")
    p = tmp_path / "v.csv"
    p.write_bytes(b"")
    with pytest.raises(ErroValidacao, match="O COD_EMPRESAS.csv esta vazio."):
        carregar_mapa_empresas(p)


def test_empresa_da_referencia():
    assert empresa_da_referencia("pc123 10.2026", {"123": "ACME"}) == ("123", "ACME")
    with pytest.raises(
        ErroValidacao,
        match="Nao foi possivel extrair o codigo apos 'PC' da referencia 'X1'.",
    ):
        empresa_da_referencia("X1", {})
    with pytest.raises(
        ErroValidacao,
        match="Codigo '5' extraido de 'PC5' nao encontrado em COD_SAUDE no COD_EMPRESAS.csv.",
    ):
        empresa_da_referencia("PC5", {"1": "A"})
