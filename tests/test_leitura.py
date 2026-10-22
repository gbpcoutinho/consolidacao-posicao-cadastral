import pytest

from consolidacao.erros import ErroValidacao
from consolidacao.leitura import ler_dependente, ler_suplementar, ler_titular
from tests.fabrica import (
    ABA_SUPLEMENTAR,
    DEPENDENTE_OK,
    SUPLEMENTAR_OK,
    TITULAR_OK,
    criar_arquivo_pc,
)


def test_titular_basico(tmp_path):
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    r = ler_titular(p, "TITULAR", "PC1")
    d = r.dados
    assert "TIPO DE REGISTRO" not in d.columns
    assert {"REGIAO", "SEXO"} <= set(d.columns)
    assert d["CODIGO DO DEPENDENTE"].tolist() == ["00", "00"]
    assert d["TITULARIDADE"].tolist() == ["TITULAR", "TITULAR"]
    assert d[".CHAVE_CERTIFICADO"].tolist() == ["001", "002"]
    assert r.avisos == ["DATA DE REATIVACAO", "CARGO DE OCUPACAO", "NUMERO DA MATRICULA"]


def test_identificador_numerico_vira_texto_sem_ponto_zero(tmp_path):
    tit = [dict(TITULAR_OK[0], **{"NUMERO DO CPF": 12345678901, "NUMERO DA SUBFATURA": 7.0})]
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, tit, DEPENDENTE_OK, SUPLEMENTAR_OK)
    d = ler_titular(p, "TITULAR", "PC1").dados
    assert d["NUMERO DO CPF"].tolist() == ["12345678901"]
    assert d["NUMERO DA SUBFATURA"].tolist() == ["7"]


def test_titular_certificado_vazio(tmp_path):
    tit = [dict(TITULAR_OK[0]), dict(TITULAR_OK[1], **{"NUMERO DO CERTIFICADO": None})]
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, tit, DEPENDENTE_OK, SUPLEMENTAR_OK)
    with pytest.raises(ErroValidacao, match=r"\[TITULAR: PC1\] Ha titular sem NUMERO DO CERTIFICADO"):
        ler_titular(p, "TITULAR", "PC1")


def test_dependente_remove_duas_linhas_de_resumo(tmp_path):
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    d = ler_dependente(p, "DEPENDENTE", "PC1").dados
    assert len(d) == 2
    assert {
        "NUMERO DO CARTAO DO DEPENDENTE", "NOME DA MAE DO DEPENDENTE", ".SUBF_DEPENDENTE"
    } <= set(d.columns)
    assert d["TITULARIDADE"].tolist() == ["DEPENDENTE", "DEPENDENTE"]


def test_dependente_ignora_linhas_vazias_formatadas_no_fim(tmp_path):
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK, linhas_vazias_formatadas_no_fim=5)
    assert len(ler_dependente(p, "DEPENDENTE", "PC1").dados) == 2


def test_dependente_sem_dados(tmp_path):
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, TITULAR_OK, [], SUPLEMENTAR_OK)
    with pytest.raises(ErroValidacao, match=r"\[DEPENDENTE: PC1\] Ha 2 linha\(s\); nao existem dados"):
        ler_dependente(p, "DEPENDENTE", "PC1")


def test_dependente_com_campo_do_titular(tmp_path):
    dep = [dict(r, PLANO="X") for r in DEPENDENTE_OK]
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, TITULAR_OK, dep, SUPLEMENTAR_OK)
    with pytest.raises(ErroValidacao, match="tambem aparecem na aba dependente: PLANO"):
        ler_dependente(p, "DEPENDENTE", "PC1")


def test_suplementar_agrega_repeticoes_e_falha_em_conflito(tmp_path):
    vazio = {"NUMERO DO CERTIFICADO": None, "CARTÃO": None, "NOME DA MÃE": None}
    sup = SUPLEMENTAR_OK + [dict(SUPLEMENTAR_OK[0]), vazio]
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, sup)
    d = ler_suplementar(p, ABA_SUPLEMENTAR, "PC1").dados
    assert list(d.columns) == [".CHAVE_CERTIFICADO", ".CARTAO_SUP", ".NOME_MAE_SUP"]
    assert len(d) == 1
    conflito = SUPLEMENTAR_OK + [dict(SUPLEMENTAR_OK[0], **{"CARTÃO": "OUTRO"})]
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, conflito)
    with pytest.raises(ErroValidacao, match="O certificado '001' tem valores conflitantes em 'CARTAO'"):
        ler_suplementar(p, ABA_SUPLEMENTAR, "PC1")
