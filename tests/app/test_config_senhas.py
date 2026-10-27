"""Senhas guardadas no config.json cifradas com a DPAPI do Windows."""
import json
import logging

import pytest

from app.config import (
    carregar_config, guardar_senhas, ler_senhas, proteger_dpapi, revelar_dpapi, salvar_config,
)

pytestmark = pytest.mark.skipif(
    not __import__("sys").platform.startswith("win"), reason="DPAPI e do Windows"
)

ENTRADA = ["Primeira#1", "segunda ç ã", "3"]
SAIDA = "Saida!Forte-9"


def test_dpapi_ida_e_volta():
    blob = proteger_dpapi("Senha com acento: ação")
    assert isinstance(blob, str) and "ação" not in blob
    assert revelar_dpapi(blob) == "Senha com acento: ação"
    assert revelar_dpapi(proteger_dpapi("")) == ""


def test_dpapi_cifra_de_verdade_e_varia():
    assert "abc" not in proteger_dpapi("abc")


def test_senhas_ida_e_volta_pelo_arquivo(tmp_path):
    cfg = {"ultima_pasta_entrada": "C:/x"}
    guardar_senhas(cfg, ENTRADA, SAIDA)
    salvar_config(tmp_path, cfg)
    bruto = (tmp_path / "config.json").read_text(encoding="utf-8")
    assert SAIDA not in bruto and all(s not in bruto for s in ENTRADA if len(s) > 2)
    lido = carregar_config(tmp_path)
    assert lido["ultima_pasta_entrada"] == "C:/x"
    assert ler_senhas(lido) == (ENTRADA, SAIDA)


def test_sem_senhas_guardadas_vira_vazio():
    assert ler_senhas({}) == ([], "")


def test_guardar_vazio_remove_as_chaves():
    cfg = {}
    guardar_senhas(cfg, ENTRADA, SAIDA)
    guardar_senhas(cfg, [], "")
    assert ler_senhas(cfg) == ([], "")
    assert "senha" not in json.dumps(cfg).replace("senha_saida", "").replace("senhas_entrada", "")


def test_falha_ao_decifrar_vira_vazio_com_aviso_sem_a_senha(caplog):
    cfg = {}
    guardar_senhas(cfg, ["Boa#1", "Outra#2"], SAIDA)
    cfg["senhas_entrada"][0] = "AAAA"  # lixo: nao e um blob da DPAPI
    cfg["senha_saida"] = "nao-e-base64!!"
    with caplog.at_level(logging.WARNING):
        entrada, saida = ler_senhas(cfg)
    assert entrada == ["Outra#2"] and saida == ""
    assert any(r.levelno == logging.WARNING for r in caplog.records)
    texto = "\n".join(r.getMessage() for r in caplog.records)
    assert "Boa#1" not in texto and "Outra#2" not in texto and SAIDA not in texto


def test_config_com_tipo_errado_nao_quebra():
    assert ler_senhas({"senhas_entrada": "texto", "senha_saida": 5}) == ([], "")
