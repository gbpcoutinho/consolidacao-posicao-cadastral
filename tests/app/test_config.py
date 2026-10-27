import json
import logging

from app import VERSAO
from app.config import carregar_config, pasta_dados, salvar_config


def test_versao():
    assert VERSAO == "1.0.0"


def test_pasta_dados_usa_appdata_e_cria(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    pasta = pasta_dados()
    assert pasta == tmp_path / "ConsolidadorPC" and pasta.is_dir()


def test_pasta_dados_injetavel(tmp_path):
    pasta = pasta_dados(tmp_path / "base")
    assert pasta == tmp_path / "base" / "ConsolidadorPC" and pasta.is_dir()


def test_config_ida_e_volta(tmp_path):
    cfg = {"arquivo_empresas": r"\\servidor\pasta\empresas.csv", "ultima_pasta_entrada": "C:/x"}
    salvar_config(tmp_path, cfg)
    assert carregar_config(tmp_path) == cfg
    assert json.loads((tmp_path / "config.json").read_text(encoding="utf-8")) == cfg


def test_config_ausente_vira_vazio(tmp_path):
    assert carregar_config(tmp_path) == {}


def test_config_json_corrompido_vira_vazio(tmp_path, caplog):
    (tmp_path / "config.json").write_text("{nao e json", encoding="utf-8")
    with caplog.at_level(logging.WARNING):
        assert carregar_config(tmp_path) == {}
    assert "config.json" in caplog.text


def test_config_json_que_nao_e_objeto_vira_vazio(tmp_path):
    (tmp_path / "config.json").write_text("[1, 2]", encoding="utf-8")
    assert carregar_config(tmp_path) == {}


def test_salvar_config_nao_deixa_temporario(tmp_path):
    salvar_config(tmp_path, {"a": 1})
    assert [p.name for p in tmp_path.iterdir()] == ["config.json"]
