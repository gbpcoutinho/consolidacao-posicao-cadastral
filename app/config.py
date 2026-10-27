"""Preferencias do usuario em %APPDATA%\\ConsolidadorPC\\config.json.

Chaves: arquivo_empresas, ultima_pasta_entrada, ultima_pasta_saida.
Nada aqui usa __file__: dentro do .exe ele aponta para uma pasta temporaria.
"""
import json
import logging
import os
import tempfile
from pathlib import Path

NOME_PASTA = "ConsolidadorPC"
ARQUIVO_CONFIG = "config.json"

log = logging.getLogger(__name__)


def pasta_dados(base: Path | None = None) -> Path:
    """%APPDATA%\\ConsolidadorPC (ou base\\ConsolidadorPC nos testes), criada se faltar."""
    if base is None:
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    pasta = Path(base) / NOME_PASTA
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


def carregar_config(pasta: Path) -> dict:
    caminho = Path(pasta) / ARQUIVO_CONFIG
    try:
        dados = json.loads(caminho.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except (OSError, ValueError) as e:
        log.warning("Preferencias ilegiveis em %s; usando padrao (%s)", caminho, e)
        return {}
    if not isinstance(dados, dict):
        log.warning("Preferencias em %s nao sao um objeto JSON; usando padrao", caminho)
        return {}
    return dados


def salvar_config(pasta: Path, cfg: dict) -> None:
    # Temporario + replace: uma queda no meio nao deixa o JSON pela metade.
    pasta = Path(pasta)
    descritor, nome = tempfile.mkstemp(dir=pasta, prefix="~$config", suffix=".tmp")
    try:
        with os.fdopen(descritor, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        os.replace(nome, pasta / ARQUIVO_CONFIG)
    except BaseException:
        Path(nome).unlink(missing_ok=True)
        raise
