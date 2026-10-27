"""Preferencias do usuario em %APPDATA%\\ConsolidadorPC\\config.json.

Chaves: arquivo_empresas, ultima_pasta_entrada, ultima_pasta_saida.
Nada aqui usa __file__: dentro do .exe ele aponta para uma pasta temporaria.
"""
import base64
import ctypes
import json
import logging
import os
import tempfile
from ctypes import wintypes
from pathlib import Path

from consolidacao.protecao import PADRAO_SENHA_AUTOMATICA, padrao_valido

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


# --- senhas ------------------------------------------------------------------
# Guardadas no config.json cifradas com a DPAPI do Windows (escopo do usuario):
# so a mesma conta, na mesma maquina, consegue le-las. Nada aqui registra o
# texto de uma senha, nem em caso de falha.

CHAVE_SENHAS_ENTRADA = "senhas_entrada"
CHAVE_SENHA_SAIDA = "senha_saida"
CHAVE_AUTOMATICA_ATIVA = "senha_automatica_ativa"
CHAVE_PADRAO_AUTOMATICA = "padrao_senha_automatica"  # cifrado: o padrao e sensivel


class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _blob(dados: bytes):
    buffer = ctypes.create_string_buffer(dados, len(dados))
    return _Blob(len(dados), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char))), buffer


def _dpapi(funcao: str, dados: bytes) -> bytes:
    entrada, _manter = _blob(dados)
    saida = _Blob()
    chamada = getattr(ctypes.windll.crypt32, funcao)
    chamada.argtypes = [ctypes.POINTER(_Blob), ctypes.c_void_p, ctypes.c_void_p,
                        ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_Blob)]
    chamada.restype = wintypes.BOOL
    if not chamada(ctypes.byref(entrada), None, None, None, None, 0, ctypes.byref(saida)):
        raise OSError(f"{funcao} falhou (codigo {ctypes.GetLastError()})")
    try:
        return ctypes.string_at(saida.pbData, saida.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(saida.pbData)


def proteger_dpapi(texto: str) -> str:
    """Texto cifrado pela DPAPI (CryptProtectData), em base64."""
    return base64.b64encode(_dpapi("CryptProtectData", texto.encode("utf-8"))).decode("ascii")


def revelar_dpapi(blob: str) -> str:
    """Inverso de proteger_dpapi; levanta se o blob nao for desta conta."""
    return _dpapi("CryptUnprotectData", base64.b64decode(blob, validate=True)).decode("utf-8")


def guardar_senhas(cfg: dict, entrada, saida: str) -> None:
    """Poe as senhas (cifradas) em `cfg`; vazias removem as chaves."""
    entrada = [s for s in entrada if s]
    if entrada:
        cfg[CHAVE_SENHAS_ENTRADA] = [proteger_dpapi(s) for s in entrada]
    else:
        cfg.pop(CHAVE_SENHAS_ENTRADA, None)
    if saida:
        cfg[CHAVE_SENHA_SAIDA] = proteger_dpapi(saida)
    else:
        cfg.pop(CHAVE_SENHA_SAIDA, None)


def _revelar_ou_vazio(blob, rotulo: str) -> str:
    if not blob or not isinstance(blob, str):
        return ""
    try:
        return revelar_dpapi(blob)
    except Exception as e:  # blob de outra conta/maquina, corrompido ou editado a mao
        log.warning("Nao foi possivel decifrar a senha salva (%s): %s", rotulo, type(e).__name__)
        return ""


def ler_senhas(cfg: dict) -> tuple:
    """(senhas de entrada, senha de saida) em texto; o que nao decifra vira vazio."""
    blobs = cfg.get(CHAVE_SENHAS_ENTRADA)
    if not isinstance(blobs, list):
        blobs = []
    entrada = [_revelar_ou_vazio(b, "entrada") for b in blobs]
    return [s for s in entrada if s], _revelar_ou_vazio(cfg.get(CHAVE_SENHA_SAIDA), "saida")


def guardar_senha_automatica(cfg: dict, ativa: bool, padrao: str, de_fabrica: str = "") -> None:
    """Liga/desliga a senha automatica e guarda o padrao (cifrado). O padrao de
    fabrica (arquivo local) nao e gravado na config: segue valendo o do arquivo."""
    cfg[CHAVE_AUTOMATICA_ATIVA] = bool(ativa)
    if padrao and padrao != de_fabrica and padrao != PADRAO_SENHA_AUTOMATICA:
        cfg[CHAVE_PADRAO_AUTOMATICA] = proteger_dpapi(padrao)
    else:
        cfg.pop(CHAVE_PADRAO_AUTOMATICA, None)


def ler_padrao_de_fabrica(pasta_recursos) -> str:
    """Primeira linha de padrao_senha.txt na pasta de recursos (fora do git, como
    o logo). Sem o arquivo, ou com conteudo invalido, vazio."""
    try:
        linhas = (Path(pasta_recursos) / "padrao_senha.txt").read_text(
            encoding="utf-8-sig").splitlines()
    except (OSError, ValueError):
        return PADRAO_SENHA_AUTOMATICA
    padrao = linhas[0].strip() if linhas else ""
    return padrao if padrao_valido(padrao) else PADRAO_SENHA_AUTOMATICA


def ler_senha_automatica(cfg: dict, de_fabrica: str = "") -> tuple:
    """(ativa, padrao). O padrao salvo na config vale; sem ele, o de fabrica (do
    arquivo de recursos); sem nenhum, vazio (= nao age)."""
    ativa = cfg.get(CHAVE_AUTOMATICA_ATIVA, True)
    ativa = ativa if isinstance(ativa, bool) else True
    padrao = _revelar_ou_vazio(cfg.get(CHAVE_PADRAO_AUTOMATICA), "padrao automatico")
    return ativa, padrao if padrao_valido(padrao) else (de_fabrica or PADRAO_SENHA_AUTOMATICA)
