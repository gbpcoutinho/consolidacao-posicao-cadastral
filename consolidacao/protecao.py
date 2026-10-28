"""Planilhas protegidas por senha (criptografia do Excel, container CFB).

A entrada e decifrada so em memoria; a saida e cifrada antes de tocar o
disco. As senhas nunca entram em mensagens, no Log nem em registros de erro.
"""
import io
import re
from pathlib import Path

import msoffcrypto
import msoffcrypto.format.ooxml

from .erros import ErroValidacao

# Todo arquivo cifrado do Excel e um container CFB (OLE), nao um zip.
ASSINATURA_CFB = bytes.fromhex("d0cf11e0a1b11ae1")

MENSAGEM_SEM_SENHA = "Arquivo protegido por senha; cadastre a senha na aba Seguranca."
MENSAGEM_SENHA_ERRADA = "Arquivo protegido por senha: nenhuma senha cadastrada abriu o arquivo."
MENSAGEM_SENHA_ERRADA_AUTO = (
    "Arquivo protegido por senha: nenhuma senha abriu o arquivo (automatica e cadastradas)."
)
DICA_CADASTRAR = " Cadastre a senha na aba Seguranca."
OCULTO = "***"

# Senha automatica: o usuario informa um padrao com {apolice}, que vira o codigo
# apos "PC" no nome do arquivo (com os zeros do nome). O padrao NAO tem valor de
# fabrica (vazio = desligada): ele e informado pelo usuario e guardado cifrado.
MARCADOR_APOLICE = "{apolice}"
PADRAO_SENHA_AUTOMATICA = ""
_APOLICE = re.compile(r"PC([0-9]+)", flags=re.IGNORECASE)


def padrao_valido(padrao) -> bool:
    return isinstance(padrao, str) and MARCADOR_APOLICE in padrao


def senha_automatica(caminho, padrao) -> str | None:
    """Senha do arquivo pelo padrao (None = desativada, padrao sem {apolice} ou
    nome sem o codigo apos PC). Mesmo codigo de empresa_da_referencia."""
    if not padrao_valido(padrao):
        return None
    m = _APOLICE.match(Path(caminho).stem)
    return padrao.replace(MARCADOR_APOLICE, m.group(1)) if m else None


def candidatas(caminho, senhas=(), padrao=None) -> list:
    """Ordem de tentativa: senha automatica do arquivo, depois as cadastradas;
    sem repeticao."""
    return senhas_validas([senha_automatica(caminho, padrao), *(senhas or ())])


def senhas_validas(senhas) -> list:
    """Senhas nao vazias, na ordem, sem repeticao."""
    return list(dict.fromkeys(s for s in (senhas or ()) if s))


def ocultar_senhas(texto: str, senhas) -> str:
    for s in sorted(senhas_validas(senhas), key=len, reverse=True):
        texto = texto.replace(s, OCULTO)
    return texto


def _cifrado(caminho: Path) -> bool:
    with open(caminho, "rb") as f:
        return f.read(len(ASSINATURA_CFB)) == ASSINATURA_CFB


def decifrar_em_memoria(caminho: Path, senhas=(), padrao=None):
    """BytesIO com o .xlsx decifrado, ou None se o arquivo nao e cifrado.

    Testa na ordem: a senha automatica (padrao com a apolice do nome) e as
    cadastradas. Sem nenhuma candidata ou com nenhuma servindo, ErroValidacao
    (mensagem sem senha alguma)."""
    caminho = Path(caminho)
    if not _cifrado(caminho):
        return None
    automatica = senha_automatica(caminho, padrao)
    tentativas = candidatas(caminho, senhas, padrao)
    with open(caminho, "rb") as f:
        try:
            oficio = msoffcrypto.OfficeFile(f)
            if not oficio.is_encrypted():
                return None  # CFB sem cifra (ex.: .xls): o erro de sempre
        except Exception:
            return None
        if not tentativas:
            raise ErroValidacao(MENSAGEM_SEM_SENHA)
        for senha in tentativas:
            try:
                oficio.load_key(password=senha)
                saida = io.BytesIO()
                oficio.decrypt(saida)
            except Exception:
                continue  # senha errada (ou arquivo que ela nao abre): proxima
            saida.seek(0)
            return saida
    if automatica is None:
        raise ErroValidacao(MENSAGEM_SENHA_ERRADA)
    raise ErroValidacao(
        MENSAGEM_SENHA_ERRADA_AUTO + ("" if senhas_validas(senhas) else DICA_CADASTRAR))


def cifrar_para(plano: io.BytesIO, senha: str, destino) -> None:
    """Grava em `destino` (caminho) o .xlsx `plano` cifrado com `senha`."""
    plano.seek(0)
    with open(destino, "wb") as saida:
        msoffcrypto.format.ooxml.OOXMLFile(plano).encrypt(senha, saida)
