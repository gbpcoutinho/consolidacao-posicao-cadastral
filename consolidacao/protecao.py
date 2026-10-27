"""Planilhas protegidas por senha (criptografia do Excel, container CFB).

A entrada e decifrada so em memoria; a saida e cifrada antes de tocar o
disco. As senhas nunca entram em mensagens, no Log nem em registros de erro.
"""
import io
from pathlib import Path

import msoffcrypto
import msoffcrypto.format.ooxml

from .erros import ErroValidacao

# Todo arquivo cifrado do Excel e um container CFB (OLE), nao um zip.
ASSINATURA_CFB = bytes.fromhex("d0cf11e0a1b11ae1")

MENSAGEM_SEM_SENHA = "Arquivo protegido por senha; cadastre a senha na aba Seguranca."
MENSAGEM_SENHA_ERRADA = (
    "Arquivo protegido por senha: nenhuma das senhas cadastradas abriu o arquivo."
)
OCULTO = "***"


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


def decifrar_em_memoria(caminho: Path, senhas=()):
    """BytesIO com o .xlsx decifrado, ou None se o arquivo nao e cifrado.

    Testa as senhas na ordem. Sem senha cadastrada ou com nenhuma servindo,
    ErroValidacao (mensagem sem a senha)."""
    caminho = Path(caminho)
    if not _cifrado(caminho):
        return None
    candidatas = senhas_validas(senhas)
    with open(caminho, "rb") as f:
        try:
            oficio = msoffcrypto.OfficeFile(f)
            if not oficio.is_encrypted():
                return None  # CFB sem cifra (ex.: .xls): o erro de sempre
        except Exception:
            return None
        if not candidatas:
            raise ErroValidacao(MENSAGEM_SEM_SENHA)
        for senha in candidatas:
            try:
                oficio.load_key(password=senha)
                saida = io.BytesIO()
                oficio.decrypt(saida)
            except Exception:
                continue  # senha errada (ou arquivo que ela nao abre): proxima
            saida.seek(0)
            return saida
    raise ErroValidacao(MENSAGEM_SENHA_ERRADA)


def cifrar_para(plano: io.BytesIO, senha: str, destino) -> None:
    """Grava em `destino` (caminho) o .xlsx `plano` cifrado com `senha`."""
    plano.seek(0)
    with open(destino, "wb") as saida:
        msoffcrypto.format.ooxml.OOXMLFile(plano).encrypt(senha, saida)
