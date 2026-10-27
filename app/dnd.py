"""Arrastar e soltar do Explorer (tkinterdnd2/tkdnd). Opcional: se a biblioteca
nao carregar, o app abre normalmente e so some o arrastar e soltar."""
import logging
from pathlib import Path

log = logging.getLogger(__name__)


def habilitar(raiz) -> bool:
    """Carrega o tkdnd na janela raiz. False (e log) se nao der."""
    try:
        from tkinterdnd2 import TkinterDnD

        versao = TkinterDnD._require(raiz)
    except Exception:
        log.warning("tkdnd indisponivel: o app abre sem arrastar e soltar", exc_info=True)
        return False
    log.info("tkdnd %s carregado: arrastar e soltar ativo", versao)
    return True


def caminhos_soltos(widget, dados: str) -> list:
    """O tkdnd entrega os caminhos numa string Tcl: os que tem espaco, acento
    ou chaves vem entre {}. tk.splitlist faz o parse certo."""
    return [Path(p) for p in widget.tk.splitlist(dados)]


def registrar_alvo(widget, funcao) -> bool:
    """Faz widget aceitar arquivos/pastas soltos; funcao(lista_de_Path)."""
    try:
        from tkinterdnd2 import DND_FILES

        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>", lambda e: (funcao(caminhos_soltos(widget, e.data)), e.action)[1])
    except Exception:
        log.warning("Nao foi possivel ativar o arrastar e soltar em %s", widget, exc_info=True)
        return False
    return True
