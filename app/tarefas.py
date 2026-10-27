"""Processamento em thread de fundo, sem Tk.

A thread so poe Eventos numa fila; quem toca na janela e a thread principal,
que consome a fila com eventos()/despachar() chamados por after(). Assim a
janela nunca congela e nenhum widget e tocado fora da thread do Tk.
"""
import logging
import queue
import threading
from dataclasses import dataclass
from typing import Any, Callable, Literal

from consolidacao.erros import ErroInesperado, ErroValidacao, ProcessamentoCancelado

MENSAGEM_INESPERADO = "Erro inesperado; detalhes em logs/app.log"
MENSAGEM_MEMORIA = (
    "Memoria insuficiente para processar estes arquivos; feche outros programas "
    "ou processe em lotes menores."
)
TERMINAIS = ("concluido", "erro", "cancelado")

log = logging.getLogger(__name__)


@dataclass(frozen=True)
class Evento:
    tipo: Literal["progresso", "concluido", "erro", "cancelado"]
    dados: Any = None


class Trabalho:
    """Roda funcao(progresso, cancelar) numa thread e entrega Eventos.

    progresso(i, total, nome) vira Evento("progresso", (i, total, nome));
    cancelar() devolve True depois de Trabalho.cancelar()."""

    def __init__(self, funcao: Callable[[Callable, Callable], Any]):
        self._funcao = funcao
        self._fila: queue.Queue = queue.Queue()
        self._pedido = threading.Event()
        self._thread: threading.Thread | None = None

    @property
    def ativo(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    @property
    def cancelamento_pedido(self) -> bool:
        return self._pedido.is_set()

    def iniciar(self) -> None:
        if self._thread is not None:
            raise RuntimeError("Trabalho ja iniciado")
        self._thread = threading.Thread(target=self._rodar, name="processamento", daemon=True)
        self._thread.start()

    def cancelar(self) -> None:
        self._pedido.set()

    def esperar(self, segundos: float | None = None) -> None:
        if self._thread is not None:
            self._thread.join(segundos)

    def eventos(self) -> list:
        """Eventos pendentes, sem bloquear."""
        pendentes = []
        while True:
            try:
                pendentes.append(self._fila.get_nowait())
            except queue.Empty:
                return pendentes

    def despachar(self, tratador: Callable[[Evento], None]) -> list:
        """Chama tratador(evento) na thread de quem chama (a do Tk)."""
        pendentes = self.eventos()
        for evento in pendentes:
            tratador(evento)
        return pendentes

    def _progresso(self, i: int, total: int, nome: str) -> None:
        self._fila.put(Evento("progresso", (i, total, nome)))

    def _rodar(self) -> None:
        try:
            resultado = self._funcao(self._progresso, self._pedido.is_set)
        except ProcessamentoCancelado:
            self._fila.put(Evento("cancelado"))
        except ErroInesperado as e:
            # Defeito embrulhado pelo motor: traceback original no log.
            causa = e.__cause__ or e
            log.error("Erro inesperado no processamento: %s", e,
                      exc_info=(type(causa), causa, causa.__traceback__))
            self._fila.put(Evento("erro", MENSAGEM_INESPERADO))
        except ErroValidacao as e:
            self._fila.put(Evento("erro", str(e)))
        except MemoryError:
            self._fila.put(Evento("erro", MENSAGEM_MEMORIA))
        except Exception:
            log.exception("Erro inesperado no processamento")
            self._fila.put(Evento("erro", MENSAGEM_INESPERADO))
        else:
            self._fila.put(Evento("concluido", resultado))
