class ErroValidacao(Exception):
    """Falha de regra do tratamento; a mensagem reproduz o stop() do R."""


class ErroInesperado(ErroValidacao):
    """Falha que nao e regra do tratamento (defeito, biblioteca); a original
    fica em __cause__. O CLI mostra a mensagem; o app registra o traceback no
    log e mostra um texto generico."""


class ProcessamentoCancelado(Exception):
    """O usuario pediu para cancelar antes do fim do processamento."""


def verificar_cancelamento(cancelar) -> None:
    """Consulta o pedido de cancelamento da interface (None = nunca cancela)."""
    if cancelar is not None and cancelar():
        raise ProcessamentoCancelado("Processamento cancelado pelo usuario.")
