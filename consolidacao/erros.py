class ErroValidacao(Exception):
    """Falha de regra do tratamento; a mensagem reproduz o stop() do R."""


class ProcessamentoCancelado(Exception):
    """O usuario pediu para cancelar antes do fim do processamento."""


def verificar_cancelamento(cancelar) -> None:
    """Consulta o pedido de cancelamento da interface (None = nunca cancela)."""
    if cancelar is not None and cancelar():
        raise ProcessamentoCancelado("Processamento cancelado pelo usuario.")
