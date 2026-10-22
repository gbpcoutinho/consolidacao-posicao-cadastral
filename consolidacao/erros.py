class ErroValidacao(Exception):
    """Falha de regra do tratamento; a mensagem reproduz o stop() do R."""


class ProcessamentoCancelado(Exception):
    """O usuario pediu para cancelar antes do fim do processamento."""
