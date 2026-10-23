# Fluxo completo reutilizavel pela linha de comando e pelo app.
# Os caminhos chegam sempre explicitos: nada aqui depende de __file__, que
# aponta para uma pasta temporaria dentro de um .exe do PyInstaller.
from pathlib import Path
from typing import Callable

from .exportacao import escrever_excel
from .processamento import Consolidado, processar_arquivos


def executar(
    arquivos: list,
    mapa: dict,
    saida: Path,
    progresso: Callable[[int, int, str], None] | None = None,
    cancelar: Callable[[], bool] | None = None,
) -> Consolidado:
    """Processa os arquivos e grava o Excel consolidado em `saida`."""
    resultado = processar_arquivos(arquivos, mapa, progresso=progresso, cancelar=cancelar)
    escrever_excel(resultado, Path(saida))
    return resultado


def resumir(c: Consolidado) -> dict:
    log = c.log
    erros = log[log["STATUS"] == "ERRO"]
    return {
        "arquivos_ok": int((log["STATUS"] == "OK").sum()),
        "arquivos_erro": len(erros),
        "registros": len(c.consolidado),
        "pendencias": len(c.pendencias),
        "datas_invalidas": int(log["DATAS_INVALIDAS"].sum()),
        "divergencias_datas": int(log["QTD_DIVERGENCIAS_DATAS"].sum()),
        "erros": list(zip(erros["ARQUIVO"], erros["MENSAGEM"])),
    }
