# Fluxo completo reutilizavel pela linha de comando e pelo app.
# Os caminhos chegam sempre explicitos: nada aqui depende de __file__, que
# aponta para uma pasta temporaria dentro de um .exe do PyInstaller.
import os
from pathlib import Path
from typing import Callable

from .erros import ErroInesperado, ErroValidacao, ProcessamentoCancelado
from .exportacao import escrever_excel
from .processamento import Consolidado, avisar_progresso, processar_arquivos
from .protecao import ocultar_senhas

# Nome passado ao progresso depois do ultimo arquivo, durante a gravacao.
ETAPA_GRAVACAO = "Gravando arquivo de saida"


def _mesmo_arquivo(a: Path, b: Path) -> bool:
    try:
        if a.exists() and b.exists():
            return a.samefile(b)
    except OSError:
        pass
    # normcase: no Windows, maiusculas e minusculas dao no mesmo arquivo.
    return os.path.normcase(a.resolve()) == os.path.normcase(b.resolve())


def _validar_saida(saida: Path, arquivos: list, arquivo_empresas: Path | None) -> None:
    # A gravacao troca o destino pelo temporario: apontar a saida para uma
    # entrada apagaria a planilha original.
    for arquivo in arquivos:
        if _mesmo_arquivo(saida, Path(arquivo)):
            raise ErroValidacao(
                f"O arquivo de saida '{saida}' e um dos arquivos de entrada; "
                "escolha outro nome para a saida."
            )
    if arquivo_empresas is not None and _mesmo_arquivo(saida, Path(arquivo_empresas)):
        raise ErroValidacao(
            f"O arquivo de saida '{saida}' e o arquivo de empresas; "
            "escolha outro nome para a saida."
        )


def _protegido(funcao, *args, _segredos=(), **kwargs):
    """Erro inesperado vira ErroValidacao (a original fica em __cause__);
    ProcessamentoCancelado e MemoryError passam direto. KeyboardInterrupt nao
    e Exception e tambem passa."""
    try:
        return funcao(*args, **kwargs)
    except (ErroValidacao, ProcessamentoCancelado, MemoryError):
        raise
    except Exception as e:
        raise ErroInesperado(
            "Erro inesperado ao gerar a consolidacao "
            f"({type(e).__name__}: {ocultar_senhas(str(e), _segredos)})."
        ) from e


def processar(
    arquivos: list,
    mapa: dict,
    progresso: Callable[[int, int, str], None] | None = None,
    cancelar: Callable[[], bool] | None = None,
    senhas=(),
) -> Consolidado:
    """Processa os arquivos sem gravar nada.

    senhas: tentadas em ordem nos arquivos protegidos por senha.
    progresso(i, total, nome) e chamado antes de cada arquivo; cancelar() True
    interrompe com ProcessamentoCancelado."""
    return _protegido(
        processar_arquivos, arquivos, mapa, progresso=progresso, cancelar=cancelar,
        senhas=senhas, _segredos=senhas,
    )


def _gravar(resultado, saida, arquivos, arquivo_empresas, progresso, senha_saida) -> None:
    _validar_saida(saida, arquivos, arquivo_empresas)
    total = len(arquivos)
    avisar_progresso(progresso, total, total, ETAPA_GRAVACAO)
    escrever_excel(resultado, saida, senha=senha_saida)


def gravar(
    resultado: Consolidado,
    saida: Path,
    arquivos: list,
    arquivo_empresas: Path | None = None,
    progresso: Callable[[int, int, str], None] | None = None,
    senha_saida: str | None = None,
) -> None:
    """Grava o Excel consolidado em `saida`, com as mesmas validacoes de saida
    de executar. progresso recebe (total, total, ETAPA_GRAVACAO). Com
    senha_saida o arquivo nasce cifrado (vazia ou None = sem senha)."""
    _protegido(
        _gravar, resultado, Path(saida), arquivos, arquivo_empresas, progresso,
        senha_saida, _segredos=[senha_saida],
    )


def executar(
    arquivos: list,
    mapa: dict,
    saida: Path,
    progresso: Callable[[int, int, str], None] | None = None,
    cancelar: Callable[[], bool] | None = None,
    arquivo_empresas: Path | None = None,
    senhas=(),
    senha_saida: str | None = None,
) -> Consolidado:
    """Processa os arquivos e grava o Excel consolidado em `saida`.

    progresso(i, total, nome) e chamado antes de cada arquivo e, por fim, com
    (total, total, ETAPA_GRAVACAO). Excecoes possiveis: ErroValidacao (inclusive
    para falhas inesperadas, com a original em __cause__), ProcessamentoCancelado
    e MemoryError."""
    saida = Path(saida)
    # Valida antes de processar: nao vale gastar minutos para recusar a saida.
    _protegido(_validar_saida, saida, arquivos, arquivo_empresas)
    resultado = processar(arquivos, mapa, progresso=progresso, cancelar=cancelar, senhas=senhas)
    gravar(
        resultado, saida, arquivos, arquivo_empresas=arquivo_empresas, progresso=progresso,
        senha_saida=senha_saida,
    )
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
