"""Ponto de entrada do executavel ConsolidadorPC.

Sem argumentos abre a janela. "--fumo ENTRADA CSV SAIDA" (escondido) processa
uma pasta sem interface e sai com 0 (ok) ou 1 (erro): e o teste do .exe.
Nada aqui usa __file__ para achar dados (dentro do .exe ele aponta para uma
pasta temporaria): preferencias, cache e logs ficam em %APPDATA%\\ConsolidadorPC.
"""
import logging
import logging.handlers
import sys
import traceback
from datetime import date
from pathlib import Path

from app.config import pasta_dados
from app.tarefas import MENSAGEM_INESPERADO
from consolidacao.empresas import carregar_mapa_empresas
from consolidacao.erros import ErroValidacao
from consolidacao.pipeline import executar, resumir
from consolidacao.processamento import listar_arquivos_entrada

log = logging.getLogger("consolidador")


def configurar_log(pasta: Path) -> None:
    pasta_logs = Path(pasta) / "logs"
    pasta_logs.mkdir(parents=True, exist_ok=True)
    handler = logging.handlers.RotatingFileHandler(
        pasta_logs / "app.log", maxBytes=1_000_000, backupCount=3, encoding="utf-8"
    )
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    raiz = logging.getLogger()
    raiz.addHandler(handler)
    if raiz.level > logging.INFO or raiz.level == logging.NOTSET:
        raiz.setLevel(logging.INFO)


def tratar_excecao(tipo, excecao, tb, mostrar=None) -> None:
    """Excecao que escapou de tudo: detalhes so no log; na tela, texto generico."""
    log.error("Excecao nao tratada:\n%s", "".join(traceback.format_exception(tipo, excecao, tb)))
    if mostrar is not None:
        try:
            mostrar(MENSAGEM_INESPERADO)
        except Exception:
            log.exception("Nao foi possivel mostrar a mensagem de erro")


def _escrever(texto: str) -> None:
    # No .exe sem console, sys.stdout e None.
    if sys.stdout is not None:
        print(texto)


def fumo(argumentos: list) -> int:
    entrada, empresas, saida = (Path(a) for a in argumentos[:3])
    # Opcionais: --senha X (repetivel) e --senha-saida X, so para o fumo do build.
    senhas, senha_saida, resto = [], None, iter(argumentos[3:])
    for opcao in resto:
        valor = next(resto, "")
        if opcao == "--senha":
            senhas.append(valor)
        elif opcao == "--senha-saida":
            senha_saida = valor
    try:
        mapa = carregar_mapa_empresas(empresas)
        arquivos = listar_arquivos_entrada(entrada, date.today())
        r = resumir(executar(
            arquivos, mapa, saida, arquivo_empresas=empresas, senhas=senhas, senha_saida=senha_saida))
    except ErroValidacao as e:
        log.error("Fumo falhou: %s", e)
        _escrever(f"ERRO: {e}")
        return 1
    _escrever(f"OK {r['arquivos_ok']} arquivo(s), {r['registros']} registro(s) -> {saida}")
    return 0 if r["arquivos_erro"] == 0 else 1


def main(argv: list | None = None, pasta: Path | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    pasta = Path(pasta) if pasta is not None else pasta_dados()
    pasta.mkdir(parents=True, exist_ok=True)
    configurar_log(pasta)
    if argv[:1] == ["--fumo"]:
        return fumo(argv[1:])

    import tkinter as tk
    from tkinter import messagebox

    from app.janela import App

    raiz = tk.Tk()

    def mostrar(texto):
        messagebox.showerror("Erro", texto, parent=raiz)

    sys.excepthook = lambda t, e, tb: tratar_excecao(t, e, tb, mostrar)
    raiz.report_callback_exception = lambda t, e, tb: tratar_excecao(t, e, tb, mostrar)
    App(raiz, pasta)
    raiz.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
