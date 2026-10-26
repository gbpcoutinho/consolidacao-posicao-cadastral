"""Mede o tempo do motor em arquivos FICTICIOS grandes (fase 2, Tarefas 3 e 4).

Uso (na raiz do projeto):
    .venv\\Scripts\\python -m scripts.medir_desempenho PASTA [--titulares 10000]
        [--perfil] [--cancelar-apos SEGUNDOS]

Gera em PASTA (de preferencia fora do OneDrive) dois arquivos no formato real:
- "real": como exemplos/gerar_exemplos.py (tudo texto);
- "pesado": o mesmo conteudo com identificadores numericos e datas como
  datas do Excel, o que exercita o texto bruto de <v> e a conversao de datas.
Com --perfil, imprime o top 15 do cProfile (tempo proprio, tottime) de cada caso.
Com --cancelar-apos, pede o cancelamento depois de N segundos e mede quanto
tempo o motor leva para parar.
"""
import argparse
import cProfile
import io
import pstats
import random
import threading
import time
from datetime import datetime
from pathlib import Path

from consolidacao.erros import ProcessamentoCancelado
from consolidacao.pipeline import gravar, processar
from exemplos.gerar_exemplos import gerar_empresa, salvar

CODIGO = "90001"
MAPA = {CODIGO: "EMPRESA FICTICIA"}


def _pesado(abas):
    # Numeros e datas de verdade no lugar do texto do arquivo real.
    def converter(v):
        if isinstance(v, str) and v.isdigit() and len(v) < 16:
            return int(v)
        if isinstance(v, str) and len(v) == 10 and v[2] == "/" and v[5] == "/" and v != "00/00/0000":
            try:
                return datetime.strptime(v, "%d/%m/%Y")
            except ValueError:
                return v
        return v

    return {k: [[converter(v) for v in linha] for linha in linhas] for k, linhas in abas.items()}


def gerar(pasta: Path, titulares: int) -> dict:
    pasta.mkdir(parents=True, exist_ok=True)
    abas = gerar_empresa(CODIGO, titulares, random.Random(26), com_casos_especiais=True)
    casos = {"real": abas, "pesado": _pesado(abas)}
    caminhos = {}
    for nome, conteudo in casos.items():
        caminho = pasta / nome / f"PC{CODIGO} 10.2026.xlsx"
        caminho.parent.mkdir(exist_ok=True)
        if not caminho.exists():
            salvar(caminho, CODIGO, conteudo)
        caminhos[nome] = caminho
    return caminhos


def medir(caminho: Path, saida: Path, perfil: bool) -> None:
    perfilador = cProfile.Profile() if perfil else None
    inicio = time.perf_counter()
    if perfilador:
        perfilador.enable()
    c = processar([caminho], MAPA)
    meio = time.perf_counter()
    gravar(c, saida, [caminho])
    fim = time.perf_counter()
    if perfilador:
        perfilador.disable()
    print(f"  registros={len(c.consolidado)} processar={meio - inicio:.2f}s "
          f"gravar={fim - meio:.2f}s total={fim - inicio:.2f}s")
    if perfilador:
        texto = io.StringIO()
        pstats.Stats(perfilador, stream=texto).sort_stats("tottime").print_stats(15)
        print(texto.getvalue())


def medir_cancelamento(caminho: Path, apos: float) -> None:
    pedido = threading.Event()
    momento = {}

    def cancelar():
        return pedido.is_set()

    def disparar():
        time.sleep(apos)
        momento["pedido"] = time.perf_counter()
        pedido.set()

    threading.Thread(target=disparar, daemon=True).start()
    try:
        processar([caminho], MAPA, cancelar=cancelar)
        print("  terminou antes do pedido de cancelamento")
    except ProcessamentoCancelado:
        print(f"  cancelou {time.perf_counter() - momento['pedido']:.2f}s depois do pedido "
              f"(pedido aos {apos:.1f}s)")


def main(argv=None) -> None:
    p = argparse.ArgumentParser()
    p.add_argument("pasta", type=Path)
    p.add_argument("--titulares", type=int, default=10000)
    p.add_argument("--perfil", action="store_true")
    p.add_argument("--cancelar-apos", type=float, nargs="*")
    p.add_argument("--casos", nargs="*", default=["real", "pesado"])
    args = p.parse_args(argv)
    caminhos = gerar(args.pasta, args.titulares)
    for nome in args.casos:
        print(f"{nome}: {caminhos[nome]}")
        if args.cancelar_apos:
            for apos in args.cancelar_apos:
                medir_cancelamento(caminhos[nome], apos)
        else:
            medir(caminhos[nome], args.pasta / f"saida_{nome}.xlsx", args.perfil)


if __name__ == "__main__":
    main()
