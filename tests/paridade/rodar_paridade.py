"""Roda TRATAMENTO.R e tratamento.py nos mesmos arquivos e compara as saidas.

Uso:
    .venv\\Scripts\\python tests\\paridade\\rodar_paridade.py --entrada dados_teste --empresas COD_EMPRESAS.csv
Tudo roda localmente, em tests/paridade/_trabalho/.
"""
import argparse
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from consolidacao.config import nome_saida, pasta_padrao  # noqa: E402
from tests.paridade.comparar import comparar_planilhas  # noqa: E402
from tratamento import main as tratamento_main  # noqa: E402

RSCRIPT_PADRAO = Path(r"C:\Program Files\R\R-4.6.1\bin\Rscript.exe")
TRABALHO = Path(__file__).resolve().parent / "_trabalho"


def rodar(
    entrada: Path, empresas: Path, rscript: Path = RSCRIPT_PADRAO, trabalho: Path = TRABALHO
) -> list:
    entrada, empresas = Path(entrada).resolve(), Path(empresas).resolve()
    hoje = date.today()

    # O OneDrive pode travar pastas recem-esvaziadas; cada execucao usa uma
    # pasta nova e as execucoes antigas sao removidas quando possivel.
    if trabalho.exists():
        for antiga in trabalho.glob("execucao_*"):
            shutil.rmtree(antiga, ignore_errors=True)
    execucao = trabalho / datetime.now().strftime("execucao_%Y%m%d_%H%M%S_%f")
    r_raiz = execucao / "r_raiz"
    pasta_r = pasta_padrao(r_raiz, hoje)
    pasta_r.mkdir(parents=True)
    # .here faz o here::here() do R apontar para r_raiz.
    (r_raiz / ".here").write_bytes(b"")
    shutil.copy2(RAIZ / "TRATAMENTO.R", r_raiz / "TRATAMENTO.R")
    shutil.copy2(empresas, r_raiz / "COD_EMPRESAS.csv")
    for arquivo in entrada.iterdir():
        if arquivo.is_file() and arquivo.suffix.lower() == ".xlsx":
            shutil.copy2(arquivo, pasta_r / arquivo.name)

    processo = subprocess.run(
        [str(rscript), "TRATAMENTO.R"], cwd=r_raiz, capture_output=True, timeout=300
    )
    saida_r = pasta_r / nome_saida(hoje)
    if processo.returncode != 0 or not saida_r.exists():
        erro = processo.stderr.decode("utf-8", errors="replace")[-2000:]
        return [f"R falhou (codigo {processo.returncode}): {erro}"]

    saida_py = execucao / "py.xlsx"
    codigo = tratamento_main(
        ["--entrada", str(entrada), "--empresas", str(empresas), "--saida", str(saida_py)]
    )
    if codigo != 0:
        return [f"Python falhou (codigo {codigo})"]

    return comparar_planilhas(saida_r, saida_py)


def principal() -> int:
    parser = argparse.ArgumentParser(description="Paridade TRATAMENTO.R x tratamento.py")
    parser.add_argument("--entrada", type=Path, required=True)
    parser.add_argument("--empresas", type=Path, required=True)
    parser.add_argument("--rscript", type=Path, default=RSCRIPT_PADRAO)
    args = parser.parse_args()
    difs = rodar(args.entrada, args.empresas, args.rscript)
    print("\n==== PARIDADE ====")
    for d in difs:
        print(d)
    print(f"{len(difs)} diferencas")
    return 0 if not difs else 1


if __name__ == "__main__":
    sys.exit(principal())
