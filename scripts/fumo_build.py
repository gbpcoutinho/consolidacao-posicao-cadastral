"""Fumo descartavel do PyInstaller: importa as dependencias e processa um exemplo.

Uso: fumo_build.exe <planilha.xlsx> <COD_EMPRESAS.csv>
"""
import sys
from pathlib import Path

import openpyxl  # noqa: F401
import pandas  # noqa: F401
import tkinter  # noqa: F401

from consolidacao.empresas import carregar_mapa_empresas
from consolidacao.processamento import processar_arquivo


def main(argv: list) -> int:
    planilha, empresas = Path(argv[0]), Path(argv[1])
    r = processar_arquivo(planilha, carregar_mapa_empresas(empresas))
    if r.status != "OK":
        print(f"FALHA: {r.mensagem}")
        return 1
    print(f"OK {len(r.final)} registros")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
