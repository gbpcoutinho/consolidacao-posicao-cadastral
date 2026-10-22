"""Gera arquivos sinteticos (dados falsos) para os testes."""
from pathlib import Path


def escrever_csv_empresas(caminho: Path, linhas: list, encoding: str = "utf-8-sig") -> None:
    texto = "EMPRESA;COD_SAUDE;COD_DENTAL\n" + "".join(
        f"{empresa};{codigo};\n" for empresa, codigo in linhas
    )
    Path(caminho).write_bytes(texto.encode(encoding))
