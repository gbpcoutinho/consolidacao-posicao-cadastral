"""Consolidacao de posicao cadastral - titulares + dependentes.

Conversao em Python do TRATAMENTO.R. Uso:
    python tratamento.py [--entrada PASTA] [--empresas CSV] [--saida XLSX]
                         [--senha SENHA ...] [--senha-saida SENHA]
Sem argumentos, usa a pasta PC_CONSOLIDADO_<mes>_<ano> ao lado deste script.
"""
import argparse
import sys
from datetime import date
from pathlib import Path

from consolidacao.config import nome_saida, pasta_padrao
from consolidacao.empresas import carregar_mapa_empresas
from consolidacao.erros import ErroValidacao
from consolidacao.pipeline import ETAPA_GRAVACAO, executar, resumir
from consolidacao.processamento import listar_arquivos_entrada


def _mostrar_progresso(i: int, total: int, nome: str) -> None:
    # O R so imprime o nome de cada arquivo.
    if nome != ETAPA_GRAVACAO:
        print(f"Processando: {nome} ")


def main(argv: list | None = None, hoje: date | None = None) -> int:
    # Padroes relativos a este script: valem so para a linha de comando.
    raiz = Path(__file__).resolve().parent
    hoje = hoje or date.today()

    parser = argparse.ArgumentParser(description="Consolida planilhas de posicao cadastral.")
    parser.add_argument("--entrada", type=Path, default=pasta_padrao(raiz, hoje))
    parser.add_argument("--empresas", type=Path, default=raiz / "COD_EMPRESAS.csv")
    parser.add_argument("--saida", type=Path, default=None)
    parser.add_argument(
        "--senha", action="append", default=[],
        help="senha de arquivos de entrada protegidos; repita para varias (testadas em ordem)",
    )
    parser.add_argument(
        "--senha-saida", default=None,
        help="protege o arquivo consolidado com esta senha (omitida = sem senha)",
    )
    args = parser.parse_args(argv)
    saida = args.saida or args.entrada / nome_saida(hoje)

    try:
        mapa = carregar_mapa_empresas(args.empresas)
        arquivos = listar_arquivos_entrada(args.entrada, hoje)
        print(f"Arquivos de entrada encontrados: {len(arquivos)} \n")
        resultado = executar(
            arquivos, mapa, saida, arquivo_empresas=args.empresas,
            progresso=_mostrar_progresso,
            senhas=args.senha, senha_saida=args.senha_saida,
        )
    except ErroValidacao as e:
        print(f"Erro: {e}")
        return 1

    r = resumir(resultado)
    print("\n====================================")
    print("RESUMO DO PROCESSAMENTO")
    print("====================================")
    print(f"Arquivos processados com sucesso: {r['arquivos_ok']} ")
    print(f"Arquivos com erro: {r['arquivos_erro']} ")
    print(f"Registros consolidados: {r['registros']} ")
    print(f"Pendencias encontradas: {r['pendencias']} ")
    print(f"Valores de data invalidos registrados: {r['datas_invalidas']} ")
    print(f"Divergencias entre datas principais e Y2K: {r['divergencias_datas']} ")
    if r["erros"]:
        print("\nArquivos com erro:")
        for arquivo, mensagem in r["erros"]:
            print(f" - {arquivo} : {mensagem} ")
    print("====================================")
    print(f"Arquivo gerado: {saida} " + ("(protegido por senha) " if args.senha_saida else ""))
    print("Abas: Base Tratada, Pendencias, Log")
    return 0


if __name__ == "__main__":
    sys.exit(main())
