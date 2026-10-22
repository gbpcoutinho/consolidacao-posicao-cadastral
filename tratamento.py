"""Consolidacao de posicao cadastral - titulares + dependentes.

Conversao em Python do TRATAMENTO.R. Uso:
    python tratamento.py [--entrada PASTA] [--empresas CSV] [--saida XLSX]
Sem argumentos, usa a pasta PC_CONSOLIDADO_<mes>_<ano> ao lado deste script.
"""
import argparse
import sys
from datetime import date
from pathlib import Path

from consolidacao.config import nome_saida, pasta_padrao
from consolidacao.empresas import carregar_mapa_empresas
from consolidacao.erros import ErroValidacao
from consolidacao.exportacao import escrever_excel
from consolidacao.processamento import listar_arquivos_entrada, processar_arquivos


def main(argv: list | None = None) -> int:
    raiz = Path(__file__).resolve().parent
    hoje = date.today()

    parser = argparse.ArgumentParser(description="Consolida planilhas de posicao cadastral.")
    parser.add_argument("--entrada", type=Path, default=pasta_padrao(raiz, hoje))
    parser.add_argument("--empresas", type=Path, default=raiz / "COD_EMPRESAS.csv")
    parser.add_argument("--saida", type=Path, default=None)
    args = parser.parse_args(argv)
    saida = args.saida or args.entrada / nome_saida(hoje)

    try:
        mapa = carregar_mapa_empresas(args.empresas)
        arquivos = listar_arquivos_entrada(args.entrada, hoje)
    except ErroValidacao as e:
        print(f"Erro: {e}")
        return 1

    print(f"Arquivos de entrada encontrados: {len(arquivos)} \n")
    resultado = processar_arquivos(
        arquivos, mapa, progresso=lambda nome: print(f"Processando: {nome} ")
    )
    escrever_excel(resultado, saida)

    log = resultado.log
    arquivos_ok = int((log["STATUS"] == "OK").sum())
    arquivos_erro = int((log["STATUS"] == "ERRO").sum())
    print("\n====================================")
    print("RESUMO DO PROCESSAMENTO")
    print("====================================")
    print(f"Arquivos processados com sucesso: {arquivos_ok} ")
    print(f"Arquivos com erro: {arquivos_erro} ")
    print(f"Registros consolidados: {len(resultado.consolidado)} ")
    print(f"Pendencias encontradas: {len(resultado.pendencias)} ")
    print(f"Valores de data invalidos registrados: {int(log['DATAS_INVALIDAS'].sum())} ")
    print(f"Divergencias entre datas principais e Y2K: {int(log['QTD_DIVERGENCIAS_DATAS'].sum())} ")
    if arquivos_erro > 0:
        print("\nArquivos com erro:")
        for _, linha in log[log["STATUS"] == "ERRO"].iterrows():
            print(f" - {linha['ARQUIVO']} : {linha['MENSAGEM']} ")
    print("====================================")
    print(f"Arquivo gerado: {saida} ")
    print("Abas: Base Tratada, Pendencias, Log")
    return 0


if __name__ == "__main__":
    sys.exit(main())
