from tests.fabrica import (
    DEPENDENTE_OK,
    SUPLEMENTAR_OK,
    TITULAR_OK,
    criar_arquivo_pc,
    escrever_csv_empresas,
)
from tratamento import main


def test_cli_gera_arquivo_e_resumo(tmp_path, capsys):
    criar_arquivo_pc(tmp_path / "PC123 x.xlsx", TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    escrever_csv_empresas(tmp_path / "e.csv", [("ACME", "123")])
    saida = tmp_path / "r.xlsx"
    codigo = main(
        ["--entrada", str(tmp_path), "--empresas", str(tmp_path / "e.csv"), "--saida", str(saida)]
    )
    out = capsys.readouterr().out
    assert codigo == 0
    assert saida.exists()
    assert "Arquivos de entrada encontrados: 1" in out
    assert "Processando: PC123 x.xlsx" in out
    assert "Arquivos processados com sucesso: 1" in out
    assert "Pendencias encontradas: 1" in out
    assert "Abas: Base Tratada, Pendencias, Log" in out


def test_cli_erro_de_entrada(tmp_path, capsys):
    escrever_csv_empresas(tmp_path / "e.csv", [("ACME", "123")])
    codigo = main(["--entrada", str(tmp_path / "nao_existe"), "--empresas", str(tmp_path / "e.csv")])
    assert codigo == 1
    assert "Erro: A pasta de entrada nao existe" in capsys.readouterr().out
