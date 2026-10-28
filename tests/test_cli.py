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


def test_cli_senhas_de_entrada_e_de_saida(tmp_path, capsys):
    import io

    import msoffcrypto

    from tests.fabrica import cifrar_arquivo

    plano = tmp_path / "plano" / "PC123 x.xlsx"
    plano.parent.mkdir()
    criar_arquivo_pc(plano, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    entrada = tmp_path / "entrada"
    entrada.mkdir()
    cifrar_arquivo(plano, entrada / plano.name, "chave-2")
    escrever_csv_empresas(tmp_path / "e.csv", [("ACME", "123")])
    saida = tmp_path / "r.xlsx"
    codigo = main([
        "--entrada", str(entrada), "--empresas", str(tmp_path / "e.csv"), "--saida", str(saida),
        "--senha", "chave-1", "--senha", "chave-2", "--senha-saida", "saida-9",
    ])
    out = capsys.readouterr().out
    assert codigo == 0
    assert "Arquivos processados com sucesso: 1" in out
    assert "saida-9" not in out and "chave-" not in out
    f = msoffcrypto.OfficeFile(open(saida, "rb"))
    f.load_key(password="saida-9")
    f.decrypt(io.BytesIO())


def test_cli_arquivo_protegido_sem_senha_e_erro_do_arquivo(tmp_path, capsys):
    from tests.fabrica import cifrar_arquivo

    plano = tmp_path / "plano" / "PC123 x.xlsx"
    plano.parent.mkdir()
    criar_arquivo_pc(plano, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    entrada = tmp_path / "entrada"
    entrada.mkdir()
    cifrar_arquivo(plano, entrada / plano.name, "chave-2")
    escrever_csv_empresas(tmp_path / "e.csv", [("ACME", "123")])
    codigo = main(["--entrada", str(entrada), "--empresas", str(tmp_path / "e.csv"),
                   "--saida", str(tmp_path / "r.xlsx")])
    out = capsys.readouterr().out
    assert codigo == 0 and "Arquivos com erro: 1" in out
    assert "cadastre a senha na aba seguranca" in out.casefold()
