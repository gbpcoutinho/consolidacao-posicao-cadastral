"""Achados da revisao por 5 agentes (rodada 2, 2026-10-23); ids C1..C15 do consenso."""
import pytest

from consolidacao.erros import ErroValidacao
from tests.fabrica import DEPENDENTE_OK, SUPLEMENTAR_OK, TITULAR_OK, criar_arquivo_pc


def _arquivo_ok(pasta, nome="PC123 x.xlsx"):
    p = pasta / nome
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    return p


# C1 ------------------------------------------------------------------------

def test_c1_saida_igual_a_entrada_e_recusada_sem_tocar_no_arquivo(tmp_path):
    from consolidacao.pipeline import executar

    a = _arquivo_ok(tmp_path)
    original = a.read_bytes()
    # Mesmo arquivo escrito de outro jeito (caixa diferente, caminho com "..").
    outro_nome = tmp_path / "sub" / ".." / a.name.upper()
    (tmp_path / "sub").mkdir()
    for saida in (a, outro_nome):
        with pytest.raises(ErroValidacao, match="e um dos arquivos de entrada"):
            executar([a], {"123": "A"}, saida)
    assert a.read_bytes() == original


def test_c1_saida_igual_ao_csv_de_empresas_e_recusada(tmp_path):
    from consolidacao.pipeline import executar

    a = _arquivo_ok(tmp_path)
    csv = tmp_path / "e.csv"
    csv.write_bytes(b"EMPRESA;COD_SAUDE\nA;123\n")
    with pytest.raises(ErroValidacao, match="e o arquivo de empresas"):
        executar([a], {"123": "A"}, csv, arquivo_empresas=csv)
    assert csv.read_bytes() == b"EMPRESA;COD_SAUDE\nA;123\n"


def test_c1_cli_recusa_saida_igual_a_entrada(tmp_path, capsys):
    from tests.fabrica import escrever_csv_empresas
    from tratamento import main

    a = _arquivo_ok(tmp_path)
    escrever_csv_empresas(tmp_path / "e.csv", [("ACME", "123")])
    codigo = main(["--entrada", str(tmp_path), "--empresas", str(tmp_path / "e.csv"),
                   "--saida", str(a)])
    assert codigo == 1
    assert "e um dos arquivos de entrada" in capsys.readouterr().out
