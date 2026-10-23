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


# C4 ------------------------------------------------------------------------

def test_c4_csv_em_codificacao_desconhecida_vira_erro_de_validacao(tmp_path):
    from consolidacao.empresas import carregar_mapa_empresas

    p = tmp_path / "e.csv"
    # 0x81 nao existe no cp1252 nem e UTF-8 valido.
    p.write_bytes(b"EMPRESA;COD_SAUDE\nA\x81;1\n")
    with pytest.raises(ErroValidacao, match="codificacao"):
        carregar_mapa_empresas(p)


def test_c4_csv_que_e_pasta_vira_erro_de_validacao(tmp_path):
    from consolidacao.empresas import carregar_mapa_empresas

    with pytest.raises(ErroValidacao, match="Nao foi possivel ler o arquivo de empresas"):
        carregar_mapa_empresas(tmp_path)


def test_c4_csv_sem_permissao_vira_erro_de_validacao(tmp_path, monkeypatch):
    from pathlib import Path

    from consolidacao.empresas import carregar_mapa_empresas

    p = tmp_path / "e.csv"
    p.write_bytes(b"EMPRESA;COD_SAUDE\nA;1\n")

    def travado(self):
        raise PermissionError(13, "Permission denied", str(self))

    monkeypatch.setattr(Path, "read_bytes", travado)
    with pytest.raises(ErroValidacao, match="em uso ou sem permissao"):
        carregar_mapa_empresas(p)


# C5 ------------------------------------------------------------------------

def _consolidado(base, pendencias=None, log=None):
    import pandas as pd

    from consolidacao.processamento import Consolidado

    vazio = pd.DataFrame({"X": pd.Series([], dtype=object)})
    return Consolidado(
        consolidado=base,
        pendencias=vazio if pendencias is None else pendencias,
        log=vazio if log is None else log,
    )


@pytest.mark.parametrize("aba", ["base", "pendencias"])
def test_c5_aba_acima_do_limite_do_excel_e_recusada_antes_de_gravar(tmp_path, aba):
    import pandas as pd

    from consolidacao.exportacao import escrever_excel

    grande = pd.DataFrame({"X": pd.Series([None] * 1_048_576, dtype=object)})
    pequeno = pd.DataFrame({"X": pd.Series(["a"], dtype=object)})
    c = _consolidado(grande, pequeno) if aba == "base" else _consolidado(pequeno, grande)
    destino = tmp_path / "saida.xlsx"
    with pytest.raises(ErroValidacao, match="1.048.575"):
        escrever_excel(c, destino)
    assert list(tmp_path.iterdir()) == []


def test_c5_aba_no_limite_exato_e_aceita():
    import pandas as pd

    from consolidacao.exportacao import validar_limite_linhas

    no_limite = pd.DataFrame({"X": pd.Series([None] * 1_048_575, dtype=object)})
    validar_limite_linhas(_consolidado(no_limite))
