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


# C6 ------------------------------------------------------------------------

def test_c6_temporario_comeca_com_til_cifrao_e_some_em_qualquer_erro(tmp_path, monkeypatch):
    from datetime import date
    from pathlib import Path

    import openpyxl

    from consolidacao.exportacao import escrever_excel
    from consolidacao.processamento import listar_arquivos_entrada

    vistos = []

    def falhar(self, nome):
        vistos.append(Path(nome))
        Path(nome).write_bytes(b"meio arquivo")
        raise RuntimeError("falha inesperada do openpyxl")

    monkeypatch.setattr(openpyxl.Workbook, "save", falhar)
    pasta = tmp_path / "saida"
    with pytest.raises(RuntimeError):
        escrever_excel(_consolidado(_tabela()), pasta / "r.xlsx")
    monkeypatch.undo()
    assert vistos and vistos[0].name.startswith("~$")
    assert list(pasta.iterdir()) == []
    # Um temporario esquecido (ex.: processo morto) nao e lido como entrada no proximo run.
    _arquivo_ok(pasta)
    (pasta / vistos[0].name).write_bytes(b"x")
    assert [p.name for p in listar_arquivos_entrada(pasta, date(2026, 10, 23))] == ["PC123 x.xlsx"]


def _tabela():
    import pandas as pd

    return pd.DataFrame({"X": pd.Series(["a"], dtype=object)})


def test_c6_progresso_avisa_a_gravacao_no_fim(tmp_path):
    from consolidacao.pipeline import ETAPA_GRAVACAO, executar

    a, b = _arquivo_ok(tmp_path, "PC123 a.xlsx"), _arquivo_ok(tmp_path, "PC123 b.xlsx")
    vistos = []
    executar([a, b], {"123": "A"}, tmp_path / "o" / "r.xlsx",
             progresso=lambda i, total, nome: vistos.append((i, total, nome)))
    assert vistos == [(1, 2, "PC123 a.xlsx"), (2, 2, "PC123 b.xlsx"), (2, 2, ETAPA_GRAVACAO)]
    assert ETAPA_GRAVACAO == "Gravando arquivo de saida"


def test_c6_erro_no_callback_de_progresso_nao_aborta(tmp_path):
    from consolidacao.pipeline import executar

    def quebrado(i, total, nome):
        raise ValueError("bug na barra de progresso")

    saida = tmp_path / "o" / "r.xlsx"
    c = executar([_arquivo_ok(tmp_path)], {"123": "A"}, saida, progresso=quebrado)
    assert saida.exists() and len(c.consolidado) == 4


def test_c6_cancelar_pelo_callback_de_progresso_continua_valendo(tmp_path):
    from consolidacao.erros import ProcessamentoCancelado
    from consolidacao.pipeline import executar

    def cancela(i, total, nome):
        raise ProcessamentoCancelado("cancelado")

    saida = tmp_path / "o" / "r.xlsx"
    with pytest.raises(ProcessamentoCancelado):
        executar([_arquivo_ok(tmp_path)], {"123": "A"}, saida, progresso=cancela)
    assert not saida.exists()


def test_c6_excecao_inesperada_vira_erro_de_validacao(tmp_path, monkeypatch):
    import consolidacao.pipeline as pipeline

    def explode(c, destino):
        raise KeyError("coluna")

    monkeypatch.setattr(pipeline, "escrever_excel", explode)
    with pytest.raises(ErroValidacao, match="Erro inesperado") as info:
        pipeline.executar([_arquivo_ok(tmp_path)], {"123": "A"}, tmp_path / "r.xlsx")
    assert isinstance(info.value.__cause__, KeyError)


@pytest.mark.parametrize("excecao", [MemoryError, KeyboardInterrupt])
def test_c6_memoria_e_interrupcao_nao_sao_engolidas(tmp_path, monkeypatch, excecao):
    import consolidacao.pipeline as pipeline

    def explode(c, destino):
        raise excecao()

    monkeypatch.setattr(pipeline, "escrever_excel", explode)
    with pytest.raises(excecao):
        pipeline.executar([_arquivo_ok(tmp_path)], {"123": "A"}, tmp_path / "r.xlsx")


# C7 ------------------------------------------------------------------------

@pytest.mark.parametrize("conteudo", [b"isto nao e um xlsx", b"", b"PK\x03\x04truncado"])
def test_c7_arquivo_corrompido_tem_mensagem_em_portugues(tmp_path, conteudo):
    from consolidacao.processamento import processar_arquivos

    ruim = tmp_path / "PC123 ruim.xlsx"
    ruim.write_bytes(conteudo)
    c = processar_arquivos([ruim, _arquivo_ok(tmp_path)], {"123": "A"})
    assert c.log["STATUS"].tolist() == ["ERRO", "OK"]
    assert c.log["MENSAGEM"][0] == "Nao foi possivel abrir o arquivo: arquivo corrompido ou nao e .xlsx."
    assert len(c.consolidado) == 4


def test_c7_arquivo_em_uso_tem_mensagem_em_portugues(tmp_path, monkeypatch):
    import consolidacao.processamento as processamento

    def em_uso(caminho):
        raise PermissionError(13, "Permission denied", str(caminho))

    monkeypatch.setattr(processamento, "abrir_planilha", em_uso)
    r = processamento.processar_arquivo(_arquivo_ok(tmp_path), {"123": "A"})
    assert (r.status, r.mensagem) == (
        "ERRO", "Nao foi possivel abrir o arquivo: arquivo em uso ou sem permissao."
    )
