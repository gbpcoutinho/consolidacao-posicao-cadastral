"""Fase 2: processar e gravar separados (o "Salvar como" vem depois do resumo)."""
import hashlib

import pytest

from consolidacao.erros import ErroValidacao
from tests.fabrica import DEPENDENTE_OK, SUPLEMENTAR_OK, TITULAR_OK, criar_arquivo_pc
from tests.paridade.comparar import comparar_planilhas

MAPA = {"123": "ACME", "456": "BETA"}


def _entradas(pasta):
    pasta.mkdir(parents=True, exist_ok=True)
    arquivos = []
    for nome in ("PC123 10.2026.xlsx", "PC456 10.2026.xlsx"):
        caminho = pasta / nome
        criar_arquivo_pc(caminho, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
        arquivos.append(caminho)
    return arquivos


def _hash(caminho):
    return hashlib.sha256(caminho.read_bytes()).hexdigest()


def test_processar_nao_grava_nada(tmp_path):
    from consolidacao.pipeline import processar

    arquivos = _entradas(tmp_path / "entrada")
    antes = sorted(p.name for p in tmp_path.rglob("*"))
    c = processar(arquivos, MAPA)
    assert sorted(p.name for p in tmp_path.rglob("*")) == antes
    assert c.log["STATUS"].tolist() == ["OK", "OK"]
    assert len(c.consolidado) == 8


def test_processar_repassa_progresso_e_cancelar(tmp_path):
    from consolidacao.erros import ProcessamentoCancelado
    from consolidacao.pipeline import processar

    arquivos = _entradas(tmp_path / "entrada")
    vistos = []
    processar(arquivos, MAPA, progresso=lambda i, n, nome: vistos.append((i, n, nome)))
    assert vistos == [(1, 2, "PC123 10.2026.xlsx"), (2, 2, "PC456 10.2026.xlsx")]
    with pytest.raises(ProcessamentoCancelado):
        processar(arquivos, MAPA, cancelar=lambda: True)


def test_processar_excecao_inesperada_vira_erro_de_validacao(tmp_path, monkeypatch):
    from consolidacao import pipeline

    def explode(*a, **k):
        raise KeyError("x")

    monkeypatch.setattr(pipeline, "processar_arquivos", explode)
    with pytest.raises(ErroValidacao, match="Erro inesperado") as info:
        pipeline.processar(_entradas(tmp_path / "e"), MAPA)
    assert isinstance(info.value.__cause__, KeyError)


def test_gravar_recusa_saida_igual_entrada(tmp_path):
    from consolidacao.pipeline import gravar, processar

    arquivos = _entradas(tmp_path / "entrada")
    c = processar(arquivos, MAPA)
    antes = _hash(arquivos[0])
    with pytest.raises(ErroValidacao, match="e um dos arquivos de entrada"):
        gravar(c, arquivos[0], arquivos)
    assert _hash(arquivos[0]) == antes


def test_gravar_recusa_saida_igual_ao_arquivo_de_empresas(tmp_path):
    from consolidacao.pipeline import gravar, processar

    arquivos = _entradas(tmp_path / "entrada")
    csv = tmp_path / "empresas.csv"
    csv.write_text("x", encoding="utf-8")
    with pytest.raises(ErroValidacao, match="e o arquivo de empresas"):
        gravar(processar(arquivos, MAPA), csv, arquivos, arquivo_empresas=csv)
    assert csv.read_text(encoding="utf-8") == "x"


def test_gravar_avisa_etapa_gravacao(tmp_path):
    from consolidacao.pipeline import ETAPA_GRAVACAO, gravar, processar

    arquivos = _entradas(tmp_path / "entrada")
    c = processar(arquivos, MAPA)
    vistos = []
    saida = tmp_path / "saida" / "r.xlsx"
    gravar(c, saida, arquivos, progresso=lambda i, n, nome: vistos.append((i, n, nome)))
    assert vistos == [(2, 2, ETAPA_GRAVACAO)]
    assert saida.exists()


def test_gravar_excecao_inesperada_vira_erro_de_validacao(tmp_path, monkeypatch):
    from consolidacao import pipeline

    arquivos = _entradas(tmp_path / "entrada")
    c = pipeline.processar(arquivos, MAPA)

    def explode(c, destino, senha=None):
        raise KeyError("x")

    monkeypatch.setattr(pipeline, "escrever_excel", explode)
    with pytest.raises(ErroValidacao, match="Erro inesperado"):
        pipeline.gravar(c, tmp_path / "r.xlsx", arquivos)


@pytest.mark.parametrize("excecao", [MemoryError, KeyboardInterrupt])
def test_gravar_memoria_e_interrupcao_passam_direto(tmp_path, monkeypatch, excecao):
    from consolidacao import pipeline

    arquivos = _entradas(tmp_path / "entrada")
    c = pipeline.processar(arquivos, MAPA)

    def explode(c, destino, senha=None):
        raise excecao()

    monkeypatch.setattr(pipeline, "escrever_excel", explode)
    with pytest.raises(excecao):
        pipeline.gravar(c, tmp_path / "r.xlsx", arquivos)


def test_executar_igual_processar_mais_gravar(tmp_path):
    from consolidacao.pipeline import executar, gravar, processar

    arquivos = _entradas(tmp_path / "entrada")
    a, b = tmp_path / "a" / "a.xlsx", tmp_path / "b" / "b.xlsx"
    executar(arquivos, MAPA, a)
    gravar(processar(arquivos, MAPA), b, arquivos)
    assert comparar_planilhas(a, b) == []


def test_executar_valida_saida_antes_de_processar(tmp_path, monkeypatch):
    from consolidacao import pipeline

    arquivos = _entradas(tmp_path / "entrada")
    chamado = []
    monkeypatch.setattr(pipeline, "processar_arquivos", lambda *a, **k: chamado.append(1))
    with pytest.raises(ErroValidacao):
        pipeline.executar(arquivos, MAPA, arquivos[0])
    assert chamado == []


def test_erro_inesperado_e_subclasse_de_erro_de_validacao(tmp_path, monkeypatch):
    from consolidacao import pipeline
    from consolidacao.erros import ErroInesperado

    monkeypatch.setattr(pipeline, "processar_arquivos", lambda *a, **k: {}["x"])
    with pytest.raises(ErroInesperado) as info:
        pipeline.processar([], MAPA)
    assert isinstance(info.value, ErroValidacao)
    assert str(info.value).startswith("Erro inesperado ao gerar a consolidacao (KeyError")
