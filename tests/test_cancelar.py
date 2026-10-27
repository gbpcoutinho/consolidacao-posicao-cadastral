"""Fase 2: Cancelar responde dentro do arquivo, nao so entre arquivos."""
import pytest

from consolidacao.erros import ProcessamentoCancelado
from consolidacao.processamento import processar_arquivo, processar_arquivos
from tests.fabrica import DEPENDENTE_OK, SUPLEMENTAR_OK, TITULAR_OK, criar_arquivo_pc

MAPA = {"123": "ACME"}


def _arquivo(tmp_path, nome="PC123 10.2026.xlsx"):
    caminho = tmp_path / nome
    criar_arquivo_pc(caminho, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    return caminho


def _cancela_na(n):
    consultas = []

    def cancelar():
        consultas.append(1)
        return len(consultas) >= n

    cancelar.consultas = consultas
    return cancelar


def test_cancelar_no_meio_do_arquivo(tmp_path):
    cancelar = _cancela_na(2)
    with pytest.raises(ProcessamentoCancelado):
        processar_arquivo(_arquivo(tmp_path), MAPA, cancelar=cancelar)
    assert len(cancelar.consultas) == 2


def test_cancelar_consulta_varias_etapas_do_arquivo(tmp_path):
    # Leitores (3), datas de cada leitor (3) e juncao: ao menos 7 consultas.
    vistos = []
    r = processar_arquivo(_arquivo(tmp_path), MAPA, cancelar=lambda: vistos.append(1) or False)
    assert r.status == "OK"
    assert len(vistos) >= 7


@pytest.mark.parametrize("n", range(1, 8))
def test_cancelar_em_qualquer_etapa_nunca_vira_erro_no_log(tmp_path, n):
    with pytest.raises(ProcessamentoCancelado):
        processar_arquivo(_arquivo(tmp_path), MAPA, cancelar=_cancela_na(n))


def test_processar_arquivos_repassa_cancelar_ao_arquivo(tmp_path):
    # 1a consulta: antes do arquivo (como hoje); 2a: ja dentro do arquivo.
    cancelar = _cancela_na(2)
    with pytest.raises(ProcessamentoCancelado):
        processar_arquivos([_arquivo(tmp_path)], MAPA, cancelar=cancelar)


def test_cancelar_none_nao_muda_resultado(tmp_path):
    caminho = _arquivo(tmp_path)
    a = processar_arquivo(caminho, MAPA)
    b = processar_arquivo(caminho, MAPA, cancelar=lambda: False)
    assert (a.status, a.mensagem, a.avisos) == (b.status, b.mensagem, b.avisos)
    assert a.final.equals(b.final) and a.pendencias.equals(b.pendencias)


def test_leitura_de_aba_grande_consulta_cancelar_por_bloco_de_linhas(tmp_path, monkeypatch):
    from consolidacao import leitura

    caminho = _arquivo(tmp_path)
    contagem = {}
    for bloco in (10_000, 1):
        monkeypatch.setattr(leitura, "LINHAS_ENTRE_CONSULTAS", bloco)
        vistos = []
        processar_arquivo(caminho, MAPA, cancelar=lambda: vistos.append(1) or False)
        contagem[bloco] = len(vistos)
    # Uma consulta por linha lida nas 3 abas (titular 5, dependente 5, suplementar 2).
    assert contagem[1] - contagem[10_000] >= 12 - 3
