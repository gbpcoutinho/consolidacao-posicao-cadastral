from datetime import date

import pytest

from consolidacao.config import ORDEM_INICIAL
from consolidacao.erros import ErroValidacao
from consolidacao.processamento import listar_arquivos_entrada, processar_arquivos
from tests.fabrica import DEPENDENTE_OK, SUPLEMENTAR_OK, TITULAR_OK, criar_arquivo_pc

HOJE = date(2026, 10, 22)


def test_listar_arquivos_entrada_filtra(tmp_path):
    for n in ["PC2 a.xlsx", "PC1 b.XLSX", "~$PC1 b.xlsx", "PC_CONSOLIDADO_9_2025.xlsx",
              "Consolidado - tratado.xlsx", "nota.txt"]:
        (tmp_path / n).write_bytes(b"x")
    (tmp_path / "sub.xlsx").mkdir()
    assert [p.name for p in listar_arquivos_entrada(tmp_path, HOJE)] == ["PC1 b.XLSX", "PC2 a.xlsx"]


def test_listar_pasta_inexistente_e_vazia(tmp_path):
    with pytest.raises(ErroValidacao, match="A pasta de entrada nao existe: "):
        listar_arquivos_entrada(tmp_path / "x", HOJE)
    with pytest.raises(ErroValidacao, match="Nenhum arquivo .xlsx de entrada encontrado em: "):
        listar_arquivos_entrada(tmp_path, HOJE)


def test_processar_arquivos_mistura_ok_e_erro(tmp_path):
    a = tmp_path / "PC123 x.xlsx"
    criar_arquivo_pc(a, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    b = tmp_path / "PC999 y.xlsx"
    criar_arquivo_pc(b, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    vistos = []
    c = processar_arquivos([a, b], {"123": "ACME"},
                           progresso=lambda i, total, nome: vistos.append(nome))
    assert vistos == ["PC123 x.xlsx", "PC999 y.xlsx"]
    assert c.log["STATUS"].tolist() == ["OK", "ERRO"]
    assert list(c.log.columns) == [
        "ARQUIVO", "REFERENCIA", "CODIGO_EMPRESA", "EMPRESA", "STATUS", "TITULARES",
        "DEPENDENTES", "AVISOS_ESTRUTURA", "DATAS_INVALIDAS", "OBSERVACOES_DATAS",
        "QTD_DIVERGENCIAS_DATAS", "DIVERGENCIAS_DATAS", "MENSAGEM",
        "OBSERVACAO_IDENTIFICADORES",
    ]
    assert c.log["TITULARES"].tolist() == [2, None]
    assert list(c.consolidado.columns[:13]) == ORDEM_INICIAL
    assert len(c.consolidado) == 4 and len(c.pendencias) == 1
    # Ordenado por REFERENCIA, SubF, NUMERO DO CERTIFICADO (estavel).
    assert c.consolidado["NUMERO DO CERTIFICADO"].tolist() == ["001", "001", "002", "999"]


def test_processar_arquivos_todos_com_erro(tmp_path):
    b = tmp_path / "PC999.xlsx"
    criar_arquivo_pc(b, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    c = processar_arquivos([b], {"1": "A"})
    assert len(c.consolidado) == 0 and "CÓDIGO UNICO" in c.consolidado.columns
    assert list(c.pendencias.columns)[-1] == "MOTIVO"
