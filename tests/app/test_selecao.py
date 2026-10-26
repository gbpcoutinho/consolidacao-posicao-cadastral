import os
from datetime import date, datetime
from pathlib import Path

from app.empresas_compartilhadas import Relacao
from app.selecao import (
    ItemArquivo,
    adicionar,
    adicionar_pasta,
    alertas_pre_processamento,
    competencia,
    mudancas,
    nome_saida_sugerido,
    retratar,
)

MAPA = {"90001": "ALFA", "90002": "BETA"}


def _arquivos(pasta, *nomes):
    pasta.mkdir(parents=True, exist_ok=True)
    caminhos = []
    for nome in nomes:
        p = pasta / nome
        p.write_bytes(b"xlsx ficticio")
        caminhos.append(p)
    return caminhos


def _relacao(hash_="h1"):
    return Relacao([], hash_, "rede", datetime(2026, 10, 26))


def test_ignora_temporario_e_consolidado(tmp_path):
    caminhos = _arquivos(
        tmp_path, "PC90001 10.2026.xlsx", "~$PC90001 10.2026.xlsx",
        "PC_CONSOLIDADO_10_2026.xlsx", "pc_consolidado_x.XLSX", "nota.txt", "PC90002 10.2026.xls",
    )
    itens = adicionar([], caminhos, MAPA)
    assert [i.caminho.name for i in itens] == ["PC90001 10.2026.xlsx"]


def test_sem_duplicados_maiusculas(tmp_path):
    (a,) = _arquivos(tmp_path, "PC90001 10.2026.xlsx")
    itens = adicionar([], [a], MAPA)
    itens = adicionar(itens, [Path(str(a).upper()), a, tmp_path / "." / a.name], MAPA)
    assert len(itens) == 1


def test_adicionar_nao_altera_a_lista_original(tmp_path):
    (a,) = _arquivos(tmp_path, "PC90001 10.2026.xlsx")
    original = []
    adicionar(original, [a], MAPA)
    assert original == []


def test_empresa_e_competencia_preenchidas(tmp_path):
    (a,) = _arquivos(tmp_path, "PC90001 10.2026.xlsx")
    (item,) = adicionar([], [a], MAPA)
    assert item == ItemArquivo(a, "ALFA", "90001", (2026, 10), None)


def test_empresa_nao_cadastrada_tem_erro(tmp_path):
    a, b = _arquivos(tmp_path, "PC77777 10.2026.xlsx", "SEM CODIGO.xlsx")
    x, y = adicionar([], [a, b], MAPA)
    assert x.empresa is None and "77777" in x.erro
    assert y.empresa is None and "PC" in y.erro


def test_adicionar_pasta_ordena_e_filtra(tmp_path):
    _arquivos(tmp_path, "PC90002 10.2026.xlsx", "PC90001 10.2026.xlsx", "~$PC90001 10.2026.xlsx")
    (tmp_path / "sub.xlsx").mkdir()
    itens = adicionar_pasta([], tmp_path, MAPA)
    assert [i.caminho.name for i in itens] == ["PC90001 10.2026.xlsx", "PC90002 10.2026.xlsx"]
    assert len(adicionar_pasta(itens, tmp_path, MAPA)) == 2


def test_competencia_mm_aaaa():
    assert competencia("PC90001 10.2026.xlsx") == (2026, 10)
    assert competencia("PC90001 01.2027 v2.xlsx") == (2027, 1)
    assert competencia("PC90001.xlsx") is None


def test_competencia_mes_invalido_none():
    assert competencia("PC90001 13.2026.xlsx") is None
    assert competencia("PC90001 00.2026.xlsx") is None
    assert competencia("PC123 10.20261.xlsx") is None


def test_alerta_competencias_misturadas(tmp_path):
    caminhos = _arquivos(tmp_path, "PC90001 10.2026.xlsx", "PC90002 09.2026.xlsx")
    alertas = alertas_pre_processamento(adicionar([], caminhos, MAPA))
    assert any("09/2026" in a and "10/2026" in a for a in alertas)


def test_alerta_sem_competencia_e_nao_cadastrada(tmp_path):
    caminhos = _arquivos(tmp_path, "PC90001.xlsx", "PC77777 10.2026.xlsx")
    alertas = alertas_pre_processamento(adicionar([], caminhos, MAPA))
    assert any("PC90001.xlsx" in a and "competencia" in a.lower() for a in alertas)
    assert any("PC77777 10.2026.xlsx" in a for a in alertas)


def test_alerta_mesma_empresa_em_dois_arquivos(tmp_path):
    caminhos = _arquivos(tmp_path, "PC90001 10.2026.xlsx", "PC90001 10.2026 v2.xlsx")
    alertas = alertas_pre_processamento(adicionar([], caminhos, MAPA))
    assert any("90001" in a and "mais de um arquivo" in a for a in alertas)


def test_sem_alertas_quando_tudo_certo(tmp_path):
    caminhos = _arquivos(tmp_path, "PC90001 10.2026.xlsx", "PC90002 10.2026.xlsx")
    assert alertas_pre_processamento(adicionar([], caminhos, MAPA)) == []


def test_nome_sugerido_usa_competencia(tmp_path):
    caminhos = _arquivos(tmp_path, "PC90001 10.2026.xlsx", "PC90002 10.2026.xlsx")
    itens = adicionar([], caminhos, MAPA)
    assert nome_saida_sugerido(itens, date(2026, 11, 3)) == "PC_CONSOLIDADO_10_2026.xlsx"


def test_nome_sugerido_sem_competencia_unica_usa_hoje(tmp_path):
    caminhos = _arquivos(tmp_path, "PC90001 10.2026.xlsx", "PC90002 09.2026.xlsx")
    itens = adicionar([], caminhos, MAPA)
    assert nome_saida_sugerido(itens, date(2026, 11, 3)) == "PC_CONSOLIDADO_11_2026.xlsx"
    assert nome_saida_sugerido([], date(2026, 11, 3)) == "PC_CONSOLIDADO_11_2026.xlsx"


def test_mudancas_nada_mudou(tmp_path):
    itens = adicionar([], _arquivos(tmp_path, "PC90001 10.2026.xlsx"), MAPA)
    retrato = retratar(itens, _relacao())
    assert mudancas(retrato, itens, _relacao()) == []


def test_mudancas_detecta_entrada_alterada(tmp_path):
    (a,) = _arquivos(tmp_path, "PC90001 10.2026.xlsx")
    itens = adicionar([], [a], MAPA)
    retrato = retratar(itens, _relacao())
    a.write_bytes(b"outro conteudo, outro tamanho")
    avisos = mudancas(retrato, itens, _relacao())
    assert len(avisos) == 1 and "PC90001 10.2026.xlsx" in avisos[0]


def test_mudancas_detecta_so_data_alterada(tmp_path):
    (a,) = _arquivos(tmp_path, "PC90001 10.2026.xlsx")
    itens = adicionar([], [a], MAPA)
    retrato = retratar(itens, _relacao())
    os.utime(a, (a.stat().st_atime, a.stat().st_mtime + 60))
    assert mudancas(retrato, itens, _relacao()) != []


def test_mudancas_detecta_entrada_apagada_e_lista_alterada(tmp_path):
    a, b = _arquivos(tmp_path, "PC90001 10.2026.xlsx", "PC90002 10.2026.xlsx")
    itens = adicionar([], [a], MAPA)
    retrato = retratar(itens, _relacao())
    a.unlink()
    assert any("PC90001" in m for m in mudancas(retrato, itens, _relacao()))
    a.write_bytes(b"xlsx ficticio")
    os.utime(a, (retrato.entradas[next(iter(retrato.entradas))][1],) * 2)
    outros = adicionar(itens, [b], MAPA)
    assert any("lista" in m.lower() for m in mudancas(retrato, outros, _relacao()))


def test_mudancas_detecta_relacao_alterada(tmp_path):
    itens = adicionar([], _arquivos(tmp_path, "PC90001 10.2026.xlsx"), MAPA)
    retrato = retratar(itens, _relacao("h1"))
    avisos = mudancas(retrato, itens, _relacao("h2"))
    assert len(avisos) == 1 and "relacao de empresas" in avisos[0].lower()
