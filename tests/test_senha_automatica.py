"""Senha automatica: padrao informado pelo usuario ({apolice} = codigo apos PC no nome),
tentada antes das cadastradas. So valores ficticios: o padrao real nao fica no repo."""
import logging

import pytest

from consolidacao import protecao
from consolidacao.processamento import processar_arquivo, processar_arquivos
from tests.fabrica import DEPENDENTE_OK, SUPLEMENTAR_OK, TITULAR_OK, cifrar_arquivo, criar_arquivo_pc

MAPA = {"90001": "ACME", "00123": "ZERO"}
PADRAO = "{apolice}#teste"


def _cifrado(pasta, nome, senha):
    pasta.mkdir(parents=True, exist_ok=True)
    plano = pasta / "plano" / nome
    plano.parent.mkdir(exist_ok=True)
    criar_arquivo_pc(plano, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    destino = pasta / nome
    cifrar_arquivo(plano, destino, senha)
    return destino


def test_abre_com_a_senha_automatica_sem_senhas_cadastradas(tmp_path):
    f = _cifrado(tmp_path, "PC90001 10.2026.xlsx", "90001#teste")
    assert processar_arquivo(f, MAPA, padrao_senha=PADRAO).status == "OK"


def test_sem_padrao_ou_desativada_nao_abre(tmp_path):
    f = _cifrado(tmp_path, "PC90001 10.2026.xlsx", "90001#teste")
    for padrao in (None, ""):  # nao ha padrao de fabrica: vazio = desligada
        r = processar_arquivo(f, MAPA, padrao_senha=padrao)
        assert r.status == "ERRO" and "cadastre a senha" in r.mensagem
    assert processar_arquivo(f, MAPA).status == "ERRO"  # o motor nao assume padrao


def test_padrao_customizado(tmp_path):
    f = _cifrado(tmp_path, "PC90001 10.2026.xlsx", "Apolice-90001!")
    assert processar_arquivo(f, MAPA, padrao_senha=PADRAO).status == "ERRO"
    assert processar_arquivo(f, MAPA, padrao_senha="Apolice-{apolice}!").status == "OK"


def test_zeros_do_nome_sao_mantidos(tmp_path):
    f = _cifrado(tmp_path, "PC00123 10.2026.xlsx", "00123#teste")
    assert processar_arquivo(f, MAPA, padrao_senha=PADRAO).status == "OK"


def test_ordem_automatica_primeiro_e_sem_repeticao():
    c = protecao.candidatas("PC90001 x.xlsx", ["90001#teste", "b", "b", ""], PADRAO)
    assert c == ["90001#teste", "b"]
    assert protecao.candidatas("PC90001 x.xlsx", ["a", "b"], None) == ["a", "b"]
    assert protecao.candidatas("PC90001 x.xlsx", ["a"], "") == ["a"]


def test_nome_sem_codigo_usa_so_as_cadastradas(tmp_path):
    assert protecao.senha_automatica("relatorio.xlsx", PADRAO) is None
    assert protecao.candidatas("relatorio.xlsx", ["a"], PADRAO) == ["a"]
    assert protecao.senha_automatica("PC90001 x.xlsx", "sem marcador") is None


def test_mensagem_quando_nenhuma_abre(tmp_path):
    f = _cifrado(tmp_path, "PC90001 10.2026.xlsx", "outra")
    r = processar_arquivo(f, MAPA, senhas=["x"], padrao_senha=PADRAO)
    assert r.mensagem == (
        "Arquivo protegido por senha: nenhuma senha abriu o arquivo (automatica e cadastradas).")


def test_cadastrada_continua_valendo_depois_da_automatica(tmp_path):
    f = _cifrado(tmp_path, "PC90001 10.2026.xlsx", "cadastrada")
    assert processar_arquivo(f, MAPA, senhas=["cadastrada"], padrao_senha=PADRAO).status == "OK"


def test_senha_automatica_nao_vaza(tmp_path, caplog, monkeypatch):
    import consolidacao.leitura as leitura

    f = _cifrado(tmp_path, "PC90001 10.2026.xlsx", "90001#teste")

    def quebra(*a, **k):
        raise RuntimeError("falhou com 90001#teste e {apolice}#teste")

    monkeypatch.setattr(leitura.openpyxl, "load_workbook", quebra)
    with caplog.at_level(logging.DEBUG):
        c = processar_arquivos([f], MAPA, padrao_senha=PADRAO)
    texto = c.log.to_csv() + "\n".join(r.getMessage() for r in caplog.records)
    assert "90001#teste" not in texto and c.log.iloc[0]["STATUS"] == "ERRO"
    assert "***" in c.log.iloc[0]["MENSAGEM"]


def _cli(tmp_path, senha, *extra):
    from tests.fabrica import escrever_csv_empresas
    from tratamento import main

    entrada = tmp_path / "entrada"
    _cifrado(entrada, "PC123 10.2026.xlsx", senha)
    escrever_csv_empresas(tmp_path / "e.csv", [("ACME", "123")])
    return main(["--entrada", str(entrada), "--empresas", str(tmp_path / "e.csv"),
                 "--saida", str(tmp_path / "r.xlsx"), *extra])


def test_cli_com_padrao_usa_a_automatica(tmp_path, capsys):
    assert _cli(tmp_path, "123#teste", "--padrao-senha", PADRAO) == 0
    out = capsys.readouterr().out
    assert "Arquivos processados com sucesso: 1" in out and "123#teste" not in out


def test_cli_sem_padrao_ou_com_sem_senha_automatica(tmp_path, capsys):
    assert _cli(tmp_path, "123#teste") == 0  # sem padrao: nao ha senha automatica
    assert "Arquivos com erro: 1" in capsys.readouterr().out
    assert _cli(tmp_path / "b", "123#teste", "--padrao-senha", PADRAO,
                "--sem-senha-automatica") == 0
    assert "Arquivos com erro: 1" in capsys.readouterr().out


def test_cli_padrao_senha_invalido(tmp_path):
    with pytest.raises(SystemExit):
        _cli(tmp_path, "x", "--padrao-senha", "sem-marcador")
