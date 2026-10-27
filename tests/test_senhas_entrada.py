"""Arquivos protegidos por senha: decifrados so em memoria, senhas em ordem."""
import hashlib
import logging

import pytest

from consolidacao.erros import ErroValidacao
from consolidacao.processamento import processar_arquivo, processar_arquivos
from tests.fabrica import DEPENDENTE_OK, SUPLEMENTAR_OK, TITULAR_OK, cifrar_arquivo, criar_arquivo_pc

MAPA = {"123": "ACME", "456": "BETA"}
SEGREDO = "Sen#ha-Secreta-9"
ASSINATURA_CFB = bytes.fromhex("d0cf11e0a1b11ae1")


def _plano(pasta, nome="PC123 10.2026.xlsx"):
    pasta.mkdir(parents=True, exist_ok=True)
    p = pasta / nome
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    return p


def _cifrado(pasta, senha=SEGREDO, nome="PC123 10.2026.xlsx"):
    plano = _plano(pasta / "plano", nome)
    destino = pasta / nome
    cifrar_arquivo(plano, destino, senha)
    assert destino.read_bytes()[:8] == ASSINATURA_CFB
    return destino, plano


def _normalizado(r):
    return r.final.reset_index(drop=True)


def test_senha_certa_le_igual_ao_arquivo_sem_senha(tmp_path):
    cifrado, plano = _cifrado(tmp_path)
    r = processar_arquivo(cifrado, MAPA, senhas=[SEGREDO])
    ref = processar_arquivo(plano, MAPA)
    assert (r.status, r.mensagem) == ("OK", "Processado com sucesso.")
    assert _normalizado(r).equals(_normalizado(ref))


def test_varias_senhas_testadas_em_ordem(tmp_path):
    cifrado, _ = _cifrado(tmp_path, senha="terceira")
    r = processar_arquivo(cifrado, MAPA, senhas=["primeira", "segunda", "terceira"])
    assert r.status == "OK"


def test_senha_errada_vira_erro_do_arquivo(tmp_path):
    cifrado, _ = _cifrado(tmp_path)
    r = processar_arquivo(cifrado, MAPA, senhas=["errada", "outra"])
    assert r.status == "ERRO"
    assert r.mensagem == "Arquivo protegido por senha: nenhuma das senhas cadastradas abriu o arquivo."


def test_sem_senha_cadastrada_vira_erro_do_arquivo(tmp_path):
    cifrado, _ = _cifrado(tmp_path)
    for senhas in ((), [], [""]):
        r = processar_arquivo(cifrado, MAPA, senhas=senhas)
        assert r.status == "ERRO"
        assert r.mensagem == "Arquivo protegido por senha; cadastre a senha na aba Seguranca."


def test_arquivo_sem_senha_nao_muda_com_ou_sem_senhas(tmp_path):
    plano = _plano(tmp_path)
    antes = hashlib.sha256(plano.read_bytes()).hexdigest()
    a = processar_arquivo(plano, MAPA)
    b = processar_arquivo(plano, MAPA, senhas=[SEGREDO])
    assert a.status == b.status == "OK"
    assert _normalizado(a).equals(_normalizado(b))
    assert hashlib.sha256(plano.read_bytes()).hexdigest() == antes


def test_decifrado_so_em_memoria_nada_novo_no_disco(tmp_path):
    cifrado, _ = _cifrado(tmp_path)
    antes = sorted(p.relative_to(tmp_path).as_posix() for p in tmp_path.rglob("*"))
    assert processar_arquivo(cifrado, MAPA, senhas=[SEGREDO]).status == "OK"
    depois = sorted(p.relative_to(tmp_path).as_posix() for p in tmp_path.rglob("*"))
    assert depois == antes


def test_erro_de_um_arquivo_nao_derruba_os_demais(tmp_path):
    ok = _plano(tmp_path / "e", "PC123 10.2026.xlsx")
    cifrado, _ = _cifrado(tmp_path / "e2", nome="PC456 10.2026.xlsx", senha="outra")
    c = processar_arquivos([ok, cifrado], MAPA, senhas=[SEGREDO])
    assert c.log["STATUS"].tolist() == ["OK", "ERRO"]
    assert "nenhuma das senhas" in c.log.iloc[1]["MENSAGEM"]


def test_senhas_chegam_por_pipeline_processar(tmp_path):
    from consolidacao.pipeline import processar

    cifrado, _ = _cifrado(tmp_path)
    assert processar([cifrado], MAPA, senhas=[SEGREDO]).log["STATUS"].tolist() == ["OK"]


def test_senha_nao_vaza_no_log_nem_nas_mensagens(tmp_path, caplog, monkeypatch):
    cifrado, _ = _cifrado(tmp_path, senha="outra")
    with caplog.at_level(logging.DEBUG):
        c = processar_arquivos([cifrado], MAPA, senhas=[SEGREDO])
    texto = c.log.to_csv() + "\n".join(r.getMessage() for r in caplog.records)
    assert SEGREDO not in texto

    # Mesmo que uma biblioteca devolva a senha no texto do erro, ela e ocultada.
    import consolidacao.leitura as leitura

    def quebra(*a, **k):
        raise RuntimeError(f"falhou com {SEGREDO}")

    monkeypatch.setattr(leitura.openpyxl, "load_workbook", quebra)
    c = processar_arquivos([_plano(tmp_path / "x")], MAPA, senhas=[SEGREDO])
    assert c.log.iloc[0]["STATUS"] == "ERRO"
    assert SEGREDO not in c.log.to_csv()
