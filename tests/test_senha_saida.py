"""Saida protegida por senha: nasce cifrada, o destino nunca tem versao sem senha."""
import io
import logging
import os

import msoffcrypto
import openpyxl
import pytest

from consolidacao import exportacao
from consolidacao.pipeline import executar, gravar, processar
from tests.fabrica import DEPENDENTE_OK, SUPLEMENTAR_OK, TITULAR_OK, criar_arquivo_pc

MAPA = {"123": "ACME"}
SENHA = "Saida#Segura-7"
CFB = bytes.fromhex("d0cf11e0a1b11ae1")


@pytest.fixture
def entrada(tmp_path):
    p = tmp_path / "entrada" / "PC123 10.2026.xlsx"
    p.parent.mkdir()
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    return p


def _conteudo(origem):
    wb = openpyxl.load_workbook(origem, read_only=True)
    try:
        return {ws.title: [list(l) for l in ws.iter_rows(values_only=True)] for ws in wb}
    finally:
        wb.close()


def _decifrar(caminho, senha):
    f = msoffcrypto.OfficeFile(open(caminho, "rb"))
    f.load_key(password=senha)
    saida = io.BytesIO()
    f.decrypt(saida)
    saida.seek(0)
    return saida


def test_saida_cifrada_decifra_com_a_senha_e_tem_o_mesmo_conteudo(tmp_path, entrada):
    c = processar([entrada], MAPA)
    sem, com = tmp_path / "sem.xlsx", tmp_path / "com.xlsx"
    gravar(c, sem, [entrada])
    gravar(c, com, [entrada], senha_saida=SENHA)
    assert sem.read_bytes()[:2] == b"PK"
    assert com.read_bytes()[:8] == CFB
    assert _conteudo(_decifrar(com, SENHA)) == _conteudo(sem)
    with pytest.raises(Exception):
        _decifrar(com, "errada")


@pytest.mark.parametrize("senha", [None, ""])
def test_senha_vazia_ou_none_grava_sem_senha(tmp_path, entrada, senha):
    destino = tmp_path / "s.xlsx"
    gravar(processar([entrada], MAPA), destino, [entrada], senha_saida=senha)
    assert destino.read_bytes()[:2] == b"PK"


def test_nenhum_arquivo_sem_senha_vai_ao_disco(tmp_path, entrada, monkeypatch):
    saida = tmp_path / "saida"
    saida.mkdir()
    destino = saida / "consolidado.xlsx"
    c = processar([entrada], MAPA)
    visto = {}
    original = os.replace

    def espiao(origem, alvo):
        # No instante da troca, o temporario ~$ ja e o arquivo cifrado.
        visto["cabecalho"] = open(origem, "rb").read(8)
        visto["alvo_existia"] = os.path.exists(alvo)
        return original(origem, alvo)

    monkeypatch.setattr(exportacao.os, "replace", espiao)
    gravar(c, destino, [entrada], senha_saida=SENHA)
    assert visto == {"cabecalho": CFB, "alvo_existia": False}
    assert [p.name for p in saida.iterdir()] == ["consolidado.xlsx"]
    assert destino.read_bytes()[:8] == CFB


def test_falha_na_cifra_nao_deixa_destino_nem_temporario(tmp_path, entrada, monkeypatch):
    destino = tmp_path / "saida" / "c.xlsx"
    c = processar([entrada], MAPA)

    def quebra(*a, **k):
        raise OSError("disco cheio")

    monkeypatch.setattr(exportacao, "cifrar_para", quebra)
    with pytest.raises(Exception):
        gravar(c, destino, [entrada], senha_saida=SENHA)
    assert list(destino.parent.iterdir()) == []


def test_destino_antigo_fica_intacto_se_a_cifra_falhar(tmp_path, entrada, monkeypatch):
    destino = tmp_path / "c.xlsx"
    destino.write_bytes(b"antigo")
    monkeypatch.setattr(exportacao, "cifrar_para", lambda *a, **k: (_ for _ in ()).throw(OSError("x")))
    with pytest.raises(Exception):
        gravar(processar([entrada], MAPA), destino, [entrada], senha_saida=SENHA)
    assert destino.read_bytes() == b"antigo"
    assert [p.name for p in tmp_path.iterdir() if p.is_file()] == ["c.xlsx"]


def test_senha_da_saida_nao_vaza_em_erros_nem_no_log(tmp_path, entrada, monkeypatch, caplog):
    c = processar([entrada], MAPA)

    def quebra(*a, **k):
        raise RuntimeError(f"falhou com {SENHA}")

    monkeypatch.setattr(exportacao, "cifrar_para", quebra)
    with caplog.at_level(logging.DEBUG):
        with pytest.raises(Exception) as e:
            gravar(c, tmp_path / "c.xlsx", [entrada], senha_saida=SENHA)
    assert SENHA not in str(e.value)
    assert SENHA not in "\n".join(r.getMessage() for r in caplog.records)


def test_executar_com_senhas_de_entrada_e_saida(tmp_path, entrada):
    destino = tmp_path / "out.xlsx"
    r = executar([entrada], MAPA, destino, senhas=["x"], senha_saida=SENHA)
    assert r.log["STATUS"].tolist() == ["OK"]
    assert "Base Tratada" in _conteudo(_decifrar(destino, SENHA))
    assert SENHA not in str(_conteudo(_decifrar(destino, SENHA))["Log"])
