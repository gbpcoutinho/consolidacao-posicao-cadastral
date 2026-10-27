"""Arrastar e soltar: parse dos caminhos do tkdnd, mesmas regras da selecao e fallback."""
import shutil

import pytest

pytest.importorskip("tkinter")

from tests.app.test_janela_fumo import Dialogos, _app, ambiente, raiz  # noqa: E402,F401


def test_parse_dos_caminhos_com_espaco_acento_e_chaves(raiz):
    from pathlib import Path

    from app.dnd import caminhos_soltos

    dados = "{C:/Pasta com espaço/PC90001 10.2026.xlsx} C:/x/PC90002.xlsx {C:/ação {x}/ç.xlsx}"
    assert caminhos_soltos(raiz, dados) == [
        Path("C:/Pasta com espaço/PC90001 10.2026.xlsx"), Path("C:/x/PC90002.xlsx"),
        Path("C:/ação {x}/ç.xlsx")]


def test_soltar_arquivos_e_pasta_segue_as_regras_da_selecao(raiz, ambiente, tmp_path):
    dados, _, arquivos = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    pasta = tmp_path / "solta"
    pasta.mkdir()
    shutil.copy(arquivos[1], pasta / "PC90002 10.2026.xlsx")
    for lixo in ("~$PC90001 10.2026.xlsx", "PC_CONSOLIDADO_10_2026.xlsx", "notas.txt"):
        (pasta / lixo).write_bytes(b"x")
    app.adicionar_soltos([arquivos[0], arquivos[0], pasta, tmp_path / "notas.txt"])
    nomes = [app.tree_arquivos.item(i, "values")[0] for i in app.tree_arquivos.get_children()]
    assert nomes == ["PC90001 10.2026.xlsx", "PC90002 10.2026.xlsx"]
    app.adicionar_soltos([pasta])  # de novo: nada duplica
    assert len(app.itens) == 2


def test_alvos_registrados_com_dnd_ativo(raiz, ambiente):
    app = _app(raiz, ambiente[0], dialogos=Dialogos())
    assert app.dnd_ativo
    assert "Arraste" in app.vazio_arquivos.cget("text")


def test_sem_tkdnd_o_app_abre_normal(raiz, ambiente, monkeypatch):
    from app import janela

    monkeypatch.setattr(janela.dnd, "habilitar", lambda r: False)
    app = _app(raiz, ambiente[0], dialogos=Dialogos())
    assert not app.dnd_ativo
    assert "Arraste" not in app.vazio_arquivos.cget("text")
    app.adicionar_caminhos(ambiente[2])  # os botoes seguem funcionando
    assert len(app.itens) == 3


def test_habilitar_devolve_false_quando_a_biblioteca_falha(raiz, monkeypatch):
    import sys

    from app import dnd

    monkeypatch.setitem(sys.modules, "tkinterdnd2", None)  # import falha
    assert dnd.habilitar(raiz) is False


def test_textos_do_bradesco_na_lista_vazia_e_no_erro_de_estrutura(raiz, ambiente):
    app = _app(raiz, ambiente[0], dialogos=Dialogos())
    assert "posição cadastral do Bradesco (.xlsx)" in app.vazio_arquivos.cget("text")
    msg = "Esperada exatamente uma aba de TITULAR; encontrada(s) 0."
    assert app._com_dica_de_estrutura(msg).endswith(
        "Confira se é uma planilha de posição cadastral do Bradesco Saúde.")
    assert app._com_dica_de_estrutura("Outro erro.") == "Outro erro."
