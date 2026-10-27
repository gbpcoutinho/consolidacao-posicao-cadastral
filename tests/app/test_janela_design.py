"""Plano de design da janela: passos, botao ambar unico, erro inline, atalhos."""
import pytest

pytest.importorskip("tkinter")

from tests.app.test_janela_fumo import (  # noqa: E402,F401
    Dialogos, _app, _consolidado_falso, _rodar_ate, ambiente, raiz,
)


def _estilo(botao):
    return str(botao.cget("style"))


def test_abre_sem_os_recursos_de_logo(raiz, ambiente, tmp_path, monkeypatch):
    from app import janela

    monkeypatch.setattr(janela, "_pasta_recursos", lambda: tmp_path / "nao_existe")
    app = _app(raiz, ambiente[0], dialogos=Dialogos())
    assert app._logo is None
    assert raiz.title() == "Consolidador de Posição Cadastral"


def test_so_um_botao_ambar_processar_depois_salvar(raiz, ambiente):
    dados, _, arquivos = ambiente
    app = _app(raiz, dados, dialogos=Dialogos(),
               processar_fn=lambda c, m, p, k: _consolidado_falso(3))
    assert app.passo == 1
    assert str(app.bt_processar.cget("state")) == "disabled"
    app.adicionar_caminhos(arquivos)
    assert app.passo == 2
    assert _estilo(app.bt_processar) == "Acao.TButton" and _estilo(app.bt_salvar) == "TButton"
    app.processar()
    _rodar_ate(raiz, lambda: app.resultado is not None)
    assert app.passo == 3
    assert _estilo(app.bt_processar) == "TButton" and _estilo(app.bt_salvar) == "Acao.TButton"


def test_enter_dispara_a_acao_ambar_so_na_aba_consolidar(raiz, ambiente):
    dados, _, arquivos = ambiente
    app = _app(raiz, dados, dialogos=Dialogos(),
               processar_fn=lambda c, m, p, k: _consolidado_falso(3))
    app.adicionar_caminhos(arquivos)
    app.abas.select(app.aba_empresas)
    app._atalho_enter()
    assert app.resultado is None and not app.ocupado
    app.abas.select(app.aba_consolidar)
    app._atalho_enter()
    _rodar_ate(raiz, lambda: app.resultado is not None)


def test_erro_do_formulario_fica_embaixo_do_campo(raiz, ambiente):
    dados, _, _ = ambiente
    dialogos = Dialogos()
    app = _app(raiz, dados, dialogos=dialogos)
    assert str(app.bt_incluir.cget("state")) == "disabled"
    assert all(v.get() == "" for v in app.erros_campo.values())  # formulario virgem
    app.campo_empresa.set("X")
    app.campo_cod.set("90001")
    assert "já cadastrado" in app.erros_campo["cod_saude"].get()
    assert str(app.bt_incluir.cget("state")) == "disabled"
    app.campo_cod.set("90077")
    assert app.erros_campo["cod_saude"].get() == ""
    assert str(app.bt_incluir.cget("state")) == "normal"
    app.campo_cod.set("")
    assert app.erros_campo["cod_saude"].get() == ""  # ainda nao passou pelo campo
    app._visitou("cod_saude")
    assert app.erros_campo["cod_saude"].get() == "Informe o código de saúde."
    assert dialogos.avisos == [] and dialogos.erros == []  # nada de pop-up


def test_botoes_de_empresa_seguem_a_selecao(raiz, ambiente):
    dados, _, _ = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    assert str(app.bt_excluir.cget("state")) == "disabled"
    app.tree_empresas.selection_set("90001")
    raiz.update()
    assert str(app.bt_excluir.cget("state")) == "normal"
    assert str(app.bt_editar.cget("state")) == "normal"
    assert str(app.bt_incluir.cget("state")) == "disabled"  # mesmo codigo da selecionada


def test_salvar_senhas_so_habilita_quando_ha_mudanca(raiz, ambiente):
    app = _app(raiz, ambiente[0], dialogos=Dialogos())
    assert str(app.bt_salvar_senhas.cget("state")) == "disabled"
    app.campo_senha_saida.set("Nova#1")
    assert str(app.bt_salvar_senhas.cget("state")) == "normal"
    assert "protegida" in app.texto_status_saida.get()
    app.bt_salvar_senhas.invoke()
    assert str(app.bt_salvar_senhas.cget("state")) == "disabled"
    assert app.texto_senhas.get() == "Senhas salvas."


def test_nova_empresa_limpa_selecao_e_formulario_com_foco(raiz, ambiente, monkeypatch):
    from tkinter import ttk

    dados, _, _ = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    app.tree_empresas.selection_set("90001")
    raiz.update()
    assert app.campo_cod.get() == "90001"
    chamadas = []
    monkeypatch.setattr(ttk.Entry, "focus_set", lambda self: chamadas.append(self))
    app.bt_nova.invoke()
    assert app.tree_empresas.selection() == ()
    assert (app.campo_empresa.get(), app.campo_cod.get(), app.campo_dental.get()) == ("", "", "")
    assert chamadas == [app.entradas["empresa"]]
    app.campo_empresa.set("OUTRA")
    app.campo_cod.set("90001")  # codigo que antes era da selecionada: agora e duplicata
    assert str(app.bt_incluir.cget("state")) == "disabled"
    app.campo_cod.set("90555")
    assert str(app.bt_incluir.cget("state")) == "normal"


def test_esc_limpa_a_empresa_so_na_aba_empresas(raiz, ambiente):
    dados, _, _ = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    app.abas.select(app.aba_empresas)
    app.tree_empresas.selection_set("90002")
    raiz.update()
    app._atalho_esc()
    assert app.tree_empresas.selection() == () and app.campo_cod.get() == ""
    app.campo_empresa.set("X")
    app.abas.select(app.aba_consolidar)
    app._atalho_esc()
    assert app.campo_empresa.get() == "X"
