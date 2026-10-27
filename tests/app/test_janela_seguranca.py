"""Aba Seguranca da janela: senhas de entrada e de saida."""
import json

import msoffcrypto  # noqa: F401
import pytest

pytest.importorskip("tkinter")

from app.config import carregar_config, ler_senhas  # noqa: E402
from tests.app.test_janela_fumo import (  # noqa: E402,F401
    Dialogos, _app, _consolidado_falso, _rodar_ate, ambiente, raiz,
)
from tests.fabrica import cifrar_arquivo  # noqa: E402

SAIDA = "Saida#Forte-5"


def _tabs(app):
    return [app.abas.tab(i, "text") for i in range(app.abas.index("end"))]


def test_aba_seguranca_existe_com_campos_mascarados_e_mostrar(raiz, ambiente):
    dados, _, _ = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    assert "Segurança" in _tabs(app)
    app.definir_senhas_entrada(["a", "b"])
    campos = app.ents_senhas_entrada + [app.ent_senha_saida]
    assert len(campos) == 3 and all(str(e.cget("show")) != "" for e in campos)
    app.bt_mostrar_senhas.invoke()
    assert all(str(e.cget("show")) == "" for e in campos)
    app.adicionar_linha_senha()  # linha nova ja nasce no modo atual (visivel)
    assert str(app.ents_senhas_entrada[-1].cget("show")) == ""
    app.bt_mostrar_senhas.invoke()
    assert all(str(e.cget("show")) != "" for e in app.ents_senhas_entrada + [app.ent_senha_saida])


def test_salvar_guarda_cifrado_e_reabrir_recarrega(raiz, ambiente):
    dados, _, _ = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    app.definir_senhas_entrada(["Um#1", " Dois#2 ", "", "  "])
    app.campo_senha_saida.set(SAIDA)
    app.bt_salvar_senhas.invoke()
    bruto = (dados / "config.json").read_text(encoding="utf-8")
    assert "Um#1" not in bruto and SAIDA not in bruto
    assert json.loads(bruto)["arquivo_empresas"]  # o resto da config segue
    assert ler_senhas(carregar_config(dados)) == (["Um#1", "Dois#2"], SAIDA)

    app2 = _app(raiz, dados, dialogos=Dialogos())
    assert [v.get() for v in app2.vars_senhas_entrada] == ["Um#1", "Dois#2"]
    assert app2.campo_senha_saida.get() == SAIDA


def test_processar_e_salvar_usam_as_senhas(raiz, ambiente, tmp_path):
    dados, _, arquivos = ambiente
    vistos = {}

    def processar_falso(caminhos, mapa, progresso, cancelar, senhas=()):
        vistos["senhas"] = list(senhas)
        return _consolidado_falso(3)

    def gravar_falso(resultado, destino, caminhos, arquivo_empresas=None, progresso=None,
                     senha_saida=None):
        vistos["senha_saida"] = senha_saida
        destino.write_bytes(b"x")

    destino = tmp_path / "s" / "c.xlsx"
    destino.parent.mkdir()
    app = _app(raiz, dados, dialogos=Dialogos(salvar=destino),
               processar_fn=processar_falso, gravar_fn=gravar_falso)
    app.definir_senhas_entrada(["a", "b"])
    app.campo_senha_saida.set(SAIDA)
    app.adicionar_caminhos(arquivos)
    app.processar()
    _rodar_ate(raiz, lambda: app.resultado is not None)
    assert vistos["senhas"] == ["a", "b"]
    assert "protegida por senha" in app.resumo_origem.get().casefold()
    app.salvar_como()
    _rodar_ate(raiz, lambda: app.ultimo_salvo is not None)
    assert vistos["senha_saida"] == SAIDA
    assert "protegid" in app.texto_salvo.get().casefold()
    assert SAIDA not in app.texto_salvo.get() + app.resumo_origem.get()


def test_sem_senhas_resumo_nao_fala_de_protecao(raiz, ambiente):
    dados, _, arquivos = ambiente
    app = _app(raiz, dados, dialogos=Dialogos(),
               processar_fn=lambda c, m, p, k: _consolidado_falso(3))
    app.adicionar_caminhos(arquivos)
    app.processar()
    _rodar_ate(raiz, lambda: app.resultado is not None)
    assert "senha" not in app.resumo_origem.get().casefold()


def test_fluxo_real_entrada_e_saida_protegidas(raiz, ambiente, tmp_path):
    import io

    import openpyxl

    dados, _, arquivos = ambiente
    pasta = tmp_path / "prot"
    pasta.mkdir()
    final = []
    for a in arquivos:
        cifrar_arquivo(a, pasta / a.name, "chave-de-entrada")
        final.append(pasta / a.name)
    destino = tmp_path / "saida" / "c.xlsx"
    app = _app(raiz, dados, dialogos=Dialogos(salvar=destino))
    app.definir_senhas_entrada(["outra", "chave-de-entrada"])
    app.campo_senha_saida.set(SAIDA)
    app.adicionar_caminhos(final)
    app.processar()
    _rodar_ate(raiz, lambda: app.resultado is not None)
    assert app.resumo_topo.get() == "3 OK, 0 com erro"
    app.salvar_como()
    _rodar_ate(raiz, lambda: app.ultimo_salvo is not None)
    f = msoffcrypto.OfficeFile(open(destino, "rb"))
    f.load_key(password=SAIDA)
    saida = io.BytesIO()
    f.decrypt(saida)
    assert "Base Tratada" in openpyxl.load_workbook(saida, read_only=True).sheetnames


def test_cada_senha_tem_sua_linha_e_pode_ter_ponto_e_virgula(raiz, ambiente):
    dados, _, _ = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    assert len(app.vars_senhas_entrada) == 1  # sempre ha uma linha para digitar
    app.definir_senhas_entrada(["a;b", "c"])
    assert app._senhas_entrada() == ["a;b", "c"]
    app.bt_salvar_senhas.invoke()
    assert ler_senhas(carregar_config(dados))[0] == ["a;b", "c"]


def test_adicionar_e_remover_linha_de_senha(raiz, ambiente):
    dados, _, _ = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    app.definir_senhas_entrada(["a"])
    foco = []
    app.bt_adicionar_senha.invoke()
    assert len(app.vars_senhas_entrada) == 2 and app._senhas_entrada() == ["a"]
    app.vars_senhas_entrada[1].set("b")
    assert str(app.bt_salvar_senhas.cget("state")) == "normal"
    app.remover_linha_senha(0)
    assert app._senhas_entrada() == ["b"] and len(app.vars_senhas_entrada) == 1
    app.remover_linha_senha(0)  # a ultima linha so e esvaziada
    assert len(app.vars_senhas_entrada) == 1 and app._senhas_entrada() == []
    assert len(app.bts_remover_senha) == 1


def test_nova_linha_recebe_o_foco(raiz, ambiente, monkeypatch):
    from tkinter import ttk

    chamadas = []
    monkeypatch.setattr(ttk.Entry, "focus_set", lambda self: chamadas.append(self))
    app = _app(raiz, ambiente[0], dialogos=Dialogos())
    app.adicionar_linha_senha(foco=True)
    assert chamadas == [app.ents_senhas_entrada[-1]]


def test_muitas_senhas_nao_empurram_a_senha_de_saida_para_fora(raiz, ambiente):
    dados, _, _ = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    raiz.deiconify()
    raiz.geometry("900x600+0+0")
    app.abas.select(app.aba_seguranca)
    app.definir_senhas_entrada([f"s{i}" for i in range(15)])
    raiz.update()
    app.adicionar_linha_senha(foco=True)
    import time

    for _ in range(10):
        raiz.update()
        time.sleep(0.02)
    altura = raiz.winfo_height()
    for w in (app.ent_senha_saida, app.bt_salvar_senhas, app.bt_mostrar_senhas):
        assert w.winfo_ismapped()
        assert w.winfo_rooty() - raiz.winfo_rooty() + w.winfo_height() <= altura
    # A lista rola: a ultima linha (nova) fica visivel dentro da area rolavel.
    canvas = app.canvas_senhas
    linha = app.ents_senhas_entrada[-1]
    y = linha.winfo_rooty() - canvas.winfo_rooty()
    assert 0 <= y and y + linha.winfo_height() <= canvas.winfo_height() + 2
    assert canvas.winfo_height() >= 100  # ~3 linhas no minimo

def test_senha_automatica_vem_ligada_mas_sem_padrao(raiz, ambiente):
    app = _app(raiz, ambiente[0], dialogos=Dialogos())
    assert app.var_senha_automatica.get() is True
    assert app.campo_padrao_senha.get() == ""  # nao ha padrao de fabrica
    assert str(app.ent_padrao_senha.cget("show")) != ""  # mascarado
    assert "Preencha o padrão para ativar" in app.erro_padrao_senha.get()
    assert str(app.bt_salvar_senhas.cget("state")) == "disabled"
    assert app._padrao_para_processar() is None  # vazio = desligada, sem erro


def test_padrao_sem_marcador_e_erro_inline_e_nao_salva(raiz, ambiente):
    app = _app(raiz, ambiente[0], dialogos=Dialogos())
    app.campo_padrao_senha.set("sem-marcador")
    assert "{apolice}" in app.erro_padrao_senha.get()
    assert str(app.bt_salvar_senhas.cget("state")) == "disabled"
    app.campo_padrao_senha.set("X-{apolice}")
    assert app.erro_padrao_senha.get() == ""
    assert str(app.bt_salvar_senhas.cget("state")) == "normal"
    app.campo_padrao_senha.set("")  # vazio nao e erro
    app.campo_senha_saida.set("x")
    assert str(app.bt_salvar_senhas.cget("state")) == "normal"


def test_padrao_e_ativacao_persistem_e_o_padrao_vai_cifrado(raiz, ambiente):
    from app.config import ler_senha_automatica

    dados, _, _ = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    app.campo_padrao_senha.set("Seg-{apolice}-Xz")
    app.chk_senha_automatica.invoke()  # desliga
    assert app.var_senha_automatica.get() is False
    app.bt_salvar_senhas.invoke()
    bruto = (dados / "config.json").read_text(encoding="utf-8")
    assert "Seg-" not in bruto
    assert ler_senha_automatica(carregar_config(dados)) == (False, "Seg-{apolice}-Xz")
    app2 = _app(raiz, dados, dialogos=Dialogos())
    assert app2.var_senha_automatica.get() is False
    assert app2.campo_padrao_senha.get() == "Seg-{apolice}-Xz"
    assert str(app2.ent_padrao_senha.cget("state")) == "disabled"


def test_processar_repassa_o_padrao_so_quando_ha_padrao_ativo(raiz, ambiente):
    dados, _, arquivos = ambiente
    vistos = []

    def falso(caminhos, mapa, progresso, cancelar, **extras):
        vistos.append(extras)
        return _consolidado_falso(3)

    app = _app(raiz, dados, dialogos=Dialogos(), processar_fn=falso)
    app.adicionar_caminhos(arquivos)
    passos = (lambda: None,                                   # sem padrao
              lambda: app.campo_padrao_senha.set("Z-{apolice}"),  # com padrao
              lambda: app.chk_senha_automatica.invoke())      # desligada
    for mexer in passos:
        mexer()
        app.resultado = None
        app.processar()
        _rodar_ate(raiz, lambda: app.resultado is not None)
    assert vistos == [{}, {"padrao_senha": "Z-{apolice}"}, {}]


def test_fluxo_real_abre_arquivos_so_com_a_senha_automatica(raiz, ambiente, tmp_path):
    dados, _, arquivos = ambiente
    pasta = tmp_path / "prot"
    pasta.mkdir()
    final = []
    for a in arquivos:
        cod = a.name[2:7]  # PC90001 ... -> 90001
        cifrar_arquivo(a, pasta / a.name, f"{cod}#teste")
        final.append(pasta / a.name)
    app = _app(raiz, dados, dialogos=Dialogos())
    app.campo_padrao_senha.set("{apolice}#teste")
    app.adicionar_caminhos(final)
    app.processar()
    _rodar_ate(raiz, lambda: app.resultado is not None)
    assert app.resumo_topo.get() == "3 OK, 0 com erro"


def _recursos_com_padrao(tmp_path, monkeypatch, texto="{apolice}#teste\n"):
    from app import janela

    pasta = tmp_path / "recursos"
    pasta.mkdir()
    (pasta / "padrao_senha.txt").write_text(texto, encoding="utf-8")
    monkeypatch.setattr(janela, "_pasta_recursos", lambda: pasta)


def test_padrao_de_fabrica_do_arquivo_local_ja_vem_preenchido(raiz, ambiente, tmp_path,
                                                              monkeypatch):
    dados, _, arquivos = ambiente
    _recursos_com_padrao(tmp_path, monkeypatch)
    vistos = []

    def falso(caminhos, mapa, progresso, cancelar, **extras):
        vistos.append(extras)
        return _consolidado_falso(3)

    app = _app(raiz, dados, dialogos=Dialogos(), processar_fn=falso)
    assert app.var_senha_automatica.get() is True
    assert app.campo_padrao_senha.get() == "{apolice}#teste"
    assert app.erro_padrao_senha.get() == ""
    assert str(app.bt_salvar_senhas.cget("state")) == "disabled"  # nada a salvar ainda
    app.adicionar_caminhos(arquivos)
    app.processar()
    _rodar_ate(raiz, lambda: app.resultado is not None)
    assert vistos == [{"padrao_senha": "{apolice}#teste"}]


def test_padrao_salvo_vale_mais_que_o_de_fabrica_e_o_de_fabrica_nao_vai_para_a_config(
        raiz, ambiente, tmp_path, monkeypatch):
    dados, _, _ = ambiente
    _recursos_com_padrao(tmp_path, monkeypatch)
    app = _app(raiz, dados, dialogos=Dialogos())
    app.campo_senha_saida.set("x")
    app.bt_salvar_senhas.invoke()
    assert "padrao_senha_automatica" not in carregar_config(dados)  # igual ao de fabrica
    app.campo_padrao_senha.set("Meu-{apolice}")
    app.bt_salvar_senhas.invoke()
    app2 = _app(raiz, dados, dialogos=Dialogos())
    assert app2.campo_padrao_senha.get() == "Meu-{apolice}"


def test_arquivo_de_padrao_invalido_ou_ausente_deixa_vazio(raiz, ambiente, tmp_path,
                                                          monkeypatch):
    dados, _, _ = ambiente
    _recursos_com_padrao(tmp_path, monkeypatch, texto="sem-marcador\n")
    assert _app(raiz, dados, dialogos=Dialogos()).campo_padrao_senha.get() == ""
