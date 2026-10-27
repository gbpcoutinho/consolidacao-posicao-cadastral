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
    assert str(app.ent_senhas_entrada.cget("show")) != ""
    assert str(app.ent_senha_saida.cget("show")) != ""
    app.bt_mostrar_senhas.invoke()
    assert str(app.ent_senhas_entrada.cget("show")) == ""
    assert str(app.ent_senha_saida.cget("show")) == ""
    app.bt_mostrar_senhas.invoke()
    assert str(app.ent_senha_saida.cget("show")) != ""


def test_salvar_guarda_cifrado_e_reabrir_recarrega(raiz, ambiente):
    dados, _, _ = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    app.campo_senhas_entrada.set("Um#1; Dois#2 ;; ")
    app.campo_senha_saida.set(SAIDA)
    app.bt_salvar_senhas.invoke()
    bruto = (dados / "config.json").read_text(encoding="utf-8")
    assert "Um#1" not in bruto and SAIDA not in bruto
    assert json.loads(bruto)["arquivo_empresas"]  # o resto da config segue
    assert ler_senhas(carregar_config(dados)) == (["Um#1", "Dois#2"], SAIDA)

    app2 = _app(raiz, dados, dialogos=Dialogos())
    assert app2.campo_senhas_entrada.get() == "Um#1;Dois#2"
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
    app.campo_senhas_entrada.set("a;b")
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
    app.campo_senhas_entrada.set("outra;chave-de-entrada")
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
