"""Fumo da janela: Tk de verdade, dialogos falsos, sem interacao manual."""
import shutil
import threading
import time
from pathlib import Path

import pandas as pd
import pytest

tk = pytest.importorskip("tkinter")

from app import empresas_compartilhadas as ec  # noqa: E402
from app.config import salvar_config  # noqa: E402
from consolidacao.processamento import COLUNAS_BASE_VAZIA, COLUNAS_LOG, COLUNAS_PENDENCIAS, Consolidado  # noqa: E402

EXEMPLOS = Path(__file__).resolve().parents[2] / "exemplos"


class Dialogos:
    def __init__(self, respostas=None, salvar=None):
        self.respostas = list(respostas or [])
        self.perguntas, self.erros, self.avisos, self.infos = [], [], [], []
        self.salvar = salvar
        self.abertos = []

    def perguntar(self, titulo, mensagem):
        self.perguntas.append(mensagem)
        return self.respostas.pop(0) if self.respostas else True

    def erro(self, titulo, mensagem):
        self.erros.append(mensagem)

    def avisar(self, titulo, mensagem):
        self.avisos.append(mensagem)

    def informar(self, titulo, mensagem):
        self.infos.append(mensagem)

    def primeiro_uso(self):
        return None

    def salvar_como(self, nome, pasta):
        self.nome_sugerido = nome
        return self.salvar

    def abrir_arquivo(self, caminho):
        self.abertos.append(("arquivo", caminho))

    def abrir_pasta(self, caminho):
        self.abertos.append(("pasta", caminho))


@pytest.fixture
def raiz():
    try:
        r = tk.Tk()
    except tk.TclError as e:
        pytest.skip(f"sem display para o Tk: {e}")
    r.withdraw()
    yield r
    try:
        r.destroy()
    except tk.TclError:
        pass


@pytest.fixture
def ambiente(tmp_path):
    """Pasta de dados, 'rede' com a relacao e copia dos exemplos ficticios."""
    dados = tmp_path / "appdata"
    dados.mkdir()
    rede = tmp_path / "rede" / "empresas.csv"
    rede.parent.mkdir()
    ec.criar_a_partir_de_csv(EXEMPLOS / "COD_EMPRESAS_EXEMPLO.csv", rede)
    salvar_config(dados, {"arquivo_empresas": str(rede)})
    entrada = tmp_path / "entrada"
    shutil.copytree(EXEMPLOS / "entrada", entrada)
    arquivos = sorted(entrada.glob("PC9000*.xlsx"))
    return dados, rede, arquivos


def _consolidado_falso(n_ok=3):
    log = pd.DataFrame(
        [{**{c: None for c in COLUNAS_LOG}, "ARQUIVO": f"PC9000{i} 10.2026.xlsx",
          "STATUS": "OK", "DATAS_INVALIDAS": 0, "QTD_DIVERGENCIAS_DATAS": 0}
         for i in range(1, n_ok + 1)],
        columns=COLUNAS_LOG,
    )
    return Consolidado(
        consolidado=pd.DataFrame(columns=COLUNAS_BASE_VAZIA, dtype=object),
        pendencias=pd.DataFrame(columns=COLUNAS_PENDENCIAS, dtype=object),
        log=log,
    )


def _rodar_ate(raiz, condicao, limite=20.0):
    fim = time.monotonic() + limite
    while time.monotonic() < fim:
        raiz.update()
        if condicao():
            return
        time.sleep(0.02)
    raise AssertionError("condicao nao atingida")


def _app(raiz, dados, **kw):
    from app.janela import App

    kw.setdefault("primeiro_uso_automatico", False)
    return App(raiz, dados, **kw)


def test_janela_abre_e_fecha(raiz, tmp_path):
    app = _app(raiz, tmp_path, dialogos=Dialogos(), primeiro_uso_automatico=True)
    raiz.update()
    # Sem relacao configurada e sem cache: Processar bloqueado.
    assert app.relacao is None
    assert str(app.bt_processar.cget("state")) == "disabled"
    assert app.bloqueio.winfo_manager() == "pack"


def test_fluxo_sem_dialogos(raiz, ambiente):
    dados, rede, arquivos = ambiente
    chamadas = []

    def processar_falso(caminhos, mapa, progresso, cancelar):
        chamadas.append((list(caminhos), dict(mapa)))
        for i, c in enumerate(caminhos, start=1):
            progresso(i, len(caminhos), c.name)
        return _consolidado_falso(3)

    app = _app(raiz, dados, dialogos=Dialogos(), processar_fn=processar_falso)
    app.adicionar_caminhos(arquivos)
    assert [app.tree_arquivos.item(i, "values")[1] for i in app.tree_arquivos.get_children()] == [
        "ALFA LOGISTICA FICTICIA LTDA", "BETA SERVICOS EXEMPLO S A", "GAMA INDUSTRIA DEMONSTRACAO",
    ]
    app.processar()
    _rodar_ate(raiz, lambda: app.resumo_topo.get() != "")
    assert app.resumo_topo.get() == "3 OK, 0 com erro"
    assert app.resumo_origem.get() == "Relação de empresas: rede"
    assert chamadas[0][0] == arquivos and chamadas[0][1]["90001"] == "ALFA LOGISTICA FICTICIA LTDA"
    assert str(app.bt_salvar.cget("state")) == "normal"


def test_processar_salvar_e_abrir_de_verdade(raiz, ambiente, tmp_path):
    dados, rede, arquivos = ambiente
    destino = tmp_path / "saida" / "consolidado.xlsx"
    # Respostas: 1a pergunta = alertas antes de processar? (nao ha) / salvar com pendencias: sim.
    dialogos = Dialogos(salvar=destino)
    app = _app(raiz, dados, dialogos=dialogos)
    app.adicionar_caminhos(arquivos)
    app.processar()
    _rodar_ate(raiz, lambda: app.resultado is not None)
    assert app.resumo_topo.get() == "3 OK, 0 com erro"
    app.salvar_como()
    _rodar_ate(raiz, lambda: app.ultimo_salvo is not None)
    assert destino.exists() and app.salvo
    assert dialogos.nome_sugerido == "PC_CONSOLIDADO_10_2026.xlsx"
    assert any("pendência" in p for p in dialogos.perguntas)  # exemplos tem 1 pendencia
    app.abrir_arquivo()
    app.abrir_pasta()
    assert dialogos.abertos == [("arquivo", destino), ("pasta", destino)]


def test_salvar_avisa_entrada_alterada_e_reprocessa(raiz, ambiente, tmp_path):
    dados, rede, arquivos = ambiente
    dialogos = Dialogos(respostas=[True], salvar=tmp_path / "x.xlsx")
    app = _app(raiz, dados, dialogos=dialogos,
               processar_fn=lambda c, m, p, k: _consolidado_falso(3))
    app.adicionar_caminhos(arquivos)
    app.processar()
    _rodar_ate(raiz, lambda: app.resultado is not None)
    arquivos[0].write_bytes(arquivos[0].read_bytes() + b"\0")  # alguem trocou a planilha
    app.salvar_como()
    assert "mudou desde o processamento" in dialogos.perguntas[0]
    assert app.tipo_trabalho == "processar"  # respondeu sim: reprocessa em vez de salvar
    _rodar_ate(raiz, lambda: app.resultado is not None)
    assert not (tmp_path / "x.xlsx").exists()


def test_erro_ao_salvar_mantem_resultado(raiz, ambiente, tmp_path):
    dados, rede, arquivos = ambiente
    dialogos = Dialogos(salvar=arquivos[0])  # saida = entrada: o pipeline recusa
    app = _app(raiz, dados, dialogos=dialogos,
               processar_fn=lambda c, m, p, k: _consolidado_falso(3))
    app.adicionar_caminhos(arquivos)
    app.processar()
    _rodar_ate(raiz, lambda: app.resultado is not None)
    app.salvar_como()
    _rodar_ate(raiz, lambda: dialogos.erros)
    assert "arquivos de entrada" in dialogos.erros[0]
    assert app.resultado is not None and str(app.bt_salvar.cget("state")) == "normal"


def test_cancelar_no_meio(raiz, ambiente):
    dados, rede, arquivos = ambiente
    comecou = threading.Event()

    def lento(caminhos, mapa, progresso, cancelar):
        from consolidacao.erros import ProcessamentoCancelado

        comecou.set()
        while not cancelar():
            time.sleep(0.01)
        raise ProcessamentoCancelado("cancelado")

    app = _app(raiz, dados, dialogos=Dialogos(), processar_fn=lento)
    app.adicionar_caminhos(arquivos)
    app.processar()
    assert comecou.wait(5)
    app.cancelar()
    _rodar_ate(raiz, lambda: app.tipo_trabalho is None)
    assert app.texto_progresso.get() == "Processamento cancelado."
    assert app.resultado is None


def test_fechar_durante_processamento_pergunta_e_cancela(raiz, ambiente):
    dados, rede, arquivos = ambiente
    cancelado = threading.Event()

    def lento(caminhos, mapa, progresso, cancelar):
        from consolidacao.erros import ProcessamentoCancelado

        while not cancelar():
            time.sleep(0.01)
        cancelado.set()
        raise ProcessamentoCancelado("cancelado")

    dialogos = Dialogos(respostas=[False, True])
    app = _app(raiz, dados, dialogos=dialogos, processar_fn=lento)
    app.adicionar_caminhos(arquivos)
    app.processar()
    app.fechar()  # respondeu nao: continua aberto
    assert raiz.winfo_exists() and app.tipo_trabalho == "processar"
    app.fechar()  # respondeu sim: cancela e fecha
    assert cancelado.wait(5)
    assert len(dialogos.perguntas) == 2 and "andamento" in dialogos.perguntas[0]


def test_fechar_com_resultado_nao_salvo_pergunta(raiz, ambiente):
    dados, rede, arquivos = ambiente
    dialogos = Dialogos(respostas=[False])
    app = _app(raiz, dados, dialogos=dialogos,
               processar_fn=lambda c, m, p, k: _consolidado_falso(3))
    app.adicionar_caminhos(arquivos)
    app.processar()
    _rodar_ate(raiz, lambda: app.resultado is not None)
    app.fechar()
    assert "não foi salvo" in dialogos.perguntas[-1]
    assert raiz.winfo_exists()


def test_incluir_empresa_grava_na_rede_e_atualiza_lista(raiz, ambiente, tmp_path):
    dados, rede, arquivos = ambiente
    novo = tmp_path / "entrada" / "PC90009 10.2026.xlsx"
    shutil.copy(arquivos[0], novo)
    app = _app(raiz, dados, dialogos=Dialogos())
    app.adicionar_caminhos([novo])
    assert "erro" in app.tree_arquivos.item("0", "tags")
    app.campo_empresa.set("  DELTA FICTICIA  ")
    app.campo_cod.set("90009")
    app.incluir_empresa()
    from consolidacao.empresas import carregar_mapa_empresas

    assert carregar_mapa_empresas(rede)["90009"] == "DELTA FICTICIA"
    assert app.tree_arquivos.item("0", "values")[1] == "DELTA FICTICIA"
    assert "erro" not in app.tree_arquivos.item("0", "tags")
    assert app.bt_reprocessar.winfo_manager() == "pack"


def test_validacao_ao_digitar_cod_saude(raiz, ambiente):
    dados, rede, arquivos = ambiente
    app = _app(raiz, dados, dialogos=Dialogos())
    app.campo_empresa.set("X")
    app.campo_cod.set("90001")
    assert "já cadastrado" in app.texto_validacao.get()
    app.campo_cod.set("")
    assert "COD_SAUDE" in app.texto_validacao.get()
    app.campo_cod.set("123")
    assert app.texto_validacao.get() == ""


def test_excluir_pede_confirmacao_com_nome(raiz, ambiente):
    dados, rede, arquivos = ambiente
    dialogos = Dialogos(respostas=[False])
    app = _app(raiz, dados, dialogos=dialogos)
    app.tree_empresas.selection_set("0")
    raiz.update()
    app.excluir_empresa()
    assert "ALFA LOGISTICA FICTICIA LTDA" in dialogos.perguntas[0]
    from consolidacao.empresas import carregar_mapa_empresas

    assert "90001" in carregar_mapa_empresas(rede)


def test_sem_rede_usa_cache_mostra_faixa_e_bloqueia_edicao(raiz, ambiente):
    dados, rede, arquivos = ambiente
    ec.ler(rede, dados)  # deixa a copia local
    rede.parent.rename(rede.parent.with_name("rede_fora"))
    app = _app(raiz, dados, dialogos=Dialogos())
    raiz.update()
    assert app.relacao.origem == "cache"
    assert app.faixa_consolidar.winfo_manager() == "pack"
    assert "cópia local" in app.faixa_consolidar.cget("text")
    assert str(app.bt_incluir.cget("state")) == "disabled"
    app.adicionar_caminhos(arquivos)
    assert str(app.bt_processar.cget("state")) == "normal"


def test_salvar_em_cima_de_entrada_removida_da_lista_e_recusado(raiz, ambiente):
    dados, rede, arquivos = ambiente
    original = arquivos[0].read_bytes()
    # 1a pergunta: "lista mudou; reprocessar?" -> nao. Depois: sim.
    dialogos = Dialogos(respostas=[False], salvar=arquivos[0])
    app = _app(raiz, dados, dialogos=dialogos,
               processar_fn=lambda c, m, p, k: _consolidado_falso(3))
    app.adicionar_caminhos(arquivos)
    app.processar()
    _rodar_ate(raiz, lambda: app.resultado is not None)
    app.tree_arquivos.selection_set("0")
    app.remover()
    app.salvar_como()
    _rodar_ate(raiz, lambda: dialogos.erros)
    assert "arquivos de entrada" in dialogos.erros[0]
    assert arquivos[0].read_bytes() == original
