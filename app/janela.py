"""Janela do Consolidador (tkinter/ttk). So telas: as regras ficam em
consolidacao/ e nos modulos sem Tk de app/.

Nenhum widget e tocado fora da thread principal: processar e gravar rodam em
app.tarefas.Trabalho, e a janela consome os eventos com after().
Dialogos (arquivos, perguntas, abrir no Explorer) passam por um objeto
injetavel, para os testes rodarem sem interacao.
"""
import getpass
import logging
import os
import socket
import re
import subprocess
import sys
import tkinter as tk
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

from consolidacao.erros import ErroValidacao
from consolidacao.pipeline import ETAPA_GRAVACAO, gravar, processar, resumir
from consolidacao.protecao import MARCADOR_APOLICE, PADRAO_SENHA_AUTOMATICA, padrao_valido

from . import NOME_APP, VERSAO
from . import dnd
from . import empresas_compartilhadas as ec
from .config import (
    carregar_config,
    guardar_senha_automatica,
    guardar_senhas,
    ler_padrao_de_fabrica,
    ler_senha_automatica,
    ler_senhas,
    salvar_config,
)
from .selecao import (
    adicionar,
    adicionar_pasta,
    alertas_pre_processamento,
    atualizar_empresas,
    mudancas,
    nome_saida_sugerido,
    retratar,
    texto_competencia,
)
from .tarefas import Trabalho

INTERVALO_MS = 100
DIAS_CACHE_VELHO = 7
MASCARA_SENHA = "•"

# Tokens do plano de design (.superpowers/design-ux.md).
PETROLEO = "#0E4A55"
NEVOA = "#EEF3F4"
PAPEL = "#FFFFFF"
ZEBRA = "#F6F9F9"
GRAFITE = "#1F2A2E"
CINZA = "#5B6B70"
AMBAR = "#E3A21A"
AMBAR_HOVER = "#C98D0F"
OK = "#2E7D4F"
ERRO = "#B3261E"
AVISO_TEXTO = "#7A5600"
AVISO_FUNDO = "#FFF4D6"
BORDA = "#D3DEE0"
TRILHO = "#DCE6E8"
FONTE = "Segoe UI"
FONTE_SEMI = "Segoe UI Semibold"
MARGEM = 16

log = logging.getLogger(__name__)


def _pasta_recursos() -> Path:
    """app/recursos (logo e icone, fora do git). Dentro do .exe o PyInstaller
    desempacota em sys._MEIPASS."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
    return base / "app" / "recursos"


class DialogosTk:
    """Dialogos reais do Windows; os testes trocam por um objeto falso."""

    def __init__(self, raiz):
        self.raiz = raiz

    def escolher_arquivos(self, pasta):
        nomes = filedialog.askopenfilenames(
            parent=self.raiz, title="Escolha as planilhas", initialdir=pasta or None,
            filetypes=[("Planilhas Excel", "*.xlsx")],
        )
        return [Path(n) for n in nomes]

    def escolher_pasta(self, pasta):
        nome = filedialog.askdirectory(parent=self.raiz, title="Escolha a pasta",
                                       initialdir=pasta or None)
        return Path(nome) if nome else None

    def salvar_como(self, nome, pasta):
        caminho = filedialog.asksaveasfilename(
            parent=self.raiz, title="Salvar consolidado como", initialfile=nome,
            initialdir=pasta or None, defaultextension=".xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")],
        )
        return Path(caminho) if caminho else None

    def primeiro_uso(self):
        """True = escolher arquivo existente; False = criar de um CSV; None = depois."""
        return messagebox.askyesnocancel(
            "Relação de empresas",
            "Já existe o arquivo da relação de empresas na rede?\n\n"
            "Sim: escolher o arquivo existente.\n"
            "Não: criar o arquivo a partir de um COD_EMPRESAS.csv.\n"
            "Cancelar: configurar depois.",
            parent=self.raiz,
        )

    def escolher_relacao(self):
        nome = filedialog.askopenfilename(
            parent=self.raiz, title="Arquivo da relação de empresas",
            filetypes=[("CSV", "*.csv"), ("Todos", "*.*")],
        )
        return Path(nome) if nome else None

    def escolher_csv_origem(self):
        nome = filedialog.askopenfilename(
            parent=self.raiz, title="CSV de origem (COD_EMPRESAS.csv)",
            filetypes=[("CSV", "*.csv")],
        )
        return Path(nome) if nome else None

    def escolher_destino_relacao(self):
        nome = filedialog.asksaveasfilename(
            parent=self.raiz, title="Onde criar a relação compartilhada",
            initialfile="empresas.csv", defaultextension=".csv", filetypes=[("CSV", "*.csv")],
        )
        return Path(nome) if nome else None

    def perguntar(self, titulo, mensagem):
        return messagebox.askyesno(titulo, mensagem, parent=self.raiz)

    def avisar(self, titulo, mensagem):
        messagebox.showwarning(titulo, mensagem, parent=self.raiz)

    def informar(self, titulo, mensagem):
        messagebox.showinfo(titulo, mensagem, parent=self.raiz)

    def erro(self, titulo, mensagem):
        messagebox.showerror(titulo, mensagem, parent=self.raiz)

    def abrir_arquivo(self, caminho):
        os.startfile(caminho)  # noqa: S606 - abre no programa padrao (Excel)

    def abrir_pasta(self, caminho):
        subprocess.Popen(["explorer", "/select,", str(caminho)])


def _data(d: datetime | None) -> str:
    return d.strftime("%d/%m/%Y %H:%M") if d else "?"


class App:
    def __init__(self, raiz, pasta_dados: Path, dialogos=None, processar_fn=None,
                 gravar_fn=None, hoje=None, primeiro_uso_automatico=True):
        self.raiz = raiz
        self.pasta_dados = Path(pasta_dados)
        self.dialogos = dialogos or DialogosTk(raiz)
        self.processar_fn = processar_fn or processar
        self.gravar_fn = gravar_fn or gravar
        self.hoje = hoje or date.today
        self.cfg = carregar_config(self.pasta_dados)
        self.usuario = getpass.getuser()
        self.maquina = socket.gethostname()

        self.relacao = None
        self.mapa = {}
        self.itens = []
        self.trabalho = None
        self.tipo_trabalho = None
        self.resultado = None
        self.retrato = None
        self.salvo = True
        self.ultimo_salvo = None
        self.saida_pendente = None
        self.saida_protegida = False
        self.fechando = False
        self.passo = 1
        self.dnd_ativo = dnd.habilitar(raiz)
        self._form_valido = False

        raiz.title(NOME_APP)
        raiz.geometry("1100x720")
        raiz.minsize(900, 600)
        raiz.protocol("WM_DELETE_WINDOW", self.fechar)
        self._montar()
        self._validar_formulario()
        self._atualizar_status_saida()
        self._atualizar_botoes()
        self.carregar_relacao(mostrar_erro=False)
        if primeiro_uso_automatico and not self.cfg.get("arquivo_empresas"):
            raiz.after(200, self.configurar_relacao)

    # --- montagem ------------------------------------------------------------

    def _estilos(self):
        s = ttk.Style(self.raiz)
        try:
            s.theme_use("clam")
        except tk.TclError:
            log.warning("Tema clam indisponivel; usando o tema padrao")
        for nome in ("TkDefaultFont", "TkTextFont", "TkHeadingFont", "TkMenuFont"):
            tkfont.nametofont(nome).configure(family=FONTE, size=10)
        corpo, semi = (FONTE, 10), (FONTE_SEMI, 10)
        s.configure(".", background=NEVOA, foreground=GRAFITE, font=corpo, bordercolor=BORDA,
                    focuscolor=PETROLEO, troughcolor=TRILHO, lightcolor=NEVOA, darkcolor=NEVOA)
        s.configure("Papel.TFrame", background=PAPEL)
        s.configure("Papel.TLabel", background=PAPEL)
        s.configure("Titulo.Papel.TLabel", font=(FONTE_SEMI, 12))
        s.configure("Negrito.Papel.TLabel", font=semi)
        s.configure("Nota.Papel.TLabel", foreground=CINZA, font=(FONTE, 9))
        s.configure("Nota.TLabel", foreground=CINZA, font=(FONTE, 9))
        s.configure("Erro.Papel.TLabel", foreground=ERRO, font=(FONTE, 9))
        s.configure("Ok.Papel.TLabel", foreground=OK, font=semi)
        s.configure("Titulo.Papel.TCheckbutton", background=PAPEL, font=(FONTE_SEMI, 12),
                    focuscolor=PAPEL)
        s.map("Titulo.Papel.TCheckbutton", background=[("active", PAPEL)])

        def botao(nome, fundo, sobre, pressionado, desativado, texto, texto_off, borda, fonte):
            s.configure(nome, padding=(14, 6), relief="flat", borderwidth=1, font=fonte,
                        background=fundo, foreground=texto, bordercolor=borda,
                        lightcolor=fundo, darkcolor=fundo)
            estados = [("disabled", desativado), ("pressed", pressionado), ("active", sobre)]
            s.map(nome, background=estados, lightcolor=estados, darkcolor=estados,
                  foreground=[("disabled", texto_off)],
                  bordercolor=[("disabled", "#D5DEE0"), ("focus", PETROLEO)])

        botao("TButton", "#E3ECEE", "#D3E0E3", "#C4D3D6", "#EBF0F1", GRAFITE, "#9AA8AC",
              "#B9C7CA", corpo)
        botao("Acao.TButton", AMBAR, AMBAR_HOVER, "#B07B0C", "#EFE6CF", GRAFITE, "#9A9484",
              "#B9800F", semi)

        s.configure("TNotebook", background=NEVOA, borderwidth=0, tabmargins=(0, 0, 0, 0),
                    bordercolor=NEVOA, lightcolor=NEVOA, darkcolor=NEVOA)
        s.configure("TNotebook.Tab", padding=(22, 8), background="#DCE6E8", foreground=GRAFITE,
                    font=semi, borderwidth=0, lightcolor="#DCE6E8", darkcolor="#DCE6E8")
        s.map("TNotebook.Tab", padding=[("selected", (22, 8))], background=[("selected", PETROLEO), ("active", "#CCDADD")],
              foreground=[("selected", "#FFFFFF")], lightcolor=[("selected", PETROLEO)],
              darkcolor=[("selected", PETROLEO)], expand=[("selected", (0, 0, 0, 0))])

        for nome in ("TEntry", "Erro.TEntry"):
            borda = ERRO if nome == "Erro.TEntry" else "#B9C7CA"
            s.configure(nome, fieldbackground=PAPEL, padding=6, bordercolor=borda,
                        lightcolor=borda, darkcolor=borda, insertcolor=GRAFITE)
            if nome == "TEntry":
                foco = [("focus", PETROLEO)]
                s.map(nome, bordercolor=foco, lightcolor=foco, darkcolor=foco)

        s.configure("Treeview", background=PAPEL, fieldbackground=PAPEL, rowheight=26,
                    borderwidth=0, font=corpo)
        s.map("Treeview", background=[("selected", PETROLEO)],
              foreground=[("selected", "#FFFFFF")])
        s.configure("Treeview.Heading", background="#E3ECEE", foreground=GRAFITE, font=semi,
                    relief="flat", padding=(8, 5), borderwidth=0, lightcolor="#E3ECEE",
                    darkcolor="#E3ECEE")
        s.map("Treeview.Heading", background=[("active", "#D3E0E3")])
        s.configure("TScrollbar", background="#C3D0D3", troughcolor=NEVOA, arrowcolor=CINZA,
                    bordercolor=NEVOA, lightcolor="#C3D0D3", darkcolor="#C3D0D3", relief="flat")
        s.configure("TProgressbar", background=PETROLEO, troughcolor="#DCE6E8", thickness=10,
                    borderwidth=0, lightcolor=PETROLEO, darkcolor=PETROLEO)

    def _identidade(self):
        """Icone e logo vem de app/recursos (nao versionados); sem eles o app
        abre igual, so com o nome na faixa."""
        pasta = _pasta_recursos()
        self._logo = None
        try:
            self.raiz.iconbitmap(default=str(pasta / "icone.ico"))
        except tk.TclError:
            pass
        try:
            self._logo = tk.PhotoImage(file=str(pasta / "logo_topo.png"))
        except tk.TclError:
            pass

    def _montar(self):
        self.raiz.configure(background=NEVOA)
        self._estilos()
        self._identidade()
        menu = tk.Menu(self.raiz)
        config = tk.Menu(menu, tearoff=False)
        config.add_command(label="Arquivo da relação de empresas...",
                           command=self.configurar_relacao)
        config.add_separator()
        config.add_command(label="Sobre", command=self.sobre)
        menu.add_cascade(label="Configurações", menu=config)
        self.raiz.config(menu=menu)

        self._montar_topo()
        self._montar_status()
        self.abas = ttk.Notebook(self.raiz)
        self.abas.pack(fill="both", expand=True, padx=MARGEM, pady=(12, 0))
        self.aba_consolidar = ttk.Frame(self.abas, padding=(0, 12, 0, 12))
        self.aba_empresas = ttk.Frame(self.abas, padding=(0, MARGEM, 0, MARGEM))
        self.aba_seguranca = ttk.Frame(self.abas, padding=(0, 12, 0, 10))
        self.abas.add(self.aba_consolidar, text="Consolidar")
        self.abas.add(self.aba_empresas, text="Empresas")
        self.abas.add(self.aba_seguranca, text="Segurança")
        self._montar_consolidar(self.aba_consolidar)
        self._montar_empresas(self.aba_empresas)
        self._montar_seguranca(self.aba_seguranca)
        self._atalhos()

    def _montar_topo(self):
        topo = tk.Frame(self.raiz, bg=PETROLEO, height=64)
        topo.pack(fill="x", side="top")
        topo.pack_propagate(False)
        textos = tk.Frame(topo, bg=PETROLEO)
        tk.Label(textos, text=NOME_APP, bg=PETROLEO, fg="#FFFFFF",
                 font=(FONTE_SEMI, 16)).pack(anchor="e")
        tk.Label(textos, text="Para planilhas de posição cadastral do Bradesco Saúde",
                 bg=PETROLEO, fg="#B7D2D7", font=(FONTE, 9)).pack(anchor="e")
        if self._logo is not None:
            tk.Label(topo, image=self._logo, bg=PETROLEO, borderwidth=0).pack(
                side="left", padx=(MARGEM, 0))
            textos.pack(side="right", padx=MARGEM, pady=8)
        else:
            textos.pack(side="left", padx=MARGEM, pady=8)

    def _montar_status(self):
        barra = tk.Frame(self.raiz, bg="#DCE6E8")
        barra.pack(fill="x", side="bottom")
        self.texto_status = tk.StringVar(value="Relação: não configurada")
        self.texto_status_saida = tk.StringVar(value="")
        tk.Label(barra, textvariable=self.texto_status, bg="#DCE6E8", fg=CINZA,
                 font=(FONTE, 9), anchor="w").pack(side="left", padx=MARGEM, pady=4)
        tk.Label(barra, text=f"v{VERSAO}", bg="#DCE6E8", fg=CINZA, font=(FONTE, 9)).pack(
            side="right", padx=MARGEM)
        tk.Label(barra, textvariable=self.texto_status_saida, bg="#DCE6E8", fg=CINZA,
                 font=(FONTE, 9)).pack(side="right", padx=(0, MARGEM))

    def _atalhos(self):
        self.raiz.bind_all("<Control-o>", self._atalho_abrir)
        self.raiz.bind_all("<Control-O>", self._atalho_abrir)
        self.raiz.bind_all("<Return>", self._atalho_enter)
        self.raiz.bind_all("<KP_Enter>", self._atalho_enter)
        self.raiz.bind_all("<Escape>", self._atalho_esc)

    def _atalho_abrir(self, evento):
        if self.ocupado:
            return "break"
        self.abas.select(self.aba_consolidar)
        if evento.state & 0x1:  # Shift
            self.adicionar_pasta()
        else:
            self.adicionar_arquivos()
        return "break"

    def _atalho_enter(self, _evento=None):
        """Enter dispara a acao ambar (a unica da tela): Processar, depois Salvar."""
        if self.abas.select() != str(self.aba_consolidar):
            return None
        botao = self.bt_processar if self.passo < 3 else self.bt_salvar
        if str(botao.cget("state")) != "disabled":
            botao.invoke()
        return "break"

    def _atalho_esc(self, _evento=None):
        """Esc na aba Empresas = Nova empresa (limpa selecao e formulario)."""
        if self.abas.select() != str(self.aba_empresas):
            return None
        self.nova_empresa()
        return "break"

    # --- blocos reutilizaveis ------------------------------------------------------

    @staticmethod
    def _cartao(pai):
        """Painel branco com filete (sem sombra). Devolve (externo, interno)."""
        externo = tk.Frame(pai, bg=PAPEL, highlightbackground=BORDA, highlightcolor=BORDA,
                           highlightthickness=1)
        interno = ttk.Frame(externo, style="Papel.TFrame", padding=(MARGEM, 10))
        interno.pack(fill="both", expand=True)
        return externo, interno

    def _cabecalho_passo(self, pai, numero, titulo):
        linha = ttk.Frame(pai, style="Papel.TFrame")
        linha.pack(fill="x")
        marca = tk.Label(linha, text=str(numero), width=3, font=(FONTE_SEMI, 11),
                         bg=PETROLEO, fg="#FFFFFF")
        marca.pack(side="left", padx=(0, 10))
        self._marcas_passo[numero] = marca
        ttk.Label(linha, text=titulo, style="Titulo.Papel.TLabel").pack(side="left")
        return linha

    def _tabela(self, pai, colunas, altura=None):
        """Treeview com rolagem. colunas = (id, titulo, largura, estica)."""
        quadro = ttk.Frame(pai, style="Papel.TFrame")
        extra = {} if altura is None else {"height": altura}
        arvore = ttk.Treeview(quadro, columns=[c[0] for c in colunas], show="headings", **extra)
        for col, titulo, largura, estica in colunas:
            arvore.heading(col, text=titulo, anchor="w")
            arvore.column(col, width=largura, minwidth=60, anchor="w", stretch=estica)
        arvore.tag_configure("zebra", background=ZEBRA)
        arvore.tag_configure("erro", foreground=ERRO)
        arvore.tag_configure("destaque", foreground=AVISO_TEXTO)
        rolagem = ttk.Scrollbar(quadro, orient="vertical", command=arvore.yview)
        arvore.configure(yscrollcommand=rolagem.set)
        rolagem.pack(side="right", fill="y")
        arvore.pack(side="left", fill="both", expand=True)
        return quadro, arvore

    def _faixa(self, pai):
        # Container sempre presente no topo da aba; a faixa aparece dentro dele.
        topo = ttk.Frame(pai)
        topo.pack(fill="x", side="top")
        return tk.Label(topo, bg=AVISO_FUNDO, fg=AVISO_TEXTO, anchor="w", padx=12, pady=6,
                        justify="left", font=(FONTE, 10))

    # --- aba Consolidar ------------------------------------------------------------

    def _montar_consolidar(self, f):
        self._marcas_passo = {}
        self.faixa_consolidar = self._faixa(f)
        self.bloqueio = tk.Frame(self.faixa_consolidar.master, bg=AVISO_FUNDO)
        tk.Label(self.bloqueio, text="Configure a relação de empresas para processar.",
                 bg=AVISO_FUNDO, fg=AVISO_TEXTO, font=(FONTE, 10), padx=12, pady=6).pack(
            side="left")
        ttk.Button(self.bloqueio, text="Configurar relação de empresas",
                   command=self.configurar_relacao).pack(side="left", padx=4, pady=4)

        # Os cartoes 2 e 3 sao empacotados antes (pelo fundo): se a janela
        # ficar baixa, quem encolhe e a tabela de arquivos, nao os botoes.
        ext3, passo3 = self._cartao(f)
        ext3.pack(fill="x", side="bottom", pady=(MARGEM, 0))
        ext2, passo2 = self._cartao(f)
        ext2.pack(fill="x", side="bottom", pady=(MARGEM, 0))
        ext1, passo1 = self._cartao(f)
        ext1.pack(fill="both", expand=True)

        # 1. Planilhas do mes
        cab = self._cabecalho_passo(passo1, 1, "Planilhas do mês")
        botoes = ttk.Frame(cab, style="Papel.TFrame")
        botoes.pack(side="right")
        self.bt_pasta = ttk.Button(botoes, text="Adicionar pasta", command=self.adicionar_pasta)
        self.bt_add = ttk.Button(botoes, text="Adicionar arquivos",
                                 command=self.adicionar_arquivos)
        self.bt_remover = ttk.Button(botoes, text="Remover", command=self.remover)
        for b in (self.bt_pasta, self.bt_add, self.bt_remover):
            b.pack(side="left", padx=(8, 0))
        quadro, self.tree_arquivos = self._tabela(passo1, (
            ("arquivo", "Arquivo", 380, True), ("empresa", "Empresa", 380, True),
            ("competencia", "Competência", 110, False)), altura=2)
        quadro.pack(fill="both", expand=True, pady=(10, 0))
        self._quadro_arquivos = quadro
        self.tree_arquivos.bind("<<TreeviewSelect>>", lambda _e: self._atualizar_botoes())
        self.vazio_arquivos = tk.Label(
            self.tree_arquivos,
            text="Adicione a pasta com as planilhas do mês\n"
                 + ("Arraste as planilhas de posição cadastral do Bradesco (.xlsx) ou a pasta "
                    "para cá, ou use os botões\n" if self.dnd_ativo else
                    "Use os botões para adicionar as planilhas de posição cadastral do "
                    "Bradesco (.xlsx)\n")
                 + "Ctrl+Shift+O adiciona uma pasta  ·  Ctrl+O adiciona arquivos",
            bg=PAPEL, fg=CINZA, font=(FONTE, 10), justify="center")
        if self.dnd_ativo:
            for alvo in (self.tree_arquivos, self.vazio_arquivos, self._quadro_arquivos):
                dnd.registrar_alvo(alvo, self.adicionar_soltos)
        self.texto_lista = tk.StringVar(value="")
        ttk.Label(cab, textvariable=self.texto_lista, style="Nota.Papel.TLabel").pack(
            side="left", padx=(MARGEM, 0))
        # Aviso amarelo (competencias misturadas etc.): so aparece quando ha o que avisar.
        self.texto_aviso_lista = tk.StringVar(value="")
        self.aviso_lista = tk.Label(passo1, textvariable=self.texto_aviso_lista, bg=AVISO_FUNDO,
                                    fg=AVISO_TEXTO, font=(FONTE, 9), anchor="w", padx=8, pady=4)

        # 2. Processar
        cab = self._cabecalho_passo(passo2, 2, "Processar")
        self.bt_processar = ttk.Button(cab, text="Processar", command=self.processar,
                                       style="Acao.TButton")
        self.bt_cancelar = ttk.Button(cab, text="Cancelar", command=self.cancelar,
                                      state="disabled")
        self.bt_processar.pack(side="left", padx=(MARGEM, 0))
        self.bt_cancelar.pack(side="left", padx=(8, 0))
        andamento = ttk.Frame(cab, style="Papel.TFrame")
        andamento.pack(side="left", fill="x", expand=True, padx=(MARGEM, 0))
        self.barra = ttk.Progressbar(andamento, mode="determinate")
        self.barra.pack(fill="x", pady=(4, 0))
        self.texto_progresso = tk.StringVar(value="")
        ttk.Label(andamento, textvariable=self.texto_progresso, style="Nota.Papel.TLabel").pack(
            anchor="w")

        # 3. Revisar e salvar
        cab = self._cabecalho_passo(passo3, 3, "Revisar e salvar")
        self.resumo_topo = tk.StringVar(value="")
        self.resumo_numeros = tk.StringVar(value="")
        self.resumo_origem = tk.StringVar(value="")
        self.dica_revisao = tk.StringVar(value="")
        self.rotulo_resumo = ttk.Label(cab, textvariable=self.resumo_topo, style="Ok.Papel.TLabel")
        self.rotulo_resumo.pack(side="left", padx=(MARGEM, 0))
        ttk.Label(cab, textvariable=self.dica_revisao, style="Nota.Papel.TLabel").pack(
            side="left", padx=(MARGEM, 0))
        notas = ttk.Frame(passo3, style="Papel.TFrame")
        notas.pack(fill="x", pady=(6, 0))
        ttk.Label(notas, textvariable=self.resumo_numeros, style="Nota.Papel.TLabel").pack(
            side="left")
        ttk.Label(notas, textvariable=self.resumo_origem, style="Nota.Papel.TLabel").pack(
            side="left", padx=(MARGEM, 0))
        salvar = ttk.Frame(passo3, style="Papel.TFrame")
        salvar.pack(fill="x", side="bottom", pady=(MARGEM, 0))
        quadro, self.tree_resumo = self._tabela(passo3, (
            ("arquivo", "Arquivo", 240, False), ("tipo", "Tipo", 90, False),
            ("mensagem", "Motivo (clique duplo para ver tudo)", 560, True)), altura=2)
        quadro.pack(fill="both", expand=True, pady=(8, 0))
        self.tree_resumo.bind("<Double-1>", self._detalhe_resumo)
        self.bt_salvar = ttk.Button(salvar, text="Salvar consolidado...", command=self.salvar_como,
                                    state="disabled", style="Acao.TButton")
        self.bt_abrir_arquivo = ttk.Button(salvar, text="Abrir arquivo",
                                           command=self.abrir_arquivo, state="disabled")
        self.bt_abrir_pasta = ttk.Button(salvar, text="Abrir pasta", command=self.abrir_pasta,
                                         state="disabled")
        self.bt_salvar.pack(side="left")
        self.bt_abrir_arquivo.pack(side="left", padx=(8, 0))
        self.bt_abrir_pasta.pack(side="left", padx=(8, 0))
        self.texto_salvo = tk.StringVar(value="")
        ttk.Label(salvar, textvariable=self.texto_salvo, style="Ok.Papel.TLabel").pack(
            side="left", padx=MARGEM)

    # --- aba Empresas ----------------------------------------------------------------

    def _montar_empresas(self, f):
        self.faixa_empresas = self._faixa(f)
        topo = ttk.Frame(f)
        topo.pack(fill="x", pady=(0, 12))
        ttk.Label(topo, text="Buscar").pack(side="left")
        self.busca = tk.StringVar()
        self.busca.trace_add("write", lambda *a: self._preencher_empresas())
        ttk.Entry(topo, textvariable=self.busca, width=44).pack(side="left", padx=8)

        self.rodape = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.rodape, style="Nota.TLabel").pack(side="bottom",
                                                                         anchor="w", pady=(8, 0))
        corpo = ttk.Frame(f)
        corpo.pack(fill="both", expand=True)
        quadro, self.tree_empresas = self._tabela(corpo, (
            ("empresa", "Empresa", 380, True), ("cod_saude", "Código de saúde", 130, False),
            ("cod_dental", "Código dental", 120, False)))
        quadro.pack(side="left", fill="both", expand=True)
        self.tree_empresas.bind("<<TreeviewSelect>>", self._selecionar_empresa)

        externo, form = self._cartao(corpo)
        externo.pack(side="left", fill="y", padx=(MARGEM, 0))
        cab = ttk.Frame(form, style="Papel.TFrame")
        cab.pack(fill="x")
        ttk.Label(cab, text="Dados da empresa", style="Titulo.Papel.TLabel").pack(side="left")
        self.bt_nova = ttk.Button(cab, text="Nova empresa", command=self.nova_empresa)
        self.bt_nova.pack(side="right")
        self.campo_empresa = tk.StringVar()
        self.campo_cod = tk.StringVar()
        self.campo_dental = tk.StringVar()
        self.erros_campo = {}
        self.entradas = {}
        self.visitados = set()
        for chave, rotulo, var in (("empresa", "Empresa", self.campo_empresa),
                                   ("cod_saude", "Código de saúde", self.campo_cod),
                                   ("cod_dental", "Código dental", self.campo_dental)):
            ttk.Label(form, text=rotulo, style="Papel.TLabel").pack(anchor="w", pady=(8, 3))
            entrada = ttk.Entry(form, textvariable=var, width=34)
            entrada.pack(anchor="w", fill="x")
            entrada.bind("<FocusOut>", lambda _e, c=chave: self._visitou(c), add="+")
            self.entradas[chave] = entrada
            self.erros_campo[chave] = tk.StringVar(value="")
            tk.Label(form, textvariable=self.erros_campo[chave], bg=PAPEL, fg=ERRO,
                     font=(FONTE, 9), wraplength=260, justify="left", anchor="nw",
                     height=2).pack(anchor="w", fill="x", pady=(2, 0))
        self.campo_cod.trace_add("write", lambda *a: self._validar_formulario())
        self.campo_empresa.trace_add("write", lambda *a: self._validar_formulario())
        self.texto_validacao = tk.StringVar(value="")

        botoes = ttk.Frame(form, style="Papel.TFrame")
        botoes.pack(anchor="w", pady=(8, 0))
        self.bt_incluir = ttk.Button(botoes, text="Incluir", command=self.incluir_empresa)
        self.bt_editar = ttk.Button(botoes, text="Salvar alteração", command=self.editar_empresa)
        self.bt_excluir = ttk.Button(botoes, text="Excluir", command=self.excluir_empresa)
        for b in (self.bt_incluir, self.bt_editar, self.bt_excluir):
            b.pack(side="left", padx=(0, 8))
        self.area_reprocessar = ttk.Frame(form, style="Papel.TFrame")
        self.area_reprocessar.pack(anchor="w", fill="x", pady=(MARGEM, 0))
        self.bt_reprocessar = ttk.Button(self.area_reprocessar, text="Processar de novo",
                                         command=self.processar_de_novo)

    # --- aba Seguranca ----------------------------------------------------------------

    def _montar_seguranca(self, f):
        self.campo_senha_saida = tk.StringVar()
        entrada, saida = ler_senhas(self.cfg)
        self.campo_senha_saida.set(saida)
        self._mostrando_senhas = False
        self.vars_senhas_entrada, self.ents_senhas_entrada = [], []
        self.bts_remover_senha, self._quadros_senha = [], []

        # Empacotado pelo fundo: o bloco da senha de saida e os botoes nunca
        # saem da tela; so a lista de senhas de entrada (rolavel) encolhe.
        botoes = ttk.Frame(f)
        botoes.pack(fill="x", side="bottom", pady=(MARGEM, 0))
        ttk.Label(botoes, text="As senhas ficam neste computador, cifradas para o seu usuário "
                               "do Windows. Elas não aparecem no Log, nos registros nem em "
                               "mensagens.",
                  style="Nota.TLabel", wraplength=400, justify="left").pack(side="right")
        self.bt_salvar_senhas = ttk.Button(botoes, text="Salvar senhas", command=self.salvar_senhas)
        self.bt_mostrar_senhas = ttk.Button(botoes, text="Mostrar senhas",
                                            command=self.alternar_mostrar_senhas)
        self.bt_salvar_senhas.pack(side="left")
        self.bt_mostrar_senhas.pack(side="left", padx=(8, 0))
        self.texto_senhas = tk.StringVar(value="")
        tk.Label(botoes, textvariable=self.texto_senhas, bg=NEVOA, fg=OK,
                 font=(FONTE_SEMI, 10)).pack(side="left", padx=MARGEM)

        linha = ttk.Frame(f)
        linha.pack(fill="x", side="top")
        linha.columnconfigure((0, 1), weight=1, uniform="bloco")
        # 1. Senha automatica (padrao informado pelo usuario, com {apolice})
        externo, bloco = self._cartao(linha)
        externo.grid(row=0, column=0, sticky="nsew", padx=(0, MARGEM // 2))
        self._padrao_de_fabrica = ler_padrao_de_fabrica(_pasta_recursos())
        if self._padrao_de_fabrica:
            log.info("Padrao de senha automatica de fabrica carregado de app/recursos")
        ativa, padrao = ler_senha_automatica(self.cfg, self._padrao_de_fabrica)
        self.var_senha_automatica = tk.BooleanVar(value=ativa)
        self.campo_padrao_senha = tk.StringVar(value=padrao)
        self.erro_padrao_senha = tk.StringVar(value="")
        self.chk_senha_automatica = ttk.Checkbutton(
            bloco, text="Tentar a senha automática", variable=self.var_senha_automatica,
            style="Titulo.Papel.TCheckbutton", command=self._senhas_mudaram)
        self.chk_senha_automatica.pack(anchor="w")
        ttk.Label(bloco, text="Usa o número da apólice do nome do arquivo no lugar de "
                              f"{MARCADOR_APOLICE}.",
                  style="Nota.Papel.TLabel", wraplength=380, justify="left").pack(
            anchor="w", pady=(4, 8))
        campo = ttk.Frame(bloco, style="Papel.TFrame")
        campo.pack(anchor="w")
        ttk.Label(campo, text="Padrão", style="Papel.TLabel").pack(side="left", padx=(0, 8))
        self.ent_padrao_senha = ttk.Entry(campo, textvariable=self.campo_padrao_senha,
                                          show=MASCARA_SENHA, width=30)
        self.ent_padrao_senha.pack(side="left")
        self._rotulo_erro_padrao = tk.Label(
            bloco, textvariable=self.erro_padrao_senha, bg=PAPEL, fg=ERRO, font=(FONTE, 9),
            anchor="w")  # so aparece quando ha erro (ver _validar_padrao)
        self.campo_padrao_senha.trace_add("write", lambda *a: self._senhas_mudaram())

        # 2. Senha do arquivo consolidado
        externo, bloco = self._cartao(linha)
        externo.grid(row=0, column=1, sticky="nsew", padx=(MARGEM // 2, 0))
        ttk.Label(bloco, text="Senha do arquivo consolidado",
                  style="Titulo.Papel.TLabel").pack(anchor="w")
        ttk.Label(bloco, text="Senha para proteger o arquivo consolidado "
                              "(deixe vazio para salvar sem senha).",
                  style="Papel.TLabel", wraplength=380, justify="left").pack(
            anchor="w", pady=(4, 8))
        self.ent_senha_saida = ttk.Entry(bloco, textvariable=self.campo_senha_saida,
                                         show=MASCARA_SENHA, width=34)
        self.ent_senha_saida.pack(anchor="w")

        externo, bloco = self._cartao(f)
        externo.pack(fill="both", expand=True, pady=(MARGEM, 0))
        cab = ttk.Frame(bloco, style="Papel.TFrame")
        cab.pack(fill="x")
        ttk.Label(cab, text="Senhas de entrada", style="Titulo.Papel.TLabel").pack(side="left")
        self.bt_adicionar_senha = ttk.Button(cab, text="+ Adicionar senha",
                                             command=lambda: self.adicionar_linha_senha(foco=True))
        self.bt_adicionar_senha.pack(side="right")
        ttk.Label(bloco, text="Senhas que o app tenta ao abrir planilhas protegidas. Cada senha "
                              "é testada, em ordem, até uma abrir a planilha.",
                  style="Papel.TLabel", wraplength=720, justify="left").pack(
            anchor="w", pady=(4, 8))
        area = ttk.Frame(bloco, style="Papel.TFrame")
        area.pack(fill="both", expand=True)
        # Lista rolavel: pelo menos ~3 linhas visiveis, o resto do espaco e dela.
        self.canvas_senhas = tk.Canvas(area, bg=PAPEL, highlightthickness=0, height=130,
                                       width=560)
        rolagem = ttk.Scrollbar(area, orient="vertical", command=self.canvas_senhas.yview)
        self.canvas_senhas.configure(yscrollcommand=rolagem.set)
        rolagem.pack(side="right", fill="y")
        self.canvas_senhas.pack(side="left", fill="both", expand=True)
        self.lista_senhas = ttk.Frame(self.canvas_senhas, style="Papel.TFrame")
        self.canvas_senhas.create_window((0, 0), window=self.lista_senhas, anchor="nw")
        self.lista_senhas.bind("<Configure>", lambda _e: self.canvas_senhas.configure(
            scrollregion=self.canvas_senhas.bbox("all")))
        self.raiz.bind_all("<MouseWheel>", self._roda_senhas, add="+")
        self.definir_senhas_entrada(entrada)
        self._senhas_guardadas = self._estado_senhas()
        self._validar_padrao()
        self.campo_senha_saida.trace_add("write", lambda *a: self._senhas_mudaram())

    def _roda_senhas(self, evento):
        """Roda do mouse sobre a lista de senhas (so se ela passa da altura)."""
        sob = self.raiz.winfo_containing(evento.x_root, evento.y_root)
        if sob is None or not str(sob).startswith(str(self.canvas_senhas)):
            return
        if self.lista_senhas.winfo_reqheight() > self.canvas_senhas.winfo_height():
            self.canvas_senhas.yview_scroll(-1 if evento.delta > 0 else 1, "units")

    def _mostrar_linha_senha(self, quadro):
        """Rola a lista ate a linha ficar inteira a vista."""
        if not quadro.winfo_exists():
            return
        self.raiz.update_idletasks()
        total = max(self.lista_senhas.winfo_reqheight(), 1)
        topo = self.canvas_senhas.canvasy(0)
        altura = self.canvas_senhas.winfo_height()
        y, h = quadro.winfo_y(), quadro.winfo_reqheight()
        if y < topo:
            self.canvas_senhas.yview_moveto(y / total)
        elif y + h > topo + altura:
            self.canvas_senhas.yview_moveto((y + h - altura) / total)

    def adicionar_linha_senha(self, valor="", foco=False):
        """Uma senha de entrada por linha: campo mascarado e botao Remover."""
        quadro = ttk.Frame(self.lista_senhas, style="Papel.TFrame")
        quadro.pack(anchor="w", pady=(0, 6))
        var = tk.StringVar(value=valor)
        entrada = ttk.Entry(quadro, textvariable=var, width=48,
                            show="" if self._mostrando_senhas else MASCARA_SENHA)
        entrada.pack(side="left")
        botao = ttk.Button(quadro, text="Remover")
        botao.config(command=lambda: self.remover_linha_senha(self.bts_remover_senha.index(botao)))
        botao.pack(side="left", padx=(8, 0))
        var.trace_add("write", lambda *a: self._senhas_mudaram())
        self.vars_senhas_entrada.append(var)
        self.ents_senhas_entrada.append(entrada)
        self.bts_remover_senha.append(botao)
        self._quadros_senha.append(quadro)
        if foco:
            entrada.focus_set()
            self._mostrar_linha_senha(quadro)
            self.raiz.after(50, lambda: self._mostrar_linha_senha(quadro))  # apos o layout
        self._senhas_mudaram()

    def remover_linha_senha(self, indice):
        if len(self.vars_senhas_entrada) == 1:  # a ultima linha so e esvaziada
            self.vars_senhas_entrada[0].set("")
            return
        for lista in (self.vars_senhas_entrada, self.ents_senhas_entrada,
                      self.bts_remover_senha):
            lista.pop(indice)
        self._quadros_senha.pop(indice).destroy()
        self._senhas_mudaram()

    def definir_senhas_entrada(self, senhas):
        for quadro in self._quadros_senha:
            quadro.destroy()
        self.vars_senhas_entrada, self.ents_senhas_entrada = [], []
        self.bts_remover_senha, self._quadros_senha = [], []
        for senha in list(senhas) or [""]:
            self.adicionar_linha_senha(senha)

    def alternar_mostrar_senhas(self):
        self._mostrando_senhas = not self._mostrando_senhas
        for e in self.ents_senhas_entrada + [self.ent_senha_saida, self.ent_padrao_senha]:
            e.config(show="" if self._mostrando_senhas else MASCARA_SENHA)
        self.bt_mostrar_senhas.config(
            text="Ocultar senhas" if self._mostrando_senhas else "Mostrar senhas")

    def _senhas_mudaram(self):
        if not hasattr(self, "_senhas_guardadas"):  # montagem ainda em andamento
            return
        self.texto_senhas.set("")
        self._validar_padrao()
        self._atualizar_botoes()
        self._atualizar_status_saida()

    def _padrao_senha(self) -> str:
        return self.campo_padrao_senha.get().strip()

    def _padrao_valido(self) -> bool:
        """Vazio (nao age) ou desligado nao e erro; preenchido precisa de {apolice}."""
        padrao = self._padrao_senha()
        return not self.var_senha_automatica.get() or not padrao or padrao_valido(padrao)

    def _validar_padrao(self):
        ativa = self.var_senha_automatica.get()
        self.ent_padrao_senha.config(state="normal" if ativa else "disabled")
        vazio = ativa and not self._padrao_senha()
        invalido = ativa and not self._padrao_valido()
        self.erro_padrao_senha.set(
            f"O padrão precisa conter {MARCADOR_APOLICE}." if invalido
            else "Preencha o padrão para ativar" if vazio else "")
        self._rotulo_erro_padrao.config(fg=ERRO if invalido else CINZA)
        self.ent_padrao_senha.config(style="Erro.TEntry" if invalido else "TEntry")
        if invalido or vazio:
            self._rotulo_erro_padrao.pack(anchor="w", fill="x")
        else:
            self._rotulo_erro_padrao.pack_forget()

    def _padrao_para_processar(self):
        """Padrao enviado ao motor: None = automatica desligada."""
        if self.var_senha_automatica.get() and self._padrao_senha() \
                and padrao_valido(self._padrao_senha()):
            return self._padrao_senha()
        return None

    def _estado_senhas(self) -> tuple:
        return (self._senhas_entrada(), self._senha_saida(),
                self.var_senha_automatica.get(), self._padrao_senha())

    def _atualizar_status_saida(self):
        self.texto_status_saida.set(
            "Saída protegida por senha" if self._senha_saida() else "Saída sem senha")

    def _senhas_entrada(self) -> list:
        """Uma senha por linha (pode conter ;). Linhas vazias nao contam."""
        partes = (v.get().strip() for v in self.vars_senhas_entrada)
        return list(dict.fromkeys(p for p in partes if p))

    def _senha_saida(self) -> str:
        return self.campo_senha_saida.get()

    def salvar_senhas(self):
        if not self._padrao_valido():
            return
        guardar_senhas(self.cfg, self._senhas_entrada(), self._senha_saida())
        guardar_senha_automatica(self.cfg, self.var_senha_automatica.get(),
                                 self._padrao_senha() or PADRAO_SENHA_AUTOMATICA,
                                 self._padrao_de_fabrica)
        try:
            salvar_config(self.pasta_dados, self.cfg)
        except OSError as e:
            self.dialogos.erro("Erro ao salvar as senhas",
                               f"Não foi possível gravar as preferências em {self.pasta_dados}: "
                               f"{e}. Confira se a pasta existe e tente de novo.")
            return
        self.definir_senhas_entrada(self._senhas_entrada())  # tira as linhas vazias
        self._senhas_guardadas = self._estado_senhas()
        self._atualizar_botoes()
        self.texto_senhas.set("Senhas salvas.")

    # --- relacao de empresas -----------------------------------------------------

    @property
    def caminho_relacao(self):
        valor = self.cfg.get("arquivo_empresas")
        return Path(valor) if valor else None

    def _caminho_para_ler(self):
        # Sem caminho configurado so serve o cache (se houver).
        return self.caminho_relacao or self.pasta_dados / "sem_relacao_configurada.csv"

    @staticmethod
    def _tarefa_ler(caminho, pasta_dados):
        """Roda fora da thread do Tk: a rede pode levar dezenas de segundos."""
        def tarefa():
            try:
                relacao = ec.ler(caminho, pasta_dados)
            except ErroValidacao as e:
                return None, None, str(e)
            try:
                return relacao, ec.mapa(relacao), None
            except ErroValidacao as e:
                return relacao, None, str(e)
        return tarefa

    def carregar_relacao(self, mostrar_erro=True, depois=None) -> None:
        """Le a relacao em segundo plano; depois() roda (na thread do Tk) so se
        a relacao lida for valida."""
        def aplicar(resultado):
            if self._aplicar_leitura(resultado, mostrar_erro) and depois is not None:
                depois()

        self._em_segundo_plano("relacao", self._tarefa_ler(self._caminho_para_ler(),
                                                            self.pasta_dados),
                               aplicar, "Lendo relação de empresas...")

    def _aplicar_leitura(self, resultado, mostrar_erro) -> bool:
        relacao, mapa, erro = resultado
        if relacao is not None:
            self.relacao = relacao
        # A tabela sempre mostra a relacao em memoria, mesmo invalida: e assim
        # que a pessoa ve (e corrige) a duplicata.
        self._preencher_empresas()
        if erro is not None:
            # Relacao ilegivel ou invalida: nada processa com ela.
            self.mapa = {}
            if mostrar_erro:
                self.dialogos.erro("Relação de empresas", erro)
            self._atualizar_estado_relacao()
            return False
        self.mapa = mapa
        self.itens = atualizar_empresas(self.itens, self.mapa)
        self._preencher_arquivos()
        self._atualizar_estado_relacao()
        return True

    def _atualizar_estado_relacao(self):
        r = self.relacao
        cache = r is not None and r.origem == "cache"
        for faixa in (self.faixa_consolidar, self.faixa_empresas):
            if cache:
                faixa.config(text=f"Usando cópia local de {_data(r.data_cache)}; "
                                  "a relação pode estar desatualizada.")
                faixa.pack(fill="x", side="top", pady=(0, 6))
            else:
                faixa.pack_forget()
        if r is None:
            self.bloqueio.pack(fill="x", side="top", pady=(0, 6))
        else:
            self.bloqueio.pack_forget()
        edicao = "normal" if self._pode_editar() else "disabled"
        self._atualizar_botoes()
        if r is None:
            self.rodape.set("Relação de empresas não configurada.")
            self.texto_status.set("Relação: não configurada")
        else:
            origem = "rede" if r.origem == "rede" else f"cópia local de {_data(r.data_cache)}"
            self.texto_status.set(
                f"Relação: rede (lida {r.lida_em:%H:%M})" if r.origem == "rede"
                else f"Relação: cópia local de {_data(r.data_cache)}")
            sem_rede = "" if edicao == "normal" else "  (edição bloqueada sem rede)"
            self.rodape.set(f"Arquivo: {self.caminho_relacao or '-'}  |  lido em "
                            f"{_data(r.lida_em)}  |  origem: {origem}{sem_rede}")

    def configurar_relacao(self):
        if self.ocupado:
            # Abertura automatica enquanto a 1a leitura ainda roda: espera.
            self.raiz.after(200, self.configurar_relacao)
            return
        escolha = self.dialogos.primeiro_uso()
        if escolha is None:
            return
        if escolha:
            caminho = self.dialogos.escolher_relacao()
            if not caminho:
                return
            origem = None
        else:
            origem = self.dialogos.escolher_csv_origem()
            if not origem:
                return
            caminho = self.dialogos.escolher_destino_relacao()
            if not caminho:
                return
        pasta = self.pasta_dados

        def tarefa():
            try:
                if origem is None:
                    ec.ler(caminho, pasta)  # valida que da para ler
                else:
                    ec.criar_a_partir_de_csv(origem, caminho)
            except ErroValidacao as e:
                return str(e)
            return None

        def depois(erro):
            if erro is not None:
                self.dialogos.erro("Relação de empresas", erro)
                return
            self.cfg["arquivo_empresas"] = str(caminho)
            salvar_config(self.pasta_dados, self.cfg)
            self.carregar_relacao()

        self._em_segundo_plano("relacao", tarefa, depois, "Lendo relação de empresas...")

    def sobre(self):
        self.dialogos.informar("Sobre", f"{NOME_APP}\nVersão {VERSAO}")

    # --- lista de arquivos --------------------------------------------------------

    def _preencher_arquivos(self):
        t = self.tree_arquivos
        t.delete(*t.get_children())
        for i, item in enumerate(self.itens):
            empresa = item.empresa or (item.erro or "")
            tags = (("erro",) if item.erro else ()) + (("zebra",) if i % 2 else ())
            t.insert("", "end", iid=str(i), tags=tags,
                     values=(item.caminho.name, empresa, texto_competencia(item.competencia)))
        if self.itens:
            self.vazio_arquivos.place_forget()
        else:
            self.vazio_arquivos.place(relx=0.5, rely=0.55, anchor="center")
        self._atualizar_resumo_lista()
        self._atualizar_botoes()

    def _atualizar_resumo_lista(self):
        """Resumo ao lado do titulo (quantos arquivos, qual competencia) e, so
        quando ha o que avisar, uma faixa amarela: competencias misturadas,
        arquivo sem competencia no nome ou sem empresa na relacao."""
        n = len(self.itens)
        resumo, avisos = [], []
        if n:
            resumo.append(f"{n} arquivo" + ("s" if n > 1 else ""))
            competencias = sorted({i.competencia for i in self.itens if i.competencia})
            if len(competencias) == 1:
                resumo.append(f"competência {texto_competencia(competencias[0])}")
            elif competencias:
                avisos.append("Competências diferentes: "
                              + ", ".join(texto_competencia(c) for c in competencias))
            sem = sum(1 for i in self.itens if not i.competencia)
            if sem:
                avisos.append(f"{sem} arquivo(s) sem competência (MM.AAAA) no nome")
            sem_empresa = sum(1 for i in self.itens if i.erro)
            if sem_empresa:
                avisos.append(f"{sem_empresa} arquivo(s) sem empresa na relação")
        self.texto_lista.set("  ·  ".join(resumo))
        if avisos:
            self.texto_aviso_lista.set("  ·  ".join(avisos))
            self.aviso_lista.pack(fill="x", pady=(8, 0), before=self._quadro_arquivos)
        else:
            self.aviso_lista.pack_forget()

    def adicionar_caminhos(self, caminhos):
        self.itens = adicionar(self.itens, caminhos, self.mapa)
        self._preencher_arquivos()

    def adicionar_soltos(self, caminhos):
        """Arquivos e pastas arrastados do Explorer: mesmas regras dos botoes
        (so .xlsx, sem ~$ nem PC_CONSOLIDADO_*, sem repetir)."""
        if self.ocupado:
            return
        caminhos = [Path(c) for c in caminhos]
        self.abas.select(self.aba_consolidar)
        itens = adicionar(self.itens, [c for c in caminhos if not c.is_dir()], self.mapa)
        for pasta in (c for c in caminhos if c.is_dir()):
            try:
                itens = adicionar_pasta(itens, pasta, self.mapa)
            except OSError as e:
                self.dialogos.erro("Adicionar pasta", f"Não foi possível ler a pasta {pasta}: {e}")
        self.itens = itens
        self._preencher_arquivos()

    def adicionar_arquivos(self):
        caminhos = self.dialogos.escolher_arquivos(self.cfg.get("ultima_pasta_entrada"))
        if caminhos:
            self._lembrar("ultima_pasta_entrada", Path(caminhos[0]).parent)
            self.adicionar_caminhos(caminhos)

    def adicionar_pasta(self):
        pasta = self.dialogos.escolher_pasta(self.cfg.get("ultima_pasta_entrada"))
        if pasta:
            self._lembrar("ultima_pasta_entrada", pasta)
            try:
                self.itens = adicionar_pasta(self.itens, pasta, self.mapa)
            except OSError as e:
                self.dialogos.erro("Adicionar pasta", f"Não foi possível ler a pasta: {e}")
            self._preencher_arquivos()

    def remover(self):
        escolhidos = {int(i) for i in self.tree_arquivos.selection()}
        self.itens = [it for i, it in enumerate(self.itens) if i not in escolhidos]
        self._preencher_arquivos()

    def _lembrar(self, chave, pasta):
        self.cfg[chave] = str(pasta)
        try:
            salvar_config(self.pasta_dados, self.cfg)
        except OSError:
            log.exception("Nao foi possivel salvar as preferencias")

    # --- processar ----------------------------------------------------------------

    @property
    def ocupado(self):
        return self.trabalho is not None and self.tipo_trabalho is not None

    def _atualizar_botoes(self):
        ocupado = self.ocupado
        normal = "disabled" if ocupado else "normal"
        for b in (self.bt_add, self.bt_pasta):
            b.config(state=normal)
        self.bt_remover.config(state="disabled" if ocupado or not self.tree_arquivos.selection()
                               else "normal")
        pode = not ocupado and bool(self.itens) and self.relacao is not None
        self.bt_processar.config(state="normal" if pode else "disabled")
        self.bt_cancelar.config(state="normal" if self.tipo_trabalho == "processar" else "disabled")
        self.bt_salvar.config(
            state="normal" if (self.resultado is not None and not ocupado) else "disabled")
        aberto = "normal" if (self.ultimo_salvo is not None and not ocupado) else "disabled"
        self.bt_abrir_arquivo.config(state=aberto)
        self.bt_abrir_pasta.config(state=aberto)
        self._atualizar_passos()
        self._atualizar_botoes_empresas(ocupado)
        mudou = self._estado_senhas() != self._senhas_guardadas
        self.bt_salvar_senhas.config(
            state="normal" if mudou and self._padrao_valido() else "disabled")

    def _atualizar_passos(self):
        """Passo atual: 1 sem arquivos, 2 com arquivos (ou processando), 3 com
        resultado. So ha um botao ambar: Processar ate processar, depois Salvar."""
        self.passo = 3 if self.resultado is not None else (2 if self.itens else 1)
        for n, marca in self._marcas_passo.items():
            atual = n == self.passo
            marca.config(bg=PETROLEO if atual else "#D3DEE0", fg="#FFFFFF" if atual else CINZA)
        self.bt_processar.config(style="Acao.TButton" if self.passo < 3 else "TButton")
        self.bt_salvar.config(style="Acao.TButton" if self.passo == 3 else "TButton")
        self.dica_revisao.set("" if self.resultado is not None
                              else "O resumo aparece aqui depois de processar.")

    def _atualizar_botoes_empresas(self, ocupado):
        pode = self._pode_editar() and not ocupado
        valido = pode and self._form_valido
        linha = self._linha_selecionada() if pode else None
        selecionada = linha is not None
        # Com uma empresa selecionada, "Incluir" so vale para um codigo novo.
        novo = not selecionada or self.campo_cod.get().strip() != linha.cod_saude
        self.bt_incluir.config(state="normal" if valido and novo else "disabled")
        self.bt_editar.config(state="normal" if valido and selecionada else "disabled")
        self.bt_excluir.config(state="normal" if selecionada else "disabled")

    def _pode_editar(self) -> bool:
        r = self.relacao
        return r is not None and r.origem == "rede" and self.caminho_relacao is not None

    def processar(self):
        if self.ocupado or not self.itens:
            return
        # A relacao e relida logo antes de processar (fora da thread do Tk).
        self.carregar_relacao(depois=self._processar_com_relacao)

    def _processar_com_relacao(self):
        if self.ocupado or not self.itens:
            return
        alertas = alertas_pre_processamento(self.itens)
        r = self.relacao
        if r.origem == "cache" and r.data_cache and (
                datetime.now() - r.data_cache > timedelta(days=DIAS_CACHE_VELHO)):
            alertas.insert(0, f"A relação de empresas é uma cópia local de {_data(r.data_cache)} "
                              f"(mais de {DIAS_CACHE_VELHO} dias).")
        if alertas and not self.dialogos.perguntar(
                "Confira antes de processar",
                "\n\n".join(alertas) + "\n\nProcessar mesmo assim?"):
            return
        caminhos = [i.caminho for i in self.itens]
        mapa = dict(self.mapa)
        self.caminhos_processados = list(caminhos)
        self.retrato = retratar(self.itens, r)
        self.origem_processada = r
        self.resultado = None
        self.ultimo_salvo = None
        self.texto_salvo.set("")
        self._limpar_resumo()
        self.barra.config(maximum=len(caminhos), value=0)
        self.texto_progresso.set("Iniciando...")
        # So repassa o que foi cadastrado: sem senhas a chamada e a de sempre.
        senhas = self._senhas_entrada()
        extras = {"senhas": senhas} if senhas else {}
        padrao = self._padrao_para_processar()
        if padrao:  # sem padrao o motor ja nao tenta a senha automatica
            extras["padrao_senha"] = padrao
        self.saida_protegida = bool(self._senha_saida())
        self._iniciar("processar", lambda p, c: self.processar_fn(caminhos, mapa, p, c, **extras))

    def processar_de_novo(self):
        self.abas.select(self.aba_consolidar)
        self.processar()

    def cancelar(self):
        if self.tipo_trabalho == "processar":
            self.trabalho.cancelar()
            self.texto_progresso.set("Cancelando...")
            self.bt_cancelar.config(state="disabled")

    def _em_segundo_plano(self, tipo, tarefa, depois, texto):
        """tarefa() roda numa thread (nunca toca em widget); depois(resultado)
        roda na thread do Tk quando ela termina."""
        self._continuacao = depois
        self.texto_progresso.set(texto)
        self._iniciar(tipo, lambda p, c: tarefa())

    def _iniciar(self, tipo, funcao):
        self.trabalho = Trabalho(funcao)
        self.tipo_trabalho = tipo
        self._atualizar_botoes()
        self.trabalho.iniciar()
        self.raiz.after(INTERVALO_MS, self._acompanhar)

    def _acompanhar(self):
        if self.trabalho is None:
            return
        terminou = False
        for evento in self.trabalho.eventos():
            terminou = self._tratar(evento) or terminou
        if not terminou:
            self.raiz.after(INTERVALO_MS, self._acompanhar)

    def _tratar(self, evento) -> bool:
        tipo = self.tipo_trabalho
        if evento.tipo == "progresso":
            i, total, nome = evento.dados
            if nome == ETAPA_GRAVACAO:
                self.texto_progresso.set("Gravando o arquivo de saída...")
            else:
                self.barra.config(maximum=total, value=i - 1)
                self.texto_progresso.set(f"Processando {i}/{total}: {nome}")
            return False
        self.trabalho, self.tipo_trabalho = None, None
        if tipo in ("relacao", "empresa"):
            self.texto_progresso.set("")
            self._atualizar_botoes()
            if evento.tipo == "concluido":
                self._continuacao(evento.dados)
            else:
                self.dialogos.erro("Relação de empresas", evento.dados)
            return True
        if evento.tipo == "concluido":
            if tipo == "processar":
                self.barra.config(value=self.barra.cget("maximum"))
                self.texto_progresso.set("Processamento concluído.")
                self.resultado = evento.dados
                self.salvo = False
                self._mostrar_resumo(self.resultado)
            else:
                self.salvo = True
                self.ultimo_salvo = self.saida_pendente
                self.texto_progresso.set("")
                self.texto_salvo.set(
                    f"Consolidado salvo em {self.ultimo_salvo}"
                    + (" (protegido por senha)" if self.saida_protegida else ""))
        elif evento.tipo == "cancelado":
            self.barra.config(value=0)
            self.texto_progresso.set("Processamento cancelado.")
        else:
            self.texto_progresso.set("")
            titulo = "Erro ao processar" if tipo == "processar" else "Erro ao salvar"
            self.dialogos.erro(titulo, evento.dados)
        self._atualizar_botoes()
        if self.fechando:
            self.raiz.destroy()
        return True

    # --- resumo -------------------------------------------------------------------

    def _limpar_resumo(self):
        self.resumo_topo.set("")
        self.resumo_numeros.set("")
        self.resumo_origem.set("")
        self.tree_resumo.delete(*self.tree_resumo.get_children())
        self.mensagens_resumo = {}

    def _mostrar_resumo(self, c):
        self._limpar_resumo()
        r = resumir(c)
        self.resumo_topo.set(f"{r['arquivos_ok']} OK, {r['arquivos_erro']} com erro")
        self.rotulo_resumo.config(
            style="Erro.Papel.TLabel" if r["arquivos_erro"] else "Ok.Papel.TLabel",
            font=(FONTE_SEMI, 10))
        self.resumo_numeros.set(
            f"Registros: {r['registros']}  |  Pendências: {r['pendencias']}  |  "
            f"Datas inválidas: {r['datas_invalidas']}  |  "
            f"Divergências de datas: {r['divergencias_datas']}")
        origem = self.origem_processada
        self.resumo_origem.set(
            "Relação de empresas: rede" if origem.origem == "rede"
            else f"Relação de empresas: cópia local de {_data(origem.data_cache)}")
        if self.saida_protegida:
            self.resumo_origem.set(self.resumo_origem.get() + "  |  Saída protegida por senha")
        self.repetidos = 0
        for n, linha in enumerate(c.log.to_dict("records")):
            if linha["STATUS"] == "ERRO":
                self._linha_resumo(n, linha["ARQUIVO"], "Erro", linha["MENSAGEM"], "erro")
            avisos = linha.get("AVISOS_ESTRUTURA")
            if avisos:
                repetido = "repetid" in avisos
                self.repetidos += repetido
                self._linha_resumo(f"a{n}", linha["ARQUIVO"],
                                   "REPETIDO" if repetido else "Aviso", avisos,
                                   "destaque" if repetido else ())

    @staticmethod
    def _com_dica_de_estrutura(mensagem):
        """Aba TITULAR/DEPENDENTE nao encontrada: diz o que conferir (so na tela;
        a mensagem das regras fica igual a do R)."""
        if "Esperada exatamente uma aba" in str(mensagem):
            return f"{mensagem} Confira se é uma planilha de posição cadastral do Bradesco Saúde."
        return mensagem

    def _linha_resumo(self, iid, arquivo, tipo, mensagem, tag):
        iid = str(iid)
        if tag == "erro":
            mensagem = self._com_dica_de_estrutura(mensagem)
        self.mensagens_resumo[iid] = mensagem
        self.tree_resumo.insert("", "end", iid=iid, values=(arquivo, tipo, mensagem),
                                tags=(tag,) if tag else ())

    def _detalhe_resumo(self, _evento=None):
        for iid in self.tree_resumo.selection():
            arquivo, tipo, _ = self.tree_resumo.item(iid, "values")
            self.dialogos.informar(f"{tipo}: {arquivo}", self.mensagens_resumo.get(iid, ""))

    # --- salvar -------------------------------------------------------------------

    def salvar_como(self):
        if self.resultado is None or self.ocupado:
            return
        caminho, pasta = self._caminho_para_ler(), self.pasta_dados

        def tarefa():
            try:
                return ec.ler(caminho, pasta)
            except ErroValidacao:
                return None

        self._em_segundo_plano("relacao", tarefa, self._salvar_com_relacao,
                               "Lendo relação de empresas...")

    def _salvar_com_relacao(self, atual):
        if self.resultado is None or self.ocupado:
            return
        atual = atual or self.relacao
        avisos = mudancas(self.retrato, self.itens, atual)
        if avisos and self.dialogos.perguntar(
                "Algo mudou desde o processamento",
                "\n".join(avisos) + "\n\nIsso mudou desde o processamento; reprocessar?"):
            self.processar()
            return
        r = resumir(self.resultado)
        problemas = []
        if r["arquivos_erro"]:
            problemas.append(f"{r['arquivos_erro']} arquivo(s) com erro")
        if r["pendencias"]:
            problemas.append(f"{r['pendencias']} pendência(s)")
        if getattr(self, "repetidos", 0):
            problemas.append(f"{self.repetidos} arquivo(s) com empresa ou referência repetida")
        if problemas and not self.dialogos.perguntar(
                "Salvar assim?", "Há " + ", ".join(problemas) + ".\n\nSalvar mesmo assim?"):
            return
        destino = self.dialogos.salvar_como(nome_saida_sugerido(self.itens, self.hoje()),
                                            self.cfg.get("ultima_pasta_saida"))
        if not destino:
            return
        destino = Path(destino)
        self._lembrar("ultima_pasta_saida", destino.parent)
        self.saida_pendente = destino
        resultado = self.resultado
        # A saida nao pode ser nenhuma entrada: nem as da lista atual nem as
        # processadas (uma planilha removida da lista continua sendo entrada).
        caminhos = list(dict.fromkeys(
            list(getattr(self, "caminhos_processados", [])) + [i.caminho for i in self.itens]))
        empresas = self.caminho_relacao
        self.texto_progresso.set("Gravando o arquivo de saída...")
        senha_saida = self._senha_saida()
        extras = {"senha_saida": senha_saida} if senha_saida else {}
        self.saida_protegida = bool(senha_saida)
        self._iniciar("gravar", lambda p, c: self.gravar_fn(
            resultado, destino, caminhos, arquivo_empresas=empresas, progresso=p, **extras))

    def abrir_arquivo(self):
        if self.ultimo_salvo:
            self.dialogos.abrir_arquivo(self.ultimo_salvo)

    def abrir_pasta(self):
        if self.ultimo_salvo:
            self.dialogos.abrir_pasta(self.ultimo_salvo)

    # --- empresas -------------------------------------------------------------------

    def _preencher_empresas(self):
        t = self.tree_empresas
        t.delete(*t.get_children())
        if self.relacao is None:
            return
        filtro = self.busca.get().strip().casefold()
        # iid = COD_SAUDE (com sufixo se alguem o duplicou); a selecao volta
        # para a linha exata da relacao em memoria, nunca por posicao.
        self._linhas_da_tabela = {}
        vistos = {}
        for l in self.relacao.linhas:
            k = ec.chave(l)
            vistos[k] = vistos.get(k, 0) + 1
            iid = k if vistos[k] == 1 and k else f"{k}#{vistos[k]}"
            self._linhas_da_tabela[iid] = l
            if filtro and filtro not in f"{l.empresa} {l.cod_saude} {l.cod_dental}".casefold():
                continue
            t.insert("", "end", iid=iid, values=(l.empresa, l.cod_saude, l.cod_dental))

    def _linha_selecionada(self):
        sel = self.tree_empresas.selection()
        if not sel or self.relacao is None:
            return None
        return getattr(self, "_linhas_da_tabela", {}).get(sel[0])

    def _selecionar_empresa(self, _evento=None):
        linha = self._linha_selecionada()
        if linha is not None:
            self.campo_empresa.set(linha.empresa)
            self.campo_cod.set(linha.cod_saude)
            self.campo_dental.set(linha.cod_dental)
        self._atualizar_botoes()

    def _formulario(self):
        return ec.Linha(self.campo_empresa.get(), self.campo_cod.get(), self.campo_dental.get())

    def _visitou(self, campo):
        self.visitados.add(campo)
        self._validar_formulario()

    @staticmethod
    def _texto_do_erro(mensagem: str) -> str:
        """Mensagem da regra, com os nomes dos campos como aparecem na tela."""
        texto = mensagem.replace("COD_SAUDE", "código de saúde").replace("EMPRESA", "empresa")
        return texto[:1].upper() + texto[1:]

    def _validar_formulario(self) -> str:
        try:
            linha = ec.validar_linha(self._formulario())
        except ErroValidacao as e:
            mensagem = str(e)
        else:
            mensagem = ""
            selecionada = self._linha_selecionada()
            for l in (self.relacao.linhas if self.relacao else []):
                if ec.chave(l) == linha.cod_saude and l is not selecionada:
                    mensagem = f"COD_SAUDE {linha.cod_saude} já cadastrado para {l.empresa}."
                    break
        vazio = not any(v.get().strip() for v in
                        (self.campo_empresa, self.campo_cod, self.campo_dental))
        self._form_valido = not mensagem
        # Formulario virgem nao mostra erro; campo obrigatorio vazio so depois
        # que a pessoa passou por ele; duplicata aparece na hora.
        self.texto_validacao.set("" if vazio else mensagem)
        campo = ("cod_saude" if "COD_SAUDE" in mensagem
                 else "empresa" if "EMPRESA" in mensagem else None)
        mostrar = bool(mensagem) and not vazio and (
            "já cadastrado" in mensagem or campo in self.visitados)
        for chave, var in self.erros_campo.items():
            ativo = mostrar and chave == campo
            var.set(self._texto_do_erro(mensagem) if ativo else "")
            self.entradas[chave].config(style="Erro.TEntry" if ativo else "TEntry")
        self._atualizar_botoes()
        return mensagem

    def _aplicar(self, alteracao, depois=None) -> None:
        caminho, base, pasta = self.caminho_relacao, self.relacao, self.pasta_dados
        usuario, maquina = self.usuario, self.maquina

        def tarefa():
            try:
                relacao = ec.aplicar(caminho, base, alteracao, usuario, maquina,
                                     pasta_dados=pasta)
            except ec.ConflitoEdicao as e:
                return "conflito", str(e)
            except ErroValidacao as e:
                return "erro", str(e)
            return "ok", (relacao, ec.mapa(relacao))

        def aplicada(resultado):
            situacao, dados = resultado
            if situacao == "conflito":
                self.dialogos.avisar("Relação de empresas", dados)
                self.carregar_relacao()
                return
            if situacao == "erro":
                # Inclui EmUso: a edicao continua no formulario para tentar de
                # novo e a tabela volta a mostrar a relacao em memoria.
                self.dialogos.erro("Relação de empresas", dados)
                self._preencher_empresas()
                return
            self.relacao, self.mapa = dados
            self.itens = atualizar_empresas(self.itens, self.mapa)
            self._preencher_arquivos()
            self._preencher_empresas()
            self._atualizar_estado_relacao()
            if depois is not None:
                depois()

        self._em_segundo_plano("empresa", tarefa, aplicada,
                               "Gravando a relação de empresas...")

    def incluir_empresa(self):
        mensagem = self._validar_formulario()
        if mensagem and "já cadastrado" not in mensagem:
            # O erro ja esta embaixo do campo; so garante que apareca.
            self.visitados.update(("empresa", "cod_saude"))
            self._validar_formulario()
            return

        def incluida():
            self._limpar_formulario()
            if self.itens:
                self.bt_reprocessar.pack(side="left")

        self._aplicar(ec.incluir(self._formulario()), incluida)

    def editar_empresa(self):
        antiga = self._linha_selecionada()
        if antiga is None:
            self.dialogos.avisar("Relação de empresas", "Selecione a empresa a editar.")
            return
        self._aplicar(ec.editar(antiga, self._formulario()), self._limpar_formulario)

    def excluir_empresa(self):
        antiga = self._linha_selecionada()
        if antiga is None:
            self.dialogos.avisar("Relação de empresas", "Selecione a empresa a excluir.")
            return
        if not self.dialogos.perguntar(
                "Excluir empresa",
                f"Excluir {antiga.empresa} (COD_SAUDE {antiga.cod_saude}) da relação?"):
            return
        self._aplicar(ec.excluir(antiga), self._limpar_formulario)

    def nova_empresa(self):
        """Limpa a selecao e o formulario e leva o foco ao campo Empresa."""
        self._limpar_formulario()
        self.entradas["empresa"].focus_set()

    def _limpar_formulario(self):
        self.tree_empresas.selection_remove(*self.tree_empresas.selection())
        for var in (self.campo_empresa, self.campo_cod, self.campo_dental):
            var.set("")
        self.visitados.clear()
        self._validar_formulario()

    # --- fechar ---------------------------------------------------------------------

    def fechar(self):
        if self.tipo_trabalho == "processar":
            if not self.dialogos.perguntar("Fechar",
                                           "O processamento está em andamento. Cancelar e fechar?"):
                return
            self.trabalho.cancelar()
        elif self.tipo_trabalho in ("gravar", "empresa"):
            self.dialogos.avisar("Fechar", "Aguarde terminar a gravação em andamento.")
            return
        elif self.resultado is not None and not self.salvo:
            if not self.dialogos.perguntar("Fechar",
                                           "O resultado do processamento não foi salvo. "
                                           "Fechar mesmo assim?"):
                return
        self.raiz.destroy()
