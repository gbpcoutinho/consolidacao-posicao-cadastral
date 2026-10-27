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
import subprocess
import tkinter as tk
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from consolidacao.erros import ErroValidacao
from consolidacao.pipeline import ETAPA_GRAVACAO, gravar, processar, resumir

from . import NOME_APP, VERSAO
from . import empresas_compartilhadas as ec
from .config import carregar_config, salvar_config
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
AMARELO = "#fff3b0"

log = logging.getLogger(__name__)


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
        self.fechando = False

        raiz.title(f"{NOME_APP} {VERSAO}")
        raiz.geometry("980x640")
        raiz.minsize(760, 520)
        raiz.protocol("WM_DELETE_WINDOW", self.fechar)
        self._montar()
        self.carregar_relacao(mostrar_erro=False)
        if primeiro_uso_automatico and not self.cfg.get("arquivo_empresas"):
            raiz.after(200, self.configurar_relacao)

    # --- montagem ------------------------------------------------------------

    def _montar(self):
        menu = tk.Menu(self.raiz)
        config = tk.Menu(menu, tearoff=False)
        config.add_command(label="Arquivo da relação de empresas...",
                           command=self.configurar_relacao)
        config.add_separator()
        config.add_command(label="Sobre", command=self.sobre)
        menu.add_cascade(label="Configurações", menu=config)
        self.raiz.config(menu=menu)

        self.abas = ttk.Notebook(self.raiz)
        self.abas.pack(fill="both", expand=True, padx=8, pady=8)
        self.aba_consolidar = ttk.Frame(self.abas, padding=8)
        self.aba_empresas = ttk.Frame(self.abas, padding=8)
        self.abas.add(self.aba_consolidar, text="Consolidar")
        self.abas.add(self.aba_empresas, text="Empresas")
        self._montar_consolidar(self.aba_consolidar)
        self._montar_empresas(self.aba_empresas)

    def _faixa(self, pai):
        # Container sempre presente no topo da aba; a faixa aparece dentro dele.
        topo = ttk.Frame(pai)
        topo.pack(fill="x", side="top")
        return tk.Label(topo, bg=AMARELO, anchor="w", padx=8, pady=4, justify="left")

    def _montar_consolidar(self, f):
        self.faixa_consolidar = self._faixa(f)
        self.bloqueio = ttk.Frame(self.faixa_consolidar.master)
        ttk.Label(self.bloqueio, text="Configure a relação de empresas para processar.").pack(
            side="left")
        ttk.Button(self.bloqueio, text="Configurar relação de empresas",
                   command=self.configurar_relacao).pack(side="left", padx=8)

        botoes = ttk.Frame(f)
        botoes.pack(fill="x", side="top")
        self.bt_add = ttk.Button(botoes, text="Adicionar arquivos", command=self.adicionar_arquivos)
        self.bt_pasta = ttk.Button(botoes, text="Adicionar pasta", command=self.adicionar_pasta)
        self.bt_remover = ttk.Button(botoes, text="Remover", command=self.remover)
        self.bt_processar = ttk.Button(botoes, text="Processar", command=self.processar)
        self.bt_cancelar = ttk.Button(botoes, text="Cancelar", command=self.cancelar,
                                      state="disabled")
        for b in (self.bt_add, self.bt_pasta, self.bt_remover):
            b.pack(side="left", padx=(0, 6))
        self.bt_cancelar.pack(side="right")
        self.bt_processar.pack(side="right", padx=6)

        lista = ttk.Frame(f)
        lista.pack(fill="both", expand=True, pady=6)
        self.tree_arquivos = ttk.Treeview(lista, columns=("arquivo", "empresa", "competencia"),
                                          show="headings", height=8)
        for col, titulo, largura in (("arquivo", "Arquivo", 340), ("empresa", "Empresa", 300),
                                     ("competencia", "Competência", 100)):
            self.tree_arquivos.heading(col, text=titulo)
            self.tree_arquivos.column(col, width=largura, anchor="w")
        self.tree_arquivos.tag_configure("erro", foreground="#c00000")
        rolagem = ttk.Scrollbar(lista, orient="vertical", command=self.tree_arquivos.yview)
        self.tree_arquivos.configure(yscrollcommand=rolagem.set)
        self.tree_arquivos.pack(side="left", fill="both", expand=True)
        rolagem.pack(side="right", fill="y")

        progresso = ttk.Frame(f)
        progresso.pack(fill="x")
        self.barra = ttk.Progressbar(progresso, mode="determinate")
        self.barra.pack(fill="x")
        self.texto_progresso = tk.StringVar(value="")
        ttk.Label(progresso, textvariable=self.texto_progresso).pack(anchor="w")

        resumo = ttk.LabelFrame(f, text="Resumo", padding=6)
        resumo.pack(fill="both", expand=True, pady=(6, 0))
        self.resumo_topo = tk.StringVar(value="")
        self.resumo_numeros = tk.StringVar(value="")
        self.resumo_origem = tk.StringVar(value="")
        ttk.Label(resumo, textvariable=self.resumo_topo, font=("Segoe UI", 11, "bold")).pack(
            anchor="w")
        ttk.Label(resumo, textvariable=self.resumo_numeros).pack(anchor="w")
        ttk.Label(resumo, textvariable=self.resumo_origem).pack(anchor="w")
        self.tree_resumo = ttk.Treeview(resumo, columns=("arquivo", "tipo", "mensagem"),
                                        show="headings", height=5)
        for col, titulo, largura in (("arquivo", "Arquivo", 220), ("tipo", "Tipo", 80),
                                     ("mensagem", "Motivo (clique duplo para ver tudo)", 520)):
            self.tree_resumo.heading(col, text=titulo)
            self.tree_resumo.column(col, width=largura, anchor="w")
        self.tree_resumo.tag_configure("erro", foreground="#c00000")
        self.tree_resumo.tag_configure("destaque", foreground="#9c5700")
        self.tree_resumo.bind("<Double-1>", self._detalhe_resumo)
        self.tree_resumo.pack(fill="both", expand=True, pady=4)

        salvar = ttk.Frame(resumo)
        salvar.pack(fill="x")
        self.bt_salvar = ttk.Button(salvar, text="Salvar como...", command=self.salvar_como,
                                    state="disabled")
        self.bt_abrir_arquivo = ttk.Button(salvar, text="Abrir arquivo",
                                           command=self.abrir_arquivo, state="disabled")
        self.bt_abrir_pasta = ttk.Button(salvar, text="Abrir pasta", command=self.abrir_pasta,
                                         state="disabled")
        self.bt_salvar.pack(side="left")
        self.bt_abrir_arquivo.pack(side="left", padx=6)
        self.bt_abrir_pasta.pack(side="left")
        self.texto_salvo = tk.StringVar(value="")
        ttk.Label(salvar, textvariable=self.texto_salvo).pack(side="left", padx=8)

    def _montar_empresas(self, f):
        self.faixa_empresas = self._faixa(f)
        topo = ttk.Frame(f)
        topo.pack(fill="x")
        ttk.Label(topo, text="Buscar:").pack(side="left")
        self.busca = tk.StringVar()
        self.busca.trace_add("write", lambda *a: self._preencher_empresas())
        ttk.Entry(topo, textvariable=self.busca, width=40).pack(side="left", padx=6)

        tabela = ttk.Frame(f)
        tabela.pack(fill="both", expand=True, pady=6)
        self.tree_empresas = ttk.Treeview(tabela, columns=("empresa", "cod_saude", "cod_dental"),
                                          show="headings")
        for col, titulo, largura in (("empresa", "EMPRESA", 420), ("cod_saude", "COD_SAUDE", 140),
                                     ("cod_dental", "COD_DENTAL", 140)):
            self.tree_empresas.heading(col, text=titulo)
            self.tree_empresas.column(col, width=largura, anchor="w")
        self.tree_empresas.bind("<<TreeviewSelect>>", self._selecionar_empresa)
        rolagem = ttk.Scrollbar(tabela, orient="vertical", command=self.tree_empresas.yview)
        self.tree_empresas.configure(yscrollcommand=rolagem.set)
        self.tree_empresas.pack(side="left", fill="both", expand=True)
        rolagem.pack(side="right", fill="y")

        form = ttk.LabelFrame(f, text="Empresa", padding=6)
        form.pack(fill="x")
        self.campo_empresa = tk.StringVar()
        self.campo_cod = tk.StringVar()
        self.campo_dental = tk.StringVar()
        for i, (rotulo, var, largura) in enumerate((("EMPRESA", self.campo_empresa, 50),
                                                    ("COD_SAUDE", self.campo_cod, 16),
                                                    ("COD_DENTAL", self.campo_dental, 16))):
            ttk.Label(form, text=rotulo).grid(row=0, column=2 * i, sticky="w", padx=(0, 4))
            ttk.Entry(form, textvariable=var, width=largura).grid(row=0, column=2 * i + 1,
                                                                  sticky="w", padx=(0, 12))
        self.campo_cod.trace_add("write", lambda *a: self._validar_formulario())
        self.campo_empresa.trace_add("write", lambda *a: self._validar_formulario())
        self.texto_validacao = tk.StringVar(value="")
        ttk.Label(form, textvariable=self.texto_validacao, foreground="#c00000").grid(
            row=1, column=0, columnspan=6, sticky="w", pady=(4, 0))

        botoes = ttk.Frame(f)
        botoes.pack(fill="x", pady=6)
        self.bt_incluir = ttk.Button(botoes, text="Incluir", command=self.incluir_empresa)
        self.bt_editar = ttk.Button(botoes, text="Salvar edição", command=self.editar_empresa)
        self.bt_excluir = ttk.Button(botoes, text="Excluir", command=self.excluir_empresa)
        self.bt_reprocessar = ttk.Button(botoes, text="Processar de novo",
                                         command=self.processar_de_novo)
        for b in (self.bt_incluir, self.bt_editar, self.bt_excluir):
            b.pack(side="left", padx=(0, 6))

        self.rodape = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.rodape, foreground="#555555").pack(anchor="w")

    # --- relacao de empresas -----------------------------------------------------

    @property
    def caminho_relacao(self):
        valor = self.cfg.get("arquivo_empresas")
        return Path(valor) if valor else None

    def carregar_relacao(self, mostrar_erro=True) -> bool:
        caminho = self.caminho_relacao
        try:
            if caminho is None:
                # Sem caminho configurado so serve o cache (se houver).
                caminho = self.pasta_dados / "sem_relacao_configurada.csv"
            self.relacao = ec.ler(caminho, self.pasta_dados)
            self.mapa = ec.mapa(self.relacao)
        except ErroValidacao as e:
            if self.relacao is None:
                self.mapa = {}
            if mostrar_erro:
                self.dialogos.erro("Relação de empresas", str(e))
            self._atualizar_estado_relacao()
            return False
        self.itens = atualizar_empresas(self.itens, self.mapa)
        self._preencher_arquivos()
        self._preencher_empresas()
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
        edicao = "normal" if (r is not None and r.origem == "rede"
                              and self.caminho_relacao is not None) else "disabled"
        for b in (self.bt_incluir, self.bt_editar, self.bt_excluir):
            b.config(state=edicao)
        self._atualizar_botoes()
        if r is None:
            self.rodape.set("Relação de empresas não configurada.")
        else:
            origem = "rede" if r.origem == "rede" else f"cópia local de {_data(r.data_cache)}"
            sem_rede = "" if edicao == "normal" else "  (edição bloqueada sem rede)"
            self.rodape.set(f"Arquivo: {self.caminho_relacao or '-'}  |  lido em "
                            f"{_data(r.lida_em)}  |  origem: {origem}{sem_rede}")

    def configurar_relacao(self):
        escolha = self.dialogos.primeiro_uso()
        if escolha is None:
            return
        try:
            if escolha:
                caminho = self.dialogos.escolher_relacao()
                if not caminho:
                    return
                ec.ler(caminho, self.pasta_dados)  # valida que da para ler
            else:
                origem = self.dialogos.escolher_csv_origem()
                if not origem:
                    return
                caminho = self.dialogos.escolher_destino_relacao()
                if not caminho:
                    return
                ec.criar_a_partir_de_csv(origem, caminho)
        except ErroValidacao as e:
            self.dialogos.erro("Relação de empresas", str(e))
            return
        self.cfg["arquivo_empresas"] = str(caminho)
        salvar_config(self.pasta_dados, self.cfg)
        self.carregar_relacao()

    def sobre(self):
        self.dialogos.informar("Sobre", f"{NOME_APP}\nVersão {VERSAO}")

    # --- lista de arquivos --------------------------------------------------------

    def _preencher_arquivos(self):
        t = self.tree_arquivos
        t.delete(*t.get_children())
        for i, item in enumerate(self.itens):
            empresa = item.empresa or (item.erro or "")
            t.insert("", "end", iid=str(i), tags=("erro",) if item.erro else (),
                     values=(item.caminho.name, empresa, texto_competencia(item.competencia)))
        self._atualizar_botoes()

    def adicionar_caminhos(self, caminhos):
        self.itens = adicionar(self.itens, caminhos, self.mapa)
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
        for b in (self.bt_add, self.bt_pasta, self.bt_remover):
            b.config(state=normal)
        pode = not ocupado and bool(self.itens) and self.relacao is not None
        self.bt_processar.config(state="normal" if pode else "disabled")
        self.bt_cancelar.config(state="normal" if self.tipo_trabalho == "processar" else "disabled")
        self.bt_salvar.config(
            state="normal" if (self.resultado is not None and not ocupado) else "disabled")
        aberto = "normal" if (self.ultimo_salvo is not None and not ocupado) else "disabled"
        self.bt_abrir_arquivo.config(state=aberto)
        self.bt_abrir_pasta.config(state=aberto)

    def processar(self):
        if self.ocupado or not self.itens:
            return
        # A relacao e relida logo antes de processar.
        if not self.carregar_relacao():
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
        self._iniciar("processar", lambda p, c: self.processar_fn(caminhos, mapa, p, c))

    def processar_de_novo(self):
        self.abas.select(self.aba_consolidar)
        self.processar()

    def cancelar(self):
        if self.tipo_trabalho == "processar":
            self.trabalho.cancelar()
            self.texto_progresso.set("Cancelando...")
            self.bt_cancelar.config(state="disabled")

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
                self.texto_salvo.set(f"Salvo em {self.ultimo_salvo}")
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
        self.resumo_numeros.set(
            f"Registros: {r['registros']}  |  Pendências: {r['pendencias']}  |  "
            f"Datas inválidas: {r['datas_invalidas']}  |  "
            f"Divergências de datas: {r['divergencias_datas']}")
        origem = self.origem_processada
        self.resumo_origem.set(
            "Relação de empresas: rede" if origem.origem == "rede"
            else f"Relação de empresas: cópia local de {_data(origem.data_cache)}")
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

    def _linha_resumo(self, iid, arquivo, tipo, mensagem, tag):
        iid = str(iid)
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
        atual = self.relacao
        try:
            atual = ec.ler(self.caminho_relacao or self.pasta_dados / "sem_relacao.csv",
                           self.pasta_dados)
        except ErroValidacao:
            pass
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
        self._iniciar("gravar", lambda p, c: self.gravar_fn(
            resultado, destino, caminhos, arquivo_empresas=empresas, progresso=p))

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
        for i, l in enumerate(self.relacao.linhas):
            if filtro and filtro not in f"{l.empresa} {l.cod_saude} {l.cod_dental}".casefold():
                continue
            t.insert("", "end", iid=str(i), values=(l.empresa, l.cod_saude, l.cod_dental))

    def _linha_selecionada(self):
        sel = self.tree_empresas.selection()
        if not sel or self.relacao is None:
            return None
        return self.relacao.linhas[int(sel[0])]

    def _selecionar_empresa(self, _evento=None):
        linha = self._linha_selecionada()
        if linha is not None:
            self.campo_empresa.set(linha.empresa)
            self.campo_cod.set(linha.cod_saude)
            self.campo_dental.set(linha.cod_dental)

    def _formulario(self):
        return ec.Linha(self.campo_empresa.get(), self.campo_cod.get(), self.campo_dental.get())

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
        self.texto_validacao.set(mensagem)
        return mensagem

    def _aplicar(self, alteracao) -> bool:
        try:
            self.relacao = ec.aplicar(self.caminho_relacao, self.relacao, alteracao,
                                      self.usuario, self.maquina, pasta_dados=self.pasta_dados)
        except ec.ConflitoEdicao as e:
            self.dialogos.avisar("Relação de empresas", str(e))
            self.carregar_relacao()
            return False
        except ErroValidacao as e:
            # Inclui EmUso: a edicao continua no formulario para tentar de novo.
            self.dialogos.erro("Relação de empresas", str(e))
            return False
        self.mapa = ec.mapa(self.relacao)
        self.itens = atualizar_empresas(self.itens, self.mapa)
        self._preencher_arquivos()
        self._preencher_empresas()
        self._atualizar_estado_relacao()
        return True

    def incluir_empresa(self):
        if self._validar_formulario() and "já cadastrado" not in self.texto_validacao.get():
            self.dialogos.avisar("Relação de empresas", self.texto_validacao.get())
            return
        if self._aplicar(ec.incluir(self._formulario())):
            self._limpar_formulario()
            if self.itens:
                self.bt_reprocessar.pack(side="right")

    def editar_empresa(self):
        antiga = self._linha_selecionada()
        if antiga is None:
            self.dialogos.avisar("Relação de empresas", "Selecione a empresa a editar.")
            return
        if self._aplicar(ec.editar(antiga, self._formulario())):
            self._limpar_formulario()

    def excluir_empresa(self):
        antiga = self._linha_selecionada()
        if antiga is None:
            self.dialogos.avisar("Relação de empresas", "Selecione a empresa a excluir.")
            return
        if not self.dialogos.perguntar(
                "Excluir empresa",
                f"Excluir {antiga.empresa} (COD_SAUDE {antiga.cod_saude}) da relação?"):
            return
        if self._aplicar(ec.excluir(antiga)):
            self._limpar_formulario()

    def _limpar_formulario(self):
        self.tree_empresas.selection_remove(*self.tree_empresas.selection())
        for var in (self.campo_empresa, self.campo_cod, self.campo_dental):
            var.set("")
        self.texto_validacao.set("")

    # --- fechar ---------------------------------------------------------------------

    def fechar(self):
        if self.tipo_trabalho == "processar":
            if not self.dialogos.perguntar("Fechar",
                                           "O processamento está em andamento. Cancelar e fechar?"):
                return
            self.trabalho.cancelar()
        elif self.tipo_trabalho == "gravar":
            self.dialogos.avisar("Fechar", "Aguarde terminar de gravar o arquivo de saída.")
            return
        elif self.resultado is not None and not self.salvo:
            if not self.dialogos.perguntar("Fechar",
                                           "O resultado do processamento não foi salvo. "
                                           "Fechar mesmo assim?"):
                return
        self.raiz.destroy()
