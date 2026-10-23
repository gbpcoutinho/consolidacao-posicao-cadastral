# Processamento de um arquivo completo e da pasta (secoes 6 e 7 do TRATAMENTO.R).
import re
import zipfile
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable

import openpyxl
import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException

from .cabecalhos import chave_locale, normalizar_nome, trimws, unificar_colunas
from .config import CODIGO_UNICO, NOMES_SAIDA_LEGADOS, ORDEM_INICIAL, nome_saida
from .datas import resumir_observacoes_datas
from .empresas import empresa_da_referencia
from .erros import ErroValidacao, ProcessamentoCancelado
from .leitura import abrir_planilha, ler_dependente, ler_suplementar, ler_titular

COLUNAS_PENDENCIAS = [
    "REFERENCIA", "EMPRESA", "NUMERO DA SUBFATURA", "NUMERO DO CERTIFICADO",
    "CODIGO DO DEPENDENTE", "NOME DEPENDENTE", "NOME DO SEGURADO", "PLANO", "MOTIVO",
]


@dataclass
class ResultadoArquivo:
    arquivo: str
    referencia: str
    status: str
    mensagem: str
    avisos: str | None = None
    observacoes_datas: list = field(default_factory=list)
    divergencias_datas: list = field(default_factory=list)
    final: pd.DataFrame | None = None
    pendencias: pd.DataFrame | None = None
    n_titulares: int | None = None
    n_dependentes: int | None = None
    codigo_empresa: str | None = None
    empresa: str | None = None


def _bind_rows(dfs: list) -> pd.DataFrame:
    """Como dplyr::bind_rows: uniao das colunas na ordem de aparicao, ausente = None."""
    colunas = list(dict.fromkeys(c for df in dfs for c in df.columns))
    partes = []
    for df in dfs:
        parte = {c: (df[c].tolist() if c in df.columns else [None] * len(df)) for c in colunas}
        partes.append(parte)
    dados = {c: [v for p in partes for v in p[c]] for c in colunas}
    return pd.DataFrame({c: pd.Series(v, dtype=object) for c, v in dados.items()}, columns=colunas)


def _preenchido(v) -> bool:
    return v is not None and trimws(str(v)) != ""


def _coluna(df: pd.DataFrame, nome: str, valores) -> None:
    df[nome] = pd.Series(list(valores), index=df.index, dtype=object)


def _exatamente_uma(abas, padrao, descricao):
    encontradas = [a for a in abas if re.search(padrao, normalizar_nome(a))]
    if len(encontradas) != 1:
        raise ErroValidacao(
            f"Esperada exatamente uma aba {descricao}; encontrada(s) {len(encontradas)}."
        )
    return encontradas[0]


def _mensagem_erro(e: Exception) -> str:
    # Erros de abertura viram texto para o usuario; o resto segue como no R.
    if isinstance(e, (zipfile.BadZipFile, InvalidFileException)):
        return "Nao foi possivel abrir o arquivo: arquivo corrompido ou nao e .xlsx."
    if isinstance(e, PermissionError):
        return "Nao foi possivel abrir o arquivo: arquivo em uso ou sem permissao."
    return str(e)


def processar_arquivo(caminho: Path, mapa: dict) -> ResultadoArquivo:
    caminho = Path(caminho)
    arquivo = caminho.name
    referencia = caminho.stem
    observacoes_datas, divergencias_datas = [], []
    wb = None

    try:
        codigo_empresa, empresa = empresa_da_referencia(referencia, mapa)

        # Abre o arquivo uma unica vez para todas as abas.
        wb = abrir_planilha(caminho)
        abas = list(wb.sheetnames)
        aba_titular = _exatamente_uma(abas, "TITULAR", "de TITULAR")
        aba_dependente = _exatamente_uma(abas, "DEPENDENT", "de DEPENDENTE")
        aba_suplementar = _exatamente_uma(
            abas, "POS.*CADASTRAL.*NOME.*M.*E", "POS. CADASTRAL (NOME M E)"
        )

        # As observacoes so entram depois que cada leitor termina, como no R.
        titular_info = ler_titular(caminho, aba_titular, referencia, wb=wb)
        observacoes_datas += titular_info.observacoes_datas
        divergencias_datas += titular_info.divergencias_datas
        dependente_info = ler_dependente(caminho, aba_dependente, referencia, wb=wb)
        observacoes_datas += dependente_info.observacoes_datas
        divergencias_datas += dependente_info.divergencias_datas
        suplementar_info = ler_suplementar(caminho, aba_suplementar, referencia, wb=wb)
        observacoes_datas += suplementar_info.observacoes_datas
        divergencias_datas += suplementar_info.divergencias_datas

        titular = titular_info.dados
        dependente = dependente_info.dados
        suplementar = suplementar_info.dados

        # Reune os nomes equivalentes da aba dependente nas colunas finais comuns.
        ctx_dep = f"DEPENDENTE: {referencia}"
        dependente = unificar_colunas(
            dependente, "NUMERO DO CARTAO", ["NUMERO DO CARTAO DO DEPENDENTE"], ctx_dep
        )
        dependente = unificar_colunas(
            dependente, "NOME DA MAE", ["NOME DA MAE DO DEPENDENTE"], ctx_dep
        )

        certificados = titular[".CHAVE_CERTIFICADO"].tolist()
        vistos, duplicados = set(), {}
        for c in certificados:
            if c in vistos:
                duplicados[c] = None
            vistos.add(c)
        duplicados = list(duplicados)
        if duplicados:
            raise ErroValidacao(
                "Certificado(s) duplicado(s) na aba TITULAR: " + ", ".join(duplicados)
            )

        # left_join com a aba suplementar (chaves unicas dos dois lados).
        sup = {
            r[".CHAVE_CERTIFICADO"]: r
            for r in suplementar.to_dict("records")
        }
        titular = titular.copy()
        _coluna(titular, ".CARTAO_SUP",
                [sup.get(c, {}).get(".CARTAO_SUP") for c in certificados])
        _coluna(titular, ".NOME_MAE_SUP",
                [sup.get(c, {}).get(".NOME_MAE_SUP") for c in certificados])

        # Mantem valores ja presentes no titular e completa com a aba suplementar.
        ctx_tit = f"TITULAR: {referencia}"
        titular = unificar_colunas(titular, "NUMERO DO CARTAO", ["CARTAO", ".CARTAO_SUP"], ctx_tit)
        titular = unificar_colunas(titular, "NOME DA MAE", [".NOME_MAE_SUP"], ctx_tit)

        # Empresa e codigo da referencia sao atribuidos antes da combinacao das abas.
        _coluna(titular, "EMPRESA", [empresa] * len(titular))
        dependente = dependente.copy()
        _coluna(dependente, "EMPRESA", [empresa] * len(dependente))

        lookup = {
            r[".CHAVE_CERTIFICADO"]: r for r in titular.to_dict("records")
        }
        # na_matches = "never": dependente sem certificado nunca encontra titular.
        encontrados = [
            lookup.get(c) if c is not None else None
            for c in dependente[".CHAVE_CERTIFICADO"].tolist()
        ]
        subf_dependente = dependente[".SUBF_DEPENDENTE"].tolist()
        _coluna(dependente, "NOME DO SEGURADO",
                [t["NOME DO SEGURADO"] if t else None for t in encontrados])
        _coluna(dependente, "PLANO", [t["PLANO"] if t else None for t in encontrados])
        _coluna(dependente, "NUMERO DA SUBFATURA", [
            t["NUMERO DA SUBFATURA"] if t and t["NUMERO DA SUBFATURA"] is not None else s
            for t, s in zip(encontrados, subf_dependente)
        ])

        sem_titular = [t is None for t in encontrados]
        pendencias = dependente.loc[sem_titular, COLUNAS_PENDENCIAS[:-1]].copy()
        _coluna(pendencias, "MOTIVO",
                ["CERTIFICADO SEM TITULAR CORRESPONDENTE"] * len(pendencias))
        pendencias = pendencias.reset_index(drop=True)

        dependente_final = dependente.drop(columns=[".SUBF_DEPENDENTE"])

        # Campos especificos do titular permanecem vazios nas linhas dependentes.
        final = _bind_rows([titular, dependente_final]).rename(
            columns={"NUMERO DA SUBFATURA": "SubF"}
        )
        _coluna(final, CODIGO_UNICO, [
            f"{cert}{cod}" if _preenchido(cert) and _preenchido(cod) else None
            for cert, cod in zip(
                final["NUMERO DO CERTIFICADO"].tolist(),
                final["CODIGO DO DEPENDENTE"].tolist(),
            )
        ])

        chaves_sup = suplementar[".CHAVE_CERTIFICADO"].tolist()
        conjunto_sup, conjunto_tit = set(chaves_sup), set(certificados)
        titulares_sem_suplemento = sum(c not in conjunto_sup for c in certificados)
        suplementares_sem_titular = sum(c not in conjunto_tit for c in chaves_sup)
        titulares_sem_cartao = sum(not _preenchido(v) for v in titular["NUMERO DO CARTAO"])
        titulares_sem_nome_mae = sum(not _preenchido(v) for v in titular["NOME DA MAE"])

        avisos = []
        if titular_info.avisos:
            avisos.append("TITULAR ausente: " + ", ".join(titular_info.avisos))
        if dependente_info.avisos:
            avisos.append("DEPENDENTE ausente: " + ", ".join(dependente_info.avisos))
        avisos += dependente_info.avisos_estrutura
        if titulares_sem_suplemento > 0:
            avisos.append(f"Titulares sem registro na aba suplementar: {titulares_sem_suplemento}")
        if suplementares_sem_titular > 0:
            avisos.append(
                f"Certificados suplementares sem titular correspondente: {suplementares_sem_titular}"
            )
        if titulares_sem_cartao > 0:
            avisos.append(f"Titulares sem CARTAO apos a juncao: {titulares_sem_cartao}")
        if titulares_sem_nome_mae > 0:
            avisos.append(f"Titulares sem NOME DA MAE apos a juncao: {titulares_sem_nome_mae}")

        return ResultadoArquivo(
            arquivo=arquivo,
            referencia=referencia,
            status="OK",
            mensagem="Processado com sucesso.",
            avisos=" | ".join(avisos) if avisos else None,
            observacoes_datas=observacoes_datas,
            divergencias_datas=divergencias_datas,
            final=final,
            pendencias=pendencias,
            n_titulares=len(titular),
            n_dependentes=len(dependente),
            codigo_empresa=codigo_empresa,
            empresa=empresa,
        )
    except Exception as e:  # tryCatch(error = ...) do R captura qualquer erro
        return ResultadoArquivo(
            arquivo=arquivo,
            referencia=referencia,
            status="ERRO",
            mensagem=_mensagem_erro(e),
            observacoes_datas=observacoes_datas,
            divergencias_datas=divergencias_datas,
        )
    finally:
        if wb is not None:
            wb.close()


# ----------------------------------------------------------------------------
# Processamento da pasta e consolidacao
# ----------------------------------------------------------------------------

COLUNAS_BASE_VAZIA = [
    "REFERENCIA", "EMPRESA", "SubF", "NUMERO DO CERTIFICADO", CODIGO_UNICO,
    "TITULARIDADE", "NOME DO SEGURADO", "NUMERO DO CARTAO", "NOME DA MAE",
    "NOME DEPENDENTE", "CODIGO DO DEPENDENTE", "GRAU PARENTESCO", "PLANO",
]

COLUNAS_LOG = [
    "ARQUIVO", "REFERENCIA", "CODIGO_EMPRESA", "EMPRESA", "STATUS", "TITULARES",
    "DEPENDENTES", "AVISOS_ESTRUTURA", "DATAS_INVALIDAS", "OBSERVACOES_DATAS",
    "QTD_DIVERGENCIAS_DATAS", "DIVERGENCIAS_DATAS", "MENSAGEM",
    "OBSERVACAO_IDENTIFICADORES",
]

OBSERVACAO_IDENTIFICADORES = (
    "Identificadores lidos como texto; zeros ausentes na origem nao foram reconstruidos."
)


@dataclass
class Consolidado:
    consolidado: pd.DataFrame
    pendencias: pd.DataFrame
    log: pd.DataFrame


def listar_arquivos_entrada(pasta: Path, hoje: date) -> list:
    pasta = Path(pasta)
    if not pasta.is_dir():
        raise ErroValidacao(f"A pasta de entrada nao existe: {pasta}")

    excluir = {n.lower() for n in [nome_saida(hoje), *NOMES_SAIDA_LEGADOS]}
    arquivos = [
        p for p in pasta.iterdir()
        if p.is_file()
        and p.suffix.lower() == ".xlsx"
        and not p.name.startswith("~$")
        and p.name.lower() not in excluir
        # Deixa de fora consolidacoes mensais anteriores, mesmo se o nome de saida mudar.
        and not re.match(r"^PC_CONSOLIDADO_.*\.xlsx$", p.name, flags=re.IGNORECASE)
    ]
    arquivos.sort(key=lambda p: chave_locale(p.name))

    if not arquivos:
        raise ErroValidacao(f"Nenhum arquivo .xlsx de entrada encontrado em: {pasta}")
    return arquivos


def _tipo(valores) -> str | None:
    for v in valores:
        if isinstance(v, date):
            return "date"
        if v is not None:
            return "character"
    return None


def _tipos(df: pd.DataFrame) -> dict:
    return {c: _tipo(df[c].tolist()) for c in df.columns}


def _incompatibilidade(tipos_acumulados: dict, df: pd.DataFrame) -> str | None:
    for coluna, tipo in _tipos(df).items():
        anterior = tipos_acumulados.get(coluna)
        if tipo and anterior and tipo != anterior:
            return f"Can't combine `{coluna}` <{anterior}> and `{coluna}` <{tipo}>."
    return None


def _chave_ordenacao(v):
    # arrange() do dplyr compara texto no locale C; NA vai para o fim.
    return (1, b"") if v is None else (0, str(v).encode("utf-8"))


def _ordenar(df: pd.DataFrame, colunas: list) -> pd.DataFrame:
    linhas = list(range(len(df)))
    valores = {c: df[c].tolist() for c in colunas}
    linhas.sort(key=lambda i: tuple(_chave_ordenacao(valores[c][i]) for c in colunas))
    return df.iloc[linhas].reset_index(drop=True)


def avisar_progresso(progresso, i: int, total: int, nome: str) -> None:
    """Chama o callback da interface; um defeito nele nao derruba o processamento.
    ProcessamentoCancelado levantado pelo callback continua cancelando."""
    if not progresso:
        return
    try:
        progresso(i, total, nome)
    except ProcessamentoCancelado:
        raise
    except Exception:
        pass


def processar_arquivos(
    caminhos: list,
    mapa: dict,
    progresso: Callable[[int, int, str], None] | None = None,
    cancelar: Callable[[], bool] | None = None,
) -> Consolidado:
    """progresso(i, total, nome) e chamado antes de cada arquivo (i comeca em 1).
    cancelar() e consultado antes de cada arquivo; True interrompe com
    ProcessamentoCancelado."""
    resultados = []
    total = len(caminhos)
    for i, caminho in enumerate(caminhos, start=1):
        if cancelar and cancelar():
            raise ProcessamentoCancelado("Processamento cancelado pelo usuario.")
        avisar_progresso(progresso, i, total, Path(caminho).name)
        resultados.append(processar_arquivo(caminho, mapa))

    # Adiciona um arquivo por vez para registrar no log uma eventual incompatibilidade.
    aceitos = []
    tipos_acumulados = {}
    for r in resultados:
        if r.status != "OK":
            continue
        problema = _incompatibilidade(tipos_acumulados, r.final)
        if problema:
            r.status = "ERRO"
            r.mensagem = (
                "Incompatibilidade ao consolidar este arquivo com os anteriores: " + problema
            )
            r.final = r.pendencias = None
            r.n_titulares = r.n_dependentes = None
            continue
        for coluna, tipo in _tipos(r.final).items():
            if tipo:
                tipos_acumulados.setdefault(coluna, tipo)
        aceitos.append(r)

    if aceitos:
        consolidado = _bind_rows([r.final for r in aceitos])
        pendencias = _bind_rows([r.pendencias for r in aceitos])
    else:
        consolidado = pd.DataFrame(columns=COLUNAS_BASE_VAZIA, dtype=object)
        pendencias = pd.DataFrame(columns=COLUNAS_PENDENCIAS, dtype=object)

    primeiras = [c for c in ORDEM_INICIAL if c in consolidado.columns]
    consolidado = consolidado[primeiras + [c for c in consolidado.columns if c not in primeiras]]
    consolidado = _ordenar(consolidado, ["REFERENCIA", "SubF", "NUMERO DO CERTIFICADO"])

    log = pd.DataFrame(
        [
            {
                "ARQUIVO": r.arquivo,
                "REFERENCIA": r.referencia,
                "CODIGO_EMPRESA": r.codigo_empresa,
                "EMPRESA": r.empresa,
                "STATUS": r.status,
                "TITULARES": r.n_titulares,
                "DEPENDENTES": r.n_dependentes,
                "AVISOS_ESTRUTURA": r.avisos,
                "DATAS_INVALIDAS": len(r.observacoes_datas),
                "OBSERVACOES_DATAS": resumir_observacoes_datas(r.observacoes_datas),
                "QTD_DIVERGENCIAS_DATAS": len(r.divergencias_datas),
                "DIVERGENCIAS_DATAS": resumir_observacoes_datas(r.divergencias_datas),
                "MENSAGEM": r.mensagem,
                "OBSERVACAO_IDENTIFICADORES": OBSERVACAO_IDENTIFICADORES,
            }
            for r in resultados
        ],
        columns=COLUNAS_LOG,
        dtype=object,
    )

    return Consolidado(consolidado=consolidado, pendencias=pendencias, log=log)
