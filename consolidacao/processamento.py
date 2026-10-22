# Processamento de um arquivo completo e da pasta (secoes 6 e 7 do TRATAMENTO.R).
import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import pandas as pd

from .cabecalhos import normalizar_nome, unificar_colunas
from .config import CODIGO_UNICO
from .empresas import empresa_da_referencia
from .erros import ErroValidacao
from .leitura import ler_dependente, ler_suplementar, ler_titular

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
    return v is not None and str(v).strip() != ""


def _coluna(df: pd.DataFrame, nome: str, valores) -> None:
    df[nome] = pd.Series(list(valores), index=df.index, dtype=object)


def _abas_do_arquivo(caminho: Path) -> list:
    wb = openpyxl.load_workbook(caminho, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _exatamente_uma(abas, padrao, descricao):
    encontradas = [a for a in abas if re.search(padrao, normalizar_nome(a))]
    if len(encontradas) != 1:
        raise ErroValidacao(
            f"Esperada exatamente uma aba {descricao}; encontrada(s) {len(encontradas)}."
        )
    return encontradas[0]


def processar_arquivo(caminho: Path, mapa: dict) -> ResultadoArquivo:
    caminho = Path(caminho)
    arquivo = caminho.name
    referencia = caminho.stem
    observacoes_datas, divergencias_datas = [], []

    try:
        codigo_empresa, empresa = empresa_da_referencia(referencia, mapa)

        abas = _abas_do_arquivo(caminho)
        aba_titular = _exatamente_uma(abas, "TITULAR", "de TITULAR")
        aba_dependente = _exatamente_uma(abas, "DEPENDENT", "de DEPENDENTE")
        aba_suplementar = _exatamente_uma(
            abas, "POS.*CADASTRAL.*NOME.*M.*E", "POS. CADASTRAL (NOME M E)"
        )

        # As observacoes so entram depois que cada leitor termina, como no R.
        titular_info = ler_titular(caminho, aba_titular, referencia)
        observacoes_datas += titular_info.observacoes_datas
        divergencias_datas += titular_info.divergencias_datas
        dependente_info = ler_dependente(caminho, aba_dependente, referencia)
        observacoes_datas += dependente_info.observacoes_datas
        divergencias_datas += dependente_info.divergencias_datas
        suplementar_info = ler_suplementar(caminho, aba_suplementar, referencia)
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
        duplicados = list(dict.fromkeys(
            c for i, c in enumerate(certificados) if c in certificados[:i]
        ))
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
        titulares_sem_suplemento = sum(c not in chaves_sup for c in certificados)
        suplementares_sem_titular = sum(c not in certificados for c in chaves_sup)
        titulares_sem_cartao = sum(not _preenchido(v) for v in titular["NUMERO DO CARTAO"])
        titulares_sem_nome_mae = sum(not _preenchido(v) for v in titular["NOME DA MAE"])

        avisos = []
        if titular_info.avisos:
            avisos.append("TITULAR ausente: " + ", ".join(titular_info.avisos))
        if dependente_info.avisos:
            avisos.append("DEPENDENTE ausente: " + ", ".join(dependente_info.avisos))
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
            mensagem=str(e),
            observacoes_datas=observacoes_datas,
            divergencias_datas=divergencias_datas,
        )
