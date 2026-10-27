"""Lista de arquivos da aba Consolidar e checagens de integridade (sem Tk).

- adicionar / adicionar_pasta: so .xlsx, sem temporarios do Excel (~$) nem
  consolidacoes anteriores (PC_CONSOLIDADO_*), sem repetir arquivo.
- competencia: MM.AAAA no nome (ex.: "PC90001 10.2026.xlsx" -> (2026, 10)).
- Retrato: o que foi processado (tamanho e data de cada entrada, hash da
  relacao), para avisar no "Salvar como" se algo mudou desde entao.
"""
import os
import re
from collections import Counter
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from consolidacao.cabecalhos import chave_locale
from consolidacao.config import nome_saida
from consolidacao.empresas import empresa_da_referencia
from consolidacao.erros import ErroValidacao

_COMPETENCIA = re.compile(r"(?<!\d)(\d{2})\.(\d{4})(?!\d)")
_CONSOLIDADO = re.compile(r"^PC_CONSOLIDADO_", re.IGNORECASE)


@dataclass
class ItemArquivo:
    caminho: Path
    empresa: str | None
    codigo: str | None
    competencia: tuple | None  # (ano, mes)
    erro: str | None


@dataclass
class Retrato:
    entradas: dict  # chave normalizada -> (tamanho, data de modificacao)
    hash_relacao: str
    nomes: dict = None  # chave normalizada -> nome do arquivo, para as mensagens


def _chave(caminho: Path) -> str:
    return os.path.normcase(str(Path(caminho).resolve()))


def _aceito(caminho: Path) -> bool:
    nome = caminho.name
    return (
        caminho.suffix.lower() == ".xlsx"
        and not nome.startswith("~$")
        and not _CONSOLIDADO.match(nome)
    )


def competencia(nome: str) -> tuple | None:
    for m in _COMPETENCIA.finditer(Path(nome).stem):
        mes, ano = int(m.group(1)), int(m.group(2))
        if 1 <= mes <= 12:
            return ano, mes
    return None


def _item(caminho: Path, mapa: dict) -> ItemArquivo:
    try:
        codigo, empresa = empresa_da_referencia(caminho.stem, mapa)
        erro = None
    except ErroValidacao as e:
        codigo, empresa, erro = None, None, str(e)
    return ItemArquivo(caminho, empresa, codigo, competencia(caminho.name), erro)


def adicionar(itens: list, caminhos: list, mapa: dict) -> list:
    novos = list(itens)
    vistos = {_chave(i.caminho) for i in novos}
    for caminho in caminhos:
        caminho = Path(caminho)
        if not _aceito(caminho):
            continue
        k = _chave(caminho)
        if k in vistos:
            continue
        vistos.add(k)
        novos.append(_item(caminho, mapa))
    return novos


def adicionar_pasta(itens: list, pasta: Path, mapa: dict) -> list:
    arquivos = [p for p in Path(pasta).iterdir() if p.is_file()]
    arquivos.sort(key=lambda p: chave_locale(p.name))
    return adicionar(itens, arquivos, mapa)


def atualizar_empresas(itens: list, mapa: dict) -> list:
    """Refaz a coluna Empresa depois que a relacao mudou."""
    return [_item(i.caminho, mapa) for i in itens]


def texto_competencia(c: tuple | None) -> str:
    return f"{c[1]:02d}/{c[0]}" if c else ""


def alertas_pre_processamento(itens: list) -> list:
    alertas = []
    competencias = sorted({i.competencia for i in itens if i.competencia})
    if len(competencias) > 1:
        alertas.append(
            "Arquivos de competencias diferentes: "
            + ", ".join(texto_competencia(c) for c in competencias) + "."
        )
    sem = [i.caminho.name for i in itens if not i.competencia]
    if sem:
        alertas.append("Arquivo(s) sem competencia (MM.AAAA) no nome: " + ", ".join(sem) + ".")
    for i in itens:
        if i.erro:
            alertas.append(f"{i.caminho.name}: {i.erro}")
    repetidos = [c for c, n in Counter(i.codigo for i in itens if i.codigo).items() if n > 1]
    for codigo in repetidos:
        nomes = [i.caminho.name for i in itens if i.codigo == codigo]
        alertas.append(
            f"Empresa {codigo} em mais de um arquivo (registros duplicados?): "
            + ", ".join(nomes) + "."
        )
    return alertas


def nome_saida_sugerido(itens: list, hoje: date) -> str:
    competencias = {i.competencia for i in itens if i.competencia}
    if len(competencias) == 1 and all(i.competencia for i in itens):
        ano, mes = competencias.pop()
        return nome_saida(date(ano, mes, 1))
    return nome_saida(hoje)


def _estado(caminho: Path):
    try:
        st = Path(caminho).stat()
    except OSError:
        return None
    return st.st_size, st.st_mtime


def retratar(itens: list, relacao) -> Retrato:
    return Retrato(
        entradas={_chave(i.caminho): _estado(i.caminho) for i in itens},
        hash_relacao=relacao.hash,
        nomes={_chave(i.caminho): i.caminho.name for i in itens},
    )


def mudancas(retrato: Retrato, itens: list, relacao_atual) -> list:
    """O que mudou entre o processamento e agora (vazio = nada)."""
    avisos = []
    atuais = {_chave(i.caminho): i.caminho.name for i in itens}
    if set(atuais) != set(retrato.entradas):
        avisos.append("A lista de arquivos mudou desde o processamento.")
    for k, estado in retrato.entradas.items():
        nome = (retrato.nomes or {}).get(k, k)
        agora = _estado(Path(k))
        if agora is None:
            avisos.append(f"O arquivo {nome} nao foi encontrado.")
        elif agora != estado:
            avisos.append(f"O arquivo {nome} foi alterado desde o processamento.")
    if relacao_atual is not None and relacao_atual.hash != retrato.hash_relacao:
        avisos.append("A relacao de empresas mudou desde o processamento.")
    return avisos
