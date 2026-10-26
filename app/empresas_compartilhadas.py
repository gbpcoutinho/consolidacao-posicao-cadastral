"""Relacao de empresas compartilhada na rede, com copia local (cache).

Formato do arquivo: o mesmo COD_EMPRESAS.csv de hoje (";", colunas
EMPRESA;COD_SAUDE;COD_DENTAL, UTF-8 com BOM). A validacao e sempre a do
carregar_mapa_empresas (consolidacao.empresas), sem regra duplicada aqui.

Gravacao (controle otimista, uma alteracao por confirmacao):
  pegar a trava -> reler -> se outra pessoa mudou o mesmo COD_SAUDE, barrar
  (ConflitoEdicao) -> aplicar e validar -> historico do dia -> temporario +
  os.replace -> soltar a trava (sempre).
Nao depende de Tk.
"""
import csv
import hashlib
import io
import json
import logging
import os
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Literal

from consolidacao.cabecalhos import normalizar_nome, trimws
from consolidacao.empresas import decodificar_csv, mapa_empresas_de_texto
from consolidacao.erros import ErroValidacao

CABECALHO = ["EMPRESA", "COD_SAUDE", "COD_DENTAL"]
ARQUIVO_CACHE = "empresas_cache.csv"
ARQUIVO_INFO_CACHE = "cache.json"
PASTA_HISTORICO = "historico"
# Trava mais velha que isto foi deixada por um app que fechou no meio.
SEGUNDOS_TRAVA_ABANDONADA = 120

log = logging.getLogger(__name__)


class ConflitoEdicao(ErroValidacao):
    """Outra pessoa alterou o mesmo COD_SAUDE desde a ultima leitura."""


class EmUso(ErroValidacao):
    """Arquivo travado por outra pessoa ou aberto em outro programa."""

    def __init__(self, mensagem: str, usuario: str | None = None, maquina: str | None = None):
        super().__init__(mensagem)
        self.usuario, self.maquina = usuario, maquina


@dataclass(frozen=True)
class Linha:
    empresa: str
    cod_saude: str
    cod_dental: str = ""


@dataclass
class Relacao:
    linhas: list
    hash: str
    origem: Literal["rede", "cache"]
    lida_em: datetime
    data_cache: datetime | None = None


@dataclass(frozen=True)
class Alteracao:
    tipo: Literal["incluir", "editar", "excluir"]
    antigo: Linha | None = None
    novo: Linha | None = None


def incluir(novo: Linha) -> Alteracao:
    return Alteracao("incluir", novo=novo)


def editar(antigo: Linha, novo: Linha) -> Alteracao:
    return Alteracao("editar", antigo=antigo, novo=novo)


def excluir(antigo: Linha) -> Alteracao:
    return Alteracao("excluir", antigo=antigo)


# --- formato -------------------------------------------------------------------

def chave(linha: Linha) -> str:
    return trimws(linha.cod_saude or "")


def validar_linha(linha: Linha) -> Linha:
    """Regra do formulario: tira os espacos das pontas e exige EMPRESA e
    COD_SAUDE (o carregador ignoraria a linha em silencio)."""
    limpa = Linha(trimws(linha.empresa or ""), trimws(linha.cod_saude or ""),
                  trimws(linha.cod_dental or ""))
    if not limpa.empresa or limpa.empresa == "NA":
        raise ErroValidacao("Informe o nome da EMPRESA.")
    if not limpa.cod_saude or limpa.cod_saude == "NA":
        raise ErroValidacao("Informe o COD_SAUDE.")
    return limpa


def _hash(dados: bytes) -> str:
    return hashlib.sha256(dados).hexdigest()


def _linhas_de_texto(texto: str) -> list:
    linhas = [l for l in texto.splitlines() if l.strip() != ""]
    if not linhas:
        raise ErroValidacao("O arquivo de empresas esta vazio.")
    registros = list(csv.reader(io.StringIO("\n".join(linhas)), delimiter=";", quotechar='"'))
    nomes = [normalizar_nome(n) for n in registros[0]]
    if "EMPRESA" not in nomes or "COD_SAUDE" not in nomes:
        raise ErroValidacao("O arquivo de empresas precisa conter as colunas EMPRESA e COD_SAUDE.")
    indices = [nomes.index(n) if n in nomes else None for n in CABECALHO]

    def campo(registro, i):
        return registro[i] if i is not None and i < len(registro) else ""

    return [Linha(*(campo(r, i) for i in indices)) for r in registros[1:]]


def _texto(linhas: list) -> str:
    saida = io.StringIO()
    escritor = csv.writer(saida, delimiter=";", quotechar='"', lineterminator="\r\n")
    escritor.writerow(CABECALHO)
    for l in linhas:
        escritor.writerow([l.empresa, l.cod_saude, l.cod_dental])
    return saida.getvalue()


def _bytes(linhas: list) -> bytes:
    # utf-8-sig: o Excel abre com acentos certos e o R/CLI continuam lendo.
    return _texto(linhas).encode("utf-8-sig")


def mapa(relacao: Relacao) -> dict:
    """codigo -> empresa, pela mesma regra do carregar_mapa_empresas."""
    return mapa_empresas_de_texto(_texto(relacao.linhas))


# --- leitura e cache -------------------------------------------------------------

def _atualizar_cache(pasta_dados: Path, dados: bytes, agora: datetime) -> None:
    try:
        _gravar_atomico(Path(pasta_dados) / ARQUIVO_CACHE, dados)
        info = {"hash": _hash(dados), "lida_em": agora.isoformat()}
        _gravar_atomico(Path(pasta_dados) / ARQUIVO_INFO_CACHE,
                        json.dumps(info).encode("utf-8"))
    except OSError:
        # Sem cache o app ainda funciona com rede; so registra.
        log.exception("Nao foi possivel atualizar a copia local da relacao")


def _ler_cache(pasta_dados: Path, caminho_rede: Path, erro: OSError, agora: datetime) -> Relacao:
    cache = Path(pasta_dados) / ARQUIVO_CACHE
    try:
        dados = cache.read_bytes()
    except OSError:
        raise ErroValidacao(
            f"Relacao de empresas indisponivel: nao foi possivel ler '{caminho_rede}' "
            "e nao ha copia local. Confira a rede ou configure o arquivo."
        ) from erro
    try:
        info = json.loads((Path(pasta_dados) / ARQUIVO_INFO_CACHE).read_text(encoding="utf-8"))
        data_cache = datetime.fromisoformat(info["lida_em"])
    except (OSError, ValueError, KeyError, TypeError):
        data_cache = datetime.fromtimestamp(cache.stat().st_mtime)
    log.warning("Rede indisponivel (%s); usando copia local de %s", erro, data_cache)
    linhas = _linhas_de_texto(decodificar_csv(dados, cache))
    return Relacao(linhas, _hash(dados), "cache", agora, data_cache)


def ler(caminho_rede: Path, pasta_dados: Path, agora: datetime | None = None) -> Relacao:
    """Le a relacao da rede e atualiza o cache; sem rede, usa o cache."""
    agora = agora or datetime.now()
    caminho_rede = Path(caminho_rede)
    try:
        dados = caminho_rede.read_bytes()
    except OSError as e:
        return _ler_cache(pasta_dados, caminho_rede, e, agora)
    linhas = _linhas_de_texto(decodificar_csv(dados, caminho_rede))
    _atualizar_cache(pasta_dados, dados, agora)
    return Relacao(linhas, _hash(dados), "rede", agora)


# --- gravacao --------------------------------------------------------------------

def _gravar_atomico(caminho: Path, conteudo: bytes) -> None:
    """Grava num temporario da mesma pasta e troca: o definitivo nunca fica
    pela metade. Prefixo "~$": a listagem de entradas ja o deixa de fora."""
    caminho = Path(caminho)
    descritor, nome = tempfile.mkstemp(dir=caminho.parent, prefix="~$", suffix=".tmp")
    try:
        with os.fdopen(descritor, "wb") as f:
            f.write(conteudo)
            f.flush()
            os.fsync(f.fileno())
        os.replace(nome, caminho)
    except BaseException:
        try:
            Path(nome).unlink(missing_ok=True)
        except OSError:
            pass  # nao esconde o erro original
        raise


def _caminho_trava(caminho_rede: Path) -> Path:
    return caminho_rede.with_name(caminho_rede.name + ".lock")


def _dono_da_trava(trava: Path) -> tuple:
    """(usuario, maquina, hora) gravados na trava; hora None se ilegivel."""
    try:
        partes = trava.read_text(encoding="utf-8").splitlines()
        return partes[0], partes[1], datetime.fromisoformat(partes[2])
    except (OSError, ValueError, IndexError, UnicodeDecodeError):
        return None, None, None


def _idade_trava(trava: Path, agora: datetime) -> float:
    _, _, hora = _dono_da_trava(trava)
    if hora is None:
        hora = datetime.fromtimestamp(trava.stat().st_mtime)
    return (agora - hora).total_seconds()


def _pegar_trava(caminho_rede: Path, usuario: str, maquina: str, agora: datetime) -> Path:
    trava = _caminho_trava(caminho_rede)
    conteudo = f"{usuario}\n{maquina}\n{agora.isoformat()}\n".encode("utf-8")
    for tentativa in range(2):
        try:
            descritor = os.open(trava, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                abandonada = _idade_trava(trava, agora) > SEGUNDOS_TRAVA_ABANDONADA
            except FileNotFoundError:
                continue  # o colega acabou de soltar
            if abandonada and tentativa == 0:
                log.warning("Removendo trava abandonada %s", trava)
                trava.unlink(missing_ok=True)
                continue
            dono, maquina_dono, _ = _dono_da_trava(trava)
            raise EmUso(
                f"A relacao de empresas esta em uso por {dono or '?'}@{maquina_dono or '?'}; "
                "tente de novo em instantes.",
                dono, maquina_dono,
            )
        with os.fdopen(descritor, "wb") as f:
            f.write(conteudo)
        return trava
    raise EmUso("A relacao de empresas esta em uso; tente de novo em instantes.")


def _soltar_trava(trava: Path) -> None:
    try:
        trava.unlink(missing_ok=True)
    except OSError:
        # Fica para o proximo, que a remove depois de 2 minutos.
        log.exception("Nao foi possivel remover a trava %s", trava)


def _da_chave(linhas: list, k: str) -> list:
    return [l for l in linhas if chave(l) == k]


def _chaves_afetadas(alteracao: Alteracao) -> set:
    return {chave(l) for l in (alteracao.antigo, alteracao.novo) if l is not None}


def _aplicar_alteracao(linhas: list, alteracao: Alteracao) -> list:
    linhas = list(linhas)
    if alteracao.tipo == "excluir":
        k = chave(alteracao.antigo)
        if not _da_chave(linhas, k):
            raise ErroValidacao(f"COD_SAUDE '{k}' nao esta mais na relacao.")
        return [l for l in linhas if chave(l) != k]

    novo = validar_linha(alteracao.novo)
    k = novo.cod_saude
    if alteracao.tipo == "incluir":
        existentes = _da_chave(linhas, k)
        if existentes:
            raise ErroValidacao(
                f"COD_SAUDE '{k}' ja esta cadastrado para '{existentes[0].empresa}'."
            )
        return linhas + [novo]

    k_antigo = chave(alteracao.antigo)
    posicoes = [i for i, l in enumerate(linhas) if chave(l) == k_antigo]
    if not posicoes:
        raise ErroValidacao(f"COD_SAUDE '{k_antigo}' nao esta mais na relacao.")
    if k != k_antigo and _da_chave(linhas, k):
        raise ErroValidacao(
            f"COD_SAUDE '{k}' ja esta cadastrado para '{_da_chave(linhas, k)[0].empresa}'."
        )
    # Trocar o COD_SAUDE vale como excluir a antiga e incluir a nova.
    resultado = [l for i, l in enumerate(linhas) if i not in posicoes[1:]]
    resultado[posicoes[0]] = novo
    return resultado


def _gravar_historico(caminho_rede: Path, anterior: bytes, agora: datetime) -> None:
    pasta = caminho_rede.parent / PASTA_HISTORICO
    pasta.mkdir(exist_ok=True)
    # Um por dia, sobrescrito no mesmo dia; nada e apagado automaticamente.
    _gravar_atomico(pasta / f"{caminho_rede.stem}_{agora:%Y%m%d}.csv", anterior)


def aplicar(
    caminho_rede: Path,
    base: Relacao,
    alteracao: Alteracao,
    usuario: str,
    maquina: str,
    agora: datetime | None = None,
    pasta_dados: Path | None = None,
) -> Relacao:
    """Grava uma alteracao na relacao compartilhada e devolve a relacao nova.

    Levanta ConflitoEdicao, EmUso ou ErroValidacao; o arquivo da rede nunca
    fica pela metade e a trava e sempre solta."""
    agora = agora or datetime.now()
    caminho_rede = Path(caminho_rede)
    trava = None
    try:
        trava = _pegar_trava(caminho_rede, usuario, maquina, agora)
        anterior = caminho_rede.read_bytes()
        atual = _linhas_de_texto(decodificar_csv(anterior, caminho_rede))
        if _hash(anterior) != base.hash:
            for k in _chaves_afetadas(alteracao):
                if _da_chave(atual, k) != _da_chave(base.linhas, k):
                    raise ConflitoEdicao(
                        "Atencao: outra pessoa alterou esta empresa agora; "
                        "confira e confirme de novo."
                    )
        novas = _aplicar_alteracao(atual, alteracao)
        conteudo = _bytes(novas)
        mapa_empresas_de_texto(conteudo.decode("utf-8-sig"))
        _gravar_historico(caminho_rede, anterior, agora)
        _gravar_atomico(caminho_rede, conteudo)
    except PermissionError as e:
        raise EmUso(
            "Nao foi possivel gravar a relacao de empresas: arquivo aberto em outro "
            "programa (Excel?) ou sem permissao. Feche-o e tente de novo."
        ) from e
    except OSError as e:
        raise ErroValidacao(
            f"Nao foi possivel gravar a relacao de empresas em '{caminho_rede}': "
            f"confira a rede e tente de novo. ({e})"
        ) from e
    finally:
        if trava is not None:
            _soltar_trava(trava)
    if pasta_dados is not None:
        _atualizar_cache(pasta_dados, conteudo, agora)
    return Relacao(novas, _hash(conteudo), "rede", agora)


def criar_a_partir_de_csv(origem: Path, caminho_rede: Path) -> Relacao:
    """Primeiro uso: valida um COD_EMPRESAS.csv e cria o arquivo compartilhado."""
    origem, caminho_rede = Path(origem), Path(caminho_rede)
    if caminho_rede.exists():
        raise ErroValidacao(
            f"O arquivo '{caminho_rede}' ja existe; escolha-o como relacao existente."
        )
    from consolidacao.empresas import carregar_mapa_empresas

    carregar_mapa_empresas(origem)
    linhas = _linhas_de_texto(decodificar_csv(origem.read_bytes(), origem))
    conteudo = _bytes(linhas)
    try:
        _gravar_atomico(caminho_rede, conteudo)
    except OSError as e:
        raise ErroValidacao(
            f"Nao foi possivel criar '{caminho_rede}': confira a pasta e a rede. ({e})"
        ) from e
    return Relacao(linhas, _hash(conteudo), "rede", datetime.now())
