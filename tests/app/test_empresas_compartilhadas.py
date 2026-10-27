"""Relacao de empresas compartilhada; tmp_path simula a pasta da rede."""
import os
from datetime import datetime, timedelta

import pytest

from app import empresas_compartilhadas as ec
from app.empresas_compartilhadas import (
    ConflitoEdicao,
    EmUso,
    Linha,
    aplicar,
    criar_a_partir_de_csv,
    editar,
    excluir,
    incluir,
    ler,
    mapa,
)
from consolidacao.empresas import carregar_mapa_empresas
from consolidacao.erros import ErroValidacao

AGORA = datetime(2026, 10, 26, 14, 30, 0)
CONTEUDO = (
    "EMPRESA;COD_SAUDE;COD_DENTAL\n"
    "ALFA FICTICIA;90001;D1\n"
    "BETA FICTICIA;90002;\n"
)


@pytest.fixture
def rede(tmp_path):
    pasta = tmp_path / "rede"
    pasta.mkdir()
    caminho = pasta / "empresas.csv"
    caminho.write_bytes(CONTEUDO.encode("utf-8-sig"))
    return caminho


@pytest.fixture
def dados(tmp_path):
    pasta = tmp_path / "appdata"
    pasta.mkdir()
    return pasta


def _aplicar(rede, base, alteracao, agora=AGORA, **kw):
    return aplicar(rede, base, alteracao, usuario="ana", maquina="PC01", agora=agora, **kw)


# --- leitura e cache -------------------------------------------------------

def test_ler_rede_atualiza_cache(rede, dados):
    r = ler(rede, dados)
    assert r.origem == "rede" and r.data_cache is None
    assert r.linhas == [
        Linha("ALFA FICTICIA", "90001", "D1"),
        Linha("BETA FICTICIA", "90002", ""),
    ]
    assert (dados / "empresas_cache.csv").read_bytes() == rede.read_bytes()
    assert (dados / "cache.json").exists()
    assert mapa(r) == {"90001": "ALFA FICTICIA", "90002": "BETA FICTICIA"}


def test_ler_sem_rede_usa_cache_origem_cache(rede, dados):
    lida = ler(rede, dados)
    rede.rename(rede.with_name("fora.csv"))  # "rede caiu"
    r = ler(rede, dados)
    assert r.origem == "cache"
    assert r.linhas == lida.linhas and r.hash == lida.hash
    assert isinstance(r.data_cache, datetime)


def test_ler_sem_rede_sem_cache_erro(rede, dados):
    with pytest.raises(ErroValidacao, match="Relacao de empresas indisponivel"):
        ler(rede.parent / "nao_existe" / "empresas.csv", dados)


def test_ler_csv_ansi_do_excel(rede, dados):
    rede.write_bytes("EMPRESA;COD_SAUDE;COD_DENTAL\nAÇÃO FICTICIA;1;\n".encode("cp1252"))
    assert ler(rede, dados).linhas == [Linha("AÇÃO FICTICIA", "1", "")]


def test_mapa_usa_a_regra_do_carregador(rede, dados):
    rede.write_bytes("EMPRESA;COD_SAUDE;COD_DENTAL\nA;1;\nB;1;\n".encode("utf-8-sig"))
    with pytest.raises(ErroValidacao, match="mais de uma EMPRESA"):
        mapa(ler(rede, dados))


# --- gravacao ----------------------------------------------------------------

def test_incluir_grava_bom_utf8_e_valida(rede, dados):
    base = ler(rede, dados)
    nova = _aplicar(rede, base, incluir(Linha("  GAMA ÇÃO  ", " 90003 ", " ")))
    dados_brutos = rede.read_bytes()
    assert dados_brutos.startswith(b"\xef\xbb\xbf")
    assert dados_brutos.decode("utf-8-sig").splitlines()[0] == "EMPRESA;COD_SAUDE;COD_DENTAL"
    assert carregar_mapa_empresas(rede)["90003"] == "GAMA ÇÃO"
    assert nova.linhas[-1] == Linha("GAMA ÇÃO", "90003", "")
    assert nova.origem == "rede" and nova.hash == ler(rede, dados).hash
    # A versao anterior foi para o historico e a trava foi solta.
    assert not rede.with_name("empresas.csv.lock").exists()


def test_incluir_atualiza_o_cache(rede, dados):
    base = ler(rede, dados)
    _aplicar(rede, base, incluir(Linha("GAMA", "90003", "")), pasta_dados=dados)
    assert (dados / "empresas_cache.csv").read_bytes() == rede.read_bytes()


def test_cod_saude_duplicado_recusado(rede, dados):
    base = ler(rede, dados)
    antes = rede.read_bytes()
    with pytest.raises(ErroValidacao, match="90001"):
        _aplicar(rede, base, incluir(Linha("OUTRA", "90001", "")))
    with pytest.raises(ErroValidacao, match="90002"):
        _aplicar(rede, base, editar(base.linhas[0], Linha("ALFA FICTICIA", "90002", "")))
    assert rede.read_bytes() == antes
    assert not rede.with_name("empresas.csv.lock").exists()


@pytest.mark.parametrize("linha", [Linha("", "9", ""), Linha("X", "  ", ""), Linha("X", "NA", "")])
def test_obrigatorios_recusados(rede, dados, linha):
    with pytest.raises(ErroValidacao):
        _aplicar(rede, ler(rede, dados), incluir(linha))


def test_editar_e_excluir(rede, dados):
    base = ler(rede, dados)
    base = _aplicar(rede, base, editar(base.linhas[1], Linha("BETA NOVA", "90002", "D2")))
    assert carregar_mapa_empresas(rede)["90002"] == "BETA NOVA"
    base = _aplicar(rede, base, excluir(base.linhas[0]))
    assert carregar_mapa_empresas(rede) == {"90002": "BETA NOVA"}
    assert base.linhas == [Linha("BETA NOVA", "90002", "D2")]


def test_editar_cod_saude_vale_como_excluir_e_incluir(rede, dados):
    base = ler(rede, dados)
    _aplicar(rede, base, editar(base.linhas[0], Linha("ALFA FICTICIA", "90009", "D1")))
    assert carregar_mapa_empresas(rede) == {"90009": "ALFA FICTICIA", "90002": "BETA FICTICIA"}


def test_linha_sem_cod_saude_ja_existente_e_preservada(rede, dados):
    rede.write_bytes((CONTEUDO + "SEM CODIGO;;X\n").encode("utf-8-sig"))
    base = ler(rede, dados)
    _aplicar(rede, base, incluir(Linha("GAMA", "90003", "")))
    assert "SEM CODIGO;;X" in rede.read_bytes().decode("utf-8-sig")


def test_campo_com_ponto_e_virgula_volta_igual(rede, dados):
    base = ler(rede, dados)
    _aplicar(rede, base, incluir(Linha('GAMA; "FILIAL"', "90003", "")))
    assert ler(rede, dados).linhas[-1] == Linha('GAMA; "FILIAL"', "90003", "")
    assert carregar_mapa_empresas(rede)["90003"] == 'GAMA; "FILIAL"'


# --- concorrencia ------------------------------------------------------------

def test_edicao_concorrente_em_outra_linha_e_aplicada(rede, dados):
    base_ana = ler(rede, dados)
    base_bia = ler(rede, dados)
    _aplicar(rede, base_ana, editar(base_ana.linhas[0], Linha("ALFA DA ANA", "90001", "D1")))
    _aplicar(rede, base_bia, editar(base_bia.linhas[1], Linha("BETA DA BIA", "90002", "")))
    assert carregar_mapa_empresas(rede) == {"90001": "ALFA DA ANA", "90002": "BETA DA BIA"}


def test_inclusoes_concorrentes_de_chaves_diferentes(rede, dados):
    base_ana = ler(rede, dados)
    base_bia = ler(rede, dados)
    _aplicar(rede, base_ana, incluir(Linha("GAMA", "90003", "")))
    _aplicar(rede, base_bia, incluir(Linha("DELTA", "90004", "")))
    assert set(carregar_mapa_empresas(rede)) == {"90001", "90002", "90003", "90004"}


def test_edicao_concorrente_mesma_linha_conflito(rede, dados):
    base_ana = ler(rede, dados)
    base_bia = ler(rede, dados)
    _aplicar(rede, base_ana, editar(base_ana.linhas[0], Linha("ALFA DA ANA", "90001", "D1")))
    depois_ana = rede.read_bytes()
    with pytest.raises(ConflitoEdicao, match="outra pessoa alterou esta empresa agora"):
        _aplicar(rede, base_bia, excluir(base_bia.linhas[0]))
    assert rede.read_bytes() == depois_ana
    assert not rede.with_name("empresas.csv.lock").exists()


def test_inclusao_concorrente_mesma_chave_conflito(rede, dados):
    base_ana = ler(rede, dados)
    base_bia = ler(rede, dados)
    _aplicar(rede, base_ana, incluir(Linha("GAMA", "90003", "")))
    with pytest.raises(ConflitoEdicao):
        _aplicar(rede, base_bia, incluir(Linha("GAMA OUTRA", "90003", "")))


def test_conflito_e_erro_de_validacao():
    assert issubclass(ConflitoEdicao, ErroValidacao) and issubclass(EmUso, ErroValidacao)


# --- trava -------------------------------------------------------------------

def _trava(rede, usuario, maquina, hora):
    rede.with_name("empresas.csv.lock").write_text(
        f"{usuario}\n{maquina}\n{hora.isoformat()}\n", encoding="utf-8"
    )


def test_trava_viva_em_uso(rede, dados):
    base = ler(rede, dados)
    _trava(rede, "bia", "PC02", AGORA - timedelta(seconds=30))
    antes = rede.read_bytes()
    with pytest.raises(EmUso, match="bia@PC02") as info:
        _aplicar(rede, base, incluir(Linha("GAMA", "90003", "")))
    assert (info.value.usuario, info.value.maquina) == ("bia", "PC02")
    assert rede.read_bytes() == antes
    # A trava do colega continua la.
    assert rede.with_name("empresas.csv.lock").exists()


def test_trava_abandonada_removida(rede, dados):
    base = ler(rede, dados)
    _trava(rede, "bia", "PC02", AGORA - timedelta(seconds=121))
    _aplicar(rede, base, incluir(Linha("GAMA", "90003", "")))
    assert "90003" in carregar_mapa_empresas(rede)
    assert not rede.with_name("empresas.csv.lock").exists()


def test_trava_ilegivel_usa_data_do_arquivo(rede, dados):
    base = ler(rede, dados)
    trava = rede.with_name("empresas.csv.lock")
    trava.write_bytes(b"\xff\xfe lixo")
    velho = (AGORA - timedelta(minutes=5)).timestamp()
    os.utime(trava, (velho, velho))
    _aplicar(rede, base, incluir(Linha("GAMA", "90003", "")))
    assert not trava.exists()


def test_trava_guarda_usuario_maquina_e_hora(rede, dados, monkeypatch):
    base = ler(rede, dados)
    vistos = []
    original = ec._gravar_atomico

    def espiar(caminho, conteudo):
        vistos.append(rede.with_name("empresas.csv.lock").read_text(encoding="utf-8"))
        original(caminho, conteudo)

    monkeypatch.setattr(ec, "_gravar_atomico", espiar)
    _aplicar(rede, base, incluir(Linha("GAMA", "90003", "")))
    # Historico e arquivo definitivo sao gravados com a trava pega.
    assert vistos == [f"ana\nPC01\n{AGORA.isoformat()}\n"] * 2


# --- falhas na gravacao --------------------------------------------------------

def test_falha_no_replace_mantem_arquivo_e_solta_trava(rede, dados, monkeypatch):
    base = ler(rede, dados)
    antes = rede.read_bytes()

    original = os.replace

    def falha(origem, destino):
        # O historico grava; a troca do arquivo definitivo falha.
        if os.path.normcase(destino) == os.path.normcase(rede):
            raise OSError("rede caiu")
        original(origem, destino)

    monkeypatch.setattr(ec.os, "replace", falha)
    with pytest.raises(ErroValidacao, match="Nao foi possivel gravar"):
        _aplicar(rede, base, incluir(Linha("GAMA", "90003", "")))
    monkeypatch.undo()
    assert rede.read_bytes() == antes
    assert not rede.with_name("empresas.csv.lock").exists()
    # Nenhum temporario sobrou na pasta da rede.
    assert sorted(p.name for p in rede.parent.iterdir() if p.is_file()) == ["empresas.csv"]


def test_arquivo_aberto_no_excel_em_uso(rede, dados, monkeypatch):
    base = ler(rede, dados)

    def negado(*a, **k):
        raise PermissionError(13, "Permission denied")

    monkeypatch.setattr(ec.os, "replace", negado)
    with pytest.raises(EmUso, match="arquivo aberto em outro programa"):
        _aplicar(rede, base, incluir(Linha("GAMA", "90003", "")))
    monkeypatch.undo()
    assert not rede.with_name("empresas.csv.lock").exists()


def test_rede_inacessivel_na_gravacao(rede, dados):
    base = ler(rede, dados)
    longe = rede.parent / "sumiu" / "empresas.csv"
    with pytest.raises(ErroValidacao, match="Nao foi possivel"):
        _aplicar(longe, base, incluir(Linha("GAMA", "90003", "")))


# --- historico ---------------------------------------------------------------

def test_historico_um_por_dia(rede, dados):
    base = ler(rede, dados)
    original = rede.read_bytes()
    base = _aplicar(rede, base, incluir(Linha("GAMA", "90003", "")))
    segunda = rede.read_bytes()
    base = _aplicar(rede, base, incluir(Linha("DELTA", "90004", "")), agora=AGORA + timedelta(hours=1))
    historico = rede.parent / "historico"
    assert [p.name for p in historico.iterdir()] == ["empresas_20261026.csv"]
    # Sobrescrito no mesmo dia: guarda a versao anterior a ultima gravacao.
    assert (historico / "empresas_20261026.csv").read_bytes() == segunda
    _aplicar(rede, base, incluir(Linha("EPSILON", "90005", "")), agora=AGORA + timedelta(days=1))
    assert sorted(p.name for p in historico.iterdir()) == [
        "empresas_20261026.csv", "empresas_20261027.csv",
    ]
    assert original != segunda


# --- primeiro uso --------------------------------------------------------------

def test_criar_a_partir_de_csv(tmp_path, dados):
    origem = tmp_path / "COD_EMPRESAS_EXEMPLO.csv"
    origem.write_text("EMPRESA;COD_SAUDE;COD_DENTAL\nALFA;1;\n", encoding="utf-8")
    destino = tmp_path / "rede" / "empresas.csv"
    destino.parent.mkdir()
    r = criar_a_partir_de_csv(origem, destino)
    assert destino.read_bytes().startswith(b"\xef\xbb\xbf")
    assert carregar_mapa_empresas(destino) == {"1": "ALFA"}
    assert r.linhas == [Linha("ALFA", "1", "")] and r.origem == "rede"


def test_criar_a_partir_de_csv_invalido_nao_grava(tmp_path):
    origem = tmp_path / "ruim.csv"
    origem.write_text("EMPRESA;COD_SAUDE\nA;1\nB;1\n", encoding="utf-8")
    destino = tmp_path / "empresas.csv"
    with pytest.raises(ErroValidacao):
        criar_a_partir_de_csv(origem, destino)
    assert not destino.exists()


def test_criar_a_partir_de_csv_nao_sobrescreve(tmp_path, rede):
    origem = tmp_path / "outro.csv"
    origem.write_text("EMPRESA;COD_SAUDE\nA;1\n", encoding="utf-8")
    antes = rede.read_bytes()
    with pytest.raises(ErroValidacao, match="ja existe"):
        criar_a_partir_de_csv(origem, rede)
    assert rede.read_bytes() == antes


DUPLICADO = CONTEUDO + "OUTRA EMPRESA;90002;\n"


def test_excluir_linha_duplicada_nao_apaga_a_legitima(rede, dados):
    rede.write_bytes(DUPLICADO.encode("utf-8-sig"))  # colega duplicou o 90002 no Excel
    base = ler(rede, dados)
    nova = _aplicar(rede, base, excluir(Linha("OUTRA EMPRESA", "90002", "")))
    assert carregar_mapa_empresas(rede) == {"90001": "ALFA FICTICIA", "90002": "BETA FICTICIA"}
    assert Linha("BETA FICTICIA", "90002", "") in nova.linhas


def test_editar_linha_duplicada_so_muda_a_escolhida(rede, dados):
    rede.write_bytes(DUPLICADO.encode("utf-8-sig"))
    base = ler(rede, dados)
    _aplicar(rede, base, editar(Linha("OUTRA EMPRESA", "90002", ""), Linha("OUTRA EMPRESA", "90005", "")))
    assert carregar_mapa_empresas(rede) == {
        "90001": "ALFA FICTICIA", "90002": "BETA FICTICIA", "90005": "OUTRA EMPRESA",
    }
