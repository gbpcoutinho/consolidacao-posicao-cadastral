import logging
import logging.handlers
from pathlib import Path

import pytest

import consolidador_app

EXEMPLOS = Path(__file__).resolve().parents[2] / "exemplos"


@pytest.fixture
def raiz_log():
    raiz = logging.getLogger()
    antes = list(raiz.handlers)
    yield raiz
    for h in raiz.handlers:
        if h not in antes:
            h.close()
            raiz.removeHandler(h)


def test_configurar_log_rotativo(tmp_path, raiz_log):
    consolidador_app.configurar_log(tmp_path)
    handlers = [h for h in raiz_log.handlers if isinstance(h, logging.handlers.RotatingFileHandler)]
    assert len(handlers) == 1
    h = handlers[0]
    assert Path(h.baseFilename) == tmp_path / "logs" / "app.log"
    assert (h.maxBytes, h.backupCount) == (1_000_000, 3)
    logging.getLogger("app.teste").error("falha ficticia")
    h.flush()
    assert "falha ficticia" in (tmp_path / "logs" / "app.log").read_text(encoding="utf-8")


def test_excecao_nao_tratada_vai_para_log_e_mostra_mensagem_generica(tmp_path, raiz_log, caplog):
    vistos = []
    try:
        raise KeyError("detalhe tecnico")
    except KeyError as e:
        with caplog.at_level(logging.ERROR):
            consolidador_app.tratar_excecao(type(e), e, e.__traceback__, mostrar=vistos.append)
    assert "detalhe tecnico" in caplog.text and "Traceback" in caplog.text
    assert vistos == ["Erro inesperado; detalhes em logs/app.log"]


def test_fumo_processa_os_exemplos(tmp_path, raiz_log):
    saida = tmp_path / "fumo.xlsx"
    codigo = consolidador_app.main(
        ["--fumo", str(EXEMPLOS / "entrada"), str(EXEMPLOS / "COD_EMPRESAS_EXEMPLO.csv"), str(saida)],
        pasta=tmp_path / "dados",
    )
    assert codigo == 0 and saida.exists()


def test_fumo_com_erro_devolve_1(tmp_path, raiz_log):
    codigo = consolidador_app.main(
        ["--fumo", str(tmp_path / "nao_existe"), str(EXEMPLOS / "COD_EMPRESAS_EXEMPLO.csv"),
         str(tmp_path / "x.xlsx")],
        pasta=tmp_path / "dados",
    )
    assert codigo == 1
