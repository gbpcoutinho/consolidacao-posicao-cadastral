import logging
import threading
import time

import pytest

from app.tarefas import Evento, Trabalho
from consolidacao.erros import ErroValidacao, ProcessamentoCancelado


def _esperar(trabalho, limite=5.0):
    """Consome os eventos ate o terminal, como a janela faz com after()."""
    vistos = []
    fim = time.monotonic() + limite
    while time.monotonic() < fim:
        vistos += trabalho.eventos()
        if vistos and vistos[-1].tipo in ("concluido", "erro", "cancelado"):
            return vistos
        time.sleep(0.01)
    raise AssertionError(f"sem evento terminal: {vistos}")


def test_eventos_em_ordem_progresso_concluido():
    def funcao(progresso, cancelar):
        for i in range(1, 4):
            progresso(i, 3, f"arq{i}")
        return "resultado"

    t = Trabalho(funcao)
    t.iniciar()
    vistos = _esperar(t)
    assert vistos == [
        Evento("progresso", (1, 3, "arq1")),
        Evento("progresso", (2, 3, "arq2")),
        Evento("progresso", (3, 3, "arq3")),
        Evento("concluido", "resultado"),
    ]
    t.esperar(1)
    assert not t.ativo


def test_eventos_nao_bloqueia():
    liberar = threading.Event()
    t = Trabalho(lambda p, c: liberar.wait(5))
    t.iniciar()
    inicio = time.monotonic()
    assert t.eventos() == []
    assert time.monotonic() - inicio < 0.5
    assert t.ativo
    liberar.set()
    _esperar(t)


def test_cancelar_gera_cancelado():
    comecou = threading.Event()

    def funcao(progresso, cancelar):
        comecou.set()
        while True:
            if cancelar():
                raise ProcessamentoCancelado("Processamento cancelado pelo usuario.")
            time.sleep(0.01)

    t = Trabalho(funcao)
    t.iniciar()
    assert comecou.wait(5)
    t.cancelar()
    vistos = _esperar(t)
    assert vistos[-1].tipo == "cancelado"


def test_erro_validacao_vira_erro_com_mensagem():
    def funcao(progresso, cancelar):
        raise ErroValidacao("Arquivo de empresas nao encontrado: x.csv")

    t = Trabalho(funcao)
    t.iniciar()
    assert _esperar(t)[-1] == Evento("erro", "Arquivo de empresas nao encontrado: x.csv")


def test_memoria_vira_erro_em_portugues():
    def funcao(progresso, cancelar):
        raise MemoryError()

    t = Trabalho(funcao)
    t.iniciar()
    ultimo = _esperar(t)[-1]
    assert ultimo.tipo == "erro" and "memoria" in ultimo.dados.lower()


def test_excecao_inesperada_vai_para_log(caplog):
    def funcao(progresso, cancelar):
        raise KeyError("detalhe tecnico")

    t = Trabalho(funcao)
    with caplog.at_level(logging.ERROR):
        t.iniciar()
        ultimo = _esperar(t)[-1]
    assert ultimo == Evento("erro", "Erro inesperado; detalhes em logs/app.log")
    assert "detalhe tecnico" in caplog.text and "Traceback" in caplog.text


def test_callbacks_so_na_thread_consumidora():
    principal = threading.get_ident()
    threads = {}

    def funcao(progresso, cancelar):
        threads["funcao"] = threading.get_ident()
        progresso(1, 1, "a")
        return 1

    tratados = []
    t = Trabalho(funcao)
    t.iniciar()
    fim = time.monotonic() + 5
    while time.monotonic() < fim and not (tratados and tratados[-1][0].tipo == "concluido"):
        t.despachar(lambda e: tratados.append((e, threading.get_ident())))
        time.sleep(0.01)
    assert threads["funcao"] != principal
    assert [e.tipo for e, _ in tratados] == ["progresso", "concluido"]
    assert {ident for _, ident in tratados} == {principal}


def test_nao_inicia_duas_vezes():
    liberar = threading.Event()
    t = Trabalho(lambda p, c: liberar.wait(5))
    t.iniciar()
    with pytest.raises(RuntimeError):
        t.iniciar()
    liberar.set()
    _esperar(t)


def test_erro_inesperado_do_motor_vai_para_log_e_tela_generica(tmp_path, monkeypatch, caplog):
    from consolidacao import pipeline

    def explode(*a, **k):
        raise KeyError("NUMERO DO CERTIFICADO")

    monkeypatch.setattr(pipeline, "processar_arquivos", explode)
    t = Trabalho(lambda p, c: pipeline.processar([tmp_path / "x.xlsx"], {}, p, c))
    with caplog.at_level(logging.ERROR):
        t.iniciar()
        ultimo = _esperar(t)[-1]
    assert ultimo == Evento("erro", "Erro inesperado; detalhes em logs/app.log")
    assert "KeyError" in caplog.text and "NUMERO DO CERTIFICADO" in caplog.text
    assert "Traceback" in caplog.text and "explode" in caplog.text
