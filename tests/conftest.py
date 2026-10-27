import copy
from pathlib import Path

import pytest

from consolidacao import leitura
from consolidacao.erros import ProcessamentoCancelado
from tests import leitura_antiga


@pytest.fixture(autouse=True)
def _oraculo_da_leitura(monkeypatch):
    """Em TODA a suite (inclusive a paridade com o R), cada aba lida pela
    leitura em uma passada e conferida contra a leitura antiga (openpyxl +
    varredura): mesmas linhas brutas ou o mesmo erro."""
    original = leitura.ler_linhas_brutas

    def conferido(wb, aba, cancelar=None):
        # Planilha decifrada em memoria: o arquivo e o proprio BytesIO.
        caminho = wb._archive.filename and Path(wb._archive.filename) or wb._archive.fp
        nome = getattr(caminho, "name", "<memoria>")
        try:
            novo = original(wb, aba, cancelar)
            erro_novo = None
        except ProcessamentoCancelado:
            raise
        except Exception as e:
            novo, erro_novo = None, e
        try:
            antigo = leitura_antiga.normalizar(leitura_antiga.linhas_brutas(caminho, aba))
            erro_antigo = None
        except Exception as e:
            antigo, erro_antigo = None, e
        if erro_novo is not None or erro_antigo is not None:
            assert (type(erro_novo), str(erro_novo)) == (type(erro_antigo), str(erro_antigo)), (
                f"{nome} [{aba}]"
            )
            raise erro_novo
        linhas, varridas = novo
        aplicadas = copy.copy([list(l) for l in linhas])
        leitura._aplicar_varredura(aplicadas, varridas)
        assert leitura_antiga.normalizar(aplicadas) == antigo, f"{nome} [{aba}]"
        return novo

    monkeypatch.setattr(leitura, "ler_linhas_brutas", conferido)
