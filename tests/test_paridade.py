import os
from pathlib import Path

import pytest

from tests.paridade.gerar_sinteticos import gerar
from tests.paridade.rodar_paridade import RSCRIPT_PADRAO, rodar

# PARIDADE_RSCRIPT troca o caminho do R; com PARIDADE_OBRIGATORIA=1 a falta do R
# reprova o teste em vez de pula-lo em silencio.
RSCRIPT = Path(os.environ.get("PARIDADE_RSCRIPT", RSCRIPT_PADRAO))


def test_paridade_sinteticos(tmp_path):
    if not RSCRIPT.exists():
        mensagem = f"R nao encontrado em {RSCRIPT} (defina PARIDADE_RSCRIPT)"
        if os.environ.get("PARIDADE_OBRIGATORIA"):
            pytest.fail(mensagem)
        pytest.skip(mensagem)
    pasta = gerar(tmp_path / "entrada")
    # Pasta propria: a suite nao apaga execucoes manuais em _trabalho/.
    assert rodar(pasta, pasta / "COD_EMPRESAS.csv", RSCRIPT, tmp_path / "trabalho") == []
