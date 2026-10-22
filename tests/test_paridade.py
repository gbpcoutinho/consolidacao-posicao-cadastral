from pathlib import Path

import pytest

from tests.paridade.gerar_sinteticos import gerar
from tests.paridade.rodar_paridade import rodar

RSCRIPT = Path(r"C:\Program Files\R\R-4.6.1\bin\Rscript.exe")


@pytest.mark.skipif(not RSCRIPT.exists(), reason="R nao instalado")
def test_paridade_sinteticos(tmp_path):
    pasta = gerar(tmp_path / "entrada")
    # Pasta propria: a suite nao apaga execucoes manuais em _trabalho/.
    assert rodar(pasta, pasta / "COD_EMPRESAS.csv", RSCRIPT, tmp_path / "trabalho") == []
