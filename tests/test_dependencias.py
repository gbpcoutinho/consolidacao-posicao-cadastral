"""A leitura e a gravacao usam partes privadas do openpyxl 3.1.5: a versao e
fixada e o build roda a suite antes do PyInstaller."""
from pathlib import Path

import lxml.etree
import openpyxl

RAIZ = Path(__file__).resolve().parents[1]
OPENPYXL = "3.1.5"


def _requisitos():
    return (RAIZ / "requirements.txt").read_text(encoding="utf-8").splitlines()


def test_openpyxl_instalado_e_o_validado():
    assert openpyxl.__version__ == OPENPYXL


def test_requirements_fixa_openpyxl_e_lxml():
    linhas = _requisitos()
    assert f"openpyxl=={OPENPYXL}" in linhas
    versao_lxml = ".".join(str(n) for n in lxml.etree.LXML_VERSION[:3])
    assert f"lxml=={versao_lxml}" in linhas


def test_build_roda_a_suite_antes_do_pyinstaller():
    ps1 = (RAIZ / "scripts" / "gerar_exe.ps1").read_text(encoding="utf-8")
    assert "-m pytest" in ps1
    assert ps1.index("-m pytest") < ps1.index("-m PyInstaller")


def test_script_de_build_so_tem_ascii_imprimivel():
    ps1 = (RAIZ / "scripts" / "gerar_exe.ps1").read_text(encoding="utf-8")
    assert all(c in "\r\n" or 32 <= ord(c) < 127 for c in ps1)
