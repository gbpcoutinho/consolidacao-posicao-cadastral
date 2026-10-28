"""Arquivos do empacotamento (Tarefa 9). O build em si e verificado por
scripts/gerar_exe.ps1, que roda o fumo do .exe nos exemplos."""
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]


def test_spec_do_pyinstaller():
    spec = (RAIZ / "ConsolidadorPC.spec").read_text(encoding="utf-8")
    assert "consolidador_app.py" in spec
    assert "name='ConsolidadorPC'" in spec
    assert "console=False" in spec
    # Modo pasta (--onedir): ha COLLECT e o EXE nao embute os binarios.
    assert "COLLECT(" in spec and "exclude_binaries=True" in spec


def test_script_de_build_roda_o_fumo_e_gera_o_zip():
    ps1 = (RAIZ / "scripts" / "gerar_exe.ps1").read_text(encoding="utf-8")
    assert "--fumo" in ps1
    assert "ConsolidadorPC-$versao.zip" in ps1
    assert "ExitCode" in ps1


def test_build_e_dist_fora_do_git():
    linhas = (RAIZ / ".gitignore").read_text(encoding="utf-8").splitlines()
    assert "build/" in linhas and "dist/" in linhas


def test_fumo_descartavel_removido():
    assert not (RAIZ / "scripts" / "fumo_build.py").exists()


def test_documentacao():
    for nome in ("manual-usuario.md", "roteiro-teste-manual.md"):
        assert (RAIZ / "docs" / nome).exists()
    roteiro = (RAIZ / "docs" / "roteiro-teste-manual.md").read_text(encoding="utf-8")
    assert all(f"\n{i}. " in roteiro for i in range(1, 11))
    assert "2026-10-26-fase2-app-design.md" in (RAIZ / "docs" / "fase2-app.md").read_text(
        encoding="utf-8")


def test_recursos_so_no_executavel_e_fora_do_git():
    spec = (RAIZ / "ConsolidadorPC.spec").read_text(encoding="utf-8")
    assert "recursos" in spec and "icon=" in spec
    linhas = (RAIZ / ".gitignore").read_text(encoding="utf-8").splitlines()
    assert "app/recursos/" in linhas and "logo_branco_GVS.png" in linhas
