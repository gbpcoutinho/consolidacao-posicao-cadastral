"""Testes da janela nao dependem dos recursos locais (logo, icone e padrao de
senha ficam em app/recursos, fora do git): por padrao a pasta aparece vazia."""
import pytest


@pytest.fixture(autouse=True)
def _sem_recursos_locais(tmp_path_factory, monkeypatch):
    from app import janela

    vazia = tmp_path_factory.mktemp("sem_recursos")
    monkeypatch.setattr(janela, "_pasta_recursos", lambda: vazia)
