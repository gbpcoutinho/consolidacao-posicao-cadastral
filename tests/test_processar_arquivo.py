from consolidacao.processamento import processar_arquivo
from tests.fabrica import DEPENDENTE_OK, SUPLEMENTAR_OK, TITULAR_OK, criar_arquivo_pc

MAPA = {"123": "ACME"}


def _arquivo(tmp_path, nome="PC123 10.2026.xlsx", **kw):
    p = tmp_path / nome
    extras = {"abas": kw["abas"]} if "abas" in kw else {}
    criar_arquivo_pc(
        p,
        kw.get("tit", TITULAR_OK),
        kw.get("dep", DEPENDENTE_OK),
        kw.get("sup", SUPLEMENTAR_OK),
        **extras,
    )
    return p


def test_processar_ok(tmp_path):
    r = processar_arquivo(_arquivo(tmp_path), MAPA)
    assert (r.status, r.mensagem, r.empresa, r.codigo_empresa) == (
        "OK", "Processado com sucesso.", "ACME", "123"
    )
    assert (r.arquivo, r.referencia) == ("PC123 10.2026.xlsx", "PC123 10.2026")
    assert (r.n_titulares, r.n_dependentes) == (2, 2)
    f = r.final
    assert "SubF" in f.columns and "NUMERO DA SUBFATURA" not in f.columns
    assert ".CHAVE_CERTIFICADO" in f.columns  # fiel ao R: a coluna auxiliar sai na base
    dep = f[f["TITULARIDADE"] == "DEPENDENTE"]
    assert dep.iloc[0]["NOME DO SEGURADO"] == TITULAR_OK[0]["NOME DO SEGURADO"]
    assert dep.iloc[0]["SubF"] == TITULAR_OK[0]["NUMERO DA SUBFATURA"]
    assert dep.iloc[1]["NOME DO SEGURADO"] is None
    assert f["CÓDIGO UNICO"].tolist() == ["00100", "00200", "00101", "99902"]
    tit001 = f[(f["TITULARIDADE"] == "TITULAR") & (f[".CHAVE_CERTIFICADO"] == "001")].iloc[0]
    assert tit001["NUMERO DO CARTAO"] == SUPLEMENTAR_OK[0]["CARTÃO"]
    assert tit001["NOME DA MAE"] == SUPLEMENTAR_OK[0]["NOME DA MÃE"]
    assert dep.iloc[0]["NUMERO DO CARTAO"] == DEPENDENTE_OK[0]["CARTÃO DO DEPENDENTE"]
    assert r.pendencias["MOTIVO"].tolist() == ["CERTIFICADO SEM TITULAR CORRESPONDENTE"]
    assert list(r.pendencias.columns) == [
        "REFERENCIA", "EMPRESA", "NUMERO DA SUBFATURA", "NUMERO DO CERTIFICADO",
        "CODIGO DO DEPENDENTE", "NOME DEPENDENTE", "NOME DO SEGURADO", "PLANO", "MOTIVO",
    ]
    assert "Titulares sem registro na aba suplementar: 1" in r.avisos
    assert r.avisos.startswith("TITULAR ausente: DATA DE REATIVACAO, CARGO DE OCUPACAO, NUMERO DA MATRICULA | ")


def test_aba_faltando(tmp_path):
    r = processar_arquivo(_arquivo(tmp_path, abas=("TITULAR", "DEPENDENTE", "OUTRA")), MAPA)
    assert r.status == "ERRO"
    assert r.mensagem == "Esperada exatamente uma aba POS. CADASTRAL (NOME M E); encontrada(s) 0."
    assert r.final is None and r.n_titulares is None and r.codigo_empresa is None


def test_certificado_duplicado(tmp_path):
    r = processar_arquivo(_arquivo(tmp_path, tit=[TITULAR_OK[0], TITULAR_OK[0]]), MAPA)
    assert r.mensagem == "Certificado(s) duplicado(s) na aba TITULAR: 001"


def test_codigo_de_empresa_inexistente(tmp_path):
    r = processar_arquivo(_arquivo(tmp_path, nome="PC555.xlsx"), MAPA)
    assert r.status == "ERRO"
    assert r.mensagem.startswith("Codigo '555' extraido de 'PC555'")


def test_observacoes_parciais_sobrevivem_ao_erro(tmp_path):
    tit = [dict(TITULAR_OK[0], **{"DATA DE NASCIMENTO": "xx"}), TITULAR_OK[1]]
    r = processar_arquivo(_arquivo(tmp_path, tit=tit, dep=[]), MAPA)
    assert r.status == "ERRO" and len(r.observacoes_datas) == 1
