"""Achados do debate dos 3 revisores (2026-10-22)."""
from datetime import date

import openpyxl

from consolidacao.leitura import ler_aba
from consolidacao.datas import validar_colunas_data


def _planilha_formatada(caminho, cabecalho, valores, formato):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "A"
    ws.append(cabecalho)
    ws.append(valores)
    for celula in ws[2]:
        celula.number_format = formato
    wb.save(caminho)


def test_a1_numero_grande_em_celula_de_data_nao_some(tmp_path):
    p = tmp_path / "x.xlsx"
    _planilha_formatada(p, ["NUMERO DO CPF", "DATA DE NASCIMENTO"], [12345678901, 19900512], "dd/mm/yyyy")
    df = ler_aba(p, "A")
    assert df["NUMERO DO CPF"].tolist() == ["12345678901"]
    _, obs, _ = validar_colunas_data(df, "C")
    assert len(obs) == 1 and "19900512" in obs[0]


def test_a1_data_normal_continua_data(tmp_path):
    p = tmp_path / "x.xlsx"
    _planilha_formatada(p, ["DATA DE NASCIMENTO"], [45000], "dd/mm/yyyy")
    df, obs, _ = validar_colunas_data(ler_aba(p, "A"), "C")
    assert df["DATA DE NASCIMENTO"].tolist() == [date(2023, 3, 15)] and obs == []


def _consolidado_minimo(tmp_path):
    from consolidacao.processamento import processar_arquivos
    from tests.fabrica import DEPENDENTE_OK, SUPLEMENTAR_OK, TITULAR_OK, criar_arquivo_pc

    a = tmp_path / "PC123 x.xlsx"
    criar_arquivo_pc(a, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    return processar_arquivos([a], {"123": "ACME"})


def test_x1_salvar_cria_pasta_e_nao_deixa_temporario(tmp_path):
    from consolidacao.exportacao import escrever_excel

    destino = tmp_path / "nova" / "saida.xlsx"
    escrever_excel(_consolidado_minimo(tmp_path), destino)
    assert destino.exists()
    assert [p.name for p in destino.parent.iterdir()] == ["saida.xlsx"]


def test_x1_falha_ao_salvar_vira_erro_claro_sem_temporario(tmp_path):
    import pytest

    from consolidacao.erros import ErroValidacao
    from consolidacao.exportacao import escrever_excel

    c = _consolidado_minimo(tmp_path)
    destino = tmp_path / "saida.xlsx"
    destino.mkdir()  # destino ocupado: os.replace falha, como com o arquivo aberto no Excel
    with pytest.raises(ErroValidacao, match="Nao foi possivel salvar o arquivo de saida"):
        escrever_excel(c, destino)
    assert sorted(p.name for p in tmp_path.iterdir()) == ["PC123 x.xlsx", "saida.xlsx"]


def test_x1_cli_retorna_1_quando_nao_salva(tmp_path, capsys):
    from tests.fabrica import DEPENDENTE_OK, SUPLEMENTAR_OK, TITULAR_OK, criar_arquivo_pc, escrever_csv_empresas
    from tratamento import main

    criar_arquivo_pc(tmp_path / "PC123 x.xlsx", TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    escrever_csv_empresas(tmp_path / "e.csv", [("ACME", "123")])
    (tmp_path / "ocupado.xlsx").mkdir()
    codigo = main(["--entrada", str(tmp_path), "--empresas", str(tmp_path / "e.csv"),
                   "--saida", str(tmp_path / "ocupado.xlsx")])
    assert codigo == 1
    assert "Erro: Nao foi possivel salvar o arquivo de saida" in capsys.readouterr().out


def test_a2_duracao_em_coluna_de_texto_vira_numero(tmp_path):
    p = tmp_path / "x.xlsx"
    _planilha_formatada(p, ["HORAS"], [1.5], "[h]:mm:ss")
    assert ler_aba(p, "A")["HORAS"].tolist() == ["1.5"]
