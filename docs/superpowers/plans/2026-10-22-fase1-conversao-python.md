# Fase 1 — Conversão `TRATAMENTO.R` → Python: Plano de Implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reproduzir em Python, regra por regra, o `TRATAMENTO.R`, com testes automáticos e um teste de paridade que compara as saídas R e Python célula a célula.

**Architecture:** Pacote `consolidacao/` com um módulo por responsabilidade (cabeçalhos, datas, empresas, leitura, processamento, exportação). `tratamento.py` é só a linha de comando. A fase 2 importará `consolidacao` sem alterá-lo. Os dados ficam em `pandas.DataFrame` com dtype `object`: texto é `str`, data é `datetime.date` e ausente é `None` (nunca `NaN`).

**Tech Stack:** Python 3.14 (instalado; piso 3.11), pandas ≥ 2.2, openpyxl ≥ 3.1, pytest ≥ 8. R 4.6.1 em `C:\Program Files\R\R-4.6.1\bin\Rscript.exe` (só para paridade).

**Spec:** `docs/fase1-conversao.md`. A referência de comportamento é `TRATAMENTO.R` (raiz do projeto). Em qualquer dúvida, o R vence. `exportar.R` é legado e está fora do escopo.

## Global Constraints

- **PROIBIDO** ler, listar, pesquisar ou escrever em `ignore/`. Nenhum glob recursivo pode partir da raiz sem excluir `ignore/`.
- Nenhum arquivo temporário fora da pasta do projeto: o pytest usa `--basetemp=.pytest_tmp`, e a paridade usa `tests/paridade/_trabalho/`.
- Mensagens de erro, nomes de abas, nomes de colunas e textos do log são copiados **verbatim** do R (sem acentos onde o R não tem; `"C\u00d3DIGO UNICO"` = `CÓDIGO UNICO`).
- Testes usam apenas dados sintéticos gerados por `tests/fabrica.py`. O `COD_EMPRESAS.csv` real não é usado em testes.
- Comandos rodam no PowerShell, a partir da raiz, com `.venv\Scripts\python`.
- Ausente = `None`. Uma função que devolve DataFrame nunca deixa `NaN`/`NaT`.

## Review Focus

1. **Faixa de leitura do readxl:** linhas vazias no início e no fim são ignoradas, e células formatadas porém vazias não contam. Se o openpyxl inflar `max_row`, a remoção das "2 últimas linhas de resumo" do DEPENDENTE corta dados reais. Teste: `test_dependente_ignora_linhas_vazias_formatadas_no_fim` (Task 4).
2. **Identificador gravado como número** (`12345678901`, `123.0`) precisa virar `"12345678901"` / `"123"`, nunca `"123.0"`. Teste: `test_identificador_numerico_vira_texto_sem_ponto_zero` (Task 4).
3. **Cabeçalhos sujos** (minúsculas, acento, `º`, espaços duplos ou nas pontas) precisam normalizar igual ao R. Teste: `test_normalizar_nome_casos_sujos` (Task 1).
4. **Pasta com `~$` temporário, `.XLSX` maiúsculo e consolidado anterior** dentro dela: só os de entrada entram. Teste: `test_listar_arquivos_entrada_filtra` (Task 6).
5. **`COD_EMPRESAS.csv` salvo pelo Excel em ANSI (cp1252)** em vez de UTF-8. Escolha do Claude: tentar `utf-8-sig` e cair para `cp1252`, para não quebrar e não gerar nome corrompido. Teste: `test_mapa_empresas_cp1252` (Task 3).

---

## Estrutura de arquivos

```
consolidacao/__init__.py
consolidacao/erros.py          ErroValidacao
consolidacao/config.py         constantes do CFG do R
consolidacao/cabecalhos.py     normalização, renomeação, validações de coluna, unificar_colunas
consolidacao/datas.py          conversão de datas, Y2K, resumo de observações
consolidacao/empresas.py       mapa COD_SAUDE → EMPRESA
consolidacao/leitura.py        ler_aba + leitores TITULAR / DEPENDENTE / suplementar
consolidacao/processamento.py  processar_arquivo, listar_arquivos_entrada, processar_arquivos
consolidacao/exportacao.py     Excel formatado
tratamento.py                  CLI
tests/fabrica.py               gera .xlsx/.csv sintéticos
tests/test_*.py                um por módulo
tests/paridade/rodar_paridade.py, comparar.py, gerar_sinteticos.py
requirements.txt, pytest.ini, .gitignore
```

---

### Task 1: Base do projeto + cabeçalhos

**Files:**
- Create: `.gitignore`, `requirements.txt`, `pytest.ini`, `consolidacao/__init__.py`, `consolidacao/erros.py`, `consolidacao/config.py`, `consolidacao/cabecalhos.py`
- Test: `tests/test_cabecalhos.py`

**Interfaces:**
- Produces:
  - `erros.ErroValidacao(Exception)`: toda falha de regra do R é levantada com esta classe, e a mensagem é exatamente o texto do `stop()` do R.
  - `config`: `ORDEM_INICIAL`, `COLUNAS_OBRIGATORIAS_TITULAR`, `COLUNAS_OBRIGATORIAS_DEPENDENTE`, `COLUNAS_OBRIGATORIAS_SUPLEMENTAR`, `COLUNAS_OPCIONAIS_TITULAR`, `COLUNAS_OPCIONAIS_DEPENDENTE`, `LARGURA_MANUAL: dict[str, float]`, `PADRAO_IDENTIFICADOR`, `NOMES_SAIDA_LEGADOS`, `CODIGO_UNICO = "CÓDIGO UNICO"`, `nome_saida(hoje: date) -> str`, `pasta_padrao(raiz: Path, hoje: date) -> Path`. Valores copiados de `TRATAMENTO.R:46-119`. O mês **não** leva zero à esquerda (`PC_CONSOLIDADO_1_2027.xlsx`).
  - `cabecalhos.normalizar_nome(x: str | None) -> str`
  - `cabecalhos.renomear_e_validar_cabecalhos(df, regras: dict[str, str], contexto: str) -> DataFrame`
  - `cabecalhos.validar_colunas(df, obrigatorias: list[str], contexto: str) -> None`
  - `cabecalhos.validar_identificadores_como_texto(df, contexto: str) -> DataFrame`
  - `cabecalhos.chave_certificado(valores: pd.Series) -> pd.Series`
  - `cabecalhos.unificar_colunas(df, nome_final: str, fontes: list[str], contexto: str) -> DataFrame`

- [ ] **Step 1: Inicializar o repositório e o ambiente**

Rodar `git init`. Criar o `.gitignore` com: `ignore/`, `.venv/`, `.pytest_tmp/`, `__pycache__/`, `tests/paridade/_trabalho/`, `tests/paridade/_sinteticos/`, `dados_teste/`, `*.xlsx`, `.Rhistory`, `.Rproj.user/`. Criar o `requirements.txt` (`pandas>=2.2`, `openpyxl>=3.1`, `pytest>=8`) e o `pytest.ini` (`[pytest]`, `testpaths = tests`, `addopts = -q --basetemp=.pytest_tmp`). Depois: `python -m venv .venv; .venv\Scripts\python -m pip install -r requirements.txt`.
Esperado: instalação sem erro, e `git status` não mostra `ignore/`.

- [ ] **Step 2: Escrever os testes que falham**

```python
def test_normalizar_nome_casos_sujos():
    assert normalizar_nome("  Região ") == "REGIAO"
    assert normalizar_nome("Nº  do Cartão") == "N_DO_CARTAO"
    assert normalizar_nome("DATA DE NASCIMENTO Y2K") == "DATA_DE_NASCIMENTO_Y2K"
    assert normalizar_nome("__x__") == "X"
    assert normalizar_nome(None) == ""

def test_renomear_remove_tipo_de_registro_e_aplica_regras():
    df = pd.DataFrame(columns=["TIPO DE REGISTRO", "Região", "Sexo do Segurado"])
    out = renomear_e_validar_cabecalhos(df, {"REGI.O": "REGIAO", "SEXO.*SEGURADO": "SEXO"}, "T")
    assert list(out.columns) == ["REGIAO", "SEXO"]

def test_renomear_coluna_sem_nome():
    with pytest.raises(ErroValidacao, match=r"^\[T\] Existe coluna sem nome; layout nao seguro\.$"):
        renomear_e_validar_cabecalhos(pd.DataFrame(columns=["A", ""]), {}, "T")

def test_renomear_padrao_em_varias_colunas():
    df = pd.DataFrame(columns=["DATA REATIVACAO", "REATIVADO"])
    with pytest.raises(ErroValidacao, match=r"O padrao 'REATIVA' corresponde a varias colunas: DATA REATIVACAO, REATIVADO"):
        renomear_e_validar_cabecalhos(df, {"REATIVA": "DATA DE REATIVACAO"}, "T")

def test_renomear_duplicado_apos_padronizar():
    with pytest.raises(ErroValidacao, match="A padronizacao gerou cabecalhos duplicados: A B"):
        renomear_e_validar_cabecalhos(pd.DataFrame(columns=["A-B", "a b"]), {}, "T")

def test_validar_colunas_faltantes():
    with pytest.raises(ErroValidacao, match=r"Coluna\(s\) obrigatoria\(s\) ausente\(s\): X, Y"):
        validar_colunas(pd.DataFrame(columns=["A"]), ["A", "X", "Y"], "T")

def test_identificador_com_data_no_nome_falha():
    df = pd.DataFrame({"DATA DO CARTAO": [None]})
    with pytest.raises(ErroValidacao, match="A coluna identificadora 'DATA DO CARTAO' nao foi lida como texto"):
        validar_identificadores_como_texto(df, "T")

def test_unificar_colunas_preenche_e_remove_alias():
    df = pd.DataFrame({"X": ["1", None, None], "NUMERO DO CARTAO": [None, "2", " "], "ALIAS": ["1", "3", "4"]})
    out = unificar_colunas(df, "NUMERO DO CARTAO", ["ALIAS"], "T")
    assert out["NUMERO DO CARTAO"].tolist() == ["1", "2", "4"]
    assert list(out.columns) == ["X", "NUMERO DO CARTAO"]

def test_unificar_colunas_conflito():
    df = pd.DataFrame({"F": ["1"], "A": ["2"]})
    with pytest.raises(ErroValidacao, match=r"Valores conflitantes nas colunas equivalentes usadas para 'F' \(registro 1\)"):
        unificar_colunas(df, "F", ["A"], "T")
```

- [ ] **Step 3: Rodar e confirmar a falha**
Run: `.venv\Scripts\python -m pytest tests/test_cabecalhos.py`
Esperado: FAIL por ImportError.

- [ ] **Step 4: Implementar `config.py`, `erros.py` e `cabecalhos.py`**

Regras que o implementador precisa respeitar:
- `normalizar_nome`: aplicar `.upper()`, depois `strip()`, depois a tabela fixa do R (`ÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇ` → `AAAAAEEEEIIIIOOOOOUUUUC`, via `str.maketrans`, **sem** `unicodedata`), depois `[^A-Z0-9]+` → `_` e por fim remover `_` das pontas.
- Nome final da coluna = normalizado com `_+` trocado por espaço e `strip()`. Regras de regex aplicadas com `re.search` sobre o nome **normalizado**.
- `validar_identificadores_como_texto`: falha se a coluna casa `PADRAO_IDENTIFICADOR` e (o nome normalizado contém `DATA` **ou** algum valor não é `str`/`None`). Isso reproduz o R, em que colunas DATA nunca são character.
- `chave_certificado`: `strip()`; vazio vira `None`.
- `unificar_colunas`: ordem de origem = `[nome_final] + fontes` (sem repetir), só as presentes. "Preenchido" = não `None` e `strip() != ""`. Comparação de conflito com `strip()`. Se `nome_final` já existe, mantém a posição; senão, vai para o fim. O valor gravado é o original, sem strip.

- [ ] **Step 5: Rodar e confirmar que passa**
Run: `.venv\Scripts\python -m pytest tests/test_cabecalhos.py`
Esperado: todos PASS.

- [ ] **Step 6: Commit**
`git add .gitignore requirements.txt pytest.ini consolidacao tests` e `git commit -m "feat: base do projeto e tratamento de cabecalhos"`

---

### Task 2: Datas

**Files:**
- Create: `consolidacao/datas.py`
- Test: `tests/test_datas.py`

**Interfaces:**
- Consumes: `cabecalhos.normalizar_nome`, `ErroValidacao`
- Produces:
  - `converter_serial_excel(serial: float) -> date | None`
  - `SEM_FORMATO` (sentinela) e `converter_data_com_ano_curto(valor: str) -> date | None | SEM_FORMATO`
  - `parse_formato_r(valor: str, formato: str) -> date | None`
  - `texto_celula(v) -> str | None`
  - `converter_celula_data(v, nome_coluna: str, linha: int, contexto: str) -> tuple[date | None, str | None]` (data, observação)
  - `validar_colunas_data(df, contexto: str) -> tuple[DataFrame, list[str], list[str]]` (df, observações, divergências)
  - `resumir_observacoes_datas(obs: list[str], max_amostras: int = 4) -> str | None`

- [ ] **Step 1: Escrever os testes que falham**

```python
def test_serial_excel():
    assert converter_serial_excel(45000) == date(2023, 3, 15)
    assert converter_serial_excel(59) == date(1900, 2, 28)
    assert converter_serial_excel(61) == date(1900, 3, 1)
    assert converter_serial_excel(60) is None
    assert converter_serial_excel(0) is None
    assert converter_serial_excel(2958466) is None
    assert converter_serial_excel(45000.9) == date(2023, 3, 15)

def test_ano_curto():
    assert converter_data_com_ano_curto("05/05/71") == date(1971, 5, 5)
    assert converter_data_com_ano_curto("5-5-29") == date(2029, 5, 5)
    assert converter_data_com_ano_curto("05.05.30 10:00") == date(1930, 5, 5)
    assert converter_data_com_ano_curto("31/02/71") is None
    assert converter_data_com_ano_curto("05/05/1971") is SEM_FORMATO

def test_parse_formato_r_leniente():
    assert parse_formato_r("05/05/2020 10:00", "%d/%m/%Y") == date(2020, 5, 5)   # ignora sobra
    assert parse_formato_r("2020-1-5", "%Y-%m-%d") == date(2020, 1, 5)
    assert parse_formato_r("20200105", "%Y%m%d") == date(2020, 1, 5)
    assert parse_formato_r("31/02/2020", "%d/%m/%Y") is None

def test_converter_celula_data_tipos():
    assert converter_celula_data(datetime(2020, 1, 2, 13, 0), "D", 1, "C") == (date(2020, 1, 2), None)
    assert converter_celula_data(45000, "D", 1, "C") == (date(2023, 3, 15), None)
    assert converter_celula_data(" 45000 ", "D", 1, "C") == (date(2023, 3, 15), None)
    assert converter_celula_data("45000,0", "D", 1, "C") == (date(2023, 3, 15), None)
    assert converter_celula_data(None, "D", 1, "C") == (None, None)
    assert converter_celula_data("  ", "D", 1, "C") == (None, None)
    assert converter_celula_data("abc", "D", 3, "CTX") == (
        None, "CTX | coluna 'D' | registro 3 | valor invalido 'abc' (convertido para vazio)")
    assert converter_celula_data(-5, "D", 1, "C")[1].endswith("valor invalido '-5' (convertido para vazio)")

def test_y2k_substitui_quando_dia_mes_coincidem_e_registra_divergencia():
    df = pd.DataFrame({"DATA DE NASCIMENTO": ["05/05/20", "01/01/90"],
                       "DATA DE NASCIMENTO Y2K": ["05/05/1920", "02/01/1990"]})
    out, obs, div = validar_colunas_data(df, "CTX")
    assert out["DATA DE NASCIMENTO"].tolist() == [date(1920, 5, 5), date(1990, 1, 1)]
    assert obs == []
    assert div == ["CTX | registro 2 | divergencia entre 'DATA DE NASCIMENTO' (01/01/1990) e "
                   "'DATA DE NASCIMENTO Y2K' (02/01/1990); mantida a data principal"]

def test_resumir_observacoes():
    obs = [f"C | coluna 'D' | registro {i} | valor invalido 'x' (convertido para vazio)" for i in range(1, 7)]
    r = resumir_observacoes_datas(obs)
    assert r.startswith("C | coluna 'D': 6 valor(es); amostras: registro 1 |")
    assert r.endswith("; +2 caso(s)")
    assert resumir_observacoes_datas([]) is None
```

- [ ] **Step 2: Rodar e confirmar a falha**
Run: `.venv\Scripts\python -m pytest tests/test_datas.py`. Esperado: FAIL (ImportError).

- [ ] **Step 3: Implementar `consolidacao/datas.py`**

Siga `TRATAMENTO.R:228-457`. Decisões que o R não deixa óbvias:
- `parse_formato_r` imita o `strptime` do R: converter o formato em regex ancorada só no início (`re.match`, sem `$`). Usar `%d`/`%m`/`%H`/`%M`/`%S` → `(\d{1,2}+)` e `%Y` → `(\d{1,4}+)` (quantificadores possessivos, sem backtracking como no R). Espaço no formato vira `\s*`; os demais literais são escapados. Validar a data com `date(...)` (ValueError → `None`), além de H ≤ 24, M ≤ 59 e S ≤ 61. Ano 0 → `None`.
- Formatos, na ordem: `"%d/%m/%Y", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d-%m-%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y%m%d", "%d%m%Y"`.
- Tipos de célula: `datetime` → `.date()`; `date` → ela mesma; `time` → `date(1899, 12, 31)`; `bool` → texto `"TRUE"`/`"FALSE"` (cai como inválido); `int`/`float` → serial. No ramo numérico, o valor da observação **não** sofre strip.
- `texto_celula` para número imita `format(x, scientific=FALSE)` do R: inteiro exato → `str(int(x))`; senão, arredondar para `max(0, 7 - dígitos_da_parte_inteira)` casas e tirar zeros finais.
- Colunas de data = nome normalizado contém `DATA`. A conversão roda em todas elas, e o resultado é coluna de `date | None`.
- Y2K: para cada coluna cujo nome normalizado termina em `_Y2K`, a base é a coluna com nome normalizado igual sem `_Y2K` e que contém `DATA`. Mais de uma base → `ErroValidacao("[ctx] Mais de uma coluna de data corresponde a '<nome>'.")`. O número do registro na divergência é 1-based, e as datas usam `%d/%m/%Y`.
- `resumir_observacoes_datas`: chave = texto antes de `" | registro N | "`; grupos em ordem alfabética da chave; amostra = parte a partir de `registro N | ...`; grupos unidos por `" || "`.

- [ ] **Step 4: Rodar e confirmar que passa**
Run: `.venv\Scripts\python -m pytest tests/test_datas.py`. Esperado: PASS.

- [ ] **Step 5: Commit**
`git commit -am "feat: conversao de datas equivalente ao R"` (com `git add consolidacao/datas.py tests/test_datas.py` antes).

---

### Task 3: Mapa de empresas

**Files:**
- Create: `consolidacao/empresas.py`, `tests/fabrica.py` (só `escrever_csv_empresas` por enquanto)
- Test: `tests/test_empresas.py`

**Interfaces:**
- Produces:
  - `carregar_mapa_empresas(caminho: Path) -> dict[str, str]` (COD_SAUDE → EMPRESA, ordem do arquivo)
  - `empresa_da_referencia(referencia: str, mapa: dict[str, str]) -> tuple[str, str]` (código, empresa)
  - `tests.fabrica.escrever_csv_empresas(caminho: Path, linhas: list[tuple[str, str]], encoding: str = "utf-8-sig") -> None`: grava `EMPRESA;COD_SAUDE;COD_DENTAL`.

- [ ] **Step 1: Escrever os testes que falham**

```python
def test_mapa_basico(tmp_path):
    p = tmp_path / "e.csv"; escrever_csv_empresas(p, [("ACME", "123"), (" BETA ", " 9 "), ("", "7"), ("NA", "8")])
    assert carregar_mapa_empresas(p) == {"123": "ACME", "9": "BETA"}

def test_mapa_cp1252(tmp_path):
    p = tmp_path / "e.csv"; escrever_csv_empresas(p, [("JOÃO SAÚDE", "1")], encoding="cp1252")
    assert carregar_mapa_empresas(p) == {"1": "JOÃO SAÚDE"}

def test_mapa_codigo_com_duas_empresas(tmp_path):
    p = tmp_path / "e.csv"; escrever_csv_empresas(p, [("A", "1"), ("B", "1")])
    with pytest.raises(ErroValidacao, match="COD_SAUDE '1' esta associado a mais de uma EMPRESA: A, B"):
        carregar_mapa_empresas(p)

def test_mapa_sem_colunas(tmp_path):
    p = tmp_path / "e.csv"; p.write_text("X;Y\n1;2\n", encoding="utf-8")
    with pytest.raises(ErroValidacao, match="O COD_EMPRESAS.csv precisa conter as colunas EMPRESA e COD_SAUDE."):
        carregar_mapa_empresas(p)

def test_empresa_da_referencia():
    assert empresa_da_referencia("pc123 10.2026", {"123": "ACME"}) == ("123", "ACME")
    with pytest.raises(ErroValidacao, match="Nao foi possivel extrair o codigo apos 'PC' da referencia 'X1'."):
        empresa_da_referencia("X1", {})
    with pytest.raises(ErroValidacao, match="Codigo '5' extraido de 'PC5' nao encontrado em COD_SAUDE no COD_EMPRESAS.csv."):
        empresa_da_referencia("PC5", {"1": "A"})
```

- [ ] **Step 2: Rodar e confirmar a falha.** `.venv\Scripts\python -m pytest tests/test_empresas.py`. Esperado: FAIL.

- [ ] **Step 3: Implementar `consolidacao/empresas.py`**
Usar o módulo `csv` com `delimiter=";"` e `quotechar='"'`. Decodificar com `utf-8-sig` e, em `UnicodeDecodeError`, com `cp1252` (escolha do Claude, Review Focus 5). Campo `"NA"` = ausente (imita `read.csv2`). Os demais erros e mensagens vêm de `TRATAMENTO.R:522-610`. Arquivo inexistente → `"Arquivo de empresas nao encontrado: <caminho>"`. Arquivo vazio → `"O COD_EMPRESAS.csv esta vazio."`. Nenhum par válido → `"O COD_EMPRESAS.csv nao tem pares EMPRESA/COD_SAUDE validos."`. Os nomes de coluna são comparados via `normalizar_nome`.

- [ ] **Step 4: Rodar e confirmar que passa.** Esperado: PASS.

- [ ] **Step 5: Commit** `git add consolidacao/empresas.py tests/fabrica.py tests/test_empresas.py` e `git commit -m "feat: mapa de empresas"`

---

### Task 4: Leitura das abas

**Files:**
- Create: `consolidacao/leitura.py`
- Modify: `tests/fabrica.py` (adicionar `criar_arquivo_pc`)
- Test: `tests/test_leitura.py`

**Interfaces:**
- Consumes: Tasks 1–2.
- Produces:
  - `ler_aba(caminho: Path, aba: str, skip: int = 0) -> DataFrame`
  - `@dataclass ResultadoAba: dados: DataFrame; avisos: list[str]; observacoes_datas: list[str]; divergencias_datas: list[str]`
  - `ler_titular(caminho, aba, referencia) -> ResultadoAba`
  - `ler_dependente(caminho, aba, referencia) -> ResultadoAba`
  - `ler_suplementar(caminho, aba, referencia) -> ResultadoAba` (avisos = `[]`; `dados` = colunas `.CHAVE_CERTIFICADO`, `.CARTAO_SUP`, `.NOME_MAE_SUP`)
  - `tests.fabrica.criar_arquivo_pc(caminho: Path, titulares: list[dict], dependentes: list[dict], suplementar: list[dict], abas: tuple[str, ...] = ("TITULAR", "DEPENDENTE", "POS. CADASTRAL (NOME MÃE)"), linhas_vazias_formatadas_no_fim: int = 0) -> None`. Na aba TITULAR, grava 2 linhas de lixo (`"RELATORIO"` e uma vazia) e depois o cabeçalho. Na DEPENDENTE, grava cabeçalho, dados e 2 linhas de resumo (`"TOTAL"`, `n`). Na suplementar, cabeçalho e dados. As colunas são as chaves do primeiro dict. Os valores são gravados com o tipo Python recebido (int/float/datetime/str). `linhas_vazias_formatadas_no_fim` aplica fill/borda em células vazias abaixo do resumo do DEPENDENTE.
  - Constantes em `tests/fabrica.py`, usadas por todos os testes: `TITULAR_OK` (colunas `TIPO DE REGISTRO`, `NUMERO DA SUBFATURA`, `NUMERO DO CERTIFICADO`, `NOME DO SEGURADO`, `PLANO`, `DATA DE NASCIMENTO`, `REGIÃO`, `SEXO DO SEGURADO`, `NUMERO DO CPF`), `DEPENDENTE_OK` (`NUMERO DO CERTIFICADO`, `CODIGO DO DEPENDENTE`, `NOME DEPENDENTE`, `GRAU PARENTESCO`, `NOME DA MÃE DO DEPENDENTE`, `CARTÃO DO DEPENDENTE`, `SEXO DO DEPENDENTE`, `DATA DE NASCIMENTO`) e `SUPLEMENTAR_OK` (`NUMERO DO CERTIFICADO`, `CARTÃO`, `NOME DA MÃE`), com 2 titulares (certificados `"001"` e `"002"`), 2 dependentes (um do `"001"`, outro do `"999"`, sem titular) e suplementar para `"001"`.

- [ ] **Step 1: Escrever os testes que falham**

```python
def test_titular_basico(tmp_path):
    p = tmp_path / "PC1.xlsx"; criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    r = ler_titular(p, "TITULAR", "PC1")
    d = r.dados
    assert "TIPO DE REGISTRO" not in d.columns
    assert {"REGIAO", "SEXO"} <= set(d.columns)
    assert d["CODIGO DO DEPENDENTE"].tolist() == ["00", "00"]
    assert d["TITULARIDADE"].tolist() == ["TITULAR", "TITULAR"]
    assert d[".CHAVE_CERTIFICADO"].tolist() == ["001", "002"]
    assert r.avisos == ["DATA DE REATIVACAO", "CARGO DE OCUPACAO", "NUMERO DA MATRICULA"]

def test_identificador_numerico_vira_texto_sem_ponto_zero(tmp_path):
    tit = [dict(TITULAR_OK[0], **{"NUMERO DO CPF": 12345678901, "NUMERO DA SUBFATURA": 7.0})]
    p = tmp_path / "PC1.xlsx"; criar_arquivo_pc(p, tit, DEPENDENTE_OK, SUPLEMENTAR_OK)
    d = ler_titular(p, "TITULAR", "PC1").dados
    assert d["NUMERO DO CPF"].tolist() == ["12345678901"]
    assert d["NUMERO DA SUBFATURA"].tolist() == ["7"]

def test_titular_certificado_duplicado_ou_vazio(tmp_path):
    tit = [dict(TITULAR_OK[0]), dict(TITULAR_OK[1], **{"NUMERO DO CERTIFICADO": None})]
    p = tmp_path / "PC1.xlsx"; criar_arquivo_pc(p, tit, DEPENDENTE_OK, SUPLEMENTAR_OK)
    with pytest.raises(ErroValidacao, match=r"\[TITULAR: PC1\] Ha titular sem NUMERO DO CERTIFICADO"):
        ler_titular(p, "TITULAR", "PC1")

def test_dependente_remove_duas_linhas_de_resumo(tmp_path):
    p = tmp_path / "PC1.xlsx"; criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    d = ler_dependente(p, "DEPENDENTE", "PC1").dados
    assert len(d) == 2
    assert {"NUMERO DO CARTAO DO DEPENDENTE", "NOME DA MAE DO DEPENDENTE", ".SUBF_DEPENDENTE"} <= set(d.columns)
    assert d["TITULARIDADE"].tolist() == ["DEPENDENTE", "DEPENDENTE"]

def test_dependente_ignora_linhas_vazias_formatadas_no_fim(tmp_path):
    p = tmp_path / "PC1.xlsx"
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK, linhas_vazias_formatadas_no_fim=5)
    assert len(ler_dependente(p, "DEPENDENTE", "PC1").dados) == 2

def test_dependente_sem_dados(tmp_path):
    p = tmp_path / "PC1.xlsx"; criar_arquivo_pc(p, TITULAR_OK, [], SUPLEMENTAR_OK)
    with pytest.raises(ErroValidacao, match=r"\[DEPENDENTE: PC1\] Ha 2 linha\(s\); nao existem dados"):
        ler_dependente(p, "DEPENDENTE", "PC1")

def test_dependente_com_campo_do_titular(tmp_path):
    dep = [dict(r, PLANO="X") for r in DEPENDENTE_OK]
    p = tmp_path / "PC1.xlsx"; criar_arquivo_pc(p, TITULAR_OK, dep, SUPLEMENTAR_OK)
    with pytest.raises(ErroValidacao, match="tambem aparecem na aba dependente: PLANO"):
        ler_dependente(p, "DEPENDENTE", "PC1")

def test_suplementar_agrega_repeticoes_e_falha_em_conflito(tmp_path):
    sup = SUPLEMENTAR_OK + [dict(SUPLEMENTAR_OK[0]), {"NUMERO DO CERTIFICADO": None, "CARTÃO": None, "NOME DA MÃE": None}]
    p = tmp_path / "PC1.xlsx"; criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, sup)
    d = ler_suplementar(p, "POS. CADASTRAL (NOME MÃE)", "PC1").dados
    assert list(d.columns) == [".CHAVE_CERTIFICADO", ".CARTAO_SUP", ".NOME_MAE_SUP"] and len(d) == 1
    conflito = SUPLEMENTAR_OK + [dict(SUPLEMENTAR_OK[0], **{"CARTÃO": "OUTRO"})]
    criar_arquivo_pc(p, TITULAR_OK, DEPENDENTE_OK, conflito)
    with pytest.raises(ErroValidacao, match="O certificado '001' tem valores conflitantes em 'CARTAO'"):
        ler_suplementar(p, "POS. CADASTRAL (NOME MÃE)", "PC1")
```

- [ ] **Step 2: Rodar e confirmar a falha.** `.venv\Scripts\python -m pytest tests/test_leitura.py`. Esperado: FAIL.

- [ ] **Step 3: Implementar `consolidacao/leitura.py`**
`ler_aba` imita o `read_excel` do readxl:
- `openpyxl.load_workbook(caminho, read_only=True, data_only=True)`.
- Pular `skip` linhas, depois pular linhas totalmente vazias. A próxima linha é o cabeçalho.
- Linhas totalmente vazias no fim são descartadas. "Vazia" = `None` ou string só com espaço; formatação não conta.
- Colunas: da A até a última coluna com algum valor no cabeçalho ou nos dados.
- Cabeçalho `None` → `""`; cabeçalho numérico → texto.
- Coluna cujo nome normalizado **não** contém `DATA`: tudo vira texto. `str` com `strip()` (vazio → `None`); `int` e `float` inteiros → `str(int(x))`; outros floats → `format(x, ".15g")`; `bool` → `"TRUE"`/`"FALSE"`; `datetime` → serial Excel em texto (`format(serial, ".15g")`).
- Coluna DATA: valor bruto (`str` com strip, vazio → `None`).
- Zero colunas → `ErroValidacao("A aba '<aba>' nao tem cabecalhos identificaveis.")`.

Leitores: seguir `TRATAMENTO.R:617-784` na mesma ordem de passos (a ordem decide qual erro aparece primeiro). Regras de renomeação copiadas verbatim (`TRATAMENTO.R:620-625`, `668-674`, `771-774`). O suplementar descarta linhas totalmente vazias **depois** da conversão de datas. `ler_dependente` com `NUMERO DA SUBFATURA` → renomeia para `.SUBF_DEPENDENTE`; sem a coluna, cria `.SUBF_DEPENDENTE = None`.

- [ ] **Step 4: Rodar e confirmar que passa.** Esperado: PASS.

- [ ] **Step 5: Commit** `git add consolidacao/leitura.py tests/fabrica.py tests/test_leitura.py` e `git commit -m "feat: leitura das abas titular, dependente e suplementar"`

---

### Task 5: Processar um arquivo

**Files:**
- Create: `consolidacao/processamento.py` (parte 1)
- Test: `tests/test_processar_arquivo.py`

**Interfaces:**
- Consumes: `ler_titular`, `ler_dependente`, `ler_suplementar`, `unificar_colunas`, `empresa_da_referencia`.
- Produces: `@dataclass ResultadoArquivo: arquivo: str; referencia: str; status: str; mensagem: str; avisos: str | None; observacoes_datas: list[str]; divergencias_datas: list[str]; final: DataFrame | None; pendencias: DataFrame | None; n_titulares: int | None; n_dependentes: int | None; codigo_empresa: str | None; empresa: str | None`, e `processar_arquivo(caminho: Path, mapa: dict[str, str]) -> ResultadoArquivo`.

- [ ] **Step 1: Escrever os testes que falham**

```python
def _arquivo(tmp_path, nome="PC123 10.2026.xlsx", **kw):
    p = tmp_path / nome
    criar_arquivo_pc(p, kw.get("tit", TITULAR_OK), kw.get("dep", DEPENDENTE_OK), kw.get("sup", SUPLEMENTAR_OK),
                     **({"abas": kw["abas"]} if "abas" in kw else {}))
    return p

MAPA = {"123": "ACME"}

def test_processar_ok(tmp_path):
    r = processar_arquivo(_arquivo(tmp_path), MAPA)
    assert (r.status, r.mensagem, r.empresa, r.codigo_empresa) == ("OK", "Processado com sucesso.", "ACME", "123")
    assert (r.n_titulares, r.n_dependentes) == (2, 2)
    f = r.final
    assert "SubF" in f.columns and "NUMERO DA SUBFATURA" not in f.columns
    assert ".CHAVE_CERTIFICADO" in f.columns          # fiel ao R: a coluna auxiliar sai na base
    dep = f[f["TITULARIDADE"] == "DEPENDENTE"]
    assert dep.iloc[0]["NOME DO SEGURADO"] == TITULAR_OK[0]["NOME DO SEGURADO"]
    assert f["CÓDIGO UNICO"].tolist()[:2] == ["00100", "00200"]
    tit001 = f[(f["TITULARIDADE"] == "TITULAR") & (f[".CHAVE_CERTIFICADO"] == "001")].iloc[0]
    assert tit001["NUMERO DO CARTAO"] == SUPLEMENTAR_OK[0]["CARTÃO"]
    assert r.pendencias["MOTIVO"].tolist() == ["CERTIFICADO SEM TITULAR CORRESPONDENTE"]
    assert "Titulares sem registro na aba suplementar: 1" in r.avisos

def test_aba_faltando(tmp_path):
    r = processar_arquivo(_arquivo(tmp_path, abas=("TITULAR", "DEPENDENTE", "OUTRA")), MAPA)
    assert r.status == "ERRO"
    assert r.mensagem == "Esperada exatamente uma aba POS. CADASTRAL (NOME M E); encontrada(s) 0."
    assert r.final is None and r.n_titulares is None and r.codigo_empresa is None

def test_certificado_duplicado(tmp_path):
    r = processar_arquivo(_arquivo(tmp_path, tit=[TITULAR_OK[0], TITULAR_OK[0]]), MAPA)
    assert r.mensagem == "Certificado(s) duplicado(s) na aba TITULAR: 001"

def test_observacoes_parciais_sobrevivem_ao_erro(tmp_path):
    tit = [dict(TITULAR_OK[0], **{"DATA DE NASCIMENTO": "xx"}), TITULAR_OK[1]]
    r = processar_arquivo(_arquivo(tmp_path, tit=tit, dep=[]), MAPA)
    assert r.status == "ERRO" and len(r.observacoes_datas) == 1
```

- [ ] **Step 2: Rodar e confirmar a falha.** Esperado: FAIL.

- [ ] **Step 3: Implementar `processar_arquivo`**
Seguir `TRATAMENTO.R:791-1029`, passo a passo e na mesma ordem.
- Detecção de abas: `re.search` sobre `normalizar_nome(aba)` com `TITULAR`, `DEPENDENT` e `POS.*CADASTRAL.*NOME.*M.*E`.
- As observações e divergências de cada leitor só são acumuladas depois que o leitor **termina** (igual ao R). O `except Exception as e` usa `mensagem = str(e)`.
- `left_join` = `merge(how="left")`, conferindo a contagem de linhas. No dependente, chave `None` nunca casa.
- `coalesce` do SubF: usar o do titular, a não ser que seja `None`.
- `bind_rows`: `pd.concat` com a união de colunas, na ordem titular e depois colunas novas do dependente, com ausentes = `None`.
- `CÓDIGO UNICO` = certificado original + código, só quando os dois têm `strip() != ""`.
- `avisos`: unidos por `" | "`, na ordem e com os textos de `TRATAMENTO.R:975-995`. Sem avisos → `None`.

- [ ] **Step 4: Rodar e confirmar que passa.** Esperado: PASS.

- [ ] **Step 5: Commit** `git commit -m "feat: processamento de um arquivo"` (com add dos arquivos).

---

### Task 6: Consolidação de vários arquivos + log

**Files:**
- Modify: `consolidacao/processamento.py` (parte 2)
- Test: `tests/test_consolidacao.py`

**Interfaces:**
- Consumes: `processar_arquivo`, `resumir_observacoes_datas`, `config`.
- Produces:
  - `listar_arquivos_entrada(pasta: Path, hoje: date) -> list[Path]`
  - `@dataclass Consolidado: consolidado: DataFrame; pendencias: DataFrame; log: DataFrame`
  - `processar_arquivos(caminhos: list[Path], mapa: dict[str, str], progresso: Callable[[str], None] | None = None) -> Consolidado`. Chama `progresso(nome_do_arquivo)` antes de cada arquivo.

- [ ] **Step 1: Escrever os testes que falham**

```python
def test_listar_arquivos_entrada_filtra(tmp_path):
    for n in ["PC2 a.xlsx", "PC1 b.XLSX", "~$PC1 b.xlsx", "PC_CONSOLIDADO_9_2025.xlsx",
              "Consolidado - tratado.xlsx", "nota.txt"]:
        (tmp_path / n).write_bytes(b"x")
    (tmp_path / "sub.xlsx").mkdir()
    assert [p.name for p in listar_arquivos_entrada(tmp_path, date(2026, 10, 22))] == ["PC1 b.XLSX", "PC2 a.xlsx"]

def test_listar_pasta_inexistente_e_vazia(tmp_path):
    with pytest.raises(ErroValidacao, match="A pasta de entrada nao existe: "):
        listar_arquivos_entrada(tmp_path / "x", date(2026, 10, 22))
    with pytest.raises(ErroValidacao, match="Nenhum arquivo .xlsx de entrada encontrado em: "):
        listar_arquivos_entrada(tmp_path, date(2026, 10, 22))

def test_processar_arquivos_mistura_ok_e_erro(tmp_path):
    a = tmp_path / "PC123 x.xlsx"; criar_arquivo_pc(a, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    b = tmp_path / "PC999 y.xlsx"; criar_arquivo_pc(b, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    vistos = []
    c = processar_arquivos([a, b], {"123": "ACME"}, progresso=vistos.append)
    assert vistos == ["PC123 x.xlsx", "PC999 y.xlsx"]
    assert c.log["STATUS"].tolist() == ["OK", "ERRO"]
    assert list(c.log.columns) == ["ARQUIVO", "REFERENCIA", "CODIGO_EMPRESA", "EMPRESA", "STATUS", "TITULARES",
        "DEPENDENTES", "AVISOS_ESTRUTURA", "DATAS_INVALIDAS", "OBSERVACOES_DATAS", "QTD_DIVERGENCIAS_DATAS",
        "DIVERGENCIAS_DATAS", "MENSAGEM", "OBSERVACAO_IDENTIFICADORES"]
    assert list(c.consolidado.columns[:13]) == ORDEM_INICIAL
    assert len(c.consolidado) == 4 and len(c.pendencias) == 1

def test_processar_arquivos_todos_com_erro(tmp_path):
    b = tmp_path / "PC999.xlsx"; criar_arquivo_pc(b, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    c = processar_arquivos([b], {"1": "A"})
    assert len(c.consolidado) == 0 and "CÓDIGO UNICO" in c.consolidado.columns
    assert list(c.pendencias.columns)[-1] == "MOTIVO"
```

- [ ] **Step 2: Rodar e confirmar a falha.** Esperado: FAIL.

- [ ] **Step 3: Implementar**
Seguir `TRATAMENTO.R:1036-1165`.
- `listar_arquivos_entrada`: `*.xlsx` sem diferenciar maiúsculas, sem diretórios. Excluir: nome que começa com `~$`; nome (minúsculo) em `nome_saida(hoje)` ou nos legados; regex `^PC_CONSOLIDADO_.*\.xlsx$` com IGNORECASE. Ordenar por `key=lambda p: (p.name.casefold(), p.name)` (aproxima o `sort()` do R no Windows; a paridade confirma).
- Consolidação: arquivos OK somados um a um. Se a coluna já existe com outro "tipo" (`date` vs `str`), o arquivo vira ERRO com `"Incompatibilidade ao consolidar este arquivo com os anteriores: <detalhe>"`.
- Ordenação final: estável, por `REFERENCIA`, `SubF`, `NUMERO DO CERTIFICADO`, comparando strings como bytes (o `arrange` do dplyr usa locale C) e com `None` por último. Colunas de `ORDEM_INICIAL` presentes vêm primeiro.
- Bases vazias: mesmas colunas de `base_final_vazia` e `pendencias_vazias`.
- Log: `TITULARES`/`DEPENDENTES` são `int` ou `None`; `OBSERVACAO_IDENTIFICADORES` é o texto fixo de `TRATAMENTO.R:1160`.

- [ ] **Step 4: Rodar e confirmar que passa.** Esperado: PASS.

- [ ] **Step 5: Commit** `git commit -m "feat: consolidacao e log"`

---

### Task 7: Exportação Excel + CLI

**Files:**
- Create: `consolidacao/exportacao.py`, `tratamento.py`
- Test: `tests/test_exportacao.py`, `tests/test_cli.py`

**Interfaces:**
- Consumes: `Consolidado`, `processar_arquivos`, `listar_arquivos_entrada`, `carregar_mapa_empresas`, `config`.
- Produces:
  - `escrever_excel(c: Consolidado, destino: Path) -> None`
  - `tratamento.main(argv: list[str] | None = None) -> int` (0 = rodou; erro fatal de entrada → mensagem e 1). Argumentos: `--entrada`, `--empresas`, `--saida`. Padrões: `pasta_padrao(raiz, hoje)`, `raiz / "COD_EMPRESAS.csv"` e `entrada / nome_saida(hoje)`, com `raiz = Path(__file__).parent`.

- [ ] **Step 1: Escrever os testes que falham**

```python
def test_excel_formatado(tmp_path):
    a = tmp_path / "PC123 x.xlsx"; criar_arquivo_pc(a, TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    destino = tmp_path / "out.xlsx"
    escrever_excel(processar_arquivos([a], {"123": "ACME"}), destino)
    wb = openpyxl.load_workbook(destino)
    assert wb.sheetnames == ["Base Tratada", "Pendencias", "Log"]
    ws = wb["Base Tratada"]
    assert "TabelaBaseTratada" in ws.tables and ws.tables["TabelaBaseTratada"].tableStyleInfo.name == "TableStyleMedium2"
    assert ws.freeze_panes == "A2" and ws.row_dimensions[1].height == 32
    assert ws["A1"].fill.fgColor.rgb.endswith("4472C4") and ws["A1"].font.bold
    col = {c.value: c.column_letter for c in ws[1]}
    assert ws.column_dimensions[col["EMPRESA"]].width == 28
    assert ws.column_dimensions[col["DATA DE NASCIMENTO"]].width == 13
    celula = ws[f'{col["DATA DE NASCIMENTO"]}2']
    assert celula.number_format == "dd/mm/yyyy" and isinstance(celula.value, datetime)
    assert ws[f'{col["CODIGO DO DEPENDENTE"]}2'].alignment.horizontal == "right"

def test_excel_pendencias_vazias_sem_tabela(tmp_path):
    a = tmp_path / "PC123 x.xlsx"; criar_arquivo_pc(a, TITULAR_OK, DEPENDENTE_OK[:1], SUPLEMENTAR_OK)
    destino = tmp_path / "out.xlsx"
    escrever_excel(processar_arquivos([a], {"123": "ACME"}), destino)
    ws = openpyxl.load_workbook(destino)["Pendencias"]
    assert not ws.tables and ws.auto_filter.ref == "A1:I1"

def test_cli_gera_arquivo_e_resumo(tmp_path, capsys):
    criar_arquivo_pc(tmp_path / "PC123 x.xlsx", TITULAR_OK, DEPENDENTE_OK, SUPLEMENTAR_OK)
    escrever_csv_empresas(tmp_path / "e.csv", [("ACME", "123")])
    saida = tmp_path / "r.xlsx"
    assert main(["--entrada", str(tmp_path), "--empresas", str(tmp_path / "e.csv"), "--saida", str(saida)]) == 0
    out = capsys.readouterr().out
    assert saida.exists()
    assert "Arquivos de entrada encontrados: 1" in out and "Arquivos processados com sucesso: 1" in out
    assert "Abas: Base Tratada, Pendencias, Log" in out
```

- [ ] **Step 2: Rodar e confirmar a falha.** Esperado: FAIL.

- [ ] **Step 3: Implementar**
`exportacao.py` segue `TRATAMENTO.R:1172-1299`, aplicando estilos nesta ordem:
- Tabela (`TableStyleMedium2`) **só se houver linhas**. Sem linhas: cabeçalho + `auto_filter`.
- Cabeçalho: fill sólido `4472C4`, fonte branca em negrito, centralizado, `wrap_text`, borda inferior fina.
- Larguras: "auto" = `max(len(texto exibido)) + 2` (datas contam 10); depois `LARGURA_MANUAL`; depois toda coluna com `DATA` no nome (sem diferenciar maiúsculas) = 13. Isso inclui colunas do Log, igual ao R.
- Linha 1 com altura 32; corpo com `vertical="center"`.
- Colunas de data (nome com DATA **e** valores `date`): `dd/mm/yyyy` centralizado.
- `CODIGO DO DEPENDENTE` alinhado à direita; lista de centralizadas de `TRATAMENTO.R:1278-1287` centralizada.
- `freeze_panes = "A2"`, `showGridLines = False`. `None` vira célula vazia, e `date` é gravada como data do Excel.

`tratamento.py` imprime exatamente as linhas de `TRATAMENTO.R:1097-1099` e `1330-1354`. Erro de mapa ou de pasta (`ErroValidacao`) → imprime `Erro: <mensagem>` e retorna 1.

- [ ] **Step 4: Rodar a suíte inteira.** `.venv\Scripts\python -m pytest`. Esperado: tudo PASS.

- [ ] **Step 5: Commit** `git commit -m "feat: exportacao Excel e linha de comando"`

---

### Task 8: Paridade com o R

**Files:**
- Create: `tests/paridade/comparar.py`, `tests/paridade/gerar_sinteticos.py`, `tests/paridade/rodar_paridade.py`
- Test: `tests/test_paridade.py`

**Interfaces:**
- Consumes: `tratamento.main`, `tests.fabrica`.
- Produces:
  - `comparar.comparar_planilhas(a: Path, b: Path, max_difs: int = 50) -> list[str]`. Compara nomes de abas, cabeçalhos, nº de linhas e cada valor (datetime → date; números comparados numericamente; `None` == vazio) e o `number_format` de células de data. Ignora larguras e estilos.
  - `gerar_sinteticos.gerar(pasta: Path) -> Path`. Cria cerca de 8 arquivos cobrindo os casos das Tasks 4–6 (OK, aba faltando, duplicado, data inválida, Y2K divergente, conflito suplementar, código inexistente, identificador numérico) + `COD_EMPRESAS.csv`. Retorna a pasta.
  - `rodar_paridade.rodar(entrada: Path, empresas: Path, rscript: Path = Path(r"C:\Program Files\R\R-4.6.1\bin\Rscript.exe")) -> list[str]`, e também uma CLI com os mesmos argumentos (`--entrada`, `--empresas`, `--rscript`), que imprime as diferenças e sai com 0 quando elas não existem.

- [ ] **Step 1: Verificar os pacotes R**
Run: `& "C:\Program Files\R\R-4.6.1\bin\Rscript.exe" -e "for (p in c('readxl','dplyr','purrr','openxlsx','lubridate','here','tibble')) if (!requireNamespace(p, quietly=TRUE)) cat('FALTA', p, '\n')"`
Esperado: nenhuma linha `FALTA`. Se faltar algum pacote, **parar e perguntar ao usuário** antes de instalar.

- [ ] **Step 2: Escrever o teste que falha**

```python
RSCRIPT = Path(r"C:\Program Files\R\R-4.6.1\bin\Rscript.exe")

@pytest.mark.skipif(not RSCRIPT.exists(), reason="R nao instalado")
def test_paridade_sinteticos(tmp_path):
    pasta = gerar(tmp_path / "entrada")
    assert rodar(pasta, pasta / "COD_EMPRESAS.csv", RSCRIPT) == []
```

- [ ] **Step 3: Rodar e confirmar a falha.** Esperado: FAIL (ImportError).

- [ ] **Step 4: Implementar**
`rodar`:
1. Limpar e recriar `tests/paridade/_trabalho/` com `r_raiz/.here` (vazio, para que o `here::here()` ache a raiz), `r_raiz/TRATAMENTO.R` (cópia da raiz do projeto), `r_raiz/COD_EMPRESAS.csv` (cópia de `empresas`) e `r_raiz/PC_CONSOLIDADO_<m>_<a>/` com os `.xlsx` de `entrada`.
2. Rodar `Rscript TRATAMENTO.R` com `cwd=r_raiz` (timeout de 300 s). Se falhar, devolver `["R falhou: <stderr>"]`.
3. Rodar `tratamento.main([...])` com saída `_trabalho/py.xlsx`.
4. Devolver `comparar_planilhas(saida_R, py.xlsx)`.

Toda diferença encontrada é bug no Python: corrigir o módulo dono **e** adicionar um teste unitário que prenda o caso, na task correspondente.

- [ ] **Step 5: Rodar até zerar as diferenças.** `.venv\Scripts\python -m pytest tests/test_paridade.py`. Esperado: PASS.

- [ ] **Step 6: Commit** `git commit -m "test: paridade R x Python com dados sinteticos"`

- [ ] **Step 7: Paridade com arquivos reais (manual, com o usuário)**
O usuário copia planilhas reais para `dados_teste/` (já no `.gitignore`; nunca usar `ignore/`) e roda:
`.venv\Scripts\python tests\paridade\rodar_paridade.py --entrada dados_teste --empresas COD_EMPRESAS.csv`
Esperado: `0 diferencas`. Este é o critério de pronto da fase 1.
