# Fase 1 — Conversão do `tratamento.R` para Python

Data: 2026-10-22
Status: aguardando aprovação final antes de iniciar a implementação.

## Objetivo

Criar `tratamento.py`, uma conversão direta do `tratamento.R`, regra por regra.
Nenhuma lógica nova é introduzida. O `tratamento.R` é a referência: o
tratamento nele é considerado correto, e qualquer divergência entre os dois
scripts é tratada como erro no Python.

Esta fase existe para validar a lógica em Python antes de construir o
aplicativo (ver `fase2-app.md`).

## Contexto

- O processo atual junta várias planilhas de posição cadastral (uma por
  empresa) em um único dataset.
- O `tratamento.R` referencia a pasta `PC_CONSOLIDADO_<mês>_<ano>`, que não
  existe mais no diretório por questões de governança de dados.
- **Restrição de acesso:** a pasta `ignore/` contém dados sigilosos e não pode
  ser lida, listada, pesquisada ou alterada de forma alguma. O bloqueio está
  configurado em `.claude/settings.json`.

## Comportamento a reproduzir (resumo do `tratamento.R`)

### Entradas
- Pasta com um `.xlsx` por empresa. O nome do arquivo, sem a extensão, é a
  REFERENCIA e deve começar com `PC<código>`.
- `COD_EMPRESAS.csv` (separador `;`, UTF-8, BOM opcional) com as colunas
  `EMPRESA` e `COD_SAUDE`. Um código ligado a mais de uma empresa é erro.
- Arquivos ignorados: temporários (`~$*`), saídas anteriores
  (`PC_CONSOLIDADO_*.xlsx`) e nomes de saída legados.

### Abas exigidas em cada arquivo (exatamente uma de cada)
1. **TITULAR**: os cabeçalhos ficam na linha 3. O titular recebe
   `CODIGO DO DEPENDENTE = "00"`. Certificado vazio ou duplicado é erro.
2. **DEPENDENTE**: as 2 últimas linhas são resumo e são descartadas. Os
   campos `NOME DO SEGURADO` e `PLANO` não podem aparecer nesta aba.
3. **POS. CADASTRAL (NOME MÃE)**: `CARTAO` e `NOME DA MAE` por certificado.
   Repetições idênticas são agregadas; valores conflitantes são erro.

### Regras de tratamento
- Cabeçalhos normalizados (maiúsculas, sem acento, não alfanuméricos viram
  espaço), com aliases resolvidos por regex, como no R.
- A coluna `TIPO DE REGISTRO` é removida.
- Colunas identificadoras (`NUMERO|CODIGO|CARTAO|CERTIFICADO|SUBFATURA|CPF|MATRICULA|INSCRICAO`)
  são lidas como texto. Zeros ausentes na origem não são reconstruídos.
- Em toda coluna sem `DATA` no nome, número e data `t="d"` ficam com o texto
  exatamente como gravado no arquivo (`<v>`), como no readxl: "001" continua
  "001", "1.2E3", 12345678901234567, "0.3333333333333333", "-0" e o serial de
  uma célula com formato de data (inclusive pasta 1904) não são reformatados.
  Lógico vira "TRUE"/"FALSE" e erro do Excel vira vazio.
- Datas: aceitar Date nativo, serial Excel (incluindo o bug do dia 60 de 1900),
  a lista de formatos de texto do R e ano curto (00–29 vira 20xx, 30–99 vira
  19xx). Data inválida fica vazia e é registrada no log.
- Coluna `*_Y2K`: substitui a data principal quando dia e mês coincidem.
  Quando não coincidem, a principal é mantida e a divergência é registrada.
- O titular é completado com o cartão e o nome da mãe da aba suplementar, pelo
  certificado (a chave é aparada apenas para comparação).
- O dependente herda `NOME DO SEGURADO`, `PLANO` e `NUMERO DA SUBFATURA` do
  titular, com a subfatura do próprio dependente como fallback.
- Dependente sem titular correspondente vai para Pendências com
  `MOTIVO = "CERTIFICADO SEM TITULAR CORRESPONDENTE"`.
- `CÓDIGO UNICO` = certificado + código do dependente.
- Toda junção confere a contagem de linhas. Se ela mudar, o arquivo é
  rejeitado.
- Arquivo com problema estrutural é rejeitado inteiro (nunca entra pela
  metade), e a mensagem vai para o Log.

### Saída
Um Excel `PC_CONSOLIDADO_<mês>_<ano>.xlsx` com 3 abas formatadas como no R
(tabela `TableStyleMedium2`, cabeçalho azul `#4472C4`, larguras manuais, datas
`dd/mm/yyyy`, primeira linha congelada):
- **Base Tratada**: ordem inicial de colunas igual a `CFG$ordem_inicial`,
  ordenada por REFERENCIA, SubF e NUMERO DO CERTIFICADO.
- **Pendencias**
- **Log**: uma linha por arquivo, com as mesmas colunas do R.

Além disso, um resumo é impresso no console.

## Decisões de implementação (escolhas do Claude, sujeitas a revisão)

- Python 3.11 ou superior, com `pandas` e `openpyxl`.
- Interface de linha de comando:
  `python tratamento.py --entrada <pasta> --empresas <csv> --saida <xlsx>`.
  Sem argumentos, o comportamento é igual ao do R (pasta e nome pelo mês e ano
  atuais).
- A lógica fica em funções separadas da execução, para que a fase 2 importe o
  mesmo código sem alterá-lo.
- Arquivos: `tratamento.py`, `tests/` e `requirements.txt` na raiz do projeto.

## Diferenças intencionais em relação ao R

- `COD_EMPRESAS.csv` salvo em ANSI (cp1252) é lido sem corromper acentos.
- Largura das colunas com `DATA` no nome: o R fixa 13. O Python usa a
  largura necessária para o cabeçalho caber em até 2 linhas, com mínimo de 13
  (pedido do usuário em 2026-10-22).

Revisão por 5 agentes (rodada 2, 2026-10-23). Desvios que protegem dados:
- **Linhas finais de DEPENDENTES**: o R sempre remove as 2 últimas linhas. O
  Python só remove se a aba tiver a coluna `TIPO DE REGISTRO` e a última linha
  tiver o valor `4` (o resumo do formato real). Com a coluna e outro valor, nada
  é removido e o Log avisa ("ultima linha sem TIPO DE REGISTRO 4"). Sem a
  coluna, remove como o R.
- **Avisos que o R não grava** (sempre no fim de `AVISOS_ESTRUTURA`): linhas
  vazias no meio de DEPENDENTES (os registros continuam, como no R) e mesma
  `REFERENCIA` ou `CODIGO_EMPRESA` em mais de um arquivo (aviso, não erro). A
  paridade ignora só esses textos (`tests/paridade/comparar.py`).
- **Saída igual a uma entrada ou ao CSV de empresas**: recusada antes de
  processar (o R sobrescreveria a planilha).
- **Mais de 1.048.575 linhas numa aba**: recusado antes de gravar.
- **Erros de abertura no Log** em português: "arquivo corrompido ou nao e
  .xlsx" e "arquivo em uso ou sem permissao" (o R mostra a mensagem do readxl).

Diferenças conhecidas e aceitas:
- Data com ano `0000` (ex.: "01/01/0000" ou o serial-texto "11060"): o R cria
  o ano 0; o Python deixa vazio e registra no Log (o ano 0 não existe em
  `datetime.date`).
- CR literal dentro de texto (gerado fora do Excel): o readxl mantém `\r\n`; o
  parser do openpyxl entrega `\n`. As chaves são aparadas, então as junções
  não mudam.

Alinhado ao R nesta rodada (antes divergia): célula com fórmula sem valor em
cache ou `<v>` vazio conta na faixa lida (varredura própria do XML da aba);
célula só com espaços e quebras de linha vira vazio; datas de texto aceitam só
dígitos e espaços ASCII; data em coluna texto de pasta 1904 volta ao serial
1904; cabeçalho do TITULAR na linha max(skip, 1ª linha com conteúdo).
Depois (2026-10-26, item P11): número em coluna texto usa o texto bruto de
`<v>` (ver Regras de tratamento); resolve também o antigo item "números
>= 1e16 em texto".

Pendências registradas (fora desta rodada):
- N3: "00/00/0000" conta como data inválida (igual ao R).
- N5: CSV de empresas com "," em vez de ";" dá a mensagem de colunas ausentes.
- N6: a guarda de tipos incompatíveis entre arquivos parece inalcançável (o
  tipo vem do nome da coluna); mantida como defesa.
- N7: desempenho (arquivo de 24 mil linhas no formato real: 7 a 30 s conforme
  a máquina; a varredura do XML acrescenta ~20%, e o texto bruto de `<v>` não
  muda o tempo no formato real, que grava tudo como texto; numa aba só de
  números, +15%) e cancelamento só entre arquivos: fase 2.

## Testes

1. **Testes automáticos com planilhas sintéticas** (dados falsos, gerados pelos
   próprios testes), cobrindo cada regra: datas inválidas, serial Excel, ano
   curto, Y2K divergente, certificado duplicado, dependente sem titular,
   conflito na aba suplementar, aba faltando, código de empresa inexistente.
2. **Paridade com o R**: rodar `tratamento.R` e `tratamento.py` nos mesmos
   arquivos e comparar as saídas célula a célula. Requer o R instalado e
   planilhas reais colocadas pelo usuário em uma subpasta do projeto (nunca em
   `ignore/`). Tudo roda localmente; nenhum dado sai do PC.

**Critério de pronto:** testes sintéticos passando e paridade confirmada com
arquivos reais.

## Fora do escopo

Interface gráfica, empacotamento em `.exe` e edição da relação de empresas.
Esses itens ficam para a fase 2.
