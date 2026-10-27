# Fase 2 — Aplicativo desktop: desenho

Data: 2026-10-26
Substitui o esboço `docs/fase2-app.md`.

## Intenção

Hoje a consolidação das planilhas de posição cadastral exige rodar um script.
O objetivo é um aplicativo para Windows. A pessoa escolhe as planilhas que
recebeu, clica em Processar, vê o resultado e salva um único Excel
consolidado. Ela não precisa saber nada de Python ou R.

A lógica de negócio já está pronta e validada contra o R (fase 1,
`consolidacao/`). O aplicativo é uma casca em volta dela e não muda nenhuma
regra de consolidação.

## Decisões do usuário

- **Quem usa:** várias pessoas da empresa, cada uma no próprio PC.
- **Relação de empresas:** uma só para todos. Todos podem cadastrar e editar.
- **Onde fica a relação:** num arquivo na rede local
  (`\\servidor\pasta\empresas.csv`). O Google Drive fica como alternativa, sem
  mudar o app, porque é só apontar o caminho para outra pasta.
- **Linguagem:** Python, reaproveitando `consolidacao/` sem alterar regras.
- **Números com zero à esquerda:** continuam texto (já entregue no PR #3).
- **Sem perguntas básicas:** o usuário pediu que eu decida o básico e pergunte
  só o que for necessário.

## Objetivos

1. **Selecionar arquivos.** A pessoa escolhe várias planilhas `.xlsx` de uma
   vez, pelo diálogo do Windows, e pode adicionar ou remover arquivos da lista
   antes de processar.
2. **Processar sem travar a janela.** Há uma barra de progresso, com o arquivo
   atual e i/N, e um botão Cancelar que responde em poucos segundos, mesmo no
   meio de um arquivo grande.
3. **Ver o resumo.**
   - Arquivos OK e com erro, com o motivo de cada erro em português.
   - Registros, pendências, datas inválidas, divergências de datas e avisos.
4. **"Salvar como".** O nome padrão vem de `nome_saida(hoje)`. Se o destino
   estiver aberto no Excel, a pessoa vê uma mensagem clara e pode tentar de
   novo sem reprocessar.
5. **Aba Empresas.**
   - Tabela com EMPRESA, COD_SAUDE e COD_DENTAL, com incluir, editar, excluir e
     buscar.
   - A validação é a mesma do `carregar_mapa_empresas`: código obrigatório e um
     mesmo COD_SAUDE não pode ter duas empresas.
   - As alterações gravam direto no arquivo compartilhado.
6. **Nunca perder a edição de um colega.** Há duas pessoas editando ao mesmo
   tempo, nenhuma sobrescreve a outra em silêncio (detalhes em "Relação de
   empresas compartilhada").
7. **Funcionar sem rede.** O app usa a última cópia local da relação e mostra
   um aviso visível de que ela pode estar desatualizada. Sem rede, a edição
   fica bloqueada.
8. **Executável que roda sem Python instalado.**

## Não-objetivos (v1)

- Atualização automática do aplicativo, instalador MSI, assinatura digital do
  executável.
- Integração direta com a API do Google Drive.
- Login, permissões por usuário, histórico de quem alterou o quê.
- Mudar regras de consolidação ou o formato da saída.

## Antiobjetivos (seria um fracasso mesmo funcionando)

- A janela congelar ("Não está respondendo") durante o processamento.
- Uma edição de empresa de um colega ser apagada sem aviso.
- O arquivo compartilhado ficar corrompido ou pela metade.
- Sobrescrever uma planilha de entrada (o pipeline já impede; o app não pode
  contornar).
- Qualquer dado de cliente sair do PC. Só a relação de empresas vai para a
  rede.
- Traceback ou mensagem técnica em inglês na tela.

## Restrições

- Windows 10 e 11. Python 3.14, pandas `>=2.2,<3`, openpyxl 3.1.
- Paridade com o R continua obrigatória. A suíte atual (165 testes) e
  `tests/test_paridade.py` continuam verdes.
- O código do app não depende de `__file__` para achar dados, porque dentro do
  .exe ele aponta para uma pasta temporária.
- Formato do arquivo de empresas: o mesmo `COD_EMPRESAS.csv` de hoje, ou seja,
  `;`, colunas `EMPRESA;COD_SAUDE;COD_DENTAL`, UTF-8 com BOM. Continua abrindo
  no Excel e servindo para o CLI e para o R.

## Desenho

### Camadas

- `consolidacao/` é o motor. Ganha só o que o app precisa (veja "Motor").
- `app/` traz o código novo.
  - `config.py` cuida das preferências em
    `%APPDATA%\ConsolidadorPC\config.json`: caminho do arquivo compartilhado,
    última pasta de entrada e última pasta de saída.
  - `empresas_compartilhadas.py` lê e grava a relação na rede e mantém o cache
    local. Não depende de Tk e é testável sozinho.
  - `tarefas.py` roda o processamento numa thread de fundo, com fila de
    eventos (progresso, concluído, erro, cancelado) consumida pela janela via
    `after()`. Não depende de Tk.
  - `janela.py` traz as telas em tkinter/ttk, sem lógica de negócio.
- `consolidador_app.py` é o ponto de entrada do executável.

### Motor (mudanças em `consolidacao/`)

1. **Processar e gravar separados.** O "Salvar como" vem depois do resumo,
   então o pipeline ganha duas funções:
   - `processar(arquivos, mapa, progresso, cancelar) -> Consolidado`;
   - `gravar(resultado, saida, arquivos, arquivo_empresas)`, com as mesmas
     validações de saída de hoje.

   `executar` continua existindo para o CLI, chamando as duas.
2. **Cancelar dentro do arquivo.** Hoje o cancelamento só é consultado entre
   arquivos. Ele passa a ser consultado também entre as abas de cada arquivo e
   entre as etapas pesadas (leitura de cada aba, datas, junção). Meta: Cancelar
   responde em até ~5 s num arquivo de 24 mil linhas.
3. **Desempenho.**
   - Medir com perfil num arquivo de 24 mil linhas no formato real e no caso
     pesado (31 s). Otimizar os pontos quentes sem mudar resultado.
   - Gravar a saída em modo `write_only`, se a formatação atual puder ser
     mantida.
   - Meta: caso pesado em ≤ 15 s. Paridade continua com 0 diferenças.

### Relação de empresas compartilhada

- **Leitura.** O app lê o arquivo da rede ao abrir, ao entrar na aba Empresas
  e logo antes de processar. Cada leitura bem-sucedida atualiza o cache em
  `%APPDATA%\ConsolidadorPC\empresas_cache.csv` e guarda a data de modificação
  e o hash do conteúdo.
- **Rede inacessível.** O app usa o cache e mostra uma faixa amarela: "Usando
  cópia local de <data>; a relação pode estar desatualizada." Se não houver
  cache nem rede, o app avisa e não processa.
- **Gravação segura.** A gravação acontece a cada confirmação de inclusão,
  edição ou exclusão.
  1. Reler o arquivo da rede.
  2. Se o hash mudou desde a última leitura, aplicar a alteração da pessoa
     sobre a versão nova. Se a mesma linha (mesmo COD_SAUDE) foi alterada pelos
     dois, mostrar as duas versões e deixar a pessoa escolher. Nunca
     sobrescrever em silêncio.
  3. Validar a relação inteira com a mesma regra de `carregar_mapa_empresas`.
  4. Antes de gravar, copiar a versão anterior para
     `<pasta>\historico\empresas_AAAAMMDD_HHMMSS.csv`, guardando as últimas 30.
  5. Gravar num temporário na mesma pasta e trocar pelo definitivo com
     `os.replace`, para não deixar o arquivo pela metade.
  6. Se o arquivo estiver em uso (aberto no Excel por alguém), mostrar
     "arquivo em uso por outra pessoa; tente de novo" e manter a edição na
     tela.
- **Primeiro uso.** O app pede o caminho do arquivo compartilhado. Ele pode
  ser um arquivo existente ou um novo, criado a partir de um CSV que a pessoa
  importe.
- **Importar e exportar.** É possível importar de um `COD_EMPRESAS.csv`
  existente (com validação) e exportar uma cópia.

### Telas

- **Aba Consolidar.**
  - Lista de arquivos com os botões Adicionar e Remover.
  - Botões Processar e Cancelar.
  - Barra de progresso e texto "Processando 3/12: PC90001 10.2026.xlsx".
  - Painel de resumo e botão "Salvar como".
- **Aba Empresas.** Busca, tabela, Incluir, Editar, Excluir, Importar e
  Exportar. No rodapé ficam o caminho do arquivo, a hora da última leitura e o
  estado (rede ou cópia local).
- **Configurações (menu).** Trocar o caminho do arquivo compartilhado e ver a
  versão do app.

### Erros

- `ErroValidacao`, `ProcessamentoCancelado` e `MemoryError` aparecem como
  mensagem em português.
- Qualquer outra exceção vira "Erro inesperado", com detalhes técnicos
  gravados em `%APPDATA%\ConsolidadorPC\logs\app.log`.
- A janela continua utilizável depois de qualquer erro.

### Empacotamento e distribuição

- PyInstaller, modo pasta (`--onedir`), com o nome `ConsolidadorPC`.
  - O modo pasta abre mais rápido e costuma ser menos bloqueado por antivírus
    do que o arquivo único.
  - A pasta é distribuída em `.zip` pela mesma pasta de rede.
- O app mostra a versão (ex.: 1.0.0) em Configurações.

## Testes

- O motor continua com a suíte atual e paridade. Entram:
  - testes de `processar` e `gravar` separados;
  - teste de cancelamento no meio de um arquivo;
  - medição de desempenho registrada, sem teste cronometrado no CI.
- `empresas_compartilhadas.py` é testado com uma pasta temporária simulando a
  rede:
  - leitura, cache e falta de rede;
  - conflito entre duas instâncias;
  - mesma linha editada pelos dois;
  - arquivo travado;
  - histórico com só as últimas 30 versões;
  - gravação atômica.
- `tarefas.py` é testado com processamento falso: progresso, cancelar, erro e
  conclusão.
- Teste de fumaça da janela (Tk criado, sem interação manual) e um roteiro
  manual curto antes de entregar o .exe.

## Decisões minhas (o usuário pode trocar)

- **Interface em tkinter/ttk, e não PySide6.** Já vem com o Python, gera um
  executável bem menor e basta para estas telas.
- **Nome do app:** "Consolidador de Posição Cadastral", executável
  `ConsolidadorPC`.
- **PyInstaller em modo pasta, distribuído em .zip.**
- **Conflito resolvido linha a linha por COD_SAUDE, com escolha manual só
  quando a mesma linha foi alterada pelos dois.**
- **Histórico das últimas 30 versões da relação na própria pasta da rede.**
- **Preferências e cache em `%APPDATA%\ConsolidadorPC`.**
- **"Salvar como" depois do resumo.** Por isso o pipeline separa processar e
  gravar.
- **Metas de desempenho:** caso pesado em ≤ 15 s; Cancelar em ≤ ~5 s.

## Decisões complementares (respostas à checagem do construtor)

1. **Chave e mesclagem da relação.**
   - **Chave:** COD_SAUDE.
   - **Base da mesclagem:** a versão lida por último, guardada em memória
     junto com a data de modificação e o hash.
   - **Trocar o COD_SAUDE de uma linha:** vale como excluir a antiga e incluir
     a nova.
   - **Conflitos que pedem escolha manual:**
     - os dois alteraram a mesma chave de formas diferentes;
     - os dois incluíram a mesma chave com EMPRESA diferente;
     - um excluiu e o outro editou a mesma chave.
   - **Sem conflito:** alterações em chaves diferentes são aplicadas sem
     perguntar.
2. **Trava de escrita.**
   - **Arquivo de trava:** cada gravação cria `empresas.csv.lock`, com
     `O_CREAT|O_EXCL`, contendo usuário, máquina e hora. A trava dura só o
     tempo da gravação.
   - **Trava abandonada:** uma trava com mais de 60 s é considerada abandonada
     e removida.
   - **Ordem da gravação:** pegar a trava, reler, mesclar, validar, salvar o
     histórico, gravar e soltar a trava.
   - **Arquivo em uso:** `PermissionError` ao gravar significa "arquivo em uso
     (aberto no Excel?)".
3. **Validação do editor = `carregar_mapa_empresas`.**
   - **Obrigatórios:** EMPRESA e COD_SAUDE (sem espaços nas pontas).
   - **Unicidade:** um COD_SAUDE tem uma única EMPRESA; a mesma EMPRESA pode
     ter vários códigos.
   - **COD_DENTAL:** é opcional, livre e o motor não usa; o editor só preserva
     o valor.
   - **Linha sem COD_SAUDE:** o carregador ignora; o editor não deixa criar.
4. **Cancelar é cooperativo.** O processamento é interrompido entre etapas,
   sem subprocesso. A meta de ~5 s é de melhor esforço: as etapas são divididas
   o suficiente para atingi-la no arquivo de 24 mil linhas, e o resultado da
   medição fica registrado.
5. **Paridade vence velocidade.** Os 15 s são meta, não requisito; se não der,
   registra-se o tempo alcançado. O modo `write_only` só entra se mantiver
   estilos, larguras, congelamento e filtro iguais aos de hoje; senão fica o
   gravador atual.
6. **Primeiro uso.**
   - O caminho do arquivo compartilhado começa vazio e é salvo por usuário em
     `config.json`.
   - O diálogo aceita caminho de rede ou local (útil para teste).
   - Sem caminho e sem cache, a aba Consolidar fica bloqueada, com o botão
     "Configurar relação de empresas".
7. **Cache e processamento sem rede.**
   - O cache é a cópia fiel do CSV.
   - Sem rede, o processamento usa o cache, e a faixa amarela aparece nas duas
     abas.
   - O resumo na tela informa a origem da relação: "rede" ou "cópia local de
     <data>".
   - O arquivo de saída não muda.
8. **Empacotamento.**
   - A primeira tarefa do plano é um teste de fumaça do PyInstaller no ambiente
     atual (Python 3.14 com pandas 2.3.3, que já cumpre `>=2.2,<3`).
   - Se o PyInstaller não suportar 3.14, o executável é gerado num ambiente
     Python 3.13 separado, só para o build. O motor continua testado nos dois.

## Ajustes do especialista (dados + UX), aprovados — prevalecem sobre o que veio antes

Contexto: app interno, menos de 10 pessoas, não vai escalar.

**Simplificações**
- **Gravação da relação.** No lugar da mesclagem em 3 vias entra um controle
  otimista. Cada confirmação grava uma única alteração.
  - Com a trava: reler e comparar o hash.
  - Se o hash for igual, gravar.
  - Se mudou, aplicar a alteração sobre a versão nova, revalidar e gravar.
  - Só barrar se o mesmo COD_SAUDE mudou na versão nova. Nesse caso aparece o
    aviso "outra pessoa alterou esta empresa agora; confira e confirme de novo",
    e a tabela é recarregada.
- **Trava abandonada após 2 min** (não 60 s). Mensagem: "em uso por
  <usuário>@<máquina>".
- **Histórico:** um `historico\empresas_AAAAMMDD.csv` por dia, sobrescrito no
  mesmo dia, sem apagar nada automaticamente.
- **Fora do v1:**
  - Importar e Exportar viram só "criar a partir de um CSV" no primeiro uso.
  - Sai a releitura ao entrar na aba Empresas. A relação é relida só ao abrir
    o app, antes de processar e antes de gravar.
  - A meta de ≤ 15 s vira desejável.

**UX**
- **Botão "Adicionar pasta".** Pega os .xlsx da pasta, ignora `~$*` e
  `PC_CONSOLIDADO_*`, e não repete arquivos. Arrastar e soltar fica fora do
  v1.
- **Lista de arquivos com colunas Arquivo | Empresa | Competência.** É
  preenchida ao adicionar. Empresa não cadastrada aparece em vermelho antes de
  processar.
- **Depois de salvar:** botões "Abrir arquivo" e "Abrir pasta".
- **Resumo.**
  - Linha de topo do tipo "12 OK, 1 com erro".
  - Lista de erros e avisos (inclusive empresa ou referência repetida em
    destaque), com o motivo ao dar clique duplo.
  - Salvar com erro, pendência ou repetido pede confirmação.
- **Aba Empresas.**
  - Formulário que tira os espaços das pontas e valida o COD_SAUDE ao
    digitar.
  - Excluir pede confirmação mostrando o nome da empresa.
  - Depois de incluir, aparece o atalho "Processar de novo".
- **Fechar a janela.** Durante o processamento, pergunta e cancela. Com
  resultado não salvo, avisa.

**Integridade**
- **Competência.**
  - É extraída dos nomes, no padrão `MM.AAAA` (ex.: `PC90001 10.2026.xlsx`).
  - O nome padrão da saída usa a competência: `nome_saida(date(AAAA, MM, 1))`.
    Sem competência reconhecível, usa a data de hoje.
- **Mais de uma competência, ou arquivo sem competência:** aviso forte antes
  de processar, com confirmação.
- **Mudanças entre processar e salvar.**
  - Ao processar, guardar o hash e a origem da relação usada, e o tamanho e a
    data de modificação de cada entrada.
  - Ao salvar, se a relação ou alguma entrada mudou, perguntar "mudou desde o
    processamento; reprocessar?".
- **Cópia local com mais de 7 dias:** a data aparece na confirmação de
  processar, além da faixa amarela.

## Fica para quem implementar

- O layout exato dos widgets, os textos dos botões e os ícones.
- Detalhes internos de como interromper as etapas do motor para cancelar.
- Quais otimizações de desempenho usar, desde que o resultado não mude.
