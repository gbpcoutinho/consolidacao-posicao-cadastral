# Fase 2 — Aplicativo desktop (esboço)

Data: 2026-10-22
Status: esboço. Será detalhado depois que a fase 1 estiver concluída e
validada.

## Ideia

Um executável para Windows em que o usuário:
1. carrega as planilhas Excel de posição cadastral que recebe;
2. processa os arquivos com a mesma lógica do `tratamento.py` (fase 1);
3. recebe um único arquivo consolidado e escolhe onde salvá-lo ("Salvar
   como").

## Já definido

- **Linguagem:** Python, pela facilidade de programação (decisão do usuário).
- **Lógica:** reaproveitar as funções do `tratamento.py` sem alterá-las.
- **Relação de empresas (`COD_EMPRESAS.csv`):** fica dentro do app, como uma
  aba ou opção onde o usuário cadastra e edita os pares EMPRESA / COD_SAUDE, e
  o conteúdo fica salvo entre usos (ideia do usuário).
  - Nota técnica: um `.exe` não consegue gravar dentro de si mesmo. Proposta:
    salvar a relação em `%APPDATA%\<NomeDoApp>\empresas.json`. Para quem usa,
    o efeito é o mesmo de "salvo no app".

## Propostas iniciais (a confirmar)

- **Empacotamento:** PyInstaller, gerando um `.exe` que roda sem Python
  instalado.
- **Interface:** `tkinter`, que já vem com o Python, ou `PySide6`, para um
  visual mais polido.
- **Fluxo de tela:** selecionar arquivos (diálogo do Windows, vários de uma
  vez), botão Processar com progresso por arquivo, resumo (sucessos, erros,
  pendências), e depois "Salvar como".
- **Erros por arquivo:** mostrar na tela os arquivos rejeitados e o motivo,
  além da aba Log.

## Em aberto

- Quem vai usar: só o autor ou outras pessoas da empresa, nos computadores
  delas? Isso afeta distribuição, atualização e onde a relação de empresas
  fica salva (por usuário ou compartilhada).
- Importar e exportar a relação de empresas a partir de um CSV existente?
- Nome do aplicativo.
- Visual: simples e funcional, ou com identidade da empresa?
- Antivírus e políticas de TI da empresa (executáveis do PyInstaller às vezes
  são bloqueados).
- Restrição de dados: o app processa tudo localmente e não envia dados para
  fora do PC.
