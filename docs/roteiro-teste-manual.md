# Roteiro de teste manual do ConsolidadorPC (antes de entregar o .exe)

Use só os dados fictícios de `exemplos/`. Simule a "rede" com uma pasta local,
por exemplo `C:\teste_rede\`. Marque cada passo como OK ou anote o problema.

1. **Primeiro uso com o CSV de exemplo.** Apague `%APPDATA%\ConsolidadorPC`,
   abra o `.exe`, responda **Não** e crie `C:\teste_rede\empresas.csv` a partir
   de `exemplos\COD_EMPRESAS_EXEMPLO.csv`. Esperado: aba Empresas com 3 linhas e
   rodapé "origem: rede".
2. **Adicionar a pasta de exemplos.** Adicionar pasta > `exemplos\entrada`.
   Esperado: 3 arquivos com empresa e competência 10/2026, nenhum em vermelho;
   adicionar a mesma pasta de novo não duplica.
3. **Processar.** Esperado: barra e "Processando i/3: ..."; resumo "3 OK, 0 com
   erro", 1 pendência e datas inválidas do PC90003; origem "rede".
4. **Cancelar no meio.** Gere um arquivo grande fictício
   (`.venv\Scripts\python -m scripts.medir_desempenho C:\temp\grande`), adicione
   `C:\temp\grande\real\PC90001 10.2026.xlsx` (cadastre 90001 se preciso),
   Processar e Cancelar após uns segundos. Esperado: para em até ~5 s com
   "Processamento cancelado."; a janela nunca fica "Não está respondendo".
5. **Salvar e abrir o arquivo.** Processe os exemplos, Salvar como (nome
   sugerido `PC_CONSOLIDADO_10_2026.xlsx`; confirme a pergunta da pendência),
   depois **Abrir arquivo** e **Abrir pasta**. Esperado: abas Base Tratada,
   Pendencias e Log, igual ao CLI.
6. **Incluir empresa.** Adicione um arquivo `PC90009 10.2026.xlsx` (cópia de um
   exemplo): aparece em vermelho. Na aba Empresas, inclua `DELTA FICTICIA` /
   `90009`. Esperado: linha deixa de ficar vermelha, botão "Processar de novo"
   aparece e `C:\teste_rede\historico\empresas_AAAAMMDD.csv` é criado.
7. **Simular duas instâncias editando.** Abra o `.exe` duas vezes. Na janela A
   edite a ALFA; na B (sem recarregar) edite a BETA: as duas mudanças ficam no
   arquivo. Agora na B edite a ALFA. Esperado: aviso "outra pessoa alterou esta
   empresa agora; confira e confirme de novo" e tabela recarregada.
8. **Desligar a rede (renomear a pasta).** Feche o app, renomeie
   `C:\teste_rede` para `C:\teste_rede_off` e abra de novo. Esperado: faixa
   amarela "Usando cópia local de <data>" nas duas abas, edição bloqueada,
   processamento funcionando e resumo com "cópia local de <data>". Volte o nome.
   Repita apontando a relação (Configurações) para um servidor que não
   responde, ex.: `\\10.255.255.1\compartilhada\empresas.csv`: a janela abre
   na hora com "Lendo relação de empresas..." e nunca fica "Não está
   respondendo"; depois de ~40 s cai na cópia local.
9. **Salvar com o destino aberto no Excel.** Abra no Excel o consolidado do
   passo 5 e salve por cima dele. Esperado: mensagem clara para fechar o
   arquivo; depois de fechar, Salvar como funciona sem reprocessar.
10. **Fechar durante o processamento.** Com o arquivo grande do passo 4,
    Processar e fechar a janela. Esperado: pergunta se deve cancelar; "Não"
    mantém a janela, "Sim" cancela e fecha. Com resultado não salvo, fechar
    também pergunta.
