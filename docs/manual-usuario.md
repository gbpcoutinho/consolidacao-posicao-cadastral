# Consolidador de Posição Cadastral — manual rápido

## Instalar

1. Copie `ConsolidadorPC-1.0.0.zip` da pasta da rede para o seu PC e extraia.
2. Abra `ConsolidadorPC\ConsolidadorPC.exe`. Não precisa instalar Python.
   Se o Windows avisar que o app é desconhecido, clique em "Mais informações" e
   "Executar assim mesmo".

## Primeiro uso: relação de empresas

Na primeira vez o app pergunta onde está a relação de empresas (o mesmo
formato do `COD_EMPRESAS.csv`: `EMPRESA;COD_SAUDE;COD_DENTAL`).

- **Sim**: escolha o arquivo que já está na rede (ex.: `\\servidor\pasta\empresas.csv`).
- **Não**: crie o arquivo da rede a partir de um `COD_EMPRESAS.csv` que você tem.

Dá para trocar depois em **Configurações > Arquivo da relação de empresas**.

## Consolidar

1. **Adicionar arquivos** ou **Adicionar pasta** (temporários `~$` e
   consolidados anteriores `PC_CONSOLIDADO_*` ficam de fora).
   Empresa não cadastrada aparece em **vermelho**.
2. **Processar**. Se houver competências diferentes, arquivo sem competência
   (`MM.AAAA` no nome) ou empresa não cadastrada, o app pergunta antes.
   **Cancelar** para em poucos segundos.
3. Confira o **Resumo**: "X OK, Y com erro", registros, pendências, datas
   inválidas e a lista de erros e avisos (clique duplo mostra o motivo inteiro;
   empresa ou referência repetida aparece em destaque).
4. **Salvar como**: o nome sugerido usa a competência
   (ex.: `PC_CONSOLIDADO_10_2026.xlsx`). Se o destino estiver aberto no Excel,
   feche-o e clique em Salvar como de novo; não precisa reprocessar.
   Depois use **Abrir arquivo** ou **Abrir pasta**.

## Empresas

Busque, selecione e use **Incluir**, **Salvar edição** ou **Excluir**. O
COD_SAUDE é obrigatório e não pode repetir. Cada confirmação grava direto no
arquivo da rede; se outra pessoa alterou a mesma empresa agora, o app avisa e
recarrega a tabela. Se o arquivo estiver aberto no Excel por alguém, tente de
novo depois. Em `historico\` na pasta da rede fica, para cada dia com
alteração, a relação como estava antes da primeira alteração daquele dia
(`empresas_AAAAMMDD.csv`); nada é apagado automaticamente.

## Sem rede

Uma **faixa amarela** mostra que o app está usando a cópia local da relação
(com a data). Dá para processar, mas a edição de empresas fica bloqueada.

## Problemas

Mensagens aparecem em português. Detalhes técnicos de erros inesperados ficam
em `%APPDATA%\ConsolidadorPC\logs\app.log` (envie esse arquivo a quem dá
suporte). Nenhum dado de cliente sai do seu PC; só a relação de empresas fica
na rede.
