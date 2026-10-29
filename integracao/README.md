# Integração com o RMR Executivo

O RMR Executivo lê o **% não produtivo** desta apuração na aba `RESUMO_NAO_PRODUTIVO`
da planilha *Controle Mensal de Capacidade Não Produtiva*.

## Contrato da aba RESUMO_NAO_PRODUTIVO

Uma linha por **mês × time**, mês mais recente primeiro.

| Coluna | Conteúdo |
|---|---|
| Mês referência | `yyyy-MM` |
| Time | Nome do time (igual à aba TIME) |
| Total Horas Esperadas | Capacidade já descontando feriados e ausências justificadas |
| Total Horas Apontadas | Horas apontadas no Tempo |
| Total Horas Não Produtivas | Esperadas − apontadas (por colaborador, mínimo 0) |
| % Não Produtivo do Time | Não produtivas ÷ esperadas |
| Total Horas Ausência (abatidas) | Férias, folgas e ausências da agenda que justificaram horas |
| Atualizado em | Quando aquele mês/time foi calculado pela última vez |

Os nomes dos cabeçalhos não mudam; o leitor localiza as colunas pelo nome.

## Como usar no RMR Executivo

1. Copie `LeitorNaoProdutivo.gs` para o projeto Apps Script do RMR Executivo.
2. A conta que executa o RMR precisa ter acesso de leitura à planilha.
3. Chame `obterNaoProdutivo('Novas Apurações', '2026-09')`.

## Quando o dado fica pronto

A rotina mensal roda **todo dia 3, às 7h**. Agende a leitura do RMR para depois disso
(por exemplo, dia 3 às 9h) ou confira `atualizadoEm`.
