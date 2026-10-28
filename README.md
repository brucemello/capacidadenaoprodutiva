# Capacidade Não Produtiva

Código do Google Apps Script com deploy automático via GitHub Actions.

## Como funciona

| Workflow | Quando roda | O que faz |
|---|---|---|
| **Deploy Apps Script** | Todo push na `main` que mexa em `src/` (ou manualmente) | `clasp push` + atualiza a implantação `DEPLOYMENT_ID` — a URL do Web App continua a mesma |
| **Importar do Apps Script** | Manual (aba Actions) | `clasp pull` do editor do Google para `src/` e faz commit |

## Fluxo de trabalho

1. Edite os arquivos em `src/` (ou peça ao Claude para implementar).
2. Faça commit/push na `main`.
3. Pronto — em ~1 min a nova versão está publicada.

> Evite editar direto no editor do Apps Script: o próximo deploy sobrescreve. Se editar, rode **Importar do Apps Script** antes.

## Secrets necessários

- `CLASPRC_JSON` — conteúdo do `~/.clasprc.json` (login do clasp)
- `SCRIPT_ID` — ID do projeto Apps Script
- `DEPLOYMENT_ID` — ID da implantação do Web App a ser atualizada
