#!/usr/bin/env bash
# Prepara o clasp no runner a partir dos secrets CLASPRC_JSON e SCRIPT_ID.
set -euo pipefail
: "${CLASPRC_JSON:?Secret CLASPRC_JSON não configurado}"
: "${SCRIPT_ID:?Secret SCRIPT_ID não configurado}"

printf '%s' "$CLASPRC_JSON" > "$HOME/.clasprc.json"

# clasp 3.x usa {"tokens":{...}}; clasp 2.x usa {"token":{...}}
if jq -e '.tokens' "$HOME/.clasprc.json" >/dev/null 2>&1; then
  npm install -g @google/clasp@3 >/dev/null
  echo "CLASP_MAJOR=3" >> "$GITHUB_ENV"
else
  npm install -g @google/clasp@2.4.2 >/dev/null
  echo "CLASP_MAJOR=2" >> "$GITHUB_ENV"
fi

printf '{"scriptId":"%s","rootDir":"src"}\n' "$SCRIPT_ID" > .clasp.json
mkdir -p src
clasp --version
