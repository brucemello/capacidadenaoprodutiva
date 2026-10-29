/*******************************************************************************
 * LEITOR DO % NÃO PRODUTIVO  (copiar para o projeto RMR Executivo)
 *
 * Lê a aba RESUMO_NAO_PRODUTIVO da planilha "Controle Mensal de Capacidade
 * Não Produtiva", gerada todo mês pela rotina automática.
 * As colunas são localizadas pelo nome do cabeçalho, então a ordem pode mudar.
 *
 * Uso:
 *   const r = obterNaoProdutivo('Novas Apurações', '2026-09');
 *   r.percNaoProdutivo  // 0.1603  (fração: 16,03%)
 *   r.atualizadoEm      // '03/11/2026 07:12'
 *
 *   obterNaoProdutivo('Novas Apurações')          // último mês disponível
 *   listarNaoProdutivo({ mes: '2026-09' })        // todos os times do mês
 *   listarNaoProdutivo({ time: 'Novas Apurações' }) // histórico do time
 *******************************************************************************/

const CNP_PLANILHA_ID = '1DZNqcA5YKSSBBN6NXfrIy-pnG75xPY3XroiT5vMj-JA';
const CNP_ABA = 'RESUMO_NAO_PRODUTIVO';

function listarNaoProdutivo(filtro) {
  filtro = filtro || {};
  const sh = SpreadsheetApp.openById(CNP_PLANILHA_ID).getSheetByName(CNP_ABA);
  if (!sh || sh.getLastRow() < 2) return [];

  const n = sh.getLastRow(), c = sh.getLastColumn();
  const valores = sh.getRange(1, 1, n, c).getValues();
  const exibidos = sh.getRange(1, 1, n, c).getDisplayValues();
  const cab = valores[0].map(h => String(h).trim().toLowerCase());
  const col = trecho => cab.findIndex(h => h.indexOf(trecho) >= 0);
  const iMes = col('mês'), iTime = col('time'), iEsp = col('esperadas'), iAp = col('apontadas'),
        iNp = col('não produtivas'), iPerc = col('% não produtivo'), iAus = col('ausência'), iAt = col('atualizado');

  const num = v => typeof v === 'number' ? v : Number(String(v).replace(/\./g, '').replace(',', '.').replace('%', '')) || 0;
  const lista = [];
  for (let r = 1; r < n; r++) {
    const mes = String(exibidos[r][iMes]).substring(0, 7);       // 'yyyy-MM'
    const time = String(valores[r][iTime]).trim();
    if (!mes || !time) continue;
    if (filtro.mes && mes !== filtro.mes) continue;
    if (filtro.time && time !== filtro.time) continue;
    let perc = valores[r][iPerc];
    perc = typeof perc === 'number' ? perc : num(perc) / 100;      // sempre fração
    lista.push({
      mes: mes,
      time: time,
      horasEsperadas: num(valores[r][iEsp]),
      horasApontadas: num(valores[r][iAp]),
      horasNaoProdutivas: num(valores[r][iNp]),
      percNaoProdutivo: perc,
      horasAusencia: iAus >= 0 ? num(valores[r][iAus]) : 0,
      atualizadoEm: iAt >= 0 ? exibidos[r][iAt] : ''
    });
  }
  return lista.sort((a, b) => b.mes.localeCompare(a.mes) || a.time.localeCompare(b.time));
}

/** Um time num mês ('yyyy-MM'); sem mês = último disponível. Retorna null se não houver. */
function obterNaoProdutivo(time, mes) {
  const lista = listarNaoProdutivo({ time: time, mes: mes });
  return lista.length ? lista[0] : null;
}

/** Teste rápido: rode no editor e veja o log. */
function testarLeitorNaoProdutivo() {
  Logger.log(JSON.stringify(obterNaoProdutivo('Novas Apurações'), null, 2));
}
