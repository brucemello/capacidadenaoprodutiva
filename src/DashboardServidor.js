/*******************************************************************************
 * DASHBOARD EXECUTIVO - CAPACIDADE NÃO PRODUTIVA
 * Abre como modal na planilha (menu/painel) ou como Web App (doGet).
 *******************************************************************************/

const CHAVE_ID_PLANILHA = 'CONFIG_SPREADSHEET_ID';

/** Web App */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Capacidade Não Produtiva · Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Modal dentro da planilha */
function abrirDashboard() {
  obterPlanilha_(); // memoriza o ID para o Web App
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(1400)
    .setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard · Capacidade Não Produtiva');
}

function obterUrlWebApp() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/** Planilha ativa (modal) ou a guardada nas propriedades (Web App). */
function obterPlanilha_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (ss) {
    if (props.getProperty(CHAVE_ID_PLANILHA) !== ss.getId()) props.setProperty(CHAVE_ID_PLANILHA, ss.getId());
    return ss;
  }
  const id = props.getProperty(CHAVE_ID_PLANILHA);
  if (!id) throw new Error('Abra a planilha uma vez (menu Capacity > Abrir Dashboard) para vincular o Web App.');
  return SpreadsheetApp.openById(id);
}

/**
 * Dados do dashboard (sem objetos Date, que o google.script.run não serializa).
 */
function obterDadosDashboard() {
  const ss = obterPlanilha_();

  // ---------- Histórico mensal por colaborador ----------
  const historico = [];
  const shHist = ss.getSheetByName(CONFIG.SHEET_HISTORICO);
  if (shHist && shHist.getLastRow() > 1) {
    const n = shHist.getLastRow() - 1;
    const nColH = Math.min(Math.max(shHist.getLastColumn(), 9), 10);
    const valores = shHist.getRange(2, 1, n, nColH).getValues();
    const exibidos = shHist.getRange(2, 1, n, 1).getDisplayValues();
    for (let i = 0; i < n; i++) {
      const v = valores[i];
      const mes = extrairAnoMesSemTimezone_(exibidos[i][0] || v[0]);
      if (!mes || !v[1] || !v[3]) continue;
      historico.push({
        mes: mes,
        time: String(v[1]),
        nome: String(v[2] || v[3]),
        usuario: String(v[3]).toLowerCase(),
        considerar: String(v[4] || 'SIM').toUpperCase(),
        esperadas: converterNumero_(v[5]),
        apontadas: converterNumero_(v[6]),
        naoProdutivas: converterNumero_(v[7]),
        ausencia: converterNumero_(v[9])
      });
    }
  }

  // ---------- Pessoas fora da capacidade (CONFIG_TIME) ----------
  const foraCapacidade = [];
  const shCfg = ss.getSheetByName(CONFIG.SHEET_CONFIG);
  if (shCfg && shCfg.getLastRow() > 1) {
    shCfg.getRange(2, 1, shCfg.getLastRow() - 1, 4).getValues().forEach(r => {
      if (r[1] && String(r[2]).trim().toUpperCase() === 'NÃO') {
        foraCapacidade.push({ nome: String(r[0] || r[1]), usuario: String(r[1]).toLowerCase(), time: String(r[3] || 'Sem Time Mapeado') });
      }
    });
  }

  // ---------- Detalhe diário (última apuração processada) ----------
  const diario = [];
  const shTs = ss.getSheetByName(CONFIG.SHEET_TIMESHEET);
  if (shTs && shTs.getLastRow() > 1) {
    const n = shTs.getLastRow() - 1;
    const nColT = Math.min(Math.max(shTs.getLastColumn(), 12), 16);
    const valores = shTs.getRange(2, 1, n, nColT).getValues();
    const exibidos = shTs.getRange(2, 1, n, 5).getDisplayValues();
    for (let i = 0; i < n; i++) {
      const v = valores[i];
      const data = normalizarData_(v[4] instanceof Date ? v[4] : exibidos[i][4]);
      if (!data || !v[3]) continue;
      diario.push({
        data: Utilities.formatDate(data, CONFIG.TIMEZONE, 'yyyy-MM-dd'),
        time: String(v[1]),
        nome: String(v[2] || v[3]),
        usuario: String(v[3]).toLowerCase(),
        horas: converterNumero_(v[7]),
        diaUtil: String(v[8]).toUpperCase() === 'SIM',
        esperadas: converterNumero_(v[10]),
        ausencia: converterNumero_(v[14]),
        motivo: String(v[15] || '')
      });
    }
  }

  return {
    historico: historico,
    foraCapacidade: foraCapacidade,
    diario: diario,
    horasDia: CONFIG.HORAS_DIA,
    planilhaUrl: ss.getUrl(),
    geradoEm: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm')
  };
}
