/*******************************************************************************
 * INTEGRAÇÃO DIRETA COM A API DO TEMPO (JIRA)
 * Substitui a importação manual do CSV: busca os apontamentos do mês fechado
 * anterior direto pela URL do Jira usando o Team ID de cada time.
 *
 * Configuração:
 *  - URL do Jira e Token (PAT): Painel lateral > Configurações da API
 *  - Team ID de cada time: aba TIME, coluna "Team ID (Tempo)"
 *******************************************************************************/

const CHAVES_TEMPO = {
  JIRA_BASE_URL: 'CONFIG_JIRA_BASE_URL',
  JIRA_TOKEN: 'CONFIG_JIRA_TOKEN'
};

const TEMPO = {
  URL_PADRAO: 'https://jiraproducao.totvs.com.br',
  SHEET_TIMES: 'TIME',
  HEADER_TEAM_ID: 'Team ID (Tempo)',
  DIA_ROTINA_MENSAL: 3,   // roda no dia 3 de cada mês (dá tempo de apontamentos atrasados)
  HORA_ROTINA_MENSAL: 7,
  FUNCAO_ROTINA: 'rotinaMensalMesFechado'
};

/* =============================================================================
 * CONFIGURAÇÕES (chamadas pelo painel)
 * ===========================================================================*/
function obterConfiguracoesTempo() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(CHAVES_TEMPO.JIRA_TOKEN) || '';
  return {
    jiraBaseUrl: props.getProperty(CHAVES_TEMPO.JIRA_BASE_URL) || TEMPO.URL_PADRAO,
    tokenConfigurado: !!token,
    rotinaAtiva: ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === TEMPO.FUNCAO_ROTINA),
    rotinaAusenciasAtiva: ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'rotinaDiariaAusencias'),
    mesFechadoPadrao: Utilities.formatDate(primeiroDiaMesAnterior_(new Date()), CONFIG.TIMEZONE, 'yyyy-MM'),
    times: obterTimesComTeamId_().map(t => ({ nome: t.nome, teamId: t.teamId }))
  };
}

function salvarConfiguracoesTempo(cfg) {
  try {
    const props = PropertiesService.getScriptProperties();
    const url = String(cfg.jiraBaseUrl || '').trim().replace(/\/+$/, '');
    if (!url) throw new Error('Informe a URL do Jira.');
    props.setProperty(CHAVES_TEMPO.JIRA_BASE_URL, url);

    const token = String(cfg.jiraToken || '').trim();
    if (token) props.setProperty(CHAVES_TEMPO.JIRA_TOKEN, token);
    if (!props.getProperty(CHAVES_TEMPO.JIRA_TOKEN)) throw new Error('Informe o Token da API (PAT do Jira).');

    garantirColunaTeamId_();
    registrarLog_('Configuração API Tempo', 'SUCESSO', `URL: ${url}`);
    return { sucesso: true, mensagem: 'Configurações salvas!' };
  } catch (e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function lerConfigTempo_() {
  const props = PropertiesService.getScriptProperties();
  const cfg = {
    jiraBaseUrl: (props.getProperty(CHAVES_TEMPO.JIRA_BASE_URL) || TEMPO.URL_PADRAO).replace(/\/+$/, ''),
    jiraToken: props.getProperty(CHAVES_TEMPO.JIRA_TOKEN) || ''
  };
  if (!cfg.jiraToken) throw new Error('Token da API não configurado. Abra o painel > Configurações da API.');
  return cfg;
}

/* =============================================================================
 * ABA TIME: nome do time + Team ID do Tempo
 * ===========================================================================*/
function garantirColunaTeamId_() {
  const ss = obterPlanilha_();
  const sh = getOrCreateSheet_(ss, TEMPO.SHEET_TIMES);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 2).setValues([['Time', TEMPO.HEADER_TEAM_ID]]);
    sh.getRange(1, 1, 1, 2).setBackground('#0F5B78').setFontColor('#FFFFFF').setFontWeight('bold');
    return 2;
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = header.findIndex(h => /team\s*id/i.test(String(h)));
  if (idx >= 0) return idx + 1;

  const novaCol = lastCol + 1;
  sh.getRange(1, novaCol).setValue(TEMPO.HEADER_TEAM_ID)
    .setBackground('#0F5B78').setFontColor('#FFFFFF').setFontWeight('bold');
  return novaCol;
}

function obterTimesComTeamId_() {
  const ss = obterPlanilha_();
  const sh = ss.getSheetByName(TEMPO.SHEET_TIMES);
  if (!sh || sh.getLastRow() < 2) return [];
  const colId = garantirColunaTeamId_();
  const n = sh.getLastRow() - 1;
  const nomes = sh.getRange(2, 1, n, 1).getValues();
  const ids = sh.getRange(2, colId, n, 1).getValues();

  const vistos = {};
  const lista = [];
  for (let i = 0; i < n; i++) {
    const nome = String(nomes[i][0] || '').trim();
    const teamId = idsDoTime_(ids[i][0]).join(', ');
    if (!nome || vistos[nome]) continue;
    vistos[nome] = true;
    lista.push({ nome: nome, teamId: teamId });
  }
  return lista;
}

/* =============================================================================
 * ENTRADA PELO PAINEL: busca o mês informado (padrão = mês fechado anterior)
 * ===========================================================================*/
function buscarTimesheetTempoPeloPainel(timeSelecionado, mesRef) {
  try {
    if (timeSelecionado === '-- Todos os Times (Geral) --') timeSelecionado = '';
    const res = buscarEProcessarMesFechado_(timeSelecionado, mesRef);
    return { sucesso: true, mensagem: res };
  } catch (e) {
    registrarLog_('Busca API Tempo', 'ERRO', e.toString());
    return { sucesso: false, mensagem: e.message || e.toString() };
  }
}

/* =============================================================================
 * ROTINA AUTOMÁTICA MENSAL (gatilho)
 * ===========================================================================*/
function rotinaMensalMesFechado() {
  try {
    const res = buscarEProcessarMesFechado_('', null);
    Logger.log(res);
  } catch (e) {
    registrarLog_('Rotina Mensal Automática', 'ERRO', e.toString());
    throw e; // o Apps Script avisa o dono por e-mail quando o gatilho falha
  }
}

function ativarRotinaMensal() {
  try {
    removerGatilhosRotina_();
    ScriptApp.newTrigger(TEMPO.FUNCAO_ROTINA)
      .timeBased()
      .onMonthDay(TEMPO.DIA_ROTINA_MENSAL)
      .atHour(TEMPO.HORA_ROTINA_MENSAL)
      .inTimezone(CONFIG.TIMEZONE)
      .create();
    ScriptApp.newTrigger('rotinaDiariaAusencias')
      .timeBased()
      .everyDays(1)
      .atHour(6)
      .inTimezone(CONFIG.TIMEZONE)
      .create();
    const msg = `Rotina ativada: todo dia ${TEMPO.DIA_ROTINA_MENSAL} às ${TEMPO.HORA_ROTINA_MENSAL}h busca o mês fechado anterior; ausências sincronizadas diariamente às 6h.`;
    registrarLog_('Rotina Mensal', 'SUCESSO', msg);
    return { sucesso: true, mensagem: msg };
  } catch (e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function desativarRotinaMensal() {
  removerGatilhosRotina_();
  registrarLog_('Rotina Mensal', 'SUCESSO', 'Rotina mensal desativada.');
  return { sucesso: true, mensagem: 'Rotina mensal desativada.' };
}

function removerGatilhosRotina_() {
  ScriptApp.getProjectTriggers()
    .filter(t => [TEMPO.FUNCAO_ROTINA, 'rotinaDiariaAusencias'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/* =============================================================================
 * NÚCLEO: busca no Tempo -> grava TIMESHEET_IMPORTAR -> processa o mês
 * ===========================================================================*/
function buscarEProcessarMesFechado_(timeSelecionado, mesRef) {
  const cfg = lerConfigTempo_();
  const periodo = calcularPeriodoMes_(mesRef);

  let times = obterTimesComTeamId_();
  if (timeSelecionado) times = times.filter(t => t.nome === timeSelecionado);

  const semId = times.filter(t => !t.teamId).map(t => t.nome);
  times = times.filter(t => t.teamId);

  if (times.length === 0) {
    throw new Error(timeSelecionado
      ? `O time "${timeSelecionado}" não tem Team ID preenchido na aba TIME (coluna "${TEMPO.HEADER_TEAM_ID}").`
      : `Nenhum time com Team ID na aba TIME. Preencha a coluna "${TEMPO.HEADER_TEAM_ID}".`);
  }

  const cacheUsuarios = {};
  const linhas = [];
  const membrosPorTime = [];
  const avisosPermissao = [];
  const resumoTimes = [];

  times.forEach(time => {
    // Um time pode ter vários Team IDs no Tempo (ex.: TSS = 1294, 895)
    const idsTime = idsDoTime_(time.teamId);
    const worklogs = deduplicarWorklogs_([].concat.apply([], idsTime.map(id => buscarWorklogsTempo_(cfg, id, periodo.deIso, periodo.ateIso))));
    const membros = [];
    const vistosMembros = {};
    idsTime.forEach(id => buscarMembrosTime_(cfg, id, periodo).forEach(m => {
      if (vistosMembros[m.name]) return;
      vistosMembros[m.name] = true;
      membros.push(m);
    }));

    membros.forEach(m => { if (m.key) cacheUsuarios[m.key.toLowerCase()] = m; });

    // Resolve worker (chave JIRAUSERxxx) -> username / nome
    const chavesSemNome = [];
    worklogs.forEach(wl => {
      const id = identificarAutor_(wl);
      if (!id.username && id.key && !cacheUsuarios[id.key.toLowerCase()]) chavesSemNome.push(id.key);
    });
    resolverUsuariosJira_(cfg, [...new Set(chavesSemNome)], cacheUsuarios);

    let horasTime = 0;
    worklogs.forEach(wl => {
      const id = identificarAutor_(wl);
      const cache = id.key ? cacheUsuarios[id.key.toLowerCase()] : null;
      const username = String(id.username || (cache && cache.name) || id.key || '').trim().toLowerCase();
      const nome = id.nome || (cache && cache.displayName) || username;
      const dataIso = String(wl.started || wl.dateStarted || wl.startDate || '').substring(0, 10);
      const horas = (Number(wl.timeSpentSeconds) || 0) / 3600;
      if (!username || !/^\d{4}-\d{2}-\d{2}$/.test(dataIso)) return;

      linhas.push([nome, username, dataIso, Number(horas.toFixed(4)), time.nome]);
      horasTime += horas;
    });

    if (!worklogs.length && membros.length) {
      avisosPermissao.push(time.nome);
      registrarLog_('Busca API Tempo', 'ERRO', `${time.nome} (Team ID ${time.teamId}): o Tempo devolveu 0 apontamentos para ${membros.length} membro(s). ` +
        'Provável falta de permissão do dono do token para ver os registros da equipe (Tempo → Equipes → Permissões) ou dos projetos do Jira.');
    }
    membrosPorTime.push({ time: time.nome, membros: membros });
    resumoTimes.push(`${time.nome}: ${worklogs.length} apontamentos (${horasTime.toFixed(1)}h)`);
  });

  // Grava a base bruta (substitui o antigo CSV)
  const ss = obterPlanilha_();
  const shImportar = getOrCreateSheet_(ss, CONFIG.SHEET_IMPORTAR);
  shImportar.clear();
  shImportar.clearFormats();
  const saida = [['Nome completo', 'Nome de Usuário', 'data de Trabalho', 'Horas', 'Time']].concat(linhas);
  // Formatos explícitos: data como texto ISO e horas como número (evita herdar formato de data)
  const maxLinhas = Math.max(shImportar.getMaxRows(), saida.length);
  shImportar.getRange(1, 3, maxLinhas, 1).setNumberFormat('@');
  shImportar.getRange(1, 4, maxLinhas, 1).setNumberFormat('0.00');
  shImportar.getRange(1, 1, saida.length, saida[0].length).setValues(saida);
  shImportar.getRange(1, 1, 1, saida[0].length).setBackground('#0F5B78').setFontColor('#FFFFFF').setFontWeight('bold');
  shImportar.setFrozenRows(1);

  // Membros do time sem nenhum apontamento no mês também entram (com 0h)
  registrarMembrosNaConfig_(membrosPorTime, new Set(linhas.map(l => l[1])));

  // Atualiza ausências (agenda + planilha de férias) dos times processados antes de calcular
  let msgAus = '';
  try {
    const fontes = obterFontesAusenciaPorTime_().filter(f => (f.agenda || f.planilhaFerias || f.timesheet) && times.some(t => t.nome === f.time));
    if (fontes.length) {
      sincronizarAusencias_(fontes.length === 1 ? fontes[0].time : timeSelecionado);
      msgAus = ' | Ausências sincronizadas: ' + fontes.map(f => f.time).join(', ');
    }
  } catch (e) {
    msgAus = ' | Aviso: ausências não sincronizadas (' + mensagemPermissao_(e) + ')';
    registrarLog_('Sincronização de Ausências', 'ERRO', e.toString());
  }

  processarTimesheetMensal(timeSelecionado, {
    inicio: periodo.inicio,
    fim: periodo.fim,
    timesProcessados: times.map(t => t.nome)
  });

  const mesTxt = Utilities.formatDate(periodo.inicio, CONFIG.TIMEZONE, 'MM/yyyy');
  let msg = `Mês ${mesTxt} processado via API Tempo. ${resumoTimes.join(' | ')}`;
  if (semId.length) msg += ` | Sem Team ID (ignorados): ${semId.join(', ')}`;
  msg += msgAus;
  if (avisosPermissao.length) msg += ` | ⚠️ Sem apontamentos visíveis (verifique a permissão no Tempo): ${avisosPermissao.join(', ')}`;
  registrarLog_('Busca API Tempo', 'SUCESSO', msg);
  return msg;
}

/** "1294, 895" / "1294;895" / 1294 → ['1294', '895'] */
function idsDoTime_(valor) {
  return String(valor == null ? '' : valor).replace(/\.0+(?=\D|$)/g, '').split(/[^\d]+/).filter(Boolean)
    .filter((v, i, a) => a.indexOf(v) === i);
}

/** Remove apontamentos repetidos (pessoa em mais de um Team ID do mesmo time). */
function deduplicarWorklogs_(lista) {
  const vistos = {};
  return lista.filter(wl => {
    const id = wl.tempoWorklogId || wl.id || wl.originId || wl.worklogId;
    const autor = (wl.worker && (wl.worker.key || wl.worker.name)) || (wl.author && (wl.author.key || wl.author.name || wl.author.accountId)) || wl.worker || '';
    const k = id ? 'id:' + id : ['x', autor, wl.started || wl.startDate, wl.timeSpentSeconds, (wl.issue && (wl.issue.key || wl.issue.id)) || ''].join('|');
    if (vistos[k]) return false;
    vistos[k] = true;
    return true;
  });
}

/* =============================================================================
 * CHAMADAS HTTP
 * ===========================================================================*/
function buscarWorklogsTempo_(cfg, teamId, deIso, ateIso) {
  const headers = { Authorization: `Bearer ${cfg.jiraToken}` };

  // Tempo Timesheets (Jira Server/Data Center)
  let resp = fetchComDiagnostico_(cfg, `${cfg.jiraBaseUrl}/rest/tempo-timesheets/4/worklogs/search`, {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify({ from: deIso, to: ateIso, teamId: [parseInt(teamId, 10)] }),
    muteHttpExceptions: true
  });

  // Fallback: Tempo Cloud
  if (resp.getResponseCode() === 404 || resp.getResponseCode() === 0) {
    resp = UrlFetchApp.fetch(`https://api.tempo.io/v3/worklogs/team/${teamId}?from=${deIso}&to=${ateIso}&limit=5000`, {
      method: 'get', headers: headers, muteHttpExceptions: true
    });
  }

  const code = resp.getResponseCode();
  if (code === 401 || code === 403) throw new Error(`API Tempo recusou o acesso (${code}). Verifique o Token no painel.`);
  if (code < 200 || code >= 300) {
    throw new Error(`Erro na API do Tempo para o Team ID ${teamId} (${code}): ${resp.getContentText().substring(0, 300)}`);
  }
  const json = JSON.parse(resp.getContentText());
  return Array.isArray(json) ? json : (json.results || []);
}

/** Envolve o fetch traduzindo o erro de host inacessível pelos servidores do Google. */
function fetchComDiagnostico_(cfg, url, opcoes) {
  // Falhas de rede intermitentes (Google -> Jira): tenta de novo antes de desistir
  const TENTATIVAS = 4;
  let ultimoErro;
  for (let i = 1; i <= TENTATIVAS; i++) {
    try {
      return UrlFetchApp.fetch(url, opcoes);
    } catch (e) {
      ultimoErro = e;
      if (!/endere[cç]o n[aã]o dispon[ií]vel|address unavailable|dns|timeout|tempo limite|unavailable/i.test(String(e))) throw e;
      if (i < TENTATIVAS) Utilities.sleep(2000 * i);
    }
  }
  throw new Error(`Não foi possível acessar ${cfg.jiraBaseUrl} depois de ${TENTATIVAS} tentativas (${ultimoErro.message}). ` +
    `Costuma ser instabilidade momentânea: tente de novo em alguns minutos. Se persistir, confira a URL do Jira em Configurações.`);
}

/** Membros ativos do time no período (Tempo Teams). Falha silenciosa: é só complemento. */
function buscarMembrosTime_(cfg, teamId, periodo) {
  try {
    const resp = UrlFetchApp.fetch(`${cfg.jiraBaseUrl}/rest/tempo-teams/2/team/${teamId}/member`, {
      headers: { Authorization: `Bearer ${cfg.jiraToken}` },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return [];
    const json = JSON.parse(resp.getContentText());
    const lista = Array.isArray(json) ? json : (json.results || []);

    return lista.map(item => {
      const m = item.member || item;
      const ms = item.membership || {};
      const de = ms.dateFrom ? normalizarData_(ms.dateFrom) : null;
      const ate = ms.dateTo ? normalizarData_(ms.dateTo) : null;
      const status = String(ms.status || 'active').toLowerCase();
      const noPeriodo = (!de || de <= periodo.fim) && (!ate || ate >= periodo.inicio);
      if ((m.type && m.type !== 'USER') || m.activeInJira === false || status === 'inactive' || !noPeriodo) return null;
      return {
        key: m.key || m.name || '',
        name: String(m.name || m.key || '').toLowerCase(),
        displayName: m.displayname || m.displayName || m.name || ''
      };
    }).filter(m => m && m.name);
  } catch (e) {
    Logger.log('Aviso: não foi possível ler membros do time ' + teamId + ': ' + e);
    return [];
  }
}

/** Converte chaves JIRAUSERxxx em username + nome (em paralelo). */
function resolverUsuariosJira_(cfg, chaves, cache) {
  for (let i = 0; i < chaves.length; i += 50) {
    const lote = chaves.slice(i, i + 50);
    const respostas = UrlFetchApp.fetchAll(lote.map(k => ({
      url: `${cfg.jiraBaseUrl}/rest/api/2/user?key=${encodeURIComponent(k)}`,
      headers: { Authorization: `Bearer ${cfg.jiraToken}` },
      muteHttpExceptions: true
    })));
    respostas.forEach((r, j) => {
      if (r.getResponseCode() !== 200) return;
      const u = JSON.parse(r.getContentText());
      cache[lote[j].toLowerCase()] = {
        key: lote[j],
        name: String(u.name || lote[j]).toLowerCase(),
        displayName: u.displayName || u.name || lote[j]
      };
    });
  }
}

function identificarAutor_(wl) {
  const a = wl.author || wl.worker || wl.user || {};
  if (typeof a === 'string') return { key: a, username: '', nome: wl.authorDisplayName || '' };
  return {
    key: a.key || a.accountId || a.name || '',
    username: a.name || a.username || '',
    nome: a.displayName || a.fullName || wl.authorDisplayName || ''
  };
}

/* =============================================================================
 * CONFIG_TIME: garante membros do time cadastrados
 * ===========================================================================*/
function registrarMembrosNaConfig_(membrosPorTime, usuariosComApontamento) {
  const ss = obterPlanilha_();
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_CONFIG);
  const dados = sh.getDataRange().getValues();
  const existentes = {};
  for (let i = 1; i < dados.length; i++) {
    const u = String(dados[i][1] || '').trim().toLowerCase();
    if (u) existentes[u] = true;
  }
  const novos = [];
  membrosPorTime.forEach(t => t.membros.forEach(m => {
    if (existentes[m.name]) return;
    existentes[m.name] = true;
    // Sem nenhum apontamento no mês: entra como NÃO para revisão (evita contas duplicadas/ex-membros com 100% não produtivo)
    const apontou = usuariosComApontamento && usuariosComApontamento.has(m.name);
    novos.push(apontou
      ? [m.displayName, m.name, 'SIM', t.time, 'Inserido automaticamente (membro do time no Tempo)']
      : [m.displayName, m.name, 'NÃO', t.time, 'Membro do time no Tempo sem apontamentos no mês. Marque SIM se fizer parte da capacidade.']);
  }));
  if (novos.length) sh.getRange(sh.getLastRow() + 1, 1, novos.length, novos[0].length).setValues(novos);
}

/* =============================================================================
 * DATAS
 * ===========================================================================*/
function primeiroDiaMesAnterior_(hoje) {
  return new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
}

/** mesRef 'yyyy-MM' (opcional). Sem informar = mês fechado anterior. */
function calcularPeriodoMes_(mesRef) {
  let inicio;
  if (mesRef && /^\d{4}-\d{2}$/.test(mesRef)) {
    const p = mesRef.split('-');
    inicio = new Date(Number(p[0]), Number(p[1]) - 1, 1);
  } else {
    inicio = primeiroDiaMesAnterior_(new Date());
  }
  const hoje = new Date();
  if (inicio > hoje) throw new Error('O mês informado ainda não começou.');
  const fim = new Date(inicio.getFullYear(), inicio.getMonth() + 1, 0);
  return {
    inicio: inicio,
    fim: fim,
    deIso: Utilities.formatDate(inicio, CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    ateIso: Utilities.formatDate(fim, CONFIG.TIMEZONE, 'yyyy-MM-dd')
  };
}
