/*******************************************************************************
 * AUSÊNCIAS: sincronização (Google Agenda + Planilha de Férias) e abatimento
 *
 * Fontes, configuradas por time na aba TIME:
 *   - "Agenda de Ausências (ID Google Agenda)": eventos da agenda do time
 *   - "Planilha de Férias (URL)": abas "Controle Consolidado*" e "Folgas e Emendas*"
 * Tudo vai para a aba AUSENCIAS (origem "Google Agenda" / "Planilha de Férias"),
 * junto com os lançamentos manuais (origem "Manual"), e é usado no cálculo:
 *   - ausência de dia inteiro  -> horas esperadas do dia = 0
 *   - ausência parcial (horas) -> horas esperadas do dia = 8 - horas
 *   - tipo "Informativo"       -> não abate (ex.: evento marcado como "Disponível")
 *******************************************************************************/

const ORIGEM_AGENDA = 'Google Agenda';
const ORIGEM_FERIAS = 'Planilha de Férias';
const ORIGEM_MANUAL = 'Manual';
const TIPO_INFORMATIVO = 'Informativo';

const AUS_JANELA_MESES_ATRAS = 4;   // agenda: sincroniza desde o 1º dia de 4 meses atrás
const AUS_JANELA_DIAS_FRENTE = 120; // ... até 120 dias à frente

/* =============================================================================
 * AUTORIZAÇÃO: executar pelo menu da planilha para o Google pedir as permissões
 * novas (Agenda, planilhas externas, Jira). Pelo Web App o pedido não aparece.
 * ===========================================================================*/
function autorizarAcessos() {
  const ok = [], falhas = [];
  try { CalendarApp.getDefaultCalendar(); ok.push('Google Agenda'); } catch (e) { falhas.push('Agenda: ' + e.message); }
  obterFontesAusenciaPorTime_().forEach(t => {
    if (t.agenda && t.agenda.indexOf('@') >= 0) {
      try { if (!CalendarApp.getCalendarById(t.agenda)) throw new Error('sem acesso'); ok.push('Agenda ' + t.time); }
      catch (e) { falhas.push(`Agenda ${t.time}: ${e.message}`); }
    }
    if (t.planilhaFerias) {
      try { SpreadsheetApp.openByUrl(t.planilhaFerias); ok.push('Férias ' + t.time); }
      catch (e) { falhas.push(`Férias ${t.time}: ${e.message}`); }
    }
  });
  try { UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); ok.push('Acesso externo (Jira)'); } catch (e) { falhas.push('Acesso externo: ' + e.message); }
  try { ScriptApp.getProjectTriggers(); ok.push('Gatilhos'); } catch (e) { falhas.push('Gatilhos: ' + e.message); }

  const msg = `Acessos OK: ${ok.join(', ')}.` + (falhas.length ? `\n\nFalhas: ${falhas.join(' | ')}` : '');
  registrarLog_('Autorização de acessos', falhas.length ? 'ERRO' : 'SUCESSO', msg.replace(/\n/g, ' '));
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* sem UI (gatilho/Web App) */ }
  return msg;
}

/* =============================================================================
 * SINCRONIZAÇÃO
 * ===========================================================================*/

/** Painel/Web App: sincroniza um time ('' = todos com fonte configurada). */
function sincronizarAusenciasPeloPainel(timeSelecionado) {
  try {
    if (timeSelecionado === '-- Todos os Times (Geral) --') timeSelecionado = '';
    return { sucesso: true, mensagem: sincronizarAusencias_(timeSelecionado) };
  } catch (e) {
    registrarLog_('Sincronização de Ausências', 'ERRO', e.toString());
    return { sucesso: false, mensagem: mensagemPermissao_(e) };
  }
}

/** Gatilho diário. */
function rotinaDiariaAusencias() {
  try { Logger.log(sincronizarAusencias_('')); }
  catch (e) { registrarLog_('Sincronização de Ausências', 'ERRO', e.toString()); throw e; }
}

function sincronizarAusencias_(timeSelecionado) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const times = obterFontesAusenciaPorTime_().filter(t => !timeSelecionado || t.time === timeSelecionado);
    const comFonte = times.filter(t => t.agenda || t.planilhaFerias);
    if (!comFonte.length) {
      throw new Error(timeSelecionado
        ? `O time "${timeSelecionado}" não tem agenda nem planilha de férias configurada (aba TIME).`
        : 'Nenhum time tem agenda de ausências ou planilha de férias configurada (aba TIME).');
    }

    const colaboradores = listarCadastro('colaboradores').linhas;
    const atuais = listarCadastro('ausencias').linhas;

    // Preserva o colaborador/tipo que alguém corrigiu manualmente em linhas sincronizadas
    const preservar = {};
    atuais.forEach(a => { if (a.origem !== ORIGEM_MANUAL && a.idEvento) preservar[a.idEvento + '|' + a.inicio] = a; });

    const hoje = new Date();
    const janelaIni = new Date(hoje.getFullYear(), hoje.getMonth() - AUS_JANELA_MESES_ATRAS, 1);
    const janelaFim = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + AUS_JANELA_DIAS_FRENTE);
    const janelaIniIso = isoData_(janelaIni), janelaFimIso = isoData_(janelaFim);

    let novas = [];
    const resumo = [];
    const naoVinculados = {};
    const timesAgendaSync = new Set();
    const timesFeriasSync = new Set();

    comFonte.forEach(t => {
      const colabsTime = colaboradores.filter(c => c.time === t.time);
      const contatos = t.planilhaFerias ? lerContatosFerias_(t.planilhaFerias) : {};
      const matcher = criarMatcherColaborador_(colabsTime, contatos);

      if (t.planilhaFerias) {
        const r = lerPlanilhaFerias_(t.time, t.planilhaFerias, matcher);
        novas = novas.concat(r.linhas);
        r.naoVinculados.forEach(n => naoVinculados[n] = true);
        timesFeriasSync.add(t.time);
        resumo.push(`${t.time} · férias/folgas: ${r.linhas.length}`);
      }
      if (t.agenda) {
        const r = lerAgendaAusencias_(t.time, t.agenda, janelaIni, janelaFim, matcher, preservar);
        novas = novas.concat(r.linhas);
        r.naoVinculados.forEach(n => naoVinculados[n] = true);
        timesAgendaSync.add(t.time);
        resumo.push(`${t.time} · agenda: ${r.linhas.length}`);
      }
    });

    // Mantém: manuais, outros times, e eventos de agenda fora da janela (histórico)
    const mantidas = atuais.filter(a => {
      if (a.origem === ORIGEM_MANUAL || !a.origem) return true;
      if (a.origem === ORIGEM_FERIAS) return !timesFeriasSync.has(a.time);
      if (a.origem === ORIGEM_AGENDA) return !timesAgendaSync.has(a.time) || a.inicio < janelaIniIso || a.inicio > janelaFimIso;
      return true;
    });

    const final = deduplicarAusencias_(mantidas.concat(novas));
    final.sort((a, b) => (a.time + a.inicio + a.nome).localeCompare(b.time + b.inicio + b.nome));
    gravarAusencias_(final);

    const nv = Object.keys(naoVinculados);
    let msg = `Ausências sincronizadas: ${resumo.join(' | ')}. Total na aba: ${final.length}.`;
    if (nv.length) msg += ` Sem colaborador na CONFIG_TIME (ignorados): ${nv.slice(0, 8).join(', ')}${nv.length > 8 ? '…' : ''}.`;
    registrarLog_('Sincronização de Ausências', 'SUCESSO', msg);
    return msg;
  } finally {
    lock.releaseLock();
  }
}

function obterFontesAusenciaPorTime_() {
  return listarCadastro('times').linhas.map(t => ({
    time: t.time,
    agenda: String(t.agenda || '').trim(),
    planilhaFerias: String(t.planilhaFerias || '').trim()
  })).filter(t => t.time);
}

/* ---------------------------- Google Agenda ---------------------------- */
function lerAgendaAusencias_(time, idAgenda, ini, fim, matcher, preservar) {
  // Aceita o ID da agenda (c_...@group.calendar.google.com) ou a planilha
  // "Agendamento Ausências" alimentada a partir da agenda (ID ou URL)
  if (idAgenda.indexOf('@') < 0) return lerPlanilhaAgendamento_(time, idAgenda, ini, fim, matcher, preservar);

  const agenda = CalendarApp.getCalendarById(idAgenda);
  if (!agenda) throw new Error(`Agenda do time "${time}" não encontrada ou sem acesso: ${idAgenda}`);

  const linhas = [];
  const naoVinculados = [];
  agenda.getEvents(ini, fim).forEach(ev => {
    const titulo = String(ev.getTitle() || '').trim();
    const criador = (ev.getCreators() || [])[0] || '';
    const diaInteiro = ev.isAllDayEvent();
    const disponivel = eventoDisponivel_(ev);

    let inicio, fimDia, horas = '';
    if (diaInteiro) {
      inicio = ev.getAllDayStartDate();
      fimDia = new Date(ev.getAllDayEndDate().getTime() - 86400000); // fim é exclusivo
    } else {
      inicio = ev.getStartTime();
      fimDia = ev.getEndTime();
      const mesmoDia = isoData_(inicio) === isoData_(fimDia);
      if (mesmoDia) {
        horas = Math.round((fimDia - inicio) / 36e5 * 100) / 100;
        if (!(horas > 0)) return;
        horas = Math.min(horas, CONFIG.HORAS_DIA);
      } // evento com horário em vários dias = dias inteiros
    }

    const chave = ev.getId() + '|' + isoData_(inicio);
    const anterior = preservar[chave];
    let colab = matcher(titulo, criador);
    if (!colab && anterior && anterior.usuario) colab = { usuario: anterior.usuario, nome: anterior.nome };
    if (!colab) { naoVinculados.push(titulo || criador); return; }

    let tipo = disponivel && diaInteiro ? TIPO_INFORMATIVO : inferirTipoAusencia_(titulo);
    if (anterior && anterior.tipo && anterior.usuario === colab.usuario) tipo = anterior.tipo;

    linhas.push({
      time: time, usuario: colab.usuario, nome: colab.nome,
      inicio: isoData_(inicio), fim: isoData_(fimDia), tipo: tipo,
      horas: horas === '' ? '' : String(horas),
      origem: ORIGEM_AGENDA, idEvento: ev.getId(),
      obs: titulo + (disponivel && diaInteiro ? ' (marcado como Disponível na agenda)' : '')
    });
  });
  return { linhas: linhas, naoVinculados: naoVinculados };
}

/* ------------------- Planilha "Agendamento Ausências" -------------------
 * Colunas: Título | Criador/Autor | Data Início | Hora Início | Data Término |
 *          Hora Término | ID do Evento   ("Dia Inteiro" = evento de dia inteiro,
 *          com Data Término exclusiva, como na agenda)
 * ------------------------------------------------------------------------ */
function lerPlanilhaAgendamento_(time, ref, ini, fim, matcher, preservar) {
  let ss;
  try {
    ss = /^https?:/i.test(ref) ? SpreadsheetApp.openByUrl(ref) : SpreadsheetApp.openById(ref);
  } catch (e) {
    throw new Error(`Agenda do time "${time}": não consegui abrir a planilha ${ref} (${e.message}). Informe o ID da agenda (c_...@group.calendar.google.com) ou uma planilha com acesso.`);
  }
  const sh = ss.getSheetByName('Página1') || ss.getSheets()[0];
  const linhas = [];
  const naoVinculados = [];
  if (sh.getLastRow() < 2) return { linhas: linhas, naoVinculados: naoVinculados };

  const iniIso = isoData_(ini), fimIso = isoData_(fim);
  const vistos = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 7).getDisplayValues().forEach(r => {
    const titulo = String(r[0] || '').trim();
    const autor = String(r[1] || '').trim();
    const dIni = normalizarData_(r[2]);
    let dFim = normalizarData_(r[4]) || dIni;
    const hIni = String(r[3] || '').trim(), hFim = String(r[5] || '').trim();
    const id = String(r[6] || '').replace(/@google\.com$/, '').trim();
    if (!dIni || (!titulo && !autor)) return;

    const diaInteiro = /dia inteiro/i.test(hIni) || /dia inteiro/i.test(hFim);
    let horas = '';
    if (diaInteiro) {
      if (dFim > dIni) dFim = new Date(dFim.getTime() - 86400000); // término exclusivo
    } else if (isoData_(dIni) === isoData_(dFim)) {
      horas = Math.min(Math.round(diferencaHoras_(hIni, hFim) * 100) / 100, CONFIG.HORAS_DIA);
      if (!(horas > 0)) return;
    }

    const inicioIso = isoData_(dIni);
    if (inicioIso > fimIso || isoData_(dFim) < iniIso) return;
    const chave = (id || titulo) + '|' + inicioIso + '|' + hIni;
    if (vistos[chave]) return; // linhas repetidas na planilha
    vistos[chave] = true;

    const anterior = preservar[(id || titulo) + '|' + inicioIso];
    let colab = matcher(titulo, '') || matcher(autor, '');
    if (!colab && anterior && anterior.usuario) colab = { usuario: anterior.usuario, nome: anterior.nome };
    if (!colab) { naoVinculados.push(titulo || autor); return; }

    let tipo = inferirTipoAusencia_(titulo);
    if (anterior && anterior.tipo && anterior.usuario === colab.usuario) tipo = anterior.tipo;

    linhas.push({
      time: time, usuario: colab.usuario, nome: colab.nome,
      inicio: inicioIso, fim: isoData_(dFim), tipo: tipo,
      horas: horas === '' ? '' : String(horas),
      origem: ORIGEM_AGENDA, idEvento: id || '', obs: titulo
    });
  });
  return { linhas: linhas, naoVinculados: naoVinculados };
}

function diferencaHoras_(hIni, hFim) {
  const m1 = String(hIni).match(/(\d{1,2}):(\d{2})/), m2 = String(hFim).match(/(\d{1,2}):(\d{2})/);
  if (!m1 || !m2) return 0;
  return ((+m2[1] * 60 + +m2[2]) - (+m1[1] * 60 + +m1[2])) / 60;
}

/* ---------------------------- Planilha de Férias ---------------------------- */
function lerPlanilhaFerias_(time, url, matcher) {
  const ss = SpreadsheetApp.openByUrl(url);
  const linhas = [];
  const naoVinculados = [];

  // Controle Consolidado (férias programadas, em gozo e finalizadas)
  ss.getSheets().filter(s => /^controle consolidado/i.test(s.getName())).forEach(sh => {
    if (sh.getLastRow() < 2) return;
    const dados = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
    const h = dados[0].map(x => normalizarTexto_(x));
    const col = nome => h.findIndex(x => x.indexOf(nome) === 0);
    const iNome = col('nome'), iSit = col('situacao'), iAg = col('agendamento');
    const iIni = col('data inicio programacao'), iFim = col('data fim programacao');
    if ([iNome, iIni, iFim].some(i => i < 0)) return;

    for (let r = 1; r < dados.length; r++) {
      const nome = dados[r][iNome];
      const sit = normalizarTexto_(dados[r][iSit] || '');
      const ag = normalizarTexto_(dados[r][iAg] || '');
      const ini = normalizarData_(dados[r][iIni]);
      const fim = normalizarData_(dados[r][iFim]);
      if (!nome || !ini || !fim || fim < ini) continue;
      const confirmada = /marcad|em gozo|finalizad/.test(sit) || ag === 'agendados' || ag === 'agendado';
      if (!confirmada) continue;
      const colab = matcher(nome, '');
      if (!colab) { naoVinculados.push(nome); continue; }
      linhas.push({
        time: time, usuario: colab.usuario, nome: colab.nome,
        inicio: isoData_(ini), fim: isoData_(fim), tipo: 'Férias', horas: '',
        origem: ORIGEM_FERIAS, idEvento: '', obs: `${dados[r][iSit] || ''} · ${dados[r][iAg] || ''}`.replace(/^[\s·-]+|[\s·-]+$/g, '')
      });
    }
  });

  // Folgas e Emendas (GIFTCARD, RESETDAY, MVT...)
  ss.getSheets().filter(s => /^folgas/i.test(s.getName())).forEach(sh => {
    if (sh.getLastRow() < 2) return;
    const dados = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
    const h = dados[0].map(x => normalizarTexto_(x));
    const col = nome => h.findIndex(x => x.indexOf(nome) === 0);
    const iNome = col('colaborador'), iBen = col('beneficio'), iIni = col('data da folga'), iQtd = col('quantidade'), iRet = col('data de retorno');
    if (iNome < 0 || iIni < 0) return;
    for (let r = 1; r < dados.length; r++) {
      const nome = dados[r][iNome];
      const ini = normalizarData_(dados[r][iIni]);
      if (!nome || !ini) continue;
      const qtd = Math.max(1, converterNumero_(iQtd >= 0 ? dados[r][iQtd] : 1));
      const fim = ultimoDiaUtilApos_(ini, qtd);
      const colab = matcher(nome, '');
      if (!colab) { naoVinculados.push(nome); continue; }
      linhas.push({
        time: time, usuario: colab.usuario, nome: colab.nome,
        inicio: isoData_(ini), fim: isoData_(fim), tipo: 'Folga', horas: '',
        origem: ORIGEM_FERIAS, idEvento: '', obs: dados[r][iBen] || 'Folga'
      });
    }
  });

  return { linhas: linhas, naoVinculados: naoVinculados };
}

/** Aba "Contatos" (Nome, Email) da planilha de férias: email -> nome. */
function lerContatosFerias_(url) {
  try {
    const sh = SpreadsheetApp.openByUrl(url).getSheetByName('Contatos');
    if (!sh || sh.getLastRow() < 2) return {};
    const mapa = {};
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(r => {
      if (r[0] && r[1]) mapa[String(r[1]).trim().toLowerCase()] = r[0];
    });
    return mapa;
  } catch (e) { return {}; }
}

/* =============================================================================
 * IDENTIFICAÇÃO DO COLABORADOR
 * Ex.: "Ausência: JOSE RIQUELMO GOMES DA SILVA [Marcadas]", "Keilla - Consulta",
 *      "Fisioterapia Ombro | Fabio Mendonça", "Consulta - Pâmela", "Edu - Exame",
 *      "Ausência Carlão" (via criador do evento), "Consulta médica" (criador)
 * ===========================================================================*/
function criarMatcherColaborador_(colabs, contatos) {
  const lista = colabs.map(c => ({
    usuario: c.usuario, nome: c.nome,
    sim: String(c.considerar).toUpperCase() !== 'NÃO',
    tokens: normalizarTexto_(c.nome).split(/[^a-z]+/).filter(t => t.length > 2 && ['dos', 'das', 'de', 'da', 'do'].indexOf(t) < 0)
  }));

  function porTexto(texto) {
    const tk = normalizarTexto_(String(texto).replace(/\[.*?\]/g, ' ')).split(/[^a-z]+/).filter(t => t.length >= 3);
    if (!tk.length) return null;
    let melhor = [], melhorScore = 0;
    lista.forEach(c => {
      let score = 0;
      tk.forEach(t => {
        if (c.tokens[0] === t) score += 3;               // primeiro nome pesa mais
        else if (c.tokens.indexOf(t) >= 0) score += 2;
        else if (t.length >= 3 && c.tokens[0] && c.tokens[0].indexOf(t) === 0) score += 1; // apelido: "edu" -> eduardo
      });
      if (score > melhorScore) { melhorScore = score; melhor = [c]; }
      else if (score === melhorScore && score > 0) melhor.push(c);
    });
    if (!melhorScore) return null;
    const preferidos = melhor.filter(c => c.sim);
    const escolhidos = preferidos.length ? preferidos : melhor;
    return escolhidos.length === 1 ? escolhidos[0] : null; // ambíguo -> tenta o criador
  }

  function porEmail(email) {
    email = String(email || '').toLowerCase().trim();
    if (!email) return null;
    const prefixo = email.split('@')[0];
    const direto = lista.filter(c => c.usuario === prefixo);
    if (direto.length === 1 && direto[0].sim) return direto[0];
    if (contatos[email]) { const c = porTexto(contatos[email]); if (c) return c; }
    return direto[0] || null;
  }

  return function (titulo, criador) {
    const c = porTexto(titulo) || porEmail(criador);
    return c ? { usuario: c.usuario, nome: c.nome } : null;
  };
}

function inferirTipoAusencia_(titulo) {
  const t = normalizarTexto_(titulo);
  // Espelhos da planilha de férias: "Ausência: NOME [Marcadas]", "[Férias + RESETDAY]", "[Sugestão do Dev]"
  const status = (t.match(/\[(.*)\]/) || [])[1] || '';
  if (/marcad|ferias|em gozo|agendad|sugest/.test(status)) return 'Férias';
  if (/resetday|giftcard|folga|mvt/.test(status)) return 'Folga';
  if (/ferias/.test(t)) return 'Férias';
  if (/resetday|giftcard|folga|emenda|mvt|day ?off/.test(t)) return 'Folga';
  if (/banco de horas/.test(t)) return 'Banco de horas';
  if (/atestado/.test(t)) return 'Atestado';
  if (/licenca/.test(t)) return 'Licença';
  if (/curso|treinamento|letramento|workshop|palestra|evento/.test(t)) return 'Treinamento';
  if (/consulta|medic|exame|dentista|fisio|nutri|terapia|cardio|hospital|cirurgia|vacina|periodico|ocupacional/.test(t)) return 'Consulta/Exame';
  return 'Ausência';
}

/* =============================================================================
 * GRAVAÇÃO E DEDUPLICAÇÃO
 * ===========================================================================*/
function deduplicarAusencias_(lista) {
  // Prioridade quando o mesmo período aparece em mais de uma fonte
  const prio = { [ORIGEM_MANUAL]: 0, [ORIGEM_FERIAS]: 1, [ORIGEM_AGENDA]: 2 };
  const ordenada = lista.slice().sort((a, b) => (prio[a.origem] || 0) - (prio[b.origem] || 0));
  const vistos = {};
  return ordenada.filter(a => {
    const k = [a.usuario, a.inicio, a.fim, a.horas || 'integral'].join('|');
    // eventos parciais diferentes no mesmo dia são mantidos (id distinto)
    const kk = a.horas ? k + '|' + (a.idEvento || a.obs) : k;
    if (vistos[kk]) return false;
    vistos[kk] = true;
    return true;
  });
}

function gravarAusencias_(linhas) {
  const def = CADASTROS.ausencias;
  const sh = garantirAbaCadastro_('ausencias');
  const nCols = def.colunas.length;
  const ultima = sh.getLastRow();
  if (ultima > 1) sh.getRange(2, 1, ultima - 1, nCols).clearContent();
  if (!linhas.length) return;
  const matriz = linhas.map(o => def.colunas.map(c => {
    const v = o[c.key];
    if (c.data) return v ? isoParaData_(v) : '';
    if (c.key === 'horas') return v === '' || v == null ? '' : converterNumero_(v);
    return v == null ? '' : String(v);
  }));
  sh.getRange(2, 1, matriz.length, nCols).setValues(matriz);
  def.colunas.forEach((c, j) => { if (c.data) sh.getRange(2, j + 1, matriz.length, 1).setNumberFormat('dd/MM/yyyy'); });
}

/* =============================================================================
 * MAPA PARA O CÁLCULO: usuario|yyyy-MM-dd -> { integral, horas, descricao }
 * ===========================================================================*/
function carregarMapaAusencias_() {
  const mapa = {};
  let linhas = [];
  try { linhas = listarCadastro('ausencias').linhas; } catch (e) { return mapa; }

  linhas.forEach(a => {
    if (!a.usuario || !a.inicio || a.tipo === TIPO_INFORMATIVO) return;
    const ini = isoParaData_(a.inicio);
    const fim = isoParaData_(a.fim || a.inicio);
    const horas = a.horas === '' || a.horas == null ? null : converterNumero_(a.horas);
    const desc = a.tipo + (a.obs ? ': ' + a.obs : '');
    for (let d = new Date(ini); d <= fim; d.setDate(d.getDate() + 1)) {
      const k = a.usuario.toLowerCase() + '|' + isoData_(d);
      const m = mapa[k] || (mapa[k] = { integral: false, horas: 0, descricoes: [] });
      if (horas === null) m.integral = true; else m.horas += horas;
      if (m.descricoes.indexOf(desc) < 0) m.descricoes.push(desc);
    }
  });
  return mapa;
}

/** Horas abatidas no dia (0..HORAS_DIA). */
function horasAbatidasNoDia_(mapaAus, usuario, dataIso) {
  const m = mapaAus[usuario + '|' + dataIso];
  if (!m) return { horas: 0, descricao: '' };
  const horas = m.integral ? CONFIG.HORAS_DIA : Math.min(m.horas, CONFIG.HORAS_DIA);
  return { horas: Math.round(horas * 100) / 100, descricao: m.descricoes.join(' | ') };
}

/* =============================================================================
 * UTILITÁRIOS
 * ===========================================================================*/
function mensagemPermissao_(e) {
  const t = String(e && e.message || e);
  if (/permission|permiss|authoriz|autoriza/i.test(t)) {
    return 'Falta autorizar o acesso à Agenda/planilhas. Na planilha, vá no menu Capacity Multi-Time > Configurações > 🔐 Autorizar acessos, aceite as permissões e tente de novo.';
  }
  return t;
}
function eventoDisponivel_(ev) {
  try {
    const t = ev.getTransparency();
    return !!t && String(t).toUpperCase().indexOf('TRANSPARENT') >= 0;
  } catch (e) { return false; }
}
function normalizarTexto_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
}
function isoData_(d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd'); }
function ultimoDiaUtilApos_(ini, qtdDias) {
  const d = new Date(ini);
  let contados = 1;
  while (contados < qtdDias) {
    d.setDate(d.getDate() + 1);
    if (d.getDay() !== 0 && d.getDay() !== 6) contados++;
  }
  return d;
}
