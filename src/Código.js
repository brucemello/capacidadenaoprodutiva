/*******************************************************************************
 * CONTROLE MENSAL - CAPACIDADE NÃO PRODUTIVA
 * Versão Final Avançada e Blindada (com Nova Aba Resumo do % Não Produtivo por Time)
 *******************************************************************************/

const CONFIG = {
  TIMEZONE: 'America/Sao_Paulo',
  HORAS_DIA: 8,

  SHEET_IMPORTAR: 'TIMESHEET_IMPORTAR',
  SHEET_TIMESHEET: 'TIMESHEET',
  SHEET_BASE: 'BASE_COLABORADOR',
  SHEET_HISTORICO: 'HISTORICO_MENSAL',
  SHEET_RESUMO_TIME: 'RESUMO_NAO_PRODUTIVO', // Nova aba dedicada
  SHEET_CONFIG: 'CONFIG_TIME',
  SHEET_FERIADOS: 'FERIADOS',
  SHEET_LOGS: 'LOGS_PROCESSAMENTO'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Capacity Multi-Time')
    .addItem('📈 Abrir Dashboard', 'abrirDashboard')
    .addItem('Abrir Painel Lateral', 'abrirPainelLateral')
    .addItem('🔄 Buscar Mês Fechado (API Tempo)', 'buscarMesFechadoPeloMenu')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Configurações')
      .addItem('Instalar/Atualizar Estrutura', 'inicializarOuAtualizarEstrutura')
      .addItem('⏰ Ativar Rotina Mensal Automática', 'ativarRotinaMensalPeloMenu')
      .addItem('🔐 Autorizar acessos (Agenda, Férias, Jira)', 'autorizarAcessos')
    )
    .addToUi();
}

function buscarMesFechadoPeloMenu() {
  const res = buscarTimesheetTempoPeloPainel('', null);
  SpreadsheetApp.getUi().alert(res.mensagem);
}

function ativarRotinaMensalPeloMenu() {
  SpreadsheetApp.getUi().alert(ativarRotinaMensal().mensagem);
}

function abrirPainelLateral() {
  const html = HtmlService.createHtmlOutputFromFile('Interface')
    .setTitle('Painel Capacity')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function obterListaDeTimes() {
  try {
    const ss = obterPlanilha_();
    const sh = ss.getSheetByName('TIME');
    if (!sh) return [];
    const dados = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    const lista = dados.map(r => r[0].toString().trim()).filter(r => r !== "");
    return [...new Set(lista)];
  } catch (e) {
    return [];
  }
}

function executarFuncaoPeloPainel(nomeFuncao, timeSelecionado) {
  try {
    if (timeSelecionado === "-- Todos os Times (Geral) --") {
      timeSelecionado = "";
    }

    if (nomeFuncao === 'processarTimesheetMensal') {
      processarTimesheetMensal(timeSelecionado);
      return { sucesso: true, mensagem: 'Processado com sucesso!' };
    } 
    else if (nomeFuncao === 'recalcularBaseEHistorico') {
      recalcularLinhasTimesheet_(timeSelecionado);
      return { sucesso: true, mensagem: 'Recalculado com sucesso!' };
    }
    throw new Error('Função inválida.');
  } catch (erro) {
    return { sucesso: false, mensagem: erro.toString() };
  }
}

function inicializarOuAtualizarEstrutura() {
  const ss = obterPlanilha_();
  const shConfig = getOrCreateSheet_(ss, CONFIG.SHEET_CONFIG);
  if (shConfig.getLastRow() < 1) {
    shConfig.appendRow(['Nome completo', 'Nome de Usuário', 'Considerar capacidade?', 'Time', 'Observação']);
    shConfig.getRange('A1:E1').setBackground('#0F5B78').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  getOrCreateSheet_(ss, CONFIG.SHEET_FERIADOS);
  getOrCreateSheet_(ss, CONFIG.SHEET_IMPORTAR);
  getOrCreateSheet_(ss, CONFIG.SHEET_TIMESHEET);
  getOrCreateSheet_(ss, CONFIG.SHEET_BASE);
  getOrCreateSheet_(ss, CONFIG.SHEET_HISTORICO);
  getOrCreateSheet_(ss, CONFIG.SHEET_RESUMO_TIME);
  getOrCreateSheet_(ss, CONFIG.SHEET_LOGS);
  garantirColunaTeamId_();
  registrarLog_('Inicialização do Sistema', 'SUCESSO', 'Estrutura de abas verificada e instalada.');
  SpreadsheetApp.getUi().alert('Estrutura validada com sucesso!');
}

function importarDadosJira(fileObject) {
  try {
    const data = Utilities.base64Decode(fileObject.data);
    const blob = Utilities.newBlob(data, fileObject.mimeType, fileObject.name);
    const tempFile = Drive.Files.create({ name: "TEMP_CAP_JIRA", mimeType: MimeType.GOOGLE_SHEETS }, blob);
    const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
    let tempSheet = tempSpreadsheet.getSheets()[0];
    const novosDados = tempSheet.getDataRange().getValues();
    const ss = obterPlanilha_();
    const shImportar = getOrCreateSheet_(ss, CONFIG.SHEET_IMPORTAR);
    shImportar.clear();
    if (novosDados.length > 0 && novosDados[0].length > 0) {
      shImportar.getRange(1, 1, novosDados.length, novosDados[0].length).setValues(novosDados);
    }
    
    Drive.Files.remove(tempFile.id);
    
    registrarLog_('Importação de Arquivo', 'SUCESSO', `Arquivo "${fileObject.name}" convertido. ${novosDados.length - 1} registros mapeados.`);
    return { sucesso: true, mensagem: `Importado com sucesso! ${novosDados.length - 1} linhas cadastradas.` };
  } catch (erro) {
    registrarLog_('Importação de Arquivo', 'ERRO', erro.toString());
    return { sucesso: false, mensagem: "Erro crítico: " + erro.toString() };
  }
}

/**
 * @param {string} timeSelecionado  Time do filtro do painel ('' = todos)
 * @param {Object} [opcoes]          Usado pela busca da API Tempo:
 *        inicio/fim (Date) = período exato do mês; timesProcessados = times buscados
 */
function processarTimesheetMensal(timeSelecionado, opcoes) {
  opcoes = opcoes || {};
  const ss = obterPlanilha_();
  const shImportar = getOrCreateSheet_(ss, CONFIG.SHEET_IMPORTAR);
  const shTimesheet = getOrCreateSheet_(ss, CONFIG.SHEET_TIMESHEET);

  const dados = shImportar.getDataRange().getValues();
  if (dados.length < 2) throw new Error('A aba TIMESHEET_IMPORTAR está vazia.');

  const header = dados[0].map(h => String(h).trim());
  const idxNome = header.indexOf('Nome completo');
  const idxUsuario = header.indexOf('Nome de Usuário');
  const idxDataTrabalho = header.indexOf('data de Trabalho');
  const idxHoras = header.indexOf('Horas');
  const idxTime = header.indexOf('Time');

  if (idxNome === -1 || idxUsuario === -1 || idxDataTrabalho === -1 || idxHoras === -1) {
    throw new Error(`Colunas necessárias não encontradas. Cabeçalho lido: ${header.join(' | ')}`);
  }

  const usuariosNoArquivoAtuais = new Set();
  const apontamentosJira = {};
  let dataMinima = null;
  let dataMaxima = null;
  const mapaUsuariosNomes = {};
  const mapaUsuariosTime = {};

  for (let r = 1; r < dados.length; r++) {
    const nome = dados[r][idxNome];
    const usuario = String(dados[r][idxUsuario]).trim().toLowerCase();
    const rawData = dados[r][idxDataTrabalho];
    const horasLançadas = converterNumero_(dados[r][idxHoras]);

    if (!nome || !dados[r][idxUsuario] || !rawData) continue;
    const dataObj = normalizarData_(rawData);
    if (!dataObj || isNaN(dataObj.getTime())) continue;

    usuariosNoArquivoAtuais.add(usuario);

    if (!dataMinima || dataObj < dataMinima) dataMinima = new Date(dataObj.getTime());
    if (!dataMaxima || dataObj > dataMaxima) dataMaxima = new Date(dataObj.getTime());

    const dataChave = Utilities.formatDate(dataObj, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const chaveUnica = usuario + '|' + dataChave;

    mapaUsuariosNomes[usuario] = nome;
    if (idxTime !== -1 && dados[r][idxTime]) mapaUsuariosTime[usuario] = String(dados[r][idxTime]).trim();

    if (!apontamentosJira[chaveUnica]) apontamentosJira[chaveUnica] = 0;
    apontamentosJira[chaveUnica] += horasLançadas;
  }

  if (Object.keys(mapaUsuariosTime).length > 0) {
    // Base vinda da API Tempo: cada usuário já traz o time de origem
    atualizarTimeNaAbaConfig_(usuariosNoArquivoAtuais, mapaUsuariosNomes, timeSelecionado, mapaUsuariosTime);
  } else if (timeSelecionado && timeSelecionado !== "") {
    atualizarTimeNaAbaConfig_(usuariosNoArquivoAtuais, mapaUsuariosNomes, timeSelecionado);
  }

  if (opcoes.inicio && opcoes.fim) {
    dataMinima = new Date(opcoes.inicio.getTime());
    dataMaxima = new Date(opcoes.fim.getTime());
  }
  if (!dataMinima || !dataMaxima) throw new Error('Nenhum apontamento válido encontrado para processar.');
  const timesPermitidos = opcoes.timesProcessados ? new Set(opcoes.timesProcessados) : null;

  const configTime = carregarConfigTime_();
  const feriados = carregarFeriados_();
  const mapaAusencias = carregarMapaAusencias_();

  const listaTodasDatas = [];
  let dataCorrente = new Date(dataMinima.getTime());
  while (dataCorrente <= dataMaxima) {
    listaTodasDatas.push(new Date(dataCorrente.getTime()));
    dataCorrente.setDate(dataCorrente.getDate() + 1);
  }

  const saida = [[
    'Mês referência', 'Time', 'Nome completo', 'Nome de Usuário', 'Data', 'Horas importadas',
    'Ajuste manual de horas', 'Horas finais', 'Dia útil?', 'Considerar capacidade?',
    'Horas esperadas', 'Horas não produtivas', '% não produtivo', 'Observação',
    'Horas ausência (abatidas)', 'Ausência (motivo)'
  ]];

  const usuariosMapeados = Object.keys(configTime);
  Object.keys(mapaUsuariosNomes).forEach(u => {
    if (!configTime[u]) usuariosMapeados.push(u);
  });
  
  const usuariosUnicos = [...new Set(usuariosMapeados)];
  usuariosUnicos.forEach(userChave => {
    const conf = configTime[userChave] || { 
      considerar: 'SIM', 
      time: 'Sem Time Mapeado', 
      nomeOriginal: mapaUsuariosNomes[userChave] || userChave 
    };

    if (timeSelecionado && timeSelecionado !== "" && conf.time !== timeSelecionado) {
      return; 
    }
    if (timesPermitidos && !timesPermitidos.has(conf.time)) {
      return;
    }

    const considerarFinal = String(conf.considerar).trim().toUpperCase();

    if (considerarFinal === 'NÃO') {
      return;
    }

    listaTodasDatas.forEach(diaObj => {
      const dataChave = Utilities.formatDate(diaObj, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      const chaveLookup = userChave + '|' + dataChave;
      const horasApontadas = apontamentosJira[chaveLookup] || 0;
      
      const mesReferenciaStr = Utilities.formatDate(diaObj, CONFIG.TIMEZONE, 'yyyy-MM-01');
      const partes = mesReferenciaStr.split('-');
      const mesReferencia = new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
      
      const diaUtil = isDiaUtil_(diaObj, feriados);
      const diaUtilTexto = diaUtil ? 'SIM' : 'NÃO';
      
      const aus = (diaUtil && considerarFinal === 'SIM') ? horasAbatidasNoDia_(mapaAusencias, userChave, dataChave) : { horas: 0, descricao: '' };
      const horasEsperadas = (diaUtil && considerarFinal === 'SIM') ? Math.max(CONFIG.HORAS_DIA - aus.horas, 0) : 0;
      const horasNaoProdutivas = Math.max(horasEsperadas - horasApontadas, 0);
      const percNaoProdutivo = horasEsperadas > 0 ? horasNaoProdutivas / horasEsperadas : 0;

      saida.push([
        mesReferencia, conf.time, conf.nomeOriginal, userChave, diaObj, horasApontadas, 
        '', horasApontadas, diaUtilTexto, considerarFinal, horasEsperadas, horasNaoProdutivas, percNaoProdutivo, '',
        aus.horas, aus.descricao
      ]);
    });
  });

  shTimesheet.clear();
  shTimesheet.getRange(1, 1, saida.length, saida[0].length).setValues(saida);
  
  shTimesheet.setFrozenRows(1);
  shTimesheet.getRange(1, 1, 1, saida[0].length).setBackground('#0F5B78').setFontColor('#FFFFFF').setFontWeight('bold');
  shTimesheet.getRange(2, 1, saida.length - 1, 1).setNumberFormat('yyyy-MM');
  shTimesheet.getRange(2, 5, saida.length - 1, 1).setNumberFormat('dd/MM/yyyy');
  shTimesheet.getRange(2, 13, saida.length - 1, 1).setNumberFormat('0.00%');
  shTimesheet.autoResizeColumns(1, saida[0].length);

  const avaliacaoConsistencia = validarConsistenciaDados_(saida);
  registrarLog_('Cálculo de Capacity', 'SUCESSO', avaliacaoConsistencia + (timeSelecionado ? ` | Filtro: ${timeSelecionado}` : ' | Geral'));

  atualizarBaseColaborador_();
}

function atualizarTimeNaAbaConfig_(usuariosSet, mapaNomes, timeSelecionado, mapaTimePorUsuario) {
  const ss = obterPlanilha_();
  const shConfig = getOrCreateSheet_(ss, CONFIG.SHEET_CONFIG);
  const dadosConfig = shConfig.getDataRange().getValues();
  
  const mapaLinhasConfig = {};
  for (let i = 1; i < dadosConfig.length; i++) {
    const user = String(dadosConfig[i][1]).trim().toLowerCase();
    if (user) mapaLinhasConfig[user] = i + 1;
  }
  
  usuariosSet.forEach(user => {
    const timeDoUsuario = (mapaTimePorUsuario && mapaTimePorUsuario[user]) || timeSelecionado;
    if (!timeDoUsuario) return;
    if (mapaLinhasConfig[user]) {
      shConfig.getRange(mapaLinhasConfig[user], 4).setValue(timeDoUsuario);
    } else {
      const novoNome = mapaNomes[user] || user;
      const origem = mapaTimePorUsuario ? 'Inserido automaticamente (API Tempo)' : 'Inserido via painel de importação';
      shConfig.appendRow([novoNome, user, 'SIM', timeDoUsuario, origem]);
      mapaLinhasConfig[user] = shConfig.getLastRow();
    }
  });
  SpreadsheetApp.flush();
}

function recalcularLinhasTimesheet_(timeSelecionado) {
  const ss = obterPlanilha_();
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_TIMESHEET);
  const dados = sh.getDataRange().getValues();
  if (dados.length < 2) return;

  const configTime = carregarConfigTime_();
  const mapaAusencias = carregarMapaAusencias_();

  // TIMESHEET gerada antes das ausências: garante as colunas O e P
  dados.forEach((linha, idx) => { while (linha.length < 16) linha.push(idx === 0 ? (linha.length === 14 ? 'Horas ausência (abatidas)' : 'Ausência (motivo)') : ''); });

  for (let i = 1; i < dados.length; i++) {
    const usuario = String(dados[i][3]).trim().toLowerCase();
    const conf = configTime[usuario] || { considerar: 'SIM', time: 'Sem Time Mapeado' };
    
    if (timeSelecionado && timeSelecionado !== "" && conf.time !== timeSelecionado) {
      continue;
    }

    const horasImportadas = converterNumero_(dados[i][5]);
    const ajusteManual = dados[i][6] !== '' ? converterNumero_(dados[i][6]) : '';
    const horasFinais = ajusteManual !== '' ? ajusteManual : horasImportadas;
    
    const considerarFinal = String(conf.considerar).trim().toUpperCase();
    const diaUtilTexto = String(dados[i][8]).trim().toUpperCase();
    
    const dataObj = normalizarData_(dados[i][4]);
    const dataIso = dataObj ? Utilities.formatDate(dataObj, CONFIG.TIMEZONE, 'yyyy-MM-dd') : '';
    const aus = (diaUtilTexto === 'SIM' && considerarFinal === 'SIM' && dataIso) ? horasAbatidasNoDia_(mapaAusencias, usuario, dataIso) : { horas: 0, descricao: '' };
    const horasEsperadas = (diaUtilTexto === 'SIM' && considerarFinal === 'SIM') ? Math.max(CONFIG.HORAS_DIA - aus.horas, 0) : 0;
    const horasNaoProdutivas = Math.max(horasEsperadas - horasFinais, 0);
    const percNaoProdutivo = horasEsperadas > 0 ? horasNaoProdutivas / horasEsperadas : 0;

    dados[i][14] = aus.horas;
    dados[i][15] = aus.descricao;
    dados[i][1] = conf.time;
    dados[i][7] = horasFinais;
    dados[i][9] = considerarFinal;
    dados[i][10] = horasEsperadas;
    dados[i][11] = horasNaoProdutivas;
    dados[i][12] = percNaoProdutivo;
  }
  sh.getRange(1, 1, dados.length, dados[0].length).setValues(dados);
  
  registrarLog_('Recálculo de Ajustes', 'SUCESSO', 'Ajustes manuais e ausências integrados à base analítica' + (timeSelecionado ? ` para o time ${timeSelecionado}.` : '.'));
  atualizarBaseColaborador_();
}

function atualizarBaseColaborador_() {
  const ss = obterPlanilha_();
  const shTimesheet = getOrCreateSheet_(ss, CONFIG.SHEET_TIMESHEET);
  const shBase = getOrCreateSheet_(ss, CONFIG.SHEET_BASE);
  const shHistorico = getOrCreateSheet_(ss, CONFIG.SHEET_HISTORICO);

  const dados = shTimesheet.getDataRange().getValues();
  if (dados.length < 2) return;

  const resumo = {};
  for (let i = 1; i < dados.length; i++) {
    const mes = dados[i][0];
    const time = dados[i][1];
    const nome = dados[i][2];
    const usuario = dados[i][3];
    if (!mes || !nome || !usuario) continue;
    
    let chaveMes = (mes instanceof Date) ? Utilities.formatDate(mes, CONFIG.TIMEZONE, 'yyyy-MM-01') : String(mes);
    const chave = chaveMes + '|' + time + '|' + usuario;
    
    if (!resumo[chave]) {
      resumo[chave] = { mes, time, nome, usuario, considerar: dados[i][9], exp: 0, ap: 0, aus: 0 };
    }
    resumo[chave].exp += converterNumero_(dados[i][10]);
    resumo[chave].ap += converterNumero_(dados[i][7]);
    resumo[chave].aus += converterNumero_(dados[i][14]);
  }

  const saida = [[
    'Mês referência', 'Time', 'Nome completo', 'Nome de Usuário', 'Considerar capacidade?',
    'Horas esperadas', 'Horas apontadas', 'Horas não produtivas', '% não produtivo', 'Horas ausência (abatidas)'
  ]];
  
  Object.values(resumo).forEach(item => {
    const np = Math.max(item.exp - item.ap, 0);
    const perc = item.exp > 0 ? np / item.exp : 0;
    
    saida.push([item.mes, item.time, item.nome, item.usuario, item.considerar, item.exp, item.ap, np, perc, Math.round(item.aus * 100) / 100]);
  });
  
  shBase.clear();
  shBase.getRange(1, 1, saida.length, saida[0].length).setValues(saida);
  
  shBase.getRange(2, 1, saida.length - 1, 1).setNumberFormat('yyyy-MM');
  shBase.getRange(2, 6, saida.length - 1, 3).setNumberFormat('#,##0.00');
  shBase.getRange(2, 9, saida.length - 1, 1).setNumberFormat('0.00%');
  shBase.getRange(2, 10, saida.length - 1, 1).setNumberFormat('#,##0.00');
  
  atualizarHistorico_(shHistorico, saida);
}

function atualizarHistorico_(shHistorico, saidaBase) {
  if (saidaBase.length < 2) return;
  const chavesProcessadas = [];
  
  for (let i = 1; i < saidaBase.length; i++) {
    let mVal = saidaBase[i][0];
    let mStr = extrairAnoMesSemTimezone_(mVal);
    
    const chave = mStr + '|' + saidaBase[i][1];
    if (!chavesProcessadas.includes(chave)) chavesProcessadas.push(chave);
  }

  let historicoValores = shHistorico.getDataRange().getValues();
  let historicoExibido = shHistorico.getDataRange().getDisplayValues();
  
  const novoHistorico = [saidaBase[0]];
  
  if (historicoValores.length > 1 && historicoValores[0][0] !== '') {
    for (let i = 1; i < historicoValores.length; i++) {
      let hValVisual = historicoExibido[i][0]; 
      let hStr = extrairAnoMesSemTimezone_(hValVisual);
      
      const chaveHist = hStr + '|' + historicoValores[i][1];
      
      if (!chavesProcessadas.includes(chaveHist)) {
        novoHistorico.push(historicoValores[i]);
      }
    }
  }

  for (let i = 1; i < saidaBase.length; i++) { 
    novoHistorico.push(saidaBase[i]);
  }
  // Meses antigos (antes das ausências) têm menos colunas: completa com vazio
  const largura = novoHistorico[0].length;
  for (let i = 0; i < novoHistorico.length; i++) {
    novoHistorico[i] = novoHistorico[i].slice(0, largura);
    while (novoHistorico[i].length < largura) novoHistorico[i].push('');
  }
  shHistorico.clear();
  shHistorico.getRange(1, 1, novoHistorico.length, novoHistorico[0].length).setValues(novoHistorico);
  shHistorico.getRange(2, 10, Math.max(novoHistorico.length - 1, 1), 1).setNumberFormat('#,##0.00');
  
  shHistorico.getRange(2, 1, novoHistorico.length - 1, 1).setNumberFormat('yyyy-MM');
  shHistorico.getRange(2, 6, novoHistorico.length - 1, 3).setNumberFormat('#,##0.00');
  shHistorico.getRange(2, 9, novoHistorico.length - 1, 1).setNumberFormat('0.00%');

  // GERA A NOVA ABA CONSOLIDADA DE FORMA TOTALMENTE ISOLADA
  gerarAbaResumoNaoProdutivo_(chavesProcessadas);
}

/**
 * NOVA FUNÇÃO ISOLADA:
 * Gera a aba RESUMO_NAO_PRODUTIVO com o % não produtivo exato do time por Mês/Time
 */
function gerarAbaResumoNaoProdutivo_(chavesProcessadas) {
  const ss = obterPlanilha_();
  const shHistorico = getOrCreateSheet_(ss, CONFIG.SHEET_HISTORICO);
  const shResumo = getOrCreateSheet_(ss, CONFIG.SHEET_RESUMO_TIME);

  const dadosHist = shHistorico.getDataRange().getValues();
  if (dadosHist.length < 2) return;

  const consolidado = {};

  for (let i = 1; i < dadosHist.length; i++) {
    const mes = dadosHist[i][0];
    const time = dadosHist[i][1];
    const exp = converterNumero_(dadosHist[i][5]);
    const ap = converterNumero_(dadosHist[i][6]);
    const np = converterNumero_(dadosHist[i][7]);
    const aus = converterNumero_(dadosHist[i][9]);

    if (!mes || !time) continue;

    let chaveMes = (mes instanceof Date) ? Utilities.formatDate(mes, CONFIG.TIMEZONE, 'yyyy-MM-01') : String(mes);
    const chave = chaveMes + '|' + time;

    if (!consolidado[chave]) {
      consolidado[chave] = { mesObj: mes, time, totalExp: 0, totalAp: 0, totalNp: 0, totalAus: 0 };
    }
    consolidado[chave].totalAus += aus;

    consolidado[chave].totalExp += exp;
    consolidado[chave].totalAp += ap;
    consolidado[chave].totalNp += np;
  }

  const saida = [[
    'Mês referência', 'Time', 'Total Horas Esperadas', 'Total Horas Apontadas', 'Total Horas Não Produtivas', '% Não Produtivo do Time', 'Total Horas Ausência (abatidas)', 'Atualizado em'
  ]];

  // "Atualizado em": mantém a data dos meses/times não reprocessados agora
  const agora = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
  const atualizadoAnterior = {};
  if (shResumo.getLastRow() > 1 && shResumo.getLastColumn() >= 8) {
    shResumo.getRange(2, 1, shResumo.getLastRow() - 1, 8).getDisplayValues().forEach(r => {
      atualizadoAnterior[extrairAnoMesSemTimezone_(r[0]) + '|' + r[1]] = r[7];
    });
  }
  const processadas = new Set(chavesProcessadas || []);

  Object.values(consolidado).forEach(item => {
    const chaveRes = extrairAnoMesSemTimezone_(item.mesObj) + '|' + item.time;
    // Aplica exatamente a fórmula do Looker Studio: SUM(Horas não produtivas) / SUM(Horas esperadas)
    const percTime = item.totalExp > 0 ? item.totalNp / item.totalExp : 0;
    
    saida.push([
      item.mesObj,
      item.time,
      item.totalExp,
      item.totalAp,
      item.totalNp,
      percTime,
      Math.round(item.totalAus * 100) / 100,
      processadas.has(chaveRes) || !atualizadoAnterior[chaveRes] ? agora : atualizadoAnterior[chaveRes]
    ]);
  });

  saida.splice(1, saida.length - 1, ...saida.slice(1).sort((a, b) =>
    (extrairAnoMesSemTimezone_(b[0]) + a[1]).localeCompare(extrairAnoMesSemTimezone_(a[0]) + b[1])));
  shResumo.clear();
  shResumo.getRange(1, 1, saida.length, saida[0].length).setValues(saida);
  shResumo.getRange(2, 8, Math.max(saida.length - 1, 1), 1).setNumberFormat('@');

  // Formatação visual da nova aba
  shResumo.setFrozenRows(1);
  shResumo.getRange(1, 1, 1, saida[0].length).setBackground('#0F5B78').setFontColor('#FFFFFF').setFontWeight('bold');
  shResumo.getRange(2, 1, saida.length - 1, 1).setNumberFormat('yyyy-MM');
  shResumo.getRange(2, 3, saida.length - 1, 3).setNumberFormat('#,##0.00');
  shResumo.getRange(2, 6, saida.length - 1, 1).setNumberFormat('0.00%');
  shResumo.autoResizeColumns(1, saida[0].length);
}

function extrairAnoMesSemTimezone_(valor) {
  if (valor instanceof Date) {
    let d = new Date(valor.getTime() + (6 * 60 * 60 * 1000));
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, '0');
  }
  
  let texto = String(valor).trim();
  
  let regexIso = /^(\d{4})[-/](\d{2})/;
  if (regexIso.test(texto)) {
    let m = texto.match(regexIso);
    return m[1] + "-" + m[2];
  }
  
  let regexBr = /^(\d{2})[-/](\d{2})[-/](\d{4})/;
  if (regexBr.test(texto)) {
    let m = texto.match(regexBr);
    return m[3] + "-" + m[2];
  }
  
  return texto.substring(0, 7);
}

function carregarConfigTime_() {
  const ss = obterPlanilha_();
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_CONFIG);
  const dados = sh.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < dados.length; i++) {
    const nomeOriginal = dados[i][0];
    const usuario = String(dados[i][1]).trim().toLowerCase();
    const considerar = dados[i][2] || 'SIM';
    const time = dados[i][3] || 'Sem Time Mapeado';
    if (dados[i][1]) {
      config[usuario] = { considerar, time, nomeOriginal };
    }
  }
  return config;
}

function carregarFeriados_() {
  const ss = obterPlanilha_();
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_FERIADOS);
  const dados = sh.getDataRange().getDisplayValues();
  const feriados = {};
  for (let i = 1; i < dados.length; i++) {
    const data = normalizarData_(dados[i][0]);
    if (data) feriados[Utilities.formatDate(data, CONFIG.TIMEZONE, 'yyyy-MM-dd')] = true;
  }
  return feriados;
}

function isDiaUtil_(data, feriados) {
  const dia = data.getDay();
  if (dia === 0 || dia === 6) return false;
  if (feriados[Utilities.formatDate(data, CONFIG.TIMEZONE, 'yyyy-MM-dd')]) return false;
  return true;
}

function converterNumero_(valor) {
  if (typeof valor === 'number') return valor;
  // Número que a planilha exibiu como data (ex.: 7,25 -> 07/01/1900): volta ao valor serial
  if (valor instanceof Date) {
    const serial = (valor.getTime() - new Date(1899, 11, 30).getTime()) / 86400000;
    return serial > 0 && serial < 1000 ? Math.round(serial * 10000) / 10000 : 0;
  }
  if (!valor) return 0;
  const texto = String(valor).trim();
  if (texto.includes(',') && texto.includes('.')) return Number(texto.replace(/\./g, '').replace(',', '.')) || 0;
  if (texto.includes(',')) return Number(texto.replace(',', '.')) || 0;
  return Number(texto) || 0;
}

function normalizarData_(valor) {
  if (valor instanceof Date) return valor;
  if (typeof valor === 'number') return new Date(Math.round((valor - 25569) * 86400 * 1000));
  const texto = String(valor).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(texto)) {
    const p = texto.split('-');
    return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2].substring(0,2)));
  }
  if (/^\d{2}\/\d{2}\/\d{4}/.test(texto)) {
    const p = texto.split('/');
    return new Date(Number(p[2].substring(0,4)), Number(p[1]) - 1, Number(p[0]));
  }
  return null;
}

function getOrCreateSheet_(ss, nome) {
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  return sh;
}

function registrarLog_(acao, status, detalhe) {
  try {
    const ss = obterPlanilha_();
    const sh = getOrCreateSheet_(ss, CONFIG.SHEET_LOGS);
    if (sh.getLastRow() < 1) {
      sh.appendRow(['Data/Hora', 'Operador', 'Ação', 'Status', 'Detalhes Matemáticos']);
      sh.getRange('A1:E1').setBackground('#0F5B78').setFontColor('#FFFFFF').setFontWeight('bold');
    }
    const dataHora = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    const operador = Session.getActiveUser().getEmail() || 'Automático';
    sh.appendRow([dataHora, operador, acao, status, detalhe]);
  } catch (e) { console.error(e); }
}

function validarConsistenciaDados_(saida) {
  if (saida.length < 2) return "Sem dados para avaliar.";
  let imp = 0, esp = 0;
  const users = new Set();
  for (let i = 1; i < saida.length; i++) {
    imp += saida[i][5];
    esp += saida[i][10];
    users.add(saida[i][3]);
  }
  return `Auditado: ${users.size} colaboradores identificados. Total apontado Jira: ${imp.toFixed(1)}h | Total esperado por contrato: ${esp.toFixed(1)}h.`;
}