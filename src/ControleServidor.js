/*******************************************************************************
 * CADASTROS DE CONTROLE (usados pelas abas Geração e Controle do dashboard)
 * TIME, CONFIG_TIME, FERIADOS e AUSENCIAS
 *******************************************************************************/

const SHEET_AUSENCIAS = 'AUSENCIAS';

const TIPOS_AUSENCIA = ['Férias', 'Folga', 'Consulta/Exame', 'Atestado', 'Licença', 'Treinamento', 'Banco de horas', 'Ausência', 'Informativo', 'Outro'];

/** Layout fixo de cada cadastro (ordem = colunas A, B, C...). */
const CADASTROS = {
  times: {
    aba: 'TIME',
    colunas: [
      { key: 'time', titulo: 'Lista de Times' },
      { key: 'teamId', titulo: 'Team ID (Tempo)' },
      { key: 'agenda', titulo: 'Agenda de Ausências (ID Google Agenda)' },
      { key: 'planilhaFerias', titulo: 'Planilha de Férias (URL)' }
    ]
  },
  colaboradores: {
    aba: 'CONFIG_TIME',
    colunas: [
      { key: 'nome', titulo: 'Nome completo' },
      { key: 'usuario', titulo: 'Nome de Usuário' },
      { key: 'considerar', titulo: 'Considerar capacidade?' },
      { key: 'time', titulo: 'Time' },
      { key: 'obs', titulo: 'Observação' }
    ]
  },
  feriados: {
    aba: 'FERIADOS',
    colunas: [
      { key: 'data', titulo: 'Data', data: true },
      { key: 'descricao', titulo: 'Descrição' }
    ]
  },
  ausencias: {
    aba: SHEET_AUSENCIAS,
    colunas: [
      { key: 'time', titulo: 'Time' },
      { key: 'usuario', titulo: 'Nome de Usuário' },
      { key: 'nome', titulo: 'Nome completo' },
      { key: 'inicio', titulo: 'Data início', data: true },
      { key: 'fim', titulo: 'Data fim', data: true },
      { key: 'tipo', titulo: 'Tipo' },
      { key: 'horas', titulo: 'Horas por dia (vazio = dia inteiro)', num: true },
      { key: 'origem', titulo: 'Origem' },
      { key: 'idEvento', titulo: 'ID Evento Agenda' },
      { key: 'obs', titulo: 'Observação' }
    ]
  }
};

/* =============================================================================
 * Leitura
 * ===========================================================================*/
function listarCadastro(tipo) {
  const def = CADASTROS[tipo];
  if (!def) throw new Error('Cadastro inválido: ' + tipo);
  const sh = garantirAbaCadastro_(tipo);
  const n = sh.getLastRow() - 1;
  const nCols = def.colunas.length;
  const linhas = [];
  if (n > 0) {
    const valores = sh.getRange(2, 1, n, nCols).getValues();
    const exibidos = sh.getRange(2, 1, n, nCols).getDisplayValues();
    for (let i = 0; i < n; i++) {
      const obj = {};
      let vazio = true;
      def.colunas.forEach((c, j) => {
        let v;
        if (c.data) {
          const d = normalizarData_(valores[i][j] instanceof Date ? valores[i][j] : exibidos[i][j]);
          v = d ? Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd') : '';
        } else if (c.num) {
          const raw = valores[i][j];
          v = raw === '' || raw == null ? '' : String(Math.round(converterNumero_(raw) * 100) / 100);
        } else {
          v = String(exibidos[i][j] == null ? '' : exibidos[i][j]).trim();
        }
        if (v) vazio = false;
        obj[c.key] = v;
      });
      // ignora linhas vazias e cabeçalhos repetidos
      if (vazio || def.colunas.some(c => obj[c.key] === c.titulo)) continue;
      linhas.push(obj);
    }
  }
  return { tipo: tipo, colunas: def.colunas, linhas: linhas };
}

/** Tudo que as abas Geração/Controle precisam em uma chamada. */
function obterDadosControle() {
  const res = {
    times: listarCadastro('times').linhas,
    colaboradores: listarCadastro('colaboradores').linhas,
    feriados: listarCadastro('feriados').linhas,
    ausencias: listarCadastro('ausencias').linhas,
    tiposAusencia: TIPOS_AUSENCIA,
    logs: listarLogs_(25)
  };
  try { res.config = obterConfiguracoesTempo(); } catch (e) { res.config = null; }
  return res;
}

function listarLogs_(qtd) {
  const sh = obterPlanilha_().getSheetByName(CONFIG.SHEET_LOGS);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = Math.min(qtd, sh.getLastRow() - 1);
  return sh.getRange(sh.getLastRow() - n + 1, 1, n, 5).getDisplayValues()
    .reverse()
    .map(r => ({ dataHora: r[0], operador: r[1], acao: r[2], status: r[3], detalhe: r[4] }));
}

/* =============================================================================
 * Gravação (substitui as linhas de dados do cadastro)
 * ===========================================================================*/
function salvarCadastro(tipo, linhas) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const def = CADASTROS[tipo];
    if (!def) throw new Error('Cadastro inválido: ' + tipo);
    const limpas = validarCadastro_(tipo, linhas || []);
    const sh = garantirAbaCadastro_(tipo);
    const nCols = def.colunas.length;

    const ultima = sh.getLastRow();
    if (ultima > 1) sh.getRange(2, 1, ultima - 1, nCols).clearContent();

    if (limpas.length) {
      const matriz = limpas.map(o => def.colunas.map(c => {
        const v = o[c.key];
        if (c.data) return v ? isoParaData_(v) : '';
        if (c.key === 'horas') return v === '' || v == null ? '' : converterNumero_(v);
        return v == null ? '' : String(v);
      }));
      sh.getRange(2, 1, matriz.length, nCols).setValues(matriz);
      def.colunas.forEach((c, j) => {
        if (c.data) sh.getRange(2, j + 1, matriz.length, 1).setNumberFormat('dd/MM/yyyy');
      });
    }
    registrarLog_('Cadastro ' + def.aba, 'SUCESSO', `${limpas.length} registro(s) salvos pelo painel.`);
    return { sucesso: true, mensagem: `${def.aba}: ${limpas.length} registro(s) salvos.`, linhas: listarCadastro(tipo).linhas };
  } catch (e) {
    return { sucesso: false, mensagem: e.message };
  } finally {
    lock.releaseLock();
  }
}

function validarCadastro_(tipo, linhas) {
  const t = s => String(s == null ? '' : s).trim();
  const erros = [];
  const out = [];
  const vistos = {};

  linhas.forEach((o, i) => {
    const n = i + 1;
    const l = {};
    Object.keys(o).forEach(k => l[k] = t(o[k]));
    if (Object.values(l).every(v => !v)) return;

    if (tipo === 'times') {
      if (!l.time) erros.push(`Linha ${n}: informe o nome do time.`);
      if (l.teamId && !/^\d+$/.test(l.teamId)) erros.push(`Linha ${n}: Team ID deve ser numérico.`);
      if (vistos[l.time]) erros.push(`Linha ${n}: time "${l.time}" repetido.`);
      vistos[l.time] = true;
    }
    if (tipo === 'colaboradores') {
      l.usuario = l.usuario.toLowerCase();
      if (!l.usuario) erros.push(`Linha ${n}: informe o nome de usuário.`);
      l.considerar = (l.considerar || 'SIM').toUpperCase() === 'NÃO' || l.considerar.toUpperCase() === 'NAO' ? 'NÃO' : 'SIM';
      if (vistos[l.usuario]) erros.push(`Linha ${n}: usuário "${l.usuario}" repetido.`);
      vistos[l.usuario] = true;
    }
    if (tipo === 'feriados') {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(l.data)) erros.push(`Linha ${n}: data inválida.`);
    }
    if (tipo === 'ausencias') {
      l.usuario = l.usuario.toLowerCase();
      if (!l.usuario) erros.push(`Linha ${n}: escolha o colaborador.`);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(l.inicio)) erros.push(`Linha ${n}: data de início inválida.`);
      if (!l.fim) l.fim = l.inicio;
      if (l.fim < l.inicio) erros.push(`Linha ${n}: data fim antes do início.`);
      if (l.horas && !(converterNumero_(l.horas) > 0 && converterNumero_(l.horas) <= 24)) erros.push(`Linha ${n}: horas por dia deve ser entre 0 e 24.`);
      if (!l.origem) l.origem = 'Manual';
      if (!l.tipo) l.tipo = 'Outro';
    }
    out.push(l);
  });

  if (erros.length) throw new Error(erros.slice(0, 6).join(' '));
  return out;
}

/* =============================================================================
 * Estrutura das abas
 * ===========================================================================*/
function garantirAbaCadastro_(tipo) {
  const def = CADASTROS[tipo];
  const ss = obterPlanilha_();
  let sh = ss.getSheetByName(def.aba);
  if (!sh) sh = ss.insertSheet(def.aba);

  const nCols = def.colunas.length;
  const atual = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, nCols).getDisplayValues()[0] : [];
  const precisa = def.colunas.some((c, j) => !String(atual[j] || '').trim());
  if (precisa) {
    const titulos = def.colunas.map((c, j) => String(atual[j] || '').trim() || c.titulo);
    sh.getRange(1, 1, 1, nCols).setValues([titulos])
      .setBackground('#0F5B78').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function isoParaData_(iso) {
  const p = String(iso).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}
