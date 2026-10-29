/*******************************************************************************
 * FERIADOS AUTOMÁTICOS
 * Calcula os feriados de cada ano (fixos + móveis pela Páscoa) e inclui na aba
 * FERIADOS com origem "Automático". Lançamentos manuais (municipais, emendas,
 * pontos facultativos da empresa) nunca são alterados nem removidos.
 *******************************************************************************/

const ORIGEM_FERIADO_AUTO = 'Automático';
const ORIGEM_FERIADO_MANUAL = 'Manual';

/** Feriados de data fixa [mês, dia, descrição]. Inclui o estadual de SP (09/07). */
const FERIADOS_FIXOS = [
  [1, 1, 'Confraternização Universal'],
  [4, 21, 'Tiradentes'],
  [5, 1, 'Dia do Trabalho'],
  [7, 9, 'Revolução Constitucionalista (SP)'],
  [9, 7, 'Independência do Brasil'],
  [10, 12, 'Nossa Senhora Aparecida'],
  [11, 2, 'Finados'],
  [11, 15, 'Proclamação da República'],
  [11, 20, 'Dia Nacional de Zumbi e da Consciência Negra'],
  [12, 25, 'Natal']
];

/** Feriados móveis: dias em relação ao Domingo de Páscoa. */
const FERIADOS_MOVEIS = [
  [-48, 'Carnaval (segunda-feira)'],
  [-47, 'Carnaval (terça-feira)'],
  [-2, 'Sexta-feira Santa'],
  [60, 'Corpus Christi']
];

/** Domingo de Páscoa (algoritmo de Meeus/Jones/Butcher, calendário gregoriano). */
function calcularPascoa_(ano) {
  const a = ano % 19, b = Math.floor(ano / 100), c = ano % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mes = Math.floor((h + l - 7 * m + 114) / 31);
  const dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(ano, mes - 1, dia);
}

/** Lista [{ data: 'yyyy-MM-dd', descricao }] do ano. */
function feriadosDoAno_(ano) {
  const lista = FERIADOS_FIXOS.map(f => ({ data: isoData_(new Date(ano, f[0] - 1, f[1])), descricao: f[2] }));
  const pascoa = calcularPascoa_(ano);
  FERIADOS_MOVEIS.forEach(f => {
    const d = new Date(pascoa.getFullYear(), pascoa.getMonth(), pascoa.getDate() + f[0]);
    lista.push({ data: isoData_(d), descricao: f[1] });
  });
  return lista.sort((a, b) => a.data.localeCompare(b.data));
}

/**
 * Garante os feriados automáticos dos anos informados na aba FERIADOS.
 * Datas já cadastradas (manuais ou automáticas) são mantidas como estão.
 * @return {number} quantidade de feriados incluídos
 */
function garantirFeriadosAnos_(anos) {
  const atuais = listarCadastro('feriados').linhas;
  const datas = new Set(atuais.map(f => f.data));
  const novos = [];
  [...new Set(anos)].forEach(ano => {
    feriadosDoAno_(ano).forEach(f => {
      if (datas.has(f.data)) return;
      datas.add(f.data);
      novos.push({ data: f.data, descricao: f.descricao, origem: ORIGEM_FERIADO_AUTO });
    });
  });
  if (!novos.length) return 0;

  const final = atuais.map(f => ({ ...f, origem: f.origem || ORIGEM_FERIADO_MANUAL }))
    .concat(novos)
    .sort((a, b) => a.data.localeCompare(b.data));
  gravarFeriados_(final);
  registrarLog_('Feriados automáticos', 'SUCESSO', `${novos.length} feriado(s) incluído(s): ${novos.map(n => n.data.split('-').reverse().join('/') + ' ' + n.descricao).join(', ')}`);
  return novos.length;
}

function gravarFeriados_(linhas) {
  const def = CADASTROS.feriados;
  const sh = garantirAbaCadastro_('feriados');
  const nCols = def.colunas.length;
  const ultima = sh.getLastRow();
  if (ultima > 1) sh.getRange(2, 1, ultima - 1, nCols).clearContent();
  if (!linhas.length) return;
  const matriz = linhas.map(o => def.colunas.map(c => c.data ? (o[c.key] ? isoParaData_(o[c.key]) : '') : (o[c.key] || '')));
  sh.getRange(2, 1, matriz.length, nCols).setValues(matriz);
  sh.getRange(2, 1, matriz.length, 1).setNumberFormat('dd/MM/yyyy');
}

/** Painel: gera os feriados do ano atual e do próximo. */
function gerarFeriadosAutomaticosPeloPainel() {
  try {
    const ano = new Date().getFullYear();
    const n = garantirFeriadosAnos_([ano, ano + 1]);
    return {
      sucesso: true,
      mensagem: n ? `${n} feriado(s) de ${ano} e ${ano + 1} incluídos.` : `Os feriados de ${ano} e ${ano + 1} já estavam cadastrados.`,
      linhas: listarCadastro('feriados').linhas
    };
  } catch (e) {
    return { sucesso: false, mensagem: e.message };
  }
}
