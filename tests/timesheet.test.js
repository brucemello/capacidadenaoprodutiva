// Ausências manuais lançadas no painel do Timesheet entram na sincronização do Capacidade.
const fs = require('fs'), vm = require('vm'), path = require('path'), assert = require('assert');
const pad = n => String(n).padStart(2, '0');
const br = d => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
const aba = (nome, dados) => ({ getName: () => nome, getLastRow: () => dados.length, getLastColumn: () => 5,
  getRange: (r, c, nr, nc) => ({ getDisplayValues: () => dados.slice(r - 1, r - 1 + nr).map(l => l.slice(c - 1, c - 1 + nc).map(v => v instanceof Date ? br(v) : String(v == null ? '' : v))) }) });
const timesheet = { getSheetByName: n => n === 'Ausencias' ? aba('Ausencias', [
  ['Identificador', 'Início', 'Fim', 'Horas', 'Motivo'],
  ['carla.latorre@totvs.com.br', new Date(2026, 9, 26), new Date(2026, 9, 30), '', 'Licença nojo'],
  ['Keilla', '05/11/2026', '05/11/2026', '3', 'Consulta'],
  ['ninguem@totvs.com.br', '05/11/2026', '05/11/2026', '', 'x']
]) : null };
const ctx = { console, Logger: { log() {} },
  Utilities: { formatDate: (d, tz, f) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` },
  SpreadsheetApp: { openByUrl: u => { if (/TS123/.test(u)) return timesheet; throw new Error('sem acesso'); }, openById: id => { if (id === 'TS123') return timesheet; throw new Error('sem acesso'); } } };
vm.createContext(ctx);
const SRC = path.join(__dirname, '..', 'src');
fs.readdirSync(SRC).filter(f => f.endsWith('.js')).forEach(f => vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f }));
ctx.__colabs = [
  { usuario: 'carla.latorre', nome: 'Carla Bianca Santos Latorre', considerar: 'SIM', time: 'TSS' },
  { usuario: 'keilla.freitas', nome: 'Keilla Silva de Freitas', considerar: 'SIM', time: 'TSS' },
  { usuario: 'carlos.silva', nome: 'Carlos Eduardo Nonato da Silva', considerar: 'SIM', time: 'TSS' }
];
const r = vm.runInContext('lerAusenciasTimesheet_("TSS", "https://docs.google.com/spreadsheets/d/TS123/edit", criarMatcherColaborador_(__colabs, {}))', ctx);
const carla = r.linhas.find(l => l.usuario === 'carla.latorre');
assert.ok(carla, 'Carla identificada pelo e-mail');
assert.strictEqual(carla.inicio, '2026-10-26'); assert.strictEqual(carla.fim, '2026-10-30');
assert.strictEqual(carla.tipo, 'Licença'); assert.strictEqual(carla.horas, ''); assert.strictEqual(carla.origem, 'Timesheet');
const keilla = r.linhas.find(l => l.usuario === 'keilla.freitas');
assert.strictEqual(keilla.horas, '3'); assert.strictEqual(keilla.tipo, 'Consulta/Exame');
assert.strictEqual(r.linhas.length, 2);
assert.strictEqual(JSON.stringify(r.naoVinculados), JSON.stringify(['ninguem@totvs.com.br']));
// coluna com texto que não é planilha é ignorada
ctx.listarCadastro = () => ({ linhas: [{ time: 'TSS', teamId: '1', agenda: '', planilhaFerias: '', timesheet: 'observação qualquer' }, { time: 'X', timesheet: 'https://docs.google.com/spreadsheets/d/TS123/edit' }] });
const f = vm.runInContext('obterFontesAusenciaPorTime_()', ctx);
assert.strictEqual(f[0].timesheet, ''); assert.ok(f[1].timesheet);
console.log('TODOS OS TESTES PASSARAM');
