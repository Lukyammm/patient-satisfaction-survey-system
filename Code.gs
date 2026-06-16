const SHEET_NAME = 'MATRIZ';
const FIRST_DATA_ROW = 5;
const LAST_COL = 48; // A:AV
const SAIDAS_SHEET = 'SAIDAS';
const SAIDAS_FIRST_ROW = 2;
const CADASTROS_SHEET = 'CADASTROS';

// Cada bloco tem suas 4 manifestações (Sugestões/Reclamações/Comentários/Elogios)
const COL = {
  setor:1, pront:2, data:3, tipo:4, dn:5, idade:6, sexo:7,
  gentilezaAcolhimento:8, agilidade:9, clareza:10, satisfacao1:11, sugestoesAcol:12, reclamacoesAcol:13, comentariosAcol:14, elogiosAcol:15,
  gentilezaAssistencia:16, identificacao:17, intimidade:18, horarioDescanso:19, esclarecimento:20, cuidados:21, confianca:22, satisfacao2:23, sugestoesAssist:24, reclamacoesAssist:25, comentariosAssist:26, elogiosAssist:27,
  acesso:28, acomodacao:29, limpeza:30, enxoval:31, alimentacao:32, locomocao:33, satisfacao3:34, sugestoesServ:35, reclamacoesServ:36, comentariosServ:37, elogiosServ:38,
  nps:39, entrevistador:40,
  otimo:41, bom:42, regular:43, ruim:44, na:45, totalConsiderado:46, taxaSatisfacao:47, encaminhamentos:48
};

// Campos de manifestação por bloco (ordem de exibição: Sugestões, Reclamações, Comentários, Elogios)
const MANIF_KEYS = [
  'sugestoesAcol','reclamacoesAcol','comentariosAcol','elogiosAcol',
  'sugestoesAssist','reclamacoesAssist','comentariosAssist','elogiosAssist',
  'sugestoesServ','reclamacoesServ','comentariosServ','elogiosServ'
];

const RATING_KEYS = [
  'gentilezaAcolhimento','agilidade','clareza','gentilezaAssistencia','identificacao','intimidade','horarioDescanso','esclarecimento','cuidados','confianca','acesso','acomodacao','limpeza','enxoval','alimentacao','locomocao'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Pesquisa de Satisfação HUC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialData() {
  const sheet = getSheet_();
  ensureHeader_(sheet);
  const lastRow = Math.max(sheet.getLastRow(), FIRST_DATA_ROW - 1);
  const numRows = Math.max(0, lastRow - FIRST_DATA_ROW + 1);
  const values = numRows ? sheet.getRange(FIRST_DATA_ROW, 1, numRows, LAST_COL).getDisplayValues() : [];
  const records = values.map((row, i) => rowToObject_(row, FIRST_DATA_ROW + i))
    .filter(r => hasContent_(r));

  const cadastros = getCadastros();
  const cadSetores = cadastros.filter(c => c.tipo === 'SETOR').map(c => c.nome);
  const cadEnts = cadastros.filter(c => c.tipo === 'ENTREVISTADOR').map(c => c.nome);

  return {
    records,
    saidas: getSaidas(),
    cadastros,
    lookups: {
      setores: unique_(records.map(r => r.setor).concat(cadSetores)),
      tipos: unique_(records.map(r => r.tipo)),
      sexos: unique_(records.map(r => r.sexo)),
      entrevistadores: unique_(records.map(r => r.entrevistador).concat(cadEnts))
    }
  };
}

function getCadastros() {
  const sheet = getCadastrosSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues()
    .map((row, i) => ({ rowNumber: i + 2, tipo: String(row[0] || '').trim().toUpperCase(), nome: String(row[1] || '').trim() }))
    .filter(c => c.nome);
}

function saveCadastro(payload) {
  const tipo = String(payload.tipo || '').trim().toUpperCase();
  const nome = String(payload.nome || '').trim();
  if (tipo !== 'SETOR' && tipo !== 'ENTREVISTADOR') throw new Error('Tipo inválido.');
  if (!nome) throw new Error('Informe o nome.');
  const exists = getCadastros().some(c => c.tipo === tipo && c.nome.toUpperCase() === nome.toUpperCase());
  if (exists) throw new Error(tipo === 'SETOR' ? 'Setor já cadastrado.' : 'Entrevistador já cadastrado.');
  getCadastrosSheet_().appendRow([tipo, nome]);
  return { ok: true, message: tipo === 'SETOR' ? 'Setor cadastrado.' : 'Entrevistador cadastrado.' };
}

function deleteCadastro(rowNumber) {
  rowNumber = Number(rowNumber);
  if (!rowNumber || rowNumber < 2) throw new Error('Linha inválida.');
  getCadastrosSheet_().deleteRow(rowNumber);
  return { ok: true, message: 'Cadastro removido.' };
}

function getCadastrosSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CADASTROS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CADASTROS_SHEET);
    sheet.getRange(1, 1, 1, 2).setValues([['TIPO', 'NOME']]).setFontWeight('bold').setBackground('#eaf2ff');
    sheet.setFrozenRows(1);
    seedCadastros_(sheet);
  }
  return sheet;
}

// Popula a aba CADASTROS com os setores e entrevistadores já presentes na MATRIZ
function seedCadastros_(sheet) {
  const m = getSheet_();
  const lastRow = m.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return;
  const n = lastRow - FIRST_DATA_ROW + 1;
  const setores = unique_(m.getRange(FIRST_DATA_ROW, COL.setor, n, 1).getDisplayValues().flat());
  const ents = unique_(m.getRange(FIRST_DATA_ROW, COL.entrevistador, n, 1).getDisplayValues().flat());
  const rows = setores.map(s => ['SETOR', s]).concat(ents.map(e => ['ENTREVISTADOR', e]));
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function getSaidas() {
  const sheet = getSaidasSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < SAIDAS_FIRST_ROW) return [];
  return sheet.getRange(SAIDAS_FIRST_ROW, 1, lastRow - SAIDAS_FIRST_ROW + 1, 3).getDisplayValues()
    .map((row, i) => ({ rowNumber: SAIDAS_FIRST_ROW + i, mes: String(row[0] || '').trim(), setor: String(row[1] || '').trim(), saidas: Number(row[2]) || 0 }))
    .filter(s => s.mes || s.setor);
}

function saveSaida(payload) {
  const sheet = getSaidasSheet_();
  const isEdit = !!Number(payload.rowNumber);
  const rowNumber = isEdit ? Number(payload.rowNumber) : sheet.getLastRow() + 1;
  const mes = payload.mes || '', setor = payload.setor || '', saidas = Number(payload.saidas) || 0;
  sheet.getRange(rowNumber, 1, 1, 3).setValues([[mes, setor, saidas]]);
  return { ok: true, rowNumber, saida: { rowNumber, mes, setor, saidas }, message: isEdit ? 'Saídas atualizadas.' : 'Saídas registradas.' };
}

function deleteSaida(rowNumber) {
  rowNumber = Number(rowNumber);
  if (!rowNumber || rowNumber < SAIDAS_FIRST_ROW) throw new Error('Linha inválida.');
  getSaidasSheet_().deleteRow(rowNumber);
  return { ok: true, message: 'Registro de saídas excluído.' };
}

function getSaidasSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SAIDAS_SHEET);
  if (!sheet) sheet = ss.insertSheet(SAIDAS_SHEET);
  if (sheet.getRange('A1').getValue() !== 'MÊS/ANO') {
    sheet.getRange(1, 1, 1, 3).setValues([['MÊS/ANO', 'SETOR', 'SAÍDAS']]).setFontWeight('bold').setBackground('#eaf2ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Critérios de cada bloco do formulário — usados para as sub-taxas K/T/AA
const BLOCO_ACOLHIMENTO = ['gentilezaAcolhimento','agilidade','clareza'];
const BLOCO_ASSISTENCIA = ['gentilezaAssistencia','identificacao','intimidade','horarioDescanso','esclarecimento','cuidados','confianca'];
const BLOCO_SERVICOS = ['acesso','acomodacao','limpeza','enxoval','alimentacao','locomocao'];

function saveRecord(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Sistema ocupado, tente novamente em instantes.');
  try {
  const sheet = getSheet_();
  ensureHeader_(sheet);
  const isEdit = !!Number(payload.rowNumber);
  const rowNumber = isEdit ? Number(payload.rowNumber) : nextDataRow_(sheet);
  const row = computeRow_(payload);

  sheet.getRange(rowNumber, 1, 1, LAST_COL).setValues([row]);
  sheet.getRangeList(pctCols_().map(c => c + rowNumber)).setNumberFormat('0.00%');

  const saved = rowToObject_(row.map(v => v === '' ? '' : String(v)), rowNumber);
  // row contém Date objects para campos de data; String(Date) produz formato ilegível.
  // Restaura do payload original que já está em YYYY-MM-DD.
  saved.data = normalizeDate_(payload.data);
  saved.dn   = normalizeDate_(payload.dn);
  return { ok: true, rowNumber, record: saved, message: isEdit ? 'Registro atualizado.' : 'Registro salvo.' };
  } finally {
    lock.releaseLock();
  }
}

// Importação em lote (CSV digitalizado): reaproveita saveRecord para cada linha
function importRecords(payloads) {
  if (!Array.isArray(payloads) || !payloads.length) throw new Error('Nada para importar.');
  if (payloads.length > 500) throw new Error('Máximo de 500 pesquisas por importação.');
  let saved = 0;
  const errors = [];
  payloads.forEach((p, i) => {
    try {
      p.rowNumber = '';
      saveRecord(p);
      saved++;
    } catch (e) {
      errors.push('Registro ' + (i + 1) + ': ' + (e && e.message ? e.message : e));
    }
  });
  return {
    ok: true, saved, errors,
    message: saved + ' pesquisa(s) importada(s)' + (errors.length ? ' · ' + errors.length + ' com erro' : '') + '.'
  };
}

function deleteRecord(rowNumber) {
  rowNumber = Number(rowNumber);
  if (!rowNumber || rowNumber < FIRST_DATA_ROW) throw new Error('Linha inválida.');
  getSheet_().deleteRow(rowNumber);
  return { ok: true, message: 'Registro excluído.' };
}

function exportCsv() {
  const data = getInitialData().records;
  const headers = ['Data','Setor','Prontuario','Tipo','DN','Idade','Sexo','Satisfacao %','NPS','Classificacao NPS',
    'Acolhimento Sugestoes','Acolhimento Reclamacoes','Acolhimento Comentarios','Acolhimento Elogios',
    'Assistencia Sugestoes','Assistencia Reclamacoes','Assistencia Comentarios','Assistencia Elogios',
    'Servicos Sugestoes','Servicos Reclamacoes','Servicos Comentarios','Servicos Elogios',
    'Entrevistador','Encaminhamentos'];
  const lines = [headers].concat(data.map(r => [
    r.data, r.setor, r.pront, r.tipo, r.dn, r.idade, r.sexo, r.taxaSatisfacao, r.nps, classifyNps_(r.nps),
    r.sugestoesAcol, r.reclamacoesAcol, r.comentariosAcol, r.elogiosAcol,
    r.sugestoesAssist, r.reclamacoesAssist, r.comentariosAssist, r.elogiosAssist,
    r.sugestoesServ, r.reclamacoesServ, r.comentariosServ, r.elogiosServ,
    r.entrevistador, r.encaminhamentos
  ])).map(row => row.map(csvCell_).join(';'));
  return lines.join('\n');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function ensureHeader_(sheet) {
  const a2 = sheet.getRange('A2').getValue();
  const l3 = sheet.getRange('L3').getValue();
  if (a2 === 'SETOR' && l3 === 'SUGESTÕES') return;                    // layout novo já presente
  if (a2 === 'SETOR' && sheet.getLastRow() >= FIRST_DATA_ROW) return;  // layout legado COM dados — rodar migrateToBlockManifestations()
  writeHeader_(sheet);
}

// Escreve as linhas de cabeçalho (1-3) montando a partir do COL — evita contagem manual de colunas
function writeHeader_(sheet) {
  const row2 = Array(LAST_COL).fill(''), row3 = Array(LAST_COL).fill('');
  const p2 = (k, v) => row2[COL[k] - 1] = v, p3 = (k, v) => row3[COL[k] - 1] = v;
  // Linha 2 — grupos
  p2('setor','SETOR'); p2('pront','PRONT'); p2('data','DATA'); p2('tipo','TIPO'); p2('dn','DN'); p2('idade','IDADE'); p2('sexo','SEXO');
  p2('gentilezaAcolhimento','1. ACOLHIMENTO');
  p2('gentilezaAssistencia','2. ASSISTÊNCIA');
  p2('acesso','3. SERVIÇOS PRESTADOS');
  p2('nps','NPS'); p2('entrevistador','ENTREVISTADOR'); p2('otimo','TABULAÇÃO DOS DADOS'); p2('encaminhamentos','ENCAMINHAMENTOS');
  // Linha 3 — colunas detalhadas
  p3('gentilezaAcolhimento','GENTILEZA E ATENÇÃO'); p3('agilidade','AGILIDADE'); p3('clareza','CLAREZA'); p3('satisfacao1','SATISFAÇÃO I');
  p3('sugestoesAcol','SUGESTÕES'); p3('reclamacoesAcol','RECLAMAÇÕES'); p3('comentariosAcol','COMENTÁRIOS'); p3('elogiosAcol','ELOGIOS');
  p3('gentilezaAssistencia','GENTILEZA E ATENÇÃO'); p3('identificacao','IDENTIFICAÇÃO'); p3('intimidade','INTIMIDADE/PRIVACIDADE'); p3('horarioDescanso','HORÁRIO DE DESCANSO'); p3('esclarecimento','ESCLARECIMENTO'); p3('cuidados','CUIDADOS PRESTADOS'); p3('confianca','CONFIANÇA E SEGURANÇA'); p3('satisfacao2','SATISFAÇÃO II');
  p3('sugestoesAssist','SUGESTÕES'); p3('reclamacoesAssist','RECLAMAÇÕES'); p3('comentariosAssist','COMENTÁRIOS'); p3('elogiosAssist','ELOGIOS');
  p3('acesso','ACESSO'); p3('acomodacao','ACOMODAÇÃO'); p3('limpeza','LIMPEZA'); p3('enxoval','ENXOVAL'); p3('alimentacao','ALIMENTAÇÃO'); p3('locomocao','LOCOMOÇÃO'); p3('satisfacao3','SATISFAÇÃO III');
  p3('sugestoesServ','SUGESTÕES'); p3('reclamacoesServ','RECLAMAÇÕES'); p3('comentariosServ','COMENTÁRIOS'); p3('elogiosServ','ELOGIOS');
  p3('otimo','ÓTIMO'); p3('bom','BOM'); p3('regular','REGULAR'); p3('ruim','RUIM'); p3('na','N/A'); p3('totalConsiderado','TOTAL CONSIDERADO'); p3('taxaSatisfacao','TAXA DE SATISFAÇÃO');
  sheet.getRange('B1').setValue('PESQUISA DE SATISFAÇÃO');
  sheet.getRange(2, 1, 1, LAST_COL).setValues([row2]);
  sheet.getRange(3, 1, 1, LAST_COL).setValues([row3]);
  sheet.getRange(1, 1, 3, LAST_COL).setFontWeight('bold').setBackground('#eaf2ff');
  sheet.setFrozenRows(3);
}

// Letra(s) da coluna A1 a partir do índice (1 -> A, 27 -> AA)
function colA1_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }

// Colunas formatadas como porcentagem (sub-taxas por bloco + taxa geral)
function pctCols_() { return [COL.satisfacao1, COL.satisfacao2, COL.satisfacao3, COL.taxaSatisfacao].map(colA1_); }

// Monta a linha completa da planilha (campos + agregados + sub-taxas + idade) a partir do payload
function computeRow_(payload) {
  const row = objectToRow_(payload);
  const cnt = v => RATING_KEYS.reduce((n, k) => n + (String(payload[k] || '').toUpperCase() === v ? 1 : 0), 0);
  const otimo = cnt('ÓTIMO'), bom = cnt('BOM'), reg = cnt('REGULAR'), ruim = cnt('RUIM'), na = cnt('N/A');
  const total = otimo + bom + reg + ruim;
  row[COL.otimo - 1] = otimo;
  row[COL.bom - 1] = bom;
  row[COL.regular - 1] = reg;
  row[COL.ruim - 1] = ruim;
  row[COL.na - 1] = na;
  row[COL.totalConsiderado - 1] = total;
  row[COL.taxaSatisfacao - 1] = total ? (otimo + bom) / total : 0;
  const blockSat = keys => {
    let ob = 0, valid = 0;
    keys.forEach(k => {
      const v = String(payload[k] || '').toUpperCase();
      if (v === 'ÓTIMO' || v === 'BOM') { ob++; valid++; }
      else if (v === 'REGULAR' || v === 'RUIM') valid++;
    });
    return valid ? ob / valid : 0;
  };
  row[COL.satisfacao1 - 1] = blockSat(BLOCO_ACOLHIMENTO);
  row[COL.satisfacao2 - 1] = blockSat(BLOCO_ASSISTENCIA);
  row[COL.satisfacao3 - 1] = blockSat(BLOCO_SERVICOS);
  const dnD = parseDateOrText_(payload.dn);
  const dtD = parseDateOrText_(payload.data);
  if (dnD instanceof Date && dtD instanceof Date) {
    const age = Math.floor((dtD - dnD) / 31557600000);
    if (age >= 0 && age <= 120) row[COL.idade - 1] = age;
  }
  return row;
}

function rowToObject_(row, rowNumber) {
  const get = key => row[COL[key]-1] || '';
  const obj = { rowNumber };
  Object.keys(COL).forEach(k => obj[k] = get(k));
  obj.data = normalizeDate_(obj.data);
  obj.dn = normalizeDate_(obj.dn);
  obj.taxaSatisfacao = numberFromPercent_(obj.taxaSatisfacao);
  obj.otimo = Number(obj.otimo) || countRating_(row, 'ÓTIMO');
  obj.bom = Number(obj.bom) || countRating_(row, 'BOM');
  obj.regular = Number(obj.regular) || countRating_(row, 'REGULAR');
  obj.ruim = Number(obj.ruim) || countRating_(row, 'RUIM');
  obj.na = Number(obj.na) || countRating_(row, 'N/A');
  if (!obj.taxaSatisfacao) obj.taxaSatisfacao = satisfaction_(row) * 100;
  return obj;
}

function objectToRow_(p) {
  const row = Array(LAST_COL).fill('');
  Object.keys(COL).forEach(k => {
    if (p[k] !== undefined) row[COL[k]-1] = p[k];
  });
  row[COL.data-1] = parseDateOrText_(p.data);
  row[COL.dn-1] = parseDateOrText_(p.dn);
  // Normalização: evita variações livres (MASC/masculino/fem...) na planilha
  row[COL.sexo-1] = normSexo_(p.sexo);
  row[COL.setor-1] = String(p.setor || '').trim().toUpperCase();
  row[COL.tipo-1] = String(p.tipo || '').trim().toUpperCase();
  row[COL.entrevistador-1] = String(p.entrevistador || '').trim().toUpperCase();
  return row;
}

function normSexo_(v) {
  const u = String(v || '').trim().toUpperCase();
  if (['M', 'MASC', 'MASCULINO', 'HOMEM'].includes(u)) return 'M';
  if (['F', 'FEM', 'FEMININO', 'MULHER'].includes(u)) return 'F';
  if (u === 'OUTRO') return 'Outro';
  return String(v || '').trim();
}

function nextDataRow_(sheet) {
  const last = Math.max(sheet.getLastRow(), FIRST_DATA_ROW - 1);
  const vals = last >= FIRST_DATA_ROW ? sheet.getRange(FIRST_DATA_ROW, 1, last - FIRST_DATA_ROW + 1, 1).getValues().flat() : [];
  const idx = vals.findIndex(v => !v);
  return idx >= 0 ? FIRST_DATA_ROW + idx : last + 1;
}

function hasContent_(r) {
  return ['setor','pront','data','tipo','sexo','nps'].concat(MANIF_KEYS).some(k => String(r[k] || '').trim());
}

function unique_(arr) {
  return [...new Set(arr.map(v => String(v || '').trim()).filter(Boolean))].sort((a,b) => a.localeCompare(b, 'pt-BR'));
}

function countRating_(row, rating) {
  return RATING_KEYS.reduce((n, key) => n + (String(row[COL[key]-1]).toUpperCase() === rating ? 1 : 0), 0);
}

function satisfaction_(row) {
  const ot = countRating_(row, 'ÓTIMO');
  const bom = countRating_(row, 'BOM');
  const reg = countRating_(row, 'REGULAR');
  const ruim = countRating_(row, 'RUIM');
  const total = ot + bom + reg + ruim;
  return total ? (ot + bom) / total : 0;
}

function classifyNps_(nps) {
  const n = Number(nps);
  if (isNaN(n)) return '';
  if (n >= 9) return 'Promotor';
  if (n >= 7) return 'Neutro';
  return 'Detrator';
}

function csvCell_(v) {
  return '"' + String(v ?? '').replace(/"/g, '""') + '"';
}

function normalizeDate_(v) {
  if (!v) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  const m = String(v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
  return v;
}

function parseDateOrText_(s) {
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const [y,m,d] = s.split('-').map(Number);
    return new Date(y, m - 1, d);
  }
  return s;
}

function numberFromPercent_(v) {
  if (v === '' || v == null) return 0;
  const s = String(v).replace('%','').replace(',','.');
  const n = Number(s);
  if (isNaN(n)) return 0;
  return n <= 1 ? n * 100 : n;
}

/**
 * Preenche as colunas K (SATISFAÇÃO I), T (SATISFAÇÃO II) e AB (SATISFAÇÃO III)
 * para todos os registros históricos que estejam com esses valores vazios.
 * Execute uma única vez pelo editor de script: Funções > backfillSatisfacaoBlocks
 */
function backfillSatisfacaoBlocks() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) { Logger.log('Nenhum registro encontrado.'); return; }

  const numRows = lastRow - FIRST_DATA_ROW + 1;
  const range = sheet.getRange(FIRST_DATA_ROW, 1, numRows, LAST_COL);
  const values = range.getValues();

  const blockSatFromRow = (row, keys) => {
    let ob = 0, valid = 0;
    keys.forEach(k => {
      const v = String(row[COL[k] - 1] || '').toUpperCase();
      if (v === 'ÓTIMO' || v === 'BOM') { ob++; valid++; }
      else if (v === 'REGULAR' || v === 'RUIM') valid++;
    });
    return valid ? ob / valid : null; // null = sem dados válidos, não sobrescreve
  };

  let updated = 0;
  const rangeListAddrs = [];

  values.forEach((row, i) => {
    const rowNum = FIRST_DATA_ROW + i;
    const isEmpty = v => v === '' || v === null || v === undefined;

    const s1 = row[COL.satisfacao1 - 1];
    const s2 = row[COL.satisfacao2 - 1];
    const s3 = row[COL.satisfacao3 - 1];

    // Só processa linhas com conteúdo e pelo menos uma sub-taxa vazia
    if (!hasContent_(rowToObject_(row.map(v => String(v || '')), rowNum))) return;
    if (!isEmpty(s1) && !isEmpty(s2) && !isEmpty(s3)) return;

    let changed = false;
    if (isEmpty(s1)) {
      const v = blockSatFromRow(row, BLOCO_ACOLHIMENTO);
      if (v !== null) { row[COL.satisfacao1 - 1] = v; rangeListAddrs.push(colA1_(COL.satisfacao1) + rowNum); changed = true; }
    }
    if (isEmpty(s2)) {
      const v = blockSatFromRow(row, BLOCO_ASSISTENCIA);
      if (v !== null) { row[COL.satisfacao2 - 1] = v; rangeListAddrs.push(colA1_(COL.satisfacao2) + rowNum); changed = true; }
    }
    if (isEmpty(s3)) {
      const v = blockSatFromRow(row, BLOCO_SERVICOS);
      if (v !== null) { row[COL.satisfacao3 - 1] = v; rangeListAddrs.push(colA1_(COL.satisfacao3) + rowNum); changed = true; }
    }

    if (changed) {
      sheet.getRange(rowNum, 1, 1, LAST_COL).setValues([row]);
      updated++;
    }
  });

  if (rangeListAddrs.length) {
    sheet.getRangeList(rangeListAddrs).setNumberFormat('0.00%');
  }

  const msg = `backfillSatisfacaoBlocks: ${updated} linha(s) atualizada(s).`;
  Logger.log(msg);
  return msg;
}

/**
 * MIGRAÇÃO ÚNICA — converte o layout legado para o novo modelo com
 * Sugestões/Reclamações/Comentários/Elogios em CADA bloco.
 *
 * Regras:
 *   - Textos livres de cada bloco -> "Comentários" do mesmo bloco
 *   - Manifestações globais antigas (Sugestões/Reclamações/Comentários/Elogios) -> bloco 1 (Acolhimento)
 *   - No Acolhimento, o texto livre antigo e o comentário global são concatenados (sem perder nada)
 *
 * Segurança: faz uma cópia de backup da aba ANTES de reescrever, e aborta se já migrado.
 * Execute UMA única vez pelo editor de script: Funções > migrateToBlockManifestations
 */
function migrateToBlockManifestations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { const m = 'Aba MATRIZ não encontrada.'; Logger.log(m); return m; }
  const LEGACY_BLOCK_TEXT_HEADER = 'OBSERVAÇÕES';

  // Já migrado? (col 12 / L na linha 3 já é "SUGESTÕES" no layout novo)
  if (sheet.getRange('A2').getValue() === 'SETOR' && sheet.getRange('L3').getValue() === 'SUGESTÕES') {
    const m = 'Migração já aplicada anteriormente — nada a fazer.'; Logger.log(m); return m;
  }
  // Confere o layout antigo esperado antes de mexer
  if (sheet.getRange('A2').getValue() !== 'SETOR' || sheet.getRange('L3').getValue() !== LEGACY_BLOCK_TEXT_HEADER) {
    const m = 'Cabeçalho não corresponde ao layout legado esperado. Abortando por segurança.';
    Logger.log(m); return m;
  }

  // Mapa de colunas do layout ANTIGO (snapshot fixo para leitura)
  const OLD_COL = {
    setor:1, pront:2, data:3, tipo:4, dn:5, idade:6, sexo:7,
    gentilezaAcolhimento:8, agilidade:9, clareza:10, satisfacao1:11, textoAcolhimento:12,
    gentilezaAssistencia:13, identificacao:14, intimidade:15, horarioDescanso:16, esclarecimento:17, cuidados:18, confianca:19, satisfacao2:20, textoAssistencia:21,
    acesso:22, acomodacao:23, limpeza:24, enxoval:25, alimentacao:26, locomocao:27, satisfacao3:28, textoServicos:29,
    sugestoes:30, reclamacoes:31, comentarios:32, elogios:33, nps:34, entrevistador:35,
    otimo:36, bom:37, regular:38, ruim:39, na:40, totalConsiderado:41, taxaSatisfacao:42, encaminhamentos:43
  };
  const OLD_LAST_COL = 43;

  const lastRow = sheet.getLastRow();
  const numRows = Math.max(0, lastRow - FIRST_DATA_ROW + 1);
  const oldValues = numRows ? sheet.getRange(FIRST_DATA_ROW, 1, numRows, OLD_LAST_COL).getDisplayValues() : [];

  // Backup ANTES de qualquer escrita
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Fortaleza', 'yyyyMMdd_HHmmss');
  const backup = sheet.copyTo(ss).setName('MATRIZ_BACKUP_' + stamp);

  const get = (row, key) => row[OLD_COL[key] - 1] || '';
  const merge = (a, b) => [a, b].map(s => String(s || '').trim()).filter(Boolean).join('\n');

  const newRows = [];
  oldValues.forEach(row => {
    const anyManif = ['sugestoes','reclamacoes','comentarios','elogios','textoAcolhimento','textoAssistencia','textoServicos']
      .some(k => String(get(row, k) || '').trim());
    const anyBase = ['setor','pront','data','tipo','sexo','nps'].some(k => String(get(row, k) || '').trim());
    const anyRating = RATING_KEYS.some(k => String(get(row, k) || '').trim());
    if (!anyManif && !anyBase && !anyRating) return; // linha vazia

    const p = {};
    ['setor','pront','tipo','sexo','idade','nps','entrevistador','encaminhamentos'].forEach(k => p[k] = get(row, k));
    p.data = normalizeDate_(get(row, 'data'));
    p.dn = normalizeDate_(get(row, 'dn'));
    RATING_KEYS.forEach(k => p[k] = get(row, k));
    p.sugestoesAcol   = get(row, 'sugestoes');
    p.reclamacoesAcol = get(row, 'reclamacoes');
    p.comentariosAcol = merge(get(row, 'textoAcolhimento'), get(row, 'comentarios'));
    p.elogiosAcol     = get(row, 'elogios');
    p.comentariosAssist = get(row, 'textoAssistencia');
    p.comentariosServ   = get(row, 'textoServicos');
    newRows.push(computeRow_(p));
  });

  // Reescreve a aba no layout novo
  sheet.clear();
  writeHeader_(sheet);
  if (newRows.length) {
    sheet.getRange(FIRST_DATA_ROW, 1, newRows.length, LAST_COL).setValues(newRows);
    const addrs = [];
    const cols = pctCols_();
    for (let i = 0; i < newRows.length; i++) { const r = FIRST_DATA_ROW + i; cols.forEach(c => addrs.push(c + r)); }
    if (addrs.length) sheet.getRangeList(addrs).setNumberFormat('0.00%');
  }

  const msg = `Migração concluída: ${newRows.length} registro(s) convertido(s). Backup salvo em "${backup.getName()}".`;
  Logger.log(msg);
  return msg;
}
