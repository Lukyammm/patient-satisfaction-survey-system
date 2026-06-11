const SHEET_NAME = 'MATRIZ';
const FIRST_DATA_ROW = 5;
const LAST_COL = 43; // A:AQ
const SAIDAS_SHEET = 'SAIDAS';
const SAIDAS_FIRST_ROW = 2;
const CADASTROS_SHEET = 'CADASTROS';

const COL = {
  setor:1, pront:2, data:3, tipo:4, dn:5, idade:6, sexo:7,
  gentilezaAcolhimento:8, agilidade:9, clareza:10, satisfacao1:11, obsAcolhimento:12,
  gentilezaAssistencia:13, identificacao:14, intimidade:15, horarioDescanso:16, esclarecimento:17, cuidados:18, confianca:19, satisfacao2:20, obsAssistencia:21,
  acesso:22, acomodacao:23, limpeza:24, enxoval:25, alimentacao:26, locomocao:27, satisfacao3:28, obsServicos:29,
  sugestoes:30, reclamacoes:31, comentarios:32, elogios:33, nps:34, entrevistador:35,
  otimo:36, bom:37, regular:38, ruim:39, na:40, totalConsiderado:41, taxaSatisfacao:42, encaminhamentos:43
};

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
  const sheet = getSheet_();
  ensureHeader_(sheet);
  const isEdit = !!Number(payload.rowNumber);
  const rowNumber = isEdit ? Number(payload.rowNumber) : nextDataRow_(sheet);
  const row = objectToRow_(payload);

  // Compute aggregate counts from payload (avoids 13+ separate formula API calls)
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

  // Sub-taxas por bloco (colunas K, T, AA — SATISFAÇÃO I/II/III)
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

  // Calcula idade a partir da data de nascimento e a data da pesquisa
  const dnD = parseDateOrText_(payload.dn);
  const dtD = parseDateOrText_(payload.data);
  if (dnD instanceof Date && dtD instanceof Date) {
    const age = Math.floor((dtD - dnD) / 31557600000);
    if (age >= 0 && age <= 120) row[COL.idade - 1] = age;
  }

  sheet.getRange(rowNumber, 1, 1, LAST_COL).setValues([row]);
  sheet.getRangeList([`K${rowNumber}`, `T${rowNumber}`, `AB${rowNumber}`, `AP${rowNumber}`]).setNumberFormat('0.00%');

  const saved = rowToObject_(row.map(v => v === '' ? '' : String(v)), rowNumber);
  return { ok: true, rowNumber, record: saved, message: isEdit ? 'Registro atualizado.' : 'Registro salvo.' };
}

function deleteRecord(rowNumber) {
  rowNumber = Number(rowNumber);
  if (!rowNumber || rowNumber < FIRST_DATA_ROW) throw new Error('Linha inválida.');
  getSheet_().deleteRow(rowNumber);
  return { ok: true, message: 'Registro excluído.' };
}

function exportCsv() {
  const data = getInitialData().records;
  const headers = ['Data','Setor','Prontuario','Tipo','DN','Idade','Sexo','Satisfacao %','NPS','Classificacao NPS','Sugestoes','Reclamacoes','Comentarios','Elogios','Entrevistador','Encaminhamentos'];
  const lines = [headers].concat(data.map(r => [
    r.data, r.setor, r.pront, r.tipo, r.dn, r.idade, r.sexo, r.taxaSatisfacao, r.nps, classifyNps_(r.nps), r.sugestoes, r.reclamacoes, r.comentarios, r.elogios, r.entrevistador, r.encaminhamentos
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
  if (sheet.getLastColumn() >= LAST_COL && sheet.getRange('A2').getValue() === 'SETOR') return;
  sheet.getRange('B1').setValue('PESQUISA DE SATISFAÇÃO');
  sheet.getRange(2,1,1,LAST_COL).setValues([[
    'SETOR','PRONT','DATA','TIPO','DN','IDADE','SEXO','1. ACOLHIMENTO','','','','','2. ASSISTÊNCIA','','','','','','','','','3. SERVIÇOS PRESTADOS','','','','','','','MANIFESTAÇÕES','','','','NPS','ENTREVISTADOR','TABULAÇÃO DOS DADOS','','','','','','ENCAMINHAMENTOS'
  ]]);
  sheet.getRange(3,1,1,LAST_COL).setValues([[
    '','','','','','','','GENTILEZA E ATENÇÃO','AGILIDADE','CLAREZA','SATISFAÇÃO I','OBSERVAÇÕES','GENTILEZA E ATENÇÃO','IDENTIFICAÇÃO','INTIMIDADE/PRIVACIDADE','HORÁRIO DE DESCANSO','ESCLARECIMENTO','CUIDADOS PRESTADOS','CONFIANÇA E SEGURANÇA','SATISFAÇÃO II','OBSERVAÇÕES','ACESSO','ACOMODAÇÃO','LIMPEZA','ENXOVAL','ALIMENTAÇÃO','LOCOMOÇÃO','SATISFAÇÃO III','OBSERVAÇÕES','SUGESTÕES','RECLAMAÇÕES','COMENTÁRIOS','ELOGIOS','','','ÓTIMO','BOM','REGULAR','RUIM','N/A','TOTAL CONSIDERADO','TAXA DE SATISFAÇÃO',''
  ]]);
  sheet.getRange('A1:AQ3').setFontWeight('bold').setBackground('#eaf2ff');
  sheet.setFrozenRows(3);
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
  return ['setor','pront','data','tipo','sexo','nps','sugestoes','reclamacoes','comentarios','elogios'].some(k => String(r[k] || '').trim());
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
      if (v !== null) { row[COL.satisfacao1 - 1] = v; rangeListAddrs.push(`K${rowNum}`); changed = true; }
    }
    if (isEmpty(s2)) {
      const v = blockSatFromRow(row, BLOCO_ASSISTENCIA);
      if (v !== null) { row[COL.satisfacao2 - 1] = v; rangeListAddrs.push(`T${rowNum}`); changed = true; }
    }
    if (isEmpty(s3)) {
      const v = blockSatFromRow(row, BLOCO_SERVICOS);
      if (v !== null) { row[COL.satisfacao3 - 1] = v; rangeListAddrs.push(`AB${rowNum}`); changed = true; }
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

// ── OCR via Google Vision API ──

function ocrPdfViaVisionApi(base64Content) {
  const props = PropertiesService.getScriptProperties();
  const keyJson = props.getProperty('GOOGLE_CLOUD_KEY');
  Logger.log('DEBUG: keyJson exists: ' + !!keyJson);
  Logger.log('DEBUG: keyJson length: ' + (keyJson ? keyJson.length : 0));
  if (!keyJson) throw new Error('Chave Google Cloud não configurada.');

  let key;
  try {
    key = JSON.parse(keyJson);
  } catch (e) {
    Logger.log('DEBUG: JSON parse error: ' + e);
    throw new Error('Erro ao processar chave: ' + e.message);
  }
  const accessToken = getGoogleAccessToken_(key);

  const payload = {
    requests: [{
      image: { content: base64Content },
      features: [{ type: 'DOCUMENT_TEXT_DETECTION' }]
    }]
  };

  const response = UrlFetchApp.fetch('https://vision.googleapis.com/v1/images:annotate', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + accessToken },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Erro Vision API: ' + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  return parseOcrResult_(result);
}

function getGoogleAccessToken_(serviceAccountKey) {
  const now = Math.floor(Date.now() / 1000);
  const expiration = now + 3600;

  const header = { alg: 'RS256', typ: 'JWT' };
  const claim = {
    iss: serviceAccountKey.client_email,
    scope: 'https://www.googleapis.com/auth/cloud-platform',
    aud: 'https://oauth2.googleapis.com/token',
    exp: expiration,
    iat: now
  };

  const headerEncoded = Utilities.base64Encode(JSON.stringify(header), true).replace(/=/g, '');
  const claimEncoded = Utilities.base64Encode(JSON.stringify(claim), true).replace(/=/g, '');
  const signatureInput = headerEncoded + '.' + claimEncoded;

  const signature = Utilities.computeRsaSha256Signature(signatureInput, serviceAccountKey.private_key);
  const signatureEncoded = Utilities.base64Encode(signature, true).replace(/=/g, '');

  const jwt = signatureInput + '.' + signatureEncoded;

  const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: 'grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=' + jwt,
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Erro OAuth: ' + response.getContentText());
  }

  return JSON.parse(response.getContentText()).access_token;
}

function parseOcrResult_(visionResponse) {
  const responses = visionResponse.responses || [];
  if (!responses.length || responses[0].error) {
    throw new Error('Erro OCR: ' + JSON.stringify(responses[0]?.error || 'Sem resposta'));
  }

  const fullText = responses[0].fullAnnotation?.text || '';
  const lines = fullText.split('\n');

  // Parsing básico — tenta extrair campos-chave
  const getText = (pattern) => {
    const regex = new RegExp(pattern, 'i');
    const match = fullText.match(regex);
    return match ? match[1]?.trim() || '' : '';
  };

  const findRating = (fieldName) => {
    const lines2 = fullText.split('\n');
    for (let i = 0; i < lines2.length; i++) {
      if (lines2[i].toLowerCase().includes(fieldName.toLowerCase())) {
        const nextLines = lines2.slice(i, Math.min(i + 3, lines2.length)).join(' ');
        if (/ótimo|excelente/i.test(nextLines)) return 'ÓTIMO';
        if (/bom|boa/i.test(nextLines)) return 'BOM';
        if (/regular|razoável/i.test(nextLines)) return 'REGULAR';
        if (/ruim|péssimo/i.test(nextLines)) return 'RUIM';
        if (/n\/a|na|não/i.test(nextLines)) return 'N/A';
      }
    }
    return '';
  };

  const findNps = () => {
    const match = fullText.match(/\b([0-9]|10)\b/);
    return match ? parseInt(match[1]) : '';
  };

  const findType = () => {
    if (/\bpaciente\b/i.test(fullText)) return 'PACIENTE';
    if (/\bacompanhante\b/i.test(fullText)) return 'ACOMPANHANTE';
    if (/\bcolaborad|colaborador\b/i.test(fullText)) return 'COLABORADA';
    return '';
  };

  return {
    setor: getText(/UNIDADE[:\s]+([^\n]+)/),
    tipo: findType(),
    sexo: /\bf\b/i.test(fullText) ? 'F' : /\bm\b/i.test(fullText) ? 'M' : '',
    data: getText(/DATA[:\s]+([^\n]+)/),
    pront: getText(/PRONT[:\s]+([^\n]+)/),
    gentilezaAcolhimento: findRating('Gentileza'),
    agilidade: findRating('Agilidade'),
    clareza: findRating('Clareza'),
    gentilezaAssistencia: findRating('Assistência'),
    identificacao: findRating('Identificação'),
    intimidade: findRating('Intimidade'),
    horarioDescanso: findRating('Horário'),
    esclarecimento: findRating('Esclarecimento'),
    cuidados: findRating('Cuidados'),
    confianca: findRating('Confiança'),
    acesso: findRating('Acesso'),
    acomodacao: findRating('Acomodação'),
    limpeza: findRating('Limpeza'),
    enxoval: findRating('Enxoval'),
    alimentacao: findRating('Alimentação'),
    locomocao: findRating('Locomoção'),
    nps: findNps()
  };
}
