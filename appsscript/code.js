const SPREADSHEET_ID = '11tAsJI5ebH1zhXXicvWi6frVf-kNjEiG-5lXLoxHNLk';
const SHEET_NAME = 'Sorteios';

// Índices (0-based) no array do getValues()
const COL_TIMESTAMP = 0;
const COL_TOKEN = 1;
const COL_GIVER_NAME = 2;
const COL_GIVER_PHONE = 3;
const COL_RECEIVER_NAME = 4;
const COL_RECEIVER_PHONE = 5;
const COL_CLAIMED_AT = 6;
const COL_DEVICE_HASH = 7;

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const mode = (p.mode || '').toLowerCase();
  const callback = (p.callback || '').trim();

  try {
    if (mode === 'sortear') {
      const namesStr = p.names || '';
      const phonesStr = p.phones || '';

      const names = namesStr.split('|').map(s => (s || '').trim()).filter(Boolean);
      const phones = phonesStr.split('|').map(s => (s || '').trim());

      const participants = names.map((name, i) => ({
        name,
        phone: (phones[i] || '')
      }));

      const data = sortearAmigoSecreto_(participants);
      return json_({ ok: true, data }, callback);
    }

    if (mode === 'peek') {
      const token = (p.t || '').trim();
      const deviceHash = (p.d || '').trim(); // opcional
      const info = peekToken_(token, deviceHash);
      return json_(info, callback);
    }

    if (mode === 'claim') {
      const token = (p.t || '').trim();
      const deviceHash = (p.d || '').trim();
      const info = claimToken_(token, deviceHash);
      return json_(info, callback);
    }

    if (mode === 'reveal') {
      const token = (p.t || '').trim();
      const deviceHash = (p.d || '').trim();
      const info = revealToken_(token, deviceHash);
      return json_(info, callback);
    }

    // Compat (se alguém chamar ?t=... sem mode) — devolve como antes (sem lock)
    if ((p.t || '').trim()) {
      const token = (p.t || '').trim();
      const info = getAmigoByTokenLegacy_(token);
      return json_(info, callback);
    }

    return json_({ ok: false, error: 'Modo inválido.' }, callback);
  } catch (err) {
    return json_({ ok: false, error: String(err) }, callback);
  }
}

function doPost(e) {
  return doGet(e);
}

function json_(obj, callback) {
  const jsonString = JSON.stringify(obj);
  if (callback && /^[a-zA-Z0-9_$.]+$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + jsonString + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(jsonString)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  ensureHeader_(sheet);
  return sheet;
}

function ensureHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp',
      'Token',
      'GiverName',
      'GiverPhone',
      'ReceiverName',
      'ReceiverPhone',
      'ClaimedAt',
      'DeviceHash'
    ]);
    return;
  }

  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // garante pelo menos 8 colunas de header
  if (header.length < 8) {
    const need = 8 - header.length;
    const extras = ['ClaimedAt', 'DeviceHash'].slice(0, need);
    sheet.getRange(1, header.length + 1, 1, extras.length).setValues([extras]);
  }
}

function findRowByToken_(sheet, token) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_TOKEN] === token) return i + 1; // row index (1-based)
  }
  return -1;
}

/* ======================
   SORTEIO (derangement)
   ====================== */

function sortearAmigoSecreto_(participants) {
  if (!participants || participants.length < 2) {
    throw new Error('Você precisa de pelo menos 2 participantes para sortear.');
  }
  const cleanParticipants = participants.filter(p => p && (p.name || '').trim().length > 0);
  if (cleanParticipants.length < 2) {
    throw new Error('Você precisa de pelo menos 2 participantes com nome válido.');
  }

  const n = cleanParticipants.length;
  const perm = criarDerivacao_(n);

  const sheet = getSheet_();

  const agora = new Date();
  const result = [];

  // cache tokens existentes
  const existing = sheet.getDataRange().getValues().map(r => r[COL_TOKEN]).slice(1);
  const existingSet = {};
  existing.forEach(t => { if (t) existingSet[t] = true; });

  for (let i = 0; i < n; i++) {
    const giver = cleanParticipants[i];
    const receiver = cleanParticipants[perm[i]];
    const token = gerarTokenUnicoFromSet_(existingSet);

    sheet.appendRow([
      agora,
      token,
      giver.name,
      giver.phone || '',
      receiver.name,
      receiver.phone || '',
      '', // ClaimedAt
      ''  // DeviceHash
    ]);

    result.push({ name: giver.name, token });
  }

  return result;
}

function criarDerivacao_(n) {
  const indices = [];
  for (let i = 0; i < n; i++) indices.push(i);

  let tentativa = 0;
  let perm;
  do {
    perm = embaralhar_(indices.slice());
    tentativa++;
    if (tentativa > 2000) {
      throw new Error('Não foi possível gerar um sorteio válido. Tente novamente.');
    }
  } while (!ehDerivacao_(perm));

  return perm;
}

function embaralhar_(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = arr[i];
    arr[i] = arr[j];
    arr[j] = tmp;
  }
  return arr;
}

function ehDerivacao_(perm) {
  for (let i = 0; i < perm.length; i++) {
    if (perm[i] === i) return false;
  }
  return true;
}

function gerarTokenUnicoFromSet_(setObj) {
  let token;
  do {
    token = gerarTokenSimples_();
  } while (setObj[token]);
  setObj[token] = true;
  return token;
}

function gerarTokenSimples_() {
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let t = '';
  for (let i = 0; i < 16; i++) {
    t += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return t;
}

/* ======================
   DEVICE LOCK (peek/claim/reveal)
   ====================== */

function peekToken_(token, deviceHashOptional) {
  if (!token) return { ok: false, status: 'notfound' };

  const sheet = getSheet_();
  const row = findRowByToken_(sheet, token);
  if (row < 0) return { ok: false, status: 'notfound' };

  const giverName = sheet.getRange(row, COL_GIVER_NAME + 1).getValue();
  const claimedAt = sheet.getRange(row, COL_CLAIMED_AT + 1).getValue();
  const storedDevice = sheet.getRange(row, COL_DEVICE_HASH + 1).getValue();

  if (!storedDevice) {
    return { ok: true, status: 'free', giverName, claimedAt: '' };
  }

  if (deviceHashOptional && storedDevice === deviceHashOptional) {
    return { ok: true, status: 'owned', giverName, claimedAt: claimedAt ? String(claimedAt) : '' };
  }

  return { ok: true, status: 'locked', giverName, claimedAt: claimedAt ? String(claimedAt) : '' };
}

function claimToken_(token, deviceHash) {
  if (!token || !deviceHash) return { ok: false, status: 'invalid' };

  const sheet = getSheet_();
  const row = findRowByToken_(sheet, token);
  if (row < 0) return { ok: false, status: 'notfound' };

  const giverName = sheet.getRange(row, COL_GIVER_NAME + 1).getValue();
  const storedDevice = sheet.getRange(row, COL_DEVICE_HASH + 1).getValue();

  if (!storedDevice) {
    sheet.getRange(row, COL_DEVICE_HASH + 1).setValue(deviceHash);
    sheet.getRange(row, COL_CLAIMED_AT + 1).setValue(new Date());
    return { ok: true, status: 'claimed', giverName };
  }

  if (storedDevice === deviceHash) {
    return { ok: true, status: 'owned', giverName };
  }

  return { ok: false, status: 'locked', giverName };
}

function revealToken_(token, deviceHash) {
  if (!token || !deviceHash) return { ok: false, status: 'invalid' };

  const sheet = getSheet_();
  const row = findRowByToken_(sheet, token);
  if (row < 0) return { ok: false, status: 'notfound' };

  const storedDevice = sheet.getRange(row, COL_DEVICE_HASH + 1).getValue();
  if (!storedDevice) {
    sheet.getRange(row, COL_DEVICE_HASH + 1).setValue(deviceHash);
    sheet.getRange(row, COL_CLAIMED_AT + 1).setValue(new Date());
  } else if (storedDevice !== deviceHash) {
    return { ok: false, status: 'locked' };
  }

  const giverName = sheet.getRange(row, COL_GIVER_NAME + 1).getValue();
  const friendName = sheet.getRange(row, COL_RECEIVER_NAME + 1).getValue();

  return { ok: true, giverName, friendName };
}

/* ======================
   Compat antigo: ?t=... (sem lock)
   ====================== */
function getAmigoByTokenLegacy_(token) {
  const sheet = getSheet_();
  const row = findRowByToken_(sheet, token);
  if (row < 0) return { ok: false };

  return {
    ok: true,
    giverName: sheet.getRange(row, COL_GIVER_NAME + 1).getValue(),
    friendName: sheet.getRange(row, COL_RECEIVER_NAME + 1).getValue()
  };
}
