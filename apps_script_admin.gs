
/*
  BACKEND — 38º SEMINÁRIO IPBM
  Versão com inscrições + painel administrativo + check-in.

  IMPORTANTE:
  - Altere ADMIN_KEY para uma senha sua.
  - Adicione um arquivo HTML chamado "admin" no Apps Script
    e cole o conteúdo do arquivo admin.html deste pacote.
  - Implante a nova versão como Web App, executando como você
    e permitindo "Qualquer pessoa".
*/

const CONFIG = {
  SHEET_NAME: 'Inscricoes',

  // TROQUE esta chave por uma senha sua.
  ADMIN_KEY: 'IPBM-ALTERE-ESTA-CHAVE-2026',

  ADMIN_EMAIL: '' // opcional
};

function doGet(e) {
  if (e && e.parameter && e.parameter.admin === '1') {
    return HtmlService
      .createHtmlOutputFromFile('admin')
      .setTitle('Painel — 38º Seminário IPBM')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ok:true, service:'38º Seminário IPBM'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const sheet = getSheet_();
    ensureHeaders_(sheet);

    const data = JSON.parse(e.postData.contents || '{}');

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    let protocol;
    try {
      const next = sheet.getLastRow();
      protocol = 'IPBM-' +
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') +
        '-' + String(Math.max(1, next)).padStart(4, '0');

      sheet.appendRow([
        protocol,
        new Date(),
        data.nome || '',
        data.email || '',
        data.instituicao || '',
        data.telefone || '',
        data.perfil || '',
        data.modalidade || '',
        data.evento || '38º Seminário do Instituto de Pesquisa da Brigada Militar',
        '',
        ''
      ]);
    } finally {
      lock.releaseLock();
    }

    if (CONFIG.ADMIN_EMAIL) {
      MailApp.sendEmail({
        to: CONFIG.ADMIN_EMAIL,
        subject: 'Nova inscrição — 38º Seminário IPBM',
        htmlBody: '<b>Nova inscrição recebida</b><br><br>' +
          '<b>Protocolo:</b> ' + esc(protocol) + '<br>' +
          '<b>Nome:</b> ' + esc(data.nome) + '<br>' +
          '<b>E-mail:</b> ' + esc(data.email) + '<br>' +
          '<b>Modalidade:</b> ' + esc(data.modalidade)
      });
    }

    if (data.email) {
      MailApp.sendEmail({
        to: data.email,
        subject: 'Confirmação de inscrição — 38º Seminário IPBM',
        htmlBody: '<p>Olá, ' + esc(data.nome) + '.</p>' +
          '<p>Sua inscrição no <b>38º Seminário do Instituto de Pesquisa da Brigada Militar</b> foi registrada com sucesso.</p>' +
          '<p><b>Protocolo:</b> ' + esc(protocol) + '<br>' +
          '<b>Data:</b> 27 de outubro de 2026<br>' +
          '<b>Local:</b> Federação Gaúcha de Futebol, Porto Alegre/RS<br>' +
          '<b>Modalidade:</b> ' + esc(data.modalidade) + '</p>' +
          '<p>As orientações complementares serão divulgadas pela organização.</p>'
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ok:true, protocol:protocol}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getAdminData(key) {
  if (String(key || '') !== CONFIG.ADMIN_KEY) {
    return {ok:false, error:'Chave inválida.'};
  }

  const sheet = getSheet_();
  ensureHeaders_(sheet);

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return {ok:true, records:[]};

  const headers = values[0];
  const idx = {};
  headers.forEach((h,i) => idx[String(h).trim()] = i);

  const records = values.slice(1).map(row => ({
    protocol: row[idx['Protocolo']] || '',
    dataHora: formatDate_(row[idx['Data/Hora']]),
    nome: row[idx['Nome']] || '',
    email: row[idx['E-mail']] || '',
    instituicao: row[idx['Instituição/Unidade']] || '',
    telefone: row[idx['Telefone']] || '',
    perfil: row[idx['Perfil']] || '',
    modalidade: row[idx['Modalidade']] || '',
    evento: row[idx['Evento']] || '',
    checkin: row[idx['Check-in']] || '',
    checkinAt: formatDate_(row[idx['Data/Hora Check-in']])
  }));

  return {ok:true, records:records.reverse()};
}

function markCheckin(key, protocol) {
  if (String(key || '') !== CONFIG.ADMIN_KEY) {
    return {ok:false, error:'Chave inválida.'};
  }

  const sheet = getSheet_();
  ensureHeaders_(sheet);

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const pCol = headers.indexOf('Protocolo') + 1;
  const cCol = headers.indexOf('Check-in') + 1;
  const tCol = headers.indexOf('Data/Hora Check-in') + 1;

  if (pCol < 1 || cCol < 1 || tCol < 1) {
    return {ok:false, error:'Colunas de check-in não encontradas.'};
  }

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][pCol - 1]) === String(protocol)) {
      if (String(values[r][cCol - 1]) === 'Confirmado') {
        return {ok:true, already:true, timestamp:formatDate_(values[r][tCol - 1])};
      }
      const now = new Date();
      sheet.getRange(r + 1, cCol).setValue('Confirmado');
      sheet.getRange(r + 1, tCol).setValue(now);
      return {ok:true, timestamp:formatDate_(now)};
    }
  }

  return {ok:false, error:'Protocolo não encontrado.'};
}

function ensureHeaders_(sheet) {
  const headers = [
    'Protocolo','Data/Hora','Nome','E-mail','Instituição/Unidade',
    'Telefone','Perfil','Modalidade','Evento','Check-in','Data/Hora Check-in'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  const current = sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];

  // Atualiza/garante as colunas.
  const map = {};
  current.forEach((h,i)=>{ if(String(h).trim()) map[String(h).trim()] = i+1; });

  headers.forEach((h,i)=>{
    if (!map[h]) sheet.getRange(1, i+1).setValue(h);
  });

  sheet.setFrozenRows(1);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
}

function formatDate_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d)) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  return String(d);
}

function esc(value) {
  return String(value || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#039;');
}
