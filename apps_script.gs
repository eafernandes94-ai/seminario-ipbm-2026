
/*
  BACKEND DE INSCRIÇÕES — 38º SEMINÁRIO IPBM

  Passos:
  1. Crie uma planilha no Google Sheets.
  2. Abra Extensões > Apps Script.
  3. Cole este código.
  4. Publique como Web App:
     - Executar como: você
     - Acesso: qualquer pessoa
  5. Copie a URL /exec e cole em config.js.

  A planilha terá a aba "Inscricoes".
*/

const CONFIG = {
  SHEET_NAME: 'Inscricoes',
  ADMIN_EMAIL: '' // opcional; coloque um e-mail para receber aviso de nova inscrição
};

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ok:true, service:'38º Seminário IPBM'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);

    const headers = ['Protocolo','Data/Hora','Nome','E-mail','Instituição/Unidade','Telefone','Perfil','Modalidade','Evento'];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1,1,1,headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }

    const data = JSON.parse(e.postData.contents || '{}');
    const protocol = 'IPBM-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    sheet.appendRow([
      protocol, new Date(), data.nome || '', data.email || '', data.instituicao || '',
      data.telefone || '', data.perfil || '', data.modalidade || '',
      data.evento || '38º Seminário do Instituto de Pesquisa da Brigada Militar'
    ]);

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
          '<p>Sua inscrição no <b>38º Seminário do Instituto de Pesquisa da Brigada Militar</b> foi registrada.</p>' +
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

function esc(value) {
  return String(value || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#039;');
}
