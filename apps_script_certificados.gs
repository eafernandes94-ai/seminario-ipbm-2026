/*
  BACKEND — 38º SEMINÁRIO IPBM
  Inscrições + painel administrativo + check-in + participação + certificados.

  Mude ADMIN_KEY antes de publicar.
*/
const CONFIG={
  SHEET_NAME:'Inscricoes',
  ADMIN_KEY:'IPBM-ALTERE-ESTA-CHAVE-2026',
  ADMIN_EMAIL:''
};
const CERT_FOLDER_NAME='Certificados — 38º Seminário IPBM';
const SITE_BASE='https://eafernandes94-ai.github.io/seminario-ipbm-2026/';

function doGet(e){
  if(e&&e.parameter&&e.parameter.admin==='1'){
    return HtmlService.createHtmlOutputFromFile('admin').setTitle('Painel — 38º Seminário IPBM').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if(e&&e.parameter&&e.parameter.cert==='1'){
    return HtmlService.createHtmlOutputFromFile('certificate').setTitle('Certificado — 38º Seminário IPBM').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return ContentService.createTextOutput(JSON.stringify({ok:true,service:'38º Seminário IPBM'})).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try{
    const sheet=getSheet_(); ensureHeaders_(sheet); const data=JSON.parse(e.postData.contents||'{}');
    const lock=LockService.getScriptLock(); lock.waitLock(10000);
    let protocol;
    try{
      protocol='IPBM-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd')+'-'+String(Math.max(1,sheet.getLastRow())).padStart(4,'0');
      sheet.appendRow([protocol,new Date(),data.nome||'',data.email||'',data.instituicao||'',data.telefone||'',data.perfil||'',data.modalidade||'',data.evento||'38º Seminário do Instituto de Pesquisa da Brigada Militar','','','','']);
    }finally{lock.releaseLock()}
    if(CONFIG.ADMIN_EMAIL){MailApp.sendEmail({to:CONFIG.ADMIN_EMAIL,subject:'Nova inscrição — 38º Seminário IPBM',htmlBody:'<b>Nova inscrição</b><br><br><b>Protocolo:</b> '+esc(protocol)+'<br><b>Nome:</b> '+esc(data.nome)+'<br><b>E-mail:</b> '+esc(data.email)+'<br><b>Modalidade:</b> '+esc(data.modalidade)})}
    if(data.email){MailApp.sendEmail({to:data.email,subject:'Confirmação de inscrição — 38º Seminário IPBM',htmlBody:'<p>Olá, '+esc(data.nome)+'.</p><p>Sua inscrição foi registrada.</p><p><b>Protocolo:</b> '+esc(protocol)+'<br><b>Data:</b> 27 de outubro de 2026<br><b>Local:</b> Federação Gaúcha de Futebol, Porto Alegre/RS<br><b>Modalidade:</b> '+esc(data.modalidade)+'</p><p>O certificado somente será liberado após confirmação da participação pela organização.</p>'})}
    return ContentService.createTextOutput(JSON.stringify({ok:true,protocol:protocol})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON)}
}

function getAdminData(key){
  if(String(key||'')!==CONFIG.ADMIN_KEY)return {ok:false,error:'Chave inválida.'};
  const sheet=getSheet_();ensureHeaders_(sheet);const vals=sheet.getDataRange().getValues();if(vals.length<=1)return {ok:true,records:[]};
  const h=vals[0],idx={};h.forEach((x,i)=>idx[String(x).trim()]=i);
  const records=vals.slice(1).map(row=>({
    protocol:row[idx['Protocolo']]||'',dataHora:formatDate_(row[idx['Data/Hora']]),nome:row[idx['Nome']]||'',email:row[idx['E-mail']]||'',instituicao:row[idx['Instituição/Unidade']]||'',telefone:row[idx['Telefone']]||'',perfil:row[idx['Perfil']]||'',modalidade:row[idx['Modalidade']]||'',evento:row[idx['Evento']]||'',checkin:row[idx['Check-in']]||'',checkinAt:formatDate_(row[idx['Data/Hora Check-in']]),participacao:row[idx['Participação']]||'',participacaoAt:formatDate_(row[idx['Data/Hora Participação']]),certificadoId:row[idx['Certificado PDF ID']]||''
  }));
  return {ok:true,records:records.reverse()};
}

function markCheckin(key,protocol){
  if(String(key||'')!==CONFIG.ADMIN_KEY)return {ok:false,error:'Chave inválida.'};
  const sheet=getSheet_();ensureHeaders_(sheet);const vals=sheet.getDataRange().getValues(),h=vals[0];const p=h.indexOf('Protocolo')+1,c=h.indexOf('Check-in')+1,t=h.indexOf('Data/Hora Check-in')+1;
  for(let r=1;r<vals.length;r++){if(String(vals[r][p-1])===String(protocol)){if(String(vals[r][c-1])==='Confirmado')return {ok:true,already:true,timestamp:formatDate_(vals[r][t-1])};const now=new Date();sheet.getRange(r+1,c).setValue('Confirmado');sheet.getRange(r+1,t).setValue(now);return {ok:true,timestamp:formatDate_(now)}}}
  return {ok:false,error:'Protocolo não encontrado.'};
}

function markParticipation(key,protocol,status){
  if(String(key||'')!==CONFIG.ADMIN_KEY)return {ok:false,error:'Chave inválida.'};
  const allowed=['Confirmado','Pendente']; if(allowed.indexOf(status)<0)return {ok:false,error:'Status inválido.'};
  const sheet=getSheet_();ensureHeaders_(sheet);const vals=sheet.getDataRange().getValues(),h=vals[0],p=h.indexOf('Protocolo')+1,s=h.indexOf('Participação')+1,t=h.indexOf('Data/Hora Participação')+1,cid=h.indexOf('Certificado PDF ID')+1;
  for(let r=1;r<vals.length;r++){
    if(String(vals[r][p-1])===String(protocol)){
      sheet.getRange(r+1,s).setValue(status);
      sheet.getRange(r+1,t).setValue(new Date());
      if(status!=='Confirmado' && vals[r][cid-1]){try{DriveApp.getFileById(String(vals[r][cid-1])).setTrashed(true)}catch(e){} sheet.getRange(r+1,cid).clearContent();}
      return {ok:true,status:status,timestamp:formatDate_(new Date())};
    }
  }
  return {ok:false,error:'Protocolo não encontrado.'};
}

function getCertificateStatus(protocol,email){
  const record=findRecord_(protocol,email); if(!record)return {ok:false,message:'Protocolo e e-mail não conferem.'};
  if(record.participacao!=='Confirmado')return {ok:false,message:'Sua inscrição foi encontrada, mas o certificado ainda não foi liberado pela organização.'};
  return {ok:true,nome:record.nome,message:'Participação confirmada. O certificado está disponível para emissão.'};
}

function generateCertificate(protocol,email){
  const record=findRecord_(protocol,email); if(!record)return {ok:false,message:'Protocolo e e-mail não conferem.'};
  if(record.participacao!=='Confirmado')return {ok:false,message:'O certificado somente pode ser emitido após a confirmação da participação pela organização.'};
  const sheet=getSheet_();ensureHeaders_(sheet);const vals=sheet.getDataRange().getValues(),h=vals[0],p=h.indexOf('Protocolo')+1,cid=h.indexOf('Certificado PDF ID')+1;
  let rowNumber=-1,currentId='';for(let r=1;r<vals.length;r++){if(String(vals[r][p-1])===String(protocol)){rowNumber=r+1;currentId=String(vals[r][cid-1]||'');break}}
  let file=null;if(currentId){try{file=DriveApp.getFileById(currentId)}catch(e){file=null}}
  if(!file){file=createCertificateFile_(record);if(rowNumber>0)sheet.getRange(rowNumber,cid).setValue(file.getId())}
  const blob=file.getBlob();
  return {ok:true,base64:Utilities.base64Encode(blob.getBytes()),filename:file.getName()};
}

function findRecord_(protocol,email){
  const sheet=getSheet_();ensureHeaders_(sheet);const vals=sheet.getDataRange().getValues();if(vals.length<=1)return null;const h=vals[0],idx={};h.forEach((x,i)=>idx[String(x).trim()]=i);
  const p=String(protocol||'').trim(),e=String(email||'').trim().toLowerCase();
  for(let r=1;r<vals.length;r++){if(String(vals[r][idx['Protocolo']]||'').trim()===p && String(vals[r][idx['E-mail']]||'').trim().toLowerCase()===e){return {protocol:vals[r][idx['Protocolo']]||'',nome:vals[r][idx['Nome']]||'',email:vals[r][idx['E-mail']]||'',instituicao:vals[r][idx['Instituição/Unidade']]||'',modalidade:vals[r][idx['Modalidade']]||'',participacao:vals[r][idx['Participação']]||'',evento:vals[r][idx['Evento']]||''}}}
  return null;
}

function createCertificateFile_(record){
  const folder=getCertificateFolder_();
  const doc=DocumentApp.create('Certificado - '+record.protocol+' - '+sanitize_(record.nome));
  const body=doc.getBody();body.clear();body.setMarginTop(44).setMarginBottom(44).setMarginLeft(60).setMarginRight(60);
  const p1=body.appendParagraph('INSTITUTO DE PESQUISA DA BRIGADA MILITAR');p1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);p1.setBold(true).setFontSize(11).setForegroundColor('#745b16');
  body.appendParagraph('');
  const title=body.appendParagraph('CERTIFICADO DE PARTICIPAÇÃO');title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);title.setBold(true).setFontSize(25).setForegroundColor('#06192d');
  const line=body.appendParagraph('38º SEMINÁRIO DO INSTITUTO DE PESQUISA DA BRIGADA MILITAR');line.setAlignment(DocumentApp.HorizontalAlignment.CENTER);line.setBold(true).setFontSize(12).setForegroundColor('#6b7789');
  body.appendParagraph('');
  const txt=body.appendParagraph('Certificamos que');txt.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(12);
  const name=body.appendParagraph(record.nome);name.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true).setFontSize(22).setForegroundColor('#06192d');
  const desc=body.appendParagraph('participou do 38º Seminário do Instituto de Pesquisa da Brigada Militar, realizado em 27 de outubro de 2026, em Porto Alegre/RS, na modalidade '+String(record.modalidade||'participação registrada').toLowerCase()+'.');desc.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(12);
  body.appendParagraph('');
  const proto=body.appendParagraph('Protocolo: '+record.protocol);proto.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(9).setForegroundColor('#687588');
  body.appendParagraph('');body.appendParagraph('');
  const sig=body.appendParagraph('Instituto de Pesquisa da Brigada Militar – IPBM');sig.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true).setFontSize(11);const org=body.appendParagraph('Brigada Militar • Departamento de Educação e Cultura');org.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(9).setForegroundColor('#687588');
  doc.saveAndClose();
  const pdf=DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF).setName('Certificado - '+record.protocol+' - '+sanitize_(record.nome)+'.pdf');
  const file=folder.createFile(pdf);DriveApp.getFileById(doc.getId()).setTrashed(true);return file;
}
function getCertificateFolder_(){const props=PropertiesService.getScriptProperties();const saved=props.getProperty('CERT_FOLDER_ID');if(saved){try{return DriveApp.getFolderById(saved)}catch(e){}}const folders=DriveApp.getFoldersByName(CERT_FOLDER_NAME);const folder=folders.hasNext()?folders.next():DriveApp.createFolder(CERT_FOLDER_NAME);props.setProperty('CERT_FOLDER_ID',folder.getId());return folder;}
function ensureHeaders_(sheet){
 const headers=['Protocolo','Data/Hora','Nome','E-mail','Instituição/Unidade','Telefone','Perfil','Modalidade','Evento','Check-in','Data/Hora Check-in','Participação','Data/Hora Participação','Certificado PDF ID'];
 if(sheet.getLastRow()===0){sheet.getRange(1,1,1,headers.length).setValues([headers]);sheet.setFrozenRows(1);return;}
 const current=sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),headers.length)).getValues()[0];const map={};current.forEach((h,i)=>{if(String(h).trim())map[String(h).trim()]=i+1});headers.forEach((h,i)=>{if(!map[h])sheet.getRange(1,i+1).setValue(h)});sheet.setFrozenRows(1);
}
function getSheet_(){const ss=SpreadsheetApp.getActiveSpreadsheet();return ss.getSheetByName(CONFIG.SHEET_NAME)||ss.insertSheet(CONFIG.SHEET_NAME)}
function formatDate_(d){if(!d)return '';if(Object.prototype.toString.call(d)==='[object Date]'&&!isNaN(d))return Utilities.formatDate(d,Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm');return String(d)}
function esc(v){return String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;')}
function sanitize_(v){return String(v||'Participante').replace(/[\\/:*?"<>|]/g,' ').replace(/\s+/g,' ').trim().slice(0,80)}
