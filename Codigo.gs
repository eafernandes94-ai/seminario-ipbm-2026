/*
  BACKEND — 38º SEMINÁRIO IPBM
  Versão com certificado personalizado a partir de um modelo Google Slides.

  IMPORTANTE:
  1) Mantenha sua chave atual em ADMIN_KEY.
  2) O modelo oficial é a apresentação Google Slides indicada em CERT_TEMPLATE_PRESENTATION_ID.
  3) O modelo deve conter uma única caixa de texto com o marcador {{NOME}}.
  4) O certificado só é gerado após participação confirmada + certificado liberado.
*/

const CONFIG = {
  SHEET_NAME: 'Inscricoes',
  ADMIN_KEY: 'COLOQUE_AQUI_SUA_CHAVE_ADMINISTRATIVA',
  ADMIN_EMAIL: ''
};

const CERT_FOLDER_NAME = 'Certificados — 38º Seminário IPBM';
const CERT_TEMPLATE_PRESENTATION_ID = '1F9qzntKj9qa5ZBFVWIDtbQp3PAwbQOL-oq4WexfkYLU';

// A caixa do {{NOME}} do modelo é centralizada e, se necessário, ampliada.
const CERT_NAME_MIN_WIDTH_PT = 300;
const CERT_NAME_MARGIN_PT = 8;

/** Execute manualmente uma vez para autorizar Drive, Docs e Slides. */
function authorizeCertificateServices(){
  const root = DriveApp.getRootFolder();
  root.getName();

  const doc = DocumentApp.create('Autorização — 38º Seminário IPBM');
  const docId = doc.getId();
  doc.getBody().appendParagraph('Teste de autorização concluído.');
  doc.saveAndClose();
  try{DriveApp.getFileById(docId).setTrashed(true);}catch(e){}

  const pres = SlidesApp.openById(CERT_TEMPLATE_PRESENTATION_ID);
  pres.getPageWidth();
  pres.getSlides()[0].getPageElements().length;

  return 'Autorização concluída. Modelo Google Slides acessível.';
}

function doGet(e){
  if(e && e.parameter && e.parameter.admin === '1'){
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('Painel — 38º Seminário IPBM')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if(e && e.parameter && e.parameter.cert === '1'){
    return HtmlService.createHtmlOutputFromFile('certificate')
      .setTitle('Certificado — 38º Seminário IPBM')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return ContentService.createTextOutput(JSON.stringify({ok:true,service:'38º Seminário IPBM'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try{
    const sheet=getSheet_();
    ensureHeaders_(sheet);
    const data=JSON.parse(e.postData.contents||'{}');
    const lock=LockService.getScriptLock();
    lock.waitLock(10000);
    let protocol;
    try{
      protocol=nextProtocol_(sheet);
      sheet.appendRow([
        protocol,
        new Date(),
        data.nome||'',
        data.email||'',
        data.instituicao||'',
        data.telefone||'',
        data.perfil||'',
        data.modalidade||'',
        data.evento||'38º Seminário do Instituto de Pesquisa da Brigada Militar',
        '', '', '', '', '', '', '',
        data.ufCorporacao||'',
        data.postoGraduacao||''
      ]);
    }finally{lock.releaseLock();}

    if(CONFIG.ADMIN_EMAIL){
      MailApp.sendEmail({to:CONFIG.ADMIN_EMAIL,subject:'Nova inscrição — 38º Seminário IPBM',htmlBody:
        '<b>Nova inscrição recebida</b><br><br><b>Protocolo:</b> '+esc(protocol)+'<br>'+ 
        '<b>Nome:</b> '+esc(data.nome)+'<br><b>E-mail:</b> '+esc(data.email)+'<br>'+ 
        '<b>Modalidade:</b> '+esc(data.modalidade)+'<br>'+
        '<b>Instituição/Força:</b> '+esc(data.perfil)+'<br>'+
        '<b>UF:</b> '+esc(data.ufCorporacao)+'<br>'+
        '<b>Posto/Graduação:</b> '+esc(data.postoGraduacao)});
    }
    if(data.email){
      MailApp.sendEmail({to:data.email,subject:'Confirmação de inscrição — 38º Seminário IPBM',htmlBody:
        '<p>Olá, '+esc(data.nome)+'.</p>'+ 
        '<p>Sua inscrição no <b>38º Seminário do Instituto de Pesquisa da Brigada Militar</b> foi registrada.</p>'+ 
        '<p><b>Protocolo:</b> '+esc(protocol)+'<br><b>Data:</b> 27 de outubro de 2026<br>'+ 
        '<b>Local:</b> Federação Gaúcha de Futebol, Porto Alegre/RS<br><b>Modalidade:</b> '+esc(data.modalidade)+'</p>'+ 
        '<p>O certificado somente ficará disponível após a confirmação da participação e a autorização de emissão pela organização.</p>'});
    }
    return ContentService.createTextOutput(JSON.stringify({ok:true,protocol:protocol})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function nextProtocol_(sheet){
  const props=PropertiesService.getScriptProperties();
  let current=Number(props.getProperty('NEXT_PROTOCOL_NUMBER')||0);
  if(!current){
    const lastRow=sheet.getLastRow();
    if(lastRow>1){
      const col=sheet.getRange(2,1,lastRow-1,1).getValues();
      for(let i=0;i<col.length;i++){
        const m=String(col[i][0]||'').match(/^IPBM-2026-(\d+)$/);
        if(m) current=Math.max(current,Number(m[1]));
      }
    }
  }
  current++; props.setProperty('NEXT_PROTOCOL_NUMBER',String(current));
  return 'IPBM-2026-'+String(current).padStart(6,'0');
}

function getAdminData(key){
  if(String(key||'')!==CONFIG.ADMIN_KEY)return {ok:false,error:'Chave inválida.'};
  const sheet=getSheet_(); ensureHeaders_(sheet);
  const vals=sheet.getDataRange().getValues();
  if(vals.length<=1)return {ok:true,records:[]};
  const h=vals[0],idx={}; h.forEach((x,i)=>idx[String(x).trim()]=i);
  const records=vals.slice(1).map(row=>({
    protocol:row[idx['Protocolo']]||'',dataHora:formatDate_(row[idx['Data/Hora']]),nome:row[idx['Nome']]||'',email:row[idx['E-mail']]||'',
    instituicao:row[idx['Instituição/Unidade']]||'',telefone:row[idx['Telefone']]||'',perfil:row[idx['Perfil']]||'',ufCorporacao:row[idx['UF da Corporação']]||'',postoGraduacao:row[idx['Posto / Graduação']]||'',modalidade:row[idx['Modalidade']]||'',evento:row[idx['Evento']]||'',
    checkin:row[idx['Check-in']]||'',checkinAt:formatDate_(row[idx['Data/Hora Check-in']]),participacao:row[idx['Participação']]||'',
    participacaoAt:formatDate_(row[idx['Data/Hora Participação']]),certAutorizado:row[idx['Certificado Autorizado']]||'',
    certAutorizadoAt:formatDate_(row[idx['Data/Hora Autorização Certificado']]),certificadoId:row[idx['Certificado PDF ID']]||''
  }));
  return {ok:true,records:records.reverse()};
}

function markCheckin(key,protocol){
  if(String(key||'')!==CONFIG.ADMIN_KEY)return {ok:false,error:'Chave inválida.'};
  const sheet=getSheet_();ensureHeaders_(sheet); const vals=sheet.getDataRange().getValues(),h=vals[0];
  const p=h.indexOf('Protocolo')+1,c=h.indexOf('Check-in')+1,t=h.indexOf('Data/Hora Check-in')+1;
  for(let r=1;r<vals.length;r++) if(String(vals[r][p-1])===String(protocol)){
    if(String(vals[r][c-1])==='Confirmado')return {ok:true,already:true,timestamp:formatDate_(vals[r][t-1])};
    const now=new Date();sheet.getRange(r+1,c).setValue('Confirmado');sheet.getRange(r+1,t).setValue(now);return {ok:true,timestamp:formatDate_(now)};
  }
  return {ok:false,error:'Protocolo não encontrado.'};
}

function markParticipation(key,protocol,status){
  if(String(key||'')!==CONFIG.ADMIN_KEY)return {ok:false,error:'Chave inválida.'};
  if(['Confirmado','Pendente'].indexOf(status)<0)return {ok:false,error:'Status inválido.'};
  const sheet=getSheet_();ensureHeaders_(sheet);const vals=sheet.getDataRange().getValues(),h=vals[0];
  const p=h.indexOf('Protocolo')+1,s=h.indexOf('Participação')+1,t=h.indexOf('Data/Hora Participação')+1,ca=h.indexOf('Certificado Autorizado')+1,cat=h.indexOf('Data/Hora Autorização Certificado')+1,cid=h.indexOf('Certificado PDF ID')+1;
  for(let r=1;r<vals.length;r++) if(String(vals[r][p-1])===String(protocol)){
    const now=new Date();sheet.getRange(r+1,s).setValue(status);sheet.getRange(r+1,t).setValue(now);
    if(status!=='Confirmado'){
      sheet.getRange(r+1,ca).setValue('Bloqueado');sheet.getRange(r+1,cat).setValue(now);
      const oldId=String(vals[r][cid-1]||'');if(oldId){try{DriveApp.getFileById(oldId).setTrashed(true)}catch(e){}}
      sheet.getRange(r+1,cid).clearContent();
    }
    return {ok:true,status:status,timestamp:formatDate_(now)};
  }
  return {ok:false,error:'Protocolo não encontrado.'};
}

function markCertificateAuthorization(key,protocol,status){
  if(String(key||'')!==CONFIG.ADMIN_KEY)return {ok:false,error:'Chave inválida.'};
  if(['Liberado','Bloqueado'].indexOf(status)<0)return {ok:false,error:'Status de certificado inválido.'};
  const sheet=getSheet_();ensureHeaders_(sheet);const vals=sheet.getDataRange().getValues(),h=vals[0];
  const p=h.indexOf('Protocolo')+1,pa=h.indexOf('Participação')+1,ca=h.indexOf('Certificado Autorizado')+1,cat=h.indexOf('Data/Hora Autorização Certificado')+1,cid=h.indexOf('Certificado PDF ID')+1;
  for(let r=1;r<vals.length;r++) if(String(vals[r][p-1])===String(protocol)){
    if(status==='Liberado'&&String(vals[r][pa-1])!=='Confirmado')return {ok:false,error:'Confirme a participação antes de liberar o certificado.'};
    const now=new Date();sheet.getRange(r+1,ca).setValue(status);sheet.getRange(r+1,cat).setValue(now);
    if(status==='Bloqueado'){
      const oldId=String(vals[r][cid-1]||'');if(oldId){try{DriveApp.getFileById(oldId).setTrashed(true)}catch(e){}}
      sheet.getRange(r+1,cid).clearContent();
    }
    return {ok:true,status:status,timestamp:formatDate_(now)};
  }
  return {ok:false,error:'Protocolo não encontrado.'};
}

function getCertificateStatus(protocol,email){
  const record=findRecord_(protocol,email);if(!record)return {ok:false,message:'Protocolo e e-mail não conferem.'};
  if(record.participacao!=='Confirmado')return {ok:false,message:'Sua inscrição foi encontrada, mas a participação ainda não foi confirmada pela organização.'};
  if(record.certAutorizado!=='Liberado')return {ok:false,message:'Sua participação foi confirmada, mas o certificado ainda não foi autorizado para emissão pela organização.'};
  return {ok:true,nome:formatCertificateIdentity_(record),message:'Participação confirmada e certificado autorizado. O certificado está disponível para emissão.'};
}

function generateCertificate(protocol,email){
  const record=findRecord_(protocol,email);if(!record)return {ok:false,message:'Protocolo e e-mail não conferem.'};
  if(record.participacao!=='Confirmado')return {ok:false,message:'O certificado somente pode ser emitido após a confirmação da participação.'};
  if(record.certAutorizado!=='Liberado')return {ok:false,message:'O certificado ainda não foi autorizado pela organização.'};
  const sheet=getSheet_();ensureHeaders_(sheet);const vals=sheet.getDataRange().getValues(),h=vals[0];
  const p=h.indexOf('Protocolo')+1,cid=h.indexOf('Certificado PDF ID')+1;let rowNumber=-1,currentId='';
  for(let r=1;r<vals.length;r++)if(String(vals[r][p-1])===String(protocol)){rowNumber=r+1;currentId=String(vals[r][cid-1]||'');break;}
  let file=null;if(currentId){try{file=DriveApp.getFileById(currentId)}catch(e){file=null}}
  if(!file){file=createCertificateFile_(record);if(rowNumber>0)sheet.getRange(rowNumber,cid).setValue(file.getId());}
  const blob=file.getBlob();return {ok:true,base64:Utilities.base64Encode(blob.getBytes()),filename:file.getName()};
}

function findRecord_(protocol,email){
  const sheet=getSheet_();ensureHeaders_(sheet);const vals=sheet.getDataRange().getValues();if(vals.length<=1)return null;
  const h=vals[0],idx={};h.forEach((x,i)=>idx[String(x).trim()]=i);const p=String(protocol||'').trim(),e=String(email||'').trim().toLowerCase();
  for(let r=1;r<vals.length;r++)if(String(vals[r][idx['Protocolo']]||'').trim()===p&&String(vals[r][idx['E-mail']]||'').trim().toLowerCase()===e){
    return {
      protocol:vals[r][idx['Protocolo']]||'',
      nome:vals[r][idx['Nome']]||'',
      email:vals[r][idx['E-mail']]||'',
      instituicao:vals[r][idx['Instituição/Unidade']]||'',
      perfil:vals[r][idx['Perfil']]||'',
      ufCorporacao:vals[r][idx['UF da Corporação']]||'',
      postoGraduacao:vals[r][idx['Posto / Graduação']]||'',
      modalidade:vals[r][idx['Modalidade']]||'',
      participacao:vals[r][idx['Participação']]||'',
      certAutorizado:vals[r][idx['Certificado Autorizado']]||'',
      evento:vals[r][idx['Evento']]||''
    };
  }
  return null;
}

/**
 * Cria o certificado em Google Slides usando a arte institucional como fundo
 * e inserindo somente o nome do participante.
 */
function createCertificateFile_(record){
  const folder=getCertificateFolder_();
  const templateFile=getCertificatePresentationTemplate_();

  // Faz uma cópia do modelo já montado pelo usuário. Assim, a página,
  // logos, assinatura, textos e proporções permanecem exatamente iguais.
  const tempFile=templateFile.makeCopy(
    'TMP Certificado - '+record.protocol+' - '+sanitize_(record.nome),
    folder
  );

  const pres=SlidesApp.openById(tempFile.getId());
  const slide=pres.getSlides()[0];
  const nomeCertificado=formatCertificateIdentity_(record);

  let found=false;
  slide.getPageElements().forEach(function(el){
    if(el.getPageElementType()!==SlidesApp.PageElementType.SHAPE) return;
    const shape=el.asShape();
    const text=shape.getText();
    const current=text.asString();
    if(current.indexOf('{{NOME}}')===-1) return;

    // Preserva a caixa criada no modelo, mas garante uma largura útil mínima
    // e centraliza a caixa na página para reduzir quebras de linha.
    const slideWidth=pres.getPageWidth();
    const currentWidth=shape.getWidth();
    const usableWidth=Math.min(
      Math.max(currentWidth, CERT_NAME_MIN_WIDTH_PT),
      slideWidth-(CERT_NAME_MARGIN_PT*2)
    );
    shape.setWidth(usableWidth);
    shape.setLeft((slideWidth-usableWidth)/2);

    // Substitui somente o marcador do nome.
    text.setText(nomeCertificado);
    const fontSize=fitCertificateNameFontSize_(nomeCertificado, usableWidth);
    text.getTextStyle()
      .setFontFamily('Times New Roman')
      .setFontSize(fontSize)
      .setBold(true)
      .setForegroundColor('#202020');
    text.getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    try{shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);}catch(e){}
    found=true;
  });

  if(!found){
    try{tempFile.setTrashed(true)}catch(e){}
    throw new Error('O modelo do Google Slides não contém uma caixa de texto com o marcador {{NOME}}.');
  }

  pres.saveAndClose();
  Utilities.sleep(1200);

  const pdf=tempFile.getBlob()
    .getAs(MimeType.PDF)
    .setName('Certificado - '+record.protocol+' - '+sanitize_(record.nome)+'.pdf');

  const file=folder.createFile(pdf);
  tempFile.setTrashed(true);
  return file;
}

function fitCertificateNameFontSize_(name, availableWidthPt){
  // Estimativa conservadora da largura em Times New Roman Bold para nomes em caixa alta.
  // A fonte começa em 10 pt e é reduzida automaticamente até caber em uma linha.
  const clean=String(name||'').trim();
  for(let size=10; size>=7; size--){
    const estimatedWidth=estimateCertificateNameWidth_(clean, size);
    if(estimatedWidth <= availableWidthPt*0.90) return size;
  }
  return 7;
}

function estimateCertificateNameWidth_(name, fontSize){
  let units=0;
  for(let i=0;i<name.length;i++){
    const ch=name.charAt(i);
    if(ch===' ') units+=0.28;
    else if('IÍJTLF'.indexOf(ch)>=0) units+=0.34;
    else if('MWÁÂÃÉÊÓÔÕ'.indexOf(ch)>=0) units+=0.88;
    else units+=0.58;
  }
  return units*fontSize;
}

function getCertificatePresentationTemplate_(){
  try{
    return DriveApp.getFileById(CERT_TEMPLATE_PRESENTATION_ID);
  }catch(e){
    throw new Error('Não foi possível acessar o modelo de certificado do Google Slides. Verifique o ID da apresentação e se ela está acessível pela conta do Apps Script.');
  }
}

function formatCertificateIdentity_(record){
  const name=formatCertificateName_(record && record.nome);
  const posto=normalizePostoGraduacao_(record && record.perfil,record && record.postoGraduacao);
  return posto ? posto+' '+name : name;
}

function normalizePostoGraduacao_(perfil,posto){
  const p=String(perfil||'').trim();
  const v=String(posto||'').trim();
  if(!v)return '';

  const lists={
    'Brigada Militar – RS':[
      'Cel PM','Ten-Cel PM','Maj PM','Cap PM','1º Ten PM','2º Ten PM','Asp Of PM','Al Of PM','ST PM','1º Sgt PM','2º Sgt PM','3º Sgt PM','Cb PM','Sd PM'
    ],
    'Polícia Militar – outro Estado':[
      'Cel PM','Ten-Cel PM','Maj PM','Cap PM','1º Ten PM','2º Ten PM','Asp Of PM','Al Of PM','ST PM','1º Sgt PM','2º Sgt PM','3º Sgt PM','Cb PM','Sd PM'
    ],
    'Corpo de Bombeiros Militar – RS':[
      'Cel BM','Ten-Cel BM','Maj BM','Cap BM','1º Ten BM','2º Ten BM','Asp Of BM','Al Of BM','ST BM','1º Sgt BM','2º Sgt BM','3º Sgt BM','Cb BM','Sd BM'
    ],
    'Corpo de Bombeiros Militar – outro Estado':[
      'Cel BM','Ten-Cel BM','Maj BM','Cap BM','1º Ten BM','2º Ten BM','Asp Of BM','Al Of BM','ST BM','1º Sgt BM','2º Sgt BM','3º Sgt BM','Cb BM','Sd BM'
    ],
    'Exército Brasileiro':[
      'Gen Ex EB','Gen Div EB','Gen Bda EB','Cel EB','Ten-Cel EB','Maj EB','Cap EB','1º Ten EB','2º Ten EB','Asp Of EB','Cad EB','S Ten EB','1º Sgt EB','2º Sgt EB','3º Sgt EB','Cb EB','Sd EB'
    ],
    'Marinha do Brasil':[
      'Alte MB','Alte Esq MB','V Alte MB','C Alte MB','CMG MB','CF MB','CC MB','CT MB','1º Ten MB','2º Ten MB','GM MB','Asp Of MB','SO MB','1º SG MB','2º SG MB','3º SG MB','CB MB','MN MB'
    ],
    'Força Aérea Brasileira':[
      'Mar FAB','Ten Brig Ar','Maj Brig Ar','Brig Ar','Cel FAB','Ten-Cel FAB','Maj FAB','Cap FAB','1º Ten FAB','2º Ten FAB','Asp Of FAB','Cad FAB','SO FAB','1S FAB','2S FAB','3S FAB','Cb FAB','S1 FAB','S2 FAB'
    ]
  };

  const allowed=lists[p]||[];
  return allowed.indexOf(v)>=0 ? v : '';
}

function formatCertificateName_(name){
  return String(name || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLocaleUpperCase('pt-BR');
}

function getCertificateFolder_(){
  const props=PropertiesService.getScriptProperties();const saved=props.getProperty('CERT_FOLDER_ID');
  if(saved){try{return DriveApp.getFolderById(saved)}catch(e){}}
  const folders=DriveApp.getFoldersByName(CERT_FOLDER_NAME);const folder=folders.hasNext()?folders.next():DriveApp.createFolder(CERT_FOLDER_NAME);
  props.setProperty('CERT_FOLDER_ID',folder.getId());return folder;
}

function ensureHeaders_(sheet){
  const headers=['Protocolo','Data/Hora','Nome','E-mail','Instituição/Unidade','Telefone','Perfil','Modalidade','Evento','Check-in','Data/Hora Check-in','Participação','Data/Hora Participação','Certificado Autorizado','Data/Hora Autorização Certificado','Certificado PDF ID','UF da Corporação','Posto / Graduação'];
  if(sheet.getLastRow()===0){sheet.getRange(1,1,1,headers.length).setValues([headers]);sheet.setFrozenRows(1);return;}
  const current=sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),headers.length)).getValues()[0],map={};current.forEach((h,i)=>{if(String(h).trim())map[String(h).trim()]=i+1});
  headers.forEach((h,i)=>{if(!map[h])sheet.getRange(1,i+1).setValue(h)});sheet.setFrozenRows(1);
}
function getSheet_(){const ss=SpreadsheetApp.getActiveSpreadsheet();return ss.getSheetByName(CONFIG.SHEET_NAME)||ss.insertSheet(CONFIG.SHEET_NAME);}
function formatDate_(d){if(!d)return '';if(Object.prototype.toString.call(d)==='[object Date]'&&!isNaN(d))return Utilities.formatDate(d,Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm');return String(d);}
function esc(v){return String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;');}
function sanitize_(v){return String(v||'Participante').replace(/[\\/:*?"<>|]/g,' ').replace(/\s+/g,' ').trim().slice(0,80);}
