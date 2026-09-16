/**
 * Rudraksha Certificate Studio - Google Drive upload backend
 *
 * The desktop app uploads PDFs here. The private registry remains private.
 * A separate "Public Verification" tab contains only Certificate No., Drive URL,
 * and Updated At. Publish ONLY that tab to the web so the GitHub verification page
 * can show the original PDF after cryptographic verification.
 *
 * IMPORTANT: replace SECRET with your current local upload secret if this script
 * is being installed fresh. Do not publish this file with a real secret.
 */
const SECRET = 'REPLACE_WITH_YOUR_UPLOAD_SECRET';
const UPLOAD_SECRET = SECRET;
const DEFAULT_FOLDER_NAME = 'Rudraksha Certificates';
const DEFAULT_SHEET_NAME = 'Rudraksha Certificate Registry';

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cert = String(p.cert || '');
  if (!cert) return HtmlService.createHtmlOutput('Certificate verification service is online.');
  return HtmlService.createHtmlOutput('This service is for certificate uploads. Please use the public verification page.');
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return json_({ok:false,error:'Empty request'});
    const body = JSON.parse(e.postData.contents);
    if (body.action !== 'upload') return json_({ok:false,error:'Unknown action'});
    if (!constantTimeEqual_(String(body.upload_secret || ''), String(UPLOAD_SECRET))) return json_({ok:false,error:'Unauthorized upload request'});
    if (!body.cert_no || !body.pdf_base64) return json_({ok:false,error:'Missing certificate number or PDF data'});

    const folder = getOrCreateFolder_();
    const bytes = Utilities.base64Decode(body.pdf_base64);
    const blob = Utilities.newBlob(bytes, body.mime_type || 'application/pdf', body.filename || (body.cert_no + '.pdf'));
    const existing = folder.getFilesByName(blob.getName());
    while (existing.hasNext()) { try { existing.next().setTrashed(true); } catch (_) {} }
    const file = folder.createFile(blob);
    let sharing = 'restricted';
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); sharing='anyone_with_link'; }
    catch (err) { sharing='restricted: '+err.message; }

    const record = {
      cert:String(body.cert_no), lot:String(body.lot_no || ''), issued:String(body.issued || ''),
      mukh:String(body.mukh || ''), type:String(body.specimen_type || ''), template:String(body.template || ''),
      created_at:String(body.created_at || new Date().toISOString()), file_id:file.getId(),
      file_url:'https://drive.google.com/file/d/'+file.getId()+'/view', sharing:sharing
    };
    upsertRecord_(record);
    upsertPublicVerificationRecord_(record);
    return json_({ok:true,file_id:record.file_id,view_url:record.file_url,sharing:sharing});
  } catch (err) { return json_({ok:false,error:String(err && err.message || err)}); }
}

function getOrCreateFolder_() {
  const props=PropertiesService.getScriptProperties(); const id=props.getProperty('DRIVE_FOLDER_ID');
  if(id){try{return DriveApp.getFolderById(id)}catch(_) {}}
  const fs=DriveApp.getFoldersByName(DEFAULT_FOLDER_NAME); const f=fs.hasNext()?fs.next():DriveApp.createFolder(DEFAULT_FOLDER_NAME);
  props.setProperty('DRIVE_FOLDER_ID',f.getId()); return f;
}
function getOrCreateSheet_() {
  const props=PropertiesService.getScriptProperties(); const id=props.getProperty('REGISTRY_SHEET_ID');
  if(id){try{return SpreadsheetApp.openById(id).getSheets()[0]}catch(_) {}}
  const ss=SpreadsheetApp.create(DEFAULT_SHEET_NAME); const sh=ss.getSheets()[0];
  sh.getRange(1,1,1,9).setValues([['Certificate No.','Lot No.','Issued To','Mukh','Type','Template','Created At','Drive File ID','Drive URL']]); sh.setFrozenRows(1);
  props.setProperty('REGISTRY_SHEET_ID',ss.getId()); return sh;
}
function upsertRecord_(r){
  const sh=getOrCreateSheet_(),last=sh.getLastRow();
  if(last>=2){const vals=sh.getRange(2,1,last-1,1).getValues();for(let i=0;i<vals.length;i++)if(String(vals[i][0])===r.cert){sh.getRange(i+2,1,1,9).setValues([[r.cert,r.lot,r.issued,r.mukh,r.type,r.template,r.created_at,r.file_id,r.file_url]]);return;}}
  sh.appendRow([r.cert,r.lot,r.issued,r.mukh,r.type,r.template,r.created_at,r.file_id,r.file_url]);
}
function getOrCreatePublicVerificationSheet_(){
  const sh=getOrCreateSheet_(),ss=sh.getParent(); let pub=ss.getSheetByName('Public Verification');
  if(!pub){pub=ss.insertSheet('Public Verification');pub.getRange(1,1,1,3).setValues([['Certificate No.','Drive URL','Updated At']]);pub.setFrozenRows(1);}
  return pub;
}
function upsertPublicVerificationRecord_(r){
  const sh=getOrCreatePublicVerificationSheet_(),last=sh.getLastRow();
  if(last>=2){const vals=sh.getRange(2,1,last-1,1).getValues();for(let i=0;i<vals.length;i++)if(String(vals[i][0])===r.cert){sh.getRange(i+2,1,1,3).setValues([[r.cert,r.file_url,new Date().toISOString()]]);return;}}
  sh.appendRow([r.cert,r.file_url,new Date().toISOString()]);
}
function constantTimeEqual_(a,b){a=String(a||'');b=String(b||'');if(a.length!==b.length)return false;let x=0;for(let i=0;i<a.length;i++)x|=a.charCodeAt(i)^b.charCodeAt(i);return x===0;}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
