const APP_NAME = 'Yubaraj Finance & Stock';
const TAB_NAMES = ['Members','Products','Transactions','Savings','Loans','LoanPayments','Transfers','Meta'];
function doGet(){return ContentService.createTextOutput(JSON.stringify({ok:true,app:APP_NAME,message:'Apps Script endpoint is ready'})).setMimeType(ContentService.MimeType.JSON);}
function doPost(e){
  try {
    const payload=JSON.parse((e.postData&&e.postData.contents)||'{}');
    const data=payload.data||{};
    const ss=getOrCreateSpreadsheet_();
    const mapping={
      Members:data.members||[], Products:data.products||[], Transactions:data.tx||[],
      Savings:data.savings||[], Loans:data.loans||[], LoanPayments:data.payments||[],
      Transfers:data.transfers||[], Meta:[{id:'last-sync',syncedAt:new Date().toISOString(),app:APP_NAME}]
    };
    const lock=LockService.getScriptLock(); lock.waitLock(20000);
    try { Object.keys(mapping).forEach(name=>upsertRows_(ss,name,mapping[name])); }
    finally { lock.releaseLock(); }
    return json_({ok:true,spreadsheetId:ss.getId(),spreadsheetUrl:ss.getUrl(),message:'Sync completed'});
  } catch(err) { return json_({ok:false,error:String(err)}); }
}
function getOrCreateSpreadsheet_(){
  const props=PropertiesService.getScriptProperties();
  let id=props.getProperty('SPREADSHEET_ID'),ss;
  if(id){try{ss=SpreadsheetApp.openById(id);}catch(e){}}
  if(!ss){ss=SpreadsheetApp.create(APP_NAME+' Data');props.setProperty('SPREADSHEET_ID',ss.getId());}
  TAB_NAMES.forEach(name=>{if(!ss.getSheetByName(name))ss.insertSheet(name);});
  const first=ss.getSheets()[0]; if(first.getName()==='Sheet1' && ss.getSheetByName('Meta'))ss.deleteSheet(first);
  return ss;
}
function upsertRows_(ss,name,records){
  const sh=ss.getSheetByName(name)||ss.insertSheet(name);
  const clean=(records||[]).filter(r=>r&&r.id);
  let existing=sh.getDataRange().getValues(),headers=existing.length?existing[0].map(String):[];
  const allKeys=[...new Set(clean.flatMap(o=>Object.keys(o)))];
  if(name==='Meta')allKeys.push(...[]);
  if(!headers.length){headers=['id',...allKeys.filter(k=>k!=='id')];sh.getRange(1,1,1,headers.length).setValues([headers]);}
  else {
    const newHeaders=allKeys.filter(k=>!headers.includes(k));
    if(newHeaders.length){headers=headers.concat(newHeaders);sh.getRange(1,1,1,headers.length).setValues([headers]);}
  }
  if(headers.length===0)return;
  const idCol=headers.indexOf('id');
  const rowById=new Map();
  if(existing.length>1){for(let r=1;r<existing.length;r++){const id=String(existing[r][idCol]||'');if(id)rowById.set(id,r+1);}}
  clean.forEach(obj=>{
    const row=headers.map(h=>serial_(obj[h]));
    const target=rowById.get(String(obj.id));
    if(target)sh.getRange(target,1,1,headers.length).setValues([row]);
    else {sh.appendRow(row);rowById.set(String(obj.id),sh.getLastRow());}
  });
  sh.setFrozenRows(1);
}
function serial_(v){if(v===undefined||v===null)return '';return typeof v==='object'?JSON.stringify(v):v;}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
