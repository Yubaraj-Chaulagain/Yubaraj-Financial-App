/** Yubaraj Finance & Stock — Google Sheets + Apps Script account approval + Gmail notifications.
 * Plain-text passwords are used only because the owner explicitly requested them. This is NOT recommended.
 * Deploy as Web App: Execute as Me. Anyone access is needed for the HTML client; every data action checks a server session.
 */
const APP_NAME='Yubaraj Finance & Stock';
const ADMIN_EMAIL='yubarajchaulagain5@gmail.com';
// Change this before running setupInitialAdmin() for the first time. Do not share this password.
const ADMIN_INITIAL_PASSWORD='ChangeThisAdminPass123!';
const DATA_SHEETS={
 Members:{key:'members',label:'Member',memberField:null,notifyAdmin:true,notifyMemberOnAdd:true,notifyMemberOnChange:false},
 Products:{key:'products',label:'Product',memberField:null,notifyAdmin:true,notifyMemberOnAdd:true,notifyMemberOnChange:true,notifyAllMembers:true},
 Transactions:{key:'tx',label:'Transaction',memberField:'memberId',notifyAdmin:true,notifyMemberOnAdd:true,notifyMemberOnChange:true},
 Savings:{key:'savings',label:'Savings',memberField:'memberId',notifyAdmin:true,notifyMemberOnAdd:true,notifyMemberOnChange:true},
 Loans:{key:'loans',label:'Loan',memberField:'memberId',notifyAdmin:true,notifyMemberOnAdd:true,notifyMemberOnChange:true},
 LoanPayments:{key:'payments',label:'Loan Payment',memberField:'loanId',notifyAdmin:true,notifyMemberOnAdd:true,notifyMemberOnChange:true},
 Transfers:{key:'transfers',label:'Transfer',memberField:'from',notifyAdmin:true,notifyMemberOnAdd:false,notifyMemberOnChange:false}
};
const HEADERS={Users:['id','name','position','gmail','whatsapp','username','password','status','createdAt','approvedAt','lastLogin'],AccountRequests:['id','name','position','gmail','whatsapp','username','password','status','createdAt','reviewedAt','reviewedBy'],Members:['id','ownerEmail','name','company','trade','gmail','whatsapp','staff','opening','date','photoUrl'],Products:['id','ownerEmail','name','desc','qty','buyRate','saleRate'],Transactions:['id','ownerEmail','kind','date','memberId','customerName','productId','productName','qty','rate','total','received','staff','payment'],Savings:['id','ownerEmail','date','memberId','kind','amount','staff','remarks'],Loans:['id','ownerEmail','memberId','date','principal','rate','method','freq','periods','staff'],LoanPayments:['id','ownerEmail','loanId','date','principal','interest','fee','total','staff'],Transfers:['id','ownerEmail','from','to','date','amount','staff','remarks'],ActivityLogs:['id','ownerEmail','time','action','entity','recordId','details'],NotificationLogs:['id','time','status','subject','recipient','details'],Meta:['id','syncedAt','app']};
function doGet(){return json_({ok:true,app:APP_NAME,message:'Apps Script endpoint ready'});}
function doPost(e){
 let p={};try{p=JSON.parse(e&&e.postData&&e.postData.contents||'{}');}catch(err){return json_({ok:false,error:'Invalid JSON'});}
 try{
  if(p.action==='health')return json_({ok:true,app:APP_NAME});
  if(p.action==='signup')return signup_(p);
  if(p.action==='login')return login_(p);
  if(p.action==='forgotPassword')return forgotPassword_(p);
  if(p.action==='resetPassword')return resetPassword_(p);
  if(p.action==='changePassword')return changePassword_(p);
  if(p.action==='logout')return logout_(p);
  if(p.action==='adminListRequests')return adminListRequests_(p);
  if(p.action==='adminReview')return adminReview_(p);
  if(p.action==='adminListUsers')return adminListUsers_(p);
  if(p.action==='adminSetUser')return adminSetUser_(p);
  if(p.action==='sync')return syncData_(p);
  return json_({ok:false,error:'Unknown action'});
 }catch(err){try{logNotification_('ERROR','Backend error','',String(err));}catch(_){}return json_({ok:false,error:String(err)});}
}
function setupInitialAdmin(){const ss=getOrCreateSpreadsheet_(),rows=readRows_(ss.getSheetByName('Users'));if(rows.some(u=>String(u.gmail).toLowerCase()===ADMIN_EMAIL.toLowerCase()))throw new Error('Admin already exists; setup stopped.');upsertRow_(ss,'Users',{id:Utilities.getUuid(),name:'Yubaraj Admin',position:'Super Admin',gmail:ADMIN_EMAIL,whatsapp:'',username:'admin',password:ADMIN_INITIAL_PASSWORD,status:'Approved',createdAt:new Date().toISOString(),approvedAt:new Date().toISOString(),lastLogin:''});Logger.log('Admin created. Login username: admin. Change ADMIN_INITIAL_PASSWORD in script and run only once.');}
function signup_(p){
 const name=clean_(p.name,100),position=clean_(p.position,80),gmail=clean_(p.gmail,180).toLowerCase(),whatsapp=clean_(p.whatsapp,40),username=clean_(p.username,40),password=String(p.password||'');
 if(!name||!position||!validEmail_(gmail)||!whatsapp||!/^[a-zA-Z0-9_.-]{3,40}$/.test(username)||password.length<6)return json_({ok:false,error:'सबै विवरण भर्नुहोस्। Username 3–40 अक्षर र Password कम्तीमा 6 अक्षर हुनुपर्छ।'});
 const ss=getOrCreateSpreadsheet_(),users=readRows_(ss.getSheetByName('Users')),reqs=readRows_(ss.getSheetByName('AccountRequests'));
 if(users.some(u=>String(u.username).toLowerCase()===username.toLowerCase()||String(u.gmail).toLowerCase()===gmail)||reqs.some(u=>String(u.username).toLowerCase()===username.toLowerCase()&&String(u.status)==='Pending'))return json_({ok:false,error:'यो Username वा Gmail पहिले नै प्रयोग भएको छ।'});
 const row={id:Utilities.getUuid(),name,position,gmail,whatsapp,username,password,status:'Pending',createdAt:new Date().toISOString(),reviewedAt:'',reviewedBy:''};
 if(gmail===ADMIN_EMAIL)return json_({ok:false,error:'Admin account सुरक्षित रूपमा initialize गर्न Apps Script मा setupInitialAdmin() एकपटक Run गर्नुहोस्।'});
 appendObjects_(ss,'AccountRequests',[row]);sendMail_(ADMIN_EMAIL,'Yubaraj Finance & Stock — New Account Request','New account request\nName: '+name+'\nPosition: '+position+'\nGmail: '+gmail+'\nWhatsApp: '+whatsapp+'\nUsername: '+username+'\nRequest ID: '+row.id+'\n\nOpen the app Admin Panel to approve or reject.');
 return json_({ok:true,status:'Pending',message:'Request पठाइयो। Admin approval कुर्नुहोस्।'});
}
function login_(p){const username=clean_(p.username,40).toLowerCase(),password=String(p.password||'');if(!username||!password)return json_({ok:false,error:'Username र Password लेख्नुहोस्।'});
 const ss=getOrCreateSpreadsheet_(),users=readRows_(ss.getSheetByName('Users'));let u=users.find(x=>String(x.username||'').toLowerCase()===username&&String(x.password||'')===password);
 if(!u){const rq=readRows_(ss.getSheetByName('AccountRequests')).find(x=>String(x.username||'').toLowerCase()===username&&String(x.password||'')===password);if(rq)return json_({ok:false,error:'तपाईंको Account '+rq.status+' छ। Admin approval आवश्यक छ।'});return json_({ok:false,error:'Username वा Password मिलेन।'});}
 if(u.status!=='Approved')return json_({ok:false,error:'Account '+u.status+' छ। Admin लाई सम्पर्क गर्नुहोस्।'});
 const token=Utilities.getUuid()+'-'+Utilities.getUuid(),exp=Date.now()+1000*60*60*24*14;PropertiesService.getScriptProperties().setProperty('SESSION_'+token,JSON.stringify({email:u.gmail,username:u.username,role:u.gmail===ADMIN_EMAIL?'Admin':(/manager/i.test(String(u.position||''))?'Manager':'Member'),exp}));u.lastLogin=new Date().toISOString();upsertRow_(ss,'Users',u);
 return json_({ok:true,token,user:{name:u.name,position:u.position,gmail:u.gmail,whatsapp:u.whatsapp,username:u.username,role:u.gmail===ADMIN_EMAIL?'Admin':(/manager/i.test(String(u.position||''))?'Manager':'Member')},message:'Login सफल भयो।'});
}

function forgotPassword_(p){
 const username=clean_(p.username,40).toLowerCase(),gmail=clean_(p.gmail,180).toLowerCase();
 if(!username||!validEmail_(gmail))return json_({ok:false,error:'Username र Gmail सही लेख्नुहोस्।'});
 const ss=getOrCreateSpreadsheet_(),u=readRows_(ss.getSheetByName('Users')).find(x=>String(x.username||'').toLowerCase()===username&&String(x.gmail||'').toLowerCase()===gmail&&x.status==='Approved');
 if(!u)return json_({ok:false,error:'Approved account सँग Username/Gmail मिलेन।'});
 const code=String(Math.floor(100000+Math.random()*900000)),key='RESET_'+username;
 PropertiesService.getScriptProperties().setProperty(key,JSON.stringify({code,gmail,exp:Date.now()+10*60*1000,tries:0}));
 sendMail_(gmail,APP_NAME+' — Password Reset Code','Your password reset verification code is: '+code+'\nThis code expires in 10 minutes. If you did not request this, ignore this email.');
 return json_({ok:true,message:'Verification code तपाईंको Gmail मा पठाइयो। 10 मिनेटभित्र प्रयोग गर्नुहोस्।'});
}
function resetPassword_(p){
 const username=clean_(p.username,40).toLowerCase(),gmail=clean_(p.gmail,180).toLowerCase(),code=clean_(p.code,10),password=String(p.password||'');
 if(password.length<6)return json_({ok:false,error:'नयाँ Password कम्तीमा 6 अक्षर हुनुपर्छ।'});
 const key='RESET_'+username,props=PropertiesService.getScriptProperties(),raw=props.getProperty(key);if(!raw)return json_({ok:false,error:'Code छैन वा म्याद सकियो। फेरि code माग्नुहोस्।'});
 let v;try{v=JSON.parse(raw)}catch(e){props.deleteProperty(key);return json_({ok:false,error:'Reset code invalid छ।'});}
 if(v.exp<Date.now()||v.gmail!==gmail){props.deleteProperty(key);return json_({ok:false,error:'Code को म्याद सकियो वा Gmail मिलेन।'});}
 if(v.code!==code){v.tries=(v.tries||0)+1;if(v.tries>=5)props.deleteProperty(key);else props.setProperty(key,JSON.stringify(v));return json_({ok:false,error:'Verification code मिलेन।'});}
 const ss=getOrCreateSpreadsheet_(),u=readRows_(ss.getSheetByName('Users')).find(x=>String(x.username||'').toLowerCase()===username&&String(x.gmail||'').toLowerCase()===gmail&&x.status==='Approved');if(!u)return json_({ok:false,error:'Account भेटिएन।'});
 u.password=password;upsertRow_(ss,'Users',u);props.deleteProperty(key);sendMail_(gmail,APP_NAME+' — Password Changed','Your account password has been reset successfully.');return json_({ok:true,message:'Password सफलतापूर्वक reset भयो। अब Login गर्नुहोस्।'});
}
function changePassword_(p){
 const a=auth_(p.token,false);if(!a.ok)return json_(a);const old=String(p.oldPassword||''),password=String(p.newPassword||'');if(password.length<6)return json_({ok:false,error:'नयाँ Password कम्तीमा 6 अक्षर हुनुपर्छ।'});
 const ss=getOrCreateSpreadsheet_(),u=readRows_(ss.getSheetByName('Users')).find(x=>String(x.gmail||'').toLowerCase()===a.email.toLowerCase());if(!u||String(u.password)!==old)return json_({ok:false,error:'हालको Password मिलेन।'});if(old===password)return json_({ok:false,error:'नयाँ Password पुरानोभन्दा फरक राख्नुहोस्।'});u.password=password;upsertRow_(ss,'Users',u);sendMail_(u.gmail,APP_NAME+' — Password Changed','Your account password was changed successfully.');return json_({ok:true,message:'Password परिवर्तन भयो।'});
}

function logout_(p){if(p.token)PropertiesService.getScriptProperties().deleteProperty('SESSION_'+p.token);return json_({ok:true});}
function adminListRequests_(p){const a=auth_(p.token,true);if(!a.ok)return json_(a);return json_({ok:true,requests:readRows_(getOrCreateSpreadsheet_().getSheetByName('AccountRequests')).filter(r=>r.status==='Pending').map(publicUser_)});}
function adminReview_(p){const a=auth_(p.token,true);if(!a.ok)return json_(a);const ss=getOrCreateSpreadsheet_(),sh=ss.getSheetByName('AccountRequests'),rows=readRows_(sh),r=rows.find(x=>String(x.id)===String(p.id));if(!r)return json_({ok:false,error:'Request भेटिएन।'});const decision=p.decision==='Approve'?'Approved':'Rejected';r.status=decision;r.reviewedAt=new Date().toISOString();r.reviewedBy=a.email;
 if(decision==='Approved'){
  r.approvedAt=r.reviewedAt;
  // Keep the approved login account in Users, and also add it to the app's Members sheet.
  upsertRow_(ss,'Users',r);
  const memberSheet=ss.getSheetByName('Members');
  if(!memberSheet)throw new Error('Members sheet बनाउन सकिएन।');
  const existingMember=readRows_(memberSheet).find(m=>String(m.gmail||'').trim().toLowerCase()===String(r.gmail||'').trim().toLowerCase() || String(m.id||'')===String(r.id));
  const memberRow={id:existingMember&&existingMember.id?existingMember.id:r.id,ownerEmail:existingMember&&existingMember.ownerEmail?existingMember.ownerEmail:ADMIN_EMAIL,name:r.name,company:existingMember&&existingMember.company?existingMember.company:'',trade:r.position||'',gmail:r.gmail,whatsapp:r.whatsapp||'',staff:r.username||'',opening:existingMember&&existingMember.opening!==undefined&&existingMember.opening!==''?existingMember.opening:0,date:existingMember&&existingMember.date?existingMember.date:(r.approvedAt||new Date().toISOString()),photoUrl:existingMember&&existingMember.photoUrl?existingMember.photoUrl:''};
  upsertRow_(ss,'Members',memberRow);
  sendMail_(r.gmail,'Yubaraj Financial App — Account Approved','Hello '+r.name+', your account has been approved. You can now log in.');
 }else sendMail_(r.gmail,'Yubaraj Financial App — Account Rejected','Hello '+r.name+', your account request was rejected. Contact the administrator if you need help.');
 upsertRow_(ss,'AccountRequests',r);return json_({ok:true,message:'Account '+decision+(decision==='Approved'?' and added to Members sheet':'' )});}
function adminListUsers_(p){const a=auth_(p.token,true);if(!a.ok)return json_(a);return json_({ok:true,users:readRows_(getOrCreateSpreadsheet_().getSheetByName('Users')).map(publicUser_)});}
function adminSetUser_(p){const a=auth_(p.token,true);if(!a.ok)return json_(a);const ss=getOrCreateSpreadsheet_(),rows=readRows_(ss.getSheetByName('Users')),u=rows.find(x=>String(x.id)===String(p.id));if(!u)return json_({ok:false,error:'User भेटिएन।'});if(u.gmail===ADMIN_EMAIL)return json_({ok:false,error:'Admin account परिवर्तन गर्न मिल्दैन।'});if(p.status==='Delete'){deleteById_(ss.getSheetByName('Users'),u.id);sendMail_(u.gmail,'Yubaraj Financial App — Account Deleted','Your account was deleted by the administrator.');}else{u.status=['Approved','Blocked','Rejected'].includes(p.status)?p.status:'Blocked';upsertRow_(ss,'Users',u);sendMail_(u.gmail,'Yubaraj Financial App — Account Status','Your account status is now: '+u.status);}return json_({ok:true});}
function syncData_(p){const a=auth_(p.token,false);if(!a.ok)return json_(a);const ss=getOrCreateSpreadsheet_(),data=p.data||{},owner=a.email;const lock=LockService.getScriptLock();lock.waitLock(25000);const activity=[];
 try{Object.keys(DATA_SHEETS).forEach(n=>{const cfg=DATA_SHEETS[n],arr=Array.isArray(data[cfg.key])?data[cfg.key]:[];const rows=arr.map(x=>Object.assign({},x,{ownerEmail:owner}));const changes=reconcileOwnerSheet_(ss,n,rows,owner);changes.forEach(c=>{const desc=describeChange_(cfg,c);activity.push({id:Utilities.getUuid(),ownerEmail:owner,time:new Date().toISOString(),action:c.action,entity:cfg.label,recordId:c.record.id||'',details:desc});sendChangeNotification_(ss,cfg,c,data,desc,owner);});});if(activity.length)appendObjects_(ss,'ActivityLogs',activity);upsertRow_(ss,'Meta',{id:'last-sync-'+owner,syncedAt:new Date().toISOString(),app:APP_NAME});}finally{lock.releaseLock();}
 return json_({ok:true,spreadsheetId:ss.getId(),spreadsheetUrl:ss.getUrl(),changes:activity.length,message:'Sync completed'});
}
function auth_(token,adminOnly){if(!token)return {ok:false,error:'Login required'};const raw=PropertiesService.getScriptProperties().getProperty('SESSION_'+token);if(!raw)return {ok:false,error:'Session expired. Please log in again.'};let s;try{s=JSON.parse(raw)}catch(e){return {ok:false,error:'Invalid session'};}if(s.exp<Date.now()){PropertiesService.getScriptProperties().deleteProperty('SESSION_'+token);return {ok:false,error:'Session expired. Please log in again.'};}const u=readRows_(getOrCreateSpreadsheet_().getSheetByName('Users')).find(x=>String(x.gmail).toLowerCase()===String(s.email).toLowerCase());if(!u||u.status!=='Approved')return {ok:false,error:'Account is not approved'};const managerAllowed=/manager/i.test(String(u.position||''));if(adminOnly&&s.email!==ADMIN_EMAIL&&!managerAllowed)return {ok:false,error:'Admin or Manager permission required'};return {ok:true,email:s.email,role:s.email===ADMIN_EMAIL?'Admin':(managerAllowed?'Manager':'Member')};}
function getOrCreateSpreadsheet_(){const props=PropertiesService.getScriptProperties();let ss,id=props.getProperty('SPREADSHEET_ID');if(id){try{ss=SpreadsheetApp.openById(id)}catch(e){}}if(!ss){ss=SpreadsheetApp.create(APP_NAME+' Data');props.setProperty('SPREADSHEET_ID',ss.getId());}Object.keys(HEADERS).forEach(n=>{let sh=ss.getSheetByName(n);if(!sh)sh=ss.insertSheet(n);ensureHeaders_(sh,HEADERS[n]);});const first=ss.getSheetByName('Sheet1');if(first&&ss.getSheets().length>1&&first.getLastRow()===0)ss.deleteSheet(first);return ss;}
function ensureHeaders_(sh,expected){let old=sh.getLastColumn()?sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].filter(String):[];if(!old.length){sh.getRange(1,1,1,expected.length).setValues([expected]);sh.setFrozenRows(1);return expected;}const merged=old.concat(expected.filter(x=>!old.includes(x)));if(merged.length!==old.length)sh.getRange(1,1,1,merged.length).setValues([merged]);sh.setFrozenRows(1);return merged;}
function readRows_(sh){const v=sh.getDataRange().getValues();if(v.length<2)return [];const h=v[0].map(String);return v.slice(1).filter(r=>r.some(x=>x!==''&&x!==null)).map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);return o;});}
function reconcileOwnerSheet_(ss,name,incoming,owner){const sh=ss.getSheetByName(name),headers=ensureHeaders_(sh,['id','ownerEmail']);const all=readRows_(sh),mine=all.filter(r=>String(r.ownerEmail||'')===owner),oldMap=new Map(mine.map(r=>[String(r.id),r])),newMap=new Map(incoming.filter(r=>r&&r.id).map(r=>[String(r.id),r])),changes=[];newMap.forEach((r,id)=>{const old=oldMap.get(id);if(!old)changes.push({action:'Add',record:r,old:null});else if(stable_(old)!==stable_(r))changes.push({action:'Update',record:r,old});});mine.forEach(old=>{if(!newMap.has(String(old.id)))changes.push({action:'Delete',record:old,old});});const byId=new Map(all.map(r=>[String(r.id),r]));mine.forEach(r=>byId.delete(String(r.id)));incoming.forEach(r=>byId.set(String(r.id),r));const final=[...byId.values()];const hs=ensureHeaders_(sh,['id','ownerEmail',...new Set(final.flatMap(o=>Object.keys(o)))]);if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();if(final.length)sh.getRange(2,1,final.length,hs.length).setValues(final.map(o=>hs.map(h=>serial_(o[h]))));return changes;}
function appendObjects_(ss,name,objects){if(!objects.length)return;const sh=ss.getSheetByName(name),keys=[...new Set(ensureHeaders_(sh,['id']).concat(...objects.map(o=>Object.keys(o))))],hs=ensureHeaders_(sh,keys);sh.getRange(sh.getLastRow()+1,1,objects.length,hs.length).setValues(objects.map(o=>hs.map(h=>serial_(o[h]))));}
function upsertRow_(ss,name,obj){const sh=ss.getSheetByName(name);if(!sh)throw new Error('Sheet not found: '+name);const rows=readRows_(sh),idx=rows.findIndex(r=>String(r.id)===String(obj.id));const hs=ensureHeaders_(sh,[...new Set(Object.keys(obj))]);const row=hs.map(h=>serial_(obj[h]));if(idx<0){sh.getRange(sh.getLastRow()+1,1,1,hs.length).setValues([row]);}else{sh.getRange(idx+2,1,1,hs.length).setValues([row]);}}
function deleteById_(sh,id){const rows=readRows_(sh),idx=rows.findIndex(r=>String(r.id)===String(id));if(idx>=0)sh.deleteRow(idx+2);}
function sendChangeNotification_(ss,cfg,c,data,desc,owner){const recipients=new Set([ADMIN_EMAIL]);const members=data.members||[];const byId=new Map(members.map(m=>[String(m.id),m]));let member=null;if(cfg.memberField==='memberId')member=byId.get(String(c.record.memberId||''));if(cfg.memberField==='loanId'){const l=(data.loans||[]).find(x=>String(x.id)===String(c.record.loanId));if(l)member=byId.get(String(l.memberId||''));}if(cfg.memberField==='from')member=byId.get(String(c.record.from||''));if(cfg.notifyAllMembers)members.forEach(m=>{if(validEmail_(m.gmail))recipients.add(String(m.gmail).trim());});if(c.action==='Add'&&cfg.notifyMemberOnAdd&&member&&validEmail_(member.gmail))recipients.add(member.gmail.trim());if(c.action!=='Add'&&cfg.notifyMemberOnChange&&member&&validEmail_(member.gmail))recipients.add(member.gmail.trim());if(cfg.key==='Members'&&c.action==='Add'&&validEmail_(c.record.gmail))recipients.add(c.record.gmail.trim());if(cfg.key==='Members'&&c.action!=='Add'){recipients.clear();recipients.add(ADMIN_EMAIL);}const subject=APP_NAME+' — '+c.action+' '+cfg.label;const body='Owner: '+owner+'\nAction: '+c.action+'\nType: '+cfg.label+'\nRecord ID: '+(c.record.id||'')+'\nDetails: '+desc+'\nDate/time: '+new Date().toISOString();recipients.forEach(to=>sendMail_(to,subject,body));}
function sendMail_(to,subject,body){try{MailApp.sendEmail({to,subject,body,name:APP_NAME});logNotification_('SENT',subject,to,'Sent');}catch(e){logNotification_('FAILED',subject,to,String(e));}}
function logNotification_(status,subject,recipient,details){try{const ss=getOrCreateSpreadsheet_();appendObjects_(ss,'NotificationLogs',[{id:Utilities.getUuid(),time:new Date().toISOString(),status,subject,recipient,details}]);}catch(e){}}
function describeChange_(cfg,c){const r=c.record||{},parts=[];['name','kind','date','memberId','customerName','productName','qty','amount','principal','interest','fee','total','gmail'].forEach(k=>{if(r[k]!==undefined&&r[k]!==''&&r[k]!==null)parts.push(k+': '+String(r[k]).slice(0,120));});return parts.join(' | ')||cfg.label+' '+c.action.toLowerCase();}
function publicUser_(u){return {id:u.id,name:u.name,position:u.position,gmail:u.gmail,whatsapp:u.whatsapp,username:u.username,status:u.status,createdAt:u.createdAt};}
function clean_(v,n){return String(v||'').trim().slice(0,n);}function validEmail_(s){return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s||'').trim());}function stable_(o){const x={};Object.keys(o||{}).sort().forEach(k=>{if(k!=='photo'&&k!=='photoUploadPending')x[k]=serial_(o[k]);});return JSON.stringify(x);}function serial_(v){return v===undefined||v===null?'':typeof v==='object'?JSON.stringify(v):v;}function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
