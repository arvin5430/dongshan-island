// 在「專案設定 → 指令碼屬性」設定 STUDENT_SHEET_ID、PROGRESS_SHEET_ID。
// 學生來源表保留私人；遊戲進度使用另一份空白試算表。禁止公開來源資料。
const CLASSES_ = ['101','102','201','202','301','302','401','402','501','502','601','602'];
const ANSWERS_ = [[0,1],[1,2],[1,0],[1,0]];
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('東山學習冒險島').addMetaTag('viewport','width=device-width, initial-scale=1');
}
function settings_() {
  const p=PropertiesService.getScriptProperties();
  const source=p.getProperty('STUDENT_SHEET_ID'), progress=p.getProperty('PROGRESS_SHEET_ID');
  if(!source||!progress||source===progress)throw new Error('請先設定兩份不同的私人試算表。');
  return {source:source,progress:progress};
}
function hash_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(value)).map(b=>('0'+(b&255).toString(16)).slice(-2)).join('');
}
function fail_() { return {ok:false,message:'登入失敗，請確認班級、帳號及密碼。'}; }
function login(classCode,account,password) {
  if(typeof classCode!=='string'||CLASSES_.indexOf(classCode)<0||typeof account!=='string'||typeof password!=='string'||!account.trim()||!password||account.length>100||password.length>128)return fail_();
  account=account.trim();
  const config=settings_(),cache=CacheService.getScriptCache(),key='attempt:'+hash_(classCode+'|'+account),lock=LockService.getScriptLock();
  // 每帳號每 15 分鐘最多 5 次失敗；使用鎖避免併發繞過。
  lock.waitLock(10000);
  try {
    let rate=JSON.parse(cache.get(key)||'{"count":0,"until":0}');
    if(Date.now()>=rate.until)rate={count:0,until:Date.now()+900000};
    if(rate.count>=5)return {ok:false,message:'嘗試次數較多，請 15 分鐘後再試，或請老師協助。'};
    rate.count++;cache.put(key,JSON.stringify(rate),900);
  } finally {lock.releaseLock();}
  const sheet=SpreadsheetApp.openById(config.source).getSheetByName(classCode);
  if(!sheet)return fail_();
  const rows=sheet.getDataRange().getDisplayValues(),headers=rows[0]||[];
  const column=label=>headers.findIndex(v=>String(v).trim()===label);
  const ix={account:column('EIP帳號'),password:column('EIP密碼'),name:column('姓名'),seat:column('座號'),id:column('學號')};
  if(Object.keys(ix).some(k=>ix[k]<0))throw new Error('班級分頁缺少必要欄位。');
  // 僅接受選定班級內唯一帳號；不跨班搜尋，也不使用其他工作表。
  const found=rows.slice(1).filter(r=>String(r[ix.account]).trim()===account);
  if(found.length!==1||String(found[0][ix.password])!==password)return fail_();
  const row=found[0],studentId=String(row[ix.id]).trim();
  if(!studentId)return {ok:false,message:'學生資料缺少學號，請老師協助。'};
  const user={id:hash_('student:'+studentId),name:String(row[ix.name]),classCode:classCode,seat:String(row[ix.seat])};
  // 只把單一學生的必要資料與短效 session 留在伺服器，絕不回傳密碼。
  const token=Utilities.getUuid()+Utilities.getUuid();
  const earned=progress_(user.id);
  cache.put('session:'+hash_(token),JSON.stringify({user:user,expires:Date.now()+3600000}),3600);
  cache.remove(key);
  return {ok:true,user:{name:user.name,classCode:user.classCode,seat:user.seat},token:token,earned:earned};
}
function session_(token) {
  if(typeof token!=='string'||token.length!==72)throw new Error('請重新登入。');
  const raw=CacheService.getScriptCache().get('session:'+hash_(token));
  if(!raw)throw new Error('登入已逾時，請重新登入。');
  const s=JSON.parse(raw);if(Date.now()>s.expires)throw new Error('登入已逾時，請重新登入。');return s.user;
}
function progressSheet_() {
  const sheet=SpreadsheetApp.openById(settings_().progress).getSheetByName('Progress');
  if(!sheet)throw new Error('請老師先執行 setupProgress。');return sheet;
}
function clean_(a) {return Array.isArray(a)?Array.from(new Set(a.filter(n=>Number.isInteger(n)&&n>=0&&n<4))):[];}
function progress_(id) {
  const rows=progressSheet_().getDataRange().getDisplayValues();
  for(let i=1;i<rows.length;i++)if(rows[i][0]===id){try{return clean_(JSON.parse(rows[i][1]));}catch(e){throw new Error('進度資料格式異常，請老師協助。');}}
  return [];
}
function completeTask(token,island,answers) {
  const user=session_(token);
  if(!Number.isInteger(island)||island<0||island>3||!Array.isArray(answers)||answers.length!==2||answers.some((a,i)=>a!==ANSWERS_[island][i]))return {ok:false,message:'請先完成這座島的兩個挑戰。'};
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const sheet=progressSheet_(),rows=sheet.getDataRange().getDisplayValues();let rowNumber=0,earned=[];
    for(let i=1;i<rows.length;i++)if(rows[i][0]===user.id){rowNumber=i+1;earned=clean_(JSON.parse(rows[i][1]));break;}
    if(earned.indexOf(island)<0)earned.push(island);
    const record=[[user.id,JSON.stringify(earned),new Date().toISOString()]];
    sheet.getRange(rowNumber||sheet.getLastRow()+1,1,1,3).setValues(record);
    return {ok:true,earned:earned};
  } finally {lock.releaseLock();}
}
function logout(token) {if(typeof token==='string')CacheService.getScriptCache().remove('session:'+hash_(token));return {ok:true};}
// 僅在 Apps Script 編輯器手動執行；底線結尾函式不能由 google.script.run 呼叫。
function setupProgress_() {
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {const ss=SpreadsheetApp.openById(settings_().progress);let sh=ss.getSheetByName('Progress');if(!sh)sh=ss.insertSheet('Progress');if(sh.getLastRow()===0){sh.appendRow(['studentId','completedIslands','updatedAt']);sh.setFrozenRows(1);}}
  finally{lock.releaseLock();}
}
