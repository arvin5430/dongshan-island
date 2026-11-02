// 只使用虛構資料；不讀正式學生檔案。執行：node auth.test.cjs
const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict'),crypto=require('node:crypto');
const header=['班級','座號','姓名','學號','EIP帳號','EIP密碼'];
const data={source:{'101':[header,['101','01','測試甲','0001','00123','000789']],'302':[header,['302','02','測試乙','0002','test302','pass302']],'轉出資料':[header,['101','03','已轉出','0003','gone','gone']]},progress:{Progress:[['studentId','completedIslands','updatedAt']]}};
const reads=[],cache=new Map();
function sheet(rows){return {getDataRange:()=>({getDisplayValues:()=>rows.map(r=>r.map(String))}),getLastRow:()=>rows.length,getRange:(r,c,n,m)=>({setValues:v=>{rows[r-1]=v[0]}})};}
const context={console,Date,PropertiesService:{getScriptProperties:()=>({getProperty:k=>({STUDENT_SHEET_ID:'source',PROGRESS_SHEET_ID:'progress'}[k])})},Utilities:{DigestAlgorithm:{SHA_256:'sha256'},computeDigest:(algo,s)=>Array.from(crypto.createHash('sha256').update(s).digest()),getUuid:()=>crypto.randomUUID()},CacheService:{getScriptCache:()=>({get:k=>cache.get(k)||null,put:(k,v)=>cache.set(k,v),remove:k=>cache.delete(k)})},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},SpreadsheetApp:{openById:id=>({getSheetByName:name=>{reads.push(name);return data[id][name]?sheet(data[id][name]):null}})}};
vm.createContext(context);vm.runInContext(fs.readFileSync(__dirname+'/Code.gs','utf8'),context);
assert.equal(context.login('轉出資料','gone','gone').ok,false);
assert.equal(context.login('302','00123','000789').ok,false);
assert.equal(context.login('101','00123','789').ok,false);
const r=context.login('101','00123','000789');assert.equal(r.ok,true);assert.equal(r.token.length,72);assert.equal(r.user.name,'測試甲');assert(!JSON.stringify(r).includes('000789'));assert(!JSON.stringify(r).includes('00123'));
assert.equal(context.completeTask(r.token,0,[0,0]).ok,false);
assert.equal(context.completeTask(r.token,0,[0,1]).earned.length,1);
assert.equal(context.completeTask(r.token,0,[0,1]).earned.length,1);
const again=context.login('101','00123','000789');assert.equal(again.earned.length,1);
const other=context.login('302','test302','pass302');assert.equal(other.earned.length,0);
context.logout(r.token);assert.throws(()=>context.completeTask(r.token,1,[1,2]));
for(let i=0;i<5;i++)assert.equal(context.login('101','missing','wrong').ok,false);
assert.match(context.login('101','missing','wrong').message,/15/);
data.source['101'].push(['101','04','重複帳號','0004','00123','000789']);assert.equal(context.login('101','00123','000789').ok,false);
assert(!reads.includes('轉出資料'));
assert.throws(()=>context.completeTask('forged',0,[0,1]));
console.log('PASS: class isolation, leading zeros, generic errors, no credential response, duplicate account rejection, per-user persistence, idempotence, server answer validation, rate limiting, logout and forged token rejection.');
