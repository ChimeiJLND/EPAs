const assert = require('node:assert/strict');
const {readFileSync} = require('node:fs');
const {createHash, randomUUID} = require('node:crypto');
const vm = require('node:vm');

// 所有資料與服務都是記憶體替身；不會連線到正式網站或Google Sheet。
let clock = Date.now();
let sheetAccesses = 0;
const properties = new Map([['EPA_ACCESS_CODE', randomUUID().slice(0,6)]]);
const initialCode = properties.get('EPA_ACCESS_CODE');
const cache = new Map();
const sheets = {
  '主題設定表': [['主題名稱','情境說明','評估項目1'],['輸血護理 ','測試情境','測試項目']],
  '評核紀錄表': [
    ['紀錄編號','狀態','填寫日期','學員姓名','期程','教師姓名','主題名稱','信賴等級','建立時間','更新時間','可再精進'],
    ['old','正式','2026-11-01','測試學員','三個月','測試教師','輸血護理 ',1,'before','before',''],
    ['draft','草稿','2026-11-01','測試學員甲','三個月','測試教師','輸血護理 ',1,'before','before','']
  ]
};
const cacheService = {
  get(k) {const v=cache.get(k);return v && v.until>clock ? v.value : null;},
  put(k,value,seconds) {cache.set(k,{value,until:clock+seconds*1000});},
  remove(k) {cache.delete(k);}
};
function sheetObject(name) {
  const data=sheets[name];
  return {
    getDataRange:()=>({getValues:()=>data.map(r=>[...r])}),
    getLastColumn:()=>data[0].length,
    getRange:(r,c,n,m)=>({getValues:()=>data.slice(r-1,r-1+n).map(row=>row.slice(c-1,c-1+m)),
      setValues:rows=>rows.forEach((row,i)=>{data[r-1+i]=[...row];})}),
    appendRow:row=>data.push(row)
  };
}
const sandbox = {
  Date: class extends Date {static now(){return clock;}},
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>properties.get(k)||null})},
  CacheService:{getScriptCache:()=>cacheService},
  LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})},
  Utilities:{DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},
    computeDigest:(_,v)=>[...createHash('sha256').update(v).digest()],getUuid:randomUUID,
    formatDate:(_,__,format)=>format.includes('HH')?'2026-11-12 12:34:56':'2026-11-12'},
  ContentService:{MimeType:{JSON:'json'},createTextOutput:text=>({text,setMimeType(){return this;}})},
  SpreadsheetApp:{openById(){sheetAccesses++;return{getSheetByName:sheetObject};},flush(){}}
};
vm.createContext(sandbox);
vm.runInContext(readFileSync('apps-script/Code.gs','utf8'),sandbox);
const post=body=>JSON.parse(sandbox.doPost({postData:{contents:JSON.stringify(body)}}).text);
const get=params=>JSON.parse(sandbox.doGet({parameter:params}).text);
let checks=0;
function check(name,fn){fn(); checks++; console.log('PASS',name);}

check('所有舊GET入口都拒絕匿名存取，含偽造查詢憑證',()=>{
  for(const action of ['getTopics','search','getRecord','submit','login']) {
    assert.equal(get({action,token:'a'.repeat(64),accessCode:initialCode}).code,'AUTH_REQUIRED');
  }
  assert.equal(sheetAccesses,0);
});
check('所有POST資料入口在驗證前均不讀寫試算表',()=>{
  for(const action of ['getTopics','search','getRecord','submit','session','logout']) {
    assert.equal(post({action,data:{'學員姓名':'不得寫入'},token:'bad'}).code,'AUTH_REQUIRED');
  }
  assert.equal(sheetAccesses,0);
});
check('空白、過短或未設定使用碼時不會放行',()=>{
  properties.delete('EPA_ACCESS_CODE');
  assert.equal(post({action:'login',accessCode:''}).code,'AUTH_NOT_CONFIGURED');
  properties.set('EPA_ACCESS_CODE','short');
  assert.equal(post({action:'search'}).code,'AUTH_NOT_CONFIGURED');
  properties.set('EPA_ACCESS_CODE',initialCode);
});
check('錯誤使用碼無法取得憑證',()=>{
  assert.equal(post({action:'login',accessCode:'wrong'}).code,'LOGIN_FAILED');
});
let session;
check('正確使用碼換取有期限憑證，伺服器快取不存明文憑證',()=>{
  session=post({action:'login',accessCode:initialCode});
  assert.match(session.token,/^[a-f0-9]{64}$/i);
  assert.ok(session.expiresAt>clock);
  assert.ok(!JSON.stringify([...cache]).includes(session.token));
  assert.ok(!JSON.stringify([...cache]).includes(initialCode));
});
check('登入後可讀取題庫，主題名稱空白已正規化',()=>{
  assert.equal(post({action:'getTopics',token:session.token})[0].code,'輸血護理');
});
check('主題篩選支援只有名稱的舊紀錄，並保留低等級正式評核',()=>{
  const rows=post({action:'search',token:session.token,params:{topic:' 輸血護理 ',status:'正式'}});
  assert.equal(rows.length,1); assert.equal(rows[0]['信賴等級'],1);
  assert.equal(rows[0]['主題代碼'],'輸血護理');
});
check('精確學員配對不會把名稱相近者合併',()=>{
  assert.equal(post({action:'search',token:session.token,params:{studentExact:'測試學員'}}).length,1);
});
check('草稿可更新、正式紀錄不覆蓋、文字不變成Sheet公式',()=>{
  const draft=post({action:'submit',token:session.token,recordId:'draft',
    data:{'學員姓名':'測試學員甲','狀態':'草稿','主題名稱':'輸血護理','可再精進':'=1+1'}});
  assert.equal(draft.status,'updated');
  assert.equal(sheets['評核紀錄表'][2][10],"'=1+1");
  const created=post({action:'submit',token:session.token,recordId:'old',
    data:{'學員姓名':'測試學員','狀態':'正式','主題名稱':'輸血護理'}});
  assert.equal(created.status,'locked_new');
  assert.equal(sheets['評核紀錄表'][1][8],'before');
});
check('不能用請求資料改寫紀錄編號與建立時間',()=>{
  post({action:'submit',token:session.token,recordId:'draft',
    data:{'學員姓名':'測試學員甲','狀態':'草稿','紀錄編號':'forged','建立時間':'forged'}});
  assert.equal(sheets['評核紀錄表'][2][0],'draft');
  assert.equal(sheets['評核紀錄表'][2][8],'before');
});
check('更改後端使用碼後，既有憑證立即失效',()=>{
  properties.set('EPA_ACCESS_CODE',randomUUID());
  assert.equal(post({action:'session',token:session.token}).code,'AUTH_REQUIRED');
  properties.set('EPA_ACCESS_CODE',initialCode);
});
check('登出後原憑證不能再次讀取',()=>{
  assert.equal(post({action:'logout',token:session.token}).ok,true);
  assert.equal(post({action:'search',token:session.token}).code,'AUTH_REQUIRED');
});
check('過期與快取遺失皆要求重新登入',()=>{
  const a=post({action:'login',accessCode:initialCode});
  clock=a.expiresAt+1;
  assert.equal(post({action:'session',token:a.token}).code,'AUTH_REQUIRED');
  const b=post({action:'login',accessCode:initialCode});
  cache.clear();
  assert.equal(post({action:'session',token:b.token}).code,'AUTH_REQUIRED');
});
check('登入頻率超過限制時不會繼續發憑證',()=>{
  cache.clear();
  for(let i=0;i<30;i++) post({action:'login',accessCode:'wrong'});
  assert.equal(post({action:'login',accessCode:initialCode}).code,'RATE_LIMIT');
  clock+=120000;
});
check('格式錯誤不回傳內部資訊',()=>{
  assert.equal(JSON.parse(sandbox.doPost({postData:{contents:'{'}}).text).code,'BAD_REQUEST');
  assert.equal(post(null).code,'BAD_REQUEST');
});

async function testFrontend() {
  const elements = new Map();
  const classes = new Set();
  const saved = new Map();
  const calls=[];
  let ready=0;
  const element=id=>{
    if(!elements.has(id)) elements.set(id,{value:'',hidden:false,disabled:false,textContent:'',replaceChildren(){this.cleared=true;}});
    return elements.get(id);
  };
  const f={
    API_URL:'https://example.invalid/exec',Date:sandbox.Date,AbortController,
    setTimeout:()=>1,clearTimeout(){},confirm:()=>true,
    sessionStorage:{getItem:k=>saved.get(k),setItem:(k,v)=>saved.set(k,v),removeItem:k=>saved.delete(k)},
    document:{body:{classList:{add:k=>classes.add(k),remove:k=>classes.delete(k)}},
      getElementById:element,querySelectorAll:()=>[]},
    fetch:async(url,options)=>{
      const body=JSON.parse(options.body); calls.push({url,options,body});
      const result=post(body);
      return {ok:true,json:async()=>result};
    },
    resetFormFields(){element('f_student').value='';},hideSearchView(){},currentTopicKey:null
  };
  vm.createContext(f);
  vm.runInContext(readFileSync('auth.js','utf8')+'\nthis.auth=EPAAuth;',f);
  await f.auth.start(async()=>{ready++;});
  assert.ok(classes.has('auth-locked'));
  await assert.rejects(f.auth.request('submit',{}));
  assert.equal(calls.length,0);
  element('authCode').value=initialCode;
  await f.auth.login({preventDefault(){}});
  assert.equal(ready,1);assert.equal(element('authCode').value,'');
  assert.ok(!classes.has('auth-locked'));
  assert.ok(![...saved.values()].join('').includes(initialCode));
  await f.auth.request('search',{params:{topic:'輸血護理'}});
  const last=calls.at(-1);
  assert.equal(last.options.method,'POST');assert.equal(last.options.credentials,'omit');
  assert.ok(last.body.token);assert.ok(!last.url.includes(last.body.token));
  element('f_student').value='尚未送出的學員';
  properties.set('EPA_ACCESS_CODE',randomUUID());
  await assert.rejects(f.auth.request('submit',{data:{}}));
  assert.ok(classes.has('auth-locked'));
  assert.equal(element('f_student').value,'尚未送出的學員');
  const writes=calls.filter(c=>c.body.action==='submit');assert.equal(writes.length,1);
  properties.set('EPA_ACCESS_CODE',initialCode);
  element('authCode').value=initialCode;
  await f.auth.login({preventDefault(){}});
  assert.equal(ready,1); // 重新登入不會重置填寫中的表單。
  await f.auth.logout();
  assert.equal(element('f_student').value,'');assert.equal(saved.size,0);
  assert.ok(element('searchResults').cleared);
  checks++; console.log('PASS 前端驗證、憑證儲存、過期保留表單、禁止重送與登出清理');
}
testFrontend().then(()=>console.log(`${checks} security checks passed`)).catch(error=>{console.error(error);process.exitCode=1;});
