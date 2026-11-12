/**
 * 共通照護EPAs — 驗證版後端。
 * 使用碼只設定在「專案設定 → 指令碼屬性 → EPA_ACCESS_CODE」。
 * 不可把使用碼填入程式碼、Google Sheet 或 GitHub。
 * 更新原部署的版本，保留原本 /exec 網址與所有評核資料。
 */
const SHEET_ID = '1pVafBIBA0mMIx1vq8wcq_srb-NjgZAG-UOQjztNEJ8o';
const SHEET_TOPICS = '主題設定表';
const SHEET_RECORDS = '評核紀錄表';
const TIMEZONE = 'Asia/Taipei';
const EPA_SESSION_SECONDS = 4 * 60 * 60;

// 僅供專案編輯器內執行的設定檢查，不輸出使用碼或登入憑證。
function checkAuthSetup() {
  accessCode_();
  console.log('EPA使用碼已設定，長度符合規定。');
}

// 所有含憑證的請求只接受 POST 本文，避免密碼或憑證出現在網址。
function doGet() {
  return jsonOut({error: '請開啟EPA填寫網站，登入後再操作。', code: 'AUTH_REQUIRED'});
}

function doPost(e) {
  try {
    const text = e && e.postData && e.postData.contents;
    if (typeof text !== 'string' || text.length > 100000) {
      throw epaError_('BAD_REQUEST', '請求格式不正確。');
    }
    let body;
    try { body = JSON.parse(text); }
    catch (_) { throw epaError_('BAD_REQUEST', '請求格式不正確。'); }
    if (!body || Array.isArray(body) || typeof body !== 'object') {
      throw epaError_('BAD_REQUEST', '請求格式不正確。');
    }
    if (body.action === 'login') return jsonOut(login_(body.accessCode));

    const session = requireSession_(body.token);
    switch (body.action) {
      case 'session': return jsonOut({ok: true, expiresAt: session.expiresAt});
      case 'logout':
        CacheService.getScriptCache().remove(sessionKey_(body.token));
        return jsonOut({ok: true});
      case 'getTopics': return jsonOut(getTopics());
      case 'search': return jsonOut(searchRecords(body.params || {}));
      case 'getRecord': return jsonOut(getRecord(body.id));
      case 'submit': return jsonOut(submitRecord(body.data, body.recordId || null));
      default: throw epaError_('BAD_REQUEST', '不支援的操作。');
    }
  } catch (err) {
    // 不回傳內部例外、試算表路徑或請求內容；不記錄密碼或憑證。
    return jsonOut({
      error: err && err.epaCode ? err.message : '服務暫時無法完成操作，請稍後再試。',
      code: err && err.epaCode ? err.epaCode : 'SERVER_ERROR'
    });
  }
}

function epaError_(code, message) {
  const error = new Error(message);
  error.epaCode = code;
  return error;
}

function digest_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(value), Utilities.Charset.UTF_8)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function equalDigest_(a, b) {
  let difference = 0;
  for (let i = 0; i < 64; i++) difference |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return difference === 0;
}

function accessCode_() {
  const code = PropertiesService.getScriptProperties().getProperty('EPA_ACCESS_CODE');
  if (!code || code.length < 6 || code.length > 256) {
    throw epaError_('AUTH_NOT_CONFIGURED', '請護理部先設定至少6字元的EPA使用碼。');
  }
  return code;
}

function sessionKey_(token) { return 'epa-session:' + digest_(token); }

function login_(candidate) {
  const configuredCode = accessCode_();
  if (typeof candidate !== 'string' || candidate.length > 256) {
    throw epaError_('LOGIN_FAILED', '使用碼不正確。');
  }
  // 共用服務的登入嘗試上限；Cache 提早失效時仍須正確使用碼，絕不放行。
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) throw epaError_('BUSY', '目前登入人數較多，請稍後再試。');
  const cache = CacheService.getScriptCache();
  try {
    const bucket = Math.floor(Date.now() / 60000);
    const key = 'epa-login:' + bucket;
    const count = Number(cache.get(key) || 0);
    if (count >= 30) throw epaError_('RATE_LIMIT', '登入嘗試過於頻繁，請稍後再試。');
    cache.put(key, String(count + 1), 120);
    if (!equalDigest_(digest_(candidate), digest_(configuredCode))) {
      throw epaError_('LOGIN_FAILED', '使用碼不正確。');
    }
    const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    const expiresAt = Date.now() + EPA_SESSION_SECONDS * 1000;
    cache.put(sessionKey_(token), JSON.stringify({
      expiresAt: expiresAt, configuration: digest_(configuredCode)
    }), EPA_SESSION_SECONDS);
    return {ok: true, token: token, expiresAt: expiresAt};
  } finally { lock.releaseLock(); }
}

function requireSession_(token) {
  const code = accessCode_();
  if (typeof token !== 'string' || !/^[a-f0-9]{64}$/i.test(token)) {
    throw epaError_('AUTH_REQUIRED', '請先登入EPA網站。');
  }
  let session;
  try { session = JSON.parse(CacheService.getScriptCache().get(sessionKey_(token)) || 'null'); }
  catch (_) { session = null; }
  if (!session || !Number.isFinite(session.expiresAt) || session.expiresAt <= Date.now() ||
      session.configuration !== digest_(code)) {
    throw epaError_('AUTH_REQUIRED', '登入已到期，請重新輸入使用碼。尚未送出的表單仍保留在目前分頁。');
  }
  return session;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet');
  return sheet;
}

function cleanText_(value) { return String(value == null ? '' : value).trim(); }

function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(r => r.some(c => c !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let value = row[i];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, TIMEZONE,
          h === '建立時間' || h === '更新時間' ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
      }
      obj[h] = value;
    });
    // 不改動舊資料列，回傳時補上主題識別碼，讓舊紀錄可正常載入。
    if (obj['主題名稱'] != null) {
      obj['主題名稱'] = cleanText_(obj['主題名稱']);
      obj['主題代碼'] = cleanText_(obj['主題代碼']) || obj['主題名稱'];
    }
    return obj;
  });
}

function getTopics() {
  return sheetToObjects(getSheet(SHEET_TOPICS)).filter(o => o['主題名稱']).map(o => {
    const items = [];
    for (let i = 1; i <= 10; i++) if (o['評估項目' + i]) items.push(o['評估項目' + i]);
    return {
      code: o['主題代碼'], title: o['主題名稱'], scenario: o['情境說明'],
      scenarioEditable: cleanText_(o['情境需現場填寫']) === 'Y',
      items: items, levelDescription: o['整體任務評估說明'] || ''
    };
  });
}

function nextRecordId() {
  return 'EPA-' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd-HHmmss') + '-' +
    Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function safeCell_(value) {
  if (typeof value === 'string' && /^[=+\-@]/.test(value)) return "'" + value;
  return value;
}

function submitRecord(data, existingId) {
  if (!data || Array.isArray(data) || typeof data !== 'object' ||
      !cleanText_(data['學員姓名']) || !['草稿', '正式'].includes(data['狀態'])) {
    throw epaError_('BAD_REQUEST', '請填寫學員姓名並選擇正確的儲存狀態。');
  }
  Object.keys(data).forEach(key => {
    if (!['string', 'number', 'boolean'].includes(typeof data[key]) && data[key] !== null) {
      throw epaError_('BAD_REQUEST', '表單資料格式不正確。');
    }
    if (String(data[key]).length > 10000) throw epaError_('BAD_REQUEST', '單一欄位內容過長。');
  });
  data['主題名稱'] = cleanText_(data['主題名稱']);
  data['主題代碼'] = cleanText_(data['主題代碼']) || data['主題名稱'];
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw epaError_('BUSY', '目前有人正在儲存，請稍後再試。');
  try {
    const sheet = getSheet(SHEET_RECORDS);
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const now = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const idCol = headers.indexOf('紀錄編號');
    const statusCol = headers.indexOf('狀態');
    if (idCol < 0 || statusCol < 0) throw new Error('Missing columns');
    if (existingId) {
      const values = sheet.getDataRange().getValues();
      for (let r = 1; r < values.length; r++) {
        if (values[r][idCol] !== existingId) continue;
        if (values[r][statusCol] === '正式') break;
        const newRow = headers.map((h, i) => {
          if (h === '紀錄編號' || h === '建立時間') return values[r][i];
          if (h === '更新時間') return now;
          return Object.prototype.hasOwnProperty.call(data, h) ? safeCell_(data[h]) : values[r][i];
        });
        sheet.getRange(r + 1, 1, 1, lastCol).setValues([newRow]);
        SpreadsheetApp.flush();
        return {recordId: existingId, status: 'updated'};
      }
    }
    const id = nextRecordId();
    sheet.appendRow(headers.map(h => {
      if (h === '紀錄編號') return id;
      if (h === '建立時間' || h === '更新時間') return now;
      return data[h] !== undefined ? safeCell_(data[h]) : '';
    }));
    SpreadsheetApp.flush();
    return {recordId: id, status: existingId ? 'locked_new' : 'created'};
  } finally { lock.releaseLock(); }
}

function searchRecords(params) {
  return sheetToObjects(getSheet(SHEET_RECORDS)).filter(o => {
    if (params.student && !cleanText_(o['學員姓名']).includes(cleanText_(params.student))) return false;
    if (params.studentExact && cleanText_(o['學員姓名']) !== cleanText_(params.studentExact)) return false;
    if (params.teacher && !cleanText_(o['教師姓名']).includes(cleanText_(params.teacher))) return false;
    if (params.topic && ![o['主題代碼'], o['主題名稱']].includes(cleanText_(params.topic))) return false;
    if (params.term && cleanText_(o['期程']) !== cleanText_(params.term)) return false;
    if (params.status && o['狀態'] !== params.status) return false;
    if (params.dateFrom && String(o['填寫日期'] || '') < params.dateFrom) return false;
    if (params.dateTo && String(o['填寫日期'] || '') > params.dateTo) return false;
    return true;
  }).sort((a, b) => String(b['填寫日期'] || '').localeCompare(String(a['填寫日期'] || '')));
}

function getRecord(id) {
  return sheetToObjects(getSheet(SHEET_RECORDS)).find(o => o['紀錄編號'] === id) || null;
}
