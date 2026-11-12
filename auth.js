/* 使用碼只會在登入時送到既有的 Apps Script；不儲存使用碼。 */
const EPAAuth = (() => {
  const storageKey = 'epa.session.v1';
  let token = '';
  let expiresAt = 0;
  let expiryTimer;
  let onFirstLogin;
  let initialized = false;

  function forget() {
    token = '';
    expiresAt = 0;
    clearTimeout(expiryTimer);
    try { sessionStorage.removeItem(storageKey); } catch (_) {}
  }

  function lock(message) {
    document.body.classList.add('auth-locked');
    document.getElementById('authGate').hidden = false;
    document.getElementById('authMessage').textContent = message || '';
    document.getElementById('authCode').value = '';
  }

  function remember(session) {
    if (typeof session.token !== 'string' || !/^[a-f0-9]{64}$/i.test(session.token) ||
        !Number.isFinite(session.expiresAt) || session.expiresAt <= Date.now()) {
      throw new Error('登入回應不完整，請稍後再試。');
    }
    token = session.token;
    expiresAt = session.expiresAt;
    try { sessionStorage.setItem(storageKey, JSON.stringify({token, expiresAt})); } catch (_) {}
    clearTimeout(expiryTimer);
    expiryTimer = setTimeout(() => {
      forget();
      lock('登入已到期，請重新輸入使用碼。尚未送出的表單仍保留在目前分頁。');
    }, expiresAt - Date.now());
  }

  async function post(body) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 25000);
    try {
      const response = await fetch(API_URL, {
        method: 'POST', credentials: 'omit', cache: 'no-store',
        headers: {'Content-Type': 'text/plain;charset=utf-8'},
        body: JSON.stringify(body), signal: controller.signal
      });
      if (!response.ok) throw new Error('後台連線失敗，請稍後再試。');
      let result;
      try { result = await response.json(); }
      catch (_) { throw new Error('後台回應格式不正確，請稍後再試。'); }
      if (result && result.error) {
        const error = new Error(result.error);
        error.code = result.code;
        throw error;
      }
      return result;
    } catch (error) {
      if (error.name === 'AbortError') {
        throw new Error('連線逾時。若剛才正在送出，請先查詢紀錄確認是否已儲存，再決定是否重試。');
      }
      throw error;
    } finally { clearTimeout(timer); }
  }

  async function request(action, values = {}) {
    if (!token || expiresAt <= Date.now()) {
      forget();
      const message = '請先登入；尚未送出的內容仍保留在目前分頁。';
      lock(message);
      throw new Error(message);
    }
    try { return await post({...values, action, token}); }
    catch (error) {
      if (error.code === 'AUTH_REQUIRED' || error.code === 'AUTH_NOT_CONFIGURED') {
        forget();
        lock(error.message);
      }
      // 不自動重送寫入，避免重複新增評核。
      throw error;
    }
  }

  async function enter() {
    document.body.classList.remove('auth-locked');
    document.getElementById('authGate').hidden = true;
    if (!initialized) {
      await onFirstLogin();
      initialized = true;
    }
  }

  async function login(event) {
    event.preventDefault();
    const input = document.getElementById('authCode');
    const button = document.getElementById('authSubmit');
    button.disabled = true;
    document.getElementById('authMessage').textContent = '正在驗證…';
    try {
      const result = await post({action: 'login', accessCode: input.value});
      input.value = '';
      remember(result);
      await enter();
    } catch (error) {
      document.getElementById('authMessage').textContent = error.message;
    } finally {
      input.value = '';
      button.disabled = false;
    }
  }

  async function logout() {
    if (!confirm('登出會清除目前畫面尚未送出的內容，確定登出？')) return;
    const previousToken = token;
    forget();
    resetFormFields();
    hideSearchView();
    document.getElementById('searchResults').replaceChildren();
    document.querySelectorAll('#searchView input, #searchView select').forEach(el => { el.value = ''; });
    document.getElementById('topicScenarioInput').value = '';
    if (currentTopicKey) loadTopic(currentTopicKey);
    lock('已登出。');
    if (previousToken) {
      try { await post({action: 'logout', token: previousToken}); }
      catch (error) {
        if (error.code !== 'AUTH_REQUIRED') {
          document.getElementById('authMessage').textContent =
            '已清除本機登入；伺服器登出尚未確認，原憑證最長4小時後失效。';
        }
      }
    }
  }

  async function start(ready) {
    onFirstLogin = ready;
    lock('請輸入護理部提供的使用碼。');
    let saved;
    try { saved = JSON.parse(sessionStorage.getItem(storageKey) || 'null'); } catch (_) {}
    if (saved && typeof saved.token === 'string' && saved.expiresAt > Date.now()) {
      try {
        remember(saved);
        await request('session');
        await enter();
        return;
      } catch (error) {
        forget();
        lock(error.message);
      }
    } else { forget(); }
  }

  return {request, login, logout, start};
})();
