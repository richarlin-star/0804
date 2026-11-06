/* HTML Slide Embed — a PowerPoint content add-in. No framework or build step. */
(() => {
  'use strict';
  const STORAGE_KEY = 'htmlSlideEmbed.settings.v1';
  // Deliberately conservative. Office's Settings API does not guarantee arbitrarily large blobs.
  const MAX_SAVED_FILE_BYTES = 240 * 1024;
  const $ = id => document.getElementById(id);
  const state = { mode: 'empty', name: '', html: '', url: '', fit: 'contain', ratio: '16/10', canPersist: true };
  const frame = $('contentFrame');
  const stage = $('stage');
  const viewport = $('viewport');
  let officeReady = false;
  let statusTimeout;

  function status(message, isError = false, sticky = false) {
    const elem = $('status');
    elem.textContent = message;
    elem.classList.toggle('error', isError);
    elem.classList.add('visible');
    clearTimeout(statusTimeout);
    if (!sticky) statusTimeout = setTimeout(() => elem.classList.remove('visible'), isError ? 9500 : 5400);
  }
  function settingClient() {
    try {
      if (officeReady && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
        return Office.context.document.settings;
      }
    } catch (_) { /* Office is not available in browser preview. */ }
    return null;
  }
  function persistedConfig() {
    return { mode: state.mode, name: state.name, html: state.mode === 'file' ? state.html : '',
      url: state.mode === 'url' ? state.url : '', fit: state.fit, ratio: state.ratio };
  }
  async function persist() {
    if (!state.canPersist) {
      status('檔案較大：目前可播放，但未存入 PPT。請改用 HTTPS 網址，或壓縮成較小的單一 HTML。', true, true);
      return false;
    }
    const payload = persistedConfig();
    const client = settingClient();
    if (!client) {
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(payload)); }
      catch (error) { status('瀏覽器暫存儲存失敗：' + error.message, true); return false; }
      status('已儲存在瀏覽器測試模式（非 PPT）。');
      return true;
    }
    try {
      client.set(STORAGE_KEY, payload);
      await new Promise((resolve, reject) => client.saveAsync(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error?.message || 'Office 設定寫入失敗'));
      }));
      status('展示設定已寫入簡報。請記得儲存 .pptx。');
      return true;
    } catch (error) {
      status('無法寫入簡報：' + error.message + '。目前仍可預覽。', true, true);
      return false;
    }
  }
  function restore() {
    let saved = null;
    try {
      const client = settingClient();
      saved = client ? client.get(STORAGE_KEY) : JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
    } catch (error) { status('讀取先前設定失敗：' + error.message, true); }
    if (saved && ['file','url','empty'].includes(saved.mode)) {
      state.mode = saved.mode;
      state.name = typeof saved.name === 'string' ? saved.name : '';
      state.html = typeof saved.html === 'string' ? saved.html : '';
      state.url = typeof saved.url === 'string' ? saved.url : '';
      state.fit = ['contain','cover','stretch'].includes(saved.fit) ? saved.fit : 'contain';
      state.ratio = ['16/10','16/9','4/3','auto'].includes(saved.ratio) ? saved.ratio : '16/10';
      $('fitMode').value = state.fit;
      $('aspectRatio').value = state.ratio;
    }
    render();
  }
  function canvasSize() {
    // A virtual CSS-pixel canvas is vital for slides authored at 2560×1600.
    const presets = { '16/10': [2560, 1600], '16/9': [1920, 1080], '4/3': [1600, 1200] };
    return presets[state.ratio] || null;
  }
  function resizeFrame() {
    if (state.mode === 'empty') return;
    const sw = Math.max(1, stage.clientWidth);
    const sh = Math.max(1, stage.clientHeight);
    const native = canvasSize();
    if (!native) {
      frame.style.width = sw + 'px'; frame.style.height = sh + 'px';
      frame.style.left = '0px'; frame.style.top = '0px';
      frame.style.transform = 'none';
      return;
    }
    const [cw, ch] = native;
    const sx = sw / cw, sy = sh / ch;
    let scaleX, scaleY;
    if (state.fit === 'stretch') { scaleX = sx; scaleY = sy; }
    else { const uniform = state.fit === 'cover' ? Math.max(sx, sy) : Math.min(sx, sy); scaleX = scaleY = uniform; }
    const visibleW = cw * scaleX, visibleH = ch * scaleY;
    frame.style.width = cw + 'px'; frame.style.height = ch + 'px';
    frame.style.left = ((sw - visibleW) / 2) + 'px';
    frame.style.top = ((sh - visibleH) / 2) + 'px';
    frame.style.transform = `scale(${scaleX}, ${scaleY})`;
  }
  function render() {
    const empty = state.mode === 'empty' || (state.mode === 'file' && !state.html) || (state.mode === 'url' && !state.url);
    $('emptyState').classList.toggle('hidden', !empty);
    viewport.classList.toggle('hidden', empty);
    if (empty) {
      frame.removeAttribute('srcdoc');
      frame.removeAttribute('src');
      return;
    }
    if (state.mode === 'file') {
      // Opaque sandbox origin: imported scripts cannot read or modify the Office add-in parent.
      frame.setAttribute('sandbox', 'allow-scripts allow-forms allow-popups');
      frame.removeAttribute('src');
      frame.srcdoc = state.html;
    } else {
      // HTTPS apps often require their own origin (fetch, storage, JS assets).
      // Only load user-trusted URLs. Other websites may prohibit framing via CSP/X-Frame-Options.
      frame.setAttribute('sandbox', 'allow-scripts allow-forms allow-popups allow-same-origin');
      frame.removeAttribute('srcdoc');
      frame.src = state.url;
    }
    resizeFrame();
  }
  async function loadFile(file) {
    if (!file) return;
    const htmlName = /\.html?$/i.test(file.name) || file.type === 'text/html';
    if (!htmlName) { status('請選擇 .html 或 .htm 檔。', true); return; }
    try {
      const html = await file.text();
      state.mode = 'file'; state.name = file.name; state.html = html; state.url = '';
      state.canPersist = new Blob([html]).size <= MAX_SAVED_FILE_BYTES;
      render();
      if (state.canPersist) await persist();
      else status(`已載入 ${file.name}；檔案大於 240 KiB，只能暫時預覽。建議部署成 HTTPS 網站。`, true, true);
    } catch (error) { status('檔案載入失敗：' + error.message, true); }
  }
  async function loadUrl(url) {
    try {
      const parsed = new URL(url);
      if (parsed.protocol !== 'https:') throw new Error('只支援 HTTPS 網址');
      state.mode = 'url'; state.name = parsed.hostname; state.html = ''; state.url = parsed.href;
      state.canPersist = true;
      render();
      await persist();
      status('已載入網址。若出現空白，請確認網站允許 iframe 嵌入。');
    } catch (error) { status('網址無效：' + error.message, true); }
  }
  async function clear() {
    state.mode = 'empty'; state.name = ''; state.html = ''; state.url = ''; state.canPersist = true;
    render(); await persist();
  }
  async function loadDemo() {
    // The small sample is responsive; real 2560×1600 demos may use the fixed canvas.
    state.ratio = 'auto'; $('aspectRatio').value = 'auto';
    await loadUrl(new URL('./demo.html', location.href).href);
  }
  function showUrlDialog(open) {
    $('urlDialog').classList.toggle('hidden', !open);
    if (open) { $('urlInput').value = state.mode === 'url' ? state.url : ''; $('urlInput').focus(); }
  }
  function togglePresentation(on) {
    $('app').classList.toggle('presentation', on);
    $('expandButton').classList.toggle('hidden', !on);
    resizeFrame();
  }
  function setFileHandlers(inputId) {
    $(inputId).addEventListener('change', async event => {
      await loadFile(event.target.files && event.target.files[0]);
      event.target.value = '';
    });
  }
  setFileHandlers('fileInput');
  setFileHandlers('emptyFileInput');
  $('demoButton').addEventListener('click', loadDemo);
  $('emptyDemoButton').addEventListener('click', loadDemo);
  $('urlButton').addEventListener('click', () => showUrlDialog(true));
  $('cancelUrl').addEventListener('click', () => showUrlDialog(false));
  $('urlDialog').addEventListener('click', event => { if (event.target === $('urlDialog')) showUrlDialog(false); });
  $('urlForm').addEventListener('submit', async event => {
    event.preventDefault();
    const url = $('urlInput').value.trim();
    showUrlDialog(false);
    await loadUrl(url);
  });
  $('fitMode').addEventListener('change', async event => { state.fit = event.target.value; resizeFrame(); await persist(); });
  $('aspectRatio').addEventListener('change', async event => { state.ratio = event.target.value; resizeFrame(); await persist(); });
  $('reloadButton').addEventListener('click', render);
  $('clearButton').addEventListener('click', clear);
  $('collapseButton').addEventListener('click', () => togglePresentation(true));
  $('expandButton').addEventListener('click', () => togglePresentation(false));
  document.addEventListener('keydown', event => {
    if (event.key === 'Escape') { showUrlDialog(false); togglePresentation(false); }
  });
  // Drag and drop self-contained HTML onto the slide add-in.
  stage.addEventListener('dragover', event => { event.preventDefault(); event.dataTransfer.dropEffect = 'copy'; });
  stage.addEventListener('drop', event => {
    event.preventDefault();
    const file = event.dataTransfer.files && event.dataTransfer.files[0];
    if (file) loadFile(file);
  });
  if ('ResizeObserver' in window) new ResizeObserver(resizeFrame).observe(stage);
  else window.addEventListener('resize', resizeFrame);

  // When not hosted by Office, make a browser-only preview available.
  let initialized = false;
  function initialize(inOffice) {
    if (initialized) return;
    initialized = true;
    officeReady = !!inOffice;
    restore();
    if (!inOffice) status('網頁預覽模式：若要儲存在 PowerPoint，請先安裝內容增益集。');
  }
  if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
    Office.onReady(() => initialize(true)).catch(() => initialize(false));
    setTimeout(() => initialize(false), 7000);
  } else initialize(false);
})();
