// MS Downloader — on-page button and dialog.
//
// Shows a small "Transcript" button on SharePoint / Stream recording pages
// (including the recording frame inside Teams). Clicking it opens a dialog to
// pick a format, preview the result and save it.

(() => {
  'use strict';

  const CHANNEL = 'ms-downloader';
  const T = globalThis.MSDLTranscript;
  const STORE_KEY = 'msdl.transcriptPrefs';

  const state = {
    file: null,        // { sitePath, driveId, itemId, name, hasTranscripts }
    bearer: null,      // "Bearer …" seen on the player's own API calls
    data: null,        // transcript JSON for the current file
    dataKey: null,
    prefs: { format: 'txt', speakers: true, timestamps: true }
  };

  // ---------- messages from page-hook.js ----------

  window.addEventListener('message', (event) => {
    if (event.source !== window) return;
    const msg = event.data;
    if (!msg || msg.channel !== CHANNEL) return;
    if (msg.kind === 'file' && msg.payload && msg.payload.itemId) {
      state.file = msg.payload;
      showLauncher();
    } else if (msg.kind === 'bearer' && msg.payload) {
      state.bearer = msg.payload.value;
    }
  });
  window.postMessage({ channel: CHANNEL, kind: 'hello' }, window.location.origin);

  try {
    chrome.storage.local.get(STORE_KEY, (res) => {
      if (res && res[STORE_KEY]) Object.assign(state.prefs, res[STORE_KEY]);
    });
  } catch (_) { /* storage is optional */ }

  function savePrefs() {
    try { chrome.storage.local.set({ [STORE_KEY]: state.prefs }); } catch (_) { /* ignore */ }
  }

  // ---------- SharePoint requests ----------

  // Microsoft Defender for Cloud Apps serves SharePoint as
  // *.sharepoint.com.mcas.ms; links it hands back may still point at the
  // plain host, which the proxied session cannot use.
  function throughProxyIfNeeded(url) {
    try {
      const u = new URL(url, location.href);
      if (location.hostname.endsWith('.mcas.ms') && !u.hostname.endsWith('.mcas.ms') &&
          u.hostname.endsWith('.sharepoint.com')) {
        u.hostname += '.mcas.ms';
      }
      return u.href;
    } catch (_) {
      return url;
    }
  }

  async function apiGet(url) {
    const headers = { Accept: 'application/json' };
    if (state.bearer) headers.Authorization = state.bearer;
    const res = await fetch(url, { credentials: 'include', headers });
    if (!res.ok) {
      const err = new Error(`SharePoint answered ${res.status}`);
      err.status = res.status;
      throw err;
    }
    return res.json();
  }

  function pickTranscript(body) {
    const list =
      (body && Array.isArray(body.value) && body.value) ||
      (body && body.media && Array.isArray(body.media.transcripts) && body.media.transcripts) ||
      [];
    const usable = list.filter((t) => t && t.temporaryDownloadUrl);
    return usable.find((t) => t.isDefault) || usable[0] || null;
  }

  async function loadTranscript() {
    const f = state.file;
    const key = f.driveId + '/' + f.itemId;
    if (state.data && state.dataKey === key) return state.data;

    if (f.hasTranscripts === false) {
      const err = new Error('This recording has no transcript.');
      err.code = 'none';
      throw err;
    }

    const item = `${location.origin}${f.sitePath}/_api/v2.1/drives/${encodeURIComponent(f.driveId)}/items/${encodeURIComponent(f.itemId)}`;
    let entry = null;
    let lastError = null;
    for (const url of [
      `${item}/media/transcripts`,
      `${item}?$select=id&$expand=media/transcripts`
    ]) {
      try {
        entry = pickTranscript(await apiGet(url));
        if (entry) break;
      } catch (e) {
        lastError = e;
      }
    }
    if (!entry) {
      const err = lastError || new Error('This recording has no transcript.');
      if (!lastError) err.code = 'none';
      throw err;
    }

    const link = new URL(throughProxyIfNeeded(entry.temporaryDownloadUrl));
    link.searchParams.set('format', 'json');
    const res = await fetch(link.href, { credentials: 'include' });
    if (!res.ok) throw new Error(`Transcript download failed (${res.status})`);
    const data = await res.json();

    state.data = data;
    state.dataKey = key;
    return data;
  }

  // ---------- UI ----------

  let host = null;
  let ui = null;

  const STYLE = `
    :host { all: initial; }
    * { box-sizing: border-box; font-family: "Segoe UI", system-ui, -apple-system, sans-serif; }
    .launcher {
      position: fixed; right: 20px; bottom: 20px; z-index: 2147483646;
      display: flex; align-items: center; gap: 8px;
      padding: 10px 16px; border: 0; border-radius: 999px; cursor: pointer;
      background: #1d4ed8; color: #fff; font-size: 13px; font-weight: 600;
      box-shadow: 0 6px 20px rgba(29, 78, 216, .35);
    }
    .launcher:hover { background: #1e40af; }
    .launcher svg { width: 16px; height: 16px; }
    .backdrop {
      position: fixed; inset: 0; z-index: 2147483647;
      background: rgba(15, 23, 42, .45); display: none;
      align-items: center; justify-content: center; padding: 16px;
    }
    .backdrop.open { display: flex; }
    .dialog {
      width: min(560px, 100%); max-height: 90vh; overflow: auto;
      background: #fff; color: #0f172a; border-radius: 14px;
      box-shadow: 0 24px 60px rgba(15, 23, 42, .3); padding: 22px;
    }
    h2 { margin: 0 0 4px; font-size: 18px; }
    .sub { margin: 0 0 18px; color: #64748b; font-size: 13px; }
    label.field { display: block; font-size: 12px; font-weight: 600; color: #334155; margin: 14px 0 6px; }
    select, input[type=text] {
      width: 100%; padding: 9px 10px; border: 1px solid #cbd5e1; border-radius: 8px;
      font-size: 14px; color: #0f172a; background: #fff;
    }
    .checks { display: flex; gap: 18px; margin-top: 12px; font-size: 13px; color: #334155; }
    .checks label { display: flex; align-items: center; gap: 6px; cursor: pointer; }
    pre {
      margin: 0; padding: 12px; height: 180px; overflow: auto; white-space: pre-wrap;
      background: #f1f5f9; border-radius: 8px; font: 12px/1.5 ui-monospace, Consolas, monospace; color: #1e293b;
    }
    .status { font-size: 13px; color: #64748b; margin-top: 10px; min-height: 18px; }
    .status.error { color: #b91c1c; }
    .row { display: flex; justify-content: flex-end; gap: 10px; margin-top: 18px; }
    button.btn {
      padding: 9px 16px; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer;
      border: 1px solid #cbd5e1; background: #fff; color: #0f172a;
    }
    button.primary { background: #1d4ed8; border-color: #1d4ed8; color: #fff; }
    button.primary:disabled { opacity: .5; cursor: default; }
  `;

  const ICON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v11"/><path d="m7 10 5 5 5-5"/><path d="M5 20h14"/></svg>';

  function build() {
    host = document.createElement('div');
    host.id = 'ms-downloader-root';
    const root = host.attachShadow({ mode: 'open' });
    root.innerHTML = `
      <style>${STYLE}</style>
      <button class="launcher" type="button" title="MS Downloader">${ICON}<span>Transcript</span></button>
      <div class="backdrop" role="dialog" aria-modal="true" aria-labelledby="msdl-title">
        <div class="dialog">
          <h2 id="msdl-title">Download transcript</h2>
          <p class="sub"></p>
          <label class="field" for="msdl-format">Format</label>
          <select id="msdl-format"></select>
          <div class="checks">
            <label><input type="checkbox" id="msdl-speakers"> Speaker names</label>
            <label><input type="checkbox" id="msdl-times"> Timestamps</label>
          </div>
          <label class="field" for="msdl-name">File name</label>
          <input type="text" id="msdl-name" spellcheck="false">
          <label class="field">Preview</label>
          <pre id="msdl-preview"></pre>
          <div class="status" id="msdl-status"></div>
          <div class="row">
            <button class="btn" type="button" id="msdl-cancel">Close</button>
            <button class="btn primary" type="button" id="msdl-save" disabled>Download</button>
          </div>
        </div>
      </div>`;

    ui = {
      launcher: root.querySelector('.launcher'),
      backdrop: root.querySelector('.backdrop'),
      sub: root.querySelector('.sub'),
      format: root.getElementById('msdl-format'),
      speakers: root.getElementById('msdl-speakers'),
      times: root.getElementById('msdl-times'),
      name: root.getElementById('msdl-name'),
      preview: root.getElementById('msdl-preview'),
      status: root.getElementById('msdl-status'),
      cancel: root.getElementById('msdl-cancel'),
      save: root.getElementById('msdl-save')
    };

    for (const [id, f] of Object.entries(T.FORMATS)) {
      const opt = document.createElement('option');
      opt.value = id;
      opt.textContent = f.label;
      ui.format.appendChild(opt);
    }

    ui.launcher.addEventListener('click', openDialog);
    ui.cancel.addEventListener('click', closeDialog);
    ui.backdrop.addEventListener('click', (e) => { if (e.target === ui.backdrop) closeDialog(); });
    root.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeDialog(); });
    ui.format.addEventListener('change', onOptionsChange);
    ui.speakers.addEventListener('change', onOptionsChange);
    ui.times.addEventListener('change', onOptionsChange);
    ui.save.addEventListener('click', save);

    (document.body || document.documentElement).appendChild(host);
  }

  function showLauncher() {
    if (!host) build();
    if (!host.isConnected) (document.body || document.documentElement).appendChild(host);
  }

  function setStatus(text, isError) {
    ui.status.textContent = text || '';
    ui.status.classList.toggle('error', !!isError);
  }

  function safeName(name) {
    return String(name || 'transcript')
      .replace(/[\\/:*?"<>|\u0000-\u001f]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 150) || 'transcript';
  }

  function refreshControls() {
    const fmt = state.prefs.format;
    ui.format.value = fmt;
    ui.speakers.checked = state.prefs.speakers;
    ui.times.checked = state.prefs.timestamps;
    ui.speakers.disabled = fmt === 'json';
    ui.times.disabled = fmt !== 'txt';
  }

  function renderPreview() {
    if (!state.data) return;
    const out = T.render(state.prefs.format, state.data, state.prefs);
    const lines = out.split('\n');
    ui.preview.textContent = lines.slice(0, 40).join('\n') + (lines.length > 40 ? '\n…' : '');
  }

  function onOptionsChange() {
    state.prefs.format = ui.format.value;
    state.prefs.speakers = ui.speakers.checked;
    state.prefs.timestamps = ui.times.checked;
    savePrefs();
    refreshControls();
    renderPreview();
  }

  async function openDialog() {
    ui.backdrop.classList.add('open');
    ui.sub.textContent = state.file ? state.file.name : '';
    ui.name.value = safeName(state.file && state.file.name);
    ui.preview.textContent = '';
    ui.save.disabled = true;
    refreshControls();
    setStatus('Loading transcript…');
    try {
      const data = await loadTranscript();
      const count = T.normalize(data).length;
      if (!count) throw Object.assign(new Error('The transcript is empty.'), { code: 'none' });
      setStatus(`${count} lines found.`);
      ui.save.disabled = false;
      renderPreview();
    } catch (e) {
      console.warn('[MS Downloader]', e);
      if (e.code === 'none') {
        setStatus('This recording has no transcript. Transcription has to be turned on during the meeting.', true);
      } else if (e.status === 401 || e.status === 403) {
        setStatus('SharePoint refused the request. Play the video for a few seconds or open its Transcript tab, then try again.', true);
      } else {
        setStatus('Could not load the transcript: ' + e.message + '. Refresh the page and try again.', true);
      }
    }
  }

  function closeDialog() {
    ui.backdrop.classList.remove('open');
  }

  function save() {
    if (!state.data) return;
    const fmt = T.FORMATS[state.prefs.format] || T.FORMATS.txt;
    const text = T.render(state.prefs.format, state.data, state.prefs);
    const blob = new Blob([state.prefs.format === 'json' ? text : '\ufeff' + text], { type: fmt.mime + ';charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${safeName(ui.name.value)}.${fmt.ext}`;
    document.documentElement.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    setStatus('Saved to your Downloads folder.');
  }
})();
