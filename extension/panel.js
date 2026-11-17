// MS Downloader — on-page Download button, menu and dialogs.
//
// On SharePoint / Stream recording pages (including the recording frame inside
// Teams) a "Download" button appears. Its menu offers:
//   Video + audio · Audio only · Video only · Transcript
// Video choices download the player's DASH stream in the browser and save one
// MP4 / M4A file. The transcript choice saves TXT, VTT, SRT or JSON.

(() => {
  'use strict';

  const CHANNEL = 'ms-downloader';
  const T = globalThis.MSDLTranscript;
  const Dash = globalThis.MSDLDash;
  const Mp4 = globalThis.MSDLMp4;
  const STORE_KEY = 'msdl.prefs';

  const MEDIA_CHOICES = {
    both: { label: 'Video + audio', hint: 'One MP4 file with picture and sound', ext: 'mp4', mime: 'video/mp4' },
    audio: { label: 'Audio only', hint: 'M4A sound file — small, good for podcasts and notes', ext: 'm4a', mime: 'audio/mp4' },
    video: { label: 'Video only', hint: 'MP4 picture without sound', ext: 'mp4', mime: 'video/mp4' }
  };

  const state = {
    file: null,        // { sitePath, driveId, itemId, name, hasTranscripts }
    bearer: null,      // Authorization header from the player's API calls
    manifest: null,    // { url, token } of the player's DASH manifest
    transcript: null,
    transcriptKey: null,
    job: null,         // running media download { controller }
    prefs: { media: 'both', format: 'txt', speakers: true, timestamps: true, parallel: 4 }
  };

  // ---------- messages from page-hook.js ----------

  window.addEventListener('message', (event) => {
    if (event.source !== window) return;
    const msg = event.data;
    if (!msg || msg.channel !== CHANNEL || !msg.payload) return;
    if (msg.kind === 'file' && msg.payload.itemId) {
      if (!state.file || state.file.itemId !== msg.payload.itemId) state.manifest = null;
      state.file = msg.payload;
      showLauncher();
    } else if (msg.kind === 'bearer') {
      state.bearer = msg.payload.value;
    } else if (msg.kind === 'manifest' && msg.payload.url) {
      state.manifest = msg.payload;
      showLauncher();
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

  // ---------- network helpers ----------

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

  function sleep(ms, signal) {
    return new Promise((resolve, reject) => {
      const t = setTimeout(resolve, ms);
      if (signal) {
        signal.addEventListener('abort', () => {
          clearTimeout(t);
          reject(signal.reason || new DOMException('Cancelled', 'AbortError'));
        }, { once: true });
      }
    });
  }

  // Cookies only go to the page's own host. Microsoft's media CDN (*.svc.ms)
  // rejects cross-site requests that carry credentials, and Chrome then
  // blocks the response ("Failed to fetch").
  function credsFor(url) {
    try {
      return new URL(url, location.href).origin === location.origin ? 'include' : 'omit';
    } catch (_) {
      return 'omit';
    }
  }

  function hostOf(url) {
    try { return new URL(url, location.href).hostname; } catch (_) { return '?'; }
  }

  // fetch with retries for throttling (429) and temporary server errors.
  async function fetchRetry(url, init, signal, onWait) {
    let wait = 1000;
    for (let attempt = 1; ; attempt++) {
      let res = null;
      try {
        res = await fetch(url, { ...init, credentials: credsFor(url), signal });
      } catch (e) {
        if (signal && signal.aborted) throw e;
        if (attempt >= 3) {
          console.warn('[MS Downloader] request blocked or offline:', hostOf(url), e);
          throw Object.assign(new Error(`could not reach ${hostOf(url)}`), { code: 'network', host: hostOf(url) });
        }
      }
      if (res && res.ok) return res;
      if (res && ![408, 429, 500, 502, 503, 504].includes(res.status)) {
        console.warn('[MS Downloader] HTTP', res.status, 'from', hostOf(url));
        throw Object.assign(new Error(`${hostOf(url)} answered ${res.status}`), { status: res.status });
      }
      if (attempt >= 6) {
        throw Object.assign(new Error(`server answered ${res ? res.status : 'no response'}`), { status: res && res.status });
      }
      const retryAfter = res && Number(res.headers.get('retry-after'));
      const delay = retryAfter > 0 ? retryAfter * 1000 : wait;
      if (onWait) onWait(delay, res && res.status);
      await sleep(delay, signal);
      wait = Math.min(wait * 2, 30000);
    }
  }

  async function apiGet(url) {
    const headers = { Accept: 'application/json' };
    if (state.bearer) headers.Authorization = state.bearer;
    const res = await fetch(url, { credentials: 'include', headers });
    if (!res.ok) throw Object.assign(new Error(`SharePoint answered ${res.status}`), { status: res.status });
    return res.json();
  }

  // ---------- transcript ----------

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
    if (!f) throw new Error('Open a recording first');
    const key = f.driveId + '/' + f.itemId;
    if (state.transcript && state.transcriptKey === key) return state.transcript;
    if (f.hasTranscripts === false) throw Object.assign(new Error('no transcript'), { code: 'none' });

    const item = `${location.origin}${f.sitePath}/_api/v2.1/drives/${encodeURIComponent(f.driveId)}/items/${encodeURIComponent(f.itemId)}`;
    let entry = null;
    let lastError = null;
    for (const url of [`${item}/media/transcripts`, `${item}?$select=id&$expand=media/transcripts`]) {
      try {
        entry = pickTranscript(await apiGet(url));
        if (entry) break;
      } catch (e) {
        lastError = e;
      }
    }
    if (!entry) throw lastError || Object.assign(new Error('no transcript'), { code: 'none' });

    const link = new URL(throughProxyIfNeeded(entry.temporaryDownloadUrl));
    link.searchParams.set('format', 'json');
    const res = await fetch(link.href, { credentials: credsFor(link.href) });
    if (!res.ok) throw new Error(`transcript download failed (${res.status})`);
    state.transcript = await res.json();
    state.transcriptKey = key;
    return state.transcript;
  }

  // ---------- video / audio ----------

  // If the hook missed the manifest request, look for it in the page's
  // resource timing list (no token available that way, but often not needed).
  function findManifest() {
    if (state.manifest) return state.manifest;
    try {
      const hit = performance.getEntriesByType('resource')
        .map((e) => e.name)
        .reverse()
        .find((n) => /videomanifest/i.test(n) && !/tempauth=/i.test(n));
      if (hit) state.manifest = { url: hit, token: null };
    } catch (_) { /* ignore */ }
    return state.manifest;
  }

  function hexBytes(hex) {
    const out = new Uint8Array(hex.length / 2);
    for (let i = 0; i < out.length; i++) out[i] = parseInt(hex.substr(i * 2, 2), 16);
    return out;
  }

  // The media CDN (*.svc.ms) wants the player's x-spopactoken.
  function headersFor(url, manifest) {
    const h = {};
    try {
      if (manifest.token && /\.svc\.ms$/i.test(new URL(url).hostname)) h['x-spopactoken'] = manifest.token;
    } catch (_) { /* ignore */ }
    return h;
  }

  async function getStreamPlan(signal) {
    const manifest = findManifest();
    if (!manifest) throw Object.assign(new Error('no manifest'), { code: 'nomanifest' });
    const res = await fetchRetry(manifest.url, { credentials: 'include', headers: headersFor(manifest.url, manifest) }, signal);
    const plan = Dash.parse(await res.text(), manifest.url);
    plan.manifest = manifest;
    return plan;
  }

  async function keyFor(stream, manifest, signal) {
    if (!stream.encryption) return null;
    const keyUrl = new URL(stream.encryption.keyUri, manifest.url).href;
    const res = await fetchRetry(keyUrl, { credentials: 'include', headers: headersFor(keyUrl, manifest) }, signal);
    const raw = await res.arrayBuffer();
    const key = await crypto.subtle.importKey('raw', raw, { name: 'AES-CBC' }, false, ['decrypt']);
    return { key, iv: hexBytes(stream.encryption.iv) };
  }

  // Shared limit on parallel requests across all streams.
  function makePool(limit, total, onProgress) {
    let active = 0;
    let finished = 0;
    let bytes = 0;
    const queue = [];
    const next = () => {
      while (active < limit && queue.length) {
        const { task, resolve, reject } = queue.shift();
        active++;
        Promise.resolve().then(task).then(resolve, reject).finally(() => { active--; next(); });
      }
    };
    return {
      run(task) { return new Promise((resolve, reject) => { queue.push({ task, resolve, reject }); next(); }); },
      done(n) { finished++; bytes += n; onProgress(finished, total, bytes, null); },
      onWait(ms, status) {
        onProgress(finished, total, bytes, `Server is busy (${status || 'network'}), waiting ${Math.round(ms / 1000)} s…`);
      }
    };
  }

  // Download one stream: returns { init, segments } as Uint8Arrays.
  async function fetchStream(stream, manifest, pool, signal) {
    const crypt = await keyFor(stream, manifest, signal);
    const open = async (buf) => new Uint8Array(crypt
      ? await crypto.subtle.decrypt({ name: 'AES-CBC', iv: crypt.iv }, crypt.key, buf)
      : buf);
    const get = async (url) => {
      const res = await fetchRetry(throughProxyIfNeeded(url),
        { credentials: 'include', headers: headersFor(url, manifest) }, signal, pool.onWait);
      const bytes = await open(await res.arrayBuffer());
      pool.done(bytes.length);
      return bytes;
    };
    const init = stream.init ? await pool.run(() => get(stream.init)) : null;
    const segments = await Promise.all(stream.segments.map((u) => pool.run(() => get(u))));
    return { init, segments };
  }

  async function downloadMedia(choice, onProgress, signal) {
    onProgress(0, 1, 0, 'Reading the video manifest…');
    const plan = await getStreamPlan(signal);
    if (plan.drm) throw Object.assign(new Error('drm'), { code: 'drm' });

    let streams;
    if (plan.combined) streams = [plan.combined];
    else if (choice === 'both') streams = [plan.video, plan.audio].filter(Boolean);
    else if (choice === 'audio') streams = [plan.audio];
    else streams = [plan.video];
    if (!streams.length || !streams[0]) {
      throw Object.assign(new Error('missing stream'), { code: choice === 'audio' ? 'noaudio' : 'novideo' });
    }

    const total = streams.reduce((n, s) => n + s.segments.length + (s.init ? 1 : 0), 0);
    const pool = makePool(Math.max(1, Math.min(16, state.prefs.parallel || 4)), total, onProgress);
    const fetched = await Promise.all(streams.map((s) => fetchStream(s, plan.manifest, pool, signal)));

    onProgress(total, total, 0, 'Putting the file together…');
    await sleep(0, signal);
    let parts;
    if (fetched.length === 2) {
      parts = Mp4.combine(fetched[0].init, fetched[0].segments, fetched[1].init, fetched[1].segments);
    } else {
      const only = fetched[0];
      parts = only.init ? Mp4.single(only.init, only.segments) : only.segments;
    }
    const info = choice !== 'audio' && plan.video && plan.video.height ? `${plan.video.height}p` : '';
    return { parts, info };
  }

  // ---------- UI ----------

  let host = null;
  let ui = null;
  let mode = 'media';

  const STYLE = `
    :host { all: initial; }
    * { box-sizing: border-box; font-family: "Segoe UI", system-ui, -apple-system, sans-serif; }
    .dock { position: fixed; right: 20px; bottom: 20px; z-index: 2147483646; display: flex; flex-direction: column; align-items: flex-end; gap: 8px; }
    .launcher {
      display: flex; align-items: center; gap: 8px;
      padding: 10px 16px; border: 0; border-radius: 999px; cursor: pointer;
      background: #1d4ed8; color: #fff; font-size: 13px; font-weight: 600;
      box-shadow: 0 6px 20px rgba(29, 78, 216, .35);
    }
    .launcher:hover { background: #1e40af; }
    .launcher svg { width: 16px; height: 16px; }
    .menu {
      display: none; min-width: 230px; padding: 6px; background: #fff; border-radius: 12px;
      box-shadow: 0 12px 32px rgba(15, 23, 42, .25); border: 1px solid #e2e8f0;
    }
    .menu.open { display: block; }
    .menu button {
      display: flex; width: 100%; align-items: center; gap: 10px; padding: 9px 10px; border: 0; border-radius: 8px;
      background: transparent; color: #0f172a; font-size: 14px; text-align: left; cursor: pointer;
    }
    .menu button:hover, .menu button:focus-visible { background: #eff6ff; outline: none; }
    .menu .sep { height: 1px; background: #e2e8f0; margin: 4px 6px; }
    .menu svg { width: 18px; height: 18px; color: #1d4ed8; flex: none; }
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
    .sub { margin: 0 0 14px; color: #64748b; font-size: 13px; word-break: break-word; }
    label.field { display: block; font-size: 12px; font-weight: 600; color: #334155; margin: 14px 0 6px; }
    select, input[type=text] {
      width: 100%; padding: 9px 10px; border: 1px solid #cbd5e1; border-radius: 8px;
      font-size: 14px; color: #0f172a; background: #fff;
    }
    .choices { display: grid; gap: 8px; }
    .choice {
      display: flex; gap: 10px; align-items: flex-start; padding: 10px 12px; border: 1px solid #cbd5e1;
      border-radius: 10px; cursor: pointer;
    }
    .choice:has(input:checked) { border-color: #1d4ed8; background: #eff6ff; }
    .choice input { margin-top: 3px; }
    .choice b { display: block; font-size: 14px; }
    .choice span { font-size: 12px; color: #64748b; }
    .checks { display: flex; gap: 18px; margin-top: 12px; font-size: 13px; color: #334155; }
    .checks label { display: flex; align-items: center; gap: 6px; cursor: pointer; }
    .inline { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; font-size: 13px; color: #334155; margin-top: 12px; }
    .inline select { width: auto; padding: 5px 8px; }
    .muted { color: #64748b; }
    pre {
      margin: 0; padding: 12px; height: 180px; overflow: auto; white-space: pre-wrap;
      background: #f1f5f9; border-radius: 8px; font: 12px/1.5 ui-monospace, Consolas, monospace; color: #1e293b;
    }
    .bar { height: 8px; background: #e2e8f0; border-radius: 99px; overflow: hidden; margin-top: 14px; display: none; }
    .bar.show { display: block; }
    .bar i { display: block; height: 100%; width: 0; background: #1d4ed8; transition: width .2s; }
    .status { font-size: 13px; color: #64748b; margin-top: 10px; min-height: 18px; }
    .status.error { color: #b91c1c; }
    .row { display: flex; justify-content: flex-end; gap: 10px; margin-top: 18px; }
    button.btn {
      padding: 9px 16px; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer;
      border: 1px solid #cbd5e1; background: #fff; color: #0f172a;
    }
    button.primary { background: #1d4ed8; border-color: #1d4ed8; color: #fff; }
    button.primary:disabled { opacity: .5; cursor: default; }
    [hidden] { display: none !important; }
  `;

  const ICONS = {
    down: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v11"/><path d="m7 10 5 5 5-5"/><path d="M5 20h14"/></svg>',
    both: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="13" height="14" rx="2"/><path d="m16 10 5-3v10l-5-3"/></svg>',
    audio: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18V5l12-2v13"/><circle cx="6" cy="18" r="3"/><circle cx="18" cy="16" r="3"/></svg>',
    video: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 3 18 18"/></svg>',
    text: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6h16M4 12h16M4 18h10"/></svg>'
  };

  function build() {
    host = document.createElement('div');
    host.id = 'ms-downloader-root';
    const root = host.attachShadow({ mode: 'open' });
    const mediaRadios = Object.entries(MEDIA_CHOICES).map(([id, c]) => `
      <label class="choice"><input type="radio" name="media" value="${id}"><div><b>${c.label}</b><span>${c.hint}</span></div></label>`).join('');
    root.innerHTML = `
      <style>${STYLE}</style>
      <div class="dock">
        <div class="menu" role="menu">
          <button type="button" role="menuitem" data-pick="both">${ICONS.both}Video + audio</button>
          <button type="button" role="menuitem" data-pick="audio">${ICONS.audio}Audio only</button>
          <button type="button" role="menuitem" data-pick="video">${ICONS.video}Video only</button>
          <div class="sep"></div>
          <button type="button" role="menuitem" data-pick="transcript">${ICONS.text}Transcript…</button>
        </div>
        <button class="launcher" type="button" aria-haspopup="menu" title="MS Downloader">${ICONS.down}<span>Download</span></button>
      </div>
      <div class="backdrop" role="dialog" aria-modal="true" aria-labelledby="msdl-title">
        <div class="dialog">
          <h2 id="msdl-title"></h2>
          <p class="sub"></p>

          <section data-mode="media">
            <div class="choices">${mediaRadios}</div>
            <div class="inline">Parallel downloads
              <select id="msdl-parallel"><option>2</option><option>4</option><option>8</option><option>12</option><option>16</option></select>
              <span class="muted">higher is faster but may be throttled</span>
            </div>
          </section>

          <section data-mode="transcript">
            <label class="field" for="msdl-format">Format</label>
            <select id="msdl-format"></select>
            <div class="checks">
              <label><input type="checkbox" id="msdl-speakers"> Speaker names</label>
              <label><input type="checkbox" id="msdl-times"> Timestamps</label>
            </div>
          </section>

          <label class="field" for="msdl-name">File name</label>
          <input type="text" id="msdl-name" spellcheck="false">

          <section data-mode="transcript">
            <label class="field">Preview</label>
            <pre id="msdl-preview"></pre>
          </section>

          <div class="bar"><i></i></div>
          <div class="status" id="msdl-status"></div>
          <div class="row">
            <button class="btn" type="button" id="msdl-cancel">Close</button>
            <button class="btn primary" type="button" id="msdl-save">Download</button>
          </div>
        </div>
      </div>`;

    const $ = (s) => root.querySelector(s);
    ui = {
      menu: $('.menu'),
      launcher: $('.launcher'),
      backdrop: $('.backdrop'),
      title: $('#msdl-title'),
      sub: $('.sub'),
      sections: [...root.querySelectorAll('section[data-mode]')],
      radios: [...root.querySelectorAll('input[name=media]')],
      parallel: $('#msdl-parallel'),
      format: $('#msdl-format'),
      speakers: $('#msdl-speakers'),
      times: $('#msdl-times'),
      name: $('#msdl-name'),
      preview: $('#msdl-preview'),
      bar: $('.bar'),
      barFill: $('.bar i'),
      status: $('#msdl-status'),
      cancel: $('#msdl-cancel'),
      save: $('#msdl-save')
    };

    for (const [id, f] of Object.entries(T.FORMATS)) {
      const opt = document.createElement('option');
      opt.value = id;
      opt.textContent = f.label;
      ui.format.appendChild(opt);
    }

    ui.launcher.addEventListener('click', (e) => { e.stopPropagation(); ui.menu.classList.toggle('open'); });
    ui.menu.addEventListener('click', (e) => {
      e.stopPropagation();
      const b = e.target.closest('button[data-pick]');
      if (!b) return;
      ui.menu.classList.remove('open');
      if (b.dataset.pick === 'transcript') openTranscript();
      else openMedia(b.dataset.pick);
    });
    document.addEventListener('click', () => ui.menu.classList.remove('open'));
    ui.cancel.addEventListener('click', onCancel);
    ui.backdrop.addEventListener('click', (e) => { if (e.target === ui.backdrop && !state.job) closeDialog(); });
    host.shadowRoot.addEventListener('keydown', (e) => {
      if (e.key !== 'Escape') return;
      ui.menu.classList.remove('open');
      if (!state.job) closeDialog();
    });
    ui.radios.forEach((r) => r.addEventListener('change', () => { state.prefs.media = r.value; savePrefs(); }));
    ui.parallel.addEventListener('change', () => { state.prefs.parallel = Number(ui.parallel.value); savePrefs(); });
    ui.format.addEventListener('change', onTranscriptOptions);
    ui.speakers.addEventListener('change', onTranscriptOptions);
    ui.times.addEventListener('change', onTranscriptOptions);
    ui.save.addEventListener('click', () => (mode === 'media' ? startMedia() : saveTranscript()));

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

  function setProgress(fraction) {
    ui.bar.classList.toggle('show', fraction != null);
    ui.barFill.style.width = `${Math.round((fraction || 0) * 100)}%`;
  }

  function safeName(name) {
    return String(name || 'recording')
      .replace(/[\\/:*?"<>|\u0000-\u001f]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 150) || 'recording';
  }

  function fileName() {
    return (state.file && state.file.name) || document.title || 'recording';
  }

  function openDialog(which, title) {
    mode = which;
    ui.sections.forEach((s) => { s.hidden = s.dataset.mode !== which; });
    ui.title.textContent = title;
    ui.sub.textContent = fileName();
    ui.name.value = safeName(fileName());
    ui.cancel.textContent = 'Close';
    setProgress(null);
    setStatus('');
    ui.backdrop.classList.add('open');
  }

  function closeDialog() {
    ui.backdrop.classList.remove('open');
  }

  function onCancel() {
    if (state.job) state.job.controller.abort();
    else closeDialog();
  }

  function saveBlob(parts, name, mime) {
    const blob = new Blob(parts, { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.documentElement.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    return blob.size;
  }

  function mb(bytes) {
    return bytes >= 1e9 ? `${(bytes / 1e9).toFixed(2)} GB` : `${(bytes / 1e6).toFixed(1)} MB`;
  }

  // ----- media dialog -----

  function openMedia(choice) {
    state.prefs.media = choice;
    savePrefs();
    openDialog('media', 'Download recording');
    ui.radios.forEach((r) => { r.checked = r.value === choice; r.disabled = false; });
    ui.parallel.value = String(state.prefs.parallel || 4);
    ui.save.disabled = false;
    if (!findManifest()) setStatus('Tip: start playing the video for a few seconds first, so the player loads it.');
  }

  async function startMedia() {
    const choice = state.prefs.media;
    const meta = MEDIA_CHOICES[choice];
    const controller = new AbortController();
    state.job = { controller };
    ui.save.disabled = true;
    ui.radios.forEach((r) => { r.disabled = true; });
    ui.parallel.disabled = true;
    ui.cancel.textContent = 'Cancel';
    setProgress(0);
    const started = Date.now();

    try {
      const result = await downloadMedia(choice, (done, total, bytes, note) => {
        setProgress(total ? done / total : 0);
        setStatus(note || `Downloading… ${done} / ${total} parts · ${mb(bytes)}`);
      }, controller.signal);
      const name = `${safeName(ui.name.value)}.${meta.ext}`;
      const size = saveBlob(result.parts, name, meta.mime);
      setProgress(1);
      const secs = Math.round((Date.now() - started) / 1000);
      setStatus(`Saved "${name}" (${mb(size)}${result.info ? ', ' + result.info : ''}) in ${secs} s. Check your Downloads folder.`);
    } catch (e) {
      setProgress(null);
      if (controller.signal.aborted) setStatus('Cancelled.');
      else setStatus(explainMediaError(e), true);
      console.warn('[MS Downloader]', e);
    } finally {
      state.job = null;
      ui.save.disabled = false;
      ui.radios.forEach((r) => { r.disabled = false; });
      ui.parallel.disabled = false;
      ui.cancel.textContent = 'Close';
    }
  }

  function explainMediaError(e) {
    switch (e && e.code) {
      case 'nomanifest': return 'The video is not loaded yet. Press play, wait a few seconds, then try again.';
      case 'drm': return 'This recording is copy-protected (DRM) by Microsoft. It can only be played, not saved.';
      case 'noaudio': return 'This recording has no separate sound track.';
      case 'novideo': return 'No video stream was found in this recording.';
      default: break;
    }
    if (e && (e.status === 401 || e.status === 403)) {
      return `Microsoft refused the request (${e.message}). Refresh the page, play the video for a few seconds and try again.`;
    }
    if (e && e.code === 'network') {
      return `The browser could not reach ${e.host}. Refresh the page, play the video for a few seconds and try again.`;
    }
    return `Download failed: ${e && e.message ? e.message : e}. Refresh the page and try again.`;
  }

  // ----- transcript dialog -----

  function refreshTranscriptControls() {
    const fmt = state.prefs.format;
    ui.format.value = fmt;
    ui.speakers.checked = state.prefs.speakers;
    ui.times.checked = state.prefs.timestamps;
    ui.speakers.disabled = fmt === 'json';
    ui.times.disabled = fmt !== 'txt';
  }

  function renderPreview() {
    if (!state.transcript) return;
    const out = T.render(state.prefs.format, state.transcript, state.prefs);
    const lines = out.split('\n');
    ui.preview.textContent = lines.slice(0, 40).join('\n') + (lines.length > 40 ? '\n…' : '');
  }

  function onTranscriptOptions() {
    state.prefs.format = ui.format.value;
    state.prefs.speakers = ui.speakers.checked;
    state.prefs.timestamps = ui.times.checked;
    savePrefs();
    refreshTranscriptControls();
    renderPreview();
  }

  async function openTranscript() {
    openDialog('transcript', 'Download transcript');
    ui.preview.textContent = '';
    ui.save.disabled = true;
    refreshTranscriptControls();
    setStatus('Loading transcript…');
    try {
      const data = await loadTranscript();
      const count = T.normalize(data).length;
      if (!count) throw Object.assign(new Error('empty'), { code: 'none' });
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
        setStatus(`Could not load the transcript: ${e.message}. Refresh the page and try again.`, true);
      }
    }
  }

  function saveTranscript() {
    if (!state.transcript) return;
    const fmt = T.FORMATS[state.prefs.format] || T.FORMATS.txt;
    const text = T.render(state.prefs.format, state.transcript, state.prefs);
    const body = state.prefs.format === 'json' ? text : '\ufeff' + text;
    saveBlob([body], `${safeName(ui.name.value)}.${fmt.ext}`, `${fmt.mime};charset=utf-8`);
    setStatus('Saved to your Downloads folder.');
  }
})();
