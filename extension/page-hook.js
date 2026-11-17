// MS Downloader — page-world helper.
//
// Runs in the page's own JavaScript world so it can see two things the
// isolated extension world cannot:
//   1. window.g_fileInfo, which SharePoint fills in with the drive and item
//      ids of the file being played.
//   2. The Authorization header the player attaches to its own SharePoint
//      API calls (some tenants refuse cookie-only API requests).
// Everything found is handed to panel.js with window.postMessage.

(() => {
  'use strict';

  const CHANNEL = 'ms-downloader';
  const API_PATH = /\/_api\/v2\.[01]\//i;

  let lastBearer = null;
  let lastItemKey = null;

  function send(kind, payload) {
    window.postMessage({ channel: CHANNEL, kind, payload }, window.location.origin);
  }

  function rememberBearer(url, value) {
    if (!value || !/^Bearer\s/i.test(value)) return;
    if (!url || !API_PATH.test(String(url))) return;
    if (value === lastBearer) return;
    lastBearer = value;
    send('bearer', { value });
  }

  function headerFrom(headers, name) {
    if (!headers) return null;
    try {
      if (headers instanceof Headers) return headers.get(name);
      if (Array.isArray(headers)) {
        const hit = headers.find(([k]) => String(k).toLowerCase() === name);
        return hit ? hit[1] : null;
      }
      for (const key of Object.keys(headers)) {
        if (key.toLowerCase() === name) return headers[key];
      }
    } catch (_) { /* ignore odd header shapes */ }
    return null;
  }

  // Watch fetch() for the player's authorized API calls.
  const realFetch = window.fetch;
  if (typeof realFetch === 'function') {
    window.fetch = function (input, init) {
      try {
        const url = typeof input === 'string' ? input : input && input.url;
        const auth = headerFrom(init && init.headers, 'authorization') ||
          (input instanceof Request ? input.headers.get('authorization') : null);
        rememberBearer(url, auth);
      } catch (_) { /* never break the page */ }
      return realFetch.apply(this, arguments);
    };
  }

  // Same for XMLHttpRequest, which parts of SharePoint still use.
  const xhrProto = XMLHttpRequest.prototype;
  const realOpen = xhrProto.open;
  const realSetHeader = xhrProto.setRequestHeader;
  xhrProto.open = function (method, url) {
    this.__msdlUrl = url;
    return realOpen.apply(this, arguments);
  };
  xhrProto.setRequestHeader = function (name, value) {
    try {
      if (String(name).toLowerCase() === 'authorization') rememberBearer(this.__msdlUrl, value);
    } catch (_) { /* ignore */ }
    return realSetHeader.apply(this, arguments);
  };

  // Read the file identity out of g_fileInfo. SharePoint is a single-page
  // app, so keep checking: the user can move to another recording.
  function readFileInfo() {
    const info = window.g_fileInfo;
    if (!info) return null;
    const itemUrl = info['.spItemUrl'];
    if (!itemUrl) return null;
    let parsed;
    try { parsed = new URL(itemUrl); } catch (_) { return null; }
    const match = parsed.pathname.match(
      /^(.*?)\/_api\/v2\.\d+\/drives\/([^/]+)\/items\/([^/?#]+)/i
    );
    if (!match) return null;
    const rawName = info.displayName || info.name || info.title || document.title || 'recording';
    return {
      sitePath: match[1],
      driveId: match[2],
      itemId: match[3],
      name: String(rawName).replace(/\.(mp4|mov|m4a|webm)$/i, ''),
      hasTranscripts: typeof info.hasTranscripts === 'boolean' ? info.hasTranscripts : null
    };
  }

  function checkFileInfo() {
    const file = readFileInfo();
    if (!file) return;
    const key = file.driveId + '/' + file.itemId;
    if (key === lastItemKey) return;
    lastItemKey = key;
    send('file', file);
  }

  // panel.js loads later than this script; it asks for a resend on startup.
  window.addEventListener('message', (event) => {
    if (event.source !== window) return;
    const data = event.data;
    if (!data || data.channel !== CHANNEL || data.kind !== 'hello') return;
    lastItemKey = null;
    checkFileInfo();
    if (lastBearer) send('bearer', { value: lastBearer });
  });

  setInterval(checkFileInfo, 1500);
  document.addEventListener('DOMContentLoaded', checkFileInfo);
})();
