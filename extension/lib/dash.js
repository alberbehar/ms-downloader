// MS Downloader — Copyright 2026 alberbehar
// Licensed under the Apache License, Version 2.0. See LICENSE and NOTICE.
// MS Downloader — reading the player's DASH manifest (MPD).
//
// Produces, for each stream (video, audio or a combined one), the address of
// its initialization segment and of every media segment, using the highest
// quality the manifest offers. Standard MPEG-DASH, nothing SharePoint-only.

(function (root) {
  'use strict';

  // Children by local name, ignoring XML namespaces.
  function kids(el, name) {
    return Array.from(el ? el.children : []).filter((c) => c.localName === name);
  }
  function kid(el, name) {
    return kids(el, name)[0] || null;
  }
  function num(v, fallback) {
    const n = Number(v);
    return v != null && v !== '' && isFinite(n) ? n : fallback;
  }

  // ISO-8601 duration ("PT1H2M3.5S") in seconds.
  function isoSeconds(text) {
    const m = String(text || '').match(/^P(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/);
    if (!m) return 0;
    return num(m[1], 0) * 86400 + num(m[2], 0) * 3600 + num(m[3], 0) * 60 + num(m[4], 0);
  }

  function fillTemplate(template, vars) {
    return template.replace(/\$(RepresentationID|Bandwidth|Number|Time)(?:%0(\d+)d)?\$/g, (_, key, width) => {
      const value = String(vars[key]);
      return width ? value.padStart(Number(width), '0') : value;
    }).replace(/\$\$/g, '$');
  }

  // Resolve nested <BaseURL> elements from the outside in.
  function baseFor(chain, start) {
    let base = start;
    for (const el of chain) {
      const b = kid(el, 'BaseURL');
      if (b && b.textContent.trim()) base = new URL(b.textContent.trim(), base).href;
    }
    return base;
  }

  function safeUrl(href) {
    const u = new URL(href);
    if (u.protocol !== 'https:' && u.protocol !== 'http:') throw new Error('Unexpected address in manifest');
    return u.href;
  }

  const DRM_SCHEMES = /edef8ba9|9a04f079|94ce86fb|1077efec|e2719d58/i; // Widevine, PlayReady, FairPlay, ClearKey(CENC), Marlin

  function streamType(set, rep) {
    const ct = set.getAttribute('contentType');
    if (ct === 'video' || ct === 'audio') return ct;
    const mime = rep.getAttribute('mimeType') || set.getAttribute('mimeType') || '';
    if (mime.startsWith('video')) return 'video';
    if (mime.startsWith('audio')) return 'audio';
    const codecs = rep.getAttribute('codecs') || set.getAttribute('codecs') || '';
    if (/avc|hev|hvc|vp0?9|av01/i.test(codecs)) return 'video';
    if (/mp4a|opus|ac-3|ec-3/i.test(codecs)) return 'audio';
    return 'other';
  }

  function segmentEncryption(set, rep) {
    const nodes = [...kids(rep, 'ContentProtection'), ...kids(set, 'ContentProtection')];
    for (const cp of nodes) {
      if ((cp.getAttribute('schemeIdUri') || '').toLowerCase() !== 'urn:mpeg:dash:sea:2012') continue;
      const se = kid(cp, 'SegmentEncryption');
      const cpd = kid(cp, 'CryptoPeriod');
      const scheme = (se && se.getAttribute('schemeIdUri')) || '';
      const keyUri = cpd && cpd.getAttribute('keyUriTemplate');
      const ivHex = ((cpd && cpd.getAttribute('IV')) || '').replace(/^0x/i, '');
      if (/aes128-cbc/i.test(scheme) && keyUri && /^[0-9a-f]{32}$/i.test(ivHex)) {
        return { keyUri, iv: ivHex };
      }
    }
    return null;
  }

  function parse(xmlText, manifestUrl, DOMParserImpl) {
    const Parser = DOMParserImpl || root.DOMParser;
    const doc = new Parser().parseFromString(xmlText, 'application/xml');
    const mpd = doc.documentElement;
    if (!mpd || mpd.localName !== 'MPD') throw new Error('Not a DASH manifest');

    const totalSeconds = isoSeconds(mpd.getAttribute('mediaPresentationDuration'));
    const period = kid(mpd, 'Period');
    if (!period) throw new Error('Manifest has no Period');
    const periodSeconds = isoSeconds(period.getAttribute('duration')) || totalSeconds;

    const sets = kids(period, 'AdaptationSet');
    const streams = [];
    let drm = false;

    for (const set of sets) {
      for (const cp of [...kids(set, 'ContentProtection'), ...set.getElementsByTagName('*')]) {
        if (cp.localName === 'ContentProtection' && DRM_SCHEMES.test(cp.getAttribute('schemeIdUri') || '')) drm = true;
      }

      const reps = kids(set, 'Representation');
      if (!reps.length) continue;
      // Best quality first: tallest picture, then highest bitrate.
      reps.sort((a, b) =>
        num(b.getAttribute('height'), 0) - num(a.getAttribute('height'), 0) ||
        num(b.getAttribute('bandwidth'), 0) - num(a.getAttribute('bandwidth'), 0));
      const rep = reps[0];

      const tpl = kid(rep, 'SegmentTemplate') || kid(set, 'SegmentTemplate') || kid(period, 'SegmentTemplate');
      if (!tpl) continue;

      const base = baseFor([mpd, period, set, rep], manifestUrl);
      const vars = { RepresentationID: rep.getAttribute('id') || '', Bandwidth: rep.getAttribute('bandwidth') || '' };
      const startNumber = num(tpl.getAttribute('startNumber'), 1);
      const timescale = num(tpl.getAttribute('timescale'), 1);
      const media = tpl.getAttribute('media') || '';
      const initTpl = tpl.getAttribute('initialization') || '';

      const segments = [];
      const timeline = kid(tpl, 'SegmentTimeline');
      if (timeline) {
        let t = 0;
        let n = startNumber;
        const entries = kids(timeline, 'S');
        entries.forEach((s, idx) => {
          if (s.hasAttribute('t')) t = num(s.getAttribute('t'), t);
          const d = num(s.getAttribute('d'), 0);
          let r = num(s.getAttribute('r'), 0);
          if (r < 0) {
            // Repeat until the next entry's start, or the end of the period.
            const next = entries[idx + 1];
            const end = next && next.hasAttribute('t') ? num(next.getAttribute('t'), 0) : periodSeconds * timescale;
            r = d > 0 ? Math.max(0, Math.ceil((end - t) / d) - 1) : 0;
          }
          for (let i = 0; i <= r; i++) {
            segments.push(safeUrl(new URL(fillTemplate(media, { ...vars, Number: n, Time: t }), base).href));
            t += d;
            n += 1;
          }
        });
      } else {
        const d = num(tpl.getAttribute('duration'), 0);
        const count = d > 0 && periodSeconds > 0 ? Math.ceil((periodSeconds * timescale) / d) : 0;
        for (let i = 0; i < count; i++) {
          segments.push(safeUrl(new URL(fillTemplate(media, { ...vars, Number: startNumber + i, Time: i * d }), base).href));
        }
      }
      if (!segments.length) continue;

      streams.push({
        type: streamType(set, rep),
        codecs: rep.getAttribute('codecs') || set.getAttribute('codecs') || '',
        width: num(rep.getAttribute('width'), 0),
        height: num(rep.getAttribute('height'), 0),
        bandwidth: num(rep.getAttribute('bandwidth'), 0),
        init: initTpl ? safeUrl(new URL(fillTemplate(initTpl, { ...vars, Number: startNumber, Time: 0 }), base).href) : null,
        segments,
        encryption: segmentEncryption(set, rep)
      });
    }

    const video = streams.find((s) => s.type === 'video') || null;
    const audio = streams.find((s) => s.type === 'audio') || null;
    // One stream that is neither, or the only stream: treat as already combined.
    const combined = !audio && streams.length === 1 ? streams[0] : null;
    return { drm, seconds: totalSeconds || periodSeconds, video, audio, combined };
  }

  const api = { parse, fillTemplate, isoSeconds };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.MSDLDash = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
