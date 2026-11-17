// MS Downloader — transcript formatting.
// Pure functions with no browser APIs, so they can be tested with Node.

(function (root) {
  'use strict';

  // SharePoint writes offsets like "00:01:02.3450000" (hh:mm:ss.fffffff).
  // Some responses use ISO-8601 durations such as "PT1M2.345S". Accept both.
  function toSeconds(value) {
    if (typeof value === 'number' && isFinite(value)) return value;
    if (!value) return 0;
    const text = String(value).trim();

    const iso = text.match(/^P(?:(\d+(?:\.\d+)?)D)?T?(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?$/i);
    if (iso && /\d/.test(text)) {
      const [, d = 0, h = 0, m = 0, s = 0] = iso;
      return Number(d) * 86400 + Number(h) * 3600 + Number(m) * 60 + Number(s);
    }

    // [d.]hh:mm:ss[.fraction]
    const clock = text.match(/^(?:(\d+)\.)?(\d+):(\d{1,2}):(\d{1,2})(?:[.,](\d+))?$/);
    if (clock) {
      const [, d = 0, h, m, s, frac = '0'] = clock;
      return Number(d) * 86400 + Number(h) * 3600 + Number(m) * 60 + Number(s) + Number('0.' + frac);
    }
    const plain = Number(text);
    return isFinite(plain) ? plain : 0;
  }

  function pad(n, width) {
    return String(n).padStart(width, '0');
  }

  // style: 'vtt' -> 00:00:01.234, 'srt' -> 00:00:01,234, 'short' -> 00:00:01
  function clock(seconds, style) {
    const totalMs = Math.max(0, Math.round(seconds * 1000));
    const h = Math.floor(totalMs / 3600000);
    const m = Math.floor((totalMs % 3600000) / 60000);
    const s = Math.floor((totalMs % 60000) / 1000);
    const ms = totalMs % 1000;
    const base = `${pad(h, 2)}:${pad(m, 2)}:${pad(s, 2)}`;
    if (style === 'short') return base;
    return base + (style === 'srt' ? ',' : '.') + pad(ms, 3);
  }

  // Turn whatever SharePoint returned into a clean, ordered list of lines.
  function normalize(data) {
    const raw = (data && (data.entries || data.value || data.transcript)) || [];
    const lines = [];
    for (const entry of raw) {
      if (!entry) continue;
      const text = String(entry.text || entry.content || '').replace(/\s+/g, ' ').trim();
      if (!text) continue;
      const start = toSeconds(entry.startOffset ?? entry.start ?? entry.startTime);
      let end = toSeconds(entry.endOffset ?? entry.end ?? entry.endTime);
      if (end <= start) end = start + 2;
      const speaker = String(entry.speakerDisplayName || entry.speakerName || entry.speaker || '').trim();
      lines.push({ start, end, speaker, text });
    }
    lines.sort((a, b) => a.start - b.start);
    return lines;
  }

  // Merge back-to-back lines from the same person into one paragraph.
  function paragraphs(lines) {
    const out = [];
    for (const line of lines) {
      const prev = out[out.length - 1];
      if (prev && prev.speaker === line.speaker && line.start - prev.end < 10) {
        prev.text += ' ' + line.text;
        prev.end = Math.max(prev.end, line.end);
      } else {
        out.push({ ...line });
      }
    }
    return out;
  }

  function escapeCue(text) {
    return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function toVtt(lines, opts) {
    const withSpeakers = !opts || opts.speakers !== false;
    const parts = ['WEBVTT', ''];
    lines.forEach((l, i) => {
      parts.push(String(i + 1));
      parts.push(`${clock(l.start, 'vtt')} --> ${clock(l.end, 'vtt')}`);
      const body = escapeCue(l.text);
      parts.push(withSpeakers && l.speaker ? `<v ${escapeCue(l.speaker)}>${body}` : body);
      parts.push('');
    });
    return parts.join('\n');
  }

  function toSrt(lines, opts) {
    const withSpeakers = !opts || opts.speakers !== false;
    return lines.map((l, i) => [
      String(i + 1),
      `${clock(l.start, 'srt')} --> ${clock(l.end, 'srt')}`,
      withSpeakers && l.speaker ? `${l.speaker}: ${l.text}` : l.text,
      ''
    ].join('\n')).join('\n');
  }

  function toText(lines, opts) {
    const withSpeakers = !opts || opts.speakers !== false;
    const withTimes = !opts || opts.timestamps !== false;
    return paragraphs(lines).map((p) => {
      const head = [];
      if (withTimes) head.push(`[${clock(p.start, 'short')}]`);
      if (withSpeakers && p.speaker) head.push(p.speaker + ':');
      return (head.length ? head.join(' ') + ' ' : '') + p.text;
    }).join('\n\n') + '\n';
  }

  const FORMATS = {
    txt: { label: 'Readable text (.txt)', ext: 'txt', mime: 'text/plain' },
    vtt: { label: 'Subtitles – WebVTT (.vtt)', ext: 'vtt', mime: 'text/vtt' },
    srt: { label: 'Subtitles – SubRip (.srt)', ext: 'srt', mime: 'application/x-subrip' },
    json: { label: 'Original data (.json)', ext: 'json', mime: 'application/json' }
  };

  function render(format, data, opts) {
    if (format === 'json') return JSON.stringify(data, null, 2);
    const lines = normalize(data);
    if (format === 'vtt') return toVtt(lines, opts);
    if (format === 'srt') return toSrt(lines, opts);
    return toText(lines, opts);
  }

  const api = { toSeconds, clock, normalize, paragraphs, toVtt, toSrt, toText, render, FORMATS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.MSDLTranscript = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
