// MS Downloader — joining fragmented MP4 streams.
//
// DASH delivers video and audio as separate fragmented MP4 streams: one
// initialization segment (ftyp + moov) and many media segments (moof + mdat).
// To make one playable file we:
//   - keep the video moov and add the audio track to it as track 2,
//   - give every audio fragment the new track number,
//   - drop per-segment index boxes (styp, sidx) whose offsets no longer apply,
//   - put all fragments in presentation order and renumber them.

(function (root) {
  'use strict';

  const CONTAINERS = new Set(['moov', 'trak', 'mdia', 'minf', 'stbl', 'mvex', 'moof', 'traf', 'edts', 'dinf', 'udta']);

  function u32(b, o) { return ((b[o] << 24) >>> 0) + (b[o + 1] << 16) + (b[o + 2] << 8) + b[o + 3]; }
  function setU32(b, o, v) { b[o] = (v >>> 24) & 255; b[o + 1] = (v >>> 16) & 255; b[o + 2] = (v >>> 8) & 255; b[o + 3] = v & 255; }
  function type4(b, o) { return String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]); }

  // List the boxes in bytes[start, end).
  function boxes(bytes, start = 0, end = bytes.length) {
    const out = [];
    let p = start;
    while (p + 8 <= end) {
      let size = u32(bytes, p);
      const type = type4(bytes, p + 4);
      let header = 8;
      if (size === 1) {
        if (p + 16 > end) break;
        size = u32(bytes, p + 8) * 4294967296 + u32(bytes, p + 12);
        header = 16;
      } else if (size === 0) {
        size = end - p;
      }
      if (size < header || p + size > end) break;
      out.push({ type, start: p, end: p + size, body: p + header });
      p += size;
    }
    return out;
  }

  function child(bytes, box, type) {
    return boxes(bytes, box.body, box.end).find((b) => b.type === type) || null;
  }

  function path(bytes, box, ...types) {
    let cur = box;
    for (const t of types) {
      if (!cur) return null;
      cur = child(bytes, cur, t);
    }
    return cur;
  }

  function slice(bytes, box) {
    return bytes.slice(box.start, box.end);
  }

  function makeBox(type, parts) {
    const size = 8 + parts.reduce((s, p) => s + p.length, 0);
    const out = new Uint8Array(size);
    setU32(out, 0, size);
    for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
    let o = 8;
    for (const p of parts) { out.set(p, o); o += p.length; }
    return out;
  }

  // ---- facts read from an init segment ----

  function readInit(bytes) {
    const top = boxes(bytes);
    const ftyp = top.find((b) => b.type === 'ftyp') || null;
    const moov = top.find((b) => b.type === 'moov');
    if (!moov) throw new Error('Init segment has no moov box');
    const traks = boxes(bytes, moov.body, moov.end).filter((b) => b.type === 'trak');
    if (!traks.length) throw new Error('Init segment has no track');
    const tracks = traks.map((trak) => {
      const tkhd = child(bytes, trak, 'tkhd');
      const version = bytes[tkhd.body];
      const idOffset = tkhd.body + (version === 1 ? 20 : 12);
      const mdhd = path(bytes, trak, 'mdia', 'mdhd');
      const mv = bytes[mdhd.body];
      const timescale = u32(bytes, mdhd.body + (mv === 1 ? 20 : 12));
      return { trak, id: u32(bytes, idOffset), timescale };
    });
    return { bytes, ftyp, moov, tracks };
  }

  // Change the track number inside a copied trak box.
  function retrack(trakBytes, newId) {
    const copy = trakBytes.slice();
    const trak = boxes(copy)[0];
    const tkhd = child(copy, trak, 'tkhd');
    setU32(copy, tkhd.body + (copy[tkhd.body] === 1 ? 20 : 12), newId);
    return copy;
  }

  function retrex(trexBytes, newId) {
    const copy = trexBytes.slice();
    const trex = boxes(copy)[0];
    setU32(copy, trex.body + 4, newId);
    return copy;
  }

  function setNextTrackId(mvhdBytes, next) {
    const copy = mvhdBytes.slice();
    const mvhd = boxes(copy)[0];
    const v = copy[mvhd.body];
    setU32(copy, mvhd.body + (v === 1 ? 108 : 96), next);
    return copy;
  }

  // ftyp + moov holding the video track(s) and the audio track.
  function joinInits(video, audio, audioId) {
    const vb = video.bytes;
    const ab = audio.bytes;
    const moovParts = [];
    const trexParts = [];
    for (const b of boxes(vb, video.moov.body, video.moov.end)) {
      if (b.type === 'mvhd') moovParts.push(setNextTrackId(slice(vb, b), audioId + 1));
      else if (b.type === 'mvex') {
        for (const m of boxes(vb, b.body, b.end)) trexParts.push(slice(vb, m));
      } else moovParts.push(slice(vb, b));
    }
    moovParts.push(retrack(slice(ab, audio.tracks[0].trak), audioId));
    const audioMvex = boxes(ab, audio.moov.body, audio.moov.end).find((b) => b.type === 'mvex');
    if (audioMvex) {
      for (const m of boxes(ab, audioMvex.body, audioMvex.end)) {
        if (m.type === 'trex') trexParts.push(retrex(slice(ab, m), audioId));
      }
    }
    moovParts.push(makeBox('mvex', trexParts));
    const ftyp = video.ftyp ? slice(vb, video.ftyp) : new Uint8Array(0);
    return [ftyp, makeBox('moov', moovParts)];
  }

  // ---- media segments ----

  // Split a media segment into (moof, mdat) pairs, noting each pair's start
  // time so audio and video can be interleaved.
  function fragments(bytes, timescale, newTrackId) {
    const out = [];
    const top = boxes(bytes);
    for (let i = 0; i < top.length; i++) {
      const box = top[i];
      if (box.type !== 'moof') continue;
      const moof = slice(bytes, box);
      const mb = boxes(moof)[0];
      let time = 0;
      for (const traf of boxes(moof, mb.body, mb.end).filter((b) => b.type === 'traf')) {
        const tfhd = child(moof, traf, 'tfhd');
        if (tfhd && newTrackId) setU32(moof, tfhd.body + 4, newTrackId);
        const tfdt = child(moof, traf, 'tfdt');
        if (tfdt) {
          const v = moof[tfdt.body];
          time = v === 1
            ? u32(moof, tfdt.body + 4) * 4294967296 + u32(moof, tfdt.body + 8)
            : u32(moof, tfdt.body + 4);
        }
      }
      const parts = [moof];
      // Everything up to the next moof belongs to this fragment (mdat, and
      // rarely extra boxes); index boxes are left out.
      let j = i + 1;
      while (j < top.length && top[j].type !== 'moof') {
        if (top[j].type === 'mdat') parts.push(bytes.subarray(top[j].start, top[j].end));
        j++;
      }
      out.push({ seconds: time / (timescale || 1), parts, moof });
    }
    return out;
  }

  function renumber(frags) {
    frags.forEach((f, i) => {
      const mb = boxes(f.moof)[0];
      const mfhd = child(f.moof, mb, 'mfhd');
      if (mfhd) setU32(f.moof, mfhd.body + 4, i + 1);
    });
  }

  // One stream on its own (video only, audio only, or already combined).
  // Returns the list of byte arrays to write, in order.
  function single(initBytes, segmentList) {
    const init = readInit(initBytes);
    const ts = init.tracks[0].timescale;
    const frags = [];
    for (const seg of segmentList) frags.push(...fragments(seg, ts, 0));
    renumber(frags);
    const out = [init.ftyp ? slice(initBytes, init.ftyp) : new Uint8Array(0), slice(initBytes, init.moov)];
    for (const f of frags) out.push(...f.parts);
    return out;
  }

  // Video stream + audio stream -> one file.
  function combine(videoInitBytes, videoSegments, audioInitBytes, audioSegments) {
    const video = readInit(videoInitBytes);
    const audio = readInit(audioInitBytes);
    const audioId = Math.max(...video.tracks.map((t) => t.id)) + 1;
    const head = joinInits(video, audio, audioId);
    const vts = video.tracks[0].timescale;
    const ats = audio.tracks[0].timescale;
    const frags = [];
    for (const seg of videoSegments) frags.push(...fragments(seg, vts, 0));
    for (const seg of audioSegments) frags.push(...fragments(seg, ats, audioId));
    frags.sort((a, b) => a.seconds - b.seconds);
    renumber(frags);
    const out = [...head];
    for (const f of frags) out.push(...f.parts);
    return out;
  }

  const api = { boxes, readInit, single, combine };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.MSDLMp4 = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
