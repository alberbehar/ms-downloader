// MS Downloader — Copyright 2026 alberbehar
// Licensed under the Apache License, Version 2.0. See LICENSE and NOTICE.
// MS Downloader — turning DASH pieces into one ordinary MP4 file.
//
// DASH delivers each stream (video, audio) as fragmented MP4: an
// initialization segment (ftyp + moov) and many media segments (moof + mdat).
// Players seek poorly in a file made by simply joining those pieces, because
// there is no index at the front. So instead we read every sample out of the
// fragments and write a classic MP4:
//
//   ftyp · moov (full sample tables: stts, ctts, stss, stsc, stsz, stco/co64) · mdat
//
// The index sits at the front ("fast start"), so every player can seek.

(function (root) {
  'use strict';

  // ---------- byte helpers ----------

  function u32(b, o) { return ((b[o] << 24) >>> 0) + (b[o + 1] << 16) + (b[o + 2] << 8) + b[o + 3]; }
  function i32(b, o) { return (b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]; }
  function u64(b, o) { return u32(b, o) * 4294967296 + u32(b, o + 4); }
  function type4(b, o) { return String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]); }

  class Writer {
    constructor(size) { this.b = new Uint8Array(size); this.o = 0; }
    u8(v) { this.b[this.o++] = v & 255; }
    u32(v) {
      const b = this.b, o = this.o;
      b[o] = (v >>> 24) & 255; b[o + 1] = (v >>> 16) & 255; b[o + 2] = (v >>> 8) & 255; b[o + 3] = v & 255;
      this.o += 4;
    }
    u64(v) { this.u32(Math.floor(v / 4294967296)); this.u32(v % 4294967296); }
    bytes(a) { this.b.set(a, this.o); this.o += a.length; }
  }

  function box(type, parts) {
    const size = 8 + parts.reduce((s, p) => s + p.length, 0);
    const w = new Writer(size);
    w.u32(size);
    for (let i = 0; i < 4; i++) w.u8(type.charCodeAt(i));
    for (const p of parts) w.bytes(p);
    return w.b;
  }

  // Full box (version + flags) from a body writer function.
  function fullBox(type, version, flags, bodySize, write) {
    const w = new Writer(12 + bodySize);
    w.u32(12 + bodySize);
    for (let i = 0; i < 4; i++) w.u8(type.charCodeAt(i));
    w.u8(version); w.u8(flags >> 16); w.u8(flags >> 8); w.u8(flags);
    write(w);
    return w.b;
  }

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
        size = u64(bytes, p + 8);
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

  function child(bytes, b, type) {
    return boxes(bytes, b.body, b.end).find((x) => x.type === type) || null;
  }

  function path(bytes, b, ...types) {
    let cur = b;
    for (const t of types) {
      if (!cur) return null;
      cur = child(bytes, cur, t);
    }
    return cur;
  }

  function copy(bytes, b) { return bytes.slice(b.start, b.end); }

  // ---------- reading init segments ----------

  // One entry per track found in an init segment.
  function readInit(bytes) {
    const top = boxes(bytes);
    const ftyp = top.find((b) => b.type === 'ftyp') || null;
    const moov = top.find((b) => b.type === 'moov');
    if (!moov) throw new Error('Init segment has no moov box');
    const mvhd = child(bytes, moov, 'mvhd');
    const mvhdV = bytes[mvhd.body];
    const movieTimescale = u32(bytes, mvhd.body + (mvhdV === 1 ? 20 : 12));

    const trexById = {};
    const mvex = child(bytes, moov, 'mvex');
    if (mvex) {
      for (const t of boxes(bytes, mvex.body, mvex.end)) {
        if (t.type !== 'trex') continue;
        trexById[u32(bytes, t.body + 4)] = {
          descIndex: u32(bytes, t.body + 8),
          duration: u32(bytes, t.body + 12),
          size: u32(bytes, t.body + 16),
          flags: u32(bytes, t.body + 20)
        };
      }
    }

    const tracks = boxes(bytes, moov.body, moov.end).filter((b) => b.type === 'trak').map((trak) => {
      const tkhd = child(bytes, trak, 'tkhd');
      const id = u32(bytes, tkhd.body + (bytes[tkhd.body] === 1 ? 20 : 12));
      const mdhd = path(bytes, trak, 'mdia', 'mdhd');
      const timescale = u32(bytes, mdhd.body + (bytes[mdhd.body] === 1 ? 20 : 12));
      const hdlr = path(bytes, trak, 'mdia', 'hdlr');
      const handler = hdlr ? type4(bytes, hdlr.body + 8) : '';
      return {
        bytes, trak, id, timescale, handler,
        trex: trexById[id] || { descIndex: 1, duration: 0, size: 0, flags: 0 }
      };
    });
    if (!tracks.length) throw new Error('Init segment has no track');
    return { bytes, ftyp, moov, movieTimescale, tracks };
  }

  // ---------- reading media segments ----------

  // Pull every sample out of a media segment.
  // Calls onRun(trackId, run) for each trun:
  //   run = { time, samples: [{ size, duration, cts, sync }], data: Uint8Array }
  function readSegment(bytes, tracksById, onRun) {
    for (const moof of boxes(bytes).filter((b) => b.type === 'moof')) {
      let nextDataStart = null; // implicit base for trafs without base offset
      for (const traf of boxes(bytes, moof.body, moof.end).filter((b) => b.type === 'traf')) {
        const tfhd = child(bytes, traf, 'tfhd');
        if (!tfhd) continue;
        const flags = u32(bytes, tfhd.body) & 0xffffff;
        const id = u32(bytes, tfhd.body + 4);
        const track = tracksById[id];
        if (!track) continue;
        let p = tfhd.body + 8;
        let base = moof.start;
        if (flags & 0x01) { base = u64(bytes, p); p += 8; }
        else if (!(flags & 0x020000) && nextDataStart !== null) base = nextDataStart;
        if (flags & 0x02) p += 4;
        let defDuration = track.trex.duration;
        let defSize = track.trex.size;
        let defFlags = track.trex.flags;
        if (flags & 0x08) { defDuration = u32(bytes, p); p += 4; }
        if (flags & 0x10) { defSize = u32(bytes, p); p += 4; }
        if (flags & 0x20) { defFlags = u32(bytes, p); p += 4; }

        const tfdt = child(bytes, traf, 'tfdt');
        let time = tfdt ? (bytes[tfdt.body] === 1 ? u64(bytes, tfdt.body + 4) : u32(bytes, tfdt.body + 4)) : null;

        let cursor = base;
        for (const trun of boxes(bytes, traf.body, traf.end).filter((b) => b.type === 'trun')) {
          const v = bytes[trun.body];
          const tf = u32(bytes, trun.body) & 0xffffff;
          const count = u32(bytes, trun.body + 4);
          let q = trun.body + 8;
          if (tf & 0x01) { cursor = base + i32(bytes, q); q += 4; }
          let firstFlags = null;
          if (tf & 0x04) { firstFlags = u32(bytes, q); q += 4; }
          const samples = new Array(count);
          let total = 0;
          for (let i = 0; i < count; i++) {
            const duration = tf & 0x100 ? u32(bytes, (q += 4) - 4) : defDuration;
            const size = tf & 0x200 ? u32(bytes, (q += 4) - 4) : defSize;
            let sflags = tf & 0x400 ? u32(bytes, (q += 4) - 4) : defFlags;
            if (i === 0 && firstFlags !== null) sflags = firstFlags;
            let cts = 0;
            if (tf & 0x800) { cts = v === 0 ? u32(bytes, q) : i32(bytes, q); q += 4; }
            samples[i] = { size, duration, cts, sync: !(sflags & 0x00010000) };
            total += size;
          }
          if (cursor + total > bytes.length) throw new Error('Media segment is shorter than its index says');
          onRun(id, { time, samples, data: bytes.subarray(cursor, cursor + total) });
          if (time !== null) for (const s of samples) time += s.duration;
          cursor += total;
        }
        nextDataStart = cursor;
      }
    }
  }

  // ---------- writing sample tables ----------

  function stts(samples) {
    const runs = [];
    for (const s of samples) {
      const last = runs[runs.length - 1];
      if (last && last[1] === s.duration) last[0]++;
      else runs.push([1, s.duration]);
    }
    return fullBox('stts', 0, 0, 4 + runs.length * 8, (w) => { w.u32(runs.length); for (const [c, d] of runs) { w.u32(c); w.u32(d); } });
  }

  function ctts(samples) {
    if (!samples.some((s) => s.cts !== 0)) return null;
    const negative = samples.some((s) => s.cts < 0);
    const runs = [];
    for (const s of samples) {
      const last = runs[runs.length - 1];
      if (last && last[1] === s.cts) last[0]++;
      else runs.push([1, s.cts]);
    }
    return fullBox('ctts', negative ? 1 : 0, 0, 4 + runs.length * 8, (w) => {
      w.u32(runs.length);
      for (const [c, o] of runs) { w.u32(c); w.u32(o >>> 0); }
    });
  }

  function stss(samples) {
    if (samples.every((s) => s.sync)) return null;
    const idx = [];
    samples.forEach((s, i) => { if (s.sync) idx.push(i + 1); });
    return fullBox('stss', 0, 0, 4 + idx.length * 4, (w) => { w.u32(idx.length); for (const i of idx) w.u32(i); });
  }

  function stsc(chunks) {
    const runs = [];
    chunks.forEach((c, i) => {
      const last = runs[runs.length - 1];
      if (!last || last[1] !== c.count) runs.push([i + 1, c.count]);
    });
    return fullBox('stsc', 0, 0, 4 + runs.length * 12, (w) => {
      w.u32(runs.length);
      for (const [first, n] of runs) { w.u32(first); w.u32(n); w.u32(1); }
    });
  }

  function stsz(samples) {
    return fullBox('stsz', 0, 0, 8 + samples.length * 4, (w) => { w.u32(0); w.u32(samples.length); for (const s of samples) w.u32(s.size); });
  }

  function chunkOffsets(chunks, wide) {
    return wide
      ? fullBox('co64', 0, 0, 4 + chunks.length * 8, (w) => { w.u32(chunks.length); for (const c of chunks) w.u64(c.offset); })
      : fullBox('stco', 0, 0, 4 + chunks.length * 4, (w) => { w.u32(chunks.length); for (const c of chunks) w.u32(c.offset); });
  }

  // Write a 32- or 64-bit duration into a copied box at the given offsets.
  function withDuration(bytes, b, v0Offset, v1Offset, value) {
    const out = copy(bytes, b);
    const head = b.body - b.start;
    const v = out[head];
    const w = new Writer(0);
    w.b = out;
    if (v === 1) { w.o = head + v1Offset; w.u64(value); }
    else { w.o = head + v0Offset; w.u32(Math.min(value, 0xffffffff)); }
    return out;
  }

  function withTrackId(bytes, tkhd, id) {
    const out = copy(bytes, tkhd);
    const head = tkhd.body - tkhd.start;
    const w = new Writer(0);
    w.b = out;
    w.o = head + (out[head] === 1 ? 20 : 12);
    w.u32(id);
    return out;
  }

  // The media time the original edit list skipped (B-frame delay), if any.
  function oldMediaTime(track) {
    const b = track.bytes;
    const elst = path(b, track.trak, 'edts', 'elst');
    if (!elst) return 0;
    const v = b[elst.body];
    const count = u32(b, elst.body + 4);
    let p = elst.body + 8;
    for (let i = 0; i < count; i++) {
      const mediaTime = v === 1 ? u64(b, p + 8) : i32(b, p + 4);
      p += v === 1 ? 20 : 12;
      if (v === 1 ? u32(b, p - 12) !== 0xffffffff : mediaTime !== -1) return mediaTime;
    }
    return 0;
  }

  // Edit list: optional empty edit (track starts later than the others),
  // then the whole media from its first shown frame.
  function edts(delay, duration, mediaTime) {
    const entries = [];
    if (delay > 0) entries.push([delay, -1]);
    entries.push([duration, mediaTime]);
    const elst = fullBox('elst', 0, 0, 4 + entries.length * 12, (w) => {
      w.u32(entries.length);
      for (const [d, m] of entries) { w.u32(Math.min(d, 0xffffffff)); w.u32(m >>> 0); w.u32(0x00010000); }
    });
    return box('edts', [elst]);
  }

  // Rebuild a trak with new sample tables and durations.
  function buildTrak(track, newId, samples, chunks, wide, movieTimescale, delaySeconds) {
    const b = track.bytes;
    const mediaDuration = samples.reduce((n, s) => n + s.duration, 0);
    const skip = Math.max(0, Math.min(oldMediaTime(track), mediaDuration));
    const shownDuration = Math.round((mediaDuration - skip) * movieTimescale / track.timescale);
    const delay = Math.round((delaySeconds || 0) * movieTimescale);
    const movieDuration = delay + shownDuration;

    const stblOld = path(b, track.trak, 'mdia', 'minf', 'stbl');
    const stbl = box('stbl', [
      copy(b, child(b, stblOld, 'stsd')),
      stts(samples), ctts(samples), stss(samples), stsc(chunks), stsz(samples), chunkOffsets(chunks, wide)
    ].filter(Boolean));

    const minfOld = path(b, track.trak, 'mdia', 'minf');
    const minf = box('minf', boxes(b, minfOld.body, minfOld.end).map((x) => (x.type === 'stbl' ? stbl : copy(b, x))));

    const mdiaOld = child(b, track.trak, 'mdia');
    const mdia = box('mdia', boxes(b, mdiaOld.body, mdiaOld.end).map((x) => {
      if (x.type === 'minf') return minf;
      if (x.type === 'mdhd') return withDuration(b, x, 16, 24, mediaDuration);
      return copy(b, x);
    }));

    const editList = edts(delay, shownDuration, skip);
    let editPlaced = false;
    const parts = boxes(b, track.trak.body, track.trak.end).map((x) => {
      if (x.type === 'mdia') {
        if (!editPlaced) { editPlaced = true; return [editList, mdia]; }
        return mdia;
      }
      if (x.type === 'tkhd') {
        const t = withTrackId(b, x, newId);
        const tb = boxes(t)[0];
        return withDuration(t, tb, 20, 28, movieDuration);
      }
      if (x.type === 'edts') { editPlaced = true; return editList; }
      return copy(b, x);
    }).flat().filter(Boolean);
    return { trak: box('trak', parts), movieDuration };
  }

  // ---------- putting it together ----------

  // inputs: [{ init: Uint8Array, segments: Uint8Array[] }, ...]  (video first)
  // Returns the byte arrays of one ordinary MP4 file, in order.
  function build(inputs) {
    const outTracks = [];
    let nextId = 1;
    let firstInit = null;

    for (const input of inputs) {
      const init = readInit(input.init);
      if (!firstInit) firstInit = init;
      const byId = {};
      for (const t of init.tracks) {
        const entry = { track: t, newId: nextId++, samples: [], runs: [] };
        byId[t.id] = t;
        t.entry = entry;
        outTracks.push(entry);
      }
      for (const seg of input.segments) {
        readSegment(seg, byId, (id, run) => {
          const entry = byId[id].entry;
          // Skip a run we already have (players sometimes refetch).
          if (run.time !== null && entry.runs.length && run.time < entry.nextTime) return;
          entry.runs.push(run);
          for (const s of run.samples) entry.samples.push(s);
          if (run.time !== null) entry.nextTime = run.time + run.samples.reduce((n, s) => n + s.duration, 0);
        });
      }
    }
    const used = outTracks.filter((e) => e.samples.length);
    if (!used.length) throw new Error('No media found in the downloaded parts');

    // Order the data so audio and video sit close together in the file.
    const order = [];
    for (const e of used) {
      let t = 0;
      for (const run of e.runs) {
        const start = run.time !== null ? run.time : t;
        order.push({ e, run, at: start / e.track.timescale });
        t = start + run.samples.reduce((n, s) => n + s.duration, 0);
      }
    }
    order.sort((a, b) => a.at - b.at);

    const dataSize = order.reduce((n, o) => n + o.run.data.length, 0);
    const wide = dataSize > 0xffffffff - 50e6;
    const movieTimescale = firstInit.movieTimescale || 1000;
    // Where each track starts (first fragment time), relative to the earliest.
    for (const e of used) {
      const first = e.runs.find((r) => r.time !== null);
      e.start = first ? first.time / e.track.timescale : 0;
    }
    const earliest = Math.min(...used.map((e) => e.start));

    // Chunk table sizes do not depend on the offsets, so build once with
    // zero offsets to learn the moov size, then fill in the real offsets.
    for (const e of used) e.chunks = e.runs.map((r) => ({ count: r.samples.length, offset: 0 }));
    const ftyp = box('ftyp', [
      new TextEncoder().encode('isom'), new Uint8Array([0, 0, 2, 0]), new TextEncoder().encode('isomiso2avc1mp41')
    ]);
    const makeMoov = () => {
      let longest = 0;
      const traks = used.map((e) => {
        const r = buildTrak(e.track, e.newId, e.samples, e.chunks, wide, movieTimescale, e.start - earliest);
        longest = Math.max(longest, r.movieDuration);
        return r.trak;
      });
      const fb = firstInit.bytes;
      const mvhdBox = child(fb, firstInit.moov, 'mvhd');
      let mvhd = withDuration(fb, mvhdBox, 16, 24, longest);
      const mb = boxes(mvhd)[0];
      const w = new Writer(0);
      w.b = mvhd;
      w.o = mb.body + (mvhd[mb.body] === 1 ? 108 : 96);
      w.u32(nextId);
      return box('moov', [mvhd, ...traks]);
    };
    const moovSize = makeMoov().length;
    const mdatHeader = dataSize + 8 > 0xffffffff ? 16 : 8;
    let offset = ftyp.length + moovSize + mdatHeader;
    const chunkIndex = new Map(used.map((e) => [e, 0]));
    for (const o of order) {
      const i = chunkIndex.get(o.e);
      o.e.chunks[o.e.runs.indexOf(o.run, i)].offset = offset;
      chunkIndex.set(o.e, i + 1);
      offset += o.run.data.length;
    }
    const moov = makeMoov();

    const head = new Writer(mdatHeader);
    if (mdatHeader === 16) { head.u32(1); head.bytes(new TextEncoder().encode('mdat')); head.u64(dataSize + 16); }
    else { head.u32(dataSize + 8); head.bytes(new TextEncoder().encode('mdat')); }

    return [ftyp, moov, head.b, ...order.map((o) => o.run.data)];
  }

  // One stream on its own (video only, audio only, or already combined).
  function single(initBytes, segmentList) {
    return build([{ init: initBytes, segments: segmentList }]);
  }

  // Video stream + audio stream -> one file.
  function combine(videoInit, videoSegments, audioInit, audioSegments) {
    return build([{ init: videoInit, segments: videoSegments }, { init: audioInit, segments: audioSegments }]);
  }

  const api = { boxes, readInit, build, single, combine };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.MSDLMp4 = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
