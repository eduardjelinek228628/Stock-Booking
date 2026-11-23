/* Probe Viewer – shared helpers. Classic script (no modules) so index.html opens by double-click. */
(function () {
  'use strict';
  const PV = (window.PV = window.PV || {});

  const ZW = /[​-‍⁠﻿]/g; // zero-width characters
  const NBSP = / /g;

  /** Normalise a header for matching: no zero-width chars, collapsed spaces, no trailing colon, lower case. */
  PV.normHeader = function (h) {
    return String(h == null ? '' : h).replace(ZW, '').replace(NBSP, ' ').replace(/\s+/g, ' ').trim().replace(/:$/, '').trim().toLowerCase();
  };

  /** Display form of a cell value (never used for writing). */
  PV.clean = function (v) {
    if (v == null) return '';
    return String(v).replace(ZW, '').replace(NBSP, ' ').trim();
  };

  /** Lower-case key for comparing categorical values ("Sent out" == "sent out", "missing" == "missing!"). */
  PV.key = function (v) {
    return PV.clean(v).toLowerCase().replace(/[!?.]+$/, '').replace(/\s+/g, ' ');
  };

  /** Parse numbers that may be stored as text ("36.6", "36,6", " -"). Returns null when not a number. */
  PV.toNum = function (v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (v == null) return null;
    const s = PV.clean(v).replace(/\s/g, '');
    if (!s) return null;
    if (/^[-+]?\d+([.,]\d+)?([eE][-+]?\d+)?$/.test(s)) return parseFloat(s.replace(',', '.'));
    return null;
  };

  PV.isNumberLikeText = function (v) {
    return typeof v === 'string' && PV.toNum(v) !== null;
  };

  // ---- A1 addresses ----
  PV.colLetter = function (n) { // 1 -> A
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };
  PV.colIndex = function (letters) { // A -> 1
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  };
  /** "'NPC-X'!A2:N553" -> {sheet:'NPC-X', c1:1, r1:2, c2:14, r2:553} */
  PV.parseAddress = function (addr) {
    let sheet = null, ref = String(addr);
    const bang = ref.lastIndexOf('!');
    if (bang >= 0) {
      sheet = ref.slice(0, bang);
      if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
      ref = ref.slice(bang + 1);
    }
    const m = ref.replace(/\$/g, '').match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
    if (!m) throw new Error('Cannot parse address ' + addr);
    const c1 = PV.colIndex(m[1]), r1 = +m[2];
    const c2 = m[3] ? PV.colIndex(m[3]) : c1, r2 = m[4] ? +m[4] : r1;
    return { sheet, c1, r1, c2, r2 };
  };
  PV.makeAddress = function (c1, r1, c2, r2) {
    const a = PV.colLetter(c1) + r1;
    return (c2 == null || (c2 === c1 && r2 === r1)) ? a : a + ':' + PV.colLetter(c2) + r2;
  };

  // ---- dates (Excel serial, 1900 system) ----
  const EPOCH = Date.UTC(1899, 11, 30);
  PV.serialToDate = function (serial) { return new Date(EPOCH + Math.round(serial * 86400000)); };
  PV.dateToSerial = function (d) {
    const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
    return Math.round((utc - EPOCH) / 86400000);
  };
  PV.isoDate = function (d) {
    const y = d.getUTCFullYear(), m = d.getUTCMonth() + 1, day = d.getUTCDate();
    return y + '-' + String(m).padStart(2, '0') + '-' + String(day).padStart(2, '0');
  };
  PV.todayISO = function () {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  };
  PV.isDateFormat = function (nf) {
    if (!nf || typeof nf !== 'string') return false;
    const f = nf.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').toLowerCase();
    if (f === 'general' || f === '@') return false;
    return /(^|[^a-z])(d{1,4}|m{1,5}|y{2,4})([^a-z]|$)/.test(f) && /[dy]/.test(f);
  };
  /** A cell as a Date, or null. Accepts Excel serials (with a date format), Date objects and ISO/Czech date text. */
  PV.cellDate = function (v, nf, text) {
    if (v instanceof Date) return v;
    if (typeof v === 'number' && PV.isDateFormat(nf) && v > 20000 && v < 80000) return PV.serialToDate(v);
    const s = PV.clean(text != null && text !== '' ? text : v);
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    m = s.match(/^(\d{1,2})\.\s?(\d{1,2})\.\s?(\d{4})$/);
    if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // Graph text for US-formatted dates (m/d/yyyy)
    if (m) return new Date(Date.UTC(+m[3], +m[1] - 1, +m[2]));
    return null;
  };

  /** Text written by the viewer must never be interpreted as a formula. */
  PV.safeText = function (s) {
    s = String(s == null ? '' : s);
    return /^[=+\-@]/.test(s) ? "'" + s : s;
  };

  PV.esc = function (s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  };

  /** Small deterministic RNG for demo data. */
  PV.rng = function (seed) {
    let s = seed >>> 0;
    const next = () => { s = (s + 0x6D2B79F5) >>> 0; let t = s; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
    return {
      next,
      int: (a, b) => a + Math.floor(next() * (b - a + 1)),
      pick: (arr) => arr[Math.floor(next() * arr.length)],
      weighted: (pairs) => { const tot = pairs.reduce((a, p) => a + p[1], 0); let r = next() * tot; for (const p of pairs) { r -= p[1]; if (r <= 0) return p[0]; } return pairs[pairs.length - 1][0]; },
      round: (a, b, dp) => +(a + next() * (b - a)).toFixed(dp),
    };
  };

  /** Run async jobs with limited parallelism. */
  PV.pool = async function (items, limit, fn) {
    const out = new Array(items.length);
    let i = 0;
    const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
      while (i < items.length) { const k = i++; out[k] = await fn(items[k], k); }
    });
    await Promise.all(workers);
    return out;
  };

  PV.storage = {
    get(k) { try { const v = localStorage.getItem(k); return v == null ? null : JSON.parse(v); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } },
    del(k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } },
  };

  /** Owner cells like "PB -> MPat" mean the probe was handed over; the last name is the current owner. */
  PV.currentOwner = function (v) {
    const parts = PV.clean(v).split(/->|→/);
    return parts[parts.length - 1].trim();
  };

  PV.monthKey = function (d) { return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0'); };
  PV.addMonths = function (d, n) { return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + n, 1)); };
})();
