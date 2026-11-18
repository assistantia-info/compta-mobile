/* Compta Mobile — chirurgie XLSX zéro perte (côté navigateur, sans dépendance)
 * Principe hérité de repare_f18_copie_api.py : un .xlsx est un zip ; on ne
 * modifie QUE la part XML de la feuille journal, par insertion de ligne dans
 * la plage déjà couverte par les SUM (E89/D90). Toutes les autres parts sont
 * recopiées octet pour octet (header local + données compressées d'origine).
 * Moteur : central directory -> slices ; inflateRaw via DecompressionStream
 * natif ; part modifiée ré-émise en STORE (method 0) ; central + EOCD
 * reconstruits. Aucune librairie externe.
 * Invariants (décisions 2026-09-04) : lignes 7–88 uniquement, colonnes A–E,
 * jamais F, aucune insertion structurelle, styles hérités s385/386/387
 * (ligne 7) puis s388/389/390 (lignes 8+), dates en série Excel, libellé en
 * inlineStr (sharedStrings jamais touché).
 */
"use strict";
window.CM = window.CM || {};
CM.xlsx = (function () {

  /* ---------- CRC32 ---------- */
  const CRC_TABLE = (function () {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  /* ---------- binaires ---------- */
  const te = new TextEncoder();
  const td = new TextDecoder("utf-8");
  const u16 = (u8, o) => new DataView(u8.buffer, u8.byteOffset, u8.byteLength).getUint16(o, true);
  const u32 = (u8, o) => new DataView(u8.buffer, u8.byteOffset, u8.byteLength).getUint32(o, true);

  async function inflateRaw(u8) {
    const ds = new DecompressionStream("deflate-raw");
    const resp = new Response(new Blob([u8]).stream().pipeThrough(ds));
    return new Uint8Array(await resp.arrayBuffer());
  }

  /* ---------- lecture du zip ---------- */
  function parseCentral(u8) {
    let eocd = -1;
    const floor = Math.max(0, u8.length - 66000);
    for (let i = u8.length - 22; i >= floor; i--) {
      if (u32(u8, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("EOCD introuvable (zip invalide ?)");
    const count = u16(u8, eocd + 10);
    let off = u32(u8, eocd + 16);
    const entries = [];
    for (let i = 0; i < count; i++) {
      if (u32(u8, off) !== 0x02014b50) throw new Error("central directory corrompu");
      entries.push({
        method: u16(u8, off + 10),
        crc: u32(u8, off + 16),
        compSize: u32(u8, off + 20),
        uncompSize: u32(u8, off + 24),
        localOff: u32(u8, off + 42),
        name: td.decode(u8.subarray(off + 46, off + 46 + u16(u8, off + 28)))
      });
      off += 46 + u16(u8, off + 28) + u16(u8, off + 30) + u16(u8, off + 32);
    }
    return entries;
  }

  function localEntrySpan(u8, e) {
    const o = e.localOff;
    if (u32(u8, o) !== 0x04034b50) throw new Error("local header corrompu : " + e.name);
    const start = o + 30 + u16(u8, o + 26) + u16(u8, o + 28);
    return { start: o, end: start + e.compSize };
  }

  async function readPart(u8, e) {
    const o = e.localOff;
    const start = o + 30 + u16(u8, o + 26) + u16(u8, o + 28);
    const raw = u8.subarray(start, start + e.compSize);
    if (e.method === 0) return new Uint8Array(raw);
    if (e.method === 8) return inflateRaw(raw);
    throw new Error("compression non supportée (" + e.method + ") : " + e.name);
  }

  /* ---------- réécriture du zip ---------- */
  function concat(parts) {
    let total = 0;
    for (const p of parts) total += p.length;
    const out = new Uint8Array(total);
    let o = 0;
    for (const p of parts) { out.set(p, o); o += p.length; }
    return out;
  }
  const w16 = (a, v) => { a.push(v & 0xFF, (v >>> 8) & 0xFF); };
  const w32 = (a, v) => { a.push(v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF); };

  function buildZip(original, entries, replacements) {
    const locals = [], central = [];
    let offset = 0;
    for (const e of entries) {
      const nameB = te.encode(e.name);
      let block, crc, compSize, method;
      if (replacements.has(e.name)) {
        const content = replacements.get(e.name);
        crc = crc32(content); compSize = content.length; method = 0;
        const h = [];
        w32(h, 0x04034b50); w16(h, 20); w16(h, 0); w16(h, method);
        w16(h, 0); w16(h, 0x21);            /* heure 00:00, date 1980-01-01 */
        w32(h, crc); w32(h, compSize); w32(h, compSize);
        w16(h, nameB.length); w16(h, 0);
        block = concat([new Uint8Array(h), nameB, content]);
      } else {
        const span = localEntrySpan(original, e);
        block = original.subarray(span.start, span.end); /* octets d'origine */
        crc = e.crc; compSize = e.compSize; method = e.method;
      }
      locals.push(block);
      const cd = [];
      w32(cd, 0x02014b50); w16(cd, 20); w16(cd, 20); w16(cd, 0); w16(cd, method);
      w16(cd, 0); w16(cd, 0x21); w32(cd, crc); w32(cd, compSize); w32(cd, e.uncompSize);
      w16(cd, nameB.length); w16(cd, 0); w16(cd, 0); w16(cd, 0); w16(cd, 0);
      w32(cd, 0); w32(cd, offset);
      central.push(concat([new Uint8Array(cd), nameB]));
      offset += block.length;
    }
    const cdBuf = concat(central);
    const eocd = [];
    w32(eocd, 0x06054b50); w16(eocd, 0); w16(eocd, 0);
    w16(eocd, entries.length); w16(eocd, entries.length);
    w32(eocd, cdBuf.length); w32(eocd, offset); w16(eocd, 0);
    return concat(locals.concat([cdBuf, new Uint8Array(eocd)]));
  }

  /* ---------- chirurgie du journal (string level) ---------- */
  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
                     .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function dateSerial(iso) {
    /* "2026-09-15" -> série Excel (jours depuis 1899-12-30) */
    const p = String(iso).slice(0, 10).split("-").map(Number);
    return Math.floor(Date.UTC(p[0], p[1] - 1, p[2]) / 86400000) + 25569;
  }
  function findRow(xml, n) {
    const needle = '<row r="' + n + '"';
    const open = xml.indexOf(needle);
    if (open < 0) return null;
    const close = xml.indexOf("</row>", open);
    const nextRow = xml.indexOf("<row ", open + needle.length);
    if (close >= 0 && (nextRow < 0 || close < nextRow)) return { start: open, end: close + 6 };
    const gt = xml.indexOf(">", open);
    if (gt > 0 && xml[gt - 1] === "/") return { start: open, end: gt + 1 };
    return null;
  }
  function isUsed(block) { return /<(v|f|is)[ >]/.test(block); }

  function findFreeRow(xml, first, last) {
    for (let n = first; n <= last; n++) {
      const b = findRow(xml, n);
      if (!b || !isUsed(xml.slice(b.start, b.end))) return n;
    }
    return null;
  }

  function stylesFor(n) {
    /* styles observés dans RELSEPT26 : ligne 7 spécifique, 8+ homogènes */
    return n === 7 ? { A: 385, B: 386, C: 387, D: 387, E: 387 }
                   : { A: 388, B: 389, C: 390, D: 390, E: 390 };
  }

  function buildRowXml(n, dateOp, dateVal, libelle, debit, credit) {
    const s = stylesFor(n);
    const c = [];
    c.push('<c r="A' + n + '" s="' + s.A + '"><v>' + dateSerial(dateOp) + '</v></c>');
    c.push('<c r="B' + n + '" s="' + s.B + '"><v>' + dateSerial(dateVal || dateOp) + '</v></c>');
    c.push('<c r="C' + n + '" s="' + s.C + '" t="inlineStr"><is><t xml:space="preserve">' + esc(libelle) + '</t></is></c>');
    c.push(debit === null || debit === undefined
      ? '<c r="D' + n + '" s="' + s.D + '"/>'
      : '<c r="D' + n + '" s="' + s.D + '"><v>' + debit + '</v></c>');
    c.push(credit === null || credit === undefined
      ? '<c r="E' + n + '" s="' + s.E + '"/>'
      : '<c r="E' + n + '" s="' + s.E + '"><v>' + credit + '</v></c>');
    return '<row r="' + n + '" spans="1:6" x14ac:dyDescent="0.25">' + c.join("") + '</row>';
  }

  function insertRow(xml, n, rowXml) {
    const existing = findRow(xml, n);
    if (existing) return xml.slice(0, existing.start) + rowXml + xml.slice(existing.end);
    for (let m = n + 1; m <= n + 200; m++) {   /* chercher la première row suivante existante */
      const b = findRow(xml, m);
      if (b) return xml.slice(0, b.start) + rowXml + xml.slice(b.start);
    }
    throw new Error("point d'insertion introuvable pour la ligne " + n);
  }

  /* ---------- correspondance nom de feuille -> part XML ---------- */
  async function sheetTargets(bytes) {
    const entries = parseCentral(bytes);
    const wbE = entries.find(function (e) { return e.name === "xl/workbook.xml"; });
    const relsE = entries.find(function (e) { return e.name === "xl/_rels/workbook.xml.rels"; });
    if (!wbE || !relsE) throw new Error("classeur sans workbook.xml");
    const wb = td.decode(await readPart(bytes, wbE));
    const rels = td.decode(await readPart(bytes, relsE));
    const rid2target = {};
    for (const tag of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /Id="([^"]+)"/.exec(tag[0]);
      const tg = /Target="([^"]+)"/.exec(tag[0]);
      if (id && tg) rid2target[id[1]] = tg[1];
    }
    const sheets = {};
    for (const tag of wb.matchAll(/<sheet\b[^>]*>/g)) {
      const nm = /name="([^"]+)"/.exec(tag[0]);
      const rid = /r:id="([^"]+)"/.exec(tag[0]);
      if (nm && rid) {
        let t = rid2target[rid[1]] || "";
        if (!t.startsWith("xl/")) t = "xl/" + t.replace(/^\//, "");
        sheets[nm[1]] = t;
      }
    }
    return { entries: entries, sheets: sheets };
  }

  /* ---------- API publique ---------- */
  async function journalState(bytes, sheetName, first, last) {
    const ctx = await sheetTargets(bytes);
    const part = ctx.sheets[sheetName];
    if (!part) throw new Error("feuille introuvable : " + sheetName + " (feuilles : " + Object.keys(ctx.sheets).join(", ") + ")");
    const pe = ctx.entries.find(function (e) { return e.name === part; });
    const xml = td.decode(await readPart(bytes, pe));
    let used = 0, free = null;
    for (let n = first; n <= last; n++) {
      const b = findRow(xml, n);
      if (b && isUsed(xml.slice(b.start, b.end))) used++;
      else if (free === null) free = n;
    }
    return { usedRows: used, firstFreeRow: free };
  }

  async function insertExpense(bytes, sheetName, exp, first, last) {
    const ctx = await sheetTargets(bytes);
    const part = ctx.sheets[sheetName];
    if (!part) throw new Error("feuille introuvable : " + sheetName);
    const pe = ctx.entries.find(function (e) { return e.name === part; });
    const xml = td.decode(await readPart(bytes, pe));
    const n = findFreeRow(xml, first, last);
    if (n === null) throw new Error("journal plein : aucune ligne libre entre " + first + " et " + last);
    const rowXml = buildRowXml(n, exp.dateOp, exp.dateVal || exp.dateOp, exp.libelle,
                               exp.debit === undefined ? null : exp.debit,
                               exp.credit === undefined ? null : exp.credit);
    const newXml = insertRow(xml, n, rowXml);
    const replacements = new Map([[part, te.encode(newXml)]]);
    return { bytes: buildZip(bytes, ctx.entries, replacements), row: n };
  }

  return { parseCentral, journalState, insertExpense, dateSerial };
})();