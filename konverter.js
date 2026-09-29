/*
 * Begrepsliste → Anki
 * Leser en utfylt begrepsliste (.docx eller innlimt tekst fra Word/Google Docs)
 * og lager en Anki-pakke (.apkg). Alt skjer lokalt i nettleseren.
 *
 * Fungerer både i nettleser (window.Konverter) og i Node (module.exports),
 * slik at logikken kan testes uten nettleser.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.Konverter = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // Faste ID-er: samme note type og samme kortstokker hver gang, slik at en ny
  // import oppdaterer kortene i stedet for å lage duplikater.
  const MODEL_ID = 1726400000101;
  const MODEL_NAME = "Begrepsliste (begrep → definisjon)";
  const GUID_SALT = "begrepsliste-v1";

  // Tekst som står i malen fra før, og som ikke teller som en definisjon.
  const PLACEHOLDER_LINES = new Set([
    "(beskriv forskjellen mellom disse to)",
    "nødvendig", "tilstrekkelig", "eksempler:", "eksempler",
    "1.", "2.", "3.", "4.",
  ]);

  // ---------- små hjelpere ----------
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const squash = (s) => String(s).replace(/[\s\u00a0\u2060]+/g, " ").trim();
  const stripHtml = (h) => squash(String(h).replace(/<br\s*\/?>/gi, " ").replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&"));

  function isHeading(text) {
    const t = squash(text);
    if (t.length < 4 || /^pensum\s*:/i.test(t)) return false;
    const letters = t.replace(/[^A-Za-zÆØÅæøåÄÖÜäöüÉé]/g, "");
    if (letters.length < 4) return false;
    const upper = letters.replace(/[^A-ZÆØÅÄÖÜÉ]/g, "").length;
    return upper / letters.length > 0.8;
  }

  function prettyTema(raw) {
    const t = squash(raw).replace(/[.:]$/, "");
    if (!t) return "Uten tema";
    const lower = t.toLocaleLowerCase("nb");
    return lower.charAt(0).toLocaleUpperCase("nb") + lower.slice(1);
  }

  function isEffectivelyEmpty(defHtml) {
    const lines = String(defHtml).split(/<br\s*\/?>/i)
      .map((l) => stripHtml(l).toLocaleLowerCase("nb"))
      .filter((l) => l !== "");
    return lines.every((l) => PLACEHOLDER_LINES.has(l));
  }

  // Finn kolonnene ut fra overskriftsraden; fall tilbake til 0/1/2.
  function columnMap(headerTexts) {
    const m = { begrep: 0, kilde: 1, def: 2 };
    if (!headerTexts) return m;
    const low = headerTexts.map((t) => squash(t).toLocaleLowerCase("nb"));
    const find = (re) => low.findIndex((t) => re.test(t));
    const b = find(/begrep/), k = find(/sidetall|side|kilde/), d = find(/definisjon|forklaring/);
    if (b >= 0) m.begrep = b;
    if (d >= 0) m.def = d;
    m.kilde = k >= 0 ? k : -1;
    if (headerTexts.length === 2 && d < 0) { m.def = 1; m.kilde = -1; }
    return m;
  }

  // ---------- .docx (word/document.xml) ----------
  function kids(node, name) {
    const out = [];
    for (let c = node.firstChild; c; c = c.nextSibling) if (c.nodeName === name) out.push(c);
    return out;
  }

  function runsHtml(parent) {
    let html = "";
    for (let c = parent.firstChild; c; c = c.nextSibling) {
      const n = c.nodeName;
      if (n === "w:r") {
        const rPr = kids(c, "w:rPr")[0];
        const on = (tag) => {
          if (!rPr) return false;
          const el = kids(rPr, tag)[0];
          if (!el) return false;
          const v = el.getAttribute("w:val");
          return !(v === "0" || v === "false" || v === "none");
        };
        let txt = "";
        for (let t = c.firstChild; t; t = t.nextSibling) {
          if (t.nodeName === "w:t") txt += esc(t.textContent);
          else if (t.nodeName === "w:tab") txt += " ";
          else if (t.nodeName === "w:br" || t.nodeName === "w:cr") txt += "<br>";
          else if (t.nodeName === "w:noBreakHyphen") txt += "-";
        }
        if (!txt) continue;
        if (on("w:b")) txt = "<b>" + txt + "</b>";
        if (on("w:i")) txt = "<i>" + txt + "</i>";
        if (on("w:u")) txt = "<u>" + txt + "</u>";
        html += txt;
      } else if (n === "w:hyperlink" || n === "w:ins" || n === "w:smartTag" || n === "w:fldSimple") {
        html += runsHtml(c);
      } else if (n === "w:sdt") {
        const content = kids(c, "w:sdtContent")[0];
        if (content) html += runsHtml(content);
      }
    }
    return html.replace(/<\/b><b>/g, "").replace(/<\/i><i>/g, "");
  }

  function cellHtml(tc) {
    const paras = [];
    (function walk(node) {
      for (let c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeName === "w:p") paras.push(runsHtml(c));
        else if (c.nodeName === "w:sdt" || c.nodeName === "w:sdtContent" || c.nodeName === "w:customXml") walk(c);
      }
    })(tc);
    while (paras.length && stripHtml(paras[paras.length - 1]) === "") paras.pop();
    while (paras.length && stripHtml(paras[0]) === "") paras.shift();
    // slå sammen flere tomme avsnitt til ett
    const out = [];
    for (const p of paras) {
      if (stripHtml(p) === "" && out.length && out[out.length - 1] === "") continue;
      out.push(stripHtml(p) === "" ? "" : p.trim());
    }
    return out.join("<br>");
  }

  function parseDocumentXml(xml, DOMParserImpl) {
    const DP = DOMParserImpl || DOMParser;
    const doc = new DP().parseFromString(xml, "text/xml");
    const body = doc.getElementsByTagName("w:body")[0];
    if (!body) throw new Error("Fant ikke innholdet i Word-filen.");
    const blocks = [];
    (function walk(node) {
      for (let c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeName === "w:p") blocks.push({ type: "p", text: stripHtml(runsHtml(c)) });
        else if (c.nodeName === "w:tbl") {
          const rows = kids(c, "w:tr").map((tr) => kids(tr, "w:tc").map(cellHtml));
          blocks.push({ type: "table", rows });
        } else if (c.nodeName === "w:sdt") {
          const content = kids(c, "w:sdtContent")[0];
          if (content) walk(content);
        }
      }
    })(body);
    return blocksToEntries(blocks);
  }

  // ---------- innlimt HTML (Google Docs, Word) ----------
  function inlineHtml(node) {
    let html = "";
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3) { html += esc(c.nodeValue.replace(/\s+/g, " ")); continue; }
      if (c.nodeType !== 1) continue;
      const tag = c.nodeName.toLowerCase();
      if (tag === "br") { html += "<br>"; continue; }
      if (tag === "style" || tag === "script") continue;
      let inner = inlineHtml(c);
      const block = /^(p|div|li|h[1-6])$/.test(tag);
      const st = (c.getAttribute && c.getAttribute("style")) || "";
      const fw = /font-weight\s*:\s*(bold|[6-9]00)/i.test(st);
      const it = /font-style\s*:\s*italic/i.test(st);
      if (inner.trim() && (tag === "b" || tag === "strong" || fw)) inner = "<b>" + inner + "</b>";
      if (inner.trim() && (tag === "i" || tag === "em" || it)) inner = "<i>" + inner + "</i>";
      if (tag === "li") inner = "• " + inner;
      html += block ? "\n" + inner + "\n" : inner;
    }
    return html;
  }

  function htmlCell(td) {
    const raw = inlineHtml(td).replace(/[ \t\u00a0]+/g, " ");
    const lines = raw.split("\n").map((l) => l.trim());
    const out = [];
    for (const l of lines) {
      const empty = stripHtml(l) === "";
      if (empty && (!out.length || out[out.length - 1] === "")) continue;
      out.push(empty ? "" : l);
    }
    while (out.length && out[out.length - 1] === "") out.pop();
    return out.join("<br>").replace(/<\/b><b>/g, "").replace(/(<br>){2,}/g, "<br>");
  }

  function parseHtml(htmlString, DOMParserImpl) {
    const DP = DOMParserImpl || DOMParser;
    const doc = new DP().parseFromString(htmlString, "text/html");
    const blocks = [];
    const els = doc.body.querySelectorAll("p, h1, h2, h3, h4, h5, h6, table, li");
    els.forEach((el) => {
      if (el.nodeName.toLowerCase() === "table") {
        if (el.parentElement && el.parentElement.closest("table")) return;
        const rows = [];
        el.querySelectorAll("tr").forEach((tr) => {
          if (tr.closest("table") !== el) return;
          rows.push(Array.from(tr.children).filter((c) => /^(td|th)$/i.test(c.nodeName)).map(htmlCell));
        });
        blocks.push({ type: "table", rows });
      } else if (!el.closest("table")) {
        blocks.push({ type: "p", text: squash(el.textContent) });
      }
    });
    return blocksToEntries(blocks);
  }

  // ---------- felles: blokker → oppføringer ----------
  function blocksToEntries(blocks) {
    const entries = [], skipped = [];
    let tema = "", temaOrder = 0, lastTemaForTable = null;
    const temaIndex = new Map();
    for (const b of blocks) {
      if (b.type === "p") { if (isHeading(b.text)) tema = prettyTema(b.text); continue; }
      if (!b.rows.length) continue;
      let rows = b.rows, header = null;
      const first = rows[0].map(stripHtml);
      if (first.some((t) => /^begrep/i.test(t))) { header = first; rows = rows.slice(1); }
      else if (lastTemaForTable === tema && b._cols) header = null;
      const cols = columnMap(header);
      const t = tema || "Uten tema";
      if (!temaIndex.has(t)) temaIndex.set(t, ++temaOrder);
      lastTemaForTable = tema;
      for (const cells of rows) {
        if (cells.length < 2) continue;
        const begrepHtml = cells[cols.begrep] || "";
        const begrep = stripHtml(begrepHtml);
        if (!begrep) continue;
        const definisjon = cells[cols.def] || "";
        const kilde = cols.kilde >= 0 ? stripHtml(cells[cols.kilde] || "") : "";
        const item = { tema: t, temaNr: temaIndex.get(t), begrep, kilde, definisjon };
        if (!definisjon || isEffectivelyEmpty(definisjon)) skipped.push(item);
        else entries.push(item);
      }
    }
    return { entries, skipped };
  }

  async function parseDocx(arrayBuffer, JSZipImpl, DOMParserImpl) {
    const JZ = JSZipImpl || JSZip;
    let zip;
    try { zip = await JZ.loadAsync(arrayBuffer); }
    catch (e) { throw new Error("Filen kunne ikke åpnes som en Word-fil (.docx). Er det kanskje en .doc, .odt eller .pdf?"); }
    const f = zip.file("word/document.xml");
    if (!f) throw new Error("Dette ser ikke ut som en Word-fil (.docx).");
    return parseDocumentXml(await f.async("string"), DOMParserImpl);
  }

  // ---------- Anki-pakke ----------
  async function sha1Hex(str) {
    const data = new TextEncoder().encode(str);
    const buf = await crypto.subtle.digest("SHA-1", data);
    return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
  }

  function hashId(str, base) {
    // enkel, stabil 32-bits hash → positivt heltall (FNV-1a)
    let h = 0x811c9dc5;
    for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return base + h;
  }

  function normBegrep(b) {
    return squash(b).toLocaleLowerCase("nb").replace(/[«»"“”'’.,;:]/g, "");
  }

  // Skjul selve begrepet i definisjonen (for baklengs-kort).
  function maskTerm(defHtml, begrep) {
    const head = begrep.split(/[(/]/)[0];
    const words = Array.from(new Set(head.toLocaleLowerCase("nb").match(/[\p{L}-]{5,}/gu) || []))
      .filter((w) => !["forskjell", "mellom", "ifølge", "hvem", "liste", "forklaring", "beskrivelse", "definisjon"].includes(w));
    if (!words.length) return defHtml;
    const stems = words.map((w) => (w.length > 7 ? w.slice(0, -2) : w).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
    const re = new RegExp("[\\p{L}-]*(" + stems.join("|") + ")[\\p{L}-]*", "giu");
    return defHtml.split(/(<[^>]+>)/).map((part) => (part.startsWith("<") ? part : part.replace(re, "[…]"))).join("");
  }

  const CSS = `.card { font-family: "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; font-size: 19px; line-height: 1.45; text-align: left; color: #1d2733; background: #fbfaf7; max-width: 38em; margin: 0 auto; padding: 0.5em 0.8em; }
.card.nightMode, .nightMode .card { color: #e6e9ee; background: #1c2129; }
.tema { font-size: 0.7em; letter-spacing: 0.06em; text-transform: uppercase; color: #6b7785; margin-bottom: 0.6em; }
.begrep { font-size: 1.35em; font-weight: 650; line-height: 1.25; }
.hint { font-size: 0.8em; color: #6b7785; margin-bottom: 0.4em; }
.def b { font-weight: 650; }
.kilde { margin-top: 1em; font-size: 0.75em; color: #6b7785; }
hr#answer { border: none; border-top: 1px solid #c9d0d8; margin: 1em 0; }`;

  function modelJson(now, did) {
    return {
      id: MODEL_ID, name: MODEL_NAME, type: 0, mod: now, usn: -1, sortf: 0, did,
      tags: [], vers: [],
      flds: ["Begrep", "Definisjon", "Kilde", "Tema", "Baklengs"].map((name, ord) => ({
        name, ord, sticky: false, rtl: false, font: "Arial", size: 20, media: [],
      })),
      tmpls: [
        {
          name: "Begrep → definisjon", ord: 0, did: null, bqfmt: "", bafmt: "",
          qfmt: '<div class="tema">{{Tema}}</div>\n<div class="begrep">{{Begrep}}</div>',
          afmt: '{{FrontSide}}\n<hr id="answer">\n<div class="def">{{Definisjon}}</div>\n{{#Kilde}}<div class="kilde">{{Kilde}}</div>{{/Kilde}}',
        },
        {
          name: "Definisjon → begrep", ord: 1, did: null, bqfmt: "", bafmt: "",
          qfmt: '{{#Baklengs}}<div class="tema">{{Tema}}</div>\n<div class="hint">Hvilket begrep er dette?</div>\n<div class="def">{{Baklengs}}</div>{{/Baklengs}}',
          afmt: '{{FrontSide}}\n<hr id="answer">\n<div class="begrep">{{Begrep}}</div>\n{{#Kilde}}<div class="kilde">{{Kilde}}</div>{{/Kilde}}',
        },
      ],
      req: [[0, "any", [0]], [1, "any", [4]]],
      css: CSS,
      latexPre: "\\documentclass[12pt]{article}\n\\special{papersize=3in,5in}\n\\usepackage[utf8]{inputenc}\n\\usepackage{amssymb,amsmath}\n\\pagestyle{empty}\n\\setlength{\\parindent}{0in}\n\\begin{document}\n",
      latexPost: "\\end{document}",
    };
  }

  function deckJson(id, name, now) {
    return {
      id, name, mod: now, usn: -1, desc: "", dyn: 0, conf: 1, collapsed: false, browserCollapsed: false,
      extendNew: 10, extendRev: 50, newToday: [0, 0], revToday: [0, 0], lrnToday: [0, 0], timeToday: [0, 0],
    };
  }

  const DCONF = {
    1: {
      id: 1, name: "Default", mod: 0, usn: 0, dyn: false, maxTaken: 60, timer: 0, autoplay: true, replayq: true,
      new: { bury: true, delays: [1, 10], initialFactor: 2500, ints: [1, 4, 7], order: 1, perDay: 20, separate: true },
      lapse: { delays: [10], leechAction: 0, leechFails: 8, minInt: 1, mult: 0 },
      rev: { bury: true, ease4: 1.3, fuzz: 0.05, ivlFct: 1, maxIvl: 36500, minSpace: 1, perDay: 200, hardFactor: 1.2 },
    },
  };

  const SCHEMA = `
CREATE TABLE col (id integer primary key, crt integer not null, mod integer not null, scm integer not null, ver integer not null, dty integer not null, usn integer not null, ls integer not null, conf text not null, models text not null, decks text not null, dconf text not null, tags text not null);
CREATE TABLE notes (id integer primary key, guid text not null, mid integer not null, mod integer not null, usn integer not null, tags text not null, flds text not null, sfld integer not null, csum integer not null, flags integer not null, data text not null);
CREATE TABLE cards (id integer primary key, nid integer not null, did integer not null, ord integer not null, mod integer not null, usn integer not null, type integer not null, queue integer not null, due integer not null, ivl integer not null, factor integer not null, reps integer not null, lapses integer not null, left integer not null, odue integer not null, odid integer not null, flags integer not null, data text not null);
CREATE TABLE revlog (id integer primary key, cid integer not null, usn integer not null, ivl integer not null, lastIvl integer not null, factor integer not null, time integer not null, type integer not null);
CREATE TABLE graves (usn integer not null, oid integer not null, type integer not null);
CREATE INDEX ix_notes_usn on notes (usn);
CREATE INDEX ix_cards_usn on cards (usn);
CREATE INDEX ix_revlog_usn on revlog (usn);
CREATE INDEX ix_cards_nid on cards (nid);
CREATE INDEX ix_cards_sched on cards (did, queue, due);
CREATE INDEX ix_revlog_cid on revlog (cid);
CREATE INDEX ix_notes_csum on notes (csum);`;

  function deckNameFor(prefix, e) {
    const clean = (s) => s.replace(/::/g, ":").trim();
    return clean(prefix) + "::" + String(e.temaNr).padStart(2, "0") + " " + clean(e.tema);
  }

  function tagFor(e) {
    const uke = e.tema.match(/uke\s*(\d+)/i);
    const slug = e.tema.toLocaleLowerCase("nb").replace(/\(.*?\)/g, "").trim()
      .replace(/[^\p{L}\p{N}]+/gu, "_").replace(/^_|_$/g, "").slice(0, 40);
    return uke ? "uke" + uke[1] : slug || "tema";
  }

  /**
   * Lager en .apkg (Uint8Array).
   * opts: { deckPrefix, reverse, courseTag }
   * libs: { SQL (initialisert sql.js), JSZip }
   */
  async function buildApkg(entries, opts, libs) {
    const { SQL, JSZip: JZ } = libs;
    const deckPrefix = (opts.deckPrefix || "Begrepsliste").trim();
    const reverse = !!opts.reverse;
    const courseTag = (opts.courseTag || deckPrefix).replace(/\s+/g, "_");
    const nowMs = Date.now();
    const now = Math.floor(nowMs / 1000);

    const db = new SQL.Database();
    db.run(SCHEMA);

    const decks = { 1: deckJson(1, "Default", now) };
    const parentId = hashId("deck:" + deckPrefix, 1500000000000);
    decks[parentId] = deckJson(parentId, deckPrefix, now);
    const deckIds = new Map();
    for (const e of entries) {
      const name = deckNameFor(deckPrefix, e);
      if (!deckIds.has(name)) {
        const id = hashId("deck:" + name, 1500000000000);
        deckIds.set(name, id);
        decks[id] = deckJson(id, name, now);
      }
    }

    const models = { [MODEL_ID]: modelJson(now, parentId) };
    const conf = {
      activeDecks: [1], curDeck: 1, newSpread: 0, collapseTime: 1200, timeLim: 0, estTimes: true,
      dueCounts: true, curModel: String(MODEL_ID), nextPos: entries.length + 1, sortType: "noteFld", sortBackwards: false, addToCur: true,
    };
    db.run("INSERT INTO col VALUES (1, ?, ?, ?, 11, 0, 0, 0, ?, ?, ?, ?, '{}')",
      [now, nowMs, nowMs, JSON.stringify(conf), JSON.stringify(models), JSON.stringify(decks), JSON.stringify(DCONF)]);

    const seen = new Set();
    let nid = nowMs, cid = nowMs, pos = 0, noteCount = 0, cardCount = 0;
    for (const e of entries) {
      const key = normBegrep(e.begrep);
      const guid = (await sha1Hex(GUID_SALT + "|" + deckPrefix + "|" + key)).slice(0, 16);
      if (seen.has(guid)) continue; // samme begrep to ganger: behold det første
      seen.add(guid);
      const baklengs = reverse ? maskTerm(e.definisjon, e.begrep) : "";
      const fields = [esc(e.begrep), e.definisjon, esc(e.kilde), esc(e.tema), baklengs];
      const sfld = stripHtml(fields[0]);
      const csum = parseInt((await sha1Hex(sfld)).slice(0, 8), 16);
      const tags = " " + [courseTag, tagFor(e)].join(" ") + " ";
      nid += 1;
      db.run("INSERT INTO notes VALUES (?, ?, ?, ?, -1, ?, ?, ?, ?, 0, '')",
        [nid, guid, MODEL_ID, now, tags, fields.join("\x1f"), sfld, csum]);
      noteCount++;
      const did = deckIds.get(deckNameFor(deckPrefix, e));
      pos += 1;
      const ords = reverse ? [0, 1] : [0];
      for (const ord of ords) {
        cid += 1;
        db.run("INSERT INTO cards VALUES (?, ?, ?, ?, ?, -1, 0, 0, ?, 0, 0, 0, 0, 0, 0, 0, 0, '')",
          [cid, nid, did, ord, now, pos]);
        cardCount++;
      }
    }

    const dbBytes = db.export();
    db.close();
    const zip = new JZ();
    zip.file("collection.anki2", dbBytes);
    zip.file("media", "{}");
    const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return { bytes, noteCount, cardCount, deckCount: deckIds.size };
  }

  return {
    parseDocx, parseDocumentXml, parseHtml, buildApkg, maskTerm, isEffectivelyEmpty, isHeading, prettyTema,
  };
});
