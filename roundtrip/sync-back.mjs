#!/usr/bin/env node
// sync-back.mjs — pull edits made in PowerPoint (saved to out/deck.pptx) back into the code.
//
//   node tools/sync-back.mjs [--dry-run] [--pptx out/deck.pptx] [--manifest out/manifest.json] [--force]
//
// --pptx      a saved copy other than out/deck.pptx (e.g. a "Save As" from PowerPoint)
// --manifest  the manifest of the build that file was edited from (default: out/manifest.json)
// --force     apply even when the file looks like it comes from a different build (see the guard below)
//
// Every shape the build creates carries a stable name "<slide id>|<kind>|<n>" and is recorded in
// out/manifest.json with its text, per-paragraph style and geometry. This script opens the saved .pptx,
// finds those shapes, and for each difference:
//   • text changed, same paragraphs   → patch the matching string literal(s) in deck.mjs
//   • paragraphs added / removed      → rewrite the array literal in deck.mjs (bullets, card bodies, code
//                                       lines); if the array cannot be located, store a text override
//   • font size / color / bold / italic / alignment changed → overrides.json "fmt" (per paragraph)
//   • shape fill / border changed     → overrides.json "fill" / "line"
//   • moved / resized                 → overrides.json {x,y,w,h}
//   • deleted                         → overrides.json {hidden:true}
//   • table cells / rows / column widths changed → cell text patched when the grid is unchanged,
//                                       otherwise overrides.json "cells" / "colW"
//   • notes changed                   → patch the `notes:` literal
// Shapes added in PowerPoint and slide reordering are reported but cannot flow back.

import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
// Theme colors resolve PowerPoint's scheme colors (tx1, accent1, …) to hex. They come from the manifest
// when the build records them (manifest.theme), else from ../theme.mjs, else scheme colors are ignored.
let THEME = { colors: {} };

const require = createRequire(import.meta.url);
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, "..");
const args = process.argv.slice(2);
const DRY = args.includes("--dry-run");
const FORCE = args.includes("--force");
const PPTX = path.resolve(ROOT, args.includes("--pptx") ? args[args.indexOf("--pptx") + 1] : "out/deck.pptx");
const MANIFEST = path.resolve(ROOT, args.includes("--manifest") ? args[args.indexOf("--manifest") + 1] : "out/manifest.json");
const DECK = path.join(ROOT, "deck.mjs");
const OVERRIDES = path.join(ROOT, "overrides.json");
const EMU = 914400;
const TOL = 0.012; // inches

if (!fs.existsSync(MANIFEST)) { console.error("sync-back: out/manifest.json missing — run `node deck.mjs` once first"); process.exit(2); }
const manifest = JSON.parse(fs.readFileSync(MANIFEST, "utf8"));
if (manifest.theme?.colors) THEME = manifest.theme;
else { try { THEME = (await import(path.join(ROOT, "theme.mjs"))).THEME; } catch { /* no theme: scheme colors stay unresolved */ } }
const overrides = fs.existsSync(OVERRIDES) ? JSON.parse(fs.readFileSync(OVERRIDES, "utf8")) : {};
let deckSrc = fs.readFileSync(DECK, "utf8");
const report = { patched: [], arrays: [], overridden: [], fmt: [], shape: [], moved: [], hidden: [], tables: [], notes: [], added: [], warnings: [] };

// ---- XML helpers (regex-based; OOXML from pptxgenjs and PowerPoint both keep these tags) ----------
const unesc = (t) => t.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, "&");
const attr = (xml, name) => { const m = xml.match(new RegExp(`\\b${name}="([^"]*)"`)); return m ? unesc(m[1]) : null; };
const SCHEME = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2", dk1: "dk1", lt1: "lt1", dk2: "dk2", lt2: "lt2", accent1: "accent1", accent2: "accent2", accent3: "accent3", accent4: "accent4", accent5: "accent5", accent6: "accent6", hlink: "hlink", folHlink: "folHlink" };
// First solid color inside an element (srgbClr or schemeClr mapped through the theme). Modifiers (lumMod…) are ignored.
function colorOf(xml) {
  const fill = xml.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
  if (!fill) return null;
  const srgb = fill[1].match(/<a:srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/);
  if (srgb) return srgb[1].toUpperCase();
  const sch = fill[1].match(/<a:schemeClr\b[^>]*val="(\w+)"/);
  if (sch && SCHEME[sch[1]] && THEME.colors[SCHEME[sch[1]]]) return THEME.colors[SCHEME[sch[1]]].toUpperCase();
  return null;
}

function paragraphsOf(txBody) {
  const paras = [];
  for (const pm of txBody.matchAll(/<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g)) {
    const p = pm[1];
    const pPr = p.match(/<a:pPr\b[^>]*>/)?.[0] ?? p.match(/<a:pPr\b[^>]*\/>/)?.[0] ?? "";
    const level = +(pPr.match(/\blvl="(\d+)"/)?.[1] ?? 0);
    const algn = pPr.match(/\balgn="(\w+)"/)?.[1];
    const align = algn === "ctr" ? "center" : algn === "r" ? "right" : algn === "just" ? "justify" : algn === "l" ? "left" : null;
    const runs = [...p.matchAll(/<a:r\b[^>]*>([\s\S]*?)<\/a:r>/g)];
    if (!runs.length) { if (/<a:fld\b/.test(p)) continue; paras.push({ text: "", level, align, nostyle: true }); continue; }
    let text = "";
    const rprs = [];
    for (const r of runs) {
      const t = r[1].match(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/);
      if (t) text += unesc(t[1]);
      const rPr = r[1].match(/<a:rPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/)?.[0] ?? "";
      rprs.push(rPr);
    }
    if (/<a:br\b/.test(p)) {
      text = "";
      for (const part of p.split(/<a:br\b[^>]*\/>/)) {
        for (const t of part.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)) text += unesc(t[1]);
        text += "\n";
      }
      text = text.replace(/\n$/, "");
    }
    // style from the first run that has text; "uniform" only if every run agrees
    const first = rprs[0] || "";
    const sz = first.match(/\bsz="(\d+)"/)?.[1];
    const style = {
      bold: /\bb="1"/.test(first), italic: /\bi="1"/.test(first),
      size: sz ? +sz / 100 : null, color: colorOf(first), align,
      uniform: rprs.every((r) => (/\bb="1"/.test(r) === /\bb="1"/.test(first)) && ((r.match(/\bsz="(\d+)"/)?.[1] ?? null) === (sz ?? null))),
    };
    paras.push({ text, level, ...style });
  }
  return paras;
}

function shapesOf(slideXml) {
  const out = [];
  const blocks = [...slideXml.matchAll(/<p:(sp|pic|graphicFrame|cxnSp)\b[\s\S]*?<\/p:\1>/g)];
  for (const b of blocks) {
    const xml = b[0], tag = b[1];
    const name = attr(xml.match(/<p:cNvPr\b[^>]*>/)?.[0] ?? "", "name");
    const off = xml.match(/<a:off\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"/);
    const ext = xml.match(/<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
    const geom = off && ext ? { x: +off[1] / EMU, y: +off[2] / EMU, w: +ext[1] / EMU, h: +ext[2] / EMU } : null;
    let paras = [], cells = null, colW = null, fill = null, line = null;
    if (tag === "sp" || tag === "cxnSp") {
      const spPr = xml.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/)?.[1] ?? "";
      const spPrNoLn = spPr.replace(/<a:ln\b[\s\S]*?<\/a:ln>/, "");
      fill = /<a:noFill\s*\/>/.test(spPrNoLn) ? "none" : colorOf(spPrNoLn);
      const ln = spPr.match(/<a:ln\b([^>]*)>([\s\S]*?)<\/a:ln>/);
      if (ln) line = { color: /<a:noFill\s*\/>/.test(ln[2]) ? "none" : colorOf(ln[2]), width: ln[1].match(/\bw="(\d+)"/) ? +ln[1].match(/\bw="(\d+)"/)[1] / 12700 : null };
      const tb = xml.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);
      if (tb) paras = paragraphsOf(tb[1]);
    }
    if (tag === "graphicFrame" && /<a:tbl>/.test(xml)) {
      cells = [...xml.matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map((tr) => [...tr[0].matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)].map((tc) => paragraphsOf(tc[0]).map((p) => p.text).join("\n")));
      colW = [...xml.matchAll(/<a:gridCol\b[^>]*\bw="(\d+)"/g)].map((m) => +(+m[1] / EMU).toFixed(3));
    }
    out.push({ name, tag, geom, paras, cells, colW, fill, line });
  }
  return out;
}

// ---- deck.mjs literal patching ------------------------------------------------------------------
function decodeLiteral(lit) {
  const body = lit.slice(1, -1);
  return body.replace(/\\(u[0-9a-fA-F]{4}|x[0-9a-fA-F]{2}|.)/g, (_, e) => {
    if (e[0] === "u") return String.fromCharCode(parseInt(e.slice(1), 16));
    if (e[0] === "x") return String.fromCharCode(parseInt(e.slice(1), 16));
    return { n: "\n", t: "\t", r: "\r", "0": "\0", b: "\b", f: "\f", v: "\v" }[e] ?? e;
  });
}
function encodeLiteral(text, q = '"') {
  const j = JSON.stringify(text);
  if (q === '"') return j;
  const inner = j.slice(1, -1).replace(/\\"/g, '"');
  if (q === "'") return "'" + inner.replace(/'/g, "\\'") + "'";
  return "`" + inner.replace(/`/g, "\\`").replace(/\$\{/g, "\\${") + "`";
}
const LITERAL_RE = /"(?:[^"\\\n]|\\.)*"|'(?:[^'\\\n]|\\.)*'|`(?:[^`\\$]|\\.|\$(?!\{))*`/g;

function findLiterals(value) {
  const hits = [];
  for (const m of deckSrc.matchAll(LITERAL_RE)) {
    if (m[0].length < 2) continue;
    if (decodeLiteral(m[0]) === value) hits.push({ start: m.index, end: m.index + m[0].length, q: m[0][0] });
  }
  return hits;
}

// Replace the unique literal whose value is `oldText` with `newText`.
function patchLiteral(oldText, newText) {
  if (oldText === newText) return true;
  const hits = findLiterals(oldText);
  if (hits.length !== 1) return false;
  const h = hits[0];
  deckSrc = deckSrc.slice(0, h.start) + encodeLiteral(newText, h.q) + deckSrc.slice(h.end);
  return true;
}

// Same paragraph count: whole-text literal, else paragraph by paragraph. All hits are located in the
// original source first and applied from the end, so shifted lists (A,B,C → A,C,D) patch correctly.
function patchText(oldParas, newParas) {
  const oldJoined = oldParas.map((p) => p.text).join("\n"), newJoined = newParas.map((p) => p.text).join("\n");
  if (patchLiteral(oldJoined, newJoined)) return true;
  if (oldParas.length !== newParas.length) return false;
  const edits = [];
  for (let i = 0; i < oldParas.length; i++) {
    if (oldParas[i].text === newParas[i].text) continue;
    const hits = findLiterals(oldParas[i].text);
    if (hits.length !== 1) return false;
    edits.push({ ...hits[0], text: newParas[i].text });
  }
  if (edits.some((e, k) => edits.some((f, l) => l !== k && f.start === e.start))) return false;
  edits.sort((a, b) => b.start - a.start);
  for (const e of edits) deckSrc = deckSrc.slice(0, e.start) + encodeLiteral(e.text, e.q) + deckSrc.slice(e.end);
  return true;
}

// Paragraph count changed: locate the array literal whose elements are the old paragraphs (strings or
// { text: "..." } objects, in order) and rewrite its element list.
function splitTopLevel(src) {
  // split `src` (the inside of [...]) at top-level commas, respecting strings, {}, [] and ()
  const parts = []; let depth = 0, cur = "", i = 0;
  while (i < src.length) {
    const c = src[i];
    if (c === '"' || c === "'" || c === "`") {
      const re = new RegExp(LITERAL_RE.source, "y"); re.lastIndex = i;
      const m = re.exec(src);
      if (m) { cur += m[0]; i += m[0].length; continue; }
    }
    if (c === "{" || c === "[" || c === "(") depth++;
    if (c === "}" || c === "]" || c === ")") depth--;
    if (c === "," && depth === 0) { parts.push(cur); cur = ""; i++; continue; }
    cur += c; i++;
  }
  if (cur.trim()) parts.push(cur);
  return parts;
}
function elementText(el) {
  const t = el.trim();
  if (/^["'`]/.test(t)) { const m = t.match(LITERAL_RE); return m && m[0] === t ? decodeLiteral(t) : null; }
  const m = t.match(/^\{[\s\S]*\btext:\s*("(?:[^"\\\n]|\\.)*"|'(?:[^'\\\n]|\\.)*')/);
  return m ? decodeLiteral(m[1]) : null;
}
function patchArray(oldParas, newParas) {
  if (!oldParas.length) return false;
  const hits = findLiterals(oldParas[0].text);
  for (const h of hits) {
    // walk back to the enclosing '[' at the same nesting level
    let i = h.start - 1, depth = 0, open = -1;
    while (i >= 0) {
      const c = deckSrc[i];
      if (c === "]" || c === "}" || c === ")") depth++;
      else if (c === "[" && depth === 0) { open = i; break; }
      else if ((c === "{" || c === "(") && depth === 0) break;
      else if (c === "[" || c === "{" || c === "(") depth--;
      i--;
    }
    if (open < 0) continue;
    // forward to the matching ']'
    let j = open + 1; depth = 0; let close = -1;
    while (j < deckSrc.length) {
      const c = deckSrc[j];
      if (c === '"' || c === "'" || c === "`") { const re = new RegExp(LITERAL_RE.source, "y"); re.lastIndex = j; const m = re.exec(deckSrc); if (m) { j += m[0].length; continue; } }
      if (c === "[" || c === "{" || c === "(") depth++;
      else if (c === "]" || c === "}" || c === ")") { if (depth === 0) { close = j; break; } depth--; }
      j++;
    }
    if (close < 0) continue;
    const inner = deckSrc.slice(open + 1, close);
    const elements = splitTopLevel(inner);
    const texts = elements.map(elementText);
    if (texts.length !== oldParas.length || texts.some((t, k) => t !== oldParas[k].text)) continue;
    const indent = (inner.match(/\n(\s*)/) || [null, " "])[1];
    const multiline = inner.includes("\n");
    const byText = new Map(elements.map((el, k) => [texts[k], el.trim()]));
    const rebuilt = newParas.map((p) => {
      if (byText.has(p.text)) return byText.get(p.text);
      const opts = [];
      if (p.level) opts.push(`level: ${p.level}`);
      if (p.bold) opts.push("bold: true");
      return opts.length ? `{ text: ${encodeLiteral(p.text)}, ${opts.join(", ")} }` : encodeLiteral(p.text);
    });
    const body = multiline ? "\n" + rebuilt.map((e) => indent + e + ",").join("\n") + "\n" + indent.replace(/ {2}$/, "") : rebuilt.join(", ");
    deckSrc = deckSrc.slice(0, open + 1) + body + deckSrc.slice(close);
    return true;
  }
  return false;
}

// ---- style diff ----------------------------------------------------------------------------------
function fmtDiff(mp, sp) {
  // returns an override object for one paragraph, or null if the PowerPoint paragraph matches the manifest
  if (sp.nostyle) return null;
  const d = {};
  if (sp.size != null && Math.abs(sp.size - mp.size) > 0.01) d.size = sp.size;
  if (sp.color && sp.color !== (mp.color || "").toUpperCase()) d.color = sp.color;
  if (sp.bold !== !!mp.bold) d.bold = sp.bold;
  if (sp.italic !== !!mp.italic) d.italic = sp.italic;
  if (sp.align && sp.align !== (mp.align || "left")) d.align = sp.align;
  return Object.keys(d).length ? d : null;
}

// ---- main ----------------------------------------------------------------------------------------
const zip = await JSZip.loadAsync(fs.readFileSync(PPTX));
// Only slides that are in the presentation's slide list count (a deleted slide can leave an orphan part).
const presXml = await zip.file("ppt/presentation.xml")?.async("string") ?? "";
const presRels = await zip.file("ppt/_rels/presentation.xml.rels")?.async("string") ?? "";
const relTarget = new Map([...presRels.matchAll(/<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"/g)].map((m) => [m[1], m[2]]));
const listed = [...presXml.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"/g)].map((m) => relTarget.get(m[1])).filter(Boolean).map((t) => "ppt/" + t.replace(/^\/?ppt\//, "").replace(/^\.\.\//, ""));
const slideFiles = listed.length
  ? listed.filter((f) => zip.files[f])
  : Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort((a, b) => +a.match(/(\d+)/)[1] - +b.match(/(\d+)/)[1]);
const found = new Map();
const slideOrder = [];
const notesBySlide = new Map();
for (const f of slideFiles) {
  const xml = await zip.file(f).async("string");
  const shapes = shapesOf(xml);
  let sid = null;
  for (const sh of shapes) {
    const ours = sh.name && sh.name.includes("|");
    if (ours && manifest.shapes[sh.name]) { found.set(sh.name, sh); sid = sid || sh.name.split("|")[0]; }
    else if (ours) report.warnings.push(`${f}: shape ${sh.name} is not in the manifest (stale build?)`);
    else if (sh.paras.length || sh.tag === "pic" || sh.cells) report.added.push(`${f.replace("ppt/slides/", "")}: new ${sh.tag}${sh.name ? ` "${sh.name}"` : ""}${sh.paras.length ? ` — "${sh.paras.map((p) => p.text).join(" / ").slice(0, 60)}"` : ""}`);
  }
  if (sid) slideOrder.push(sid);
  const rels = zip.file(f.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels");
  if (sid && rels) {
    const rx = await rels.async("string");
    const nm = rx.match(/Target="\.\.\/notesSlides\/(notesSlide\d+\.xml)"/);
    if (nm) {
      const nx = await zip.file("ppt/notesSlides/" + nm[1])?.async("string");
      if (nx) {
        const bodies = [...nx.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)].filter((m) => /<p:ph\b[^>]*type="body"/.test(m[0]));
        if (bodies.length) notesBySlide.set(sid, paragraphsOf(bodies[0][0].match(/<p:txBody>([\s\S]*?)<\/p:txBody>/)?.[1] ?? "").map((p) => p.text).join("\n"));
      }
    }
  }
}

const presentSlides = new Set(slideOrder);
const visibleSlides = manifest.slides.filter((s) => !s.hidden);           // slides the build actually emitted
const missingSlides = visibleSlides.filter((s) => !presentSlides.has(s.id)).map((s) => s.id);
{
  // Guard: a file from a different build (stale copy, or the manifest was replaced under it) would turn
  // every absent shape into a {hidden:true} override and mis-patch literals. Refuse unless --force.
  const expected = Object.values(manifest.shapes).filter((m) => !missingSlides.includes(m.slide)).length;
  const missingShapes = Object.keys(manifest.shapes).filter((n) => !missingSlides.includes(manifest.shapes[n].slide) && !found.has(n)).length;
  const foreign = report.warnings.filter((w) => w.includes("not in the manifest")).length;
  if (!FORCE && expected && (missingShapes > expected * 0.3 || foreign > expected * 0.3)) {
    console.error(`sync-back: ${PPTX} does not look like a save of the build in ${MANIFEST} (${missingShapes}/${expected} manifest shapes missing, ${foreign} shapes not in the manifest). Nothing written. Pass --manifest <the manifest of the build you edited> or --force.`);
    process.exit(3);
  }
}
const manifestOrder = visibleSlides.map((s) => s.id).filter((id) => presentSlides.has(id));
if (manifestOrder.join(",") !== slideOrder.join(",")) report.warnings.push("slide order differs from deck.mjs — reorder the blocks in deck.mjs to keep it");

const setOv = (name, patch) => { overrides[name] = { ...(overrides[name] || {}), ...patch }; };
// A slide deleted in PowerPoint: the build skips it from now on (theme.mjs startSlide honours
// overrides["slide|<id>"].hidden) and keeps its block in deck.mjs; remove the entry to bring it back.
for (const id of missingSlides) { setOv(`slide|${id}`, { hidden: true }); report.hidden.push(`slide "${id}" (whole slide; remove "slide|${id}" from overrides.json to restore)`); }

for (const [name, m] of Object.entries(manifest.shapes)) {
  if (missingSlides.includes(m.slide)) continue;
  const sh = found.get(name);
  if (!sh) { setOv(name, { hidden: true }); report.hidden.push(name); continue; }

  // ---- geometry
  if (sh.geom) {
    const keys = m.kind === "table" ? ["x", "y"] : ["x", "y", "w", "h"];
    const d = keys.filter((k) => Math.abs(sh.geom[k] - m[k]) > TOL);
    if (d.length) { const g = {}; for (const k of keys) g[k] = +sh.geom[k].toFixed(3); setOv(name, g); report.moved.push(`${name} (${d.join(",")})`); }
  }

  // ---- shape fill / border (shapes the build filled)
  if (m.fill !== undefined && m.fill !== null) {
    if (sh.fill && sh.fill !== "none" && sh.fill !== m.fill) { setOv(name, { fill: sh.fill }); report.shape.push(`${name} fill ${sh.fill}`); }
    if (sh.line && m.line) {
      const lp = {};
      if (sh.line.color && sh.line.color !== "none" && sh.line.color !== m.line.color) lp.color = sh.line.color;
      if (sh.line.width != null && Math.abs(sh.line.width - m.line.width) > 0.05) lp.width = +sh.line.width.toFixed(2);
      if (Object.keys(lp).length) { setOv(name, { line: { ...(overrides[name]?.line || {}), ...lp } }); report.shape.push(`${name} line ${JSON.stringify(lp)}`); }
    }
  }

  // ---- tables
  if (m.kind === "table") {
    if (sh.colW && m.colW && (sh.colW.length !== m.colW.length || sh.colW.some((w, i) => Math.abs(w - m.colW[i]) > TOL))) {
      setOv(name, { colW: sh.colW }); report.tables.push(`${name} column widths`);
    }
    if (sh.cells && m.cells) {
      const sameGrid = sh.cells.length === m.cells.length && sh.cells.every((r, i) => r.length === m.cells[i].length);
      if (sameGrid) {
        const flatOld = m.cells.flat(), flatNew = sh.cells.flat();
        let fallback = false;
        const snapshot = deckSrc;
        for (let i = 0; i < flatOld.length; i++) {
          if (flatOld[i] === flatNew[i]) continue;
          if (overrides[name]?.cells || !patchLiteral(flatOld[i], flatNew[i])) { fallback = true; break; }
          report.patched.push(`${name} cell ${i}: "${flatNew[i].slice(0, 40)}"`);
        }
        if (fallback) { deckSrc = snapshot; report.patched = report.patched.filter((p) => !p.startsWith(name + " cell")); setOv(name, { cells: sh.cells }); report.tables.push(`${name} cells (override)`); }
      } else { setOv(name, { cells: sh.cells }); report.tables.push(`${name} rows/columns ${m.cells.length}×${m.cells[0]?.length} → ${sh.cells.length}×${sh.cells[0]?.length} (override)`); }
    }
    continue;
  }

  // ---- text + formatting
  if (!m.paras || !m.paras.length) continue;
  const oldJ = m.paras.map((p) => p.text).join("\n"), newJ = sh.paras.map((p) => p.text).join("\n");
  const sameCount = sh.paras.length === m.paras.length;
  if (oldJ !== newJ) {
    const hasTextOv = overrides[name]?.text !== undefined;
    let done = false;
    if (!hasTextOv && sameCount && patchText(m.paras, sh.paras)) { report.patched.push(`${name}: "${newJ.slice(0, 60).replace(/\n/g, " ⏎ ")}"`); done = true; }
    else if (!hasTextOv && patchArray(m.paras, sh.paras)) { report.arrays.push(`${name}: ${m.paras.length} → ${sh.paras.length} paragraphs`); done = true; }
    if (!done) {
      // store the whole text (with any style that differs from the build's defaults for that paragraph)
      const base = m.paras[0];
      setOv(name, { text: sh.paras.map((p, i) => { const mp = m.paras[i] || base; const d = fmtDiff(mp, p) || {}; return { text: p.text, level: p.level, ...d }; }) });
      delete overrides[name].fmt;
      report.overridden.push(name);
      continue;
    }
  }
  // formatting (same paragraph structure as the code now produces)
  if (sameCount || oldJ === newJ) {
    const fmt = sh.paras.map((p, i) => fmtDiff(m.paras[i], p));
    const existing = overrides[name]?.fmt;
    if (fmt.some(Boolean)) { setOv(name, { fmt: fmt.map((d) => d || {}) }); report.fmt.push(`${name}: ${fmt.map((d, i) => d ? `¶${i + 1} ${Object.entries(d).map(([k, v]) => `${k}=${v}`).join(" ")}` : "").filter(Boolean).join("; ")}`); }
    else if (existing) { delete overrides[name].fmt; if (!Object.keys(overrides[name]).length) delete overrides[name]; }
  }
}

// notes
for (const s of manifest.slides) {
  const n = notesBySlide.get(s.id);
  if (n == null || n === (s.notes || "")) continue;
  if (s.notes && patchLiteral(s.notes, n)) report.notes.push(s.id);
  else report.warnings.push(`notes of "${s.id}" changed but could not be patched (add a notes: literal in deck.mjs)`);
}

// ---- write ----------------------------------------------------------------------------------------
const changedDeck = deckSrc !== fs.readFileSync(DECK, "utf8");
if (!DRY) {
  if (changedDeck) fs.writeFileSync(DECK, deckSrc);
  fs.writeFileSync(OVERRIDES, JSON.stringify(overrides, null, 2) + "\n");
}
const lines = [`sync-back${DRY ? " (dry run)" : ""}: ${PPTX}`];
const section = (title, arr) => { if (arr.length) lines.push(`  ${title} (${arr.length}):\n    ` + arr.join("\n    ")); };
section("text patched into deck.mjs", report.patched);
section("bullet lists rewritten in deck.mjs", report.arrays);
if (report.notes.length) lines.push(`  notes patched: ${report.notes.join(", ")}`);
section("formatting → overrides.json fmt", report.fmt);
section("fill / border → overrides.json", report.shape);
section("moved / resized → overrides.json", report.moved);
section("deleted → hidden in overrides.json", report.hidden);
section("tables → overrides.json", report.tables);
section("text stored as override (could not patch deck.mjs)", report.overridden);
section("ADDED in PowerPoint — not representable, add to deck.mjs", report.added);
section("warnings", report.warnings);
if (lines.length === 1) lines.push("  no differences");
console.log(lines.join("\n"));
