#!/usr/bin/env node
// build.mjs — one entry point for the live preview and for the command line.
//
//   node tools/build.mjs            decide what changed and do the right thing
//   node tools/build.mjs --force    regenerate from code no matter what (after syncing an unsynced save)
//   node tools/build.mjs --no-sync  never run sync-back (render whatever .pptx is there)
//
// Two directions:
//   code → pptx : deck.mjs / theme.mjs / overrides.json / assets changed  ⇒ node deck.mjs, then render PNGs
//   pptx → code : out/deck.pptx changed (saved from PowerPoint)            ⇒ tools/sync-back.mjs, then render
//                 PNGs from the saved file (the pptx is NOT regenerated, so PowerPoint keeps its file)
//
// A state file (out/.build-state.json) remembers the hashes after each run, so the watcher's follow-up
// triggers (the build itself touches deck.pptx; sync-back touches deck.mjs) are recognised as "up to date".
//
// Guards (added after an edit was lost to a code push that replaced deck.pptx and deck.mjs together):
//   • Nothing is overwritten without a copy in out/backup/<timestamp>/ (deck.pptx, deck.mjs,
//     overrides.json, manifest.json). The last BACKUP_KEEP (20) backups are kept.
//   • If the .pptx AND the code changed since the last state, the change came from outside (a sync, a
//     checkout, a push): the .pptx is treated as the build of that code — no sync-back, manifest refreshed
//     from the code, PNGs rendered from the .pptx. The live preview reports the changed files in
//     PPTX_PREVIEW_CHANGED, which is used the same way.
//   • If the code changed while the .pptx holds an unsynced PowerPoint save, the save is synced back
//     first, then the code is rebuilt — a PowerPoint edit is never silently discarded.
//   • A sync-back that would hide most of the manifest's shapes (stale build vs. edited file) is refused
//     by sync-back itself; see its --force.

import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, "..");
const OUT = path.join(ROOT, "out");
const PPTX = path.join(OUT, "deck.pptx");
const RENDER_DIR = path.join(OUT, "rendered");
const STATE = path.join(OUT, ".build-state.json");
const BACKUP_DIR = path.join(OUT, "backup");
const BACKUP_KEEP = 20;
const FORCE = process.argv.includes("--force");
const NO_SYNC = process.argv.includes("--no-sync");
const CODE_FILES = ["deck.mjs", "theme.mjs", "overrides.json"];

const sha = (file) => (fs.existsSync(file) ? crypto.createHash("sha256").update(fs.readFileSync(file)).digest("hex") : "");
function codeSha() {
  const files = CODE_FILES.map((f) => path.join(ROOT, f));
  const assets = fs.existsSync(path.join(ROOT, "assets")) ? fs.readdirSync(path.join(ROOT, "assets")).sort().map((f) => path.join(ROOT, "assets", f)) : [];
  return crypto.createHash("sha256").update([...files, ...assets].map((f) => f + ":" + sha(f)).join("\n")).digest("hex");
}
const run = (cmd, args, env = {}) => { const r = spawnSync(cmd, args, { cwd: ROOT, stdio: "inherit", env: { ...process.env, ...env } }); if (r.status !== 0) process.exit(r.status ?? 1); };
const stamp = () => new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d+Z$/, "Z");

// Copy the files that are about to change into out/backup/<stamp>-<why>/ and prune old backups.
function backup(why) {
  const dir = path.join(BACKUP_DIR, `${stamp()}-${why}`);
  fs.mkdirSync(dir, { recursive: true });
  for (const f of [PPTX, path.join(OUT, "manifest.json"), ...CODE_FILES.map((c) => path.join(ROOT, c))]) {
    if (fs.existsSync(f)) fs.copyFileSync(f, path.join(dir, path.basename(f)));
  }
  const all = fs.readdirSync(BACKUP_DIR).filter((d) => fs.statSync(path.join(BACKUP_DIR, d)).isDirectory()).sort();
  for (const d of all.slice(0, Math.max(0, all.length - BACKUP_KEEP))) fs.rmSync(path.join(BACKUP_DIR, d), { recursive: true, force: true });
  console.log(`build: backup → ${path.relative(ROOT, dir)}`);
  return dir;
}
const writeState = (from, extra = {}) => fs.writeFileSync(STATE, JSON.stringify({ pptx: sha(PPTX), code: codeSha(), at: new Date().toISOString(), from, ...extra }, null, 2));

const state = fs.existsSync(STATE) ? JSON.parse(fs.readFileSync(STATE, "utf8")) : null;
const pptxNow = sha(PPTX), codeNow = codeSha();
const pptxChanged = !!(state && pptxNow && pptxNow !== state.pptx);
const codeChanged = !state || codeNow !== state.code;
const changedList = (process.env.PPTX_PREVIEW_CHANGED || "").split(",").filter(Boolean).map((f) => path.basename(f));
const externalByWatcher = changedList.includes("deck.pptx") && changedList.some((f) => CODE_FILES.includes(f));

if (!pptxNow) {
  console.log("build: no out/deck.pptx yet → node deck.mjs + render");
  run("node", [path.join(ROOT, "deck.mjs")]);
  run("bash", [path.join(HERE, "render.sh"), PPTX, RENDER_DIR]);
  writeState("code");
} else if (FORCE) {
  if (pptxChanged && !NO_SYNC) {
    console.log("build: --force, but out/deck.pptx holds an unsynced PowerPoint save → sync-back first");
    backup("sync");
    run("node", [path.join(HERE, "sync-back.mjs")]);
  }
  console.log("build: --force → node deck.mjs + render");
  backup("rebuild");
  run("node", [path.join(ROOT, "deck.mjs")]);
  run("bash", [path.join(HERE, "render.sh"), PPTX, RENDER_DIR]);
  writeState("code");
} else if (pptxChanged && (codeChanged || externalByWatcher)) {
  // Both sides changed at once: an external update (sync, checkout, cloud push), not a PowerPoint save.
  console.log(`build: out/deck.pptx and the code changed together (${externalByWatcher ? "watcher: " + changedList.join(", ") : "hashes"}) → external update: no sync-back; manifest from code; render from the .pptx`);
  backup("external");
  { const r = spawnSync("node", [path.join(ROOT, "deck.mjs")], { cwd: ROOT, stdio: "inherit", env: { ...process.env, DECK_MANIFEST_ONLY: "1" } }); if (r.status !== 0) process.exit(r.status ?? 1); }
  run("bash", [path.join(HERE, "render.sh"), PPTX, RENDER_DIR]);
  writeState("external", { changed: changedList });
} else if (pptxChanged) {
  if (NO_SYNC) {
    console.log("build: out/deck.pptx changed (--no-sync) → render from the saved file only");
    run("bash", [path.join(HERE, "render.sh"), PPTX, RENDER_DIR]);
    writeState("powerpoint-nosync");
  } else {
    console.log("build: out/deck.pptx changed outside the build (PowerPoint save) → sync-back + render from the saved file");
    backup("sync");
    run("node", [path.join(HERE, "sync-back.mjs")]);
    // refresh out/manifest.json from the patched code (no pptx overwrite), so the next save diffs correctly
    { const r = spawnSync("node", [path.join(ROOT, "deck.mjs")], { cwd: ROOT, stdio: "inherit", env: { ...process.env, DECK_MANIFEST_ONLY: "1" } }); if (r.status !== 0) process.exit(r.status ?? 1); }
    run("bash", [path.join(HERE, "render.sh"), PPTX, RENDER_DIR]);
    writeState("powerpoint");
  }
} else if (codeChanged) {
  console.log("build: code changed → node deck.mjs + render");
  backup("rebuild");
  run("node", [path.join(ROOT, "deck.mjs")]);
  run("bash", [path.join(HERE, "render.sh"), PPTX, RENDER_DIR]);
  writeState("code");
} else {
  console.log("build: up to date");
}
