#!/usr/bin/env node

import fs from "node:fs";
import fsp from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import { spawn } from "node:child_process";
import { pathToFileURL } from "node:url";

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function tail(text, length = 3000) {
  return text.length <= length ? text : text.slice(text.length - length);
}

function slideNumber(file) {
  return Number(file.match(/(\d+)/)?.[1] ?? Number.MAX_SAFE_INTEGER);
}

function contentType(file) {
  if (file.endsWith(".png")) return "image/png";
  if (file.endsWith(".webp")) return "image/webp";
  if (file.endsWith(".pptx")) {
    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  }
  if (file.endsWith(".json")) return "application/json";
  return "application/octet-stream";
}

async function serveFile(res, file) {
  try {
    const data = await fsp.readFile(file);
    res.writeHead(200, { "Content-Type": contentType(file), "Cache-Control": "no-store" });
    res.end(data);
  } catch {
    res.writeHead(404);
    res.end("not found");
  }
}

function normalizeBuild(build) {
  if (!build) return null;
  if (typeof build === "string") return { command: build, args: [], shell: true };
  return {
    command: build.command,
    args: build.args ?? [],
    cwd: build.cwd,
    env: build.env,
    shell: build.shell ?? false,
  };
}

export async function startLivePreview(options) {
  const pptx = path.resolve(options.pptx);
  const renderDir = path.resolve(options.renderDir);
  const watchFiles = (options.watchFiles ?? []).map((file) => path.resolve(file));
  const build = normalizeBuild(options.build);
  const settleMs = Number.isFinite(options.settleMs) ? Math.max(0, options.settleMs) : 800;
  // Files that changed since the last build was started; handed to the build command as
  // PPTX_PREVIEW_CHANGED (comma-separated absolute paths) so it can tell a PowerPoint save
  // (only the .pptx changed) from an external update that replaced the sources and the .pptx together.
  let changedFiles = new Set();
  const title = options.title ?? path.basename(pptx, path.extname(pptx));
  const host = options.host ?? "127.0.0.1";
  const port = Number(options.port ?? 8765);

  let generation = 0;
  let running = false;
  let pendingReason = null;
  let lastLog = "";
  let status = {
    state: build ? "starting" : "ready",
    generation,
    success: !build,
    reason: build ? "initial" : "static",
    startedAt: null,
    finishedAt: null,
    message: build ? "Starting first build" : "Serving existing rendered slides",
  };
  const clients = new Set();

  function publish() {
    const payload = `data: ${JSON.stringify(status)}\n\n`;
    for (const res of clients) res.write(payload);
  }

  function setStatus(next) {
    status = { ...status, ...next };
    publish();
  }

  function rebuild(reason) {
    if (!build) return;
    if (running) {
      pendingReason = reason;
      setStatus({ state: "queued", reason, message: `Queued rebuild: ${reason}` });
      return;
    }

    running = true;
    lastLog = "";
    const changed = [...changedFiles];
    changedFiles = new Set();
    setStatus({
      state: "building",
      reason,
      startedAt: new Date().toISOString(),
      finishedAt: null,
      message: `Building after ${reason}`,
    });

    const child = spawn(build.command, build.args, {
      cwd: build.cwd,
      env: { ...process.env, ...(build.env ?? {}), PPTX_PREVIEW_CHANGED: changed.join(","), PPTX_PREVIEW_REASON: reason },
      shell: build.shell,
      stdio: ["ignore", "pipe", "pipe"],
    });
    child.stdout.on("data", (chunk) => {
      lastLog += chunk.toString();
      setStatus({ message: tail(lastLog) });
    });
    child.stderr.on("data", (chunk) => {
      lastLog += chunk.toString();
      setStatus({ message: tail(lastLog) });
    });
    child.on("error", (error) => {
      running = false;
      setStatus({
        state: "error",
        success: false,
        finishedAt: new Date().toISOString(),
        message: String(error),
      });
    });
    child.on("close", (code) => {
      running = false;
      if (code === 0) generation += 1;
      setStatus({
        state: code === 0 ? "ready" : "error",
        generation,
        success: code === 0,
        finishedAt: new Date().toISOString(),
        message: code === 0
          ? tail(lastLog || "Build complete")
          : tail(lastLog || `Build failed with exit code ${code}`),
      });
      if (pendingReason) {
        const next = pendingReason;
        pendingReason = null;
        setTimeout(() => rebuild(next), 120);
      }
    });
  }

  // A change starts (or extends) a settle window; the build runs once the watched files have been quiet
  // for `settleMs`. A PowerPoint save and a multi-file sync (scp, git checkout, a cloud push) each
  // arrive as one build this way, with every changed file listed in PPTX_PREVIEW_CHANGED.
  let scheduleTimer = null;
  function schedule(file) {
    changedFiles.add(file);
    clearTimeout(scheduleTimer);
    scheduleTimer = setTimeout(() => {
      const names = [...changedFiles].map((f) => path.basename(f));
      rebuild(names.length > 1 ? `${names.length} files (${names.join(", ")})` : names[0] ?? "change");
    }, Math.max(settleMs, 250));
  }

  function startWatchers() {
    for (const file of watchFiles) {
      if (!fs.existsSync(file)) continue;
      fs.watchFile(file, { interval: 500 }, (current, previous) => {
        if (current.mtimeMs !== previous.mtimeMs) schedule(file);
      });
    }
  }

  async function listSlides() {
    try {
      const files = await fsp.readdir(renderDir);
      return files
        .filter((file) => /^slide-\d+\.png$/.test(file))
        .sort((a, b) => slideNumber(a) - slideNumber(b));
    } catch {
      return [];
    }
  }

  async function renderIndex() {
    const slides = await listSlides();
    const cards = slides.map((file) => `
      <figure>
        <img src="/rendered/${encodeURIComponent(file)}?g=${generation}" alt="${escapeHtml(file)}">
        <figcaption>${escapeHtml(file)}</figcaption>
      </figure>
    `).join("\n");
    const download = fs.existsSync(pptx) ? '<a href="/pptx">download/open PPTX</a>' : "";

    return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(title)} live preview</title>
  <style>
    :root { --accent: #2457d6; --dark: #16213b; --error: #b42318; }
    * { box-sizing: border-box; }
    body { margin: 0; background: #f4f6fa; color: #111827; font-family: Helvetica, Arial, sans-serif; }
    header { position: sticky; top: 0; z-index: 2; display: flex; align-items: center; justify-content: space-between; gap: 18px; padding: 14px 22px; background: var(--dark); color: white; box-shadow: 0 2px 10px rgba(0,0,0,.18); }
    h1 { margin: 0; font-size: 18px; }
    .meta, .view-controls { display: flex; align-items: center; gap: 10px; }
    .pill { border-radius: 999px; padding: 5px 10px; background: var(--accent); font-weight: 700; }
    .pill.error { background: var(--error); }
    main { padding: 22px 28px 40px; }
    .grid { display: grid; grid-template-columns: minmax(720px, 1120px); justify-content: center; gap: 32px; }
    body.grid-2 .grid { grid-template-columns: repeat(2, minmax(420px, 1fr)); justify-content: stretch; }
    body.grid-3 .grid { grid-template-columns: repeat(3, minmax(300px, 1fr)); justify-content: stretch; }
    figure { margin: 0; background: white; border: 1px solid #d8dce5; box-shadow: 0 3px 14px rgba(0,0,0,.07); }
    figure img { display: block; width: 100%; aspect-ratio: 16 / 9; object-fit: contain; background: white; }
    figcaption { padding: 7px 10px; color: #343741; font-size: 12px; border-top: 1px solid #edf0f5; }
    pre { white-space: pre-wrap; max-height: 140px; overflow: auto; margin: 0 0 20px; padding: 12px; background: white; border-left: 5px solid var(--accent); color: #343741; font-size: 12px; }
    a { color: white; }
    button { border: 1px solid rgba(255,255,255,.5); border-radius: 999px; background: rgba(255,255,255,.12); color: white; padding: 5px 9px; font: inherit; cursor: pointer; }
    button.active { background: white; color: var(--dark); font-weight: 700; }
    .empty { padding: 50px; text-align: center; color: #667085; background: white; border: 1px dashed #c9ced8; }
    @media (max-width: 860px) {
      main { padding: 14px 12px 30px; }
      .grid, body.grid-2 .grid, body.grid-3 .grid { grid-template-columns: 1fr; }
      header { align-items: flex-start; flex-direction: column; }
      .meta { flex-wrap: wrap; }
    }
  </style>
</head>
<body>
  <header>
    <h1>${escapeHtml(title)} live preview</h1>
    <div class="meta">
      <span id="state" class="pill">${escapeHtml(status.state)}</span>
      <span id="gen">generation ${generation}</span>
      <span class="view-controls" aria-label="slide grid size">
        <button type="button" data-grid="1" class="active">1 up</button>
        <button type="button" data-grid="2">2 up</button>
        <button type="button" data-grid="3">3 up</button>
      </span>
      ${download}
    </div>
  </header>
  <main>
    <pre id="log">${escapeHtml(status.message)}</pre>
    <section class="grid">${cards || '<div class="empty">No slide images found</div>'}</section>
  </main>
  <script>
    let currentGeneration = ${generation};
    const state = document.getElementById("state");
    const gen = document.getElementById("gen");
    const log = document.getElementById("log");
    const params = new URLSearchParams(location.search);
    const storageKey = "pptx-live-preview-grid";
    const initialGrid = params.get("grid") || localStorage.getItem(storageKey) || "1";
    function setGrid(columns) {
      document.body.classList.toggle("grid-2", columns === "2");
      document.body.classList.toggle("grid-3", columns === "3");
      document.querySelectorAll("button[data-grid]").forEach((button) => {
        button.classList.toggle("active", button.dataset.grid === columns);
      });
      localStorage.setItem(storageKey, columns);
    }
    document.querySelectorAll("button[data-grid]").forEach((button) => {
      button.addEventListener("click", () => setGrid(button.dataset.grid));
    });
    setGrid(["1", "2", "3"].includes(initialGrid) ? initialGrid : "1");
    const events = new EventSource("/events");
    events.onmessage = (event) => {
      const next = JSON.parse(event.data);
      state.textContent = next.state;
      state.classList.toggle("error", next.state === "error");
      gen.textContent = "generation " + next.generation;
      log.textContent = next.message || "";
      if (next.success && next.generation !== currentGeneration) location.reload();
    };
  </script>
</body>
</html>`;
  }

  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url ?? "/", `http://${host}:${port}`);
    if (url.pathname === "/") {
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" });
      res.end(await renderIndex());
      return;
    }
    if (url.pathname === "/events") {
      res.writeHead(200, {
        "Content-Type": "text/event-stream",
        "Cache-Control": "no-store",
        Connection: "keep-alive",
      });
      clients.add(res);
      res.write(`data: ${JSON.stringify(status)}\n\n`);
      req.on("close", () => clients.delete(res));
      return;
    }
    if (url.pathname === "/status.json") {
      res.writeHead(200, { "Content-Type": "application/json", "Cache-Control": "no-store" });
      res.end(JSON.stringify(status, null, 2));
      return;
    }
    if (url.pathname === "/pptx") {
      await serveFile(res, pptx);
      return;
    }
    if (url.pathname.startsWith("/rendered/")) {
      const requested = path.basename(decodeURIComponent(url.pathname));
      await serveFile(res, path.join(renderDir, requested));
      return;
    }
    res.writeHead(404);
    res.end("not found");
  });

  server.listen(port, host, () => {
    console.log(`Live preview: http://${host}:${port}/`);
    console.log(`PPTX: ${pptx}`);
    console.log(`Rendered slides: ${renderDir}`);
    console.log(`Watching ${watchFiles.filter((file) => fs.existsSync(file)).length} files`);
    startWatchers();
    if (build) rebuild("initial");
  });

  function close() {
    for (const file of watchFiles) fs.unwatchFile(file);
    clearTimeout(scheduleTimer);
    server.close();
  }

  process.once("SIGINT", close);
  process.once("SIGTERM", close);
  return server;
}

function usage() {
  console.log(`Usage:
  node live-preview.mjs --pptx FILE --render-dir DIR [options]

Options:
  --watch FILE          Rebuild when FILE changes. Repeatable.
  --build-command CMD   Shell command that regenerates the PPTX and slide PNGs.
  --settle MS           Quiet time after the last change before a build starts. Default: 800.
                        Files changed in that window are passed to the build command as
                        PPTX_PREVIEW_CHANGED (comma-separated paths); the build reason as PPTX_PREVIEW_REASON.
  --title TEXT          Browser title. Defaults to the PPTX filename.
  --host HOST           Listen address. Default: 127.0.0.1.
  --port PORT           Listen port. Default: 8765.
  --help                Show this help.

Rendered images must be named slide-1.png, slide-2.png, and so on.
Round trip (edit the .pptx in PowerPoint, get the edits back into the code): see roundtrip/README.md.`);
}

function parseArgs(argv) {
  const options = { watchFiles: [] };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--help") return { help: true };
    const value = argv[index + 1];
    if (!value) throw new Error(`Missing value for ${argument}`);
    if (argument === "--pptx") options.pptx = value;
    else if (argument === "--render-dir") options.renderDir = value;
    else if (argument === "--watch") options.watchFiles.push(value);
    else if (argument === "--build-command") options.build = value;
    else if (argument === "--title") options.title = value;
    else if (argument === "--host") options.host = value;
    else if (argument === "--port") options.port = Number(value);
    else if (argument === "--settle") options.settleMs = Number(value);
    else throw new Error(`Unknown option: ${argument}`);
    index += 1;
  }
  return options;
}

const isMain = process.argv[1] && pathToFileURL(path.resolve(process.argv[1])).href === import.meta.url;
if (isMain) {
  try {
    const options = parseArgs(process.argv.slice(2));
    if (options.help) {
      usage();
    } else if (!options.pptx || !options.renderDir) {
      usage();
      process.exitCode = 2;
    } else {
      await startLivePreview(options);
    }
  } catch (error) {
    console.error(error.message);
    process.exitCode = 2;
  }
}
