# PPTX Live Preview

A dependency-free local web viewer for generated PowerPoint decks. It serves
pre-rendered slide PNGs, optionally runs an external build command, and reloads
the browser after watched source files change.

The viewer provides:

- one-, two-, and three-column slide layouts;
- live build status and build logs;
- automatic browser reload after successful rebuilds;
- a link to open or download the generated PPTX.

The viewer does not contain slide content. The optional `roundtrip/` scripts
connect it to a code-generated deck (pptxgenjs) in both directions.
Your build command remains responsible for producing:

- a `.pptx` file;
- a directory containing `slide-1.png`, `slide-2.png`, and so on.

## Static preview

```bash
node live-preview.mjs \
  --pptx /absolute/path/to/deck.pptx \
  --render-dir /absolute/path/to/rendered
```

Open <http://127.0.0.1:8765/>.

## Live rebuild

```bash
node live-preview.mjs \
  --pptx /absolute/path/to/deck.pptx \
  --render-dir /absolute/path/to/rendered \
  --watch /absolute/path/to/deck-source.mjs \
  --build-command '/absolute/path/to/build-deck.sh'
```

Repeat `--watch FILE` to monitor more than one source file. The build command
can use any presentation toolchain as long as it updates the PPTX and rendered
slide PNGs. Changes are collected for `--settle` milliseconds (default 800) and
built once; the files that changed in that window are passed to the build
command as `PPTX_PREVIEW_CHANGED` (comma-separated paths).

## Round trip: edit in PowerPoint, keep the code

Watch the `.pptx` too, and point `--build-command` at `roundtrip/build.mjs`
(copied into the deck project as `tools/`). A save from PowerPoint is then
synced back into the generator (`deck.mjs` / `overrides.json`) and re-rendered;
a code change regenerates the `.pptx`; a change that replaces both at once (a
sync or a checkout) is rendered without touching the code. Every overwrite is
preceded by a backup in `out/backup/`. See [roundtrip/README.md](roundtrip/README.md).

## Install as a command

After cloning the repository:

```bash
npm link
pptx-live-preview --help
```

The project has no runtime package dependencies.

## Options

```text
--title TEXT          Browser title
--host HOST           Listen address, default 127.0.0.1
--port PORT           Listen port, default 8765
--watch FILE          Watched source file, repeatable
--build-command CMD   External build command
--settle MS           Quiet time before a build starts (default 800)
```

Use `--host 0.0.0.0` only on a trusted network. The server has no authentication.

## Requirements

- Node.js 18 or newer
- slide images rendered as `slide-1.png`, `slide-2.png`, and so on

The viewer intentionally does not render PowerPoint files itself. This keeps it
independent of PowerPoint, LibreOffice, and any specific presentation generator.

## License

MIT
