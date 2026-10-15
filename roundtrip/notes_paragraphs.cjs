// pptxgenjs writes speaker notes as one run whose text keeps raw "\n" characters, which PowerPoint
// does not treat as paragraph breaks. This rewrites each notes body so every line is its own <a:p>
// (blank lines become empty paragraphs), so the notes read as written in PowerPoint's notes pane and
// tools/sync-back.mjs reads them back identically (paragraphs joined with "\n").
const fs = require("fs");
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));

async function splitNotes(deckPath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(deckPath));
  const parts = Object.keys(zip.files).filter((f) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(f));
  for (const f of parts) {
    let xml = await zip.file(f).async("string");
    xml = xml.replace(/<a:p><a:r>(<a:rPr\b[^>]*\/>)<a:t>([\s\S]*?)<\/a:t><\/a:r>(<a:endParaRPr\b[^>]*\/>)<\/a:p>/g, (m, rPr, text, end) => {
      if (!text.includes("\n")) return m;
      return text.split(/\r?\n/).map((line) => (line.length ? `<a:p><a:r>${rPr}<a:t>${line}</a:t></a:r>${end}</a:p>` : `<a:p>${end}</a:p>`)).join("");
    });
    zip.file(f, xml);
  }
  fs.writeFileSync(deckPath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}
module.exports = { splitNotes };
