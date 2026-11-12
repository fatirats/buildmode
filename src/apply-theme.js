// PptxGenJS can set the theme fonts but always writes Office's stock colour
// palette. This step rewrites ppt/theme/theme1.xml so the deck carries the
// BUILD MODE palette as real theme colours (editable via Design > Variants >
// Colors in PowerPoint).
const fs = require("fs");
const JSZip = require("jszip");

const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

async function applyTheme(deckPath, theme) {
  for (const k of SLOTS) {
    if (!/^[0-9A-Fa-f]{6}$/.test(String(theme.colors[k]))) {
      throw new Error(`theme.colors.${k} must be six hex digits without "#"`);
    }
  }
  const zip = await JSZip.loadAsync(fs.readFileSync(deckPath));
  const part = "ppt/theme/theme1.xml";
  const xml = await zip.file(part).async("string");
  const name = String(theme.name).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
  const scheme =
    `<a:clrScheme name="${name}">` +
    SLOTS.map((k) => `<a:${k}><a:srgbClr val="${theme.colors[k].toUpperCase()}"/></a:${k}>`).join("") +
    "</a:clrScheme>";
  const out = xml
    .replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme)
    .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, head) => `${head}${name}"`);
  if (!out.includes(scheme)) throw new Error(`${part} has no colour scheme to replace`);
  zip.file(part, out);
  // Drop bare folder entries; they are not package parts.
  for (const n of Object.keys(zip.files)) if (zip.files[n].dir) delete zip.files[n];
  fs.writeFileSync(deckPath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

module.exports = { applyTheme };
