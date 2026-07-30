// patch-fix-documentspage-loading-ref.cjs — Remplace {loading ? ...} par {!data ? ...}
// Usage : node patch-fix-documentspage-loading-ref.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-docsloadingref-${Date.now()}`);

if (!fs.existsSync(FILE)) {
  console.error("❌ Fichier introuvable :", FILE);
  process.exit(1);
}

const raw = fs.readFileSync(FILE, "utf8");
fs.writeFileSync(BACKUP, raw, "utf8");
console.log("✓ Sauvegarde créée :", BACKUP);

const eol = raw.includes("\r\n") ? "\r\n" : "\n";
const lines = raw.split(/\r\n|\n/);

const idx = lines.findIndex(l => l.trim() === "{loading ? (");
if (idx === -1) {
  console.error("❌ Motif '{loading ? (' introuvable.");
  process.exit(1);
}

lines[idx] = lines[idx].replace("{loading ? (", "{!data ? (");
console.log(`✓ Corrigé à la ligne ${idx+1} : {loading ? ( → {!data ? (`);

fs.writeFileSync(FILE, lines.join(eol), "utf8");
console.log("✅ Terminé.");
