// fix-patch4.cjs — Corrige patch 4b et 4c en tenant compte du CRLF Windows
// Usage : node fix-patch4.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-fix4-${Date.now()}`);

if (!fs.existsSync(FILE)) {
  console.error("❌ Fichier introuvable :", FILE);
  process.exit(1);
}

const raw = fs.readFileSync(FILE, "utf8");
fs.writeFileSync(BACKUP, raw, "utf8");
console.log("✓ Sauvegarde créée :", BACKUP);

const eol = raw.includes("\r\n") ? "\r\n" : "\n";
const lines = raw.split(/\r\n|\n/);

let fixed4b = false;
let fixed4c = false;

for (let i = 0; i < lines.length; i++) {
  const line = lines[i];
  const trimmed = line.trim();

  // Patch 4b : "const d = await loadAllData();" précédée par "try {" (refreshData)
  if (
    !fixed4b &&
    trimmed === "const d = await loadAllData();" &&
    lines[i - 1] &&
    lines[i - 1].trim() === "try {"
  ) {
    lines[i] = line.replace("loadAllData()", "loadAllData(deptIdRef.current)");
    fixed4b = true;
    console.log(`✓ Patch 4b appliqué (ligne ${i + 1})`);
    continue;
  }

  // Patch 4c : "const d = await loadAllData();" précédée par une ligne contenant "loadStaticData"
  if (
    !fixed4c &&
    trimmed === "const d = await loadAllData();" &&
    lines[i - 1] &&
    lines[i - 1].includes("loadStaticData")
  ) {
    const indent = line.match(/^(\s*)/)[1];
    lines[i] =
      `${indent}deptIdRef.current = isAdminRole(acc.role) && acc.role !== "proviseur" ? acc.departement_id : null;` +
      eol +
      line.replace("loadAllData()", "loadAllData(deptIdRef.current)");
    fixed4c = true;
    console.log(`✓ Patch 4c appliqué (ligne ${i + 1})`);
    continue;
  }
}

if (!fixed4b) console.error("❌ Patch 4b — motif introuvable.");
if (!fixed4c) console.error("❌ Patch 4c — motif introuvable.");

fs.writeFileSync(FILE, lines.join(eol), "utf8");
console.log(`\n✅ Terminé — ${[fixed4b, fixed4c].filter(Boolean).length}/2 patchs corrigés.`);
