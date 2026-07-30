// patch-fix-suiviprogramme-dept.cjs — Corrige le fallback DEMO_ACCOUNTS dans SuiviProgrammePage précisément
// Usage : node patch-fix-suiviprogramme-dept.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-suiviprog-${Date.now()}`);

if (!fs.existsSync(FILE)) {
  console.error("❌ Fichier introuvable :", FILE);
  process.exit(1);
}

let src = fs.readFileSync(FILE, "utf8");
fs.writeFileSync(BACKUP, src, "utf8");
console.log("✓ Sauvegarde créée :", BACKUP);

const FN_MARKER = "function SuiviProgrammePage() {";
const TARGET = "supabaseEns.length>0";

const startIdx = src.indexOf(FN_MARKER);
if (startIdx === -1) {
  console.error("❌ Fonction SuiviProgrammePage introuvable.");
  process.exit(1);
}

const localIdx = src.indexOf(TARGET, startIdx);
if (localIdx === -1) {
  console.error("❌ Motif 'supabaseEns.length>0' introuvable après SuiviProgrammePage — déjà patché ?");
  process.exit(1);
}

// Sanity check : le motif doit être proche du début de la fonction (< 3000 caractères), sinon on a dépassé la fonction.
if (localIdx - startIdx > 3000) {
  console.error("❌ Motif trouvé trop loin de SuiviProgrammePage — abandon par précaution.");
  process.exit(1);
}

const before = src.slice(0, localIdx);
const after = src.slice(localIdx);
const patchedAfter = after.replace(TARGET, "(supabaseEns.length>0 || data?.deptFilterActive)");

if (patchedAfter === after) {
  console.error("❌ Le remplacement n'a rien changé (motif non trouvé en tête de segment).");
  process.exit(1);
}

src = before + patchedAfter;
fs.writeFileSync(FILE, src, "utf8");
console.log("✅ SuiviProgrammePage corrigé : le filtre département s'applique désormais correctement.");
