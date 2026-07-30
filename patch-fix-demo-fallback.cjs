// patch-fix-demo-fallback.cjs — Empêche le fallback DEMO_ACCOUNTS de masquer un filtre département vide
// Usage : node patch-fix-demo-fallback.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-demofallback-${Date.now()}`);

if (!fs.existsSync(FILE)) {
  console.error("❌ Fichier introuvable :", FILE);
  process.exit(1);
}

let src = fs.readFileSync(FILE, "utf8");
fs.writeFileSync(BACKUP, src, "utf8");
console.log("✓ Sauvegarde créée :", BACKUP);

let ok = 0;
function apply(label, search, replace, expected = 1) {
  const count = src.split(search).length - 1;
  if (count === 0) { console.error(`❌ ${label} — motif introuvable.`); return; }
  if (count !== expected) console.warn(`⚠ ${label} — ${count} occurrence(s) (attendu ${expected}).`);
  src = src.split(search).join(replace);
  ok++;
  console.log(`✓ ${label}`);
}

// ── 0 — Marquer les données comme "filtrées par département" ──────────
apply(
  "Patch 0 (flag deptFilterActive)",
  `  return { ...data, users, prog, absences, epreuves, exceptions, edtBase };\n}`,
  `  return { ...data, users, prog, absences, epreuves, exceptions, edtBase, deptFilterActive: true };\n}`
);

// ── 1 — EnsGerer ────────────────────────────────────────────────────
apply(
  "Patch 1 (EnsGerer)",
  `  const supabaseEns = Object.values(data?.users||{}).filter(u=>u.role!=="animatrice");
  const enseignants = supabaseEns.length > 0
    ? supabaseEns.map(u=>({...u, col:u.col||getColor(u.id), ini:u.ini||getIni(u.nom), classes:(u.classes||[]).length>0?u.classes:(ENS_CLASSES_REF[u.id]||[])}))
    : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant").map(a=>({...a, col:getColor(a.id), ini:getIni(a.nom), classes:ENS_CLASSES_REF[a.id]||[]}));`,
  `  const supabaseEns = Object.values(data?.users||{}).filter(u=>u.role!=="animatrice");
  const enseignants = (supabaseEns.length > 0 || data?.deptFilterActive)
    ? supabaseEns.map(u=>({...u, col:u.col||getColor(u.id), ini:u.ini||getIni(u.nom), classes:(u.classes||[]).length>0?u.classes:(ENS_CLASSES_REF[u.id]||[])}))
    : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant").map(a=>({...a, col:getColor(a.id), ini:getIni(a.nom), classes:ENS_CLASSES_REF[a.id]||[]}));`
);

// ── 2 — EnsListe ────────────────────────────────────────────────────
apply(
  "Patch 2 (EnsListe)",
  `  const sourceData = supabaseEns.length > 0`,
  `  const sourceData = (supabaseEns.length > 0 || data?.deptFilterActive)`
);

// ── 3 — Bloc useMemo (EdtPage ou similaire) ──────────────────────────
apply(
  "Patch 3 (useMemo enseignants)",
  `    const source = users.length>0 ? users : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant");`,
  `    const source = (users.length>0 || data?.deptFilterActive) ? users : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant");`
);

// ── 4 — SuiviProgrammePage (sans ini) ─────────────────────────────────
apply(
  "Patch 4 (SuiviProgrammePage)",
  `  const enseignants = supabaseEns.length > 0
    ? supabaseEns.map(u=>({...u, col:u.col||getColor(u.id), classes:(u.classes||[]).length>0?u.classes:(ENS_CLASSES_REF[u.id]||[])}))
    : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant").map(a=>({...a, classes:ENS_CLASSES_REF[a.id]||[]}));`,
  `  const enseignants = (supabaseEns.length > 0 || data?.deptFilterActive)
    ? supabaseEns.map(u=>({...u, col:u.col||getColor(u.id), classes:(u.classes||[]).length>0?u.classes:(ENS_CLASSES_REF[u.id]||[])}))
    : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant").map(a=>({...a, classes:ENS_CLASSES_REF[a.id]||[]}));`
);

// ── 5 — Dernière page (EpreuvesPage/DocumentsPage, sans espaces) ─────
apply(
  "Patch 5 (page compacte)",
  `  const enseignants = (supabaseEns.length>0
    ? supabaseEns.map(u=>({...u,col:u.col||getColor(u.id),ini:u.ini||getIni(u.nom),classes:(u.classes||[]).length>0?u.classes:(ENS_CLASSES_REF[u.id]||[])}))
    : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant").map(a=>({...a,col:getColor(a.id),ini:getIni(a.nom),classes:ENS_CLASSES_REF[a.id]||[]}))
  );`,
  `  const enseignants = ((supabaseEns.length>0 || data?.deptFilterActive)
    ? supabaseEns.map(u=>({...u,col:u.col||getColor(u.id),ini:u.ini||getIni(u.nom),classes:(u.classes||[]).length>0?u.classes:(ENS_CLASSES_REF[u.id]||[])}))
    : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant").map(a=>({...a,col:getColor(a.id),ini:getIni(a.nom),classes:ENS_CLASSES_REF[a.id]||[]}))
  );`
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${ok}/6 patchs appliqués.`);
