// patch-app.js — Applique les 4 patchs "département" sur src/App.jsx
// Usage : node patch-app.js
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-${Date.now()}`);

if (!fs.existsSync(FILE)) {
  console.error("❌ Fichier introuvable :", FILE, "— lance ce script depuis la racine du projet.");
  process.exit(1);
}

let src = fs.readFileSync(FILE, "utf8");
fs.writeFileSync(BACKUP, src, "utf8");
console.log("✓ Sauvegarde créée :", BACKUP);

let patchCount = 0;

function apply(label, search, replace, expectedOccurrences = 1) {
  const count = src.split(search).length - 1;
  if (count === 0) {
    console.error(`❌ ${label} — motif introuvable, patch ignoré.`);
    return;
  }
  if (count !== expectedOccurrences) {
    console.warn(`⚠ ${label} — ${count} occurrence(s) trouvée(s) (attendu ${expectedOccurrences}), patch quand même appliqué.`);
  }
  src = src.split(search).join(replace);
  patchCount++;
  console.log(`✓ ${label} — ${count} remplacement(s)`);
}

// ── Patch 1 — Helper de rôle ────────────────────────────────────────
apply(
  "Patch 1 (helper isAdminRole)",
  `const useApp = () => useContext(AppCtx);`,
  `const useApp = () => useContext(AppCtx);
const ADMIN_ROLES = ["animatrice", "animateur", "proviseur"];
const isAdminRole = (role) => ADMIN_ROLES.includes(role);`
);

// ── Patch 2 — Signature loadAllData + filtre département ───────────
apply(
  "Patch 2a (signature loadAllData)",
  `async function loadAllData() {`,
  `async function loadAllData(departementId = null) {`
);

apply(
  "Patch 2b (utilisateurs)",
  `sb.get("utilisateurs","?select=id,nom,role,classes,photo"),`,
  `sb.get("utilisateurs", departementId ? \`?select=id,nom,role,classes,photo,departement_id&departement_id=eq.\${departementId}\` : "?select=id,nom,role,classes,photo,departement_id"),`
);

apply(
  "Patch 2c (prog_suivi)",
  `sb.get("prog_suivi","?select=ens_id,classe,faites"),`,
  `sb.get("prog_suivi", departementId ? \`?select=ens_id,classe,faites&departement_id=eq.\${departementId}\` : "?select=ens_id,classe,faites"),`
);

apply(
  "Patch 2d (absences)",
  `sb.get("absences","?select=ens_id,classe,seance,absents"),`,
  `sb.get("absences", departementId ? \`?select=ens_id,classe,seance,absents&departement_id=eq.\${departementId}\` : "?select=ens_id,classe,seance,absents"),`
);

// ── Patch 3 — isAdmin utilise le helper (2 variantes de formatage) ─
apply(
  "Patch 3a (isAdmin, avec espaces)",
  `const isAdmin = user?.role === "animatrice";`,
  `const isAdmin = isAdminRole(user?.role);`,
  2 // lignes 4659 et 6512
);

apply(
  "Patch 3b (isAdmin, sans espaces)",
  `const isAdmin = user?.role==="animatrice";`,
  `const isAdmin = isAdminRole(user?.role);`
);

// ── Patch 4 — Scoping département au login / refresh ───────────────
apply(
  "Patch 4a (déclaration deptIdRef)",
  `const refreshData = useCallback(async(silent=false)=>{`,
  `const deptIdRef = useRef(null);
  const refreshData = useCallback(async(silent=false)=>{`
);

apply(
  "Patch 4b (refreshData utilise deptIdRef)",
  `    setSyncing(true);
    try {
      const d = await loadAllData();
      if (d) {`,
  `    setSyncing(true);
    try {
      const d = await loadAllData(deptIdRef.current);
      if (d) {`
);

apply(
  "Patch 4c (handleLogin fixe deptIdRef puis charge)",
  `await loadStaticData(); // Charger les données statiques en parallèle
    const d = await loadAllData();`,
  `await loadStaticData(); // Charger les données statiques en parallèle
    deptIdRef.current = isAdminRole(acc.role) && acc.role !== "proviseur" ? acc.departement_id : null;
    const d = await loadAllData(deptIdRef.current);`
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${patchCount}/9 patchs appliqués sur ${FILE}`);
console.log(`ℹ Vérifie que "useRef" est bien importé depuis 'react' en haut du fichier.`);
