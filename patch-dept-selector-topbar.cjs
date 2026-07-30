// patch-dept-selector-topbar.cjs — Ajoute le sélecteur de département au Topbar
// Usage : node patch-dept-selector-topbar.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-topbar-${Date.now()}`);

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

apply(
  "Patch A (destructure viewDeptId)",
  `const {user, realtimeStatus} = useApp();`,
  `const {user, realtimeStatus, viewDeptId, setViewDeptId} = useApp();`
);

apply(
  "Patch B (sélecteur département, avant DarkModeToggle)",
  `      <DarkModeToggle/>`,
  `      {user?.role==="proviseur" && (
        <select value={viewDeptId||""} onChange={e=>setViewDeptId(e.target.value?parseInt(e.target.value):null)}
          title="Filtrer par département"
          style={{padding:"5px 10px", borderRadius:8, border:\`1px solid \${C.border}\`, background:C.white, fontSize:11, fontWeight:700, color:C.txt, fontFamily:"inherit", cursor:"pointer"}}>
          <option value="">🏛️ Tous les départements</option>
          {DEPARTEMENTS_LIST.map(d=><option key={d.id} value={d.id}>{d.emoji} {d.nom}</option>)}
        </select>
      )}
      <DarkModeToggle/>`
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${ok}/2 patchs appliqués.`);
