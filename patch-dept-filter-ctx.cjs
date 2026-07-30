// patch-dept-filter-ctx.cjs — Filtrage département centralisé (Provider)
// Usage : node patch-dept-filter-ctx.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-deptfilter-${Date.now()}`);

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

// ── A — Fonction de filtrage (avant App()) ──────────────────────────
const FILTER_FN = `
// ─── Filtrage département (scope Proviseur) ────────────────────────
function filterDataByDept(data, deptId) {
  if (!data || !deptId) return data;
  const deptOf = {};
  Object.values(data.users||{}).forEach(u=>{ deptOf[u.id]=u.departement_id; });

  const users = {};
  Object.entries(data.users||{}).forEach(([id,u])=>{ if(u.departement_id===deptId) users[id]=u; });

  const prog = {};
  Object.entries(data.prog||{}).forEach(([k,v])=>{ if(deptOf[k.split("||")[0]]===deptId) prog[k]=v; });

  const absences = {};
  Object.entries(data.absences||{}).forEach(([k,v])=>{ if(deptOf[k.split("||")[0]]===deptId) absences[k]=v; });

  const epreuves = (data.epreuves||[]).filter(e=>deptOf[e.ens_id]===deptId);

  const exceptions = {};
  Object.entries(data.exceptions||{}).forEach(([ensId,v])=>{ if(deptOf[ensId]===deptId) exceptions[ensId]=v; });

  const edtBase = {};
  Object.entries(data.edtBase||{}).forEach(([ensId,v])=>{ if(deptOf[ensId]===deptId) edtBase[ensId]=v; });

  return { ...data, users, prog, absences, epreuves, exceptions, edtBase };
}

`;
apply("Patch A (fonction filterDataByDept)", `export default function App() {`, FILTER_FN + `export default function App() {`);

// ── B — État viewDeptId ──────────────────────────────────────────────
apply(
  "Patch B (état viewDeptId)",
  `const [data,setData]     = useState(null);`,
  `const [data,setData]     = useState(null);
  const [viewDeptId,setViewDeptId] = useState(null); // filtre département actif (Proviseur uniquement)`
);

// ── C — Données filtrées (avant ctx) ────────────────────────────────
apply(
  "Patch C (scopedData mémoïsé)",
  `const ctx = {user,setUser,page,setPage,data,setData,online,syncing`,
  `const scopedData = useMemo(()=>filterDataByDept(data, user?.role==="proviseur"?viewDeptId:null), [data, viewDeptId, user?.role]);

  const ctx = {user,setUser,page,setPage,data:scopedData,rawData:data,setData,viewDeptId,setViewDeptId,online,syncing`
);

// ── D — DashboardProviseur reste sur la vue globale ─────────────────
apply(
  "Patch D (DashboardProviseur → données brutes)",
  `function DashboardProviseur() {
  const {data,refreshData} = useApp();`,
  `function DashboardProviseur() {
  const {rawData:data,refreshData} = useApp();`
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${ok}/4 patchs appliqués.`);
