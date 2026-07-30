// patch-fix-documentspage-context-v3.cjs — DocumentsPage utilise useApp() (ligne par ligne, robuste CRLF)
// Usage : node patch-fix-documentspage-context-v3.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-docspagev3-${Date.now()}`);

if (!fs.existsSync(FILE)) {
  console.error("❌ Fichier introuvable :", FILE);
  process.exit(1);
}

const raw = fs.readFileSync(FILE, "utf8");
fs.writeFileSync(BACKUP, raw, "utf8");
console.log("✓ Sauvegarde créée :", BACKUP);

const eol = raw.includes("\r\n") ? "\r\n" : "\n";
const lines = raw.split(/\r\n|\n/);

const dpLineIdx = lines.findIndex(l => l.trim() === "function DocumentsPage() {");
if (dpLineIdx === -1) { console.error("❌ DocumentsPage introuvable."); process.exit(1); }

const done = {};
const LIMIT = dpLineIdx + 400;

for (let i = dpLineIdx; i < lines.length && i < LIMIT; i++) {
  if (lines[i] === null) continue;
  const t = lines[i].trim();

  if (!done.l1 && t === `const {showToast, pendingFicheEns, setPendingFicheEns} = useApp();`) {
    lines[i] = lines[i].replace(
      `const {showToast, pendingFicheEns, setPendingFicheEns} = useApp();`,
      `const {data, refreshData, showToast, pendingFicheEns, setPendingFicheEns} = useApp();`
    );
    done.l1 = i+1; continue;
  }

  if (!done.l2 && t === `const [data, setData]         = useState(null);`) {
    lines[i] = null; done.l2 = i+1; continue;
  }

  if (!done.l3 && t === `const [loading, setLoading]   = useState(true);`) {
    lines[i] = null; done.l3 = i+1; continue;
  }

  if (!done.l4 && t === `useEffect(() => {` && lines[i+1] && lines[i+1].trim() === `loadData().then(d => {`) {
    // Bloc de 6 lignes : useEffect(()=>{ / loadData().then(d=>{ / setData(d); / setLoading(false); / }); / }, []);
    for (let k = i; k <= i+5; k++) lines[k] = null;
    done.l4 = `${i+1}-${i+6}`; continue;
  }

  if (!done.l5 && t === `const d = await loadData();` && lines[i+1] && lines[i+1].trim() === `setData(d);`) {
    lines[i] = lines[i].replace(`const d = await loadData();`, `await refreshData();`);
    lines[i+1] = null;
    done.l5 = `${i+1}-${i+2}`; continue;
  }

  if (!done.l6 && t.startsWith(`const enseignants = data ? Object.values(data?.users||{}).filter(u => u.role !==`)) {
    lines[i] = lines[i].replace(`"animatrice"`, `"proviseur"`);
    done.l6 = i+1; continue;
  }
}

["l1","l2","l3","l4","l5","l6"].forEach(k=>{
  if (done[k]) console.log(`✓ ${k} — ligne(s) ${done[k]}`);
  else console.error(`❌ ${k} — motif introuvable, non appliqué.`);
});

const out = lines.filter(l => l !== null).join(eol);
fs.writeFileSync(FILE, out, "utf8");

const okCount = ["l1","l2","l3","l4","l5","l6"].filter(k=>done[k]).length;
console.log(`\n✅ ${okCount}/6 patchs appliqués.`);
