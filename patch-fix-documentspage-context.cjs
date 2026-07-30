// patch-fix-documentspage-context.cjs — DocumentsPage utilise useApp() au lieu d'un état local
// Usage : node patch-fix-documentspage-context.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-docspage-${Date.now()}`);

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

// ── 1 — États : data/loading viennent du contexte, plus d'état local ──
apply(
  "Patch 1 (data/loading depuis useApp)",
  `  const {showToast, pendingFicheEns, setPendingFicheEns} = useApp();
  const [data, setData]         = useState(null);
  const [loading, setLoading]   = useState(true);
  const [refreshing, setRefreshing] = useState(false);`,
  `  const {data, refreshData, showToast, pendingFicheEns, setPendingFicheEns} = useApp();
  const [refreshing, setRefreshing] = useState(false);`
);

// ── 2 — Suppression du chargement local (doublon) ───────────────────
apply(
  "Patch 2 (suppression useEffect loadData local)",
  `  useEffect(() => {
    loadData().then(d => {
      setData(d);
      setLoading(false);
    });
  }, []);
  // Navigation ciblée depuis le Tableau de bord ("voir la fiche de cet enseignant")`,
  `  // Navigation ciblée depuis le Tableau de bord ("voir la fiche de cet enseignant")`
);

// ── 3 — Actualiser passe par le contexte partagé ────────────────────
apply(
  "Patch 3 (refreshDocsData → refreshData du contexte)",
  `  const refreshDocsData = async () => {
    setRefreshing(true);
    try {
      const d = await loadData();
      setData(d);
      showToast("✓ Données actualisées — fiches à jour");
    } catch {
      showToast("⚠ Actualisation impossible — vérifiez la connexion", false);
    } finally {
      setRefreshing(false);
    }
  };`,
  `  const refreshDocsData = async () => {
    setRefreshing(true);
    try {
      await refreshData();
      showToast("✓ Données actualisées — fiches à jour");
    } catch {
      showToast("⚠ Actualisation impossible — vérifiez la connexion", false);
    } finally {
      setRefreshing(false);
    }
  };`
);

// ── 4 — Cohérence du filtre rôle (proviseur, pas animatrice) ────────
apply(
  "Patch 4 (filtre rôle cohérent)",
  `  const enseignants = data ? Object.values(data?.users||{}).filter(u => u.role !== "animatrice") : [];`,
  `  const enseignants = data ? Object.values(data?.users||{}).filter(u => u.role !== "proviseur") : [];`
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${ok}/4 patchs appliqués. DocumentsPage respecte désormais le filtre département.`);
