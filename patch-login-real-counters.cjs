// patch-login-real-counters.cjs — Remplace les compteurs figés par des données réelles
// Usage : node patch-login-real-counters.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-logincounters-${Date.now()}`);

if (!fs.existsSync(FILE)) {
  console.error("❌ Fichier introuvable :", FILE);
  process.exit(1);
}

let src = fs.readFileSync(FILE, "utf8");
fs.writeFileSync(BACKUP, src, "utf8");
console.log("✓ Sauvegarde créée :", BACKUP);

const search = `  // ── Compteurs animés au montage ──────────────────────────────────────
  useEffect(()=>{
    const targets = {eleves:1166, ens:9};
    const duration = 1800;
    const steps = 60;
    const interval = duration / steps;
    let step = 0;
    const timer = setInterval(()=>{
      step++;
      const p = Math.min(step/steps, 1);
      const ease = 1 - Math.pow(1-p, 3); // easeOutCubic
      setCounter({
        eleves: Math.round(targets.eleves * ease),
        ens:    Math.round(targets.ens    * ease),
      });
      if(step >= steps) clearInterval(timer);
    }, interval);
    return()=>clearInterval(timer);
  },[]);`;

const replace = `  // ── Compteur enseignants réel (requête légère, élèves = données locales déjà chargées) ──
  const [ensCountReal, setEnsCountReal] = useState(null);
  useEffect(()=>{
    sb.get("utilisateurs", "?select=id&role=eq.enseignant").then(rows=>{ if(rows) setEnsCountReal(rows.length); });
  },[]);

  // ── Compteurs animés au montage (déclenchés dès que le total réel est connu) ──
  useEffect(()=>{
    if (ensCountReal === null) return;
    const targets = {eleves:getTotalEleves(), ens:ensCountReal};
    const duration = 1800;
    const steps = 60;
    const interval = duration / steps;
    let step = 0;
    const timer = setInterval(()=>{
      step++;
      const p = Math.min(step/steps, 1);
      const ease = 1 - Math.pow(1-p, 3); // easeOutCubic
      setCounter({
        eleves: Math.round(targets.eleves * ease),
        ens:    Math.round(targets.ens    * ease),
      });
      if(step >= steps) clearInterval(timer);
    }, interval);
    return()=>clearInterval(timer);
  },[ensCountReal]);`;

const count = src.split(search).length - 1;
if (count === 0) {
  console.error("❌ Motif introuvable.");
  process.exit(1);
}
if (count !== 1) console.warn(`⚠ ${count} occurrence(s) (attendu 1).`);

src = src.split(search).join(replace);
fs.writeFileSync(FILE, src, "utf8");
console.log("✅ Compteurs réels branchés (élèves via ELEVES_DB, enseignants via requête).");
