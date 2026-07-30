// patch-edt-dept-resync.cjs — Corrige la sélection figée dans EdtPage lors d'un changement de département
// Usage : node patch-edt-dept-resync.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-edtresync-${Date.now()}`);

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
  "Patch 1 (selEns se réajuste si invalide pour le département courant)",
  `useEffect(()=>{ if (enseignants.length>0 && !selEns) setSelEns(enseignants[0].id); }, [enseignants]);`,
  `useEffect(()=>{ if (enseignants.length>0 && !enseignants.some(e=>e.id===selEns)) setSelEns(enseignants[0].id); }, [enseignants]);`
);

apply(
  "Patch 2 (selCl se réajuste si invalide pour le département courant)",
  `useEffect(()=>{ if (toutesClasses.length>0 && !selCl) setSelCl(toutesClasses[0]); }, [toutesClasses]);`,
  `useEffect(()=>{ if (toutesClasses.length>0 && !toutesClasses.includes(selCl)) setSelCl(toutesClasses[0]); }, [toutesClasses]);`
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${ok}/2 patchs appliqués.`);
