// patch-option-b-animateur-visible.cjs — Les animateurs comptent comme enseignants (sauf proviseur)
// Usage : node patch-option-b-animateur-visible.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-optionb-${Date.now()}`);

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

// ── 1 — Listes enseignants (5 occurrences) : exclure seulement le proviseur ──
apply(
  "Patch 1 (u.role!==\"animatrice\" → proviseur)",
  `u.role!=="animatrice"`,
  `u.role!=="proviseur"`,
  5
);

// ── 2 — Dashboards (Admin + Proviseur) : les animateurs comptent comme enseignants ──
apply(
  "Patch 2 (dashboards — animateurs inclus dans les stats)",
  `const ens=Object.values(data.users||{}).filter(u=>!isAdminRole(u.role));`,
  `const ens=Object.values(data.users||{}).filter(u=>u.role!=="proviseur");`,
  2
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${ok}/2 patchs appliqués.`);
