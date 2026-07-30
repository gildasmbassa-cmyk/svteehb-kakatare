// patch-secure-matieres-crud.cjs — Remplace les écritures directes matieres par des RPC
// Usage : node patch-secure-matieres-crud.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-securematieres-${Date.now()}`);

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
  "Patch 1 (ajouterMatiere → RPC)",
  `    const ok = await sb.upsert("matieres", { nom: newMatiere.trim(), departement_id: deptId });`,
  `    const ok = await sb.rpc("admin_add_matiere", { p_nom: newMatiere.trim(), p_departement_id: deptId });`
);

apply(
  "Patch 2 (supprimerMatiere → RPC)",
  `    const ok = await sb.del("matieres", \`?id=eq.\${m.id}\`);`,
  `    const ok = await sb.rpc("admin_delete_matiere", { p_id: m.id });`
);

apply(
  "Patch 3 (renommerMatiere → RPC)",
  `    const ok = await sb.patchRow("matieres", m.id, { nom: nom.trim() });`,
  `    const ok = await sb.rpc("admin_rename_matiere", { p_id: m.id, p_nom: nom.trim() });`
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${ok}/3 patchs appliqués.`);
