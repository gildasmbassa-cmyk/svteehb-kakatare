// patch-secure-progsuivi-delete.cjs — Remplace sb.del("prog_suivi",...) par les RPC admin_delete_prog_*
// Usage : node patch-secure-progsuivi-delete.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-progsuividelete-${Date.now()}`);

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
  console.log(`✓ ${label} (${count}x)`);
}

apply(
  "A (ens.id + classe)",
  `await sb.del("prog_suivi", \`?ens_id=eq.\${encodeURIComponent(ens.id)}&classe=eq.\${encodeURIComponent(classe)}\`);`,
  `await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: ens.id, p_classe: classe });`,
  2
);

apply(
  "B (ens.id + classe||dig)",
  `await sb.del("prog_suivi", \`?ens_id=eq.\${encodeURIComponent(ens.id)}&classe=eq.\${encodeURIComponent(classe+"||dig")}\`);`,
  `await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: ens.id, p_classe: classe+"||dig" });`,
  2
);

apply(
  "C (wipe enseignant)",
  `const ok2 = await sb.del("prog_suivi", \`?ens_id=eq.\${ens.id}\`);`,
  `const ok2 = await sb.rpc("admin_delete_prog_by_teacher", { p_ens_id: ens.id });`
);

apply(
  "D (ens_id + classe, variable ens_id)",
  `const res = await sb.del("prog_suivi", \`?ens_id=eq.\${ens_id}&classe=eq.\${encodeURIComponent(classe)}\`);`,
  `const res = await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: ens_id, p_classe: classe });`
);

apply(
  "E (wipe total)",
  `const res = await sb.del("prog_suivi", "?ens_id=not.is.null"); // supprimer toutes les progressions`,
  `const res = await sb.rpc("admin_delete_all_prog", {}); // supprimer toutes les progressions`
);

apply(
  "F (autreId + cl)",
  `await sb.del("prog_suivi", \`?ens_id=eq.\${encodeURIComponent(autreId)}&classe=eq.\${encodeURIComponent(cl)}\`);`,
  `await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: autreId, p_classe: cl });`
);

apply(
  "G (autreId + cl||dig)",
  `await sb.del("prog_suivi", \`?ens_id=eq.\${encodeURIComponent(autreId)}&classe=eq.\${encodeURIComponent(cl+"||dig")}\`);`,
  `await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: autreId, p_classe: cl+"||dig" });`
);

apply(
  "H (selEns + cl)",
  `await sb.del("prog_suivi", \`?ens_id=eq.\${encodeURIComponent(selEns)}&classe=eq.\${encodeURIComponent(cl)}\`);`,
  `await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: selEns, p_classe: cl });`
);

apply(
  "I (selEns + cl||dig)",
  `await sb.del("prog_suivi", \`?ens_id=eq.\${encodeURIComponent(selEns)}&classe=eq.\${encodeURIComponent(cl+"||dig")}\`);`,
  `await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: selEns, p_classe: cl+"||dig" });`
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${ok}/9 règles appliquées (couvrant les 11 occurrences).`);
