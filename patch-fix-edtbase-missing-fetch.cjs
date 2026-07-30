// patch-fix-edtbase-missing-fetch.cjs — Corrige l'appel edt_base manquant dans loadAllData
// Usage : node patch-fix-edtbase-missing-fetch.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-edtbasefetch-${Date.now()}`);

if (!fs.existsSync(FILE)) {
  console.error("❌ Fichier introuvable :", FILE);
  process.exit(1);
}

let src = fs.readFileSync(FILE, "utf8");
fs.writeFileSync(BACKUP, src, "utf8");
console.log("✓ Sauvegarde créée :", BACKUP);

const search = `    sb.get("absences", departementId ? \`?select=ens_id,classe,seance,absents&departement_id=eq.\${departementId}\` : "?select=ens_id,classe,seance,absents"),
  ]);`;

const replace = `    sb.get("absences", departementId ? \`?select=ens_id,classe,seance,absents&departement_id=eq.\${departementId}\` : "?select=ens_id,classe,seance,absents"),
    sb.get("edt_base", departementId ? \`?select=ens_id,slot,lbl&departement_id=eq.\${departementId}\` : "?select=ens_id,slot,lbl"),
  ]);`;

const count = src.split(search).length - 1;
if (count === 0) {
  console.error("❌ Motif introuvable — le fichier a peut-être déjà été modifié différemment.");
  process.exit(1);
}
if (count !== 1) console.warn(`⚠ ${count} occurrence(s) trouvée(s) (attendu 1).`);

src = src.split(search).join(replace);
fs.writeFileSync(FILE, src, "utf8");
console.log("✅ Requête edt_base ajoutée à loadAllData (8ᵉ position, alignée sur la déstructuration).");
