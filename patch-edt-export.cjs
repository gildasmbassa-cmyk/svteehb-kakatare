// patch-edt-export.cjs — Ajoute l'export PDF (par enseignant) et Excel (groupé) de l'EDT
// Usage : node patch-edt-export.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-edtexport-${Date.now()}`);

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

// ── 1 — Fonction de génération HTML (avant EdtPage) ─────────────────
const GEN_EDT_HTML = `function genEdtHTML(ens, edtRt) {
  const rt = edtRt[ens.id] || {};
  let rows = "";
  HEURES.forEach((h, hi) => {
    rows += \`<tr><td style="padding:6px;border:1px solid #ccc;font-size:9px;font-weight:700;white-space:nowrap">\${h}</td>\`;
    JKEYS.forEach(jk => {
      const val = (rt[jk] || [])[hi] || "";
      rows += \`<td style="padding:6px;border:1px solid #ccc;font-size:9px;text-align:center">\${val}</td>\`;
    });
    rows += \`</tr>\`;
  });
  const headerCols = JOURS.map(j => \`<th style="padding:6px;border:1px solid #ccc;background:#1a6b3c;color:#fff;font-size:9px">\${j}</th>\`).join("");
  return \`<!DOCTYPE html><html><head><meta charset="utf-8"><title>EDT \${ens.nom}</title></head>
  <body style="font-family:Arial,sans-serif;padding:20px">
    \${enteteOfficiel("EMPLOI DU TEMPS", ens.nom)}
    <h3 style="text-align:center;margin:14px 0;font-size:13px">Emploi du temps — \${ens.nom}</h3>
    <table style="width:100%;border-collapse:collapse">
      <thead><tr><th style="padding:6px;border:1px solid #ccc;background:#1a6b3c;color:#fff;font-size:9px">Heures</th>\${headerCols}</tr></thead>
      <tbody>\${rows}</tbody>
    </table>
  </body></html>\`;
}

`;

apply("Patch 1 (fonction genEdtHTML)", `function EdtPage() {`, GEN_EDT_HTML + `function EdtPage() {`);

// ── 2 — Boutons export (après les onglets) ──────────────────────────
apply(
  "Patch 2 (boutons export PDF/Excel)",
  `      </div>
      {/* ── Onglet Maintenant : qui enseigne en ce moment ── */}`,
  `      </div>
      {/* ── Export EDT ── */}
      <div style={{ display:"flex", gap:8, justifyContent:"flex-end", flexWrap:"wrap" }}>
        {onglet==="parEnseignant" && ensActuel && (
          <button onClick={()=>imprimerHTML(genEdtHTML(ensActuel, edtRt))}
            style={{ display:"flex", alignItems:"center", gap:6, padding:"7px 14px", borderRadius:8, border:\`1px solid \${C.border}\`, background:C.white, color:C.txtMuted, fontSize:11, fontWeight:700, cursor:"pointer", fontFamily:"inherit" }}>
            📄 PDF ({getNomCourt(ensActuel.nom)})
          </button>
        )}
        <button onClick={()=>{
            const rows = [];
            enseignants.forEach(ens => {
              JKEYS.forEach((jk, ji) => {
                HEURES.forEach((h, hi) => {
                  const val = (edtRt[ens.id]?.[jk] || [])[hi];
                  if (val) rows.push({ Enseignant: ens.nom, Jour: JOURS[ji], Heure: h, Classe: val });
                });
              });
            });
            if (rows.length === 0) { showToast("⚠ Aucune donnée à exporter", false); return; }
            const ok2 = exportToExcel(\`EDT_\${new Date().toISOString().slice(0,10)}\`, "EDT", rows);
            showToast(ok2 ? "✓ Export Excel généré" : "⚠ Échec de l'export", ok2);
          }}
          style={{ display:"flex", alignItems:"center", gap:6, padding:"7px 14px", borderRadius:8, border:\`1px solid \${C.border}\`, background:C.white, color:C.txtMuted, fontSize:11, fontWeight:700, cursor:"pointer", fontFamily:"inherit" }}>
          📊 Excel (tous)
        </button>
      </div>
      {/* ── Onglet Maintenant : qui enseigne en ce moment ── */}`
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${ok}/2 patchs appliqués.`);
