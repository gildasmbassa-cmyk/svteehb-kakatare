// patch-page-departements.cjs — Ajoute la page de gestion Départements/Matières (proviseur)
// Usage : node patch-page-departements.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-deptpage-${Date.now()}`);

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

// ── 1 — Nav proviseur (NAV_ADMIN + entrée Départements) + titre de page ──
apply(
  "Patch 1 (NAV_PROVISEUR + PAGE_TITLES)",
  `const PAGE_TITLES = {`,
  `const NAV_PROVISEUR = [...NAV_ADMIN, {id:"departements", emoji:"🏛️", label:"Départements", sub:"Matières · Animateurs"}];

const PAGE_TITLES = {
  departements:"Départements",`
);

// ── 2 — Sidebar : nav à 3 branches ──────────────────────────────────
apply(
  "Patch 2 (Sidebar nav proviseur)",
  `const nav = isAdmin ? NAV_ADMIN : NAV_TEACHER;`,
  `const nav = user?.role==="proviseur" ? NAV_PROVISEUR : isAdmin ? NAV_ADMIN : NAV_TEACHER;`
);

// ── 3 — Route de page ────────────────────────────────────────────────
apply(
  "Patch 3 (route departements)",
  `if(page==="gestion-annuelle")  return <W>{isAdmin?<GestionAnnuellePage/>:null}</W>`,
  `if(page==="gestion-annuelle")  return <W>{isAdmin?<GestionAnnuellePage/>:null}</W>
    if(page==="departements")      return <W>{user?.role==="proviseur"?<DepartementsPage/>:null}</W>`
);

// ── 4 — Composant DepartementsPage (avant ChangePasswordPage) ────────
const DEPARTEMENTS_PAGE = `function DepartementsPage() {
  const {data, showToast} = useApp();
  const {isMobile} = useDevice();
  const [matieres, setMatieres] = useState(null);
  const [loading, setLoading] = useState(true);
  const [openDept, setOpenDept] = useState(null);
  const [newMatiere, setNewMatiere] = useState("");
  const [editingNom, setEditingNom] = useState(null);
  const [savingId, setSavingId] = useState(null);

  const loadMatieres = async () => {
    setLoading(true);
    const rows = await sb.get("matieres", "?select=id,nom,departement_id&order=departement_id,nom");
    setMatieres(rows || []);
    setLoading(false);
  };
  useEffect(() => { loadMatieres(); }, []);

  const nbEnsParDept = {};
  Object.values(data?.users||{}).filter(u=>u.role!=="proviseur").forEach(u=>{
    const d = u.departement_id||1;
    nbEnsParDept[d] = (nbEnsParDept[d]||0) + 1;
  });

  const ajouterMatiere = async (deptId) => {
    if (!newMatiere.trim()) return;
    setSavingId(\`new-\${deptId}\`);
    const ok = await sb.upsert("matieres", { nom: newMatiere.trim(), departement_id: deptId });
    if (ok) { showToast(\`✓ \${newMatiere.trim()} ajoutée\`); setNewMatiere(""); await loadMatieres(); }
    else showToast("⚠ Échec de l'ajout", false);
    setSavingId(null);
  };

  const supprimerMatiere = async (m) => {
    if (!window.confirm(\`Supprimer "\${m.nom}" ?\`)) return;
    setSavingId(m.id);
    const ok = await sb.del("matieres", \`?id=eq.\${m.id}\`);
    if (ok) { showToast(\`✓ \${m.nom} supprimée\`); await loadMatieres(); }
    else showToast("⚠ Échec de la suppression", false);
    setSavingId(null);
  };

  const renommerMatiere = async (m, nom) => {
    if (!nom.trim() || nom.trim()===m.nom) { setEditingNom(null); return; }
    setSavingId(m.id);
    const ok = await sb.patchRow("matieres", m.id, { nom: nom.trim() });
    if (ok) { showToast("✓ Renommée"); await loadMatieres(); }
    else showToast("⚠ Échec du renommage", false);
    setEditingNom(null); setSavingId(null);
  };

  return (
    <div style={{padding:"20px 20px 40px", display:"flex", flexDirection:"column", gap:16}}>
      <div>
        <h2 style={{fontSize:18, fontWeight:800, color:C.txt, margin:0}}>🏛️ Départements & Matières</h2>
        <p style={{color:C.txtMuted, margin:"3px 0 0", fontSize:12}}>8 départements pédagogiques · Lycée de Kakatare</p>
      </div>

      <div style={{display:"grid", gridTemplateColumns: isMobile?"1fr":"repeat(auto-fill, minmax(280px, 1fr))", gap:12}}>
        {DEPARTEMENTS_LIST.map(d => {
          const deptMatieres = (matieres||[]).filter(m=>m.departement_id===d.id);
          const isOpen = openDept===d.id;
          return (
            <div key={d.id} style={{background:C.white, borderRadius:12, border:\`1px solid \${C.border}\`, padding:16}}>
              <div onClick={()=>setOpenDept(isOpen?null:d.id)} style={{display:"flex", alignItems:"center", gap:10, cursor:"pointer"}}>
                <span style={{fontSize:20}}>{d.emoji}</span>
                <div style={{flex:1}}>
                  <div style={{fontSize:13, fontWeight:700, color:C.txt}}>{d.nom}</div>
                  <div style={{fontSize:10.5, color:C.txtMuted}}>{nbEnsParDept[d.id]||0} enseignant{(nbEnsParDept[d.id]||0)>1?"s":""} · {deptMatieres.length} matière{deptMatieres.length>1?"s":""}</div>
                </div>
                <span style={{fontSize:12, color:C.txtMuted}}>{isOpen?"▲":"▼"}</span>
              </div>

              {isOpen && (
                <div style={{marginTop:12, paddingTop:12, borderTop:\`1px solid \${C.border}\`, display:"flex", flexDirection:"column", gap:6}}>
                  {loading ? <Sk h={16} w="60%"/> : deptMatieres.length===0 ? (
                    <div style={{fontSize:11, color:C.txtLight, fontStyle:"italic"}}>Aucune matière</div>
                  ) : deptMatieres.map(m => (
                    <div key={m.id} style={{display:"flex", alignItems:"center", gap:8, padding:"6px 8px", background:"#f8fafc", borderRadius:7}}>
                      {editingNom===m.id ? (
                        <input autoFocus defaultValue={m.nom}
                          onBlur={e=>renommerMatiere(m, e.target.value)}
                          onKeyDown={e=>{ if(e.key==="Enter") e.target.blur(); if(e.key==="Escape") setEditingNom(null); }}
                          style={{flex:1, border:\`1px solid \${C.green}\`, borderRadius:5, padding:"3px 6px", fontSize:11.5, fontFamily:"inherit"}}/>
                      ) : (
                        <span onClick={()=>setEditingNom(m.id)} style={{flex:1, fontSize:11.5, color:C.txt, cursor:"pointer"}}>{m.nom}</span>
                      )}
                      <button onClick={()=>supprimerMatiere(m)} disabled={savingId===m.id}
                        style={{border:"none", background:"transparent", color:C.red, cursor:"pointer", fontSize:13, padding:2}}>
                        {savingId===m.id ? <Spinner size={11} color={C.red}/> : "🗑️"}
                      </button>
                    </div>
                  ))}

                  <div style={{display:"flex", gap:6, marginTop:4}}>
                    <input value={openDept===d.id?newMatiere:""} onChange={e=>setNewMatiere(e.target.value)}
                      onKeyDown={e=>{ if(e.key==="Enter") ajouterMatiere(d.id); }}
                      placeholder="Nouvelle matière…"
                      style={{flex:1, border:\`1px solid \${C.border}\`, borderRadius:6, padding:"6px 8px", fontSize:11.5, fontFamily:"inherit"}}/>
                    <button onClick={()=>ajouterMatiere(d.id)} disabled={savingId===\`new-\${d.id}\`}
                      style={{padding:"6px 12px", borderRadius:6, border:"none", background:C.green, color:"#fff", fontSize:11.5, fontWeight:700, cursor:"pointer", fontFamily:"inherit"}}>
                      {savingId===\`new-\${d.id}\` ? <Spinner size={11}/> : "+ Ajouter"}
                    </button>
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

`;

apply(
  "Patch 4 (composant DepartementsPage)",
  `function ChangePasswordPage() {`,
  DEPARTEMENTS_PAGE + `function ChangePasswordPage() {`
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${ok}/4 patchs appliqués.`);
