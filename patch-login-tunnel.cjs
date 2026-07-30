// patch-login-tunnel.cjs — Ajoute le tunnel Année → Département/Proviseur → Identifiants
// dans le LoginPage existant, sans toucher au design (hélice ADN, carte, etc.)
// Usage : node patch-login-tunnel.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-logintunnel-${Date.now()}`);

if (!fs.existsSync(FILE)) {
  console.error("❌ Fichier introuvable :", FILE);
  process.exit(1);
}

let src = fs.readFileSync(FILE, "utf8");
fs.writeFileSync(BACKUP, src, "utf8");
console.log("✓ Sauvegarde créée :", BACKUP);

const LP_MARKER = "function LoginPage({onLogin}){";
const lpStart = src.indexOf(LP_MARKER);
if (lpStart === -1) { console.error("❌ LoginPage introuvable."); process.exit(1); }

function insertAfter(marker, text, label, fromIdx = lpStart) {
  const idx = src.indexOf(marker, fromIdx);
  if (idx === -1) { console.error(`❌ ${label} — marqueur introuvable.`); process.exit(1); }
  const insertPoint = idx + marker.length;
  src = src.slice(0, insertPoint) + text + src.slice(insertPoint);
  console.log(`✓ ${label}`);
}

function insertBefore(marker, text, label, fromIdx = lpStart) {
  const idx = src.indexOf(marker, fromIdx);
  if (idx === -1) { console.error(`❌ ${label} — marqueur introuvable.`); process.exit(1); }
  src = src.slice(0, idx) + text + src.slice(idx);
  console.log(`✓ ${label}`);
}

// ── 1 — Nouveaux états ──────────────────────────────────────────────
insertAfter(
  `const [counter,setCounter] = useState({eleves:0,ens:0});`,
  `
  const [step,setStep] = useState(1); // 1=année/département, 2=identifiants
  const [annee,setAnnee] = useState("2025-2026");
  const [selDept,setSelDept] = useState("");
  const [isProviseurMode,setIsProviseurMode] = useState(false);`,
  "Patch 1 (états step/annee/selDept/isProviseurMode)"
);

// ── 2 — Validation département/rôle dans submit() ───────────────────
insertBefore(
  `// Persister l'identifiant si "Se souvenir de moi"`,
  `if (isProviseurMode && authUser.role !== "proviseur") {
      setErr("Ce compte n'est pas un compte proviseur.");
      setLoading(false);
      return;
    }
    if (!isProviseurMode && authUser.role === "proviseur") {
      setErr("Utilisez l'accès Proviseur.");
      setLoading(false);
      return;
    }
    if (!isProviseurMode && authUser.departement_id && String(authUser.departement_id) !== String(selDept)) {
      setErr("Ce compte n'appartient pas à ce département.");
      setLoading(false);
      return;
    }

    `,
  "Patch 2 (validation rôle/département dans submit())"
);

// ── 3 — Bloc étape 1 (avant le bloc Erreur/Identifiants existant) ───
const STEP1_BLOCK = `{step===1 && (
            <>
              <div style={{marginBottom:14}}>
                <label style={{display:"flex",alignItems:"center",gap:6,fontSize:13,fontWeight:700,color:"#1e293b",marginBottom:7}}>Année scolaire</label>
                <select value={annee} onChange={e=>setAnnee(e.target.value)}
                  style={{width:"100%",padding:"13px 14px",border:"1.5px solid #e2e8f0",borderRadius:14,fontSize:14,color:"#1e293b",background:"#f8fafc",fontFamily:"inherit"}}>
                  <option value="2025-2026">2025 – 2026</option>
                  <option value="2024-2025">2024 – 2025</option>
                </select>
              </div>

              <div onClick={()=>{setIsProviseurMode(!isProviseurMode); setSelDept(""); setErr("");}}
                style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 14px",borderRadius:14,border:\`1.5px solid \${isProviseurMode?"#16a34a":"#e2e8f0"}\`,background:isProviseurMode?"#f0fdf4":"#f8fafc",marginBottom:14,cursor:"pointer"}}>
                <span style={{fontSize:13,fontWeight:700,color:isProviseurMode?"#166534":"#1e293b"}}>👤 Accès Proviseur</span>
                <div style={{width:38,height:20,borderRadius:10,background:isProviseurMode?"#16a34a":"#cbd5e1",position:"relative",transition:"all .2s"}}>
                  <div style={{position:"absolute",top:2,left:isProviseurMode?20:2,width:16,height:16,borderRadius:"50%",background:"#fff",transition:"all .2s"}}/>
                </div>
              </div>

              {!isProviseurMode && (
                <div style={{marginBottom:14}}>
                  <label style={{display:"flex",alignItems:"center",gap:6,fontSize:13,fontWeight:700,color:"#1e293b",marginBottom:7}}>Département</label>
                  <select value={selDept} onChange={e=>setSelDept(e.target.value)}
                    style={{width:"100%",padding:"13px 14px",border:"1.5px solid #e2e8f0",borderRadius:14,fontSize:14,color:"#1e293b",background:"#f8fafc",fontFamily:"inherit"}}>
                    <option value="">— Sélectionner —</option>
                    {DEPARTEMENTS_LIST.map(d=><option key={d.id} value={d.id}>{d.emoji} {d.nom}</option>)}
                  </select>
                </div>
              )}

              {err && (
                <div style={{display:"flex",alignItems:"center",gap:8,background:"#fef2f2",border:"1px solid #fecaca",borderRadius:10,padding:"10px 12px",marginBottom:14}}>
                  <span style={{fontSize:14}}>⚠️</span>
                  <span style={{fontSize:12,color:"#dc2626",fontWeight:500}}>{err}</span>
                </div>
              )}

              <button onClick={()=>{
                  if(!isProviseurMode && !selDept){ setErr("Sélectionnez un département."); return; }
                  setErr(""); setStep(2);
                }}
                style={{width:"100%",padding:"16px",background:"linear-gradient(160deg,#166534 0%,#16a34a 100%)",color:"#fff",border:"none",borderRadius:16,fontSize:15,fontWeight:800,cursor:"pointer",fontFamily:"inherit",letterSpacing:".08em",textTransform:"uppercase",boxShadow:"0 8px 28px rgba(22,163,74,.45)",marginBottom:20}}>
                Continuer →
              </button>
            </>
          )}
          {step===2 && (
          <>
          <div onClick={()=>{setStep(1); setErr("");}}
            style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:"#16a34a",fontWeight:700,cursor:"pointer",marginBottom:14}}>
            ← Retour
          </div>
          `;

insertBefore(`{/* Erreur */}`, STEP1_BLOCK, "Patch 3 (bloc étape 1 + ouverture étape 2)");

// ── 4 — Fermeture du fragment étape 2 (avant le pied de carte) ──────
insertBefore(
  `{/* ── Pied de carte ── */}`,
  `</>
          )}
          `,
  "Patch 4 (fermeture étape 2)"
);

fs.writeFileSync(FILE, src, "utf8");
console.log("\n✅ Tunnel branché : Année → Proviseur/Département → Identifiants.");
