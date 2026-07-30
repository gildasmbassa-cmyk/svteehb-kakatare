// patch-dashboard-proviseur.cjs — Ajoute DashboardProviseur (vue tous départements)
// Usage : node patch-dashboard-proviseur.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-proviseur-${Date.now()}`);

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
  if (count === 0) {
    console.error(`❌ ${label} — motif introuvable.`);
    return;
  }
  if (count !== expected) {
    console.warn(`⚠ ${label} — ${count} occurrence(s) (attendu ${expected}), appliqué quand même.`);
  }
  src = src.split(search).join(replace);
  ok++;
  console.log(`✓ ${label}`);
}

// ── Patch A — Corriger le filtre "animatrice" en dur dans DashboardAdmin ──
apply(
  "Patch A (filtre rôle DashboardAdmin)",
  `const ens=Object.values(data.users||{}).filter(u=>u.role!=="animatrice");`,
  `const ens=Object.values(data.users||{}).filter(u=>!isAdminRole(u.role));`
);

// ── Patch B — Insérer DashboardProviseur avant DashboardTeacher ──────────
const DASHBOARD_PROVISEUR = `
// ─── Dashboard Proviseur (vue tous départements) ───────────────────
const DEPARTEMENTS_LIST = [
  {id:1,nom:"SVT",emoji:"🌿"},
  {id:2,nom:"Mathématiques",emoji:"📐"},
  {id:3,nom:"Sciences Physiques",emoji:"🧪"},
  {id:4,nom:"Lettres",emoji:"📖"},
  {id:5,nom:"Sciences Humaines",emoji:"🌍"},
  {id:6,nom:"Langues Vivantes",emoji:"🗣️"},
  {id:7,nom:"EPS",emoji:"🏃"},
  {id:8,nom:"Informatique",emoji:"💻"},
];

function DashboardProviseur() {
  const {data,refreshData} = useApp();
  const {isMobile} = useDevice();
  const [loading,setLoading] = useState(true);
  const [refreshing,setRefreshing] = useState(false);
  const [stats,setStats] = useState({nbEns:0,nbClasses:0,nbEleves:0,tauxMoyen:0,parDept:[]});

  useEffect(()=>{
    if(!data)return;
    const ens=Object.values(data.users||{}).filter(u=>!isAdminRole(u.role));
    const nbEleves=CLASSES_REELLES.reduce((s,c)=>s+c.effectif,0);

    const tauxParEns=ens.map(e=>{
      let tf=0,tr=0;
      (e.classes||[]).forEach(cl=>{const k=\`\${e.id}||\${cl}\`;const f=((data.prog||{})[k]||[]).length;const code=resolveProgCode(cl);const meta=code?PROG_META[code]:null;if(meta){tf+=f;tr+=meta.lpRef;}});
      return{id:e.id,nom:e.nom,departement_id:e.departement_id||1,taux:tr>0?Math.min(100,Math.round(tf/tr*100)):0};
    });

    const parDept = DEPARTEMENTS_LIST.map(d=>{
      const teachers = tauxParEns.filter(e=>e.departement_id===d.id);
      const taux = teachers.length ? Math.round(teachers.reduce((s,e)=>s+e.taux,0)/teachers.length) : null;
      const nbAlerte = teachers.filter(e=>e.taux<50).length;
      return {...d, nbEns:teachers.length, taux, nbAlerte};
    });

    const tauxMoyen = tauxParEns.length ? Math.round(tauxParEns.reduce((s,e)=>s+e.taux,0)/tauxParEns.length) : 0;

    setStats({nbEns:ens.length,nbClasses:CLASSES_REELLES.length,nbEleves,tauxMoyen,parDept});
    setLoading(false);
  },[data]);

  const tauCol=t=>t===null?C.txtMuted:t>=75?C.green:t>=50?C.amber:C.red;

  return(
    <div style={{padding:"20px 20px 40px",display:"flex",flexDirection:"column",gap:18}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div><h2 style={{fontSize:20,fontWeight:800,color:C.txt,margin:0}}>Bonjour, Proviseur 🏫</h2><p style={{color:C.txtMuted,margin:"3px 0 0",fontSize:12}}>{new Date().toLocaleDateString("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric"})} · Lycée de Kakatare</p></div>
        <div style={{textAlign:"right"}}><div style={{fontSize:11,color:C.txtMuted}}>2025–2026</div><div style={{fontSize:13,fontWeight:700,color:C.green}}>Vue globale · {DEPARTEMENTS_LIST.length} départements ↗</div></div>
      </div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <KpiCard label="Enseignants" value={stats.nbEns} sub="Tous départements" iconEmoji="👥" bg={C.greenPale} loading={loading} delay={0}/>
        <KpiCard label="Classes" value={stats.nbClasses} sub="Toutes séries" iconEmoji="📚" bg={C.bluePale} subColor={C.blue} loading={loading} delay={0.05}/>
        <KpiCard label="Élèves" value={stats.nbEleves} sub="Total effectifs" iconEmoji="🎓" bg={C.amberPale} subColor={C.amber} loading={loading} delay={0.1}/>
        <KpiCard label="Couverture" value={\`\${stats.tauxMoyen}%\`} sub={stats.tauxMoyen>=75?"✓ Objectif atteint":"⚠ Sous objectif"} subColor={tauCol(stats.tauxMoyen)} iconEmoji="📊" bg={C.greenPale} loading={loading} delay={0.15}/>
      </div>

      <div style={{background:C.white,borderRadius:12,border:\`1px solid \${C.border}\`,padding:18}}>
        <div style={{display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:10}}>
          <div>
            <h3 style={{margin:"0 0 2px",fontSize:13,fontWeight:700,color:C.txt}}>🏛️ Vue par département</h3>
            <p style={{margin:0,fontSize:10.5,color:C.txtMuted}}>Couverture moyenne et effectifs enseignants</p>
          </div>
          <button onClick={async()=>{ setRefreshing(true); await refreshData(); setRefreshing(false); }}
            disabled={refreshing}
            style={{display:"flex", alignItems:"center", gap:6, padding: isMobile?"7px":"6px 12px", borderRadius:8, border:\`1px solid \${C.border}\`, background:C.white, color:C.txtMuted, fontSize:11, fontWeight:700, cursor:refreshing?"not-allowed":"pointer", fontFamily:"inherit", flexShrink:0}}>
            {refreshing ? <Spinner size={11} color={C.txtMuted}/> : "🔄"} {!isMobile && "Actualiser"}
          </button>
        </div>
        <div style={{marginTop:14, display:"grid", gridTemplateColumns: isMobile?"1fr":"repeat(auto-fill, minmax(220px, 1fr))", gap:10}}>
          {loading?[1,2,3,4].map(i=><Sk key={i} h={78} br={10}/>):(
            stats.parDept.map(d=>(
              <div key={d.id} style={{padding:"14px",background: d.nbEns>0 ? "#f8fafc":"#fafafa",borderRadius:10,border:\`1px solid \${C.border}\`,borderLeft:\`3px solid \${d.nbEns>0?tauCol(d.taux):C.border}\`}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                  <span style={{fontSize:18}}>{d.emoji}</span>
                  <span style={{fontSize:12,fontWeight:700,color:C.txt,flex:1}}>{d.nom}</span>
                  {d.nbAlerte>0 && <span style={{fontSize:9,fontWeight:700,color:"#b91c1c",background:"#fef2f2",padding:"2px 6px",borderRadius:10}}>⚠ {d.nbAlerte}</span>}
                </div>
                {d.nbEns>0 ? (
                  <>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                      <span style={{fontSize:10,color:C.txtMuted}}>{d.nbEns} enseignant{d.nbEns>1?"s":""}</span>
                      <span style={{fontSize:11,fontWeight:800,color:tauCol(d.taux)}}>{d.taux}%</span>
                    </div>
                    <ProgBar value={d.taux}/>
                  </>
                ) : (
                  <div style={{fontSize:10.5,color:C.txtLight,fontStyle:"italic"}}>Aucun enseignant rattaché</div>
                )}
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  );
}

`;

apply(
  "Patch B (insertion DashboardProviseur)",
  `function DashboardTeacher() {`,
  DASHBOARD_PROVISEUR + `function DashboardTeacher() {`
);

// ── Patch C — Routing par rôle sur la page dashboard ──────────────────────
apply(
  "Patch C (routing dashboard par rôle)",
  `if(page==="dashboard")         return <W>{isAdmin?<DashboardAdmin/>:<DashboardTeacher/>}</W>`,
  `if(page==="dashboard")         return <W>{user?.role==="proviseur"?<DashboardProviseur/>:isAdmin?<DashboardAdmin/>:<DashboardTeacher/>}</W>`
);

fs.writeFileSync(FILE, src, "utf8");
console.log(`\n✅ ${ok}/3 patchs appliqués.`);
