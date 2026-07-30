// patch-dashboard-proviseur-v2.cjs — Remplace DashboardProviseur par une version enrichie
// Usage : node patch-dashboard-proviseur-v2.cjs
const fs = require("fs");
const path = require("path");

const FILE = path.join("src", "App.jsx");
const BACKUP = path.join("src", `App.jsx.backup-provv2-${Date.now()}`);

if (!fs.existsSync(FILE)) {
  console.error("❌ Fichier introuvable :", FILE);
  process.exit(1);
}

let src = fs.readFileSync(FILE, "utf8");
fs.writeFileSync(BACKUP, src, "utf8");
console.log("✓ Sauvegarde créée :", BACKUP);

const startMarker = "function DashboardProviseur() {";
const endMarker = "function DashboardTeacher() {";

const startIdx = src.indexOf(startMarker);
const endIdx = src.indexOf(endMarker);

if (startIdx === -1) { console.error("❌ Marqueur de début introuvable (function DashboardProviseur)."); process.exit(1); }
if (endIdx === -1)   { console.error("❌ Marqueur de fin introuvable (function DashboardTeacher)."); process.exit(1); }
if (endIdx < startIdx) { console.error("❌ Ordre des marqueurs incohérent."); process.exit(1); }

const NEW_DASHBOARD_PROVISEUR = `function DashboardProviseur() {
  const {rawData:data,refreshData} = useApp();
  const {isMobile} = useDevice();
  const [loading,setLoading] = useState(true);
  const [refreshing,setRefreshing] = useState(false);
  const [stats,setStats] = useState({nbEns:0,nbClasses:0,nbEleves:0,tauxMoyen:0,parDept:[],niveaux:[],evolution:[],absencesParDept:[]});

  useEffect(()=>{
    if(!data)return;
    const ens=Object.values(data.users||{}).filter(u=>u.role!=="proviseur");
    const nbEleves=CLASSES_REELLES.reduce((s,c)=>s+c.effectif,0);
    const deptOf={}; ens.forEach(e=>{deptOf[e.id]=e.departement_id||1;});

    const tauxParEns=ens.map(e=>{
      let tf=0,tr=0;
      (e.classes||[]).forEach(cl=>{const k=\`\${e.id}||\${cl}\`;const f=((data.prog||{})[k]||[]).length;const code=resolveProgCode(cl);const meta=code?PROG_META[code]:null;if(meta){tf+=f;tr+=meta.lpRef;}});
      return{id:e.id,departement_id:e.departement_id||1,taux:tr>0?Math.min(100,Math.round(tf/tr*100)):0};
    });

    const parDept = DEPARTEMENTS_LIST.map(d=>{
      const teachers = tauxParEns.filter(e=>e.departement_id===d.id);
      const taux = teachers.length ? Math.round(teachers.reduce((s,e)=>s+e.taux,0)/teachers.length) : null;
      const nbAlerte = teachers.filter(e=>e.taux<50).length;
      return {...d, nbEns:teachers.length, taux, nbAlerte};
    });

    const tauxMoyen = tauxParEns.length ? Math.round(tauxParEns.reduce((s,e)=>s+e.taux,0)/tauxParEns.length) : 0;

    const groupNiveau = n => n && n.startsWith("2nde") ? "2nde" : n && n.startsWith("1ère") ? "1ère" : n && n.startsWith("Terminale") ? "Terminale" : n;
    const niveauMap = {};
    CLASSES_REELLES.forEach(c=>{
      const niv = groupNiveau(getNiveau(c.code)) || "Autre";
      niveauMap[niv] = (niveauMap[niv]||0) + c.effectif;
    });
    const ordreNiveaux = ["6ème","5ème","4ème","3ème","2nde","1ère","Terminale","Autre"];
    const niveauxColors = {"6ème":C.blue,"5ème":C.teal,"4ème":C.amber,"3ème":C.purple,"2nde":C.orange,"1ère":C.pink,"Terminale":C.green,"Autre":"#94a3b8"};
    const niveaux = ordreNiveaux.filter(n=>niveauMap[n]).map(n=>({label:n,value:niveauMap[n],color:niveauxColors[n]}));

    const evolution = ["T1","T2","T3","ANN"].map(trim=>{
      let totalFait=0, totalRef=0;
      ens.forEach(e=>{
        (e.classes||[]).forEach(cl=>{
          const code=resolveProgCode(cl);
          const meta=code?PROG_META[code]:null;
          if(!meta) return;
          const prog=(data.prog||{})[\`\${e.id}||\${cl}\`]||[];
          let lp=meta.lpRef, lf=prog.length;
          if(trim!=="ANN"){
            const range=getTrimRange(code,trim);
            if(range){
              lp=LECONS_DATA[code]?.filter(l=>l.n>=range[0]&&l.n<=range[1]).length||lp;
              lf=prog.filter(n=>n>=range[0]&&n<=range[1]).length;
            }
          }
          totalFait+=lf; totalRef+=lp;
        });
      });
      return {label:trim==="ANN"?"Annuel":trim, value: totalRef>0?Math.min(100,Math.round(totalFait/totalRef*100)):null};
    });

    const absParDept = {};
    Object.entries(data.absences||{}).forEach(([k,absents])=>{
      const ensId = k.split("||")[0];
      const dId = deptOf[ensId]||1;
      absParDept[dId] = (absParDept[dId]||0) + (absents?absents.length:0);
    });
    const absencesParDept = DEPARTEMENTS_LIST
      .map(d=>({...d, total: absParDept[d.id]||0}))
      .filter(d=>d.total>0)
      .sort((a,b)=>b.total-a.total);

    setStats({nbEns:ens.length,nbClasses:CLASSES_REELLES.length,nbEleves,tauxMoyen,parDept,niveaux,evolution,absencesParDept});
    setLoading(false);
  },[data]);

  const tauCol=t=>t===null?C.txtMuted:t>=75?C.green:t>=50?C.amber:C.red;

  const Donut = ({segments, size=130, thickness=16}) => {
    const total = segments.reduce((s,x)=>s+x.value,0) || 1;
    const r = (size-thickness)/2;
    const circ = 2*Math.PI*r;
    let acc = 0;
    return (
      <svg width={size} height={size} viewBox={\`0 0 \${size} \${size}\`}>
        <g transform={\`rotate(-90 \${size/2} \${size/2})\`}>
          {segments.map((s,i)=>{
            const frac = s.value/total;
            const dash = frac*circ;
            const el = (
              <circle key={i} cx={size/2} cy={size/2} r={r} fill="none" stroke={s.color} strokeWidth={thickness}
                strokeDasharray={\`\${dash} \${circ-dash}\`} strokeDashoffset={-acc}/>
            );
            acc += dash;
            return el;
          })}
        </g>
        <text x="50%" y="47%" textAnchor="middle" fontSize="20" fontWeight="800" fill={C.txt}>{total}</text>
        <text x="50%" y="62%" textAnchor="middle" fontSize="9" fill={C.txtMuted}>Élèves</text>
      </svg>
    );
  };

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

      <div style={{display:"grid", gridTemplateColumns: isMobile?"1fr":"1fr 1fr 1fr", gap:14}}>
        <div style={{background:C.white,borderRadius:12,border:\`1px solid \${C.border}\`,padding:16}}>
          <h3 style={{margin:"0 0 12px",fontSize:12.5,fontWeight:700,color:C.txt}}>🎓 Élèves par niveau</h3>
          {loading ? <Sk h={130} br={65}/> : stats.niveaux.length>0 ? (
            <div style={{display:"flex",alignItems:"center",gap:14}}>
              <Donut segments={stats.niveaux}/>
              <div style={{display:"flex",flexDirection:"column",gap:4,fontSize:10.5}}>
                {stats.niveaux.map(n=>(
                  <div key={n.label} style={{display:"flex",alignItems:"center",gap:5}}>
                    <span style={{width:8,height:8,borderRadius:2,background:n.color,flexShrink:0}}/>
                    <span style={{color:C.txt,fontWeight:600}}>{n.label}</span>
                    <span style={{color:C.txtMuted}}>{n.value}</span>
                  </div>
                ))}
              </div>
            </div>
          ) : <div style={{fontSize:11,color:C.txtLight,textAlign:"center",padding:"30px 0"}}>Aucune donnée</div>}
        </div>

        <div style={{background:C.white,borderRadius:12,border:\`1px solid \${C.border}\`,padding:16}}>
          <h3 style={{margin:"0 0 12px",fontSize:12.5,fontWeight:700,color:C.txt}}>📈 Évolution de la couverture</h3>
          {loading ? <Sk h={130} br={8}/> : <EvolutionChartLarge series={stats.evolution} height={130}/>}
        </div>

        <div style={{background:C.white,borderRadius:12,border:\`1px solid \${C.border}\`,padding:16}}>
          <h3 style={{margin:"0 0 4px",fontSize:12.5,fontWeight:700,color:C.txt}}>📋 Absences par département</h3>
          <p style={{margin:"0 0 12px",fontSize:10,color:C.txtMuted}}>Volume total signalé</p>
          {loading ? <Sk h={100} br={8}/> : stats.absencesParDept.length>0 ? (
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {stats.absencesParDept.slice(0,5).map(d=>(
                <div key={d.id} style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:14,flexShrink:0}}>{d.emoji}</span>
                  <span style={{fontSize:11,color:C.txt,flex:1}}>{d.nom}</span>
                  <span style={{fontSize:11,fontWeight:800,color:C.red}}>{d.total}</span>
                </div>
              ))}
            </div>
          ) : <div style={{fontSize:11,color:C.txtLight,textAlign:"center",padding:"30px 0"}}>Aucune absence enregistrée</div>}
        </div>
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

      {!loading && (stats.parDept.some(d=>d.nbAlerte>0) || stats.tauxMoyen<50) && (
        <div style={{background:C.white,borderRadius:12,border:\`1px solid \${C.border}\`,padding:18}}>
          <h3 style={{margin:"0 0 12px",fontSize:13,fontWeight:700,color:C.txt}}>⚠️ Alertes & points de vigilance</h3>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {stats.parDept.filter(d=>d.nbAlerte>0).map(d=>(
              <div key={d.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",background:"#fef2f2",border:"1px solid #fecaca",borderRadius:9}}>
                <span style={{fontSize:16}}>{d.emoji}</span>
                <span style={{fontSize:11.5,color:"#7f1d1d"}}>{d.nbAlerte} enseignant{d.nbAlerte>1?"s":""} en dessous de 50% de couverture en <b>{d.nom}</b></span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

`;

const before = src.slice(0, startIdx);
const after = src.slice(endIdx);
src = before + NEW_DASHBOARD_PROVISEUR + after;

fs.writeFileSync(FILE, src, "utf8");
console.log("✅ DashboardProviseur remplacé avec succès.");
