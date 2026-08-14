import React, { useState, useEffect, useCallback, useRef, createContext, useContext, useMemo } from "react";
import { TRANSLATIONS_EN, DEPARTEMENTS_LIST } from "./lib/constants.js";
import ReactDOM from "react-dom/client";
import ELEVES_DB from "./data/eleves.json";
import EDT_REEL from "./data/edt_reel.json";
import * as XLSX from "xlsx";
import { RealtimeClient } from "@supabase/realtime-js";

// ════════════════════════════════════════════════════════════════════
// SVTEEHB — Application Pédagogique Unifiée
// Lycée de Kakatare · Maroua · Cameroun · 2025–2026
// Version : React SaaS v1.0 — Toutes pages intégrées
// ════════════════════════════════════════════════════════════════════

const SB_URL = import.meta.env.VITE_SUPABASE_URL;
const SB_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!SB_URL || !SB_KEY) {
  console.error("Variables d'environnement Supabase manquantes (VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY).");
}
const REALTIME_URL = SB_URL.replace("https://", "wss://") + "/realtime/v1";
// Tables surveillées pour la synchronisation multi-sessions (Realtime)
const REALTIME_TABLES = ["notes","absences","prog_suivi","epreuves","edt_exceptions","edt_base","utilisateurs","classes"];

// ── API Supabase (fetch léger, pas de SDK) ────────────────────────
const sb = {
  h: () => ({ apikey:SB_KEY, Authorization:`Bearer ${SB_KEY}`, "Content-Type":"application/json" }),
  async get(t, q="") {
    try { const r=await fetch(`${SB_URL}/rest/v1/${t}${q}`,{headers:sb.h()}); return r.ok?r.json():null; }
    catch { return null; }
  },
  async patchRow(t, id, fields) {
    try {
      const r = await fetch(`${SB_URL}/rest/v1/${t}?id=eq.${encodeURIComponent(id)}`,{
        method:"PATCH", headers:{...sb.h(),Prefer:"return=minimal"},
        body:JSON.stringify(fields)
      });
      if (r.ok||r.status===204) return true;
      const errBody = await r.text().catch(()=>"");
      console.error(`sb.patchRow(${t}) → ${r.status}: ${errBody}`);
      sb.lastError = errBody;
      return false;
    } catch (e) { sb.lastError = e?.message||String(e); return false; }
  },
  async upsert(t, data, conflict) {
    try {
      const r=await fetch(`${SB_URL}/rest/v1/${t}${conflict?"?on_conflict="+conflict:""}`,{
        method:"POST", headers:{...sb.h(),Prefer:"resolution=merge-duplicates,return=minimal"},
        body:JSON.stringify(Array.isArray(data)?data:[data])
      });
      if (r.ok||r.status===201||r.status===204) return true;
      const errBody = await r.text().catch(()=>"");
      console.error(`sb.upsert(${t}) → ${r.status}: ${errBody}`);
      sb.lastError = errBody;
      return false;
    } catch (e) { sb.lastError = e?.message||String(e); return false; }
  },
  async del(t, q) {
    try { const r=await fetch(`${SB_URL}/rest/v1/${t}${q}`,{method:"DELETE",headers:sb.h()}); return r.ok||r.status===204; }
    catch { return false; }
  },
  async patch(t, q, data) {
    try {
      const r=await fetch(`${SB_URL}/rest/v1/${t}${q}`,{
        method:"PATCH", headers:{...sb.h(),Prefer:"return=minimal"},
        body:JSON.stringify(data)
      });
      return r.ok||r.status===204;
    } catch { return false; }
  },
  async uploadFile(path, file) {
    try {
      const r=await fetch(`${SB_URL}/storage/v1/object/epreuves/${encodeURIComponent(path)}`,{
        method:"POST", headers:{apikey:SB_KEY,Authorization:`Bearer ${SB_KEY}`,"Content-Type":file.type||"application/octet-stream","x-upsert":"true"},
        body:file
      });
      return r.ok;
    } catch { return false; }
  },
  fileUrl: (path) => `${SB_URL}/storage/v1/object/public/epreuves/${encodeURIComponent(path)}`,
  async uploadPhoto(path, file) {
    try {
      const r=await fetch(`${SB_URL}/storage/v1/object/teacher-photos/${encodeURIComponent(path)}`,{
        method:"POST", headers:{apikey:SB_KEY,Authorization:`Bearer ${SB_KEY}`,"Content-Type":file.type||"application/octet-stream","x-upsert":"true"},
        body:file
      });
      return r.ok;
    } catch { return false; }
  },
  photoUrl: (path) => path ? `${SB_URL}/storage/v1/object/public/teacher-photos/${encodeURIComponent(path)}` : null,
  async rpc(fn, params={}) {
    try {
      const PROTECTED_RPCS = [
        "admin_delete_all_prog","admin_delete_matiere","admin_set_teacher_classes",
        "admin_add_matiere","admin_rename_matiere","admin_delete_absences_by_teacher",
        "admin_delete_all_epreuves","admin_delete_edt_slots_by_teacher","admin_delete_epreuves_by_teacher",
        "admin_delete_prog_by_classe","admin_delete_prog_by_teacher","admin_delete_teacher",
        "admin_set_edt_slots","admin_set_password","admin_upsert_teacher",
        "submit_absence","submit_note","submit_prog","submit_epreuve","submit_eleves_import","submit_vie_scolaire","submit_fiche_inspection",
      ];
      const body = PROTECTED_RPCS.includes(fn) ? {...params, p_token: window.__svtSessionToken||null} : params;
      const r = await fetch(`${SB_URL}/rest/v1/rpc/${fn}`, {
        method:"POST",
        headers:{...sb.h(), Prefer:"return=representation"},
        body:JSON.stringify(body)
      });
      if(!r.ok) {
        const errBody = await r.text().catch(()=>"");
        console.error(`sb.rpc(${fn}) → ${r.status}: ${errBody}`);
        sb.lastError = errBody;
        return null;
      }
      const d = await r.json();
      return Array.isArray(d) ? (d.length>0 ? d[0] : null) : (d||null);
    } catch (e) { sb.lastError = e?.message||String(e); return null; }
  },
};

// ── Données métier ─────────────────────────────────────────────────
const PROG_MAP = {"6ème 1":"SIX","6ème 2":"SIX","6ème 3":"SIX","6e II":"SIX","5ème 1":"CINQ","5ème 2":"CINQ","5ème 3":"CINQ","5e II":"CINQ","4ème 1 ALL/CHI":"QUATRE","4ème 2 ARA/ITA":"QUATRE","4ème 3 ESP":"QUATRE","4e II":"QUATRE","3ème 1 ALL/CHI":"TROIS","3ème 2 ARA/ITA":"TROIS","3ème 3 ESP":"TROIS","3e II":"TROIS","2nde C":"SEC_C","2nde L1 ALL/ARA/CHI":"SEC_A","2nde L2 ESP/ITA":"SEC_A","1e C":"PREM_CTI","1e Ti":"PREM_CTI","P C":"PREM_CTI","P ère Ti":"PREM_CTI","1ère S2 C/TI":"PREM_CTI","1e D":"PREM_D","1ère S1 D":"PREM_D","1e A":"PREM_A","1ère L1 ALL/ARA/CHI":"PREM_A","1ère L1 ALL/CHI":"PREM_A","1ère L2 ARA/ITA/ESP":"PREM_A","1ère L3 ESP 1":"PREM_A","1ère L3 ESP":"PREM_A","Tle D":"TERM_D","Tle S1 D":"TERM_D","Tle S2 C/TI":"TERM_CTI","Tle Ti":"TERM_CTI","Tle L1 ALL/ARA/CHI":"TERM_A","TLE Esp":"TERM_A","TLE Ita":"TERM_A","Tle L2 ESP/ITA":"TERM_A","1ère L2 ARA/ITA/ESP 2":"PREM_A","4ème ALL":"QUATRE","4ème ARB":"QUATRE","4ème CHN":"QUATRE","4ème ITA":"QUATRE","4ème ESP":"QUATRE","3ème ALL":"TROIS","3ème ARB":"TROIS","3ème CHN":"TROIS","3ème ESP":"TROIS","3ème ITA":"TROIS","2nde ALL":"SEC_A","2nde ARB":"SEC_A","2nde CHN":"SEC_A","2nde ITA":"SEC_A","2nde ESP":"SEC_A","1ère A4 ALL":"PREM_A","1ère A4 ARB":"PREM_A","1ère A4 ESP":"PREM_A","1ère CHN":"PREM_A","1ère ITA":"PREM_A","1ère C":"PREM_CTI","1ère Ti":"PREM_CTI","1ère D":"PREM_D","Tle A4 ALL":"TERM_A","Tle A4 ARB":"TERM_A","Tle A4 CHN":"TERM_A","Tle A4 ITA":"TERM_A","Tle A4 ESP":"TERM_A","Tle C":"TERM_CTI","1ère C/Ti":"PREM_CTI","Tle C/Ti":"TERM_CTI","Tle A4 ESP/ITA":"TERM_A"};
const PROG_META = {"SIX":{vh:2,hd:50,lpRef:33,tp:[2,3,4,9,10,11,14,15,16,21,22]},"CINQ":{vh:2,hd:50,lpRef:33,tp:[6,7,10,13,15,29,33]},"QUATRE":{vh:2,hd:50,lpRef:36,tp:[1,3,4,9,17,21,32]},"TROIS":{vh:2,hd:62,lpRef:45,tp:[4,5,7,9,13,25,28,32,43]},"SEC_C":{vh:2,hd:50,lpRef:29,tp:[1,12]},"SEC_A":{vh:1,hd:25,lpRef:20,tp:[]},"PREM_A":{vh:1,hd:25,lpRef:26,tp:[]},"PREM_CTI":{vh:2,hd:52,lpRef:41,tp:[]},"PREM_D":{vh:6,hd:168,lpRef:75,tp:[1,4,8,9,10,17,18,19,21,25,27,31,53]},"TERM_D":{vh:6,hd:186,lpRef:84,tp:[1,2,3,9,10,11,24,31,32,36,37,49,50,55,67,69,74,75,76,79,80,81,88,92,94,96]},"TERM_A":{vh:1,hd:31,lpRef:28,tp:[3,5,16,34,37,38]},"TERM_C":{vh:2,hd:50,lpRef:36,tp:[]},"TERM_CTI":{vh:2,hd:62,lpRef:36,tp:[1,3,6,8,15,17,18,21,24,25,30,32,37,38,39,40,41,42]}};
const PROGRAMMES_LABELS = {
  SIX:"6ème", CINQ:"5ème", QUATRE:"4ème", TROIS:"3ème",
  SEC_C:"2nde C", SEC_A:"2nde A (L1/L2)",
  PREM_A:"1ère A", PREM_CTI:"1ère C/Ti", PREM_D:"1ère D",
  TERM_D:"Tle D", TERM_A:"Tle A4", TERM_CTI:"Tle C/Ti", TERM_C:"Tle Littéraire",
};
const PROG_TRIM = {"SIX":[19,40],"QUATRE":[15,33],"TROIS":[17,33],"TERM_A":[13,30],"CINQ":[19,39],"SEC_A":[12,19],"SEC_C":[11,24],"PREM_A":[9,15],"PREM_CTI":[13,29],"PREM_D":[30,62],"TERM_D":[35,73],"TERM_C":[13,30],"TERM_CTI":[17,34]};
const TRIM_LABELS = {T1:"Trimestre 1 · Oct–Déc",T2:"Trimestre 2 · Jan–Mars",T3:"Trimestre 3 · Avr–Juin",ANN:"Année complète"};
const EP_SLOTS = [{trim:"T1",ep:"E1",label:"T1 · Épreuve 1"},{trim:"T1",ep:"E2",label:"T1 · Épreuve 2"},{trim:"T2",ep:"E1",label:"T2 · Épreuve 1"},{trim:"T2",ep:"E2",label:"T2 · Épreuve 2"},{trim:"T3",ep:"E1",label:"T3 · Épreuve 1"},{trim:"T3",ep:"E2",label:"T3 · Épreuve 2"}];
const TRIM_COLORS = {T1:"#1a6b3c",T2:"#e67e22",T3:"#8e44ad"};
const HEURES = ["07:30-08:25","08:25-09:20","09:20-10:15","10:30-11:25","11:25-12:20","12:50-13:45","13:45-14:40","14:40-15:35"];
const JOURS  = ["Lundi","Mardi","Mercredi","Jeudi","Vendredi"];
const JKEYS  = ["L","Ma","Me","J","V"];
const PLAGES_DEC = [[7.5,8.417],[8.417,9.333],[9.333,10.25],[10.5,11.417],[11.417,12.333],[12.833,13.75],[13.75,14.667],[14.667,15.583]];
// EDT_REEL importé depuis ./data/edt_reel.json (9 enseignants, base réelle)

// ── Helpers ───────────────────────────────────────────────────────
const LOGO_LYCEE_B64 = "data:image/webp;base64,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";
const LOGIN_BG_B64 = "data:image/webp;base64,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";

const SPLASH_BG_B64 = "data:image/webp;base64,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";

const ENS_COLORS = {mbassam:"#1a6b3c",boubam:"#c0392b",douniaroud:"#2980b9",hayatouh:"#16a085",aissatous:"#8e44ad",essambas:"#d35400",koffa:"#27ae60",mawiyak:"#e67e22",sadjot:"#2c3e50",sylvie:"#c8a951"};
const CLASS_COLORS = ["#1a6b3c","#2980b9","#8e44ad","#d35400","#16a085","#c0392b","#27ae60","#e67e22","#2c3e50","#0891b2","#7c3aed","#b45309"];
const getColor = id => ENS_COLORS[id]||"#1a6b3c";
const getDureeSVT = (classe) => {
  if (!classe) return 2;
  const c = classe.toLowerCase();
  const is1A4 = c.startsWith("1") && (c.includes("a4") || c.includes("ita") || c.includes("chn"));
  const isTleA4 = c.startsWith("tle") && (c.includes("a4") || c.includes("ita") || c.includes("chn") || c.includes("esp") || c.includes("arb") || c.includes("all"));
  if (is1A4 || isTleA4) return 1;
  return 2;
};
const getIni   = nom => (nom||"").replace("Mme ","").split(" ").filter(Boolean).map(w=>w[0]).join("").slice(0,2).toUpperCase();
const getNomCourt = nom => {
  const sansTitre = (nom||"").replace(/^(Mme|M\.)\s+/,"");
  return sansTitre.split(" ")[0] || sansTitre || "—";
};

// Redimensionne/compresse une image côté client avant upload (connexions lentes)
function resizeImageFile(file, maxDim = 300, quality = 0.82) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const reader = new FileReader();
    reader.onload = (e) => { img.src = e.target.result; };
    reader.onerror = reject;
    img.onload = () => {
      let { width, height } = img;
      if (width > height) { if (width > maxDim) { height = Math.round(height*maxDim/width); width = maxDim; } }
      else { if (height > maxDim) { width = Math.round(width*maxDim/height); height = maxDim; } }
      const canvas = document.createElement("canvas");
      canvas.width = width; canvas.height = height;
      canvas.getContext("2d").drawImage(img, 0, 0, width, height);
      canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("toBlob a échoué")), "image/jpeg", quality);
    };
    img.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// Avatar enseignant — photo si disponible, sinon initiales colorées (fallback automatique)
const Avatar = ({ ens, size = 30, fontSize }) => {
  const [erreur, setErreur] = useState(false);
  const fs = fontSize || Math.round(size*0.34);
  if (ens?.photo && !erreur) {
    return (
      <img src={sb.photoUrl(ens.photo)} alt={ens?.nom||""} onError={()=>setErreur(true)}
        style={{ width:size, height:size, borderRadius:"50%", objectFit:"cover", flexShrink:0, background:"#e2e8f0" }}/>
    );
  }
  return (
    <div style={{ width:size, height:size, borderRadius:"50%", background:ens?.col||"#94a3b8",
      display:"flex", alignItems:"center", justifyContent:"center", fontSize:fs, fontWeight:800, color:"#fff", flexShrink:0 }}>
      {ens?.ini||"?"}
    </div>
  );
};

// Export Excel générique — un tableau d'objets devient une feuille téléchargeable
function exportToExcel(filename, sheetName, dataRows) {
  try {
    const ws = XLSX.utils.json_to_sheet(dataRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0,31));
    XLSX.writeFile(wb, `${filename}.xlsx`);
    return true;
  } catch { return false; }
}
function normLabel(s){return String(s||"").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/\b1ere\b|\b1re\b/g,"1e").replace(/\b2nde\b|\b2nd\b/g,"2nde").replace(/\b3eme\b/g,"3e").replace(/\b4eme\b/g,"4e").replace(/\b5eme\b/g,"5e").replace(/\b6eme\b/g,"6e").replace(/\bterminale\b/g,"tle").replace(/\s+/g," ").trim();}
const PROG_MAP_NORM=(()=>{const o={};for(const k in PROG_MAP)o[normLabel(k)]=PROG_MAP[k];return o;})();
function resolveProgCode(cl){return PROG_MAP[cl]||PROG_MAP_NORM[normLabel(cl)]||null;}
function getTrimRange(code,trim){if(!code)return null;const b=PROG_TRIM[code];if(!b)return null;if(trim==="T1")return[1,b[0]];if(trim==="T2")return[b[0]+1,b[1]];if(trim==="T3")return[b[1]+1,9999];return null;}

// ── Coefficient SVTEEHB par classe (officiel MINESEC) ──────────────
function getNowInfo(){const now=new Date();const h=now.getHours()+now.getMinutes()/60;const dIdx=now.getDay();const jk=dIdx>=1&&dIdx<=5?JKEYS[dIdx-1]:null;let hi=-1;PLAGES_DEC.forEach((p,i)=>{if(h>=p[0]&&h<p[1])hi=i;});return{jk,hi,isWeekend:!jk,now};}
function progDocMeta(classe,trim,progIndex){const code=resolveProgCode(classe);const meta=code?PROG_META[code]:null;if(!code||!meta)return null;const trimRange=trim&&trim!=="ANN"?getTrimRange(code,trim):null;const allFaites=new Set();Object.entries(progIndex).forEach(([key,arr])=>{if(key.endsWith("||"+classe))(arr||[]).forEach(n=>allFaites.add(n));});let lpRef=meta.lpRef;let faites=allFaites.size;if(trimRange){const[tMin,tMax]=trimRange;lpRef=Math.round(meta.lpRef*(tMax-tMin+1)/meta.lpRef)||meta.lpRef;faites=[...allFaites].filter(n=>n>=tMin&&n<=tMax).length;}const taux=lpRef>0?Math.min(100, Math.round(faites/lpRef*100)):0;return{vh:meta.vh,hd:meta.hd,lpRef,lfRef:Math.min(lpRef,faites),taux};}
const taux2col = t => t >= 75 ? "#16a34a" : t >= 50 ? "#f59e0b" : "#ef4444";
const fmtDateFr = (v) => {
  if (!v) return "—";
  // Déjà au format français (ancien data ou saisie directe) → afficher tel quel
  if (typeof v === "string" && /^\d{2}\/\d{2}\/\d{4}/.test(v)) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? String(v) : d.toLocaleDateString("fr-FR",{day:"2-digit",month:"2-digit",year:"numeric"});
};
const taux2bg  = t => t >= 75 ? "#f0fdf4" : t >= 50 ? "#fffbeb" : "#fef2f2";
const classColorCache={};let colorIdx=0;
function getClassColor(cl){if(!cl)return"#94a3b8";if(!classColorCache[cl])classColorCache[cl]=CLASS_COLORS[colorIdx++%CLASS_COLORS.length];return classColorCache[cl];}

// ── Comptes démo ──────────────────────────────────────────────────
// Classes issues de ENSEIGNANTS1.pdf (EDT Lantiv officiel Kakatare)
// Classes officielles (source : CLASSES.pdf Lantiv Kakatare)
// Clés = noms exacts du PDF utilisés pour PROG_MAP ; display = nom court feuille de calcul

// ════════════════════════════════════════════════════════════════
// CHARGEMENT LAZY DES DONNÉES STATIQUES (JSON servis par Vercel)
// ════════════════════════════════════════════════════════════════
let _elevesLoaded = false;

function loadElevesDB() {
  _elevesLoaded = true;
  ALL_CLASSES = getAllClasses();
  TOTAL_ELEVES = getTotalEleves();
  TOTAL_FILLES = getTotalFilles();
  TOTAL_GARCONS = getTotalGarcons();
}

// LECONS_DATA est désormais dans le bundle — chargement direct
async function loadStaticData() {
  await loadElevesDB(); // Seul ELEVES_DB est chargé dynamiquement
}

const ENS_CLASSES_REF = {
  // Libellés alignés sur EDT_REEL harmonisé + Relevé de notes
  mbassam:    ["5ème 2","1ère C","1ère Ti","Tle A4 ESP","Tle A4 ITA"],  // 1ère C/Ti éclatée + Tle ESP/ITA séparées  // Tle ESP/ITA (classe mixte)
  boubam:     ["2nde C","3ème ARB","4ème ARB"],
  douniaroud: ["6ème 1","Tle A4 ALL","Tle C/Ti"],
  hayatouh:   ["1ère D"],
  aissatous:  ["Tle D"],
  essambas:   ["1ère A4 ARB","3ème ALL","4ème ALL"],
  koffa:      ["1ère A4 ESP","5ème 3","6ème 3"],
  mawiyak:    ["2nde ALL","2nde ESP","4ème ESP","6ème 2"],
  sadjot:     ["1ère A4 ALL","3ème ESP","5ème 1"],
  sylvie:     [],
};

// Noms d'affichage courts (feuille de calcul MINESEC)
const CLASS_DISPLAY = {
  // Source : Relevé de notes officiel — Lycée de Kakatare 2025–2026
  // 38 classes séparées + alias EDT
  "6ème 1":"6ème I","6ème 2":"6ème II","6ème 3":"6ème III",
  "5ème 1":"5ème I","5ème 2":"5ème II","5ème 3":"5ème III",
  "4ème ALL":"4ème Allemand","4ème ARB":"4ème Arabe",
  "4ème CHN":"4ème Chinois","4ème ITA":"4ème Italien","4ème ESP":"4ème Espagnol",
  "3ème ALL":"3ème Allemand","3ème ARB":"3ème Arabe",
  "3ème CHN":"3ème Chinois","3ème ESP":"3ème Espagnol","3ème ITA":"3ème Italien",
  "2nde ALL":"2nde Allemand","2nde ARB":"2nde Arabe","2nde CHN":"2nde Chinois",
  "2nde ITA":"2nde Italien","2nde ESP":"2nde Espagnol","2nde C":"2nde C",
  "1ère A4 ALL":"1ère A4 Allemand","1ère A4 ARB":"1ère A4 Arabe",
  "1ère A4 ESP":"1ère A4 Espagnol","1ère CHN":"1ère Chinois","1ère ITA":"1ère Italien",
  "1ère C":"1ère C","1ère D":"1ère D","1ère Ti":"1ère Ti",
  "1ère C/Ti":"1ère C & Ti",
  "Tle A4 ALL":"Tle A4 Allemand","Tle A4 ARB":"Tle A4 Arabe","Tle A4 CHN":"Tle A4 Chinois",
  "Tle A4 ITA":"Tle A4 Italien","Tle A4 ESP":"Tle A4 Espagnol",
  "Tle C":"Tle C","Tle D":"Tle D","Tle Ti":"Tle Ti",
  "Tle C/Ti":"Tle C & Ti",
  "TLE Esp":"Tle A4 Espagnol","TLE Ita":"Tle A4 Italien",
"Tle A4 ESP/ITA":"Tle A4 Espagnol / Italien",};

const DEMO_ACCOUNTS = [
  // Source : ENSEIGNANTS1.pdf + CLASSES.pdf Lantiv — Lycée de Kakatare 2025–2026
  // Mots de passe à personnaliser avant déploiement définitif
  {id:"sylvie",nom:"AÏSSATOU SYLVIE",   role:"animatrice", sub:"Animatrice Pédagogique · PCEG", classes:[]},
  {id:"mbassam", nom:"MBASSA André Gildas",role:"enseignant", sub:"5ème 2 · 1ère C/Ti · Tle A4 ESP", classes:["5ème 2","1ère C","1ère Ti","Tle A4 ESP","Tle A4 ITA"]},
  {id:"boubam", nom:"BOUBA M",            role:"enseignant", sub:"2nde C · 3ème ARB · 4ème ARB",   classes:["2nde C","3ème ARB","4ème ARB"]},
  {id:"douniaroud",nom:"DOUNIAROU D",        role:"enseignant", sub:"6ème 1 · Tle A4 ALL · Tle C/Ti", classes:["6ème 1","Tle A4 ALL","Tle C/Ti"]},
  {id:"hayatouh", nom:"HAYATOU H",          role:"enseignant", sub:"1ère D",                         classes:["1ère D"]},
  {id:"aissatous", nom:"Mme AÏSSATOU S",     role:"enseignant", sub:"Tle D",                          classes:["Tle D"]},
  {id:"essambas", nom:"Mme ESSAMBA S",       role:"enseignant", sub:"1ère A4 ARB · 3ème ALL · 4ème ALL", classes:["1ère A4 ARB","3ème ALL","4ème ALL"]},
  {id:"koffa", nom:"Mme KOFFA",           role:"enseignant", sub:"1ère A4 ESP · 5ème 3 · 6ème 3", classes:["1ère A4 ESP","5ème 3","6ème 3"]},
  {id:"mawiyak", nom:"Mme MAWIYA K",        role:"enseignant", sub:"2nde ALL · 2nde ESP · 4ème ESP · 6ème 2", classes:["2nde ALL","2nde ESP","4ème ESP","6ème 2"]},
  {id:"sadjot", nom:"Mme SADJO T",          role:"enseignant", sub:"1ère A4 ALL · 3ème ESP · 5ème 1", classes:["1ère A4 ALL","3ème ESP","5ème 1"]},
];


// ── Chargement données Supabase ───────────────────────────────────
// Fusionne les ajouts/modifications d'élèves persistés (eleves_import) dans ELEVES_DB
async function syncElevesImport() {
  try {
    const rows = await sb.get("eleves_import", "?select=classe,donnees");
    (rows||[]).forEach(r => {
      let arr = r.donnees;
      if (typeof arr === "string") { try { arr = JSON.parse(arr); } catch { arr = null; } }
      if (Array.isArray(arr) && arr.length > 0) {
        ELEVES_DB[r.classe] = arr;
      }
    });
  } catch { /* échec silencieux — on garde les données par défaut */ }
}

async function loadAllData(departementId = null) {
  await loadTrimestres();
  await loadCoefficients();
  await syncElevesImport();
  const [classes, users, prog, epreuves, exceptions, notes, absences, edtBase] = await Promise.all([
    sb.get("classes", departementId ? `?select=code,effectif,enseignant,departement_id&departement_id=eq.${departementId}&order=code` : "?select=code,effectif,enseignant,departement_id&order=code"),
    sb.get("utilisateurs", departementId ? `?select=id,nom,role,classes,photo,departement_id&departement_id=eq.${departementId}` : "?select=id,nom,role,classes,photo,departement_id"),
    sb.get("prog_suivi", departementId ? `?select=ens_id,classe,faites&departement_id=eq.${departementId}` : "?select=ens_id,classe,faites"),
    sb.get("epreuves",""),
    sb.get("edt_exceptions","?select=ens_id,slot,lbl"),
    sb.get("notes","?select=classe,evaluation,eleve_id,note"),
    sb.get("absences", departementId ? `?select=ens_id,classe,seance,absents&departement_id=eq.${departementId}` : "?select=ens_id,classe,seance,absents"),
  ]);
  const progIndex={};
  (prog||[]).forEach(r=>{progIndex[`${r.ens_id}||${r.classe}`]=Array.isArray(r.faites)?r.faites:[];});
  const notesIndex={};
  (notes||[]).forEach(r=>{
    const k=`${r.classe}||${r.evaluation}`;
    if(!notesIndex[k]) notesIndex[k]={};
    if(r.note!==null) notesIndex[k][r.eleve_id]=r.note;
  });
  const absIndex={};
  (absences||[]).forEach(r=>{absIndex[`${r.ens_id}||${r.classe}||${r.seance}`]=Array.isArray(r.absents)?r.absents:[];});
  const usersMap={};
  (users||[]).forEach(u=>{usersMap[u.id]={...u,classes:u.classes||[],col:getColor(u.id),ini:getIni(u.nom)};});
  const excMap={};
  (exceptions||[]).forEach(e=>{if(!excMap[e.ens_id])excMap[e.ens_id]={};excMap[e.ens_id][e.slot]=e.lbl;});
  const eps=(epreuves||[]).map(e=>({...e,stockPath:e.stockpath||e.stockPath||"",commentaires:Array.isArray(e.commentaires)?e.commentaires:[]}))
    .sort((a,b)=>{const da=a.soumis||a.created_at||a.id||0, db=b.soumis||b.created_at||b.id||0; return db>da?1:-1;});
  const edtBaseMap={};
  (edtBase||[]).forEach(r=>{if(!edtBaseMap[r.ens_id])edtBaseMap[r.ens_id]={};edtBaseMap[r.ens_id][r.slot]=r.lbl||null;});
  return{classes:classes||[],users:usersMap,prog:progIndex,epreuves:eps,exceptions:excMap,notes:notesIndex,absences:absIndex,edtBase:edtBaseMap};
}

// ── Palette ───────────────────────────────────────────────────────
const C = {
  sidebar:"var(--c-sidebar)",sidebarBorder:"var(--c-sidebarBorder)",sidebarActive:"var(--c-sidebarActive)",sidebarActiveText:"var(--c-sidebarActiveText)",sidebarText:"var(--c-sidebarText)",sidebarHover:"var(--c-sidebarHover)",
  green:"var(--c-green)",greenLight:"var(--c-greenLight)",greenDark:"var(--c-greenDark)",greenPale:"var(--c-greenPale)",greenBorder:"var(--c-greenBorder)",
  gold:"var(--c-gold)",goldPale:"var(--c-goldPale)",goldBorder:"var(--c-goldBorder)",
  white:"var(--c-white)",bg:"var(--c-bg)",border:"var(--c-border)",
  txt:"var(--c-txt)",txtMuted:"var(--c-txtMuted)",txtLight:"var(--c-txtLight)",
  blue:"var(--c-blue)",bluePale:"var(--c-bluePale)",
  orange:"var(--c-orange)",orangePale:"var(--c-orangePale)",orangeBorder:"var(--c-orangeBorder)",
  red:"var(--c-red)",redPale:"var(--c-redPale)",redBorder:"var(--c-redBorder)",
  amber:"var(--c-amber)",amberPale:"var(--c-amberPale)",
  purple:"var(--c-purple)",purplePale:"var(--c-purplePale)",
  teal:"var(--c-teal)",tealPale:"var(--c-tealPale)",
  pink:"var(--c-pink)",pinkPale:"var(--c-pinkPale)",
  // Tokens avec alpha pré-calculé (remplacent les anciennes concaténations C.xxx+"NN" incompatibles avec var())
  greenPaleA30:"var(--c-greenPaleA30)",greenPaleA40:"var(--c-greenPaleA40)",greenPaleA60:"var(--c-greenPaleA60)",
  blueA40:"var(--c-blueA40)",pinkA40:"var(--c-pinkA40)",redA30:"var(--c-redA30)",redA40:"var(--c-redA40)",
};

// ── Contexte ──────────────────────────────────────────────────────
const AppCtx = createContext(null);
const useApp = () => useContext(AppCtx);

const ADMIN_ROLES = ["animatrice", "animateur", "proviseur"];
const isAdminRole = (role) => ADMIN_ROLES.includes(role);
// ══════════════════════════════════════════════════════════════════════
// HOOK : Détecteur automatique mobile / tablette / desktop
// Réactif au resize — met à jour en temps réel
// ══════════════════════════════════════════════════════════════════════
function useDevice() {
  const getState = () => {
    if (typeof window === "undefined") return {device:"desktop", orientation:"landscape"};
    const w = window.innerWidth;
    const h = window.innerHeight;
    const orientation = w > h ? "landscape" : "portrait";
    const device = w < 640 ? "mobile"
                 : w < 1024 ? "tablet"
                 : "desktop";
    return {device, orientation};
  };

  const [state, setState] = useState(getState);

  useEffect(() => {
    const update = () => setState(getState());
    const orientUpdate = () => setTimeout(update, 120);
    window.addEventListener("resize", update);
    window.addEventListener("orientationchange", orientUpdate);
    return () => {
      window.removeEventListener("resize", update);
      window.removeEventListener("orientationchange", orientUpdate);
    };
  }, []);

  const {device, orientation} = state;
  return {
    device,
    orientation,
    isMobile:    device === "mobile",
    isTablet:    device === "tablet",
    isDesktop:   device === "desktop",
    isSmall:     device !== "desktop",
    isPortrait:  orientation === "portrait",
    isLandscape: orientation === "landscape",
    // Cas spécifiques
    mobileLandscape: device === "mobile"  && orientation === "landscape",
    mobilePortrait:  device === "mobile"  && orientation === "portrait",
    tabletLandscape: device === "tablet"  && orientation === "landscape",
    tabletPortrait:  device === "tablet"  && orientation === "portrait",
  };
}

// ════════════════════════════════════════════════════════════════
// HOOK MODE SOMBRE
// ════════════════════════════════════════════════════════════════
function useDarkMode() {
  const [dark, setDark] = useState(()=>{
    try { return localStorage.getItem("svteehb-dark")==="1"; } catch{return false;}
  });
  useEffect(()=>{
    document.body.classList.toggle("dark-mode", dark);
    try { localStorage.setItem("svteehb-dark", dark?"1":"0"); } catch{}
  }, [dark]);
  return [dark, setDark];
}



// ── Composants UI partagés ────────────────────────────────────────
// ─── Logo officiel SVTEEHB — microscope + feuille + globe + bandeau SVT ──
const SVTLogo = ({size=36}) => (
  <svg viewBox="0 0 160 160" width={size} height={size} style={{flexShrink:0}}>
    {/* Fond cercle */}
    <circle cx="80" cy="80" r="77" fill="#dcfce7" stroke="#166534" strokeWidth="3.5"/>
    <circle cx="80" cy="80" r="71" fill="#f0fdf4" stroke="#4ade80" strokeWidth="1" strokeDasharray="3 2"/>

    {/* Microscope (gauche) */}
    <g transform="translate(18,28)" fill="#166534">
      <rect x="14" y="0" width="10" height="6" rx="2"/>
      <rect x="15" y="5" width="8" height="20" rx="1"/>
      <rect x="10" y="22" width="18" height="5" rx="2"/>
      <rect x="15" y="26" width="8" height="30" rx="2"/>
      <rect x="8" y="54" width="22" height="5" rx="2.5"/>
      <ellipse cx="19" cy="26" rx="4" ry="3" fill="#4ade80"/>
      <circle cx="18" cy="2" r="3" fill="none" stroke="#166534" strokeWidth="1.5"/>
    </g>

    {/* Plante / feuille centrale */}
    <g transform="translate(72,18)">
      <path d="M8,0 Q24,-4 28,14 Q16,24 8,20 Q0,16 8,0Z" fill="#22c55e"/>
      <path d="M8,0 Q-8,-4 -12,14 Q0,24 8,20 Q16,16 8,0Z" fill="#4ade80"/>
      <line x1="8" y1="20" x2="8" y2="38" stroke="#166534" strokeWidth="2" strokeLinecap="round"/>
      <line x1="8" y1="28" x2="14" y2="24" stroke="#166534" strokeWidth="1" strokeLinecap="round"/>
      <line x1="8" y1="32" x2="2" y2="28" stroke="#166534" strokeWidth="1" strokeLinecap="round"/>
    </g>

    {/* Globe terrestre (droite) */}
    <g transform="translate(100,38)">
      <circle cx="18" cy="18" r="18" fill="none" stroke="#166534" strokeWidth="2"/>
      <ellipse cx="18" cy="18" rx="9" ry="18" fill="none" stroke="#166534" strokeWidth="1.5"/>
      <path d="M0,18 Q9,12 18,18 Q27,24 36,18" fill="none" stroke="#166534" strokeWidth="1.5"/>
      <path d="M2,10 Q12,6 26,10" fill="none" stroke="#166534" strokeWidth="1"/>
      <path d="M2,26 Q12,30 26,26" fill="none" stroke="#166534" strokeWidth="1"/>
      <path d="M8,10 Q12,8 16,12 Q14,16 10,15Z" fill="rgba(22,163,74,.4)"/>
      <path d="M20,14 Q24,12 28,16 Q26,20 22,19Z" fill="rgba(22,163,74,.4)"/>
      <path d="M10,22 Q14,20 18,24 Q16,28 12,27Z" fill="rgba(22,163,74,.4)"/>
    </g>

    {/* Banderole SVT */}
    <rect x="8" y="112" width="144" height="34" rx="17" fill="#166534"/>
    <text x="80" y="134" textAnchor="middle" fill="white" fontSize="18" fontWeight="900" fontFamily="'DM Sans',system-ui,sans-serif" letterSpacing="3">SVT</text>
  </svg>
);
const LogoSVG = ({size=36}) => (<SVTLogo size={size}/>);
const Spinner = ({size=16,color="#fff"}) => (<span style={{width:size,height:size,border:`2px solid ${color}40`,borderTopColor:color,borderRadius:"50%",display:"inline-block",animation:"spin .7s linear infinite",flexShrink:0}}/>);
const Sk = ({h=14,w="100%",br=6}) => (<div style={{height:h,width:w,borderRadius:br,background:"linear-gradient(90deg,#e2e8f0 25%,#f1f5f9 50%,#e2e8f0 75%)",backgroundSize:"200% 100%",animation:"shimmer 1.4s infinite"}}/>);
const Pill = ({ch,color=C.green,bg}) => (<span style={{display:"inline-flex",alignItems:"center",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700,background:bg||color+"18",color,whiteSpace:"nowrap"}}>{ch}</span>);
const ProgBar = ({value,h=5}) => {const c=value>=75?C.green:value>=50?C.amber:C.red;return(<div style={{height:h,background:"#e2e8f0",borderRadius:h,overflow:"hidden"}}><div style={{width:`${Math.min(100,value||0)}%`,height:"100%",background:c,borderRadius:h,transition:"width .5s"}}/></div>);};
const Toast = ({msg,ok}) => (<div style={{position:"fixed",bottom:24,right:24,zIndex:9999,display:"flex",alignItems:"center",gap:10,padding:"12px 18px",borderRadius:12,background:ok===false?C.red:C.greenDark,color:"#fff",fontSize:13,fontWeight:600,boxShadow:"0 8px 24px rgba(0,0,0,.25)",animation:"fadeUp .3s ease",maxWidth:360}}>{msg}</div>);
const STATUT_CFG = {attente:{label:"En attente",emoji:"⏳",color:C.orange,bg:C.orangePale,border:C.orangeBorder},validee:{label:"Validée",emoji:"✅",color:C.green,bg:C.greenPale,border:C.greenBorder},rejetee:{label:"Rejetée",emoji:"❌",color:C.red,bg:C.redPale,border:C.redBorder},vide:{label:"Non soumise",emoji:"○",color:C.txtLight,bg:"#f8fafc",border:C.border}};

// ── CSS Global ────────────────────────────────────────────────────
const GLOBAL_CSS = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=Playfair+Display:wght@700;800&display=swap');
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}

  /* ── Palette claire (par défaut) ──────────────────────────────── */
  :root{
    --c-sidebar:#0B3D20;--c-sidebarBorder:rgba(212,175,55,0.15);--c-sidebarActive:rgba(212,175,55,0.14);--c-sidebarActiveText:#D4AF37;--c-sidebarText:rgba(255,255,255,0.60);--c-sidebarHover:rgba(255,255,255,0.06);
    --c-green:#0B4D2C;--c-greenLight:#166534;--c-greenDark:#083D22;--c-greenPale:#f0f7f2;--c-greenBorder:#a7d4b5;
    --c-gold:#D4AF37;--c-goldPale:#fdf8e8;--c-goldBorder:rgba(212,175,55,.4);
    --c-white:#ffffff;--c-bg:#f8fafc;--c-border:#e2e8f0;
    --c-txt:#0f172a;--c-txtMuted:#64748b;--c-txtLight:#94a3b8;
    --c-blue:#3b82f6;--c-bluePale:#eff6ff;
    --c-orange:#f97316;--c-orangePale:#fff7ed;--c-orangeBorder:#fed7aa;
    --c-red:#ef4444;--c-redPale:#fef2f2;--c-redBorder:#fecaca;
    --c-amber:#f59e0b;--c-amberPale:#fffbeb;
    --c-purple:#8b5cf6;--c-purplePale:#f5f3ff;
    --c-teal:#14b8a6;--c-tealPale:#f0fdfa;
    --c-pink:#ec4899;--c-pinkPale:#fdf2f8;
    --c-greenPaleA30:rgba(240,253,244,.3);--c-greenPaleA40:rgba(240,253,244,.4);--c-greenPaleA60:rgba(240,253,244,.6);
    --c-blueA40:rgba(59,130,246,.4);--c-pinkA40:rgba(236,72,153,.4);--c-redA30:rgba(239,68,68,.3);--c-redA40:rgba(239,68,68,.4);
  }

  /* ── Palette sombre — activée via body.dark-mode ─────────────── */
  body.dark-mode{
    --c-white:#1a2420;--c-bg:#0f1411;--c-border:#2d3a33;
    --c-txt:#e8f0ea;--c-txtMuted:#94a89e;--c-txtLight:#6b7d73;
    --c-green:#22c55e;--c-greenLight:#4ade80;--c-greenPale:#0f2a1c;--c-greenBorder:#1f4a35;
    --c-gold:#d4b860;--c-goldPale:#2a2418;--c-goldBorder:rgba(212,184,96,.3);
    --c-blue:#60a5fa;--c-bluePale:#16202e;
    --c-orange:#fb923c;--c-orangePale:#2b1f12;--c-orangeBorder:#4a3320;
    --c-red:#f87171;--c-redPale:#2a1515;--c-redBorder:#4a2424;
    --c-amber:#fbbf24;--c-amberPale:#2a2310;
    --c-purple:#a78bfa;--c-purplePale:#221e33;
    --c-teal:#2dd4bf;--c-tealPale:#0f2a26;
    --c-pink:#f472b6;--c-pinkPale:#2a1622;
    --c-greenPaleA30:rgba(15,42,28,.3);--c-greenPaleA40:rgba(15,42,28,.4);--c-greenPaleA60:rgba(15,42,28,.6);
    --c-blueA40:rgba(96,165,250,.4);--c-pinkA40:rgba(244,114,182,.4);--c-redA30:rgba(248,113,113,.3);--c-redA40:rgba(248,113,113,.4);
  }

  body{font-family:'DM Sans',system-ui,sans-serif;background:var(--c-bg);transition:background .25s ease;}
  @keyframes spin{to{transform:rotate(360deg)}}
  @keyframes shimmer{0%{background-position:200% 0}100%{background-position:-200% 0}}
  @keyframes fadeUp{from{opacity:0;transform:translateY(8px);}to{opacity:1;transform:none;}}
  @keyframes logoIn{from{opacity:0;transform:scale(.75);}to{opacity:1;transform:none;}}
  @keyframes pulse{0%,100%{opacity:.5;transform:scale(1);}50%{opacity:1;transform:scale(1.4);}}
  input,button,select,textarea{font-family:inherit;}
  ::-webkit-scrollbar{width:4px;height:4px;}::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:4px;}
  ::-webkit-scrollbar-track{background:transparent;}
  input::placeholder{color:#94a3b8;}
`;


// useState, useMemo: importés en tête de fichier

// ═══════════════════════════════════════════════════
// BASE NOMINALE COMPLÈTE — 1163 élèves / 38 classes
// Source : Fiche_Suivi_Annuel_Kakatare_2025_2026.xlsx
// ═══════════════════════════════════════════════════
// ELEVES_DB importé depuis ./data/eleves.json (38 classes, 1166 élèves) // 38 classes, 1166 élèves — données inline

// Liste fiable classes+effectifs, dérivée d'ELEVES_DB — à utiliser PARTOUT
// au lieu de la table Supabase "classes" (peu fiable : vide ou codes désynchronisés)
const CLASSES_REELLES = Object.keys(ELEVES_DB).sort().map(code => ({
  code, effectif: (ELEVES_DB[code]||[]).length,
}));


// ── Référentiel programmes par niveau (avec coefficients) ──────────
const PROGRAMME_SVTEEHB = {
  "6ème":        { lp_t1:19, lp_t2:21, lp_t3:12, tp:12, coeff:2 },
  "5ème":        { lp_t1:19, lp_t2:20, lp_t3:5,  tp:9,  coeff:2 },
  "4ème":        { lp_t1:15, lp_t2:18, lp_t3:12, tp:8,  coeff:2 },
  "3ème":        { lp_t1:17, lp_t2:16, lp_t3:11, tp:10, coeff:2 },
  "2nde C":      { lp_t1:11, lp_t2:8,  lp_t3:11, tp:4,  coeff:2 },
  "2nde A":      { lp_t1:9,  lp_t2:6,  lp_t3:5,  tp:0,  coeff:1 },
  "1ère D":      { lp_t1:30, lp_t2:32, lp_t3:16, tp:21, coeff:6 },
  "1ère C/Ti":   { lp_t1:13, lp_t2:16, lp_t3:12, tp:8,  coeff:2 },
  "1ère A":      { lp_t1:9,  lp_t2:6,  lp_t3:5,  tp:0,  coeff:1 },
  "Terminale D": { lp_t1:35, lp_t2:38, lp_t3:24, tp:28, coeff:6 },
  "Terminale C": { lp_t1:17, lp_t2:17, lp_t3:9,  tp:20, coeff:2 },
  "Terminale A": { lp_t1:13, lp_t2:12, lp_t3:8,  tp:4,  coeff:1 },
};

function getNiveau(code) {
  const c = code.toLowerCase();
  if (c.startsWith("6")) return "6ème";
  if (c.startsWith("5")) return "5ème";
  if (c.startsWith("4")) return "4ème";
  if (c.startsWith("3")) return "3ème";
  if (c.includes("2nde c")) return "2nde C";
  if (c.includes("2nde")) return "2nde A";
  if (c.includes("1ère") && (c.includes(" d") || c.includes("s1"))) return "1ère D";
  if (c.includes("1ère") && (c.includes("c") || c.includes("ti"))) return "1ère C/Ti";
  if (c.includes("1ère")) return "1ère A";
  if (c.includes("tle") && (c.includes(" d") || c.includes("s1"))) return "Terminale D";
  if (c.includes("tle") && (c.includes("c"))) return "Terminale C";
  if (c.includes("tle") || c.includes("terminale")) return "Terminale A";
  return null;
}

// ── Palette ────────────────────────────────────────────────────────
// const C: défini et enrichi globalement (pink, amber ajoutés)

// Niveaux pour le groupement de navigation
const NIVEAUX_ORDER = ["6ème","5ème","4ème","3ème","2nde","1ère","Terminale"];
// getNiveauGroupe supprimé — utiliser niveauGroupe() (ligne 110)

// Sk: défini globalement

// ── Stats globales ─────────────────────────────────────────────────
// Stats globales — calculées dynamiquement après chargement ELEVES_DB
function getAllClasses() { return Object.keys(ELEVES_DB); }
function getTotalEleves() { return Object.values(ELEVES_DB).reduce((s,cl)=>s+cl.length,0); }
function getTotalFilles() { return Object.values(ELEVES_DB).reduce((s,cl)=>s+cl.filter(e=>e.g==="F").length,0); }
function getTotalGarcons() { return getTotalEleves() - getTotalFilles(); }
// Aliases statiques (ne pas utiliser dans les composants React — utiliser les fonctions ci-dessus)
let ALL_CLASSES = [], TOTAL_ELEVES = 0, TOTAL_FILLES = 0, TOTAL_GARCONS = 0;

// ── Impression liste de classe ──────────────────────────────────────
let printGuard = false;
function imprimerListeClasse(code, eleves) {
  if (printGuard) return;
  printGuard = true;
  const date = new Date().toLocaleDateString("fr-FR",{day:"2-digit",month:"long",year:"numeric"});
  const filles = eleves.filter(e=>e.g==="F").length;
  const garcons = eleves.filter(e=>e.g==="M").length;
  const rows = eleves.map((e,i) => `
    <tr style="border-bottom:1px solid #e5e7eb;${i%2===0?"background:#f9fafb":""}">
      <td style="padding:5px 10px;text-align:center;font-weight:700;color:#6b7280">${i+1}</td>
      <td style="padding:5px 10px;font-weight:600">${e.nom}</td>
      <td style="padding:5px 10px;text-align:center">
        <span style="padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;
          background:${e.g==="F"?"#fdf2f8":"#eff6ff"};
          color:${e.g==="F"?"#db2777":"#2563eb"};">${e.g}</span>
      </td>
      <td style="padding:5px 10px"></td>
      <td style="padding:5px 10px"></td>
    </tr>`).join("");

  const html = `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<style>
  body{font-family:Arial,sans-serif;font-size:10px;margin:20px;color:#111;}
  table{width:100%;border-collapse:collapse;}
  thead tr{background:#0f1f14;color:#fff;}
  th{padding:7px 10px;text-align:left;font-size:9px;letter-spacing:.05em;}
  @media print{body{margin:10px}}
</style></head><body>
  <div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:10px;border-bottom:2px solid #16a34a;padding-bottom:8px">
    <div style="text-align:center;font-size:8px;line-height:1.5;width:40%">
      <strong>REPUBLIQUE DU CAMEROUN</strong><br><em>Paix – Travail – Patrie</em><br>
      MINISTERE DES ENSEIGNEMENTS SECONDAIRES<br><strong>LYCÉE DE KAKATARE — MAROUA</strong>
    </div>
    <div style="text-align:center;width:20%">
      <div style="width:50px;height:50px;border-radius:50%;background:linear-gradient(135deg,#34b06c,#0c3d24);margin:0 auto;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:900;font-size:16px">SV</div>
    </div>
    <div style="text-align:center;font-size:8px;line-height:1.5;width:40%">
      <strong>REPUBLIC OF CAMEROON</strong><br><em>Peace – Work – Fatherland</em><br>
      MINISTRY OF SECONDARY EDUCATION<br><strong>LYCÉE DE KAKATARE — MAROUA</strong>
    </div>
  </div>
  <h2 style="text-align:center;font-size:12px;font-weight:900;text-transform:uppercase;border:2px solid #222;padding:6px;margin:8px 0;letter-spacing:.04em">
    LISTE OFFICIELLE DES ÉLÈVES — CLASSE : ${code}
  </h2>
  <div style="font-size:9px;margin-bottom:8px;display:flex;gap:20px">
    <span>Année scolaire : <strong>2025–2026</strong></span>
    <span>Effectif total : <strong>${eleves.length}</strong></span>
    <span>Filles : <strong>${filles}</strong></span>
    <span>Garçons : <strong>${garcons}</strong></span>
    <span>Date : ${date}</span>
  </div>
  <table>
    <thead><tr>
      <th style="width:40px;text-align:center">N°</th>
      <th>NOM ET PRÉNOM(S)</th>
      <th style="width:50px;text-align:center">GENRE</th>
      <th style="width:100px">SIGNATURE</th>
      <th style="width:120px">OBSERVATIONS</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table>
  <div style="margin-top:20px;display:grid;grid-template-columns:1fr 1fr;gap:30px;font-size:9px">
    <div><div style="font-weight:bold;margin-bottom:10px">Le Professeur Principal</div><div style="border-top:1px solid #999;margin-top:30px;padding-top:4px">Signature</div></div>
    <div><div style="font-weight:bold;margin-bottom:10px">Le Proviseur</div><div style="border-top:1px solid #999;margin-top:30px;padding-top:4px">Signature &amp; Cachet</div></div>
  </div>
  <script>window.onload=()=>window.print();</script>
</body></html>`;

  const iframe = document.createElement("iframe");
  iframe.style.cssText = "position:fixed;top:-9999px;left:-9999px;width:1px;height:1px;";
  document.body.appendChild(iframe);
  const doc = iframe.contentDocument || iframe.contentWindow.document;
  doc.open(); doc.write(html); doc.close();
  setTimeout(()=>{ iframe.contentWindow.focus(); iframe.contentWindow.print();
    setTimeout(()=>{ document.body.removeChild(iframe); printGuard=false; }, 2000);
  }, 400);
}

// ══════════════════════════════════════════════════════════════════
// PAGE PRINCIPALE ÉLÈVES
// ══════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════
// MES CLASSES (enseignant) — liste élèves + présences + notes
// ═══════════════════════════════════════════════════════════════

function genererConvocation(eleve, classe, absH, retards, sanctions) {
  const date = new Date().toLocaleDateString("fr-FR",{day:"2-digit",month:"long",year:"numeric"});
  const motif = absH>=15?"Blame d'assiduite ("+absH+"h)":absH>=6?"Avertissement ("+absH+"h)":retards>=3?"Retards repetes ("+retards+")":"Comportement";
  const html = '<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>Convocation</title>'
    + '<style>body{font-family:Georgia,serif;max-width:700px;margin:40px auto;color:#1f2937;font-size:13px;line-height:1.7}'
    + 'h1{font-size:16px;text-align:center;border-bottom:2px solid #0B4D2C;padding-bottom:8px;color:#0B4D2C}'
    + '.bloc{background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;padding:16px 20px;margin:16px 0}'
    + '.sign{display:grid;grid-template-columns:1fr 1fr;gap:32px;margin-top:40px;font-size:12px}'
    + '.sign div{border-top:1px solid #d1d5db;padding-top:8px;text-align:center}'
    + '@media print{.np{display:none}}</style></head><body>'
    + '<div style="text-align:right;font-size:11px;color:#6b7280">Lycee de Kakatare-Maroua | 2025-2026 | Le '+date+'</div>'
    + '<h1>CONVOCATION DES PARENTS / TUTEURS</h1>'
    + '<div class="bloc"><b>Nom :</b> '+eleve.nom+'<br><b>Classe :</b> '+classe+'<br><b>Motif :</b> '+motif+'<br>'
    + '<b>Detail :</b> '+absH+'h absence'+(retards>0?', '+retards+' retard(s)':'')+(sanctions>0?', '+sanctions+' sanction(s)':'')+'</div>'
    + '<p>de se presenter au bureau de la <b>Surveillance Generale</b> muni(e) du present document.</p>'
    + '<div class="sign"><div>Le Parent/Tuteur<br><br><br>Signature :</div><div>Le Surveillant General<br><br><br>Signature & Cachet :</div></div>'
    + '<div class="np" style="margin-top:24px;text-align:center"><button onclick="window.print()" style="padding:10px 24px;background:#0B4D2C;color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:14px;font-weight:700">Imprimer</button></div>'
    + '</body></html>';
  const w=window.open("","_blank"); w.document.write(html); w.document.close();
}

function FicheEleveSG({eleve, data, onClose}) {
  const {isMobile} = useDevice();
  const [vieSco, setVieSco] = useState([]);
  const [loading, setLoading] = useState(true);
  useEffect(()=>{
    sb.get("vie_scolaire","?select=*&eleve_id=eq."+eleve.id+"&order=date.desc").then(rows=>{
      setVieSco(rows||[]); setLoading(false);
    });
  },[eleve.id]);
  let totalSeances=0, totalHeures=0;
  Object.entries(data?.absences||{}).forEach(([k,absents])=>{
    const [,cl]=k.split("||");
    if(cl!==eleve.classe)return;
    if((absents||[]).includes(eleve.id)){totalSeances++; totalHeures+=getDureeSVT(cl);}
  });
  const retards=vieSco.filter(v=>v.type==="retard").length;
  const sanctions=vieSco.filter(v=>v.type==="sanction").length;
  const aC=totalHeures>=15?"#b91c1c":totalHeures>=6?"#d97706":"#15803d";
  const aB=totalHeures>=15?"#fef2f2":totalHeures>=6?"#fffbeb":"#f0fdf4";
  const aL=totalHeures>=15?"Blame — Convocation obligatoire":totalHeures>=6?"Avertissement — Convocation recommandee":"Assidu";
  const needsConvoc=totalHeures>=6||retards>=3||sanctions>=1;
  return(
    <div style={{position:"fixed",top:0,right:0,bottom:0,width:isMobile?"100vw":380,
      background:"#fff",boxShadow:"-4px 0 32px rgba(0,0,0,.18)",zIndex:1000,display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <div style={{background:"#0B3D20",padding:"16px 18px",display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
        <div style={{width:40,height:40,borderRadius:"50%",background:"rgba(212,175,55,.2)",border:"2px solid #D4AF37",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>{eleve.g==="F"?"👧":"👦"}</div>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:13,fontWeight:800,color:"#fff",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{eleve.nom}</div>
          <div style={{fontSize:10,color:"rgba(255,255,255,.5)",marginTop:2}}>{eleve.classe}</div>
        </div>
        <button onClick={onClose} style={{background:"rgba(255,255,255,.1)",border:"none",color:"#fff",borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:16}}>✕</button>
      </div>
      <div style={{padding:"10px 18px",background:aB,borderBottom:"1px solid #e5e7eb",flexShrink:0}}>
        <span style={{fontSize:11,fontWeight:700,color:aC}}>{aL}</span>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,padding:"14px 18px",flexShrink:0}}>
        {[{label:"Seances abs.",value:totalSeances,color:"#3b82f6"},{label:"Heures abs.",value:totalHeures+"h",color:aC},{label:"Retards",value:retards,color:"#d97706"},{label:"Sanctions",value:sanctions,color:"#b91c1c"}].map((k,i)=>(
          <div key={i} style={{background:"#f8fafc",borderRadius:10,padding:"10px 14px",border:"1px solid #e5e7eb"}}>
            <div style={{fontSize:10,color:"#6b7280",fontWeight:600}}>{k.label}</div>
            <div style={{fontSize:20,fontWeight:800,color:k.color,marginTop:2}}>{k.value}</div>
          </div>
        ))}
      </div>
      <div style={{flex:1,overflowY:"auto",padding:"0 18px 18px"}}>
        <div style={{fontSize:11,fontWeight:700,color:"#6b7280",marginBottom:8,textTransform:"uppercase"}}>Historique vie scolaire</div>
        {loading?<div style={{fontSize:12,color:"#9ca3af",textAlign:"center",padding:20}}>Chargement...</div>
        :vieSco.length===0?<div style={{fontSize:12,color:"#9ca3af",textAlign:"center",padding:20}}>Aucun evenement</div>
        :vieSco.map((v,i)=>(
          <div key={v.id||i} style={{display:"flex",gap:10,padding:"8px 0",borderBottom:"1px solid #f1f5f9",alignItems:"flex-start"}}>
            <span style={{fontSize:14,flexShrink:0}}>{v.type==="retard"?"⏰":v.type==="sanction"?"⚠️":"🚨"}</span>
            <div><div style={{fontSize:11,fontWeight:700,color:"#374151",textTransform:"capitalize"}}>{v.type}</div>
            <div style={{fontSize:10,color:"#6b7280"}}>{v.date}{v.motif?" · "+v.motif:""}</div></div>
          </div>
        ))}
      </div>
      {needsConvoc&&(
        <div style={{padding:"14px 18px",borderTop:"1px solid #e5e7eb",flexShrink:0}}>
          <button onClick={()=>genererConvocation(eleve,eleve.classe,totalHeures,retards,sanctions)}
            style={{width:"100%",padding:"12px 0",background:"#D4AF37",color:"#0B3D20",border:"none",borderRadius:10,fontWeight:800,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>
            📨 Generer la convocation parent
          </button>
        </div>
      )}
    </div>
  );
}

function MesClassesPage() {
  const {user, data, setData, showToast} = useApp();
  const {isMobile} = useDevice();
  const mesClasses = useMemo(()=>(user?.classes||[]).filter(Boolean), [user]);
  const [selClasse, setSelClasse] = useState(()=> (user?.classes||[]).filter(Boolean)[0] || null);
  const [search,    setSearch]    = useState("");
  const [filtreG,   setFiltreG]   = useState("all"); // all|M|F
  const [onglet,    setOnglet]    = useState("liste"); // liste | notes
  const [showAddEleve, setShowAddEleve] = useState(false);
  const [newEleveNom,  setNewEleveNom]  = useState("");
  const [newEleveGenre,setNewEleveGenre]= useState("M");

  // Présences
  const todayStr = () => new Date().toISOString().slice(0,10);
  const [datePresence, setDatePresence] = useState(todayStr());

  // Notes
  const [selTrim,   setSelTrim]   = useState("T1");
  const [selEval,   setSelEval]   = useState("E1");
  const [savingNote, setSavingNote] = useState({}); // {eleveId: 'pending'|'saved'|'error'}
  const syncTimer = useRef({});
  const [ficheEleveSG, setFicheEleveSG] = useState(null);

  useEffect(()=>{
    if (!selClasse && mesClasses.length>0) setSelClasse(mesClasses[0]);
  }, [mesClasses]);

  const eleves = selClasse ? (ELEVES_DB[selClasse]||[]) : [];
  const elevesFiltres = eleves.filter(e => {
    if (filtreG!=="all" && e.g!==filtreG) return false;
    if (search.trim() && !e.nom.toLowerCase().includes(search.trim().toLowerCase())) return false;
    return true;
  });

  // ── Présences ────────────────────────────────────────────────
  const absKey = selClasse ? `${user.id}||${selClasse}||${datePresence}` : null;
  const absentsJour = absKey ? (data?.absences?.[absKey]||[]) : [];

  const toggleAbsence = async (eleveId) => {
    if (!selClasse) return;
    const newAbsents = absentsJour.includes(eleveId)
      ? absentsJour.filter(id=>id!==eleveId)
      : [...absentsJour, eleveId];
    setData(prev => ({...prev, absences:{...(prev?.absences||{}), [absKey]:newAbsents}}));
    try {
      // Écriture via RPC sécurisée (submit_absence) plutôt que PATCH/POST direct.
      const ok = await sb.rpc("submit_absence", {
        p_ens_id: user.id, p_classe: selClasse, p_seance: datePresence, p_absents: newAbsents
      });
      if (!ok) showToast("⚠ Présence non sauvegardée — réessaie", false);
    } catch { showToast("⚠ Erreur réseau — présence non sauvegardée", false); }
  };

  // ── Notes ────────────────────────────────────────────────────
  const evalCode = `${selTrim}-${selEval}`;
  const notesEval = selClasse ? (data?.notes?.[`${selClasse}||${evalCode}`]||{}) : {};

  const saveNote = (eleveId, valeur) => {
    if (!selClasse) return;
    const num = valeur==="" ? null : Math.max(0, Math.min(20, Number(valeur)));
    setData(prev => {
      const k = `${selClasse}||${evalCode}`;
      const cur = {...(prev?.notes?.[k]||{})};
      if (num===null) delete cur[eleveId]; else cur[eleveId]=num;
      return {...prev, notes:{...(prev?.notes||{}), [k]:cur}};
    });
    setSavingNote(prev=>({...prev, [`${eleveId}-${evalCode}`]:"pending"}));
    const timerKey = `${eleveId}-${evalCode}`;
    if (syncTimer.current[timerKey]) clearTimeout(syncTimer.current[timerKey]);
    syncTimer.current[timerKey] = setTimeout(async () => {
      try {
        // Écriture via RPC sécurisée (submit_note) plutôt que PATCH/POST direct sur la table —
        // notes passe en RLS avec écriture directe bloquée ; seule cette fonction peut écrire.
        const ok = await sb.rpc("submit_note", {
          p_classe: selClasse, p_evaluation: evalCode, p_eleve_id: eleveId, p_note: num
        });
        setSavingNote(prev=>({...prev, [`${eleveId}-${evalCode}`]: ok?"saved":"error"}));
        if (!ok) showToast("⚠ Note non sauvegardée — réessaie", false);
      } catch {
        setSavingNote(prev=>({...prev, [`${eleveId}-${evalCode}`]:"error"}));
        showToast("⚠ Erreur réseau — note non sauvegardée", false);
      }
    }, 600);
  };

  // ── Ajout / retrait élève (persisté dans eleves_import) ──────
  const persisterClasse = async (nouvelleListe) => {
    const ok = await sb.upsert("eleves_import", {classe:selClasse, donnees:JSON.stringify(nouvelleListe)}, "classe");
    if (!ok) showToast("⚠ Sauvegarde Supabase échouée — élève visible ici seulement", false);
    return ok;
  };

  const addEleve = async () => {
    const nom = newEleveNom.trim().toUpperCase().split(/\s+/).join(" ");
    if (!selClasse || nom.length < 3) { showToast("⚠ Nom trop court (min 3 car.)", false); return; }
    const id = selClasse.replace(/[^a-zA-Z0-9]/g,"_") + "_" + Date.now();
    const nouvel = {id, nom, g:newEleveGenre};
    const nouvelleListe = [...(ELEVES_DB[selClasse]||[]), nouvel];
    ELEVES_DB[selClasse] = nouvelleListe;
    setNewEleveNom(""); setShowAddEleve(false);
    showToast(`✓ ${nom} ajouté(e) à ${selClasse}`);
    await persisterClasse(nouvelleListe);
  };

  const retirerEleve = async (eleve) => {
    const nouvelleListe = (ELEVES_DB[selClasse]||[]).filter(e=>e.id!==eleve.id);
    ELEVES_DB[selClasse] = nouvelleListe;
    showToast(`✓ ${eleve.nom} retiré(e)`);
    await persisterClasse(nouvelleListe);
  };

  // ── Rendu ──────────────────────────────────────────────────────
  if (mesClasses.length===0) {
    return (
      <div style={{padding:isMobile?16:24}}>
        <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:40, textAlign:"center"}}>
          <div style={{fontSize:32, marginBottom:10}}>📚</div>
          <p style={{color:C.txtMuted, fontSize:13}}>Aucune classe ne t'est encore assignée.</p>
        </div>
      </div>
    );
  }

  return (
    <div style={{display:"flex",flexDirection:"column",flex:1,minHeight:0,overflow:"hidden"}}>
    <div style={{flex:1, minHeight:0, overflowY:"auto", padding:isMobile?12:24, display:"flex", flexDirection:"column", gap:14}}>
      {/* Sélecteur de classe */}
      <div style={{display:"flex", gap:8, overflowX:"auto", paddingBottom:4, flexShrink:0}}>
        {mesClasses.map(cl=>(
          <button key={cl} onClick={()=>setSelClasse(cl)}
            style={{flexShrink:0, padding: isMobile?"9px 14px":"7px 14px", borderRadius:20, fontSize: isMobile?12.5:12, fontWeight:700, fontFamily:"inherit", cursor:"pointer", whiteSpace:"nowrap",
              border:`1.5px solid ${selClasse===cl?C.green:C.border}`,
              background:selClasse===cl?C.green:"#eef1f5",
              color:selClasse===cl?"#fff":C.txtMuted}}>
            {displayCl(cl)} <span style={{opacity:.75}}>({(ELEVES_DB[cl]||[]).length})</span>
          </button>
        ))}
      </div>

      {/* Onglets Liste / Notes */}
      <div style={{display:"flex", gap:8, flexShrink:0}}>
        {[{id:"liste",label:"👥 Liste & présences"},...(user?.role==="surveillant_general"?[]:[{id:"notes",label:"📝 Notes"}])].map(t=>(
          <button key={t.id} onClick={()=>setOnglet(t.id)}
            style={{flex: isMobile?1:"none", padding: isMobile?"10px":"8px 16px", borderRadius:9, fontSize:12.5, fontWeight:700, fontFamily:"inherit", cursor:"pointer",
              border:`1.5px solid ${onglet===t.id?C.green:C.border}`,
              background:onglet===t.id?C.greenPale:"#eef1f5",
              color:onglet===t.id?C.green:C.txtMuted}}>
            {t.label}
          </button>
        ))}
      </div>

      {onglet==="liste" ? (
        <>
          {/* Barre présences */}
          <div style={{background:C.greenPale, border:`1px solid ${C.greenBorder}`, borderRadius:11, padding: isMobile?"12px":"12px 16px",
            display:"flex", flexDirection: isMobile?"column":"row", alignItems: isMobile?"stretch":"center", gap:10, flexShrink:0}}>
            <span style={{fontSize:12.5, fontWeight:700, color:C.green, flexShrink:0}}>📅 Présences du jour</span>
            <div style={{display:"flex", gap:8, flex:1}}>
              <input type="date" value={datePresence} onChange={e=>setDatePresence(e.target.value)}
                style={{flex:1, padding: isMobile?"9px":"6px 10px", borderRadius:8, border:`1px solid ${C.greenBorder}`, fontSize:12.5, fontFamily:"inherit"}}/>
              <button onClick={()=>setDatePresence(todayStr())}
                style={{padding: isMobile?"9px 14px":"6px 12px", borderRadius:8, border:"none", background:C.green, color:"#fff", fontSize:12, fontWeight:700, cursor:"pointer", fontFamily:"inherit", flexShrink:0}}>
                Aujourd'hui
              </button>
            </div>
          </div>

          {/* Recherche + filtre genre */}
          <div style={{display:"flex", flexDirection: isMobile?"column":"row", gap:8, flexShrink:0}}>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Rechercher un élève…"
              style={{flex:1, padding: isMobile?"10px 12px":"8px 12px", borderRadius:9, border:`1px solid ${C.border}`, fontSize:13, fontFamily:"inherit"}}/>
            <div style={{display:"flex", gap:6}}>
              {[{id:"all",label:"Tous"},{id:"M",label:"♂ Garçons"},{id:"F",label:"♀ Filles"}].map(f=>(
                <button key={f.id} onClick={()=>setFiltreG(f.id)}
                  style={{flex: isMobile?1:"none", padding: isMobile?"9px 10px":"6px 12px", borderRadius:8, fontSize:11.5, fontWeight:700, fontFamily:"inherit", cursor:"pointer", whiteSpace:"nowrap",
                    border:`1.5px solid ${filtreG===f.id?C.blue:C.border}`,
                    background:filtreG===f.id?C.bluePale:"#eef1f5",
                    color:filtreG===f.id?C.blue:C.txtMuted}}>
                  {f.label}
                </button>
              ))}
            </div>
          </div>

          {/* Liste élèves */}
          <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, overflow:"hidden", flexShrink:0}}>
            {elevesFiltres.length===0 ? (
              <div style={{padding:32, textAlign:"center", color:C.txtLight, fontSize:13}}>Aucun élève trouvé</div>
            ) : elevesFiltres.map((e,i)=>{
              const absent = absentsJour.includes(e.id);
              return (
                <div key={e.id} style={{display:"flex", alignItems:"center", gap:10, padding: isMobile?"10px 12px":"9px 14px",
                  borderBottom: i<elevesFiltres.length-1?`1px solid #f1f5f9`:"none"}}>
                  <span style={{fontSize:11, color:C.txtLight, width:20, flexShrink:0}}>{i+1}</span>
                  <span style={{fontSize:14, flexShrink:0}}>{e.g==="F"?"👧":"👦"}</span>
                  <span onClick={()=>user?.role==="surveillant_general"?setFicheEleveSG({...e,classe:selClasse}):null}
                    style={{flex:1,fontSize:isMobile?12.5:13,fontWeight:600,color:C.txt,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                      cursor:user?.role==="surveillant_general"?"pointer":"default"}}>{e.nom}</span>
                  <button onClick={()=>toggleAbsence(e.id)}
                    style={{padding: isMobile?"7px 12px":"5px 12px", borderRadius:7, border:"none", fontSize:11, fontWeight:700, cursor:"pointer", fontFamily:"inherit", flexShrink:0,
                      background:absent?"#fef2f2":"#f0fdf4", color:absent?"#b91c1c":"#166534"}}>
                    {absent ? "✕ Absent" : "✓ Présent"}
                  </button>
                  <button onClick={()=>retirerEleve(e)} title="Retirer cet élève"
                    style={{width:26, height:26, borderRadius:6, border:"none", background:"transparent", color:"#cbd5e1", fontSize:14, cursor:"pointer", flexShrink:0}}>
                    ✕
                  </button>
                </div>
              );
            })}
          </div>

          {/* Ajout élève */}
          {showAddEleve ? (
            <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.green}`, padding: isMobile?14:16, display:"flex", flexDirection: isMobile?"column":"row", gap:10, flexShrink:0}}>
              <input value={newEleveNom} onChange={e=>setNewEleveNom(e.target.value)} placeholder="Nom complet de l'élève"
                style={{flex:1, padding: isMobile?"10px 12px":"8px 12px", borderRadius:8, border:`1px solid ${C.border}`, fontSize:13, fontFamily:"inherit"}}/>
              <select value={newEleveGenre} onChange={e=>setNewEleveGenre(e.target.value)}
                style={{padding: isMobile?"10px 12px":"8px 12px", borderRadius:8, border:`1px solid ${C.border}`, fontSize:13, fontFamily:"inherit"}}>
                <option value="M">Garçon</option>
                <option value="F">Fille</option>
              </select>
              <div style={{display:"flex", gap:8}}>
                <button onClick={addEleve} style={{flex:1, padding: isMobile?"10px":"8px 16px", borderRadius:8, border:"none", background:C.green, color:"#fff", fontSize:12.5, fontWeight:700, cursor:"pointer", fontFamily:"inherit"}}>✓ Ajouter</button>
                <button onClick={()=>{setShowAddEleve(false);setNewEleveNom("");}} style={{padding: isMobile?"10px 14px":"8px 14px", borderRadius:8, border:"none", background:"transparent", color:C.txtMuted, fontSize:12.5, cursor:"pointer", fontFamily:"inherit"}}>Annuler</button>
              </div>
            </div>
          ) : (
            <button onClick={()=>setShowAddEleve(true)}
              style={{padding: isMobile?"12px":"10px", borderRadius:10, border:`1.5px dashed ${C.border}`, background:"transparent", color:C.txtMuted, fontSize:12.5, fontWeight:700, cursor:"pointer", fontFamily:"inherit", flexShrink:0}}>
              + Nouvel élève
            </button>
          )}
        </>
      ) : (
        <>
          {/* Sélecteurs Trimestre / Évaluation */}
          <div style={{display:"flex", flexDirection: isMobile?"column":"row", gap:10, flexShrink:0}}>
            <div style={{display:"flex", gap:6, flex:1}}>
              {["T1","T2","T3"].map(t=>(
                <button key={t} onClick={()=>setSelTrim(t)}
                  style={{flex:1, padding: isMobile?"9px":"7px", borderRadius:8, fontSize:12, fontWeight:700, fontFamily:"inherit", cursor:"pointer",
                    border:`1.5px solid ${selTrim===t?C.green:"#cbd5e1"}`, background:selTrim===t?C.green:"#eef1f5", color:selTrim===t?"#fff":"#475569"}}>
                  {t}
                </button>
              ))}
            </div>
            <div style={{display:"flex", gap:6, flex:1}}>
              {["E1","E2"].map(ev=>(
                <button key={ev} onClick={()=>setSelEval(ev)}
                  style={{flex:1, padding: isMobile?"9px":"7px", borderRadius:8, fontSize:12, fontWeight:700, fontFamily:"inherit", cursor:"pointer",
                    border:`1.5px solid ${selEval===ev?C.blue:"#cbd5e1"}`, background:selEval===ev?C.bluePale:"#eef1f5", color:selEval===ev?C.blue:"#475569"}}>
                  {ev==="E1"?"Évaluation 1":"Évaluation 2"}
                </button>
              ))}
            </div>
          </div>

          {/* Liste notes */}
          <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, overflow:"hidden", flexShrink:0}}>
            {eleves.length===0 ? (
              <div style={{padding:32, textAlign:"center", color:C.txtLight, fontSize:13}}>Aucun élève dans cette classe</div>
            ) : eleves.map((e,i)=>{
              const val = notesEval[e.id];
              const status = savingNote[`${e.id}-${evalCode}`];
              return (
                <div key={e.id} style={{display:"flex", alignItems:"center", gap:10, padding: isMobile?"10px 12px":"9px 14px",
                  borderBottom: i<eleves.length-1?`1px solid #f1f5f9`:"none"}}>
                  <span style={{fontSize:11, color:C.txtLight, width:20, flexShrink:0}}>{i+1}</span>
                  <span style={{flex:1, fontSize: isMobile?12.5:13, fontWeight:600, color:C.txt, minWidth:0, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>{e.nom}</span>
                  <input key={`${e.id}-${evalCode}`} type="number" min="0" max="20" step="0.5" defaultValue={val??""} onChange={ev=>saveNote(e.id, ev.target.value)}
                    placeholder="—"
                    style={{width: isMobile?60:64, padding: isMobile?"8px":"6px 8px", borderRadius:7, border:`1px solid ${C.border}`, fontSize:13, fontFamily:"inherit", textAlign:"center", flexShrink:0}}/>
                  <span style={{width:16, flexShrink:0, fontSize:13}}>
                    {status==="pending"?"⏳":status==="saved"?"✅":status==="error"?"⚠️":""}
                  </span>
                </div>
              );
            })}
          </div>
        </>
      )}
      </div>
    </div>
  );
}


// ═══════════════════════════════════════════════════════════════
// MON PROGRAMME (enseignant) — synthèse de couverture par classe
// ═══════════════════════════════════════════════════════════════
function MonProgrammePage() {
  const {user, data} = useApp();
  const {isMobile} = useDevice();
  const [trim, setTrim] = useState(0); // 0=Annuel,1=T1,2=T2,3=T3

  if (!data) return (
    <div style={{padding:"60px",textAlign:"center",color:C.txtMuted}}>
      <Spinner size={28} color={C.green}/><div style={{marginTop:12}}>Chargement…</div>
    </div>
  );

  const mesClasses = [...new Set((user?.classes||[]).filter(Boolean))];

  const rows = mesClasses.map(cl => {
    const code = resolveProgCode(cl);
    const meta = code ? PROG_META[code] : null;
    const key  = `${user.id}||${cl}`;
    const prog = (data?.prog||{})[key]||[];

    let lp = meta?.lpRef||0, tpP = meta?.tp?.length||0;
    let range = null;
    if (trim > 0 && code) {
      const tk    = ["T1","T2","T3"][trim-1];
      range = getTrimRange(code, tk);
      if (range) {
        const lecons = LECONS_DATA[code]||[];
        lp  = lecons.filter(l=>l.n>=range[0]&&l.n<=range[1]).length||lp;
        tpP = (meta?.tp||[]).filter(n=>n>=range[0]&&n<=range[1]).length;
      }
    }
    const lfTrim = trim===0 ? prog.length : prog.filter(n=>{
      const tk=["T1","T2","T3"][trim-1]; const r=code?getTrimRange(code,tk):null;
      return r&&n>=r[0]&&n<=r[1];
    }).length;

    const taux   = lp>0 ? Math.min(100, Math.round(lfTrim/lp*100)) : null;
    // tpFait limité à la même plage que tpP (range) — sinon des TP/TD faits en avance
    // sur un autre trimestre se comptaient dans le trimestre affiché, faussant le taux.
    const tpFait = (meta?.tp||[]).filter(n => prog.includes(n) && (trim===0 || (range && n>=range[0] && n<=range[1]))).length;
    const tauxTP = tpP>0 ? Math.min(100, Math.round(tpFait/tpP*100)) : null;
    const ef     = (ELEVES_DB[cl]||[]).length;

    const digKey   = `${user.id}||${cl}||dig`;
    const digProg  = (data?.prog?.[digKey]||[]);
    const leconsList = code ? (LECONS_DATA[code]||[]) : [];
    const ldTot  = leconsList.filter(l=>l.d===1).length;
    const ldFait = digProg.filter(n=>{ const l=leconsList.find(x=>x.n===n); return l&&l.d===1; }).length;
    const tauxDig = ldTot>0 ? Math.min(100, Math.round(ldFait/ldTot*100)) : null;

    return {cl, lp, lf:lfTrim, taux, tpP, tpFait, tauxTP, ef, ldTot, ldFait, tauxDig};
  });

  const tauxValides = rows.map(r=>r.taux).filter(v=>v!==null);
  const tauxMoyen = tauxValides.length>0 ? Math.round(tauxValides.reduce((s,v)=>s+v,0)/tauxValides.length) : 0;
  const totalFait = rows.reduce((s,r)=>s+r.lf,0);
  const totalPrevu = rows.reduce((s,r)=>s+r.lp,0);

  const badge = (val, taux, color) => (
    <div style={{display:"inline-flex", alignItems:"center", gap:7}}>
      <span style={{fontSize:12.5, fontWeight:600, color:"#334155"}}>{val}</span>
      {taux!==null ? (
        <span style={{fontSize:10.5, fontWeight:800, padding:"2px 7px", borderRadius:20, background:`${color}15`, color}}>{taux}%</span>
      ) : (
        <span style={{fontSize:10.5, fontWeight:600, padding:"2px 7px", borderRadius:20, background:"#f1f5f9", color:"#94a3b8"}}>—</span>
      )}
    </div>
  );

  return (
    <div style={{padding:isMobile?12:24, display:"flex", flexDirection:"column", gap:14}}>
      {/* KPIs */}
      <div style={{display:"flex", gap:10, flexWrap:"wrap"}}>
        <KpiCard label="Couverture moyenne" value={`${tauxMoyen}%`} sub={tauxMoyen>=75?"Objectif atteint ✓":"Sous l'objectif"} subColor={taux2col(tauxMoyen)} iconEmoji="📊" bg={C.greenPale} loading={false} delay={0}/>
        <KpiCard label="Leçons dispensées" value={totalFait} sub={`sur ${totalPrevu} prévues`} iconEmoji="✅" bg={C.bluePale} subColor={C.blue} loading={false} delay={0.05}/>
        <KpiCard label="Mes classes" value={mesClasses.length} sub={`${rows.reduce((s,r)=>s+r.ef,0)} élèves`} iconEmoji="📚" bg={C.amberPale} subColor={C.amber} loading={false} delay={0.1}/>
      </div>

      {/* Sélecteur période */}
      <div style={{display:"flex", gap:6}}>
        {[{id:0,label:"Année"},{id:1,label:"T1"},{id:2,label:"T2"},{id:3,label:"T3"}].map(t=>(
          <button key={t.id} onClick={()=>setTrim(t.id)}
            style={{flex: isMobile?1:"none", padding: isMobile?"9px":"7px 16px", borderRadius:8, fontSize:12, fontWeight:700, fontFamily:"inherit", cursor:"pointer",
              border:`1.5px solid ${trim===t.id?C.green:"#cbd5e1"}`, background:trim===t.id?C.green:"#eef1f5", color:trim===t.id?"#fff":"#475569"}}>
            {t.label}
          </button>
        ))}
      </div>

      {/* Liste classes */}
      {mesClasses.length===0 ? (
        <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:40, textAlign:"center", color:C.txtMuted}}>
          Aucune classe assignée.
        </div>
      ) : isMobile ? (
        <div style={{display:"flex", flexDirection:"column", gap:10}}>
          {rows.map(r=>{
            const statut = r.taux===null ? null : r.taux<50
              ? {label:"Alerte", bg:"#fef2f2", fg:"#b91c1c", dot:"#ef4444"}
              : r.taux>=75
              ? {label:"Objectif", bg:"#f0fdf4", fg:"#166534", dot:"#16a34a"}
              : {label:"En cours", bg:"#fffbeb", fg:"#92400e", dot:"#f59e0b"};
            return (
              <div key={r.cl} style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, borderLeft:`3px solid ${r.taux!==null?taux2col(r.taux):"#cbd5e1"}`, padding:14}}>
                <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10}}>
                  <div>
                    <div style={{fontSize:13.5, fontWeight:800, color:C.txt}}>{displayCl(r.cl)}</div>
                    <div style={{fontSize:10.5, color:C.txtMuted}}>{r.ef} élève{r.ef>1?"s":""}</div>
                  </div>
                  {statut && (
                    <span style={{display:"inline-flex", alignItems:"center", gap:5, padding:"4px 10px", borderRadius:20, background:statut.bg, color:statut.fg, fontSize:10.5, fontWeight:700}}>
                      <span style={{width:5,height:5,borderRadius:"50%",background:statut.dot}}/>{statut.label}
                    </span>
                  )}
                </div>
                <div style={{display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:8}}>
                  <div><div style={{fontSize:9.5, color:C.txtLight, marginBottom:3}}>Programme</div>{badge(`${r.lf}/${r.lp}`, r.taux, r.taux!==null?taux2col(r.taux):"#94a3b8")}</div>
                  <div><div style={{fontSize:9.5, color:C.txtLight, marginBottom:3}}>TP/TD</div>{badge(`${r.tpFait}/${r.tpP}`, r.tauxTP, r.tauxTP!==null?taux2col(r.tauxTP):"#94a3b8")}</div>
                  <div><div style={{fontSize:9.5, color:C.txtLight, marginBottom:3}}>Digital</div>{badge(`${r.ldFait}/${r.ldTot}`, r.tauxDig, "#0369a1")}</div>
                </div>
              </div>
            );
          })}
        </div>
      ) : (
        <div style={{background:C.white, borderRadius:14, border:`1px solid ${C.border}`, overflow:"hidden", boxShadow:"0 1px 3px rgba(0,0,0,.04)"}}>
          <table style={{width:"100%", borderCollapse:"collapse", fontSize:12.5}}>
            <thead>
              <tr style={{background:"#fafbfc", borderBottom:`2px solid ${C.border}`}}>
                {["Classe","Programme","TP / TD","Digitalisation","Statut"].map((h,i)=>(
                  <th key={i} style={{padding:"11px 14px", textAlign:i===4?"center":"left", color:"#64748b", fontSize:10.5, fontWeight:700, textTransform:"uppercase", letterSpacing:".04em"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map(r=>{
                const statut = r.taux===null ? null : r.taux<50
                  ? {label:"Alerte", bg:"#fef2f2", fg:"#b91c1c", dot:"#ef4444"}
                  : r.taux>=75
                  ? {label:"Objectif", bg:"#f0fdf4", fg:"#166534", dot:"#16a34a"}
                  : {label:"En cours", bg:"#fffbeb", fg:"#92400e", dot:"#f59e0b"};
                return (
                  <tr key={r.cl} style={{borderBottom:`1px solid #f1f5f9`}}>
                    <td style={{padding:"12px 14px"}}>
                      <div style={{fontSize:12.5,fontWeight:700,color:"#1e293b"}}>{displayCl(r.cl)}</div>
                      <div style={{fontSize:10.5,color:"#94a3b8",marginTop:1}}>{r.ef} élève{r.ef>1?"s":""}</div>
                    </td>
                    <td style={{padding:"12px 14px"}}>{badge(`${r.lf}/${r.lp}`, r.taux, r.taux!==null?taux2col(r.taux):"#94a3b8")}</td>
                    <td style={{padding:"12px 14px"}}>{badge(`${r.tpFait}/${r.tpP}`, r.tauxTP, r.tauxTP!==null?taux2col(r.tauxTP):"#94a3b8")}</td>
                    <td style={{padding:"12px 14px"}}>{badge(`${r.ldFait}/${r.ldTot}`, r.tauxDig, "#0369a1")}</td>
                    <td style={{padding:"12px 14px", textAlign:"center"}}>
                      {statut && (
                        <span style={{display:"inline-flex", alignItems:"center", gap:6, padding:"5px 12px", borderRadius:20, background:statut.bg, color:statut.fg, fontSize:11, fontWeight:700}}>
                          <span style={{width:6,height:6,borderRadius:"50%",background:statut.dot}}/>{statut.label}
                        </span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// CAHIER DE TEXTE (enseignant) — wrapper de sélection de classe
// réutilise EnsClasLecons (liste des leçons + cases à cocher),
// déjà existant et fonctionnel pour la vue admin par enseignant.
// ═══════════════════════════════════════════════════════════════
function CahierDeTextePage() {
  const {user, data, setData, showToast} = useApp();
  const {isMobile} = useDevice();
  const mesClasses = useMemo(()=>(user?.classes||[]).filter(Boolean), [user]);
  const [selClasse, setSelClasse] = useState(()=>{
    const cls = (user?.classes||[]).filter(Boolean);
    return cls.length===1 ? cls[0] : null;
  });

  if (mesClasses.length===0) {
    return (
      <div style={{padding:isMobile?16:24}}>
        <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:40, textAlign:"center"}}>
          <div style={{fontSize:32, marginBottom:10}}>📖</div>
          <p style={{color:C.txtMuted, fontSize:13}}>Aucune classe ne t'est encore assignée.</p>
        </div>
      </div>
    );
  }

  if (selClasse) {
    return (
      <div style={{display:"flex",flexDirection:"column",flex:1,minHeight:0,overflow:"hidden"}}>
        <div style={{flex:1, minHeight:0, overflowY:"auto"}}>
          <EnsClasLecons ens={user} cl={selClasse} data={data} setData={setData} showToast={showToast}
            onBack={()=>setSelClasse(null)}/>
        </div>
      </div>
    );
  }

  return (
    <div style={{display:"flex",flexDirection:"column",flex:1,minHeight:0,overflow:"hidden"}}>
    <div style={{flex:1, minHeight:0, overflowY:"auto", padding:isMobile?16:24, display:"flex", flexDirection:"column", gap:14}}>
      <h2 style={{fontSize:15, fontWeight:800, color:C.txt, margin:0}}>📖 Choisis une classe</h2>
      <div style={{display:"grid", gridTemplateColumns: isMobile?"1fr":"repeat(auto-fill,minmax(200px,1fr))", gap:10}}>
        {mesClasses.map(cl=>{
          const code = resolveProgCode(cl);
          const meta = code ? PROG_META[code] : null;
          const key  = `${user.id}||${cl}`;
          const fait = ((data?.prog||{})[key]||[]).length;
          const taux = meta?.lpRef>0 ? Math.min(100, Math.round(fait/meta.lpRef*100)) : 0;
          return (
            <button key={cl} onClick={()=>setSelClasse(cl)}
              style={{textAlign:"left", padding:16, borderRadius:12, border:`1px solid ${C.border}`, background:C.white, cursor:"pointer", fontFamily:"inherit"}}>
              <div style={{fontSize:13.5, fontWeight:800, color:C.txt, marginBottom:6}}>{displayCl(cl)}</div>
              <div style={{fontSize:11, color:C.txtMuted, marginBottom:8}}>{(ELEVES_DB[cl]||[]).length} élèves</div>
              <div style={{height:6, borderRadius:3, background:"#e2e8f0", overflow:"hidden"}}>
                <div style={{width:`${taux}%`, height:"100%", background:taux2col(taux)}}/>
              </div>
              <div style={{fontSize:10.5, color:taux2col(taux), fontWeight:700, marginTop:5}}>{fait} leçons faites · {taux}%</div>
            </button>
          );
        })}
      </div>
    </div>
      {ficheEleveSG&&user?.role==="surveillant_general"&&(
        <FicheEleveSG eleve={ficheEleveSG} data={data} onClose={()=>setFicheEleveSG(null)}/>
      )}
    </div>
  );
}

function ElevesPage() {
  const {isMobile} = useDevice();
  const {pendingClasseSelect, setPendingClasseSelect} = useApp();
  // ── État principal ────────────────────────────────────────────────
  const [selClasse, setSelClasse] = useState(() => "6ème 1");
  const [search, setSearch]       = useState("");
  const [filtreGenre, setFiltreGenre] = useState("all");
  const [vue, setVue]             = useState("registre");

  // Navigation ciblée depuis la recherche globale
  useEffect(() => {
    if (!pendingClasseSelect) return;
    setSelClasse(pendingClasseSelect);
    setPendingClasseSelect(null);
  }, [pendingClasseSelect]);

  // ── Base locale — modifiable (ajout / retrait) ───────────────────
  const [localDB, setLocalDB] = useState(() => {
    // Cloner ELEVES_DB pour pouvoir le modifier localement
    const clone = {};
    for (const k in ELEVES_DB) clone[k] = [...ELEVES_DB[k]];
    return clone;
  });

  // ── Modal ajout / retrait ─────────────────────────────────────────
  const [modal, setModal]           = useState(null); // null | "ajout" | {type:"retrait",eleve}
  const [newNom, setNewNom]         = useState("");
  const [newGenre, setNewGenre]     = useState("M");
  const [confirmRetrait, setConfirmRetrait] = useState(null);
  const [toast, setToast]           = useState(null);

  function showToast(msg, ok=true) {
    setToast({msg,ok});
    setTimeout(()=>setToast(null), 2800);
  }

  // ── Ajout d'un élève ─────────────────────────────────────────────
  function ajouterEleve() {
    const nom = newNom.trim().toUpperCase();
    if (!nom || nom.length < 3) return showToast("⚠ Nom trop court", false);
    const safe = selClasse.replace(/[^a-zA-Z0-9]/g,'_');
    const id   = `${safe}_new_${Date.now()}`;
    const nouvelleListe = [...(localDB[selClasse]||[]), {id, nom, g:newGenre}];
    setLocalDB(prev => ({ ...prev, [selClasse]: nouvelleListe }));
    ELEVES_DB[selClasse] = nouvelleListe;
    showToast(`✓ ${nom} ajouté(e) en ${selClasse}`);
    setNewNom(""); setModal(null);
    sb.upsert("eleves_import", {classe:selClasse, donnees:JSON.stringify(nouvelleListe)}, "classe")
      .then(ok=>{ if(!ok) showToast("⚠ Sauvegarde Supabase échouée — élève visible ici seulement", false); })
      .catch(()=>showToast("⚠ Sauvegarde Supabase échouée", false));
  }

  // ── Retrait d'un élève ───────────────────────────────────────────
  function retirerEleve(eleve) {
    const nouvelleListe = (localDB[selClasse]||[]).filter(e => e.id !== eleve.id);
    setLocalDB(prev => ({ ...prev, [selClasse]: nouvelleListe }));
    ELEVES_DB[selClasse] = nouvelleListe;
    showToast(`✓ ${eleve.nom} retiré(e)`);
    setConfirmRetrait(null);
    sb.upsert("eleves_import", {classe:selClasse, donnees:JSON.stringify(nouvelleListe)}, "classe")
      .then(ok=>{ if(!ok) showToast("⚠ Sauvegarde Supabase échouée", false); })
      .catch(()=>showToast("⚠ Sauvegarde Supabase échouée", false));
  }

  // ── Données de la classe sélectionnée ────────────────────────────
  const eleves   = localDB[selClasse] || [];
  const filles   = eleves.filter(e=>e.g==="F").length;
  const garcons  = eleves.length - filles;

  const elevesFiltres = useMemo(() => eleves.filter(e => {
    const ms = !search || e.nom.toLowerCase().includes(search.toLowerCase());
    const mg = filtreGenre==="all" || e.g===filtreGenre;
    return ms && mg;
  }), [eleves, search, filtreGenre]);

  // ── Stats globales ────────────────────────────────────────────────
  const TOTAL_LOC   = Object.values(localDB).reduce((s,cl)=>s+cl.length, 0);
  const FILLES_LOC  = Object.values(localDB).reduce((s,cl)=>s+cl.filter(e=>e.g==="F").length, 0);
  const GARCONS_LOC = TOTAL_LOC - FILLES_LOC;
  const ALL_CLS     = Object.keys(localDB);

  // ── Groupement sidebar ────────────────────────────────────────────
  const GROUPES_ORDRE = [
    {label:"6ème",     classes:["6ème 1","6ème 2","6ème 3"]},
    {label:"5ème",     classes:["5ème 1","5ème 2","5ème 3"]},
    {label:"4ème",     classes:["4ème ALL","4ème ARB","4ème CHN","4ème ITA","4ème ESP"]},
    {label:"3ème",     classes:["3ème ALL","3ème ARB","3ème CHN","3ème ESP","3ème ITA"]},
    {label:"2nde A4",  classes:["2nde ALL","2nde ARB","2nde CHN","2nde ITA","2nde ESP"]},
    {label:"2nde C",   classes:["2nde C"]},
    {label:"1ère A4",  classes:["1ère A4 ALL","1ère A4 ARB","1ère A4 ESP","1ère CHN","1ère ITA"]},
    {label:"1ère S/D", classes:["1ère C","1ère D","1ère Ti"]},
    {label:"Tle A4",   classes:["Tle A4 ALL","Tle A4 ARB","Tle A4 CHN","Tle A4 ITA","Tle A4 ESP"]},
    {label:"Tle S/D",  classes:["Tle C","Tle D","Tle Ti"]},
  ];

  const niveau = getNiveau(selClasse);
  const prog   = niveau ? PROGRAMME_SVTEEHB[niveau] : null;

  return (
    <div style={{display:"flex",flexDirection:"column",flex:1,minHeight:0,overflow:"hidden"}}>

      {/* ── Barre d'onglets + KPIs (remplace l'ancien header propre) ── */}
      <div style={{background:C.white,borderBottom:`1px solid ${C.border}`,padding:"8px 20px",display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
        <div style={{display:"flex",gap:4,background:"#f1f5f9",borderRadius:9,padding:3}}>
          {[["registre","📋 Registre"],["stats","📊 Stats"],["programmes","📚 Programmes"]].map(([id,label])=>(
            <button key={id} onClick={()=>setVue(id)}
              style={{padding:"5px 12px",borderRadius:7,fontSize:12,fontWeight:700,border:"none",cursor:"pointer",background:vue===id?C.white:"transparent",color:vue===id?C.txt:C.txtMuted,boxShadow:vue===id?"0 1px 4px rgba(0,0,0,.08)":"none"}}>
              {label}
            </button>
          ))}
        </div>
        <div style={{flex:1}}/>
        <div style={{display:"flex",gap:8}}>
          {[
            {v:TOTAL_LOC,  l:"Élèves", col:C.green},
            {v:FILLES_LOC, l:"Filles",  col:C.pink},
            {v:GARCONS_LOC,l:"Garçons", col:C.blue},
            {v:ALL_CLS.length,l:"Classes",col:C.purple},
          ].map((k,i)=>(
            <div key={i} style={{textAlign:"center",padding:"3px 10px",background:C.greenPale,border:`1px solid ${C.greenBorder}`,borderRadius:8}}>
              <div style={{fontSize:14,fontWeight:800,color:k.col}}>{k.v}</div>
              <div style={{fontSize:9,color:C.txtMuted,fontWeight:600}}>{k.l}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{display:"flex", flexDirection: isMobile?"column":"row", flex:1,minHeight:0,overflow:"hidden"}}>

        {/* ── Sidebar classes — vertical (desktop) ou chips horizontaux (mobile) ── */}
        {isMobile ? (
          <div style={{display:"flex", gap:6, overflowX:"auto", padding:"10px 12px", background:C.sidebar, flexShrink:0, scrollbarWidth:"none", WebkitOverflowScrolling:"touch"}}>
            {GROUPES_ORDRE.flatMap(g=>g.classes).map(cl=>{
              const ef = (localDB[cl]||[]).length;
              const isActive = cl===selClasse;
              return (
                <button key={cl} onClick={()=>setSelClasse(cl)}
                  style={{flexShrink:0, display:"flex", alignItems:"center", gap:5,
                    padding:"6px 11px", borderRadius:20, border:"none",
                    background:isActive?C.green:"rgba(255,255,255,.08)",
                    color:isActive?"#fff":"rgba(255,255,255,.7)",
                    fontSize:11.5, fontWeight:isActive?700:500, whiteSpace:"nowrap",
                    cursor:"pointer", fontFamily:"inherit"}}>
                  {cl}
                  <span style={{fontSize:9, opacity:.8}}>({ef})</span>
                </button>
              );
            })}
          </div>
        ) : (
          <aside style={{width:200,minWidth:200,background:C.sidebar,overflowY:"auto",scrollbarWidth:"none",paddingBottom:20,flexShrink:0}}>
            {GROUPES_ORDRE.map(({label, classes}) => (
              <div key={label}>
                <div style={{padding:"8px 12px 4px",fontSize:9,fontWeight:800,color:"rgba(255,255,255,.35)",textTransform:"uppercase",letterSpacing:".1em"}}>{label}</div>
                {classes.map(cl => {
                  const ef = (localDB[cl]||[]).length;
                  const isActive = cl===selClasse;
                  return (
                    <div key={cl} onClick={()=>setSelClasse(cl)}
                      style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"6px 12px",cursor:"pointer",background:isActive?"rgba(34,197,94,.15)":"transparent",color:isActive?"#4ade80":"rgba(255,255,255,.55)",fontSize:11,fontWeight:isActive?700:400,transition:"all .15s"}}
                      onMouseEnter={e=>{if(!isActive)e.currentTarget.style.background="rgba(255,255,255,.05)";}}
                      onMouseLeave={e=>{if(!isActive)e.currentTarget.style.background="transparent";}}>
                      <span style={{whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{cl}</span>
                      <span style={{fontSize:9,fontWeight:700,color:isActive?"#4ade80":"rgba(255,255,255,.3)",flexShrink:0,marginLeft:4}}>{ef}</span>
                    </div>
                  );
                })}
              </div>
            ))}
          </aside>
        )}

        {/* ── Contenu principal ────────────────────────────────── */}
        <main style={{flex:1,padding: isMobile?12:20,overflowY:"auto",minHeight:0}}>

          {/* ══ VUE REGISTRE ══════════════════════════════════════ */}
          {vue==="registre" && (
            <div style={{display:"flex",flexDirection:"column",gap:14}}>

              {/* Header classe */}
              <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:"16px 20px",display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:16}}>
                <div>
                  <h2 style={{fontSize:18,fontWeight:800,color:C.txt,margin:0}}>{selClasse}</h2>
                  <div style={{fontSize:12,color:C.txtMuted,marginTop:4}}>
                    Lycée de Kakatare · Maroua · 2025–2026
                    {niveau && <span style={{marginLeft:8,padding:"1px 8px",background:C.greenPale,border:`1px solid ${C.greenBorder}`,borderRadius:20,fontSize:10,fontWeight:700,color:C.green}}>{niveau}</span>}
                  </div>
                  {prog && (
                    <div style={{marginTop:8,display:"flex",gap:8,flexWrap:"wrap"}}>
                      <span style={{fontSize:11,padding:"2px 8px",borderRadius:20,background:C.greenPale,color:C.green,fontWeight:600}}>Coeff. {prog.coeff}</span>
                      <span style={{fontSize:11,padding:"2px 8px",borderRadius:20,background:C.bluePale,color:C.blue,fontWeight:600}}>T1: {prog.lp_t1} leçons</span>
                      <span style={{fontSize:11,padding:"2px 8px",borderRadius:20,background:"#fffbeb",color:C.amber,fontWeight:600}}>T2: {prog.lp_t2} leçons</span>
                      <span style={{fontSize:11,padding:"2px 8px",borderRadius:20,background:C.purplePale,color:C.purple,fontWeight:600}}>T3: {prog.lp_t3} leçons</span>
                    </div>
                  )}
                </div>
                {/* KPIs classe */}
                <div style={{display:"flex",gap:10,flexShrink:0}}>
                  {[
                    {v:eleves.length,l:"Total", col:C.txt, bg:"#f8fafc"},
                    {v:filles,       l:"Filles", col:C.pink, bg:C.pinkPale},
                    {v:garcons,      l:"Garçons",col:C.blue, bg:C.bluePale},
                  ].map((k,i)=>(
                    <div key={i} style={{textAlign:"center",padding:"8px 14px",background:k.bg,borderRadius:9,border:`1px solid ${C.border}`}}>
                      <div style={{fontSize:22,fontWeight:900,color:k.col}}>{k.v}</div>
                      <div style={{fontSize:10,color:C.txtMuted,fontWeight:600}}>{k.l}</div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Barre genre */}
              <div style={{background:C.white,borderRadius:10,border:`1px solid ${C.border}`,padding:"10px 14px"}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                  <span style={{fontSize:11,color:C.pink,fontWeight:700}}>♀ {filles} filles ({eleves.length?Math.min(100, Math.round(filles/eleves.length*100)):0}%)</span>
                  <span style={{fontSize:11,color:C.blue,fontWeight:700}}>♂ {garcons} garçons ({eleves.length?Math.min(100, Math.round(garcons/eleves.length*100)):0}%)</span>
                </div>
                <div style={{height:8,borderRadius:4,overflow:"hidden",background:"#e2e8f0",display:"flex"}}>
                  <div style={{width:`${eleves.length?filles/eleves.length*100:50}%`,background:C.pink,transition:"width .5s"}}/>
                  <div style={{flex:1,background:C.blue}}/>
                </div>
              </div>

              {/* Barre d'outils */}
              <div style={{display:"flex",gap:9,alignItems:"center",flexWrap:"wrap"}}>
                {/* Recherche */}
                <div style={{position:"relative",flex:1,minWidth:180}}>
                  <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",fontSize:13}}>🔍</span>
                  <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Rechercher un élève…"
                    style={{width:"100%",padding:"8px 12px 8px 32px",border:`1px solid ${C.border}`,borderRadius:9,fontSize:12,color:C.txt,background:"#f8fafc",outline:"none"}}
                    onFocus={e=>{e.target.style.borderColor=C.green;e.target.style.background=C.white;}}
                    onBlur={e=>{e.target.style.borderColor=C.border;e.target.style.background="#f8fafc";}}/>
                </div>
                {/* Filtre genre */}
                <div style={{display:"flex",gap:4}}>
                  {[["all","Tous"],["F","♀ Filles"],["M","♂ Garçons"]].map(([val,lab])=>(
                    <button key={val} onClick={()=>setFiltreGenre(val)}
                      style={{padding:"7px 11px",borderRadius:8,fontSize:11,fontWeight:700,cursor:"pointer",border:`1.5px solid ${filtreGenre===val?(val==="F"?C.pink:val==="M"?C.blue:C.green):C.border}`,background:filtreGenre===val?(val==="F"?C.pinkPale:val==="M"?C.bluePale:C.greenPale):C.white,color:filtreGenre===val?(val==="F"?C.pink:val==="M"?C.blue:C.green):C.txtMuted}}>
                      {lab}
                    </button>
                  ))}
                </div>
                {/* Bouton Ajouter */}
                <button onClick={()=>{setModal("ajout");setNewNom("");setNewGenre("M");}}
                  style={{padding:"7px 14px",background:`linear-gradient(135deg,${C.greenDark},${C.green})`,color:"#fff",border:"none",borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",gap:6}}>
                  ➕ Ajouter un élève
                </button>
                {/* Impression */}
                <button onClick={()=>imprimerListeClasse(selClasse, eleves)}
                  style={{padding:"7px 14px",background:C.white,color:C.txt,border:`1px solid ${C.border}`,borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer"}}>
                  🖨️ Imprimer
                </button>
                <span style={{fontSize:11,color:C.txtMuted}}>{elevesFiltres.length}/{eleves.length}</span>
              </div>

              {/* Table des élèves */}
              <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`}}><div style={{overflowX:"auto",WebkitOverflowScrolling:"touch"}}><table style={{minWidth:480,width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead>
                    <tr style={{background:"#f8fafc",borderBottom:`1px solid ${C.border}`}}>
                      <th style={{padding:"9px 12px",textAlign:"center",fontSize:10,fontWeight:700,color:C.txtMuted,width:48}}>N°</th>
                      <th style={{padding:"9px 12px",textAlign:"left",fontSize:10,fontWeight:700,color:C.txtMuted}}>NOM ET PRÉNOM(S)</th>
                      <th style={{padding:"9px 12px",textAlign:"center",fontSize:10,fontWeight:700,color:C.txtMuted,width:80}}>GENRE</th>
                      <th style={{padding:"9px 12px",textAlign:"center",fontSize:10,fontWeight:700,color:C.txtMuted,width:80}}>ACTION</th>
                    </tr>
                  </thead>
                  <tbody>
                    {elevesFiltres.length===0 ? (
                      <tr><td colSpan={4} style={{padding:"32px",textAlign:"center",color:C.txtLight}}>
                        <div style={{fontSize:28,marginBottom:6}}>🔍</div>Aucun élève trouvé
                      </td></tr>
                    ) : elevesFiltres.map((e,i)=>(
                      <tr key={e.id}
                        style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?C.white:"#fafafa",transition:"background .1s"}}
                        onMouseEnter={ev=>ev.currentTarget.style.background=C.greenPaleA60}
                        onMouseLeave={ev=>ev.currentTarget.style.background=i%2===0?C.white:"#fafafa"}>
                        <td style={{padding:"10px 12px",textAlign:"center",fontSize:11,fontWeight:700,color:C.txtLight}}>{i+1}</td>
                        <td style={{padding:"10px 12px"}}>
                          <div style={{display:"flex",alignItems:"center",gap:9}}>
                            <div style={{width:28,height:28,borderRadius:"50%",background:e.g==="F"?C.pinkPale:C.bluePale,border:`1.5px solid ${e.g==="F"?C.pink:C.blue}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:800,color:e.g==="F"?C.pink:C.blue,flexShrink:0}}>
                              {e.nom.split(" ")[0]?.[0]||"?"}{e.nom.split(" ").slice(-1)[0]?.[0]||""}
                            </div>
                            <span style={{fontWeight:700,color:C.txt}}>{e.nom}</span>
                          </div>
                        </td>
                        <td style={{padding:"10px 12px",textAlign:"center"}}>
                          <span style={{display:"inline-block",padding:"3px 10px",borderRadius:20,fontSize:10,fontWeight:800,background:e.g==="F"?C.pinkPale:C.bluePale,color:e.g==="F"?C.pink:C.blue,border:`1px solid ${e.g==="F"?C.pinkA40:C.blueA40}`}}>
                            {e.g==="F"?"♀ F":"♂ M"}
                          </span>
                        </td>
                        <td style={{padding:"10px 12px",textAlign:"center"}}>
                          <button onClick={()=>setConfirmRetrait(e)}
                            style={{padding:"4px 10px",background:C.redPale,border:`1px solid ${C.redBorder}`,borderRadius:7,color:C.red,fontSize:11,fontWeight:700,cursor:"pointer"}}>
                            ✕ Retirer
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>
            </div>
          )}

          {/* ══ VUE STATISTIQUES ══════════════════════════════════ */}
          {vue==="stats" && (
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              <h2 style={{fontSize:16,fontWeight:800,color:C.txt}}>📊 Statistiques — 2025–2026</h2>
              <div style={{display:"grid",gridTemplateColumns: isMobile?"repeat(2,1fr)":"repeat(4,1fr)",gap:12}}>
                {[
                  {v:TOTAL_LOC,  l:"Total élèves", col:C.green,  bg:C.greenPale,  emoji:"🎓"},
                  {v:FILLES_LOC, l:"Filles",        col:C.pink,   bg:C.pinkPale,   emoji:"♀"},
                  {v:GARCONS_LOC,l:"Garçons",       col:C.blue,   bg:C.bluePale,   emoji:"♂"},
                  {v:ALL_CLS.length,l:"Classes",   col:C.purple, bg:C.purplePale, emoji:"📚"},
                ].map((k,i)=>(
                  <div key={i} style={{background:k.bg,borderRadius:11,border:`1px solid ${C.border}`,padding:"16px",textAlign:"center"}}>
                    <div style={{fontSize:22}}>{k.emoji}</div>
                    <div style={{fontSize:30,fontWeight:900,color:k.col,marginTop:4}}>{k.v}</div>
                    <div style={{fontSize:11,color:C.txtMuted,fontWeight:600}}>{k.l}</div>
                  </div>
                ))}
              </div>
              {/* Tableau par groupe */}
              <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,overflow:"hidden"}}>
                <div style={{padding:"12px 16px",borderBottom:`1px solid ${C.border}`,fontSize:13,fontWeight:700,color:C.txt}}>Répartition par niveau</div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead>
                    <tr style={{background:"#0f1f14",color:"#fff"}}>
                      {["Niveau","Classes","Total","Filles","Garçons","Ratio ♀"].map((h,i)=>(
                        <th key={i} style={{padding:"9px 12px",textAlign:i===0?"left":"center",fontSize:10,fontWeight:700}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {GROUPES_ORDRE.map((gr,idx)=>{
                      const tot  = gr.classes.reduce((s,cl)=>s+(localDB[cl]||[]).length,0);
                      const fll  = gr.classes.reduce((s,cl)=>s+(localDB[cl]||[]).filter(e=>e.g==="F").length,0);
                      const grc  = tot-fll;
                      const rf   = tot>0?Math.min(100, Math.round(fll/tot*100)):0;
                      return tot===0?null:(
                        <tr key={gr.label} style={{borderBottom:`1px solid ${C.border}`,background:idx%2===0?C.white:"#f8fafc"}}>
                          <td style={{padding:"10px 12px",fontWeight:800,color:C.txt}}>{gr.label}</td>
                          <td style={{padding:"10px 12px",textAlign:"center",color:C.txtMuted}}>{gr.classes.filter(cl=>(localDB[cl]||[]).length>0).length}</td>
                          <td style={{padding:"10px 12px",textAlign:"center",fontWeight:800,fontSize:14,color:C.txt}}>{tot}</td>
                          <td style={{padding:"10px 12px",textAlign:"center",fontWeight:700,color:C.pink}}>{fll}</td>
                          <td style={{padding:"10px 12px",textAlign:"center",fontWeight:700,color:C.blue}}>{grc}</td>
                          <td style={{padding:"10px 12px",textAlign:"center"}}>
                            <div style={{display:"flex",alignItems:"center",gap:6}}>
                              <div style={{height:5,borderRadius:3,overflow:"hidden",background:"#e2e8f0",display:"flex",flex:1}}>
                                <div style={{width:`${rf}%`,background:C.pink}}/><div style={{flex:1,background:C.blue}}/>
                              </div>
                              <span style={{fontSize:10,color:C.txtMuted}}>{rf}%</span>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                    <tr style={{background:"#0f1f14",color:"#fff",fontWeight:800}}>
                      <td style={{padding:"10px 12px"}}>TOTAL</td>
                      <td style={{padding:"10px 12px",textAlign:"center"}}>{ALL_CLS.length}</td>
                      <td style={{padding:"10px 12px",textAlign:"center",fontSize:15,color:C.greenLight}}>{TOTAL_LOC}</td>
                      <td style={{padding:"10px 12px",textAlign:"center",color:"#f9a8d4"}}>{FILLES_LOC}</td>
                      <td style={{padding:"10px 12px",textAlign:"center",color:"#93c5fd"}}>{GARCONS_LOC}</td>
                      <td style={{padding:"10px 12px",textAlign:"center",color:"rgba(255,255,255,.6)",fontSize:11}}>{Math.min(100, Math.round(FILLES_LOC/TOTAL_LOC*100))}% filles</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ══ VUE PROGRAMMES ════════════════════════════════════ */}
          {vue==="programmes" && (
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              <h2 style={{fontSize:16,fontWeight:800,color:C.txt}}>📚 Référentiel programmes SVTEEHB</h2>
              <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:14}}>
                {Object.entries(PROGRAMME_SVTEEHB).map(([niv,prog])=>{
                  const total = prog.lp_t1+prog.lp_t2+prog.lp_t3;
                  return (
                    <div key={niv} style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:"16px 18px"}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
                        <div>
                          <h3 style={{fontSize:14,fontWeight:800,color:C.txt,margin:0}}>{niv}</h3>
                          <div style={{fontSize:11,color:C.txtMuted,marginTop:2}}>{total} leçons · {prog.tp} TP · Coeff. {prog.coeff}</div>
                        </div>
                        <div style={{padding:"4px 10px",background:C.greenPale,border:`1px solid ${C.greenBorder}`,borderRadius:20,fontSize:12,fontWeight:800,color:C.green}}>Coeff. {prog.coeff}</div>
                      </div>
                      <div style={{height:10,borderRadius:5,overflow:"hidden",display:"flex",gap:1,marginBottom:10}}>
                        <div style={{flex:prog.lp_t1,background:C.green,minWidth:4}}/>
                        <div style={{flex:prog.lp_t2,background:C.amber,minWidth:4}}/>
                        <div style={{flex:prog.lp_t3,background:C.purple,minWidth:4}}/>
                      </div>
                      <div style={{display:"grid",gridTemplateColumns: isMobile?"repeat(2,1fr)":"repeat(3,1fr)",gap:8}}>
                        {[{t:"T1",v:prog.lp_t1,col:C.green,bg:C.greenPale},{t:"T2",v:prog.lp_t2,col:C.amber,bg:C.amberPale},{t:"T3",v:prog.lp_t3,col:C.purple,bg:C.purplePale}].map((d,i)=>(
                          <div key={i} style={{background:d.bg,borderRadius:8,padding:"8px",textAlign:"center"}}>
                            <div style={{fontSize:10,color:d.col,fontWeight:700,marginBottom:2}}>{d.t}</div>
                            <div style={{fontSize:18,fontWeight:900,color:d.col}}>{d.v}</div>
                            <div style={{fontSize:9,color:C.txtMuted}}>leçons</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </main>
      </div>

      {/* ══ MODAL AJOUT ═══════════════════════════════════════════ */}
      {modal==="ajout" && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,animation:"fadeIn .2s"}}>
          <div style={{background:C.white,borderRadius:16,padding:"28px 32px",width:"100%",maxWidth:420,boxShadow:"0 24px 80px rgba(0,0,0,.25)"}}>
            <h3 style={{fontSize:16,fontWeight:800,color:C.txt,marginBottom:4}}>➕ Ajouter un élève</h3>
            <p style={{fontSize:12,color:C.txtMuted,marginBottom:20}}>Classe : <strong>{selClasse}</strong></p>

            <div style={{marginBottom:14}}>
              <label style={{fontSize:11,fontWeight:600,color:C.txtMuted,display:"block",marginBottom:5}}>Nom et Prénom(s) <span style={{color:C.red}}>*</span></label>
              <input
                autoFocus
                value={newNom}
                onChange={e=>setNewNom(e.target.value.toUpperCase())}
                onKeyDown={e=>{ if(e.key==="Enter") ajouterEleve(); }}
                placeholder="ex. AMADOU BOUKAR"
                style={{width:"100%",padding:"10px 14px",border:`1.5px solid ${C.border}`,borderRadius:9,fontSize:13,color:C.txt,outline:"none"}}
                onFocus={e=>e.target.style.borderColor=C.green}
                onBlur={e=>e.target.style.borderColor=C.border}/>
            </div>

            <div style={{marginBottom:20}}>
              <label style={{fontSize:11,fontWeight:600,color:C.txtMuted,display:"block",marginBottom:5}}>Genre</label>
              <div style={{display:"flex",gap:10}}>
                {[["M","♂ Masculin",C.blue,C.bluePale],["F","♀ Féminin",C.pink,C.pinkPale]].map(([val,lab,col,bg])=>(
                  <button key={val} onClick={()=>setNewGenre(val)}
                    style={{flex:1,padding:"10px",borderRadius:9,border:`2px solid ${newGenre===val?col:C.border}`,background:newGenre===val?bg:C.white,color:newGenre===val?col:C.txtMuted,fontSize:13,fontWeight:700,cursor:"pointer"}}>
                    {lab}
                  </button>
                ))}
              </div>
            </div>

            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setModal(null)}
                style={{flex:1,padding:"11px",borderRadius:9,border:`1px solid ${C.border}`,background:C.white,color:C.txtMuted,fontSize:13,fontWeight:700,cursor:"pointer"}}>
                Annuler
              </button>
              <button onClick={ajouterEleve}
                style={{flex:2,padding:"11px",borderRadius:9,border:"none",background:`linear-gradient(135deg,${C.greenDark},${C.green})`,color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                ✓ Confirmer l'ajout
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ══ MODAL CONFIRMATION RETRAIT ════════════════════════════ */}
      {confirmRetrait && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,animation:"fadeIn .2s"}}>
          <div style={{background:C.white,borderRadius:16,padding:"28px 32px",width:"100%",maxWidth:400,boxShadow:"0 24px 80px rgba(0,0,0,.25)"}}>
            <div style={{fontSize:32,textAlign:"center",marginBottom:12}}>⚠️</div>
            <h3 style={{fontSize:15,fontWeight:800,color:C.txt,textAlign:"center",marginBottom:8}}>Retirer cet élève ?</h3>
            <p style={{fontSize:13,color:C.txtMuted,textAlign:"center",marginBottom:6}}>
              <strong style={{color:C.txt}}>{confirmRetrait.nom}</strong>
            </p>
            <p style={{fontSize:12,color:C.txtMuted,textAlign:"center",marginBottom:24}}>
              sera retiré(e) de <strong>{selClasse}</strong>. Cette action est réversible uniquement par rechargement de la page.
            </p>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setConfirmRetrait(null)}
                style={{flex:1,padding:"11px",borderRadius:9,border:`1px solid ${C.border}`,background:C.white,color:C.txtMuted,fontSize:13,fontWeight:700,cursor:"pointer"}}>
                Annuler
              </button>
              <button onClick={()=>retirerEleve(confirmRetrait)}
                style={{flex:1,padding:"11px",borderRadius:9,border:"none",background:C.red,color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                ✕ Confirmer le retrait
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


// ══════════════════════════════════════════════════════════════════════
// SECTION ENSEIGNANTS — Interface Animatrice Pédagogique
// Vue : liste globale → détail par enseignant → détail par classe
// ══════════════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════════
// GESTION ENSEIGNANTS — Ajout · Suppression · Données fictives
// Intégré dans EnseignantsPage comme onglet "Gérer"
// ══════════════════════════════════════════════════════════════════════

function EnseignantsPage() {
  const {data, setData, showToast, refreshData} = useApp();
  const {isMobile} = useDevice();
  const [onglet,  setOnglet]  = useState("liste");   // "liste" | "gerer"
  const [viewEns, setViewEns] = useState(null);
  const [viewCl,  setViewCl]  = useState(null);
  const [refreshing, setRefreshing] = useState(false);

  if (!data) return (
    <div style={{padding:"40px",textAlign:"center",color:C.txtMuted}}>
      <Spinner size={28} color={C.green}/><br/><br/>Chargement…
    </div>
  );

  // Niveau 3 — Leçons d'une classe
  if (viewEns && viewCl) {
    return <EnsClasLecons ens={viewEns} cl={viewCl}
      data={data} setData={setData} showToast={showToast}
      onBack={()=>setViewCl(null)}/>;
  }

  // Niveau 2 — Détail enseignant
  if (viewEns) {
    return <EnsDetail ens={viewEns}
      data={data} setData={setData} showToast={showToast}
      onBack={()=>setViewEns(null)}
      onViewClass={(cl)=>setViewCl(cl)}/>;
  }

  return (
    <div style={{display:"flex", flexDirection:"column", flex:1}}>

      {/* ── Onglets ── */}
      <div style={{background:C.white, borderBottom:`1px solid ${C.border}`, padding:"0 20px", display:"flex", alignItems:"center", justifyContent:"space-between", gap:0, flexShrink:0}}>
        <div style={{display:"flex"}}>
          {[
            {id:"liste", label:"👥 Enseignants",     sub:"Vue globale"},
            {id:"gerer", label:"⚙️ Gérer",            sub:"Ajouter · Supprimer"},
            {id:"clean", label:"🧹 Données fictives", sub:"Nettoyer la base"},
          ].map(o=>(
            <button key={o.id} onClick={()=>setOnglet(o.id)}
              style={{padding:"12px 18px", border:"none", borderBottom:`3px solid ${onglet===o.id?C.green:"transparent"}`, background:"transparent", cursor:"pointer", fontFamily:"inherit", textAlign:"left"}}>
              <div style={{fontSize:12, fontWeight:700, color:onglet===o.id?C.green:C.txt}}>{o.label}</div>
              <div style={{fontSize:10, color:C.txtMuted, marginTop:1}}>{o.sub}</div>
            </button>
          ))}
        </div>
        <button onClick={async()=>{
            setRefreshing(true);
            await refreshData();
            setRefreshing(false);
          }}
          disabled={refreshing}
          style={{display:"flex", alignItems:"center", gap:6, padding: isMobile?"8px":"7px 14px", borderRadius:8, border:`1px solid ${C.border}`, background:C.white, color:C.txtMuted, fontSize:11.5, fontWeight:700, cursor:refreshing?"not-allowed":"pointer", fontFamily:"inherit", flexShrink:0}}>
          {refreshing ? <Spinner size={12} color={C.txtMuted}/> : "🔄"} {!isMobile && "Actualiser"}
        </button>
      </div>

      {/* ── Contenu ── */}
      <div style={{flex:1, overflowY:"auto"}}>
        {onglet==="liste" && <EnsListe data={data} onSelect={ens=>{setViewEns(ens);}}/>}
        {onglet==="gerer" && <EnsGerer data={data} setData={setData} showToast={showToast}/>}
        {onglet==="clean" && <EnsClean data={data} setData={setData} showToast={showToast}/>}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════
// ONGLET GÉRER — Ajouter / Modifier / Supprimer un enseignant
// ══════════════════════════════════════════════════════════════════════
// ── Détecteur de conflits : une classe assignée à plusieurs enseignants à la fois ──
function ConflitsClassesPanel({ enseignants, data }) {
  const { showToast, refreshData } = useApp();
  const [resolving, setResolving] = useState(null); // nom de la classe ou id enseignant en cours de résolution

  const conflits = useMemo(() => {
    const parClasse = {};
    enseignants.forEach(ens => {
      (ens.classes||[]).forEach(cl => { (parClasse[cl] ||= []).push(ens); });
    });
    return Object.entries(parClasse).filter(([, list]) => list.length > 1);
  }, [enseignants]);

  // Incohérences par enseignant : ses "classes assignées" ne correspondent pas à ce qui
  // est réellement dans son EDT (cas vécu : Essamba avait "1ère A4 ARB" en classes
  // assignées mais "1ère L2 ARA/ITA/ESP" dans son EDT_REEL d'origine).
  const incoherences = useMemo(() => {
    const result = [];
    enseignants.forEach(ens => {
      const rt = buildEdtRuntime(data?.exceptions||{}, data?.edtBase||{})[ens.id] || {};
      const classesEdt = new Set();
      JKEYS.forEach(jk => (rt[jk]||[]).forEach(cl => { if (cl) classesEdt.add(cl); }));
      const classesAssignees = new Set(ens.classes||[]);
      const manquantes = [...classesEdt].filter(cl => !classesAssignees.has(cl));
      const enTrop = [...classesAssignees].filter(cl => !classesEdt.has(cl));
      if (manquantes.length>0 || enTrop.length>0) result.push({ ens, manquantes, enTrop, classesEdt:[...classesEdt] });
    });
    return result;
  }, [enseignants, data?.exceptions]);

  if (conflits.length === 0 && incoherences.length === 0) return null;

  const resoudreConflit = async (classe, ensGardeId) => {
    setResolving(classe);
    const aNettoyer = enseignants.filter(e => (e.classes||[]).includes(classe) && e.id !== ensGardeId);
    let ok = true;
    for (const ens of aNettoyer) {
      const nouvellesClasses = (ens.classes||[]).filter(cl => cl !== classe);
      const rt = buildEdtRuntime(data?.exceptions||{}, data?.edtBase||{})[ens.id] || {};
      const slots = [];
      JKEYS.forEach(jk => { (rt[jk]||[]).forEach((lbl,hi)=>{ if (lbl===classe) slots.push(`${jk}-${hi}`); }); });
      if (slots.length > 0) {
        await sb.rpc("admin_set_edt_slots", { p_ens_id: ens.id, p_slots: slots.map(slot=>({ slot, lbl:"" })) });
      }
      // Cascade identique à ModalEnsForm : la classe transférée ne doit pas laisser
      // de progression (ni de suivi digital) fantôme rattachée au perdant du conflit.
      // Nettoyée seulement si le retrait de classe a réellement réussi (r).
      const r = await sb.rpc("admin_set_teacher_classes", { p_id: ens.id, p_classes: nouvellesClasses });
      if (r) {
        await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: ens.id, p_classe: classe });
        await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: ens.id, p_classe: classe+"||dig" });
      } else {
        ok = false;
      }
    }
    setResolving(null);
    const nomGarde = getNomCourt(enseignants.find(e=>e.id===ensGardeId)?.nom);
    showToast(ok ? `✓ "${classe}" attribuée uniquement à ${nomGarde}` : "⚠ Échec partiel — vérifiez la connexion", ok);
    await refreshData?.();
  };

  const reconcilier = async (item) => {
    const { ens, manquantes, classesEdt } = item;
    setResolving(ens.id);
    // Vérifier qu'aucune des classes manquantes n'est déjà chez un autre enseignant
    // (évite de créer un nouveau conflit en réconciliant celui-ci)
    const dejaAilleurs = manquantes.filter(cl => enseignants.some(autre => autre.id!==ens.id && (autre.classes||[]).includes(cl)));
    if (dejaAilleurs.length > 0) {
      showToast(`⚠ Réconciliation bloquée : ${dejaAilleurs.join(", ")} déjà chez un autre enseignant — résous d'abord ce conflit ci-dessus`, false);
      setResolving(null);
      return;
    }
    const ok = await sb.rpc("admin_set_teacher_classes", { p_id: ens.id, p_classes: classesEdt });
    setResolving(null);
    showToast(ok ? `✓ Classes de ${getNomCourt(ens.nom)} alignées sur son EDT réel` : "⚠ Échec — vérifiez la connexion", ok);
    await refreshData?.();
  };

  // Calcul stats pour la fiche
  const totalFait = (ens.classes||[]).reduce((s,cl)=>{
    const fait = ((data?.prog||{})[ens.id+"||"+cl]||[]).length;
    return s+fait;
  },0);
  const totalRef = (ens.classes||[]).reduce((s,cl)=>{
    const code = resolveProgCode(cl);
    return s+(code?PROG_META[code]?.lpRef||0:0);
  },0);
  const tauxCouv = totalRef>0?Math.min(100,Math.round(totalFait/totalRef*100)):0;
  const nbAbsences = Object.entries(data?.absences||{}).filter(([k])=>k.startsWith(ens.id+"||")).length;
  const nbEpreuves = (data?.epreuves||[]).filter(e=>e.ens_id===ens.id).length;
  const nbEpAttente = (data?.epreuves||[]).filter(e=>e.ens_id===ens.id&&e.statut==="attente").length;
  const deptNom = DEPARTEMENTS_LIST.find(d=>d.id===ens.departement_id)?.nom||"—";
  return (
    <div style={{display:"flex", flexDirection:"column", gap:12}}>

      {/* ── Fiche profil ─────────────────────────────────────────── */}
      <div style={{background:"#0B3D20",borderRadius:16,padding:"20px 24px",display:"flex",gap:18,alignItems:"flex-start",flexWrap:"wrap"}}>
        {/* Photo */}
        <div style={{flexShrink:0}}>
          <Avatar ens={ens} size={72} fontSize={22}/>
        </div>
        {/* Infos */}
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:16,fontWeight:800,color:"#fff",marginBottom:2}}>{ens.nom}</div>
          <div style={{fontSize:11,color:"rgba(255,255,255,.5)",marginBottom:10}}>
            {deptNom} · {(ens.classes||[]).join(", ")||"Aucune classe"}
          </div>
          {/* KPIs */}
          <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
            {[
              {label:"Couverture prog.",value:tauxCouv+"%",color:tauxCouv>=75?"#4ade80":tauxCouv>=50?"#fbbf24":"#f87171"},
              {label:"Leçons faites",value:totalFait+"/"+totalRef,color:"rgba(255,255,255,.8)"},
              {label:"Séances saisies",value:nbAbsences,color:"rgba(255,255,255,.8)"},
              {label:"Épreuves",value:nbEpreuves+(nbEpAttente>0?" ("+nbEpAttente+" en attente)":""),color:nbEpAttente>0?"#fbbf24":"rgba(255,255,255,.8)"},
              {label:"Classes",value:(ens.classes||[]).length,color:"rgba(255,255,255,.8)"},
            ].map((k,i)=>(
              <div key={i} style={{background:"rgba(255,255,255,.08)",borderRadius:10,padding:"8px 14px",minWidth:90,textAlign:"center"}}>
                <div style={{fontSize:16,fontWeight:800,color:k.color}}>{k.value}</div>
                <div style={{fontSize:9,color:"rgba(255,255,255,.4)",marginTop:2,fontWeight:600}}>{k.label}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {conflits.length > 0 && (
        <div style={{background:C.redPale, border:`1px solid ${C.redBorder}`, borderRadius:12, padding:"16px 18px", display:"flex", flexDirection:"column", gap:14}}>
          <div>
            <h3 style={{fontSize:13, fontWeight:800, color:C.red, margin:0}}>⚠️ {conflits.length} classe{conflits.length>1?"s":""} assignée{conflits.length>1?"s":""} à plusieurs enseignants</h3>
            <p style={{fontSize:11.5, color:C.txtMuted, margin:"4px 0 0", lineHeight:1.5}}>Choisis qui garde chaque classe — elle sera retirée automatiquement des autres (classes assignées + leur EDT).</p>
          </div>
          {conflits.map(([classe, list]) => (
            <div key={classe} style={{background:C.white, borderRadius:9, padding:"12px 14px", display:"flex", flexDirection:"column", gap:8}}>
              <div style={{fontWeight:800, fontSize:13, color:C.txt}}>{classe}</div>
              <div style={{display:"flex", gap:8, flexWrap:"wrap"}}>
                {list.map(ens=>(
                  <button key={ens.id} disabled={resolving===classe} onClick={()=>resoudreConflit(classe, ens.id)}
                    style={{padding:"7px 14px", borderRadius:8, border:`1.5px solid ${ens.col||C.border}`, background:C.white, fontSize:12, fontWeight:700,
                      color:ens.col||C.txt, cursor:resolving===classe?"not-allowed":"pointer", fontFamily:"inherit", opacity:resolving===classe?.6:1}}>
                    Garder chez {getNomCourt(ens.nom)}
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}

      {incoherences.length > 0 && (
        <div style={{background:C.amberPale, border:`1px solid #fde68a`, borderRadius:12, padding:"16px 18px", display:"flex", flexDirection:"column", gap:14}}>
          <div>
            <h3 style={{fontSize:13, fontWeight:800, color:"#92400e", margin:0}}>🔶 {incoherences.length} enseignant{incoherences.length>1?"s":""} dont les classes assignées ne correspondent pas à l'EDT</h3>
            <p style={{fontSize:11.5, color:C.txtMuted, margin:"4px 0 0", lineHeight:1.5}}>Ses "classes assignées" et le contenu réel de son emploi du temps ont divergé — probablement un ancien réglage jamais mis à jour.</p>
          </div>
          {incoherences.map(item => (
            <div key={item.ens.id} style={{background:C.white, borderRadius:9, padding:"12px 14px", display:"flex", flexDirection:"column", gap:8}}>
              <div style={{fontWeight:800, fontSize:13, color:C.txt}}>{getNomCourt(item.ens.nom)}</div>
              {item.manquantes.length>0 && <div style={{fontSize:11.5, color:C.txtMuted}}>Dans son EDT mais pas dans ses classes assignées : <strong>{item.manquantes.join(", ")}</strong></div>}
              {item.enTrop.length>0 && <div style={{fontSize:11.5, color:C.txtMuted}}>Dans ses classes assignées mais absent de son EDT : <strong>{item.enTrop.join(", ")}</strong></div>}
              <button disabled={resolving===item.ens.id} onClick={()=>reconcilier(item)}
                style={{alignSelf:"flex-start", padding:"7px 14px", borderRadius:8, border:"1.5px solid #92400e", background:C.white, fontSize:12, fontWeight:700,
                  color:"#92400e", cursor:resolving===item.ens.id?"not-allowed":"pointer", fontFamily:"inherit", opacity:resolving===item.ens.id?.6:1}}>
                ✓ Aligner ses classes assignées sur son EDT réel
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function EnsGerer({ data, setData, showToast }) {
  const [modal,    setModal]    = useState(null); // null | "ajout" | {ens} pour édition
  const [confirm,  setConfirm]  = useState(null); // enseignant à supprimer
  const [saving,   setSaving]   = useState(false);

  // Source enseignants
  const supabaseEns = Object.values(data?.users||{}).filter(u=>u.role!=="proviseur");
  const enseignants = supabaseEns.length > 0
    ? supabaseEns.map(u=>({...u, col:u.col||getColor(u.id), ini:u.ini||getIni(u.nom), classes:(u.classes||[]).length>0?u.classes:(ENS_CLASSES_REF[u.id]||[])}))
    : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant").map(a=>({...a, col:getColor(a.id), ini:getIni(a.nom), classes:ENS_CLASSES_REF[a.id]||[]}));

  // ── Supprimer un enseignant ──
  const supprimerEnseignant = async(ens) => {
    setSaving(true);
    // 1. Supprimer de Supabase utilisateurs
    const ok1 = await sb.rpc("admin_delete_teacher", { p_id: ens.id });
    // 2. Supprimer ses prog_suivi
    const ok2 = await sb.rpc("admin_delete_prog_by_teacher", { p_ens_id: ens.id });
    // 3. Supprimer ses epreuves
    const ok3 = await sb.rpc("admin_delete_epreuves_by_teacher", { p_ens_id: ens.id });
    // 4. Supprimer ses exceptions d'EDT (sinon orphelines — pourraient ressurgir si l'id est réutilisé)
    await sb.rpc("admin_delete_edt_slots_by_teacher", { p_ens_id: ens.id });
    // 5. Supprimer ses absences et ses notes
    await sb.rpc("admin_delete_absences_by_teacher", { p_ens_id: ens.id });
    // 6. Mettre à jour le contexte local
    setData(prev => {
      const newUsers = {...(prev.users||{})};
      delete newUsers[ens.id];
      const newProg  = {...(prev.prog||{})};
      Object.keys(newProg).forEach(k=>{ if(k.startsWith(ens.id+"||")) delete newProg[k]; });
      const newEps   = (prev.epreuves||[]).filter(e=>e.ens_id!==ens.id);
      const newExc   = {...(prev.exceptions||{})};
      delete newExc[ens.id];
      return {...prev, users:newUsers, prog:newProg, epreuves:newEps, exceptions:newExc};
    });
    setSaving(false);
    setConfirm(null);
    if(ok1) showToast(`✓ ${ens.nom} supprimé(e)`);
    else     showToast("⚠ Erreur Supabase — compte local supprimé", false);
  };

  return (
    <div style={{padding:"20px", display:"flex", flexDirection:"column", gap:16}}>

      {/* Header + bouton ajouter */}
      <div style={{display:"flex", justifyContent:"space-between", alignItems:"center"}}>
        <div>
          <h2 style={{fontSize:16, fontWeight:800, color:C.txt, margin:0}}>⚙️ Gestion des enseignants</h2>
          <p style={{fontSize:12, color:C.txtMuted, margin:"4px 0 0"}}>Département SVTEEHB · Lycée de Kakatare</p>
        </div>
        <button onClick={()=>setModal("ajout")}
          style={{padding:"9px 18px", background:`linear-gradient(135deg,${C.greenDark},${C.green})`, color:"#fff", border:"none", borderRadius:10, fontSize:13, fontWeight:700, cursor:"pointer", fontFamily:"inherit", display:"flex", alignItems:"center", gap:7}}>
          ➕ Ajouter un enseignant
        </button>
      </div>

      <ConflitsClassesPanel enseignants={enseignants} data={data}/>

      {/* Liste enseignants */}
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, overflow:"hidden"}}>
        <div style={{padding:"12px 18px", borderBottom:`1px solid ${C.border}`, fontSize:12, fontWeight:700, color:C.txt}}>
          {enseignants.length} enseignant{enseignants.length>1?"s":""} dans le département
        </div>
        {enseignants.map((ens,i)=>(
          <div key={ens.id} style={{display:"flex", alignItems:"center", gap:12, padding:"14px 18px", borderBottom:i<enseignants.length-1?`1px solid ${C.border}`:"none"}}>
            {/* Avatar */}
            <Avatar ens={ens} size={42} fontSize={13}/>
            {/* Infos */}
            <div style={{flex:1, minWidth:0}}>
              <div style={{fontSize:14, fontWeight:700, color:C.txt}}>{ens.nom}</div>
              <div style={{fontSize:11, color:C.txtMuted, marginTop:2}}>
                id: <code style={{background:"#f1f5f9", padding:"1px 5px", borderRadius:4, fontSize:11}}>{ens.id}</code>
                &nbsp;·&nbsp; {(ens.classes||[]).length} classe{(ens.classes||[]).length>1?"s":""}
              </div>
              {(ens.classes||[]).length > 0 && (
                <div style={{display:"flex", flexWrap:"wrap", gap:4, marginTop:5}}>
                  {(ens.classes||[]).map(cl=>(
                    <span key={cl} style={{fontSize:9, padding:"1px 7px", borderRadius:20, background:ens.col+"18", color:ens.col, fontWeight:700, border:`1px solid ${ens.col}30`}}>{cl}</span>
                  ))}
                </div>
              )}
            </div>
            {/* Actions */}
            <div style={{display:"flex", gap:8, flexShrink:0}}>
              <button onClick={()=>setModal(ens)}
                style={{padding:"6px 12px", background:C.greenPale, border:`1px solid ${C.greenBorder}`, borderRadius:8, fontSize:11, fontWeight:700, cursor:"pointer", color:C.green, fontFamily:"inherit"}}>
                ✎ Modifier
              </button>
              <button onClick={()=>setConfirm(ens)}
                style={{padding:"6px 12px", background:C.redPale, border:`1px solid ${C.redBorder||"#fca5a5"}`, borderRadius:8, fontSize:11, fontWeight:700, cursor:"pointer", color:C.red, fontFamily:"inherit"}}>
                🗑 Supprimer
              </button>
            </div>
          </div>
        ))}
      </div>

      {/* Modal Ajout / Modification */}
      {modal && (
        <ModalEnsForm
          ens={modal==="ajout" ? null : modal}
          data={data} setData={setData}
          showToast={showToast}
          onClose={()=>setModal(null)}/>
      )}

      {/* Modal Confirmation Suppression */}
      {confirm && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}}>
          <div style={{background:C.white,borderRadius:16,padding:"28px 32px",width:"100%",maxWidth:400,boxShadow:"0 24px 80px rgba(0,0,0,.25)"}}>
            <div style={{fontSize:36,textAlign:"center",marginBottom:12}}>⚠️</div>
            <h3 style={{fontSize:15,fontWeight:800,color:C.txt,textAlign:"center",marginBottom:8}}>
              Supprimer cet enseignant ?
            </h3>
            <p style={{fontSize:13,color:C.txtMuted,textAlign:"center",marginBottom:4}}>
              <strong style={{color:C.txt}}>{confirm.nom}</strong>
            </p>
            <p style={{fontSize:12,color:C.txtMuted,textAlign:"center",marginBottom:24,lineHeight:1.5}}>
              Son compte, ses leçons cochées et ses épreuves seront définitivement supprimés de Supabase.
            </p>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setConfirm(null)} disabled={saving}
                style={{flex:1,padding:"11px",borderRadius:9,border:`1px solid ${C.border}`,background:C.white,color:C.txtMuted,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                Annuler
              </button>
              <button onClick={()=>supprimerEnseignant(confirm)} disabled={saving}
                style={{flex:1,padding:"11px",borderRadius:9,border:"none",background:C.red,color:"#fff",fontSize:13,fontWeight:700,cursor:saving?"not-allowed":"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
                {saving?<><Spinner size={14} color="#fff"/>Suppression…</>:"🗑 Confirmer"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Formulaire Ajout / Modification enseignant ────────────────────────
function ModalEnsForm({ ens, data, setData, showToast, onClose }) {
  const {isMobile} = useDevice();
  const {user} = useApp();
  const isEdit = !!ens;
  const [nom,     setNom]     = useState(ens?.nom||"");
  const [ensId,   setEnsId]   = useState(ens?.id||"");
  const [mdp,     setMdp]     = useState("");
  const [classes, setClasses] = useState(ens?.classes||[]);
  const [newCl,   setNewCl]   = useState("");
  const [saving,  setSaving]  = useState(false);
  const [err,     setErr]     = useState("");
  const [photoFile, setPhotoFile] = useState(null);   // nouveau fichier choisi (pas encore envoyé)
  const [photoPreview, setPhotoPreview] = useState(null); // aperçu local (URL objet)
  const [photoExistant, setPhotoExistant] = useState(ens?.photo||null); // chemin déjà en base
  const fileInputRef = useRef(null);

  const choisirPhoto = (file) => {
    if (!file) return;
    if (!file.type.startsWith("image/")) { showToast("⚠ Choisis une image (jpg, png…)", false); return; }
    if (file.size > 8*1024*1024) { showToast("⚠ Image trop lourde (max 8 Mo)", false); return; }
    setPhotoFile(file);
    setPhotoPreview(URL.createObjectURL(file));
  };

  const toId = (nom) => nom.toLowerCase()
    .normalize("NFD").replace(/[̀-ͯ]/g,"")
    .replace(/[^a-z0-9]/g,"").slice(0,12);

  useEffect(()=>{
    if(!isEdit && nom) setEnsId(toId(nom));
  },[nom, isEdit]);

  const ajouterClasse = () => {
    const cl = newCl.trim();
    if(cl && !classes.includes(cl)) {
      setClasses(prev=>[...prev, cl]);
      setNewCl("");
    }
  };

  const sauvegarder = async() => {
    if(!nom.trim()) return setErr("Le nom est requis.");
    if(!ensId.trim()) return setErr("L'identifiant est requis.");
    if(!isEdit && !mdp.trim()) return setErr("Le mot de passe est requis pour un nouvel enseignant.");
    if(!isEdit && mdp.trim().length < 6) return setErr("Le mot de passe doit contenir au moins 6 caractères.");

    setSaving(true); setErr("");

    let photoPath = photoExistant;
    if (photoFile) {
      try {
        const compressed = await resizeImageFile(photoFile, 300, 0.82);
        const path = `${ensId.trim()}_${Date.now()}.jpg`;
        const ok = await sb.uploadPhoto(path, compressed);
        if (ok) photoPath = path;
        else showToast("⚠ Photo non envoyée — le reste a été sauvegardé", false);
      } catch {
        showToast("⚠ Erreur lors du traitement de la photo", false);
      }
    }

    const userData = {
      id: ensId.trim(),
      nom: nom.trim().toUpperCase(),
      role: "enseignant",
      classes,
      photo: photoPath,
      // mdp retiré : plus jamais écrit en clair — voir admin_set_password() après l'upsert
    };

    // Classes retirées des "Classes assignées" lors de cette modification (vide si nouvel enseignant)
    const classesRetirees = isEdit ? (ens.classes||[]).filter(c => !classes.includes(c)) : [];
    // Classes ajoutées — à vérifier pour conflit avec un autre enseignant
    const classesAjouteesIci = isEdit ? classes.filter(c => !(ens.classes||[]).includes(c)) : classes;

    const currentDept = user?.departement_id ?? ens?.departement_id ?? null;
    const autresEns = Object.values(data?.users||{}).filter(u=>u.role!=="proviseur" && u.id!==ensId.trim() && (currentDept==null || (u.departement_id||1)===currentDept));
    const conflits = [];
    classesAjouteesIci.forEach(cl => {
      autresEns.forEach(autre => { if ((autre.classes||[]).includes(cl)) conflits.push({ classe:cl, autreEns:autre }); });
    });
    if (conflits.length > 0) {
      const liste = conflits.map(c => `• ${c.classe} (actuellement chez ${getNomCourt(c.autreEns.nom)})`).join("\n");
      const continuer = window.confirm(
        `${conflits.length>1?"Ces classes sont":"Cette classe est"} déjà assignée(s) à un autre enseignant :\n\n${liste}\n\n`+
        `Continuer va la/les retirer de l'autre enseignant (classes assignées + son EDT) pour la/les transférer ici.\n\nContinuer ?`
      );
      if (!continuer) { setSaving(false); return; }
    }

    const ok = await sb.rpc("admin_upsert_teacher", { p_id: userData.id, p_nom: userData.nom, p_classes: userData.classes, p_photo: userData.photo });

    // Si un mot de passe a été saisi (création ou réinitialisation), le définir via la RPC
    // sécurisée (hachage côté serveur — jamais plus écrit en clair depuis le client).
    let mdpOk = true;
    if (ok && mdp.trim()) {
      mdpOk = await sb.rpc("admin_set_password", { p_id: ensId.trim(), p_new_mdp: mdp.trim() });
    }

    if(ok) {
      let nouvellesExceptions = data?.exceptions || {};

      // Nettoyage chez les enseignants en conflit : retirer la classe transférée de leurs
      // classes assignées ET de tous leurs créneaux EDT correspondants.
      if (conflits.length > 0) {
        const parAutreEns = {};
        conflits.forEach(c => { (parAutreEns[c.autreEns.id] ||= []).push(c.classe); });
        for (const [autreId, classesAEnlever] of Object.entries(parAutreEns)) {
          const autre = autresEns.find(e=>e.id===autreId);
          const nouvellesClassesAutre = (autre?.classes||[]).filter(cl=>!classesAEnlever.includes(cl));
          const rtAutre = buildEdtRuntime(data?.exceptions||{}, data?.edtBase||{})[autreId] || {};
          const slotsAEffacer = [];
          JKEYS.forEach(jk => { (rtAutre[jk]||[]).forEach((lbl,hi)=>{ if (lbl && classesAEnlever.includes(lbl)) slotsAEffacer.push(`${jk}-${hi}`); }); });
          if (slotsAEffacer.length > 0) {
            await sb.rpc("admin_set_edt_slots", { p_ens_id: autreId, p_slots: slotsAEffacer.map(slot=>({ slot, lbl:"" })) });
            nouvellesExceptions = { ...nouvellesExceptions, [autreId]: { ...(nouvellesExceptions[autreId]||{}) } };
            slotsAEffacer.forEach(slot => { nouvellesExceptions[autreId][slot] = ""; });
          }
          await sb.rpc("admin_set_teacher_classes", { p_id: autreId, p_classes: nouvellesClassesAutre });
        }
      }

      // Cascade : effacer automatiquement les classes retirées de tous ses créneaux EDT,
      // pour éviter qu'une classe reste affichée chez deux enseignants à la fois.
      let progKeysASupprimer = [];
      if (classesRetirees.length > 0) {
        const rt = buildEdtRuntime(data?.exceptions||{}, data?.edtBase||{})[ens.id] || {};
        const slotsAEffacer = [];
        JKEYS.forEach(jk => {
          (rt[jk]||[]).forEach((lbl, hi) => {
            if (lbl && classesRetirees.includes(lbl)) slotsAEffacer.push(`${jk}-${hi}`);
          });
        });
        if (slotsAEffacer.length > 0) {
          await sb.rpc("admin_set_edt_slots", { p_ens_id: ens.id, p_slots: slotsAEffacer.map(slot=>({ slot, lbl:"" })) });
          nouvellesExceptions = { ...nouvellesExceptions, [ens.id]: { ...(nouvellesExceptions[ens.id]||{}) } };
          slotsAEffacer.forEach(slot => { nouvellesExceptions[ens.id][slot] = ""; });
        }

        // Cascade identique pour prog_suivi (progression de cours + suivi digital) :
        // une classe qui quitte un enseignant ne doit pas laisser de traces de
        // progression rattachées à lui — même logique que la suppression complète
        // d'un enseignant, mais limitée à la seule classe retirée.
        for (const cl of classesRetirees) {
          await sb.del("prog_suivi", `?ens_id=eq.${encodeURIComponent(ens.id)}&classe=eq.${encodeURIComponent(cl)}`);
          await sb.del("prog_suivi", `?ens_id=eq.${encodeURIComponent(ens.id)}&classe=eq.${encodeURIComponent(cl+"||dig")}`);
          progKeysASupprimer.push(`${ens.id}||${cl}`, `${ens.id}||${cl}||dig`);
        }
      }

      // Mettre à jour le contexte local
      setData(prev=>{
        const newProg = {...(prev.prog||{})};
        progKeysASupprimer.forEach(k => delete newProg[k]);
        return {
          ...prev,
          users: {
            ...(prev.users||{}),
            [userData.id]: {
              ...userData,
              col: getColor(userData.id),
              ini: getIni(userData.nom),
            }
          },
          exceptions: nouvellesExceptions,
          prog: newProg,
        };
      });
      const msgParts = [];
      if (classesRetirees.length>0) msgParts.push(`${classesRetirees.join(", ")} retirée(s) de son EDT`);
      if (conflits.length>0) msgParts.push(`${[...new Set(conflits.map(c=>c.classe))].join(", ")} transférée(s) depuis ${[...new Set(conflits.map(c=>getNomCourt(c.autreEns.nom)))].join(", ")}`);
      if (!mdpOk) {
        showToast(`⚠ ${userData.nom} sauvegardé(e), mais la définition du mot de passe a échoué — réessaie depuis "Gérer enseignants"`, false);
      } else {
        showToast(msgParts.length>0
          ? `✓ ${userData.nom} mis à jour · ${msgParts.join(" · ")}`
          : `✓ ${userData.nom} ${isEdit?"mis à jour":"ajouté(e)"}`);
      }
      onClose();
    } else {
      setErr("Erreur Supabase. Vérifiez les données.");
    }
    setSaving(false);
  };

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,overflowY:"auto",padding:20}}>
      <div style={{background:C.white,borderRadius:16,padding:"28px 30px",width:"100%",maxWidth:480,boxShadow:"0 24px 80px rgba(0,0,0,.25)"}}>
        <h3 style={{fontSize:16,fontWeight:800,color:C.txt,margin:"0 0 4px"}}>
          {isEdit?"✎ Modifier l'enseignant":"➕ Nouvel enseignant"}
        </h3>
        <p style={{fontSize:12,color:C.txtMuted,margin:"0 0 20px"}}>Département SVTEEHB · Lycée de Kakatare</p>

        {err && (
          <div style={{background:C.redPale,border:`1px solid ${C.redBorder||"#fca5a5"}`,borderRadius:8,padding:"9px 12px",marginBottom:14,fontSize:12,color:C.red,fontWeight:600}}>
            ⚠️ {err}
          </div>
        )}

        {/* Nom */}
        <div style={{marginBottom:14}}>
          <label style={{fontSize:11,fontWeight:700,color:C.txtMuted,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:".06em"}}>
            Nom complet <span style={{color:C.red}}>*</span>
          </label>
          <input value={nom} onChange={e=>setNom(e.target.value.toUpperCase().trim())}
            placeholder="ex. DUPONT Marie Claire"
            maxLength={60}
            style={{width:"100%",padding:"10px 14px",border:`1.5px solid ${C.border}`,borderRadius:9,fontSize:13,color:C.txt,fontFamily:"inherit"}}
            onFocus={e=>e.target.style.borderColor=C.green}
            onBlur={e=>e.target.style.borderColor=C.border}/>
        </div>

        {/* Photo */}
        <div style={{marginBottom:14}}>
          <label style={{fontSize:11,fontWeight:700,color:C.txtMuted,display:"block",marginBottom:7,textTransform:"uppercase",letterSpacing:".06em"}}>
            Photo
          </label>
          <div style={{display:"flex",flexDirection: isMobile?"column":"row",alignItems:"center",gap: isMobile?12:14}}>
            <div style={{width: isMobile?80:64,height: isMobile?80:64,borderRadius:"50%",overflow:"hidden",flexShrink:0,background:"#f1f5f9",display:"flex",alignItems:"center",justifyContent:"center",border:`1.5px solid ${C.border}`}}>
              {photoPreview ? (
                <img src={photoPreview} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/>
              ) : photoExistant ? (
                <img src={sb.photoUrl(photoExistant)} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/>
              ) : (
                <span style={{fontSize: isMobile?28:22,color:"#cbd5e1"}}>👤</span>
              )}
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:7,width: isMobile?"100%":"auto"}}>
              <input ref={fileInputRef} type="file" accept="image/*" style={{display:"none"}}
                onChange={e=>choisirPhoto(e.target.files?.[0])}/>
              <button type="button" onClick={()=>fileInputRef.current?.click()}
                style={{padding: isMobile?"11px 14px":"7px 14px",borderRadius:9,border:`1.5px solid ${C.green}`,background:C.greenPale,color:C.green,fontSize: isMobile?13:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",width:"100%",textAlign:"center"}}>
                📷 {photoExistant||photoPreview ? "Changer la photo" : "Ajouter une photo"}
              </button>
              {(photoExistant||photoPreview) && (
                <button type="button" onClick={()=>{setPhotoFile(null);setPhotoPreview(null);setPhotoExistant(null);}}
                  style={{padding: isMobile?"9px":"5px 10px",borderRadius:8,border:"none",background:"transparent",color:C.txtMuted,fontSize: isMobile?12:11,cursor:"pointer",fontFamily:"inherit",textAlign:"center",width:"100%"}}>
                  Retirer la photo
                </button>
              )}
            </div>
          </div>
        </div>

      {/* Identifiant */}
        <div style={{marginBottom:14}}>
          <label style={{fontSize:11,fontWeight:700,color:C.txtMuted,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:".06em"}}>
            Identifiant de connexion <span style={{color:C.red}}>*</span>
          </label>
          <input value={ensId} onChange={e=>setEnsId(e.target.value.toLowerCase().replace(/[^a-z0-9]/g,""))}
            maxLength={12}
            placeholder="ex. dupont" disabled={isEdit}
            style={{width:"100%",padding:"10px 14px",border:`1.5px solid ${C.border}`,borderRadius:9,fontSize:13,color:C.txt,fontFamily:"monospace",background:isEdit?"#f8fafc":"#fff"}}
            onFocus={e=>e.target.style.borderColor=C.green}
            onBlur={e=>e.target.style.borderColor=C.border}/>
          <div style={{fontSize:10,color:C.txtMuted,marginTop:3}}>
            Généré automatiquement · {isEdit?"non modifiable en édition":""}
          </div>
        </div>

        {/* Mot de passe */}
        <div style={{marginBottom:14}}>
          <label style={{fontSize:11,fontWeight:700,color:C.txtMuted,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:".06em"}}>
            Mot de passe {!isEdit && <span style={{color:C.red}}>*</span>}
          </label>
          <input type="password" value={mdp} onChange={e=>setMdp(e.target.value)}
            placeholder={isEdit?"Laisser vide pour ne pas modifier":"Minimum 8 caractères"}
            style={{width:"100%",padding:"10px 14px",border:`1.5px solid ${C.border}`,borderRadius:9,fontSize:13,color:C.txt,fontFamily:"inherit"}}
            onFocus={e=>e.target.style.borderColor=C.green}
            onBlur={e=>e.target.style.borderColor=C.border}/>
        </div>

        {/* Classes assignées */}
        <div style={{marginBottom:20}}>
          <label style={{fontSize:11,fontWeight:700,color:C.txtMuted,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:".06em"}}>
            Classes assignées
          </label>
          {/* Tags classes */}
          <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:8}}>
            {classes.map(cl=>(
              <span key={cl} style={{display:"inline-flex",alignItems:"center",gap:5,padding:"4px 10px",borderRadius:20,background:C.greenPale,border:`1px solid ${C.greenBorder}`,fontSize:12,fontWeight:700,color:C.green}}>
                {cl}
                <button onClick={()=>setClasses(prev=>prev.filter(c=>c!==cl))}
                  style={{background:"none",border:"none",cursor:"pointer",color:C.green,fontSize:14,padding:0,lineHeight:1}}>×</button>
              </span>
            ))}
            {classes.length===0 && <span style={{fontSize:12,color:C.txtLight}}>Aucune classe assignée</span>}
          </div>
          {/* Ajouter une classe */}
          <div style={{display:"flex",gap:8}}>
            <select value={newCl} onChange={e=>setNewCl(e.target.value)}
              style={{flex:1,padding:"8px 12px",border:`1.5px solid ${C.border}`,borderRadius:8,fontSize:12,color:C.txt,fontFamily:"inherit",background:C.white}}>
              <option value="">— Choisir une classe —</option>
              {Object.keys(ELEVES_DB).sort().filter(cl=>!classes.includes(cl)).map(cl=><option key={cl} value={cl}>{cl}</option>)}
            </select>
            <button onClick={ajouterClasse}
              style={{padding:"8px 14px",background:C.greenPale,border:`1px solid ${C.greenBorder}`,borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer",color:C.green,fontFamily:"inherit"}}>
              + Ajouter
            </button>
          </div>
        </div>

        {/* Boutons */}
        <div style={{display:"flex",gap:10}}>
          <button onClick={onClose} disabled={saving}
            style={{flex:1,padding:"11px",borderRadius:9,border:`1px solid ${C.border}`,background:C.white,color:C.txtMuted,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
            Annuler
          </button>
          <button onClick={sauvegarder} disabled={saving}
            style={{flex:2,padding:"11px",borderRadius:9,border:"none",background:`linear-gradient(135deg,${C.greenDark},${C.green})`,color:"#fff",fontSize:13,fontWeight:700,cursor:saving?"not-allowed":"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
            {saving?<><Spinner size={14} color="#fff"/>Sauvegarde…</>:(isEdit?"✓ Enregistrer":"✓ Créer l'enseignant")}
          </button>
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════
// ONGLET NETTOYER — Suppression des données fictives
// ══════════════════════════════════════════════════════════════════════
function EnsClean({ data, setData, showToast }) {
  const {isMobile} = useDevice();
  const [loading,  setLoading]  = useState(false);
  const [results,  setResults]  = useState(null);
  const [confirm,  setConfirm]  = useState(null); // action à confirmer

  // ── Analyser ce qui existe dans Supabase ──────────────────────────
  const epreuves    = data?.epreuves||[];
  const progEntries = Object.entries(data?.prog||{});

  // Détection données fictives / test
  const epsTest = epreuves.filter(e=>
    e.type==="test" || e.type==="demo" ||
    !e.ens_id || !DEMO_ACCOUNTS.find(a=>a.id===e.ens_id) ||
    e.titre?.toLowerCase().includes("test") ||
    e.titre?.toLowerCase().includes("démo")
  );
  const progVides = progEntries.filter(([k,v])=>!v||v.length===0);

  // ── Actions de nettoyage ──────────────────────────────────────────
  const actions = [
    {
      id:"eps-test",
      icon:"📋", titre:"Épreuves de test",
      desc:`${epsTest.length} épreuve${epsTest.length>1?"s":""} détectée${epsTest.length>1?"s":""} comme fictive${epsTest.length>1?"s":""}`,
      count: epsTest.length,
      danger: true,
      exec: async()=>{
        setLoading(true);
        let ok = true;
        for(const ep of epsTest) {
          const res = await sb.rpc("delete_epreuve", { p_id: ep.id });
          if(!res) ok = false;
        }
        setData(prev=>({...prev, epreuves:(prev.epreuves||[]).filter(e=>!epsTest.find(t=>t.id===e.id))}));
        setLoading(false);
        showToast(ok ? `✓ ${epsTest.length} épreuve${epsTest.length>1?"s":""} supprimée${epsTest.length>1?"s":""}` : "⚠ Erreurs partielles", ok);
        setResults({done:"eps-test", count:epsTest.length});
      }
    },
    {
      id:"prog-vides",
      icon:"📖", titre:"Progressions vides",
      desc:`${progVides.length} entrée${progVides.length>1?"s":""} prog_suivi sans leçons cochées`,
      count: progVides.length,
      danger: false,
      exec: async()=>{
        setLoading(true);
        let ok = true;
        for(const [key] of progVides) {
          // Split sur le PREMIER séparateur seulement : une classe peut elle-même
          // contenir "||dig" (clé digitalisation), qu'il ne faut jamais tronquer
          // sous peine de cibler/supprimer la mauvaise ligne (perte de vraies données).
          const sep = key.indexOf("||");
          const ens_id = key.slice(0, sep);
          const classe = key.slice(sep + 2);
          const res = await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: ens_id, p_classe: classe });
          if(!res) ok = false;
        }
        setData(prev=>{
          const newProg = {...(prev.prog||{})};
          progVides.forEach(([k])=>delete newProg[k]);
          return {...prev, prog:newProg};
        });
        setLoading(false);
        showToast(ok ? `✓ ${progVides.length} entrée${progVides.length>1?"s":""} nettoyée${progVides.length>1?"s":""}` : "⚠ Erreurs partielles", ok);
        setResults({done:"prog-vides", count:progVides.length});
      }
    },
    {
      id:"tout-prog",
      icon:"🔄", titre:"Réinitialiser toutes les progressions",
      desc:"Supprime TOUTES les leçons cochées de TOUS les enseignants",
      count: progEntries.filter(([,v])=>v?.length>0).length,
      danger: true,
      exec: async()=>{
        setLoading(true);
        const res = await sb.rpc("admin_delete_all_prog", {}); // supprimer toutes les progressions
        setData(prev=>({...prev, prog:{}}));
        setLoading(false);
        showToast(res ? "✓ Toutes les progressions réinitialisées" : "⚠ Erreur Supabase", res);
        setResults({done:"tout-prog", count:progEntries.length});
      }
    },
    {
      id:"tout-epreuves",
      icon:"📋", titre:"Vider toutes les épreuves",
      desc:`Supprime les ${epreuves.length} épreuve${epreuves.length>1?"s":""} de la base`,
      count: epreuves.length,
      danger: true,
      exec: async()=>{
        setLoading(true);
        const res = await sb.rpc("admin_delete_all_epreuves", {});
        setData(prev=>({...prev, epreuves:[]}));
        setLoading(false);
        showToast(res ? `✓ ${epreuves.length} épreuve${epreuves.length>1?"s":""} supprimée${epreuves.length>1?"s":""}` : "⚠ Erreur Supabase", res);
        setResults({done:"tout-epreuves", count:epreuves.length});
      }
    },
  ];

  return (
    <div style={{padding:"20px", display:"flex", flexDirection:"column", gap:16}}>

      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"16px 20px"}}>
        <h2 style={{fontSize:16, fontWeight:800, color:C.txt, margin:"0 0 4px"}}>🧹 Nettoyage des données</h2>
        <p style={{fontSize:12, color:C.txtMuted, margin:0}}>
          Supprime les données fictives, de test ou non désirées de la base Supabase.
          <strong style={{color:C.red}}> Actions irréversibles.</strong>
        </p>
      </div>

      {/* État actuel */}
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"14px 18px"}}>
        <div style={{fontSize:12, fontWeight:700, color:C.txt, marginBottom:12}}>📊 État actuel de la base</div>
        <div style={{display:"grid", gridTemplateColumns: isMobile?"repeat(2,1fr)":"repeat(3,1fr)", gap:10}}>
          {[
            {label:"Épreuves total",   val:epreuves.length,      col:epreuves.length>0?C.amber:C.green},
            {label:"Épreuves test",    val:epsTest.length,       col:epsTest.length>0?C.red:C.green},
            {label:"Progressions",     val:progEntries.length,   col:C.blue},
          ].map((k,i)=>(
            <div key={i} style={{textAlign:"center", padding:"12px 8px", background:"#f8fafc", borderRadius:9, border:`1px solid ${C.border}`}}>
              <div style={{fontSize:22, fontWeight:900, color:k.col}}>{k.val}</div>
              <div style={{fontSize:10, color:C.txtMuted, marginTop:3}}>{k.label}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Actions */}
      {actions.map(action=>{
        const done = results?.done === action.id;
        return (
          <div key={action.id} style={{background:C.white, borderRadius:12, border:`1.5px solid ${action.danger?(done?C.green:C.redA30):C.border}`, padding:"16px 18px", display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:16, opacity:action.count===0?.5:1}}>
            <div style={{flex:1}}>
              <div style={{display:"flex", alignItems:"center", gap:8, marginBottom:5}}>
                <span style={{fontSize:20}}>{action.icon}</span>
                <span style={{fontSize:14, fontWeight:700, color:C.txt}}>{action.titre}</span>
                {action.danger && <span style={{fontSize:9, padding:"1px 7px", borderRadius:20, background:C.redPale, color:C.red, fontWeight:700}}>DANGER</span>}
                {done && <span style={{fontSize:9, padding:"1px 7px", borderRadius:20, background:C.greenPale, color:C.green, fontWeight:700}}>✓ FAIT</span>}
              </div>
              <div style={{fontSize:12, color:C.txtMuted}}>{action.desc}</div>
            </div>
            <button
              disabled={loading || action.count===0 || done}
              onClick={()=>action.danger ? setConfirm(action) : action.exec()}
              style={{
                padding:"8px 16px", borderRadius:9, fontSize:12, fontWeight:700,
                cursor:(loading||action.count===0||done)?"not-allowed":"pointer",
                border:`1px solid ${action.danger?C.redA40:C.greenBorder}`,
                background:done?C.greenPale:action.danger?C.redPale:C.greenPale,
                color:done?C.green:action.danger?C.red:C.green,
                fontFamily:"inherit", flexShrink:0, opacity:(action.count===0||done)?.5:1,
              }}>
              {done ? "✓ Terminé" : loading ? "…" : action.count===0 ? "Rien à faire" : "Exécuter"}
            </button>
          </div>
        );
      })}

      {/* Modal confirmation danger */}
      {confirm && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}}>
          <div style={{background:C.white,borderRadius:16,padding:"28px 32px",width:"100%",maxWidth:420,boxShadow:"0 24px 80px rgba(0,0,0,.3)"}}>
            <div style={{fontSize:40,textAlign:"center",marginBottom:12}}>⚠️</div>
            <h3 style={{fontSize:16,fontWeight:800,color:C.txt,textAlign:"center",marginBottom:8}}>{confirm.titre}</h3>
            <p style={{fontSize:13,color:C.txtMuted,textAlign:"center",marginBottom:24,lineHeight:1.5}}>
              {confirm.desc}<br/>
              <strong style={{color:C.red}}>Cette action est irréversible.</strong>
            </p>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setConfirm(null)}
                style={{flex:1,padding:"11px",borderRadius:9,border:`1px solid ${C.border}`,background:C.white,color:C.txtMuted,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                Annuler
              </button>
              <button onClick={()=>{setConfirm(null);confirm.exec();}}
                style={{flex:1,padding:"11px",borderRadius:9,border:"none",background:C.red,color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                ✓ Confirmer
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


function EnsListe({ data, onSelect }) {
  const {isMobile} = useDevice();
  const [search, setSearch] = useState("");

  // Fallback : si Supabase ne retourne pas les utilisateurs, utiliser DEMO_ACCOUNTS
  // Source des enseignants : fusionner Supabase + ENS_CLASSES_REF
  // Si un enseignant Supabase n'a pas de classes → utiliser ENS_CLASSES_REF comme fallback
  const supabaseEns = Object.values(data?.users||{}).filter(u=>u.role!=="proviseur");
  const sourceData = (supabaseEns.length > 0 || data?.deptFilterActive)
    // Enrichir les users Supabase avec ENS_CLASSES_REF si classes vides
    ? supabaseEns.map(u=>({
        ...u,
        col: u.col||getColor(u.id), ini: u.ini||getIni(u.nom),
        classes: (u.classes||[]).length > 0 ? u.classes : (ENS_CLASSES_REF[u.id]||[])
      }))
    // Fallback complet : DEMO_ACCOUNTS + ENS_CLASSES_REF
    : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant").map(a=>({
        ...a, col:getColor(a.id), ini:getIni(a.nom),
        classes: ENS_CLASSES_REF[a.id]||[]
      }));

  const enseignants = sourceData
    .map(ens => {
      let tf = 0, tr = 0;
      (ens.classes||[]).forEach(cl => {
        const fait = ((data?.prog||{})[`${ens.id}||${cl}`]||[]).length;
        const code = resolveProgCode(cl);
        const ref  = code ? (PROG_META[code]?.lpRef||0) : 0;
        tf += fait; tr += ref;
      });
      const taux    = tr>0 ? Math.min(100, Math.round(tf/tr*100)) : 0;
      const epAttente = (data.epreuves||[]).filter(e=>e.ens_id===ens.id&&e.statut==="attente").length;
      return {...ens, totalFait:tf, totalRef:tr, taux, epAttente};
    })
    .filter(e => !search || e.nom.toLowerCase().includes(search.toLowerCase()))
    .sort((a,b) => b.taux - a.taux);

  const tauxMoyen = enseignants.length
    ? Math.round(enseignants.reduce((s,e)=>s+e.taux,0)/enseignants.length) : 0;

  const alertes = enseignants.filter(e=>e.taux<50).length;

  return (
    <div style={{padding:"20px", display:"flex", flexDirection:"column", gap:16}}>

      {/* KPIs globaux */}
      <div style={{display:"grid", gridTemplateColumns: isMobile?"repeat(2,1fr)":"repeat(4,1fr)", gap:10}}>
        {[
          {label:"Enseignants",       val:enseignants.length, emoji:"👥", col:C.green,  bg:C.greenPale},
          {label:"Couverture moyenne",val:`${tauxMoyen}%`,    emoji:"📊", col:taux2col(tauxMoyen), bg:taux2bg(tauxMoyen)},
          {label:"Objectif atteint ≥75%", val:enseignants.filter(e=>e.taux>=75).length, emoji:"✅", col:C.green, bg:C.greenPale},
          {label:"En alerte <50%",   val:alertes, emoji:"⚠️", col:alertes>0?C.red:C.green, bg:alertes>0?C.redPale:C.greenPale},
        ].map((k,i)=>(
          <div key={i} style={{background:k.bg, borderRadius:11, border:`1px solid ${C.border}`, padding:"14px 16px"}}>
            <div style={{display:"flex", justifyContent:"space-between", marginBottom:6}}>
              <span style={{fontSize:10, fontWeight:600, color:C.txtMuted}}>{k.label}</span>
              <span style={{fontSize:16}}>{k.emoji}</span>
            </div>
            <div style={{fontSize:26, fontWeight:900, color:k.col, lineHeight:1}}>{k.val}</div>
          </div>
        ))}
      </div>

      {/* Barre de recherche */}
      <div style={{position:"relative"}}>
        <span style={{position:"absolute", left:12, top:"50%", transform:"translateY(-50%)", fontSize:14}}>🔍</span>
        <input value={search} onChange={e=>setSearch(e.target.value)}
          placeholder="Rechercher un enseignant…"
          style={{width:"100%", padding:"10px 14px 10px 36px", border:`1px solid ${C.border}`, borderRadius:10, fontSize:13, color:C.txt, background:"#f8fafc", outline:"none", fontFamily:"inherit"}}
          onFocus={e=>{e.target.style.borderColor=C.green; e.target.style.background=C.white;}}
          onBlur={e=>{e.target.style.borderColor=C.border; e.target.style.background="#f8fafc";}}/>
      </div>

      {/* Tableau enseignants */}
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, overflow:"hidden"}}>
        <div style={{padding:"12px 18px", borderBottom:`1px solid ${C.border}`, display:"flex", justifyContent:"space-between", alignItems:"center"}}>
          <h3 style={{margin:0, fontSize:13, fontWeight:700, color:C.txt}}>👥 Tous les enseignants</h3>
          <span style={{fontSize:11, color:C.txtMuted}}>{enseignants.length} enseignant{enseignants.length>1?"s":""}</span>
        </div>

        <div style={{display:"flex", flexDirection:"column"}}>
          {enseignants.length === 0 ? (
            <div style={{padding:"40px", textAlign:"center", color:C.txtLight}}>
              <div style={{fontSize:32, marginBottom:8}}>🔍</div>
              Aucun enseignant trouvé
            </div>
          ) : enseignants.map((ens, i) => {
            const alerte = ens.taux < 50;
            const bon    = ens.taux >= 75;
            return (
              <div key={ens.id}
                onClick={()=>onSelect(ens)}
                style={{display:"flex", alignItems:"center", gap:12, padding:"14px 18px", borderBottom:`1px solid ${C.border}`, cursor:"pointer", transition:"background .15s", background:alerte?"rgba(239,68,68,.02)":"transparent"}}
                onMouseEnter={e=>e.currentTarget.style.background=C.greenPaleA30}
                onMouseLeave={e=>e.currentTarget.style.background=alerte?"rgba(239,68,68,.02)":"transparent"}>

                {/* Rang */}
                <div style={{width:24, textAlign:"center", fontSize:13, fontWeight:800, color:i<3?[C.green,C.amber,C.orange][i]:C.txtLight, flexShrink:0}}>
                  #{i+1}
                </div>

                {/* Avatar */}
                <Avatar ens={ens} size={38} fontSize={11}/>

                {/* Infos */}
                <div style={{flex:1, minWidth:0}}>
                  <div style={{display:"flex", alignItems:"center", gap:8, marginBottom:4}}>
                    <span style={{fontSize:13, fontWeight:700, color:C.txt, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis"}}>{ens.nom}</span>
                    {alerte && <span style={{fontSize:9, padding:"1px 6px", borderRadius:20, background:C.redPale, color:C.red, fontWeight:700, flexShrink:0}}>⚠ Alerte</span>}
                    {bon    && <span style={{fontSize:9, padding:"1px 6px", borderRadius:20, background:C.greenPale, color:C.green, fontWeight:700, flexShrink:0}}>✓ Objectif</span>}
                    {ens.epAttente>0 && <span style={{fontSize:9, padding:"1px 6px", borderRadius:20, background:C.amberPale, color:C.amber, fontWeight:700, flexShrink:0}}>⏳ {ens.epAttente} épreuve{ens.epAttente>1?"s":""}</span>}
                  </div>
                  <div style={{display:"flex", alignItems:"center", gap:6, flexWrap:"wrap"}}>
                    <div style={{flex:"1 1 80px", minWidth:60}}><ProgBar value={ens.taux}/></div>
                    <span style={{fontSize:12, fontWeight:800, color:taux2col(ens.taux), flexShrink:0}}>{ens.taux}%</span>
                    <span style={{fontSize:10, color:C.txtMuted, flexShrink:0}}>{ens.totalFait}/{ens.totalRef} leç.</span>
                    <span style={{fontSize:10, color:C.txtMuted, flexShrink:0}}>{(ens.classes||[]).length} cl.</span>
                  </div>
                </div>

                {/* Flèche */}
                <div style={{color:C.txtLight, fontSize:16, flexShrink:0}}>›</div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ── NIVEAU 2 : Détail d'un enseignant ─────────────────────────────────
function EnsDetail({ ens, data, setData, showToast, onBack, onViewClass }) {
  const [trim, setTrim] = useState(0);

  const classes = (ens.classes||[]).map(cl => {
    const key  = `${ens.id}||${cl}`;
    const fait = ((data?.prog||{})[key]||[]).length;
    const code = resolveProgCode(cl);
    const meta = code ? PROG_META[code] : null;
    const ref  = meta?.lpRef||0;
    const taux = ref>0 ? Math.min(100, Math.round(fait/ref*100)) : 0;
    const ef   = (ELEVES_DB[cl]||[]).length;
    // LP/LF par trimestre
    let lpTrim = ref, lfTrim = fait;
    let range = null;
    if (trim > 0 && code) {
      range = getTrimRange(code, ["T1","T2","T3"][trim-1]);
      if (range) {
        const lecons = LECONS_DATA[code]||[];
        lpTrim = lecons.filter(l=>l.n>=range[0]&&l.n<=range[1]).length || ref;
        lfTrim = ((data?.prog||{})[key]||[]).filter(n=>n>=range[0]&&n<=range[1]).length;
      }
    }
    // tpPrevu/tpFait limités à la même plage que lpTrim/lfTrim (sinon incohérent avec le trimestre affiché)
    const tpAll   = meta?.tp||[];
    const tpPrevu = range ? tpAll.filter(n=>n>=range[0]&&n<=range[1]).length : tpAll.length;
    const tpFait  = tpAll.filter(n => ((data?.prog||{})[key]||[]).includes(n) && (!range || (n>=range[0]&&n<=range[1]))).length;
    const tauxTrim = lpTrim>0 ? Math.min(100, Math.round(lfTrim/lpTrim*100)) : 0;
    const eps = (data.epreuves||[]).filter(e=>e.ens_id===ens.id&&e.classe===cl);
    return {cl, fait, ref, taux, tauxTrim, lpTrim, lfTrim, ef, tpPrevu, tpFait, eps, meta};
  });

  const totalFait = classes.reduce((s,c)=>s+c.lfTrim,0);
  const totalRef  = classes.reduce((s,c)=>s+c.lpTrim,0);
  const tauxGlobal= totalRef>0?Math.min(100, Math.round(totalFait/totalRef*100)):0;

  // EDT de cet enseignant
  const edt = EDT_REEL[ens.id]||{};
  const {jk:nowJk, hi:nowHi} = getNowInfo();
  const coursActuel = nowJk&&nowHi>=0 ? edt[nowJk]?.[nowHi] : null;

  return (
    <div style={{padding:"20px", display:"flex", flexDirection:"column", gap:16}}>

      {/* Breadcrumb */}
      <div style={{display:"flex", alignItems:"center", gap:8}}>
        <button onClick={onBack}
          style={{padding:"6px 12px", background:C.white, border:`1px solid ${C.border}`, borderRadius:8, fontSize:12, color:C.txtMuted, cursor:"pointer", fontFamily:"inherit", display:"flex", alignItems:"center", gap:5}}>
          ← Retour
        </button>
        <span style={{fontSize:12, color:C.txtMuted}}>Enseignants</span>
        <span style={{fontSize:12, color:C.txtMuted}}>›</span>
        <span style={{fontSize:12, fontWeight:700, color:C.txt}}>{ens.nom}</span>
      </div>

      {/* Header enseignant */}
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"18px 20px"}}>
        <div style={{display:"flex", alignItems:"center", gap:14, marginBottom:14}}>
          <Avatar ens={ens} size={48} fontSize={16}/>
          <div style={{flex:1}}>
            <div style={{fontSize:17, fontWeight:800, color:C.txt}}>{ens.nom}</div>
            <div style={{fontSize:12, color:C.txtMuted, marginTop:2}}>
              Enseignant SVTEEHB · {classes.length} classe{classes.length>1?"s":""} · {classes.reduce((s,c)=>s+c.ef,0)} élèves
            </div>
            {coursActuel && (
              <div style={{marginTop:4, fontSize:11, color:C.green, fontWeight:700}}>
                🟢 En cours : {coursActuel}
              </div>
            )}
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:28, fontWeight:900, color:taux2col(tauxGlobal)}}>{tauxGlobal}%</div>
            <div style={{fontSize:10, color:C.txtMuted}}>couverture globale</div>
          </div>
        </div>
        <ProgBar value={tauxGlobal} h={10}/>

        {/* Filtre trimestre */}
        <div style={{display:"flex", gap:6, marginTop:12}}>
          {[{l:"Annuel",v:0},{l:"T1",v:1},{l:"T2",v:2},{l:"T3",v:3}].map(t=>(
            <button key={t.v} onClick={()=>setTrim(t.v)}
              style={{flex:1, padding:"6px 0", borderRadius:8, fontSize:11, fontWeight:700, cursor:"pointer", fontFamily:"inherit", border:`1.5px solid ${trim===t.v?C.green:C.border}`, background:trim===t.v?C.green:"transparent", color:trim===t.v?"#fff":C.txtMuted, transition:"all .15s"}}>
              {t.l}
            </button>
          ))}
        </div>
      </div>

      {/* Cartes par classe */}
      <div style={{display:"flex", flexDirection:"column", gap:10}}>
        {classes.map(c => {
          const alerte = c.tauxTrim < 50;
          const bon    = c.tauxTrim >= 75;
          return (
            <div key={c.cl} style={{background:C.white, borderRadius:11, border:`1.5px solid ${alerte?C.redA30:bon?C.greenBorder:C.border}`, overflow:"hidden"}}>
              {/* Header classe */}
              <div style={{padding:"12px 16px", display:"flex", alignItems:"center", justifyContent:"space-between", borderBottom:`1px solid ${C.border}`, background:alerte?"rgba(239,68,68,.03)":bon?C.greenPaleA40:"transparent"}}>
                <div style={{display:"flex", alignItems:"center", gap:10}}>
                  <div style={{width:34, height:34, borderRadius:9, background:getColor(ens.id), display:"flex", alignItems:"center", justifyContent:"center", fontSize:10, fontWeight:800, color:"#fff"}}>
                    {c.cl.slice(0,3).replace(/\s/g,"")}
                  </div>
                  <div>
                    <div style={{fontSize:13, fontWeight:800, color:C.txt}}>{c.cl}</div>
                    <div style={{fontSize:10, color:C.txtMuted}}>{c.ef} élèves · {c.meta?.vh||0}h/sem</div>
                  </div>
                  {alerte && <span style={{fontSize:9, padding:"2px 8px", borderRadius:20, background:C.redPale, color:C.red, fontWeight:700}}>⚠ En retard</span>}
                  {bon    && <span style={{fontSize:9, padding:"2px 8px", borderRadius:20, background:C.greenPale, color:C.green, fontWeight:700}}>✓ Objectif</span>}
                </div>
                <div style={{textAlign:"right"}}>
                  <div style={{fontSize:20, fontWeight:900, color:taux2col(c.tauxTrim)}}>{c.tauxTrim}%</div>
                  <div style={{fontSize:10, color:C.txtMuted}}>{c.lfTrim}/{c.lpTrim}</div>
                </div>
              </div>

              {/* Métriques */}
              <div style={{padding:"10px 16px"}}>
                <ProgBar value={c.tauxTrim} h={6}/>
                <div style={{display:"flex", gap:12, marginTop:10, flexWrap:"wrap"}}>
                  <div style={{display:"flex", gap:4, alignItems:"center"}}>
                    <span style={{fontSize:10, color:C.txtMuted}}>Leçons :</span>
                    <span style={{fontSize:11, fontWeight:700, color:C.txt}}>{c.lfTrim}</span>
                    <span style={{fontSize:10, color:C.txtMuted}}>/ {c.lpTrim}</span>
                  </div>
                  <div style={{display:"flex", gap:4, alignItems:"center"}}>
                    <span style={{fontSize:10, color:C.txtMuted}}>Heures dues :</span>
                    <span style={{fontSize:11, fontWeight:700, color:C.txt}}>{c.meta?.hd||"—"}</span>
                  </div>
                  {c.tpPrevu > 0 && (
                    <div style={{display:"flex", gap:4, alignItems:"center"}}>
                      <span style={{fontSize:10, color:C.txtMuted}}>TP :</span>
                      <span style={{fontSize:11, fontWeight:700, color:c.tpFait===c.tpPrevu?C.green:C.amber}}>{c.tpFait}/{c.tpPrevu}</span>
                    </div>
                  )}
                  {/* Épreuves */}
                  {c.eps.length > 0 && (
                    <div style={{display:"flex", gap:4, alignItems:"center"}}>
                      <span style={{fontSize:10, color:C.txtMuted}}>Épreuves :</span>
                      {c.eps.map((ep,ei)=>{
                        const cfg = {attente:{bg:C.amberPale,col:C.amber,label:"⏳"},validee:{bg:C.greenPale,col:C.green,label:"✅"},rejetee:{bg:C.redPale,col:C.red,label:"❌"}};
                        const s = cfg[ep.statut]||cfg.attente;
                        return <span key={ei} style={{fontSize:9, padding:"1px 6px", borderRadius:20, background:s.bg, color:s.col, fontWeight:700}}>{s.label} {ep.trim} {ep.ep}</span>;
                      })}
                    </div>
                  )}
                </div>

                {/* Actions */}
                <div style={{display:"flex", gap:8, marginTop:10}}>
                  <button onClick={()=>onViewClass(c.cl)}
                    style={{flex:1, padding:"7px 0", background:C.greenPale, border:`1px solid ${C.greenBorder}`, borderRadius:8, fontSize:11, fontWeight:700, cursor:"pointer", color:C.green, fontFamily:"inherit"}}>
                    📖 Voir les leçons
                  </button>
                </div>
              </div>
            </div>
          );
        })}
      </div>

      {/* EDT résumé */}
      <div style={{background:C.white, borderRadius:11, border:`1px solid ${C.border}`, overflow:"hidden"}}>
        <div style={{padding:"12px 16px", borderBottom:`1px solid ${C.border}`, fontSize:12, fontWeight:700, color:C.txt}}>
          🗓 Emploi du temps SVTEEHB
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%", borderCollapse:"collapse", fontSize:11}}>
            <thead>
              <tr style={{background:"#f8fafc"}}>
                <th style={{padding:"7px 10px", textAlign:"left", fontSize:10, color:C.txtMuted, fontWeight:600, width:90}}>Heure</th>
                {JOURS.map((j,ji)=>(
                  <th key={j} style={{padding:"7px 10px", textAlign:"center", fontSize:10, color:nowJk===JKEYS[ji]?C.green:C.txtMuted, fontWeight:nowJk===JKEYS[ji]?800:600}}>{j}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {HEURES.map((h,hi)=>(
                <tr key={hi} style={{borderTop:`1px solid ${C.border}`}}>
                  <td style={{padding:"5px 10px", fontSize:10, color:C.txtMuted, fontWeight:600, background:"#fafafa", whiteSpace:"nowrap"}}>{h}</td>
                  {JKEYS.map((jk,ji)=>{
                    const cl = edt[jk]?.[hi];
                    const isNow = jk===nowJk&&hi===nowHi;
                    return (
                      <td key={jk} style={{padding:4, verticalAlign:"top", background:isNow?"rgba(22,163,74,.06)":"transparent"}}>
                        {cl && (
                          <div style={{padding:"5px 7px", borderRadius:6, background:isNow?C.green:getColor(ens.id)+"18", color:isNow?"#fff":getColor(ens.id), fontSize:10, fontWeight:700, border:isNow?`2px solid ${C.green}`:"none"}}>
                            {cl}
                          </div>
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── NIVEAU 3 : Leçons d'une classe (lecture seule pour l'animatrice) ──
function EnsClasLecons({ ens, cl, data, setData, showToast, onBack }) {
  const [trim, setTrim] = useState(0);
  const [search, setSearch] = useState("");
  const [saving,     setSaving]     = useState(false);
  const [closedSeqs, setClosedSeqs] = useState(new Set());
  const syncTimer = useRef(null);

  useEffect(()=>()=>{ if(syncTimer.current) clearTimeout(syncTimer.current); },[]);

  const toggleSeq = (seq) => setClosedSeqs(prev => {
    const next = new Set(prev);
    if(next.has(seq)) next.delete(seq); else next.add(seq);
    return next;
  });

  const progKey = `${ens.id}||${cl}`;
  const faites  = (data?.prog||{})[progKey]||[];
  const code    = resolveProgCode(cl);
  const meta    = code ? PROG_META[code] : null;
  const lecons  = code ? (LECONS_DATA[code]||[]) : [];

  const trimKey = trim>0 ? ["T1","T2","T3"][trim-1] : null;
  const range   = trimKey&&code ? getTrimRange(code,trimKey) : null;
  const filtered = lecons.filter(l=>{
    if(range&&(l.n<range[0]||l.n>range[1])) return false;
    if(search&&!l.t.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });

  const totalRef  = range ? (lecons.filter(l=>l.n>=range[0]&&l.n<=range[1]).length || meta?.lpRef||lecons.length) : (meta?.lpRef||lecons.length);
  const totalFait = range ? faites.filter(n=>n>=range[0]&&n<=range[1]).length : faites.length;
  const taux      = totalRef>0?Math.min(100, Math.round(totalFait/totalRef*100)):0;
  const seqs      = [...new Set(filtered.map(l=>(l.seq||l.s)))];

  // L'animatrice peut aussi cocher pour corriger
  const toggleLecon = useCallback(async(n)=>{
    const current = [...((data?.prog||{})[progKey]||[])];
    const newFaites = current.includes(n)?current.filter(x=>x!==n):[...current,n];
    setData(prev=>({...prev,prog:{...prev.prog,[progKey]:newFaites}}));
    if(syncTimer.current) clearTimeout(syncTimer.current);
    syncTimer.current = setTimeout(async()=>{
      setSaving(true);
      let ok = false;
      const _pr = await fetch(`${SB_URL}/rest/v1/prog_suivi?ens_id=eq.${ens.id}&classe=eq.${encodeURIComponent(cl)}`,{
        method:"PATCH", headers:{...sb.h(),"Prefer":"count=exact"},
        body:JSON.stringify({faites:newFaites})
      });
      if(_pr.ok){
        const _cr = _pr.headers.get("Content-Range")||"";
        const _m = _cr ? parseInt(_cr.split("/")[1]||"0") : 0;
        if(_m > 0){ ok=true; }
        else {
          const _ir = await fetch(`${SB_URL}/rest/v1/prog_suivi`,{
            method:"POST", headers:{...sb.h(),"Prefer":"return=minimal"},
            body:JSON.stringify({ens_id:ens.id,classe:cl,faites:newFaites})
          });
          ok = _ir.ok || _ir.status===201;
        }
      }
      setSaving(false);
      if(ok) showToast("✓ Progression mise à jour");
      else showToast("⚠ Erreur",false);
    },600);
  },[data,ens,cl,progKey]);

  // Clé et timer dédiés à la digitalisation (indépendants de toggleLecon)
  const digKey   = `${ens.id}||${cl}||dig`;
  const digFaites = (data?.prog||{})[digKey]||[];
  const digTimer = useRef(null);

  const toggleDigital = useCallback(async(n, e)=>{
    e.stopPropagation(); // ne pas déclencher toggleLecon sur la même ligne
    const current = [...((data?.prog||{})[digKey]||[])];
    const newDig  = current.includes(n) ? current.filter(x=>x!==n) : [...current,n];
    setData(prev=>({...prev, prog:{...prev.prog, [digKey]:newDig}}));
    if(digTimer.current) clearTimeout(digTimer.current);
    digTimer.current = setTimeout(async()=>{
      setSaving(true);
      let ok = false;
      const _pr = await fetch(`${SB_URL}/rest/v1/prog_suivi?ens_id=eq.${ens.id}&classe=eq.${encodeURIComponent(cl+"||dig")}`,{
        method:"PATCH", headers:{...sb.h(),"Prefer":"count=exact"},
        body:JSON.stringify({faites:newDig})
      });
      if(_pr.ok){
        const _cr = _pr.headers.get("Content-Range")||"";
        const _m  = _cr ? parseInt(_cr.split("/")[1]||"0") : 0;
        if(_m > 0){ ok=true; }
        else {
          const _ir = await fetch(`${SB_URL}/rest/v1/prog_suivi`,{
            method:"POST", headers:{...sb.h(),"Prefer":"return=minimal"},
            body:JSON.stringify({ens_id:ens.id, classe:cl+"||dig", faites:newDig})
          });
          ok = _ir.ok || _ir.status===201;
        }
      }
      setSaving(false);
      if(ok) showToast("✓ Digitalisation mise à jour");
      else showToast("⚠ Erreur sauvegarde digitalisation",false);
    },600);
  },[data,ens,cl,digKey]);

  return (
    <div style={{padding:"20px", display:"flex", flexDirection:"column", gap:14}}>

      {/* Breadcrumb */}
      <div style={{display:"flex", alignItems:"center", gap:8}}>
        <button onClick={onBack}
          style={{padding:"6px 12px", background:C.white, border:`1px solid ${C.border}`, borderRadius:8, fontSize:12, color:C.txtMuted, cursor:"pointer", fontFamily:"inherit"}}>
          ← {ens.nom}
        </button>
        <span style={{fontSize:12, color:C.txtMuted}}>›</span>
        <span style={{fontSize:12, fontWeight:700, color:C.txt}}>{cl}</span>
        {saving && <Pill ch="🔄 Sauvegarde…" color={C.blue}/>}
      </div>

      {/* Header */}
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"14px 18px"}}>
        <div style={{display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:10}}>
          <div>
            <div style={{fontSize:15, fontWeight:800, color:C.txt}}>{cl} — {PROGRAMMES_LABELS[code]||code||"Programme"}</div>
            <div style={{fontSize:11, color:C.txtMuted, marginTop:3}}>{ens.nom} · {meta?.vh||0}h/sem · {meta?.hd||0}h annuelles</div>
          </div>
          <div style={{textAlign:"right", display:"flex", flexDirection:"column", gap:4}}>
            <div style={{fontSize:24, fontWeight:900, color:taux2col(taux)}}>{taux}%</div>
            <div style={{fontSize:10, color:C.txtMuted}}>{totalFait}/{totalRef} leçons</div>
            {(() => {
              const digTotal = lecons.filter(l=>l.d===1 && (!range || (l.n>=range[0]&&l.n<=range[1]))).length;
              const digFait  = digFaites.filter(n=>{ const l=lecons.find(x=>x.n===n); return l&&l.d===1&&(!range||(n>=range[0]&&n<=range[1])); }).length;
              return digTotal>0 ? (
                <div style={{display:"inline-flex", alignItems:"center", gap:5, padding:"2px 8px", borderRadius:20,
                  background:digFait>0?"#eff6ff":"#f8fafc", border:`1px solid ${digFait>0?"#bfdbfe":"#e2e8f0"}`,
                  fontSize:10, fontWeight:700, color:digFait>0?"#0369a1":"#94a3b8"}}>
                  💻 {digFait}/{digTotal} digital
                </div>
              ) : null;
            })()}
          </div>
        </div>
        <ProgBar value={taux} h={8}/>
        <div style={{display:"flex", gap:6, marginTop:10}}>
          {[{l:"Annuel",v:0},{l:"T1",v:1},{l:"T2",v:2},{l:"T3",v:3}].map(t=>(
            <button key={t.v} onClick={()=>setTrim(t.v)}
              style={{flex:1, padding:"6px 0", borderRadius:7, fontSize:11, fontWeight:700, cursor:"pointer", fontFamily:"inherit", border:`1.5px solid ${trim===t.v?C.green:C.border}`, background:trim===t.v?C.green:"transparent", color:trim===t.v?"#fff":C.txtMuted}}>
              {t.l}
            </button>
          ))}
        </div>
      </div>

      {/* Recherche */}
      <input value={search} onChange={e=>setSearch(e.target.value)}
        placeholder="🔍 Rechercher une leçon…"
        style={{padding:"8px 14px", border:`1px solid ${C.border}`, borderRadius:9, fontSize:12, color:C.txt, background:"#f8fafc", outline:"none", fontFamily:"inherit"}}
        onFocus={e=>{e.target.style.borderColor=C.green; e.target.style.background=C.white;}}
        onBlur={e=>{e.target.style.borderColor=C.border; e.target.style.background="#f8fafc";}}/>

      {/* Leçons par séquence */}
      {seqs.map(seq=>{
        const leconsSec = filtered.filter(l=>(l.seq||l.s)===seq);
        const secFait = leconsSec.filter(l=>faites.includes(l.n)).length;
        const secTaux = leconsSec.length>0?Math.min(100, Math.round(secFait/leconsSec.length*100)):0;
        return (
          <div key={seq} style={{background:C.white, borderRadius:11, border:`1px solid ${C.border}`, overflow:"hidden"}}>
            <div style={{padding:"10px 14px", background:"#f8fafc", borderBottom:`1px solid ${C.border}`, display:"flex", justifyContent:"space-between", alignItems:"center"}}>
              <div>
                <div style={{fontSize:12, fontWeight:700, color:C.txt}}>{seq}</div>
                <div style={{fontSize:10, color:C.txtMuted}}>{secFait}/{leconsSec.length} dispensées</div>
              </div>
              <div style={{display:"flex", alignItems:"center", gap:8}}>
                <div style={{width:80}}><ProgBar value={secTaux}/></div>
                <span style={{fontSize:11, fontWeight:700, color:taux2col(secTaux)}}>{secTaux}%</span>
              </div>
            </div>
            {leconsSec.map((l,li)=>{
              const done      = faites.includes(l.n);
              const isTP      = meta?.tp?.includes(l.n);
              const isDigital = l.d===1; // leçon disponible en version digitalisée
              const digDone   = digFaites.includes(l.n); // ressource digitale utilisée
              return (
                <div key={l.n} onClick={()=>toggleLecon(l.n)}
                  style={{display:"flex", alignItems:"center", gap:12, padding:"9px 14px", borderTop:li>0?`1px solid ${C.border}`:"none", cursor:"pointer", background:done?"rgba(22,163,74,.04)":"transparent", transition:"background .12s"}}
                  onMouseEnter={e=>{if(!done)e.currentTarget.style.background="#f8fafc";}}
                  onMouseLeave={e=>e.currentTarget.style.background=done?"rgba(22,163,74,.04)":"transparent"}>
                  {/* Checkbox principale — dispensée */}
                  <div style={{width:20, height:20, borderRadius:6, border:`2px solid ${done?C.green:C.border}`, background:done?C.green:"transparent", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0}}>
                    {done && <span style={{color:"#fff", fontSize:12, fontWeight:900}}>✓</span>}
                  </div>
                  <span style={{width:22, textAlign:"center", fontSize:11, fontWeight:700, color:done?C.green:C.txtLight}}>{l.n}</span>
                  <span style={{flex:1, fontSize:12, fontWeight:done?600:400, color:done?C.txt:C.txtMuted}}>{l.t}</span>
                  <div style={{display:"flex", gap:5, alignItems:"center"}}>
                    {isTP && <Pill ch="TP" color={C.blue}/>}
                    {done && <Pill ch="✓" color={C.green}/>}
                    {/* Bouton de bascule digital — visible uniquement pour les leçons digitalisées */}
                    {isDigital && (
                      <button
                        onClick={(e)=>toggleDigital(l.n, e)}
                        title={digDone ? "Marquer comme non utilisé en digital" : "Marquer comme utilisé en digital"}
                        style={{
                          display:"inline-flex", alignItems:"center", gap:4,
                          padding:"2px 8px", borderRadius:20, fontSize:10, fontWeight:700,
                          cursor:"pointer", fontFamily:"inherit", flexShrink:0,
                          border:`1.5px solid ${digDone?"#0369a1":"#cbd5e1"}`,
                          background:digDone?"#eff6ff":"transparent",
                          color:digDone?"#0369a1":"#94a3b8",
                          transition:"all .15s",
                        }}>
                        💻 {digDone?"Dig ✓":"Dig"}
                      </button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        );
      })}
    </div>
  );
}



// ═══════════════════════════════════════════════════
// VUE ENSEIGNANT — détail par classe + cochage leçons
// ═══════════════════════════════════════════════════
function ViewTeacher({ ens, data, setData, isAdmin }) {
  const {isMobile} = useDevice();
  const [selClasse, setSelClasse] = useState(null);
  const [trim, setTrim] = useState(0); // 0 = annuel
  const [toast, setToast] = useState(null);
  const [saving, setSaving] = useState(false);
  const [search, setSearch] = useState("");
  const [showOnlyDone, setShowOnlyDone] = useState(false);
  const syncTimer = useRef(null);
  useEffect(()=>()=>{ if(syncTimer.current) clearTimeout(syncTimer.current); },[]);

  const classes = (ens.classes || []);

  // Sélection auto première classe
  useEffect(() => { if (classes.length && !selClasse) setSelClasse(classes[0]); }, [classes]);

  const progKey = selClasse ? `${ens.id}||${selClasse}` : null;
  const faites = progKey ? ((data?.prog||{})[progKey] || []) : [];
  const digKeyVT = selClasse ? `${ens.id}||${selClasse}||dig` : null;
  const digFaitesVT = digKeyVT ? ((data?.prog||{})[digKeyVT] || []) : [];
  const code = selClasse ? resolveProgCode(selClasse) : null;
  const meta = code ? PROG_META[code] : null;
  const programme = code ? buildLecons(code) : [];

  function buildLecons(code) {
    // Génère les leçons depuis PROG_META/PROG_MAP
    // Données inline compactes issues du prototype pour chaque code
    return LECONS_DATA[code] || [];
  }

  // Filtrage par trimestre
  // trim est un Number (0/1/2/3) — getTrimRange attend "T1"/"T2"/"T3"
  const trimKey = trim > 0 ? ["T1","T2","T3"][trim-1] : null;
  const trimRange = (trimKey && code) ? getTrimRange(code, trimKey) : null;
  const leconsFiltered = programme.filter(l => {
    if (trimRange && (l.n < trimRange[0] || l.n > trimRange[1])) return false;
    if (search && !l.t.toLowerCase().includes(search.toLowerCase())) return false;
    if (showOnlyDone && !faites.includes(l.n)) return false;
    return true;
  });

  // Stats
  const totalRef = meta?.lpRef || programme.length;
  const totalFait = faites.length;
  const tauxCouv = totalRef > 0 ? Math.min(100, Math.round(totalFait / totalRef * 100)) : 0;
  const tpTotal = meta?.tp?.length || 0;
  const tpFait = (meta?.tp || []).filter(n => faites.includes(n)).length;
  const digTotal = programme.filter(l=>l.d===1).length;
  const digFaitCount = digFaitesVT.filter(n=>{ const l=programme.find(x=>x.n===n); return l&&l.d===1; }).length;

  // Séquences uniques pour le résumé
  const seqs = [...new Set(leconsFiltered.map(l => (l.seq||l.s)))];

  // Toggle leçon (cochage)
  const toggleLecon = useCallback(async (n) => {
    if (isAdmin) return; // animateur ne coche pas
    const key = `${ens.id}||${selClasse}`;
    const current = [...((data?.prog||{})[key] || [])];
    const newFaites = current.includes(n) ? current.filter(x => x !== n) : [...current, n];

    // Mise à jour optimiste
    setData(prev => ({ ...prev, prog: { ...prev.prog, [key]: newFaites } }));

    // Debounce sync Supabase (500ms)
    if (syncTimer.current) clearTimeout(syncTimer.current);
    syncTimer.current = setTimeout(async () => {
      setSaving(true);
      let ok = false;
      const q1 = `?ens_id=eq.${ens.id}&classe=eq.${encodeURIComponent(selClasse)}`;
      const pr1 = await fetch(`${SB_URL}/rest/v1/prog_suivi${q1}`,{
        method:"PATCH", headers:{...sb.h(),"Prefer":"count=exact"},
        body:JSON.stringify({faites:newFaites})
      });
      if(pr1.ok){
        const cr1 = pr1.headers.get("Content-Range")||"";
        const m1 = cr1 ? parseInt(cr1.split("/")[1]||"0") : 0;
        if(m1>0){ ok=true; }
        else {
          const ir1 = await fetch(`${SB_URL}/rest/v1/prog_suivi`,{
            method:"POST", headers:{...sb.h(),"Prefer":"return=minimal"},
            body:JSON.stringify({ens_id:ens.id, classe:selClasse, faites:newFaites})
          });
          ok = ir1.ok || ir1.status===201;
        }
      }
      setSaving(false);
      showToast(ok ? "✓ Progression sauvegardée" : "⚠ Erreur de sauvegarde", ok);
    }, 500);
  }, [data, ens, selClasse, isAdmin]);

  // Toggle digital — même mécanique que toggleLecon mais sur la clé ||dig
  const digTimerVT = useRef(null);
  useEffect(()=>()=>{ if(digTimerVT.current) clearTimeout(digTimerVT.current); },[]);
  const toggleDigitalVT = useCallback(async (n, e) => {
    e.stopPropagation();
    const key = `${ens.id}||${selClasse}||dig`;
    const current = [...((data?.prog||{})[key] || [])];
    const newDig = current.includes(n) ? current.filter(x => x !== n) : [...current, n];
    setData(prev => ({ ...prev, prog: { ...prev.prog, [key]: newDig } }));
    if (digTimerVT.current) clearTimeout(digTimerVT.current);
    digTimerVT.current = setTimeout(async () => {
      setSaving(true);
      let ok = false;
      const classeEnc = encodeURIComponent(selClasse + "||dig");
      const pr = await fetch(`${SB_URL}/rest/v1/prog_suivi?ens_id=eq.${ens.id}&classe=eq.${classeEnc}`,{
        method:"PATCH", headers:{...sb.h(),"Prefer":"count=exact"},
        body:JSON.stringify({faites:newDig})
      });
      if(pr.ok){
        const cr = pr.headers.get("Content-Range")||"";
        const m  = cr ? parseInt(cr.split("/")[1]||"0") : 0;
        if(m>0){ ok=true; }
        else {
          const ir = await fetch(`${SB_URL}/rest/v1/prog_suivi`,{
            method:"POST", headers:{...sb.h(),"Prefer":"return=minimal"},
            body:JSON.stringify({ens_id:ens.id, classe:selClasse+"||dig", faites:newDig})
          });
          ok = ir.ok || ir.status===201;
        }
      }
      setSaving(false);
      showToast(ok ? "✓ Digitalisation mise à jour" : "⚠ Erreur de sauvegarde", ok);
    }, 500);
  }, [data, ens, selClasse]);

  function showToast(msg, ok) {
    setToast({ msg, ok });
    setTimeout(() => setToast(null), 2500);
  }

  if (!classes.length) {
    return (
      <div style={{ textAlign:"center",padding:"60px 24px",color:C.txtMuted }}>
        <div style={{ fontSize:40,marginBottom:12 }}>📭</div>
        <div style={{ fontSize:14,fontWeight:600 }}>Aucune classe assignée à {ens.nom}</div>
      </div>
    );
  }

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      {toast && <Toast msg={toast.msg} ok={toast.ok}/>}

      {/* Sélecteur de classes */}
      <div style={{ display:"flex",gap:8,flexWrap:"wrap",alignItems:"center" }}>
        <span style={{ fontSize:12,color:C.txtMuted,fontWeight:600 }}>Classe :</span>
        {classes.map(cl => {
          const k = `${ens.id}||${cl}`;
          const fait = ((data?.prog||{})[k]||[]).length;
          const code2 = resolveProgCode(cl);
          const ref2 = code2 ? (PROG_META[code2]?.lpRef || 0) : 0;
          const t = ref2 > 0 ? Math.min(100, Math.round(fait/ref2*100)) : 0;
          const isActive = cl === selClasse;
          return (
            <button key={cl} onClick={() => { setSelClasse(cl); setSearch(""); }}
              style={{ padding:"6px 12px",borderRadius:20,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",border:`2px solid ${isActive ? taux2col(t) : C.border}`,background:isActive ? taux2bg(t) : C.white,color:isActive ? taux2col(t) : C.txtMuted,transition:"all .15s",display:"flex",alignItems:"center",gap:5 }}>
              {cl}
              <span style={{ fontSize:11,fontWeight:800,color:taux2col(t) }}>{t}%</span>
            </button>
          );
        })}
      </div>

      {selClasse && (
        <>
          {/* Stats classe sélectionnée */}
          <div style={{ display:"grid",gridTemplateColumns: isMobile?"repeat(2,1fr)":"repeat(4,1fr)",gap:10 }}>
            {[
              { label:"Leçons dispensées", val:totalFait, sub:`/ ${totalRef} prévues`, col:taux2col(tauxCouv), emoji:"✅" },
              { label:"Couverture", val:`${tauxCouv}%`, sub:tauxCouv>=75?"Objectif atteint ✓":"Sous l'objectif", col:taux2col(tauxCouv), emoji:"📊" },
              { label:"TP effectués", val:tpFait, sub:`/ ${tpTotal} prévus`, col:C.blue, emoji:"🔬" },
              ...(digTotal>0 ? [{ label:"Ressources digitales", val:digFaitCount, sub:`/ ${digTotal} disponibles`, col:"#0369a1", emoji:"💻" }] : []),
              { label:"Volume horaire", val:meta?`${meta.vh}h/sem`:"—", sub:meta?`${meta.hd}h annuelles`:"", col:C.purple, emoji:"⏱" },
            ].map((s,i) => (
              <div key={i} style={{ background:C.white,borderRadius:10,border:`1px solid ${C.border}`,padding:"12px 14px" }}>
                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4 }}>
                  <span style={{ fontSize:10,color:C.txtMuted,fontWeight:600 }}>{s.label}</span>
                  <span style={{ fontSize:14 }}>{s.emoji}</span>
                </div>
                <div style={{ fontSize:22,fontWeight:800,color:s.col,lineHeight:1 }}>{s.val}</div>
                <div style={{ fontSize:10,color:C.txtMuted,marginTop:4 }}>{s.sub}</div>
              </div>
            ))}
          </div>

          {/* Barre de progression globale */}
          <div style={{ background:C.white,borderRadius:10,border:`1px solid ${C.border}`,padding:"14px 16px" }}>
            <div style={{ display:"flex",justifyContent:"space-between",marginBottom:8 }}>
              <span style={{ fontSize:12,fontWeight:700,color:C.txt }}>{selClasse} — {PROGRAMMES_LABELS[code] || code}</span>
              <span style={{ fontSize:13,fontWeight:800,color:taux2col(tauxCouv) }}>{tauxCouv}%</span>
            </div>
            <ProgBar value={tauxCouv} h={10}/>
          </div>

          {/* Filtres */}
          <div style={{ display:"flex",gap:10,flexWrap:"wrap",alignItems:"center" }}>
            {/* Trimestre */}
            <div style={{ display:"flex",gap:5 }}>
              {[{l:"Annuel",v:0},{l:"T1",v:1},{l:"T2",v:2},{l:"T3",v:3}].map(t => (
                <button key={t.v} onClick={() => setTrim(t.v)}
                  style={{ padding:"5px 10px",borderRadius:7,fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",border:`1.5px solid ${trim===t.v?C.green:C.border}`,background:trim===t.v?C.greenPale:C.white,color:trim===t.v?C.green:C.txtMuted }}>
                  {t.l}
                </button>
              ))}
            </div>
            {/* Recherche */}
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Rechercher une leçon…"
              style={{ padding:"5px 12px",border:`1px solid ${C.border}`,borderRadius:8,fontSize:12,color:C.txt,background:"#f8fafc",outline:"none",fontFamily:"inherit",flex:1,minWidth:160 }}
              onFocus={e=>{ e.target.style.borderColor=C.green; e.target.style.background=C.white; }}
              onBlur={e=>{ e.target.style.borderColor=C.border; e.target.style.background="#f8fafc"; }}/>
            {/* Toggle faites */}
            <label style={{ display:"flex",alignItems:"center",gap:6,cursor:"pointer",fontSize:12,color:C.txtMuted,userSelect:"none" }}>
              <input type="checkbox" checked={showOnlyDone} onChange={e=>setShowOnlyDone(e.target.checked)} style={{ width:14,height:14,accentColor:C.green }}/>
              Seulement faites
            </label>
            {/* Indicateur sauvegarde */}
            {saving && <Pill ch="🔄 Sauvegarde…" color={C.blue}/>}
            {!isAdmin && <Pill ch="✏️ Cliquez sur une leçon pour la cocher/décocher" color={C.txtMuted}/>}
          </div>

          {/* Liste des leçons par séquence */}
          <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
            {seqs.map(seq => {
              const lecons = leconsFiltered.filter(l => (l.seq||l.s) === seq);
              if (!lecons.length) return null;
              const seqFait = lecons.filter(l => faites.includes(l.n)).length;
              const seqTotal = lecons.length;
              return (
                <div key={seq} style={{ background:C.white,borderRadius:11,border:`1px solid ${C.border}`,overflow:"hidden" }}>
                  {/* Header séquence */}
                  <div style={{ padding:"10px 14px",background:"#f8fafc",borderBottom:`1px solid ${C.border}`,display:"flex",alignItems:"center",justifyContent:"space-between" }}>
                    <div style={{ flex:1 }}>
                      <div style={{ fontSize:12,fontWeight:700,color:C.txt }}>{seq}</div>
                      <div style={{ fontSize:10,color:C.txtMuted,marginTop:2 }}>{seqFait}/{seqTotal} leçons dispensées</div>
                    </div>
                    <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                      <div style={{ width:80 }}><ProgBar value={seqTotal>0?Math.min(100, Math.round(seqFait/seqTotal*100)):0}/></div>
                      <span style={{ fontSize:11,fontWeight:700,color:taux2col(seqTotal>0?Math.min(100, Math.round(seqFait/seqTotal*100)):0),minWidth:30,textAlign:"right" }}>{seqTotal>0?Math.min(100, Math.round(seqFait/seqTotal*100)):0}%</span>
                    </div>
                  </div>
                  {/* Leçons */}
                  <div>
                    {lecons.map((l, li) => {
                      const done = faites.includes(l.n);
                      const isTP = meta?.tp?.includes(l.n);
                      const isInteg = /intégration/i.test(l.t);
                      const isDigital = l.d===1;
                      const digDoneVT = digFaitesVT.includes(l.n);
                      return (
                        <div key={l.n} onClick={() => toggleLecon(l.n)}
                          style={{ display:"flex",alignItems:"center",gap:12,padding:"10px 14px",borderTop:li>0?`1px solid ${C.border}`:"none",cursor:isAdmin?"default":"pointer",background:done?"rgba(22,163,74,.04)":"transparent",transition:"background .15s" }}
                          onMouseEnter={e=>{ if(!isAdmin && !done) e.currentTarget.style.background="#f8fafc"; }}
                          onMouseLeave={e=>{ if(!isAdmin) e.currentTarget.style.background=done?"rgba(22,163,74,.04)":"transparent"; }}>
                          {/* Checkbox */}
                          <div style={{ width:20,height:20,borderRadius:6,border:`2px solid ${done?C.green:C.border}`,background:done?C.green:"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,transition:"all .2s" }}>
                            {done && <span style={{ color:"#fff",fontSize:12,fontWeight:800 }}>✓</span>}
                          </div>
                          {/* Numéro */}
                          <div style={{ width:24,textAlign:"center",fontSize:11,fontWeight:700,color:done?C.green:C.txtLight,flexShrink:0 }}>{l.n}</div>
                          {/* Titre */}
                          <div style={{ flex:1 }}>
                            <span style={{ fontSize:12,fontWeight:done?700:400,color:done?C.txt:C.txtMuted }}>{l.t}</span>
                          </div>
                          {/* Badges */}
                          <div style={{ display:"flex",gap:5,flexShrink:0,alignItems:"center" }}>
                            {isTP && <Pill ch="TP" color={C.blue}/>}
                            {isInteg && <Pill ch="Intégration" color={C.amber}/>}
                            {done && <Pill ch="✓ Fait" color={C.green}/>}
                            {isDigital && (
                              <button onClick={(e)=>toggleDigitalVT(l.n,e)}
                                title={digDoneVT?"Retirer la ressource digitale":"Marquer comme utilisé en digital"}
                                style={{ display:"inline-flex",alignItems:"center",gap:4,padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flexShrink:0,
                                  border:`1.5px solid ${digDoneVT?"#0369a1":"#cbd5e1"}`,background:digDoneVT?"#eff6ff":"transparent",color:digDoneVT?"#0369a1":"#94a3b8",transition:"all .15s" }}>
                                💻 {digDoneVT?"Dig ✓":"Dig"}
                              </button>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
            {leconsFiltered.length === 0 && (
              <div style={{ textAlign:"center",padding:"40px",color:C.txtMuted }}>
                <div style={{ fontSize:32,marginBottom:8 }}>🔍</div>
                <div style={{ fontSize:13 }}>Aucune leçon trouvée avec ces filtres</div>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════
// DONNÉES LEÇONS (subset compact — les 12 programmes)
// On génère les leçons depuis les données embarquées du prototype
// ═══════════════════════════════════════════════════

const LECONS_DATA = {
  "TERM_D": [
    {n:1,s:"SÉQ 1",t:"Échanges d'eau",c:1,d:1},
    {n:2,s:"SÉQ 1",t:"Interprétation des échanges d'eau",c:1,d:1},
    {n:3,s:"SÉQ 1",t:"Échanges des substances dissoutes, des particules",c:1,d:1},
    {n:4,s:"SÉQ 1",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:5,s:"SÉQ 1",t:"Structure d'une fibre musculaire squelettique striée",c:1,d:1},
    {n:6,s:"SÉQ 1",t:"Mécanisme de la contraction musculaire",c:1,d:1},
    {n:7,s:"SÉQ 2",t:"Voies de restauration de l'ATP",c:1,d:1},
    {n:8,s:"SÉQ 2",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:9,s:"SÉQ 2",t:"Structures et rôles des gonades chez les mammifères",c:1,d:1},
    {n:10,s:"SÉQ 2",t:"Structures et rôles des gonades chez les spermaphytes",c:1,d:1},
    {n:11,s:"SÉQ 2",t:"La méiose",c:1,d:1},
    {n:12,s:"SÉQ 2",t:"Gamétogenèse chez les mammifères",c:1,d:1},
    {n:13,s:"SÉQ 3",t:"Gamétogenèse chez les spermaphytes",c:1,d:1},
    {n:14,s:"SÉQ 3",t:"Fécondation chez les mammifères",c:1,d:1},
    {n:15,s:"SÉQ 3",t:"Fécondation chez les spermaphytes",c:1,d:1},
    {n:16,s:"SÉQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:17,s:"SÉQ 3",t:"Monohybridisme avec dominance",c:1,d:1},
    {n:18,s:"SÉQ 3",t:"Monohybridisme avec dominance intermédiaire",c:1,d:1},
    {n:19,s:"SÉQ 4",t:"Monohybridisme gonosomal",c:1,d:1},
    {n:20,s:"SÉQ 4",t:"Dihybridisme",c:1,d:1},
    {n:21,s:"SÉQ 4",t:"Trihybridisme",c:1,d:1},
    {n:22,s:"SÉQ 4",t:"Brassage intrachromosomique",c:1,d:1},
    {n:23,s:"SÉQ 4",t:"Quelques exceptions à la monogénie",c:1,d:1},
    {n:24,s:"SÉQ 4",t:"Origine des nouveaux allèles",c:1,d:1},
    {n:25,s:"SÉQ 5",t:"Notion d'arbre généalogique",c:1,d:1},
    {n:26,s:"SÉQ 5",t:"Transmissions autosomiques",c:1,d:1},
    {n:27,s:"SÉQ 5",t:"Transmissions gonosomiques",c:1,d:1},
    {n:28,s:"SÉQ 5",t:"Évaluation d'un risque génétique",c:1,d:0},
    {n:29,s:"SÉQ 5",t:"Applications et implications des connaissances",c:1,d:1},
    {n:30,s:"SÉQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:31,s:"SÉQ 6",t:"Structure du tissu nerveux",c:1,d:1},
    {n:32,s:"SÉQ 6",t:"Réflexes innés (réflexe médullaire)",c:1,d:1},
    {n:33,s:"SÉQ 6",t:"Réflexes innés (réflexe myotatique)",c:1,d:1},
    {n:34,s:"SÉQ 6",t:"Réflexes acquis",c:1,d:1},
    {n:35,s:"SÉQ 6",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:36,s:"SÉQ 6",t:"Le potentiel de repos",c:1,d:1},
    {n:37,s:"SÉQ 7",t:"Le potentiel d'action",c:1,d:1},
    {n:38,s:"SÉQ 7",t:"Naissance du potentiel d'action",c:1,d:1},
    {n:39,s:"SÉQ 7",t:"Propagation du potentiel d'action",c:1,d:1},
    {n:40,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:41,s:"SÉQ 7",t:"Synapse : Structure",c:1,d:1},
    {n:42,s:"SÉQ 7",t:"Synapse : Fonctionnement",c:1,d:1},
    {n:43,s:"SÉQ 8",t:"L'intégration neuronale",c:1,d:1},
    {n:44,s:"SÉQ 8",t:"Effets de certaines substances",c:1,d:1},
    {n:45,s:"SÉQ 8",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:46,s:"SÉQ 8",t:"Les différentes aires de l'encéphale",c:1,d:1},
    {n:47,s:"SÉQ 8",t:"Quelques aspects de la motricité dirigée",c:1,d:1},
    {n:48,s:"SÉQ 8",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:49,s:"SÉQ 9",t:"Origine des cellules immunitaires",c:1,d:1},
    {n:50,s:"SÉQ 9",t:"Structure de reconnaissance du non-soi",c:1,d:1},
    {n:51,s:"SÉQ 9",t:"Mécanisme de la réponse immunitaire non spécifique",c:1,d:1},
    {n:52,s:"SÉQ 9",t:"Mécanisme de la réponse immunitaire spécifique",c:1,d:1},
    {n:53,s:"SÉQ 9",t:"Mémoire immunitaire",c:1,d:1},
    {n:54,s:"SÉQ 9",t:"Maladies auto-immunes",c:1,d:1},
    {n:55,s:"SÉQ 10",t:"Pandémies et leurs conséquences",c:1,d:1},
    {n:56,s:"SÉQ 10",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:57,s:"SÉQ 10",t:"Régulation du taux d'hormones sexuelles chez l'homme (1)",c:1,d:1},
    {n:58,s:"SÉQ 10",t:"Régulation du taux d'hormones sexuelles chez l'homme (2)",c:1,d:1},
    {n:59,s:"SÉQ 10",t:"Régulation du taux d'hormones sexuelles chez la femme (1)",c:1,d:1},
    {n:60,s:"SÉQ 10",t:"Régulation du taux d'hormones sexuelles chez la femme (2)",c:1,d:1},
    {n:61,s:"SÉQ 11",t:"Régulation du taux d'hormones sexuelles chez la femme (3)",c:1,d:1},
    {n:62,s:"SÉQ 11",t:"Maîtrise de la reproduction",c:1,d:1},
    {n:63,s:"SÉQ 11",t:"Dérèglements des hormones sexuelles",c:1,d:1},
    {n:64,s:"SÉQ 11",t:"Stérilité",c:1,d:1},
    {n:65,s:"SÉQ 11",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:66,s:"SÉQ 11",t:"Glycémie : mesure et causes",c:1,d:1},
    {n:67,s:"SÉQ 12",t:"Glycémie : régulation",c:1,d:1},
    {n:68,s:"SÉQ 12",t:"Glycémie : conséquences",c:1,d:1},
    {n:69,s:"SÉQ 12",t:"Pression artérielle : définition et mesure",c:1,d:1},
    {n:70,s:"SÉQ 12",t:"Pression artérielle : régulation nerveuse",c:1,d:1},
    {n:71,s:"SÉQ 12",t:"Pression artérielle : régulation hormonale",c:1,d:1},
    {n:72,s:"SÉQ 12",t:"Pression artérielle : conséquences",c:1,d:1},
    {n:73,s:"SÉQ 13",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:74,s:"SÉQ 13",t:"Arrêt cardiaque",c:1,d:1},
    {n:75,s:"SÉQ 13",t:"Étouffement total",c:1,d:1},
    {n:76,s:"SÉQ 13",t:"Perte de connaissance",c:1,d:1},
    {n:77,s:"SÉQ 13",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:78,s:"SÉQ 13",t:"Plancher océanique : morphologie",c:1,d:1},
    {n:79,s:"SÉQ 14",t:"Plancher océanique : fonctionnement d'un rift",c:1,d:1},
    {n:80,s:"SÉQ 14",t:"Plancher océanique : mouvements de convergence",c:1,d:1},
    {n:81,s:"SÉQ 14",t:"Plancher océanique : mouvements de coulissage",c:1,d:1},
    {n:82,s:"SÉQ 14",t:"Catastrophes naturelles associées aux mouvements des plaques",c:1,d:1},
    {n:83,s:"SÉQ 14",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:84,s:"SÉQ 14",t:"Les preuves de l'évolution de l'Homme",c:1,d:1},
    {n:85,s:"SÉQ 15",t:"Les critères de l'hominisation",c:1,d:1},
    {n:86,s:"SÉQ 15",t:"Capacité d'adaptation de l'Homme",c:1,d:1},
    {n:87,s:"SÉQ 15",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:88,s:"SÉQ 15",t:"Transformation et conservation des fruits (mangue, tomate)",c:1,d:1},
    {n:89,s:"SÉQ 15",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:90,s:"SÉQ 15",t:"Insectes comestibles et lutte biologique",c:1,d:1},
    {n:91,s:"SÉQ 16",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:92,s:"SÉQ 16",t:"Production des biocarburants",c:1,d:1},
    {n:93,s:"SÉQ 16",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:94,s:"SÉQ 16",t:"Valorisation des déchets papiers",c:1,d:1},
    {n:95,s:"SÉQ 16",t:"Valorisation des déchets plastiques",c:1,d:1},
    {n:96,s:"SÉQ 16",t:"Valorisation des déchets plastiques (suite)",c:1,d:1},
    {n:97,s:"SÉQ 17",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "PREM_D": [
    {n:1,s:"SEQ 1",t:"Cellules animale et végétale vues au microscope",c:1,d:1},
    {n:2,s:"SEQ 1",t:"Ultrastructures et rôles des organites",c:1,d:1},
    {n:3,s:"SEQ 1",t:"Structure et composition chimique d'un chromosome",c:1,d:1},
    {n:4,s:"SEQ 1",t:"L'ADN, support de l'information génétique",c:1,d:1},
    {n:5,s:"SEQ 1",t:"Division cellulaire",c:1,d:1},
    {n:6,s:"SEQ 1",t:"Cycle cellulaire et rôles de la mitose",c:1,d:1},
    {n:7,s:"SEQ 1",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:8,s:"SEQ 2",t:"Le renouvellement permanent des molécules",c:1,d:1},
    {n:9,s:"SEQ 2",t:"La biosynthèse des protéines : la transcription",c:1,d:1},
    {n:10,s:"SEQ 2",t:"Biosynthèse des protéines : la traduction",c:1,d:1},
    {n:11,s:"SEQ 2",t:"Biosynthèse des protéines : le devenir des protéines",c:1,d:1},
    {n:12,s:"SEQ 2",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:13,s:"SEQ 2",t:"Génie génétique : caractéristiques des plantes transgéniques",c:1,d:1},
    {n:14,s:"SEQ 2",t:"Génie génétique : technique d'obtention",c:1,d:1},
    {n:15,s:"SEQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:16,s:"SEQ 3",t:"Relation entre l'équipement enzymatique et les réactions",c:1,d:1},
    {n:17,s:"SEQ 3",t:"Caractéristiques de l'activité enzymatique",c:1,d:1},
    {n:18,s:"SEQ 3",t:"Relation entre structure et fonction de la protéine",c:1,d:1},
    {n:19,s:"SEQ 3",t:"Quelques applications de la catalyse enzymatique",c:1,d:1},
    {n:20,s:"SEQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:21,s:"SEQ 3",t:"Évaluation de la dépense énergétique",c:1,d:0},
    {n:22,s:"SEQ 4",t:"Facteurs de variation de la dépense énergétique",c:1,d:1},
    {n:23,s:"SEQ 4",t:"Le métabolisme de base",c:1,d:1},
    {n:24,s:"SEQ 4",t:"Les apports énergétiques",c:1,d:1},
    {n:25,s:"SEQ 4",t:"Dépenses énergétiques produites par la respiration",c:1,d:1},
    {n:26,s:"SEQ 4",t:"Le mécanisme de la respiration",c:1,d:1},
    {n:27,s:"SEQ 4",t:"Dépenses énergétiques produites par la fermentation",c:1,d:1},
    {n:28,s:"SEQ 4",t:"Les mécanismes de la fermentation",c:1,d:1},
    {n:29,s:"SEQ 5",t:"Rendement énergétique de la respiration et de la fermentation",c:1,d:1},
    {n:30,s:"SEQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:31,s:"SEQ 5",t:"Le captage de l'énergie lumineuse",c:1,d:1},
    {n:32,s:"SEQ 5",t:"Mécanisme de la photosynthèse : la phase chimique",c:1,d:1},
    {n:33,s:"SEQ 5",t:"Mécanisme de la photosynthèse : la phase photochimique",c:1,d:1},
    {n:34,s:"SEQ 5",t:"Importance de la photosynthèse",c:1,d:1},
    {n:35,s:"SEQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:36,s:"SEQ 6",t:"Organisation d'un écosystème",c:1,d:1},
    {n:37,s:"SEQ 6",t:"Le flux et la dissipation de l'énergie",c:1,d:1},
    {n:38,s:"SEQ 6",t:"Le cycle biogéochimique du carbone (réservoirs)",c:1,d:1},
    {n:39,s:"SEQ 6",t:"Le cycle biogéochimique du carbone (modifications)",c:1,d:1},
    {n:40,s:"SEQ 6",t:"Le cycle biogéochimique de l'azote",c:1,d:1},
    {n:41,s:"SEQ 6",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:42,s:"SEQ 6",t:"Le soi",c:1,d:1},
    {n:43,s:"SEQ 7",t:"Le non-soi",c:1,d:1},
    {n:44,s:"SEQ 7",t:"Les principales cellules immunitaires",c:1,d:1},
    {n:45,s:"SEQ 7",t:"Le virus et son mode d'action",c:1,d:1},
    {n:46,s:"SEQ 7",t:"La multiplication du virus",c:1,d:1},
    {n:47,s:"SEQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:48,s:"SEQ 7",t:"La sexualité précoce et ses conséquences",c:1,d:1},
    {n:49,s:"SEQ 7",t:"Les mutilations génitales",c:1,d:1},
    {n:50,s:"SEQ 8",t:"La prise des stupéfiants",c:1,d:1},
    {n:51,s:"SEQ 8",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:52,s:"SEQ 8",t:"L'alimentation équilibrée : définition",c:1,d:1},
    {n:53,s:"SEQ 8",t:"L'alimentation équilibrée : un menu équilibré",c:1,d:1},
    {n:54,s:"SEQ 8",t:"Les conséquences de la mauvaise alimentation",c:1,d:1},
    {n:55,s:"SEQ 8",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:56,s:"SEQ 8",t:"L'origine de l'énergie reçue par la planète",c:1,d:1},
    {n:57,s:"SEQ 9",t:"Le devenir du rayonnement solaire (albédo)",c:1,d:1},
    {n:58,s:"SEQ 9",t:"Le devenir du rayonnement solaire (bilan radiatif)",c:1,d:1},
    {n:59,s:"SEQ 9",t:"L'effet de serre",c:1,d:1},
    {n:60,s:"SEQ 9",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:61,s:"SEQ 9",t:"Les causes de l'inégale répartition",c:1,d:1},
    {n:62,s:"SEQ 9",t:"Les mouvements atmosphériques",c:1,d:1},
    {n:63,s:"SEQ 9",t:"Les mouvements océaniques",c:1,d:1},
    {n:64,s:"SEQ 10",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:65,s:"SEQ 10",t:"L'altération des roches",c:1,d:1},
    {n:66,s:"SEQ 10",t:"Le devenir des produits d'altération",c:1,d:1},
    {n:67,s:"SEQ 10",t:"Sédimentation",c:1,d:1},
    {n:68,s:"SEQ 10",t:"Diagenèse et roches sédimentaires",c:1,d:1},
    {n:69,s:"SEQ 10",t:"Séries et cycles sédimentaires",c:1,d:1},
    {n:70,s:"SEQ 10",t:"La chronologie relative",c:1,d:1},
    {n:71,s:"SEQ 11",t:"La reconstitution des milieux sédimentaires",c:1,d:1},
    {n:72,s:"SEQ 11",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:73,s:"SEQ 11",t:"Les éléments de sismologie (ondes)",c:1,d:1},
    {n:74,s:"SEQ 11",t:"Les éléments de sismologie (comportement)",c:1,d:1},
    {n:75,s:"SEQ 11",t:"La structure interne de la Terre (discontinuités)",c:1,d:1},
    {n:76,s:"SEQ 11",t:"La structure interne de la Terre (composition)",c:1,d:1},
    {n:77,s:"SEQ 11",t:"L'énergie interne de la Terre",c:1,d:1},
    {n:78,s:"SEQ 12",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "PREM_CTI": [
    {n:1,s:"SEQ 1",t:"Rappels sur les rôles des aliments et de la digestion",c:1,d:1},
    {n:2,s:"SEQ 1",t:"Rappel sur l'absorption intestinale et la circulation",c:1,d:1},
    {n:3,s:"SEQ 1",t:"La respiration cellulaire : les étapes",c:1,d:1},
    {n:4,s:"SEQ 1",t:"La respiration cellulaire : rôle des transporteurs",c:1,d:1},
    {n:5,s:"SEQ 1",t:"Quelques exemples de fermentation",c:1,d:1},
    {n:6,s:"SEQ 2",t:"Comparaison respiration-fermentation",c:1,d:1},
    {n:7,s:"SEQ 2",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:8,s:"SEQ 2",t:"Évaluation de la dépense énergétique",c:1,d:0},
    {n:9,s:"SEQ 2",t:"Les facteurs de variation de la dépense énergétique",c:1,d:1},
    {n:10,s:"SEQ 2",t:"Apports énergétiques des repas",c:1,d:1},
    {n:11,s:"SEQ 3",t:"Crise cardiaque et AVC",c:1,d:1},
    {n:12,s:"SEQ 3",t:"Les moyens de lutte contre l'AVC",c:1,d:1},
    {n:13,s:"SEQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:14,s:"SEQ 3",t:"Le soi et le non-soi",c:1,d:1},
    {n:15,s:"SEQ 3",t:"Les principales cellules immunitaires",c:1,d:1},
    {n:16,s:"SEQ 4",t:"La contamination par le VIH",c:1,d:1},
    {n:17,s:"SEQ 4",t:"Les différentes phases de la maladie",c:1,d:1},
    {n:18,s:"SEQ 4",t:"La prévention et le traitement",c:1,d:1},
    {n:19,s:"SEQ 4",t:"Les causes et les conséquences (IVG, alcoolisme…)",c:1,d:1},
    {n:20,s:"SEQ 4",t:"Les moyens de lutte",c:1,d:1},
    {n:21,s:"SEQ 5",t:"Les différents groupes d'aliments simples",c:1,d:1},
    {n:22,s:"SEQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:23,s:"SEQ 5",t:"Les mouvements atmosphériques",c:1,d:1},
    {n:24,s:"SEQ 5",t:"Les mouvements océaniques",c:1,d:1},
    {n:25,s:"SEQ 5",t:"Altération et érosion des roches",c:1,d:1},
    {n:26,s:"SEQ 6",t:"Transport des produits d'altération",c:1,d:1},
    {n:27,s:"SEQ 6",t:"Sédimentation et diagenèse",c:1,d:1},
    {n:28,s:"SEQ 6",t:"Les roches sédimentaires et leur importance",c:1,d:1},
    {n:29,s:"SEQ 6",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:30,s:"SEQ 6",t:"La propagation des ondes sismiques",c:1,d:1},
    {n:31,s:"SEQ 7",t:"Les enveloppes internes du globe",c:1,d:1},
    {n:32,s:"SEQ 7",t:"Les propriétés physicochimiques des enveloppes",c:1,d:1},
    {n:33,s:"SEQ 7",t:"Origine de l'énergie interne",c:1,d:1},
    {n:34,s:"SEQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:35,s:"SEQ 7",t:"La dérive des continents : les preuves",c:1,d:1},
    {n:36,s:"SEQ 8",t:"La plaque lithosphérique",c:1,d:1},
    {n:37,s:"SEQ 8",t:"Les mouvements des plaques",c:1,d:1},
    {n:38,s:"SEQ 8",t:"La dissipation de l'énergie des écosystèmes",c:1,d:1},
    {n:39,s:"SEQ 8",t:"Les modifications du réservoir de carbone",c:1,d:1},
    {n:40,s:"SEQ 8",t:"L'effet de serre",c:1,d:1},
    {n:41,s:"SEQ 9",t:"Rôle de la couche d'ozone",c:1,d:1},
    {n:42,s:"SEQ 9",t:"Les moyens de lutte contre le réchauffement climatique",c:1,d:1},
    {n:43,s:"SEQ 9",t:"Définition et applications de la biotechnologie",c:1,d:1},
    {n:44,s:"SEQ 9",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "PREM_A": [
    {n:1,s:"SÉQ 1",t:"Les principaux constituants de la matière vivante",c:1,d:1},
    {n:2,s:"SÉQ 1",t:"Notion d'aliments simples et composés",c:1,d:1},
    {n:3,s:"SÉQ 1",t:"Rôles des aliments",c:1,d:1},
    {n:4,s:"SÉQ 1",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:5,s:"SÉQ 2",t:"La notion et les caractéristiques des caryotypes",c:1,d:1},
    {n:6,s:"SÉQ 2",t:"L'hérédité hétérochromosomique",c:1,d:1},
    {n:7,s:"SÉQ 2",t:"Relation d'agressivité et de dominance chez l'Homme",c:1,d:1},
    {n:8,s:"SÉQ 2",t:"Relation émotionnelle chez l'Homme",c:1,d:1},
    {n:9,s:"SÉQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:10,s:"SÉQ 3",t:"Le soi, le non-soi et les cellules immunitaires",c:1,d:1},
    {n:11,s:"SÉQ 3",t:"Les dysfonctionnements : les allergies",c:1,d:1},
    {n:12,s:"SÉQ 3",t:"Les déficiences : le VIH-Sida",c:1,d:1},
    {n:13,s:"SÉQ 4",t:"Les maladies nutritionnelles par carence",c:1,d:1},
    {n:14,s:"SÉQ 4",t:"Les maladies nutritionnelles par excès",c:1,d:1},
    {n:15,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:16,s:"SÉQ 4",t:"Maîtrise de la reproduction",c:1,d:1},
    {n:17,s:"SÉQ 5",t:"Quelques comportements affectant la santé reproductive",c:1,d:1},
    {n:18,s:"SÉQ 5",t:"VIH/Sida et grossesse",c:1,d:1},
    {n:19,s:"SÉQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:20,s:"SÉQ 5",t:"L'assainissement de l'habitat",c:1,d:1},
    {n:21,s:"SÉQ 6",t:"Les déchets ménagers",c:1,d:1},
    {n:22,s:"SÉQ 6",t:"La pollution par les déchets ménagers",c:1,d:1},
    {n:23,s:"SÉQ 6",t:"Concept de biotechnologie et applications",c:1,d:1},
    {n:24,s:"SÉQ 6",t:"La production du biogaz",c:1,d:1},
    {n:25,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "TERM_CTI": [
    {n:1,s:"SÉQ 1",t:"Cellule en microscopie optique",c:1,d:1},
    {n:2,s:"SÉQ 1",t:"Cellule en microscopie électronique",c:1,d:1},
    {n:3,s:"SÉQ 1",t:"Échanges d'eau : l'osmose et la dialyse",c:1,d:1},
    {n:4,s:"SÉQ 1",t:"Échanges de substances dissoutes",c:1,d:1},
    {n:5,s:"SÉQ 1",t:"Échanges de particules",c:1,d:1},
    {n:6,s:"SÉQ 2",t:"Structure des acides nucléiques",c:1,d:1},
    {n:7,s:"SÉQ 2",t:"L'étude de la biosynthèse des protéines",c:1,d:1},
    {n:8,s:"SÉQ 2",t:"Structure, rôle des gonades et méiose",c:1,d:1},
    {n:9,s:"SÉQ 2",t:"Gamétogenèse et fécondation",c:1,d:1},
    {n:10,s:"SÉQ 2",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:11,s:"SÉQ 3",t:"Formes alléliques et brassage interchromosomique (mono)",c:1,d:1},
    {n:12,s:"SÉQ 3",t:"Brassage interchromosomique (dihybridisme)",c:1,d:1},
    {n:13,s:"SÉQ 3",t:"Brassage intrachromosomique",c:1,d:1},
    {n:14,s:"SÉQ 3",t:"Hérédité autosomique",c:1,d:1},
    {n:15,s:"SÉQ 3",t:"Hérédité gonosomique et mutations",c:1,d:1},
    {n:16,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:17,s:"SÉQ 4",t:"Le tissu nerveux",c:1,d:1},
    {n:18,s:"SÉQ 4",t:"Mise en évidence du potentiel de repos et d'action",c:1,d:1},
    {n:19,s:"SÉQ 4",t:"La conduction du message nerveux",c:1,d:1},
    {n:20,s:"SÉQ 4",t:"La notion de synapse",c:1,d:1},
    {n:21,s:"SÉQ 5",t:"L'origine des cellules immunitaires",c:1,d:1},
    {n:22,s:"SÉQ 5",t:"Les mécanismes de la réponse non spécifique",c:1,d:1},
    {n:23,s:"SÉQ 5",t:"La réponse à médiation humorale",c:1,d:1},
    {n:24,s:"SÉQ 5",t:"Les maladies auto-immunes",c:1,d:1},
    {n:25,s:"SÉQ 5",t:"Le VIH/Sida et ses conséquences",c:1,d:1},
    {n:26,s:"SÉQ 6",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:27,s:"SÉQ 6",t:"Régulation des hormones sexuelles chez l'homme",c:1,d:1},
    {n:28,s:"SÉQ 6",t:"Régulation des hormones sexuelles chez la femme",c:1,d:1},
    {n:29,s:"SÉQ 6",t:"L'infertilité",c:1,d:1},
    {n:30,s:"SÉQ 6",t:"La glycémie : facteurs de variation",c:1,d:1},
    {n:31,s:"SÉQ 7",t:"La glycémie : conséquences",c:1,d:1},
    {n:32,s:"SÉQ 7",t:"La pression artérielle : mesure",c:1,d:1},
    {n:33,s:"SÉQ 7",t:"La pression artérielle : conséquences",c:1,d:1},
    {n:34,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:35,s:"SÉQ 7",t:"Définitions et classification des catastrophes",c:1,d:1},
    {n:36,s:"SÉQ 8",t:"La gestion des catastrophes",c:1,d:1},
    {n:37,s:"SÉQ 8",t:"Transformation des fruits : cas de la mangue",c:1,d:1},
    {n:38,s:"SÉQ 8",t:"Transformation des fruits : cas de la tomate",c:1,d:1},
    {n:39,s:"SÉQ 8",t:"Énergies renouvelables : biocarburants",c:1,d:1},
    {n:40,s:"SÉQ 8",t:"Énergies renouvelables : biocarburants (suite)",c:1,d:1},
    {n:41,s:"SÉQ 9",t:"Valorisation des déchets papiers",c:1,d:1},
    {n:42,s:"SÉQ 9",t:"Valorisation des déchets plastiques",c:1,d:1},
    {n:43,s:"SÉQ 9",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "TERM_A": [
    {n:1,s:"SÉQ 1",t:"Cellule en microscopie optique",c:1,d:1},
    {n:2,s:"SÉQ 1",t:"Organisation de la cellule en microscopie électronique",c:1,d:1},
    {n:3,s:"SÉQ 1",t:"Principaux organites cellulaires et leurs rôles",c:1,d:1},
    {n:4,s:"SÉQ 1",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:5,s:"SÉQ 1",t:"Structure des acides nucléiques et duplication",c:1,d:1},
    {n:6,s:"SÉQ 2",t:"Étude de la biosynthèse des protéines (étapes)",c:1,d:1},
    {n:7,s:"SÉQ 2",t:"Étude de la biosynthèse des protéines (relation)",c:1,d:1},
    {n:8,s:"SÉQ 2",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:9,s:"SÉQ 2",t:"De la diploïdie à l'haploïdie : les étapes de la méiose",c:1,d:1},
    {n:10,s:"SÉQ 2",t:"Nécessité de la méiose dans la pérennisation",c:1,d:1},
    {n:11,s:"SÉQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:12,s:"SÉQ 3",t:"Absorption intestinale",c:1,d:1},
    {n:13,s:"SÉQ 3",t:"Devenir et rôle des éléments absorbés",c:1,d:1},
    {n:14,s:"SÉQ 3",t:"Sort des résidus de la digestion",c:1,d:1},
    {n:15,s:"SÉQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:16,s:"SÉQ 4",t:"Constituants du sang et leurs rôles",c:1,d:1},
    {n:17,s:"SÉQ 4",t:"Milieu intérieur",c:1,d:1},
    {n:18,s:"SÉQ 4",t:"Nécessité du maintien de la constance",c:1,d:1},
    {n:19,s:"SÉQ 4",t:"Importance de l'élimination urinaire",c:1,d:1},
    {n:20,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:21,s:"SÉQ 5",t:"Rôle du système nerveux dans les relations",c:1,d:1},
    {n:22,s:"SÉQ 5",t:"Comportements psychosociaux",c:1,d:1},
    {n:23,s:"SÉQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:24,s:"SÉQ 5",t:"Grossesse",c:1,d:1},
    {n:25,s:"SÉQ 5",t:"Stérilité",c:1,d:1},
    {n:26,s:"SÉQ 6",t:"La procréation médicalement assistée",c:1,d:1},
    {n:27,s:"SÉQ 6",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:28,s:"SÉQ 6",t:"Dysfonctionnements du système immunitaire",c:1,d:1},
    {n:29,s:"SÉQ 6",t:"Le VIH-Sida et conséquences socioculturelles",c:1,d:1},
    {n:30,s:"SÉQ 6",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:31,s:"SÉQ 7",t:"Définition et classification des catastrophes",c:1,d:1},
    {n:32,s:"SÉQ 7",t:"La gestion des catastrophes",c:1,d:1},
    {n:33,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:34,s:"SÉQ 7",t:"Transformation et conservation des fruits : la mangue",c:1,d:1},
    {n:35,s:"SÉQ 7",t:"Transformation et conservation des fruits : la tomate",c:1,d:1},
    {n:36,s:"SÉQ 8",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:37,s:"SÉQ 8",t:"Valorisation des déchets papiers",c:1,d:1},
    {n:38,s:"SÉQ 8",t:"Valorisation des déchets plastiques",c:1,d:1},
    {n:39,s:"SÉQ 8",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "SIX": [
    {n:1,s:"SÉQ 1",t:"Influence du climat sur la production végétale",c:1,d:1},
    {n:2,s:"SÉQ 1",t:"Influence du sol sur la production végétale",c:1,d:1},
    {n:3,s:"SÉQ 2",t:"Multiplication par voie sexuée des plantes",c:1,d:1},
    {n:4,s:"SÉQ 2",t:"Multiplication végétative",c:1,d:1},
    {n:5,s:"SÉQ 2",t:"Rôle de la reproduction dans l'amélioration",c:1,d:1},
    {n:6,s:"SÉQ 3",t:"Sols et production végétale (compost)",c:1,d:1},
    {n:7,s:"SÉQ 3",t:"Moyens de lutte contre les parasites",c:1,d:1},
    {n:8,s:"SÉQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:9,s:"SÉQ 4",t:"Rôle des microorganismes : fermentations",c:1,d:1},
    {n:10,s:"SÉQ 4",t:"Transformation de quelques produits agricoles",c:1,d:1},
    {n:11,s:"SÉQ 4",t:"Technique d'extraction d'une huile végétale",c:1,d:1},
    {n:12,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:13,s:"SÉQ 5",t:"Les caractéristiques de la matière",c:1,d:1},
    {n:14,s:"SÉQ 5",t:"Mesure et calcul de quelques caractéristiques",c:1,d:1},
    {n:15,s:"SÉQ 5",t:"Les propriétés physiques de la matière",c:1,d:1},
    {n:16,s:"SÉQ 5",t:"Les propriétés chimiques de la matière",c:1,d:1},
    {n:17,s:"SÉQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:18,s:"SÉQ 6",t:"Notion d'énergie",c:1,d:1},
    {n:19,s:"SÉQ 6",t:"Formes et sources d'énergie",c:1,d:1},
    {n:20,s:"SÉQ 7",t:"Énergie et environnement",c:1,d:1},
    {n:21,s:"SÉQ 7",t:"Les modes de transferts de chaleur",c:1,d:1},
    {n:22,s:"SÉQ 8",t:"Notion de circuit électrique",c:1,d:1},
    {n:23,s:"SÉQ 9",t:"La lumière",c:1,d:1},
    {n:24,s:"SÉQ 10",t:"Énergie mécanique",c:1,d:1},
    {n:25,s:"SÉQ 10",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:26,s:"SÉQ 11",t:"La puberté, la fécondation et la grossesse",c:1,d:1},
    {n:27,s:"SÉQ 11",t:"Les grossesses précoces",c:1,d:1},
    {n:28,s:"SÉQ 11",t:"Les pratiques culturelles néfastes",c:1,d:1},
    {n:29,s:"SÉQ 11",t:"Les IST et le VIH/SIDA",c:1,d:1},
    {n:30,s:"SÉQ 11",t:"Hygiène des organes reproducteurs",c:1,d:1},
    {n:31,s:"SÉQ 11",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:32,s:"SÉQ 12",t:"Les différentes catégories d'aliments",c:1,d:1},
    {n:33,s:"SÉQ 12",t:"Les maladies nutritionnelles de carence",c:1,d:1},
    {n:34,s:"SÉQ 12",t:"Les maladies nutritionnelles de l'excès",c:1,d:1},
    {n:35,s:"SÉQ 12",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:36,s:"SÉQ 13",t:"L'eau dans l'environnement",c:1,d:1},
    {n:37,s:"SÉQ 13",t:"La pollution de l'eau",c:1,d:1},
    {n:38,s:"SÉQ 13",t:"Lutte contre la pollution des eaux",c:1,d:1},
    {n:39,s:"SÉQ 13",t:"Les différents usages de l'eau",c:1,d:1},
    {n:40,s:"SÉQ 13",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:41,s:"SÉQ 14",t:"La pollution des sols",c:1,d:1},
    {n:42,s:"SÉQ 14",t:"Les moyens de lutte contre la pollution des sols",c:1,d:1},
    {n:43,s:"SÉQ 14",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:44,s:"SÉQ 15",t:"Les plantes médicinales",c:1,d:1},
    {n:45,s:"SÉQ 15",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:46,s:"SÉQ 16",t:"Technologie : définition et démarche",c:1,d:1},
    {n:47,s:"SÉQ 17",t:"Objet technique",c:1,d:1},
    {n:48,s:"SÉQ 18",t:"Cahier des charges",c:1,d:1},
    {n:49,s:"SÉQ 19",t:"Notion de projet",c:1,d:1},
    {n:50,s:"SÉQ 19",t:"Exemple : fabrication d'un filtre",c:1,d:1},
    {n:51,s:"SÉQ 19",t:"Exemple : création d'un jardin potager",c:1,d:1},
    {n:52,s:"SÉQ 19",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "CINQ": [
    {n:1,s:"SÉQ 1",t:"Influence de l'espace vital et compétition",c:1,d:1},
    {n:2,s:"SÉQ 1",t:"Reproduction sexuée des animaux",c:1,d:1},
    {n:3,s:"SÉQ 1",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:4,s:"SÉQ 1",t:"Sélection des espèces et reproduction croisée",c:1,d:1},
    {n:5,s:"SÉQ 1",t:"Lutte contre les parasites des animaux",c:1,d:1},
    {n:6,s:"SÉQ 2",t:"La transformation du lait en fromage et beurre",c:1,d:1},
    {n:7,s:"SÉQ 2",t:"La transformation de la viande en saucisse",c:1,d:1},
    {n:8,s:"SÉQ 2",t:"Le fumage, le séchage, le saumurage",c:1,d:1},
    {n:9,s:"SÉQ 2",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:10,s:"SÉQ 2",t:"Notions de masse volumique et de densité",c:1,d:1},
    {n:11,s:"SÉQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:12,s:"SÉQ 3",t:"La vaporisation et la condensation",c:1,d:1},
    {n:13,s:"SÉQ 3",t:"La solidification, la liquéfaction, la sublimation",c:1,d:1},
    {n:14,s:"SÉQ 3",t:"Définitions, exemples de mélanges",c:1,d:1},
    {n:15,s:"SÉQ 3",t:"Séparation des constituants d'un mélange",c:1,d:1},
    {n:16,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:17,s:"SÉQ 4",t:"Position d'un mobile dans un repère",c:1,d:1},
    {n:18,s:"SÉQ 4",t:"Relativité du mouvement et trajectoire",c:1,d:1},
    {n:19,s:"SÉQ 4",t:"Notions de vitesse moyenne et instantanée",c:1,d:1},
    {n:20,s:"SÉQ 4",t:"Diagrammes des espaces et des vitesses",c:1,d:1},
    {n:21,s:"SÉQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:22,s:"SÉQ 5",t:"Les comportements émergents néfastes",c:1,d:1},
    {n:23,s:"SÉQ 5",t:"Le VIH/Sida",c:1,d:1},
    {n:24,s:"SÉQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:25,s:"SÉQ 5",t:"Rations alimentaires équilibrées",c:1,d:1},
    {n:26,s:"SÉQ 6",t:"Définition et germes impliqués (intoxications)",c:1,d:1},
    {n:27,s:"SÉQ 6",t:"Les règles d'hygiène alimentaire",c:1,d:1},
    {n:28,s:"SÉQ 6",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:29,s:"SÉQ 6",t:"L'air et les polluants atmosphériques",c:1,d:1},
    {n:30,s:"SÉQ 6",t:"Causes et conséquences de l'effet de serre",c:1,d:1},
    {n:31,s:"SÉQ 7",t:"Localisation et protection de la couche d'ozone",c:1,d:1},
    {n:32,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:33,s:"SÉQ 7",t:"Origines et types de roches sédimentaires",c:1,d:1},
    {n:34,s:"SÉQ 7",t:"Intérêts des roches sédimentaires",c:1,d:1},
    {n:35,s:"SÉQ 7",t:"Relation entre propriétés et utilisation",c:1,d:1},
    {n:36,s:"SÉQ 8",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:37,s:"SÉQ 8",t:"Recensement des espèces animales",c:1,d:1},
    {n:38,s:"SÉQ 8",t:"Classification sommaire et causes de disparition",c:1,d:1},
    {n:39,s:"SÉQ 8",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:40,s:"SÉQ 8",t:"Étude d'un objet technique",c:1,d:1},
    {n:41,s:"SÉQ 9",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:42,s:"SÉQ 9",t:"Élevage des poulets",c:1,d:1},
    {n:43,s:"SÉQ 9",t:"Élevage du tilapia",c:1,d:1},
    {n:44,s:"SÉQ 9",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "QUATRE": [
    {n:1,s:"SÉQ 1",t:"Les besoins nutritifs des végétaux",c:1,d:1},
    {n:2,s:"SÉQ 1",t:"Les besoins nutritifs des animaux",c:1,d:1},
    {n:3,s:"SÉQ 1",t:"La production de la matière végétale",c:1,d:1},
    {n:4,s:"SÉQ 1",t:"La production de la matière animale",c:1,d:1},
    {n:5,s:"SÉQ 1",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:6,s:"SÉQ 2",t:"Recensement des espèces et notion de biodiversité",c:1,d:1},
    {n:7,s:"SÉQ 2",t:"Diminution de la biodiversité",c:1,d:1},
    {n:8,s:"SÉQ 2",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:9,s:"SÉQ 2",t:"Les organes de l'appareil moteur",c:1,d:1},
    {n:10,s:"SÉQ 2",t:"Accidents des muscles et secourisme",c:1,d:1},
    {n:11,s:"SÉQ 3",t:"Accidents des os et articulations",c:1,d:1},
    {n:12,s:"SÉQ 3",t:"Rôle de l'alimentation sur l'appareil moteur",c:1,d:1},
    {n:13,s:"SÉQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:14,s:"SÉQ 3",t:"La peau : structure, rôle et fonctions",c:1,d:1},
    {n:15,s:"SÉQ 3",t:"Maladies liées au décapage de la peau",c:1,d:1},
    {n:16,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:17,s:"SÉQ 4",t:"Organisation sommaire du système nerveux",c:1,d:1},
    {n:18,s:"SÉQ 4",t:"La fatigue nerveuse",c:1,d:1},
    {n:19,s:"SÉQ 4",t:"Les toxicomanies",c:1,d:1},
    {n:20,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:21,s:"SÉQ 5",t:"Organisation de l'appareil digestif",c:1,d:1},
    {n:22,s:"SÉQ 5",t:"Rôle de l'appareil digestif",c:1,d:1},
    {n:23,s:"SÉQ 5",t:"Importance et hygiène de la digestion",c:1,d:1},
    {n:24,s:"SÉQ 5",t:"Les maladies par excès et par carence",c:1,d:1},
    {n:25,s:"SÉQ 5",t:"Les maladies du péril fécal",c:1,d:1},
    {n:26,s:"SÉQ 6",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:27,s:"SÉQ 6",t:"Étude de quelques IST",c:1,d:1},
    {n:28,s:"SÉQ 6",t:"Le VIH/SIDA",c:1,d:1},
    {n:29,s:"SÉQ 6",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:30,s:"SÉQ 6",t:"Les différentes phases d'une éruption volcanique",c:1,d:1},
    {n:31,s:"SÉQ 7",t:"Les différents types de dynamisme volcanique",c:1,d:1},
    {n:32,s:"SÉQ 7",t:"Du magma aux roches volcaniques",c:1,d:1},
    {n:33,s:"SÉQ 7",t:"Les types de risques volcaniques",c:1,d:1},
    {n:34,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:35,s:"SÉQ 7",t:"Quelques exemples d'énergies fossiles",c:1,d:1},
    {n:36,s:"SÉQ 8",t:"Le pétrole",c:1,d:1},
    {n:37,s:"SÉQ 8",t:"Les énergies fossiles et le développement durable",c:1,d:1},
    {n:38,s:"SÉQ 8",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:39,s:"SÉQ 8",t:"Les ressources minières du Cameroun",c:1,d:1},
    {n:40,s:"SÉQ 8",t:"Gestion des ressources minières",c:1,d:1},
    {n:41,s:"SÉQ 9",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:42,s:"SÉQ 9",t:"Biodiversité dans un écosystème aquatique",c:1,d:1},
    {n:43,s:"SÉQ 9",t:"Activités humaines détruisant les écosystèmes",c:1,d:1},
    {n:44,s:"SÉQ 9",t:"Restauration et conservation de la biodiversité",c:1,d:1},
    {n:45,s:"SÉQ 9",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "TROIS": [
    {n:1,s:"SÉQ 1",t:"Ressemblances entre les individus",c:1,d:1},
    {n:2,s:"SÉQ 1",t:"Différences entre les individus",c:1,d:1},
    {n:3,s:"SÉQ 1",t:"Localisation de l'information génétique",c:1,d:1},
    {n:4,s:"SÉQ 1",t:"Nature de l'information génétique : chromosome et ADN",c:1,d:1},
    {n:5,s:"SÉQ 1",t:"Nature de l'information génétique : chromosomes humains",c:1,d:1},
    {n:6,s:"SÉQ 2",t:"Les gènes humains",c:1,d:1},
    {n:7,s:"SÉQ 2",t:"Gènes et diversité humaine (groupes sanguins)",c:1,d:1},
    {n:8,s:"SÉQ 2",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:9,s:"SÉQ 2",t:"Différents groupes de microorganismes",c:1,d:1},
    {n:10,s:"SÉQ 2",t:"Mode de vie des microbes : reproduction",c:1,d:1},
    {n:11,s:"SÉQ 3",t:"Mode de vie des microbes : nutrition et respiration",c:1,d:1},
    {n:12,s:"SÉQ 3",t:"Contamination par les microorganismes",c:1,d:1},
    {n:13,s:"SÉQ 3",t:"Pratiques pour éviter la contamination",c:1,d:1},
    {n:14,s:"SÉQ 3",t:"La réponse immunitaire non spécifique (peau, muqueuses)",c:1,d:1},
    {n:15,s:"SÉQ 3",t:"La réponse immunitaire non spécifique (phagocytose)",c:1,d:1},
    {n:16,s:"SÉQ 4",t:"La réponse immunitaire spécifique",c:1,d:1},
    {n:17,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:18,s:"SÉQ 4",t:"La multiplication du VIH",c:1,d:1},
    {n:19,s:"SÉQ 4",t:"VIH/SIDA : phases, prévention et traitement",c:1,d:1},
    {n:20,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:21,s:"SÉQ 5",t:"Antibiothérapie et sérothérapie",c:1,d:1},
    {n:22,s:"SÉQ 5",t:"Vaccinothérapie",c:1,d:1},
    {n:23,s:"SÉQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:24,s:"SÉQ 5",t:"Siège de la circulation : les vaisseaux",c:1,d:1},
    {n:25,s:"SÉQ 5",t:"Siège de la circulation : le cœur",c:1,d:1},
    {n:26,s:"SÉQ 6",t:"Hygiène de la circulation : hémorragies",c:1,d:1},
    {n:27,s:"SÉQ 6",t:"Hygiène de la circulation : maladies cardiovasculaires",c:1,d:1},
    {n:28,s:"SÉQ 6",t:"Soins de premiers secours (AVC, hémorragie)",c:1,d:1},
    {n:29,s:"SÉQ 6",t:"Anatomie de l'œil, anomalies et maladies",c:1,d:1},
    {n:30,s:"SÉQ 6",t:"Hygiène de la vision",c:1,d:1},
    {n:31,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:32,s:"SÉQ 7",t:"Un exemple d'endémie : le paludisme",c:1,d:1},
    {n:33,s:"SÉQ 7",t:"Quelques exemples d'épidémies",c:1,d:1},
    {n:34,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:35,s:"SÉQ 7",t:"Manifestations et origine des séismes",c:1,d:1},
    {n:36,s:"SÉQ 8",t:"Localisation des séismes",c:1,d:1},
    {n:37,s:"SÉQ 8",t:"Causes des risques liés aux mouvements de terrain",c:1,d:1},
    {n:38,s:"SÉQ 8",t:"Causes : action mécanique de l'eau",c:1,d:1},
    {n:39,s:"SÉQ 8",t:"Techniques de prévention des accidents",c:1,d:1},
    {n:40,s:"SÉQ 8",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:41,s:"SÉQ 9",t:"Biodiversité dans les écosystèmes",c:1,d:1},
    {n:42,s:"SÉQ 9",t:"Activités humaines détruisant les écosystèmes",c:1,d:1},
    {n:43,s:"SÉQ 9",t:"Restauration et conservation de la biodiversité",c:1,d:1},
    {n:44,s:"SÉQ 9",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "SEC_C": [
    {n:1,s:"SÉQ 1",t:"La cellule chlorophyllienne, usine photosynthétique",c:1,d:1},
    {n:2,s:"SÉQ 1",t:"L'influence de certains facteurs sur la production végétale",c:1,d:1},
    {n:3,s:"SÉQ 1",t:"Les plantes performantes",c:1,d:1},
    {n:4,s:"SÉQ 1",t:"Les sociétés des fourmis, termites, abeilles",c:1,d:1},
    {n:5,s:"SÉQ 1",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:6,s:"SÉQ 2",t:"Définition de la notion de plan d'organisation",c:1,d:1},
    {n:7,s:"SÉQ 2",t:"Déduction de la notion d'homologie, analogie",c:1,d:1},
    {n:8,s:"SÉQ 2",t:"Liens de parenté entre Homme et autres espèces",c:1,d:1},
    {n:9,s:"SÉQ 2",t:"La transmission de l'information génétique : mitose",c:1,d:1},
    {n:10,s:"SÉQ 2",t:"Structure de l'appareil respiratoire",c:1,d:1},
    {n:11,s:"SÉQ 3",t:"Échanges gazeux respiratoires",c:1,d:1},
    {n:12,s:"SÉQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:13,s:"SÉQ 3",t:"Variations des paramètres physiologiques (1)",c:1,d:1},
    {n:14,s:"SÉQ 3",t:"Variations des paramètres physiologiques (2)",c:1,d:1},
    {n:15,s:"SÉQ 3",t:"Structure et rôle de l'appareil urinaire",c:1,d:1},
    {n:16,s:"SÉQ 4",t:"Insuffisances rénales",c:1,d:1},
    {n:17,s:"SÉQ 4",t:"Définition, mécanisme, causes des allergies",c:1,d:1},
    {n:18,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:19,s:"SÉQ 4",t:"Complications des allergies",c:1,d:1},
    {n:20,s:"SÉQ 4",t:"La filariose lymphatique et l'onchocercose",c:1,d:1},
    {n:21,s:"SÉQ 5",t:"La dracunculose et la téniase",c:1,d:1},
    {n:22,s:"SÉQ 5",t:"Les schistosomiases et l'ulcère de Buruli",c:1,d:1},
    {n:23,s:"SÉQ 5",t:"La fièvre jaune",c:1,d:1},
    {n:24,s:"SÉQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:25,s:"SÉQ 5",t:"Situation de la planète Terre",c:1,d:1},
    {n:26,s:"SÉQ 6",t:"Le réchauffement climatique",c:1,d:1},
    {n:27,s:"SÉQ 6",t:"La mangrove et l'écosystème aquatique",c:1,d:1},
    {n:28,s:"SÉQ 6",t:"Restauration et conservation de la biodiversité",c:1,d:1},
    {n:29,s:"SÉQ 6",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:30,s:"SÉQ 6",t:"Les déchets biodégradables et non biodégradables",c:1,d:1},
    {n:31,s:"SÉQ 7",t:"Le génie génétique",c:1,d:1},
    {n:32,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "SEC_A": [
    {n:1,s:"SÉQ 1",t:"Cellule chlorophyllienne, usine photosynthétique",c:1,d:1},
    {n:2,s:"SÉQ 1",t:"Cellule chlorophyllienne et approvisionnement",c:1,d:1},
    {n:3,s:"SÉQ 1",t:"L'influence de certains facteurs sur la production",c:1,d:1},
    {n:4,s:"SÉQ 1",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:5,s:"SÉQ 2",t:"Les plantes performantes",c:1,d:1},
    {n:6,s:"SÉQ 2",t:"La transmission de l'information génétique",c:1,d:1},
    {n:7,s:"SÉQ 2",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:8,s:"SÉQ 2",t:"Structure de l'appareil respiratoire",c:1,d:1},
    {n:9,s:"SÉQ 3",t:"Échanges gazeux respiratoires",c:1,d:1},
    {n:10,s:"SÉQ 3",t:"Hygiène de l'appareil respiratoire",c:1,d:1},
    {n:11,s:"SÉQ 3",t:"Variations des paramètres physiologiques (1)",c:1,d:1},
    {n:12,s:"SÉQ 3",t:"Variations des paramètres physiologiques (2)",c:1,d:1},
    {n:13,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:14,s:"SÉQ 4",t:"Structure et rôle de l'appareil urinaire",c:1,d:1},
    {n:15,s:"SÉQ 4",t:"Insuffisances rénales",c:1,d:1},
    {n:16,s:"SÉQ 4",t:"Cycles sexuels",c:1,d:1},
    {n:17,s:"SÉQ 5",t:"Formation des gamètes et fécondation",c:1,d:1},
    {n:18,s:"SÉQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:19,s:"SÉQ 5",t:"Apprentissage de l'intégration (suite)",c:1,d:0},
    {n:20,s:"SÉQ 5",t:"Les comportements à risque et le VIH/Sida",c:1,d:1},
    {n:21,s:"SÉQ 6",t:"Caractéristiques de la planète Terre",c:1,d:1},
    {n:22,s:"SÉQ 6",t:"Le réchauffement climatique",c:1,d:1},
    {n:23,s:"SÉQ 6",t:"La mangrove",c:1,d:1},
    {n:24,s:"SÉQ 6",t:"Activités humaines détruisant les écosystèmes",c:1,d:1},
    {n:25,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:26,s:"SÉQ 7",t:"Les étapes du génie génétique",c:1,d:1},
    {n:27,s:"SÉQ 7",t:"Un exemple d'application du génie génétique",c:1,d:1},
    {n:28,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
  ],
  "TERM_C": [
    {n:1,s:"SÉQ 1",t:"Prise de contact ; présentation du programme ; évaluation diagnostique",c:1,d:0},
    {n:2,s:"SÉQ 1",t:"Cellule en microscopie optique",c:1,d:1},
    {n:3,s:"SÉQ 1",t:"Organisation de la cellule en microscopie électronique",c:1,d:1},
    {n:4,s:"SÉQ 1",t:"Principaux organites cellulaires et leurs rôles",c:1,d:1},
    {n:5,s:"SÉQ 1",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:6,s:"SÉQ 2",t:"Structure des acides nucléiques et duplication de l'ADN",c:1,d:1},
    {n:7,s:"SÉQ 2",t:"Étude de la biosynthèse des protéines (étapes, localisation, acteurs et produits)",c:1,d:1},
    {n:8,s:"SÉQ 2",t:"Étude de la biosynthèse des protéines (la relation entre la protéine et le caractère)",c:1,d:1},
    {n:9,s:"SÉQ 2",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:10,s:"SÉQ 3",t:"De la diploïdie à l'haploïdie : les étapes de la méiose",c:1,d:1},
    {n:11,s:"SÉQ 3",t:"Nécessité de la méiose dans la pérennisation de l'espèce",c:1,d:1},
    {n:12,s:"SÉQ 3",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:13,s:"SÉQ 3",t:"Évaluation / Correction, remise des copies et remédiations",c:1,d:0},
    {n:14,s:"SÉQ 4",t:"Absorption intestinale : définition et troubles y afférents",c:1,d:1},
    {n:15,s:"SÉQ 4",t:"Devenir et le rôle des éléments absorbés",c:1,d:1},
    {n:16,s:"SÉQ 4",t:"Sort des résidus de la digestion : élimination des déchets",c:1,d:1},
    {n:17,s:"SÉQ 4",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:18,s:"SÉQ 5",t:"Constituants du sang et leurs rôles",c:1,d:1},
    {n:19,s:"SÉQ 5",t:"Milieu intérieur",c:1,d:1},
    {n:20,s:"SÉQ 5",t:"Nécessité du maintien de la constance du milieu intérieur",c:1,d:1},
    {n:21,s:"SÉQ 5",t:"Importance de l'élimination urinaire",c:1,d:1},
    {n:22,s:"SÉQ 5",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:23,s:"SÉQ 6",t:"Rôle du système nerveux dans les relations interpersonnelles",c:1,d:1},
    {n:24,s:"SÉQ 6",t:"Comportements psychosociaux dans la gestion d'un malade contagieux et d'une épidémie",c:1,d:1},
    {n:25,s:"SÉQ 6",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:26,s:"SÉQ 7",t:"Grossesse",c:1,d:1},
    {n:27,s:"SÉQ 7",t:"Stérilité",c:1,d:1},
    {n:28,s:"SÉQ 7",t:"La procréation médicalement assistée",c:1,d:1},
    {n:29,s:"SÉQ 7",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:30,s:"SÉQ 8",t:"Dysfonctionnements du système immunitaire : les maladies auto-immunes",c:1,d:1},
    {n:31,s:"SÉQ 8",t:"Le VIH-Sida et conséquences socioculturelles",c:1,d:1},
    {n:32,s:"SÉQ 8",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:33,s:"SÉQ 9",t:"Définition et classification des catastrophes",c:1,d:1},
    {n:34,s:"SÉQ 9",t:"La gestion des catastrophes (prévision, prédiction, prévention)",c:1,d:1},
    {n:35,s:"SÉQ 9",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:36,s:"SÉQ 10",t:"Transformation et conservation de fruits : cas de la mangue",c:1,d:1},
    {n:37,s:"SÉQ 10",t:"Transformation et conservation de fruits : cas de la tomate",c:1,d:1},
    {n:38,s:"SÉQ 10",t:"Apprentissage de l'intégration",c:1,d:0},
    {n:39,s:"SÉQ 11",t:"Valorisation des déchets papiers et plastiques",c:1,d:1},
  ],
};

function ProgrammePage() {
  // Vue Animatrice — Suivi programme global
  // Utilise le contexte AppContext (données partagées avec tout AppLayout)
  const {data} = useApp();

  if (!data) return (
    <div style={{padding:"60px",textAlign:"center",color:C.txtMuted}}>
      <Spinner size={28} color={C.green}/>
      <div style={{marginTop:12,fontSize:13}}>Chargement des données…</div>
    </div>
  );

  return <EnseignantsPage/>;
}

// const EP_SLOTS: défini globalement ligne ~4092
// constante définie globalement

// [DEDUPLIQUÉ] const ENS_COLORS = { mbassam:"#1a6b3c",boubam:"#c0392b",douniaroud:"#2980b9",hayatouh:"#16a085",aissatous:"#8e44ad",essambas:"#d35400",koffa:"#27ae60",mawiyak:"#e67e22",sadjot:"#2c3e50",sylvie:"#c8a951" };
// const getColor = id => ENS_COLORS[id] || "#1a6b3c"; — défini globalement
// getIni: défini globalement

// ═══════════════════════════════════════════════════
// PALETTE
// ═══════════════════════════════════════════════════
// const C: référence au bloc global ci-dessus

const STATUT_CONFIG = {
  attente:  { label:"En attente",  emoji:"⏳", color:C.orange,  bg:C.orangePale, border:C.orangeBorder },
  validee:  { label:"Validée",     emoji:"✅", color:C.green,   bg:C.greenPale,  border:C.greenBorder  },
  rejetee:  { label:"Rejetée",     emoji:"❌", color:C.red,     bg:C.redPale,    border:C.redBorder    },
  vide:     { label:"Non soumise", emoji:"○",  color:C.txtLight,bg:"#f8fafc",    border:C.border       },
};

// ═══════════════════════════════════════════════════
// COMPOSANTS UI
// ═══════════════════════════════════════════════════
// EpSpinner: utilise Spinner global

// (composant Ep dédupliqué)

// (composant Ep dédupliqué)

const StatutBadge = ({ statut }) => {
  const cfg = STATUT_CONFIG[statut] || STATUT_CONFIG.vide;
  return <Pill ch={`${cfg.emoji} ${cfg.label}`} color={cfg.color} bg={cfg.bg}/>;
};

// ═══════════════════════════════════════════════════
// CHARGEMENT INITIAL
// ═══════════════════════════════════════════════════
async function loadData() {
  await syncElevesImport();
  const [epreuves, users, classes, prog, notes, absences] = await Promise.all([
    sb.get("epreuves", ""),
    sb.get("utilisateurs", "?select=id,nom,role,classes,photo"),
    sb.get("classes", "?select=code,effectif&order=code"),
    sb.get("prog_suivi", "?select=ens_id,classe,faites"),
    sb.get("notes", "?select=classe,evaluation,eleve_id,note"),
    sb.get("absences", "?select=ens_id,classe,seance,absents"),
  ]);
  const eps = (epreuves||[]).map(e => ({
    ...e,
    stockPath: e.stockpath || e.stockPath || "",
    commentaires: Array.isArray(e.commentaires) ? e.commentaires : [],
  }));
  const usersMap = {};
  (users||[]).forEach(u => { usersMap[u.id] = { ...u, classes:u.classes||[], col:getColor(u.id), ini:getIni(u.nom) }; });
  const progIndex = {};
  (prog||[]).forEach(r=>{progIndex[`${r.ens_id}||${r.classe}`]=Array.isArray(r.faites)?r.faites:[];});
  const notesIndex = {};
  (notes||[]).forEach(r=>{
    const k=`${r.classe}||${r.evaluation}`;
    if(!notesIndex[k]) notesIndex[k]={};
    if(r.note!==null) notesIndex[k][r.eleve_id]=r.note;
  });
  const absIndex = {};
  (absences||[]).forEach(r=>{absIndex[`${r.ens_id}||${r.classe}||${r.seance}`]=Array.isArray(r.absents)?r.absents:[];});
  return { epreuves: eps, users: usersMap, classes: classes||[], prog: progIndex, notes: notesIndex, absences: absIndex };
}

// ═══════════════════════════════════════════════════
// GRILLE DE COUVERTURE (vue animateur)
// ═══════════════════════════════════════════════════
function GrilleCouverture({ classes, epreuves, onCardClick }) {
  return (
    <div style={{ display:"flex",flexDirection:"column",gap:18 }}>
      {classes.map(cl => {
        const eps = epreuves.filter(e => e.classe === cl.code);
        const nb = eps.filter(e => e.statut === "validee").length;
        const total = EP_SLOTS.length;
        return (
          <div key={cl.code} style={{ background:C.white,borderRadius:11,border:`1px solid ${C.border}`,overflow:"hidden" }}>
            {/* Header classe */}
            <div style={{ padding:"10px 14px",background:"#f8fafc",borderBottom:`1px solid ${C.border}`,display:"flex",alignItems:"center",justifyContent:"space-between" }}>
              <div>
                <span style={{ fontSize:13,fontWeight:800,color:C.txt }}>{cl.code}</span>
                <span style={{ fontSize:11,color:C.txtMuted,marginLeft:8 }}>{cl.effectif} élèves</span>
              </div>
              <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                <div style={{ width:80,height:5,background:"#e2e8f0",borderRadius:3,overflow:"hidden" }}>
                  <div style={{ height:"100%",width:`${Math.min(100, Math.round(nb/total*100))}%`,background:nb===total?C.green:nb>0?C.orange:C.border,borderRadius:3,transition:"width .5s" }}/>
                </div>
                <span style={{ fontSize:11,fontWeight:700,color:nb===total?C.green:C.txtMuted }}>{nb}/{total} validées</span>
              </div>
            </div>
            {/* Slots */}
            <div style={{ display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:0 }}>
              {EP_SLOTS.map((slot, i) => {
                const ep = eps.find(e => e.trim === slot.trim && e.num === slot.ep);
                const st = ep?.statut || "vide";
                const cfg = STATUT_CONFIG[st];
                return (
                  <div key={i} onClick={() => ep && onCardClick(ep)}
                    style={{ padding:"10px 8px",textAlign:"center",borderRight:i<5?`1px solid ${C.border}`:"none",borderLeft:"none",background:cfg.bg,cursor:ep?"pointer":"default",transition:"all .15s" }}
                    onMouseEnter={e=>{ if(ep) e.currentTarget.style.opacity=".8"; }}
                    onMouseLeave={e=>{ e.currentTarget.style.opacity="1"; }}>
                    <div style={{ fontSize:9,fontWeight:800,color:TRIM_COLORS[slot.trim],marginBottom:3,letterSpacing:".05em" }}>{slot.label}</div>
                    <div style={{ fontSize:18,lineHeight:1,marginBottom:2 }}>{cfg.emoji}</div>
                    <div style={{ fontSize:9,color:cfg.color,fontWeight:600 }}>{cfg.label}</div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ═══════════════════════════════════════════════════
// CARTE ÉPREUVE (animateur — avec actions)
// ═══════════════════════════════════════════════════
function EpreuveCard({ ep, users, onAction, isAdmin }) {
  const [comment, setComment] = useState("");
  const [sendingCom, setSendingCom] = useState(false);
  const [actioning, setActioning] = useState(null);
  const ens = users[ep.ens] || { nom: ep.ens, ini: "??", col: C.txtLight };
  const cfg = STATUT_CONFIG[ep.statut] || STATUT_CONFIG.vide;

  const handleAction = async (statut) => {
    setActioning(statut);
    await onAction(ep.id, statut, null);
    setActioning(null);
  };

  const handleComment = async () => {
    if (!comment.trim()) return;
    setSendingCom(true);
    await onAction(ep.id, null, comment.trim());
    setComment("");
    setSendingCom(false);
  };

  return (
    <div style={{ background:C.white,borderRadius:12,border:`1.5px solid ${cfg.border}`,overflow:"hidden",animation:"fadeUp .3s ease" }}>
      {/* Header */}
      <div style={{ padding:"14px 16px",background:cfg.bg,borderBottom:`1px solid ${cfg.border}`,display:"flex",alignItems:"flex-start",gap:12 }}>
        {/* Icône type */}
        <div style={{ width:42,height:42,borderRadius:10,background:ep.type==="pdf"?C.redPale:C.bluePale,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0 }}>
          {ep.type==="pdf"?"📕":"📘"}
        </div>
        {/* Infos */}
        <div style={{ flex:1,minWidth:0 }}>
          <div style={{ fontSize:14,fontWeight:800,color:C.txt }}>
            {ep.trim}·{ep.num} — {ep.classe} · SVTEEHB
          </div>
          <div style={{ display:"flex",flexWrap:"wrap",gap:10,marginTop:5 }}>
            <div style={{ display:"flex",alignItems:"center",gap:6 }}>
              <Avatar ens={ens} size={20} fontSize={8}/>
              <span style={{ fontSize:12,color:C.txtMuted }}>{ens.nom}</span>
            </div>
            <span style={{ fontSize:12,color:C.txtMuted }}>📅 {fmtDateFr(ep.soumis)}</span>
            {ep.fichier && <span style={{ fontSize:12,color:C.txtMuted }}>📄 {ep.fichier}</span>}
            {ep.taille && <span style={{ fontSize:12,color:C.txtMuted }}>{ep.taille}</span>}
          </div>
        </div>
        <StatutBadge statut={ep.statut}/>
      </div>

      {/* Actions animateur */}
      {isAdmin && ep.statut === "attente" && (
        <div style={{ padding:"12px 16px",display:"flex",gap:8,borderBottom:`1px solid ${C.border}`,background:"rgba(255,255,255,.7)" }}>
          <button onClick={() => handleAction("validee")} disabled={!!actioning}
            style={{ flex:1,padding:"9px",background:actioning==="validee"?C.green:`linear-gradient(135deg,${C.greenDark},${C.greenLight})`,color:"#fff",border:"none",borderRadius:9,fontSize:13,fontWeight:700,cursor:actioning?"not-allowed":"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:7 }}>
            {actioning==="validee" ? <><Spinner s={14}/>Validation…</> : "✓ Valider"}
          </button>
          <button onClick={() => handleAction("rejetee")} disabled={!!actioning}
            style={{ flex:1,padding:"9px",background:actioning==="rejetee"?C.red:C.redPale,color:actioning==="rejetee"?"#fff":C.red,border:`1.5px solid ${C.redBorder}`,borderRadius:9,fontSize:13,fontWeight:700,cursor:actioning?"not-allowed":"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:7 }}>
            {actioning==="rejetee" ? <><Spinner size={14} color={C.red}/>Rejet…</> : "✗ Rejeter"}
          </button>
          {ep.stockPath ? (
            <button onClick={() => window.open(sb.fileUrl(ep.stockPath), "_blank")}
              style={{ padding:"9px 14px",background:"#f8fafc",border:`1.5px solid ${C.border}`,borderRadius:9,fontSize:13,fontWeight:600,cursor:"pointer",color:C.txtMuted,fontFamily:"inherit",whiteSpace:"nowrap" }}>
              👁️ Prévisualiser
            </button>
          ) : (
            <span title="Le fichier n'a pas pu être envoyé lors de la soumission" style={{ padding:"9px 12px",background:C.redPale,border:`1.5px solid ${C.redBorder}`,borderRadius:9,fontSize:12,fontWeight:700,color:C.red,whiteSpace:"nowrap" }}>
              ⚠️ Fichier manquant
            </span>
          )}
        </div>
      )}

      {/* Bouton prévisualisation seul (validée / rejetée) */}
      {isAdmin && ep.statut !== "attente" && (
        <div style={{ padding:"10px 16px",borderBottom:`1px solid ${C.border}` }}>
          {ep.stockPath ? (
            <button onClick={() => window.open(sb.fileUrl(ep.stockPath), "_blank")}
              style={{ padding:"7px 14px",background:"#f8fafc",border:`1px solid ${C.border}`,borderRadius:8,fontSize:12,fontWeight:600,cursor:"pointer",color:C.txtMuted,fontFamily:"inherit" }}>
              👁️ Prévisualiser / Imprimer
            </button>
          ) : (
            <span style={{ fontSize:12,fontWeight:700,color:C.red }}>⚠️ Fichier manquant — non envoyé lors de la soumission</span>
          )}
        </div>
      )}

      {/* Commentaires existants */}
      {ep.commentaires.length > 0 && (
        <div style={{ padding:"12px 16px",borderBottom:`1px solid ${C.border}`,display:"flex",flexDirection:"column",gap:8 }}>
          <div style={{ fontSize:11,fontWeight:700,color:C.txtMuted,textTransform:"uppercase",letterSpacing:".06em",marginBottom:2 }}>
            {isAdmin ? "Échanges" : "Retour de l'animatrice"}
          </div>
          {ep.commentaires.map((c, i) => (
            <div key={i} style={{ display:"flex",gap:9 }}>
              <div style={{ width:26,height:26,borderRadius:"50%",background:c.col||C.gold,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:800,color:"#fff",flexShrink:0 }}>
                {c.av}
              </div>
              <div style={{ flex:1,background:"#f8fafc",borderRadius:9,padding:"8px 11px",border:`1px solid ${C.border}` }}>
                <div style={{ display:"flex",justifyContent:"space-between",marginBottom:3 }}>
                  <span style={{ fontSize:11,fontWeight:700,color:C.txt }}>{c.aut}</span>
                  <span style={{ fontSize:10,color:C.txtLight }}>{c.date}</span>
                </div>
                <div style={{ fontSize:12,color:C.txtMuted,lineHeight:1.5 }}>{c.txt}</div>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Zone commentaire */}
      <div style={{ padding:"12px 16px",display:"flex",gap:8,alignItems:"flex-end" }}>
        <textarea value={comment} onChange={e=>setComment(e.target.value)}
          placeholder="Ajouter un commentaire…"
          rows={1} onInput={e=>{ e.target.style.height="auto"; e.target.style.height=e.target.scrollHeight+"px"; }}
          style={{ flex:1,padding:"9px 12px",border:`1px solid ${C.border}`,borderRadius:9,fontSize:12,color:C.txt,fontFamily:"inherit",resize:"none",outline:"none",lineHeight:1.4,background:"#f8fafc",transition:"border .2s" }}
          onFocus={e=>{ e.target.style.borderColor=C.green; e.target.style.background=C.white; }}
          onBlur={e=>{ e.target.style.borderColor=C.border; e.target.style.background="#f8fafc"; }}
          onKeyDown={e=>{ if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();handleComment();} }}/>
        <button onClick={handleComment} disabled={!comment.trim()||sendingCom}
          style={{ padding:"9px 14px",background:comment.trim()?C.green:"#e2e8f0",color:comment.trim()?"#fff":C.txtLight,border:"none",borderRadius:9,fontSize:13,fontWeight:700,cursor:comment.trim()?"pointer":"not-allowed",fontFamily:"inherit",flexShrink:0,display:"flex",alignItems:"center",gap:6 }}>
          {sendingCom ? <Spinner size={14}/> : "➤"}
        </button>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════
// FORMULAIRE DE SOUMISSION (enseignant)
// ═══════════════════════════════════════════════════
function FormSoumission({ user, onSubmit }) {
  const [file, setFile] = useState(null);
  const [classe, setClasse] = useState(user.classes?.[0] || "");
  const [trim, setTrim] = useState("T1");
  const [num, setNum] = useState("E1");
  const [note, setNote] = useState("");
  const [dragging, setDragging] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const fileRef = useRef();

  const handleFile = (f) => { if (f && f.size <= 10*1024*1024) setFile(f); };
  const handleDrop = (e) => { e.preventDefault(); setDragging(false); handleFile(e.dataTransfer.files[0]); };

  const submit = async () => {
    if (!file) return;
    setSubmitting(true);
    const id = "ep" + Date.now();
    const ext = file.name.split(".").pop().toLowerCase();
    const date = new Date().toISOString();
    const stockPath = `${id}_${file.name.replace(/[^a-zA-Z0-9._-]/g,"_")}`;
    const commentaires = note ? [{ av:user.ini, col:user.col, aut:getNomCourt(user.nom), txt:note, date:new Date().toLocaleString("fr-FR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"}) }] : [];
    // Upload Storage
    const uploaded = await sb.uploadFile(stockPath, file);
    const ep = { id, ens:user.id, classe, trim, num, fichier:file.name, stockpath:uploaded?stockPath:"", type:ext, taille:`${(file.size/1024).toFixed(0)} Ko`, soumis:date, statut:"attente", commentaires };
    const ok = await sb.rpc("submit_epreuve", {
      p_id: ep.id, p_ens: ep.ens, p_classe: ep.classe, p_trim: ep.trim, p_num: ep.num,
      p_fichier: ep.fichier, p_stockpath: ep.stockpath, p_type: ep.type, p_taille: ep.taille,
      p_soumis: ep.soumis, p_commentaires: ep.commentaires
    });
    setSubmitting(false);
    if (ok) { setFile(null); setNote(""); onSubmit(ep, true, !uploaded); }
    else onSubmit(null, false);
  };

  // Slots disponibles selon trim
  const slots = EP_SLOTS.filter(s => s.trim === trim);

  return (
    <div style={{ background:C.white,borderRadius:12,border:`1px solid ${C.border}`,overflow:"hidden" }}>
      <div style={{ padding:"14px 16px",background:"#f8fafc",borderBottom:`1px solid ${C.border}` }}>
        <h3 style={{ margin:0,fontSize:14,fontWeight:700,color:C.txt }}>📤 Soumettre une épreuve</h3>
        <p style={{ margin:"3px 0 0",fontSize:12,color:C.txtMuted }}>L'animatrice sera notifiée automatiquement</p>
      </div>
      <div style={{ padding:"18px 16px",display:"flex",flexDirection:"column",gap:14 }}>
        {/* Sélecteurs */}
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10 }}>
          <div>
            <label style={{ display:"block",fontSize:10,fontWeight:700,letterSpacing:".07em",textTransform:"uppercase",color:C.txtMuted,marginBottom:4 }}>Classe</label>
            <select value={classe} onChange={e=>setClasse(e.target.value)}
              style={{ width:"100%",padding:"9px 11px",border:`1.5px solid ${C.border}`,borderRadius:9,fontSize:13,color:C.txt,background:C.white,fontFamily:"inherit",outline:"none",cursor:"pointer" }}>
              {(user.classes||[]).map(cl=><option key={cl} value={cl}>{cl}</option>)}
            </select>
          </div>
          <div>
            <label style={{ display:"block",fontSize:10,fontWeight:700,letterSpacing:".07em",textTransform:"uppercase",color:C.txtMuted,marginBottom:4 }}>Trimestre</label>
            <select value={trim} onChange={e=>{ setTrim(e.target.value); setNum(EP_SLOTS.find(s=>s.trim===e.target.value)?.ep||"E1"); }}
              style={{ width:"100%",padding:"9px 11px",border:`1.5px solid ${C.border}`,borderRadius:9,fontSize:13,color:C.txt,background:C.white,fontFamily:"inherit",outline:"none",cursor:"pointer" }}>
              {["T1","T2","T3"].map(t=><option key={t} value={t}>{TRIM_LABELS[t]}</option>)}
            </select>
          </div>
          <div>
            <label style={{ display:"block",fontSize:10,fontWeight:700,letterSpacing:".07em",textTransform:"uppercase",color:C.txtMuted,marginBottom:4 }}>Épreuve</label>
            <select value={num} onChange={e=>setNum(e.target.value)}
              style={{ width:"100%",padding:"9px 11px",border:`1.5px solid ${C.border}`,borderRadius:9,fontSize:13,color:C.txt,background:C.white,fontFamily:"inherit",outline:"none",cursor:"pointer" }}>
              {slots.map(s=><option key={s.ep} value={s.ep}>{s.label}</option>)}
            </select>
          </div>
        </div>

        {/* Zone upload */}
        {!file ? (
          <div
            onDragOver={e=>{ e.preventDefault(); setDragging(true); }}
            onDragLeave={()=>setDragging(false)}
            onDrop={handleDrop}
            onClick={()=>fileRef.current?.click()}
            style={{ border:`2px dashed ${dragging?C.green:C.border}`,borderRadius:11,padding:"28px 20px",textAlign:"center",cursor:"pointer",background:dragging?C.greenPale:"#f8fafc",transition:"all .2s" }}>
            <div style={{ fontSize:32,marginBottom:8 }}>📎</div>
            <div style={{ fontSize:13,fontWeight:700,color:C.txt }}>Déposer le fichier ici</div>
            <div style={{ fontSize:12,color:C.txtMuted,marginTop:3 }}>ou cliquer pour parcourir — PDF, Word, image — max 10 Mo</div>
            <input ref={fileRef} type="file" accept=".pdf,.doc,.docx,.png,.jpg,.jpeg" style={{ display:"none" }} onChange={e=>handleFile(e.target.files[0])}/>
          </div>
        ) : (
          <div style={{ background:C.greenPale,border:`1.5px solid ${C.greenBorder}`,borderRadius:11,padding:"12px 14px",display:"flex",alignItems:"center",gap:12 }}>
            <div style={{ fontSize:22,flexShrink:0 }}>{file.name.endsWith(".pdf")?"📕":"📘"}</div>
            <div style={{ flex:1,minWidth:0 }}>
              <div style={{ fontSize:13,fontWeight:700,color:C.txt,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis" }}>{file.name}</div>
              <div style={{ fontSize:11,color:C.txtMuted }}>{(file.size/1024).toFixed(0)} Ko</div>
            </div>
            <button onClick={()=>setFile(null)} style={{ padding:"4px 9px",background:C.redPale,border:`1px solid ${C.redBorder}`,borderRadius:7,color:C.red,fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flexShrink:0 }}>Changer</button>
          </div>
        )}

        {/* Note facultative */}
        <div>
          <label style={{ display:"block",fontSize:10,fontWeight:700,letterSpacing:".07em",textTransform:"uppercase",color:C.txtMuted,marginBottom:4 }}>Note à l'animatrice (optionnel)</label>
          <textarea value={note} onChange={e=>setNote(e.target.value)} rows={2} placeholder="Informations complémentaires, consignes, contexte…"
            style={{ width:"100%",padding:"9px 12px",border:`1.5px solid ${C.border}`,borderRadius:9,fontSize:12,color:C.txt,fontFamily:"inherit",resize:"vertical",outline:"none",background:"#f8fafc",transition:"border .2s" }}
            onFocus={e=>{ e.target.style.borderColor=C.green; e.target.style.background=C.white; }}
            onBlur={e=>{ e.target.style.borderColor=C.border; e.target.style.background="#f8fafc"; }}/>
        </div>

        {/* Bouton */}
        <button onClick={submit} disabled={!file||submitting}
          style={{ width:"100%",padding:"12px",background:file&&!submitting?`linear-gradient(135deg,${C.greenDark},${C.greenLight})`:"#e2e8f0",color:file&&!submitting?"#fff":C.txtLight,border:"none",borderRadius:10,fontSize:14,fontWeight:700,cursor:file&&!submitting?"pointer":"not-allowed",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:9 }}>
          {submitting ? <><Spinner/>Envoi en cours…</> : "📤 Soumettre à l'animatrice"}
        </button>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════
// VUE ANIMATEUR COMPLÈTE
// ═══════════════════════════════════════════════════
function EpViewAdmin({ data, setData, showToast }) {
  const {isMobile} = useDevice();
  const [activeTab, setActiveTab] = useState("pending");
  const [selectedEp, setSelectedEp] = useState(null);
  const [filterEns, setFilterEns] = useState("all");

  const pending   = (data?.epreuves||[]).filter(e => e.statut === "attente");
  const validated = (data?.epreuves||[]).filter(e => e.statut === "validee");
  const rejected  = (data?.epreuves||[]).filter(e => e.statut === "rejetee");
  const enseignants = Object.values(data?.users||{}).filter(u => u.role !== "animatrice");

  const filtered = (list) => filterEns === "all" ? list : list.filter(e => e.ens === filterEns);

  const handleAction = useCallback(async (epId, statut, commentaire) => {
    const eps = [...data?.epreuves||[]];
    const ep  = eps.find(e => e.id === epId);
    if (!ep) return;

    if (statut) {
      ep.statut = statut;
      const ok = await sb.rpc("update_epreuve_statut", { p_id: ep.id, p_statut: statut });
      if (ok) {
        showToast(statut==="validee" ? "✅ Épreuve validée — enseignant notifié" : "⚠️ Épreuve rejetée — enseignant notifié", true);
      }
    }
    if (commentaire) {
      const now = new Date().toLocaleString("fr-FR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"});
      const nouveauCommentaire = { av:"AS", col:C.gold, aut:"Sylvie", txt:commentaire, date:now };
      ep.commentaires.push(nouveauCommentaire);
      await sb.rpc("add_epreuve_commentaire", { p_id: ep.id, p_commentaire: nouveauCommentaire });
      showToast("💬 Commentaire envoyé", true);
    }
    setData(prev => ({ ...prev, epreuves: eps }));
    if (selectedEp?.id === epId) setSelectedEp(eps.find(e => e.id === epId));
  }, [data?.epreuves||[], selectedEp]);

  const tabs = [
    { id:"pending",   label:"En attente",  count:pending.length,   color:C.orange },
    { id:"validated", label:"Validées",     count:validated.length, color:C.green  },
    { id:"rejected",  label:"Rejetées",     count:rejected.length,  color:C.red    },
    { id:"grille",    label:"Grille globale",count:null,            color:C.blue   },
  ];

  const currentList = activeTab==="pending" ? filtered(pending) : activeTab==="validated" ? filtered(validated) : filtered(rejected);

  return (
    <div style={{ display:"grid",gridTemplateColumns:"1fr 380px",gap:16,alignItems:"start" }}>
      {/* Gauche — liste */}
      <div style={{ display:"flex",flexDirection:"column",gap:14 }}>
        {/* Stats */}
        <div style={{ display:"grid",gridTemplateColumns: isMobile?"repeat(2,1fr)":"repeat(4,1fr)",gap:10 }}>
          {[
            { label:"En attente", val:pending.length,   col:C.orange, bg:C.orangePale },
            { label:"Validées",   val:validated.length, col:C.green,  bg:C.greenPale  },
            { label:"Rejetées",   val:rejected.length,  col:C.red,    bg:C.redPale    },
            { label:"Total",      val:(data?.epreuves||[]).length, col:C.blue, bg:C.bluePale },
          ].map((s,i)=>(
            <div key={i} style={{ background:s.bg,borderRadius:10,border:`1px solid ${s.col}20`,padding:"12px 14px" }}>
              <div style={{ fontSize:10,color:s.col,fontWeight:700,marginBottom:4 }}>{s.label}</div>
              <div style={{ fontSize:28,fontWeight:900,color:s.col }}>{s.val}</div>
            </div>
          ))}
        </div>

        {/* Onglets */}
        <div style={{ display:"flex",gap:0,background:"#f8fafc",borderRadius:10,padding:4,border:`1px solid ${C.border}` }}>
          {tabs.map(t=>(
            <button key={t.id} onClick={()=>setActiveTab(t.id)}
              style={{ flex:1,padding:"7px 6px",borderRadius:7,border:"none",background:activeTab===t.id?C.white:"transparent",color:activeTab===t.id?C.txt:C.txtMuted,fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:5,boxShadow:activeTab===t.id?"0 1px 4px rgba(0,0,0,.08)":"none",transition:"all .15s" }}>
              {t.label}
              {t.count !== null && <span style={{ minWidth:18,height:18,borderRadius:"50%",background:t.count>0?t.color:C.border,color:"#fff",fontSize:9,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center",padding:"0 4px" }}>{t.count}</span>}
            </button>
          ))}
        </div>

        {/* Filtre enseignant */}
        {activeTab !== "grille" && (
          <div style={{ display:"flex",alignItems:"center",gap:8 }}>
            <span style={{ fontSize:12,color:C.txtMuted,fontWeight:600,flexShrink:0 }}>Enseignant :</span>
            <select value={filterEns} onChange={e=>setFilterEns(e.target.value)}
              style={{ padding:"6px 10px",border:`1px solid ${C.border}`,borderRadius:8,fontSize:12,color:C.txt,background:C.white,fontFamily:"inherit",outline:"none",cursor:"pointer" }}>
              <option value="all">Tous les enseignants</option>
              {enseignants.map(u=><option key={u.id} value={u.id}>{u.nom}</option>)}
            </select>
          </div>
        )}

        {/* Contenu */}
        {activeTab === "grille" ? (
          <GrilleCouverture classes={CLASSES_REELLES} epreuves={data?.epreuves||[]} onCardClick={setSelectedEp}/>
        ) : (
          <div style={{ display:"flex",flexDirection:"column",gap:10 }}>
            {currentList.length === 0 ? (
              <div style={{ textAlign:"center",padding:"40px",color:C.txtMuted,background:C.white,borderRadius:12,border:`1px solid ${C.border}` }}>
                <div style={{ fontSize:32,marginBottom:8 }}>
                  {activeTab==="pending"?"📭":activeTab==="validated"?"✅":"📋"}
                </div>
                <div style={{ fontSize:13,fontWeight:600 }}>
                  {activeTab==="pending"?"Aucune épreuve en attente":activeTab==="validated"?"Aucune épreuve validée":"Aucune épreuve rejetée"}
                </div>
              </div>
            ) : (
              currentList.map(ep=>(
                <EpreuveCard key={ep.id} ep={ep} users={data?.users||{}} onAction={handleAction} isAdmin={true}/>
              ))
            )}
          </div>
        )}
      </div>

      {/* Droite — détail sélectionné */}
      <div style={{ position:"sticky",top:70 }}>
        {selectedEp ? (
          <div>
            <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10 }}>
              <h3 style={{ margin:0,fontSize:13,fontWeight:700,color:C.txt }}>Détail de l'épreuve</h3>
              <button onClick={()=>setSelectedEp(null)} style={{ padding:"4px 9px",background:"#f8fafc",border:`1px solid ${C.border}`,borderRadius:7,color:C.txtMuted,fontSize:11,cursor:"pointer",fontFamily:"inherit" }}>✕ Fermer</button>
            </div>
            <EpreuveCard ep={(data?.epreuves||[]).find(e=>e.id===selectedEp.id)||selectedEp} users={data?.users||{}} onAction={handleAction} isAdmin={true}/>
          </div>
        ) : (
          <div style={{ background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:"28px 20px",textAlign:"center",color:C.txtMuted }}>
            <div style={{ fontSize:32,marginBottom:10 }}>👆</div>
            <div style={{ fontSize:13,fontWeight:600,marginBottom:5 }}>Sélectionner une épreuve</div>
            <div style={{ fontSize:12 }}>Cliquez sur une carte de la grille ou d'une épreuve pour voir le détail et agir</div>
          </div>
        )}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════
// VUE ENSEIGNANT COMPLÈTE
// ═══════════════════════════════════════════════════
function EpViewTeacher({ user, data, setData, showToast }) {
  const {isMobile} = useDevice();
  const mesEpreuves = (data?.epreuves||[]).filter(e => e.ens_id === user.id || e.ens === user.id);
  const pending   = mesEpreuves.filter(e => e.statut === "attente");
  const validated = mesEpreuves.filter(e => e.statut === "validee");
  const rejected  = mesEpreuves.filter(e => e.statut === "rejetee");

  const handleAction = useCallback(async (epId, statut, commentaire) => {
    const eps = [...data?.epreuves||[]];
    const ep = eps.find(e => e.id === epId);
    if (!ep || !commentaire) return;
    const now = new Date().toLocaleString("fr-FR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"});
    const nouveauCommentaire = { av:user.ini, col:user.col, aut:getNomCourt(user.nom), txt:commentaire, date:now };
    ep.commentaires.push(nouveauCommentaire);
    await sb.rpc("add_epreuve_commentaire", { p_id: ep.id, p_commentaire: nouveauCommentaire });
    showToast("💬 Message envoyé à l'animatrice", true);
    setData(prev => ({ ...prev, epreuves: eps }));
  }, [data?.epreuves||[], user]);

  const onSubmit = (ep, ok, uploadFailed) => {
    if (ok && ep) {
      setData(prev => ({ ...prev, epreuves: [ep, ...prev.epreuves] }));
      if (uploadFailed) {
        showToast("⚠️ Fiche enregistrée mais le FICHIER n'a pas pu être envoyé — réessayez ou contactez l'animatrice", false);
      } else {
        showToast("✅ Épreuve soumise — animatrice notifiée", true);
      }
    } else {
      showToast("⚠️ Erreur lors de la soumission", false);
    }
  };

  return (
    <div style={{ display:"grid",gridTemplateColumns:"1fr 340px",gap:16,alignItems:"start" }}>
      {/* Gauche — mes épreuves */}
      <div style={{ display:"flex",flexDirection:"column",gap:14 }}>
        {/* Stats */}
        <div style={{ display:"grid",gridTemplateColumns: isMobile?"repeat(2,1fr)":"repeat(3,1fr)",gap:10 }}>
          {[
            { label:"En attente",   val:pending.length,   col:C.orange, bg:C.orangePale },
            { label:"Validées",     val:validated.length, col:C.green,  bg:C.greenPale  },
            { label:"À corriger",   val:rejected.length,  col:C.red,    bg:C.redPale    },
          ].map((s,i)=>(
            <div key={i} style={{ background:s.bg,borderRadius:10,border:`1px solid ${s.col}20`,padding:"12px 14px" }}>
              <div style={{ fontSize:10,color:s.col,fontWeight:700,marginBottom:4 }}>{s.label}</div>
              <div style={{ fontSize:28,fontWeight:900,color:s.col }}>{s.val}</div>
            </div>
          ))}
        </div>

        <h3 style={{ margin:0,fontSize:14,fontWeight:700,color:C.txt }}>📋 Mes épreuves soumises</h3>

        {mesEpreuves.length === 0 ? (
          <div style={{ textAlign:"center",padding:"40px",color:C.txtMuted,background:C.white,borderRadius:12,border:`1px solid ${C.border}` }}>
            <div style={{ fontSize:32,marginBottom:8 }}>📭</div>
            <div style={{ fontSize:13,fontWeight:600 }}>Aucune épreuve soumise</div>
            <div style={{ fontSize:12,marginTop:4 }}>Utilisez le formulaire ci-contre pour soumettre votre première épreuve</div>
          </div>
        ) : (
          <div style={{ display:"flex",flexDirection:"column",gap:10 }}>
            {/* Rejetées en premier — action requise */}
            {rejected.length > 0 && (
              <div style={{ background:C.redPale,border:`1px solid ${C.redBorder}`,borderRadius:10,padding:"10px 14px",marginBottom:4 }}>
                <div style={{ fontSize:12,fontWeight:700,color:C.red }}>⚠️ {rejected.length} épreuve{rejected.length>1?"s":""} rejetée{rejected.length>1?"s":""} — action requise</div>
                <div style={{ fontSize:11,color:C.txtMuted,marginTop:2 }}>Consultez les commentaires et soumettez une version corrigée</div>
              </div>
            )}
            {mesEpreuves.map(ep => (
              <EpreuveCard key={ep.id} ep={ep} users={data?.users||{}} onAction={handleAction} isAdmin={false}/>
            ))}
          </div>
        )}
      </div>

      {/* Droite — formulaire */}
      <div style={{ position:"sticky",top:70 }}>
        <FormSoumission user={user} onSubmit={onSubmit}/>
        {/* Grille slots */}
        <div style={{ marginTop:14,background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:"14px 16px" }}>
          <h4 style={{ margin:"0 0 12px",fontSize:13,fontWeight:700,color:C.txt }}>📊 Mes slots d'épreuves</h4>
          <div style={{ display:"grid",gridTemplateColumns: isMobile?"repeat(2,1fr)":"repeat(2,1fr)",gap:7 }}>
            {EP_SLOTS.map((slot,i) => {
              const submitted = mesEpreuves.filter(e => e.trim===slot.trim && e.num===slot.ep);
              // On prend la plus récente
              const ep = submitted.sort((a,b)=>b.soumis?.localeCompare(a.soumis)||0)[0];
              const st = ep?.statut || "vide";
              const cfg = STATUT_CONFIG[st];
              return (
                <div key={i} style={{ padding:"8px 10px",borderRadius:9,border:`1.5px solid ${cfg.border}`,background:cfg.bg }}>
                  <div style={{ fontSize:9,fontWeight:800,color:TRIM_COLORS[slot.trim],marginBottom:3,letterSpacing:".05em" }}>{slot.label}</div>
                  <div style={{ display:"flex",alignItems:"center",gap:5 }}>
                    <span style={{ fontSize:14 }}>{cfg.emoji}</span>
                    <span style={{ fontSize:10,color:cfg.color,fontWeight:700 }}>{cfg.label}</span>
                  </div>
                  {ep?.classe && <div style={{ fontSize:9,color:C.txtLight,marginTop:2 }}>{ep.classe}</div>}
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════
// PAGE PRINCIPALE
// ═══════════════════════════════════════════════════

function EpreuvesPage() {
  const {user, data, setData, showToast} = useApp();
  const isAdmin = isAdminRole(user?.role);

  if (!data) return (
    <div style={{padding:"60px",textAlign:"center",color:C.txtMuted}}>
      <Spinner size={28} color={C.green}/>
      <div style={{marginTop:12,fontSize:13}}>Chargement…</div>
    </div>
  );

  if (isAdmin) return <EpViewAdmin data={data} setData={setData} showToast={showToast}/>;

  // Vue enseignant — utiliser l'utilisateur connecté
  const ens = user;
  return <EpViewTeacher user={ens} data={data} setData={setData} showToast={showToast}/>;
}


// ═══════════════════════════════════════════════════
// PALETTE
// ═══════════════════════════════════════════════════
// const C: référence au bloc global ci-dessus

// ═══════════════════════════════════════════════════
// CHARGEMENT DONNÉES
// ═══════════════════════════════════════════════════
// loadData: définie globalement (charge classes, users, prog_suivi, epreuves, edt_exceptions)

// progDocMeta: définie globalement dans le header

// ═══════════════════════════════════════════════════
// GÉNÉRATION HTML DOCUMENT OFFICIEL MINESEC
// ═══════════════════════════════════════════════════
function enteteOfficiel(titre, sousTitre, dept) {
  return `
  <div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:12px;border-bottom:2px solid #1a6b3c;padding-bottom:10px">
    <div style="text-align:center;font-size:9px;line-height:1.5;width:45%">
      <div style="font-weight:bold">REPUBLIQUE DU CAMEROUN</div>
      <div style="font-style:italic">Paix – Travail – Patrie</div>
      <div style="font-size:7px">– – – – – –</div>
      <div>MINISTERE DES ENSEIGNEMENTS SECONDAIRES</div>
      <div style="font-size:7px">– – – – – –</div>
      <div>INSPECTION GENERALE DES ENSEIGNEMENTS</div>
      <div style="font-size:7px">– – – – – –</div>
      <div>INSPECTION DE PEDAGOGIE CHARGEE DE L'ENSEIGNEMENT DES SCIENCES</div>
      <div style="font-size:7px">– – – – – –</div>
      <div style="font-weight:bold">SECTION : ${dept || "SVTEEHB"}</div>
    </div>
    <div style="text-align:center;width:10%;display:flex;flex-direction:column;align-items:center;gap:6px">
      <img src="${LOGO_LYCEE_B64}" alt="" style="width:64px;height:64px;object-fit:contain;"/>
    </div>
    <div style="text-align:center;font-size:9px;line-height:1.5;width:45%">
      <div style="font-weight:bold">REPUBLIC OF CAMEROON</div>
      <div style="font-style:italic">Peace – Work – Fatherland</div>
      <div style="font-size:7px">– – – – – –</div>
      <div>MINISTRY OF SECONDARY EDUCATION</div>
      <div style="font-size:7px">– – – – – –</div>
      <div>INSPECTORATE GENERAL OF EDUCATION</div>
      <div style="font-size:7px">– – – – – –</div>
      <div>INSPECTORATE OF PEDAGOGY IN CHARGE OF THE TEACHING OF SCIENCES</div>
      <div style="font-size:7px">– – – – – –</div>
      <div style="font-weight:bold">SECTION : LESEEHB</div>
    </div>
  </div>
  <h2 style="text-align:center;font-size:13px;font-weight:900;text-transform:uppercase;border:2px solid #222;padding:7px;margin:8px 0;letter-spacing:.04em">${titre}</h2>
  ${sousTitre ? `<p style="text-align:center;font-size:10px;color:#444;margin:0 0 8px">${sousTitre}</p>` : ""}
  <div style="text-align:center;font-size:9px;margin-bottom:2px">Établissement : Lycée de Kakatare &nbsp;|&nbsp; Discipline : SVTEEHB &nbsp;|&nbsp; Année scolaire : 2025-2026</div>`;
}

// Fiche de suivi pédagogique (vue animatrice par enseignant)
// ═══════════════════════════════════════════════════════════════════
// genFicheSuivi — Fiche de Suivi Pédagogique MINESEC officielle
// Source modèle : Modele_de_FP.pdf — MINESEC-DRES/EXTREME-NORD
// Structure : identique au modèle officiel — toutes colonnes présentes
// ═══════════════════════════════════════════════════════════════════
function genFicheSuivi(enseignant, classes, progIndex, trim = "ANN", notesIndex = {}, absencesIndex = {}, deptNom = "SVTEEHB", animateurNom = "—") {
  const periode = TRIM_LABELS[trim] || "Année complète";
  const nbSem   = trim === "ANN" ? 36 : 12;
  const dateJour = new Date().toLocaleDateString("fr-FR",{day:"2-digit",month:"long",year:"numeric"});
  const evalLabel = trim === "ANN" ? "N° 1 · 2 · 3 · 4 · 5 · 6" : trim === "T1" ? "N° 1 et 2" : trim === "T2" ? "N° 3 et 4" : "N° 5 et 6";

  // ── Calcul des lignes ────────────────────────────────────────────
  const statsObs = [];
  const rows = classes.map((cl,idx) => {
    const eleves = ELEVES_DB[cl] || [];
    const g = eleves.filter(e=>e.g==="M").length;
    const f = eleves.filter(e=>e.g==="F").length;
    const total = eleves.length;
    const code = resolveProgCode(cl);
    const meta = code ? PROG_META[code] : null;
    if (!meta) return `<tr><td style="font-weight:700;padding:3px 4px;text-align:center;font-size:7.5pt;background:#f0f0f0;border:1px solid #999">${cl}</td><td colspan="22" style="border:1px solid #999;padding:3px;text-align:center;color:#999;font-style:italic;font-size:7pt">Programme non mappé</td></tr>`;

    const key    = `${enseignant.id}||${cl}`;
    const faites = progIndex[key] || [];
    const leconsCodeAll = LECONS_DATA[code]||[];
    const tdLeconsAll = leconsCodeAll.filter(l=>/intégration/i.test(l.t));

    let lp = meta.lpRef, tpP = meta.tp?.length||0, tdP = tdLeconsAll.length;
    if (trim !== "ANN") {
      const range = code ? getTrimRange(code, trim) : null;
      if (range) {
        const lecons = leconsCodeAll;
        lp  = lecons.filter(l=>l.n>=range[0]&&l.n<=range[1]).length || meta.lpRef;
        tpP = (meta.tp||[]).filter(n=>n>=range[0]&&n<=range[1]).length;
        tdP = tdLeconsAll.filter(l=>l.n>=range[0]&&l.n<=range[1]).length;
      }
    }
    const lfTrim = trim==="ANN" ? faites.length : faites.filter(n=>{ const r=getTrimRange(code,trim); return r&&n>=r[0]&&n<=r[1]; }).length;
    const tauxLP = lp>0 ? Math.min(100, Math.round(lfTrim/lp*100)) : 0;
    const tpFT   = trim==="ANN"
      ? (meta.tp||[]).filter(n=>faites.includes(n)).length
      : (meta.tp||[]).filter(n=>{ const r=getTrimRange(code,trim); return faites.includes(n)&&r&&n>=r[0]&&n<=r[1]; }).length;
    const tdFait = trim==="ANN"
      ? tdLeconsAll.filter(l=>faites.includes(l.n)).length
      : tdLeconsAll.filter(l=>{ const r=getTrimRange(code,trim); return faites.includes(l.n)&&r&&l.n>=r[0]&&l.n<=r[1]; }).length;
    // Nombre prévu/fait total = TP + TD (activités d'intégration, obligatoires)
    const tpTdP  = tpP + tdP;
    const tpTdFT = tpFT + tdFait;
    const tauxTP = tpTdP>0 ? Math.min(100, Math.round(tpTdFT/tpTdP*100)) : 0;
    const hDues  = meta.hd;
    const hFaites = Math.round(lfTrim * (meta.hd/meta.lpRef));
    const tauxHD = hDues>0 ? Math.min(100, Math.round(hFaites/hDues*100)) : 0;

    // Utilisation des ressources digitalisées — depuis prog_suivi clé "||dig"
    const digFaites = progIndex[`${key}||dig`] || [];
    const leconsCode = LECONS_DATA[code]||[];
    const digRange = trim!=="ANN" ? getTrimRange(code,trim) : null;
    const leconsDigitalisables = digRange
      ? leconsCode.filter(l=>l.d===1 && l.n>=digRange[0] && l.n<=digRange[1])
      : leconsCode.filter(l=>l.d===1);
    const ldTot  = leconsDigitalisables.length;
    const ldFait = leconsDigitalisables.filter(l=>digFaites.includes(l.n)).length;
    const tauxLD = ldTot>0 ? Math.min(100, Math.round(ldFait/ldTot*100)) : 0;

    // Taux de réussite depuis notes
    const getN = (eid, t, e) => { const v = (notesIndex[`${cl}||${t}-${e}`]||{})[eid]; return (v===undefined||v===""||v===null) ? null : +v; };
    const moyTrimEl = (eid, t) => { const e1=getN(eid,t,"E1"), e2=getN(eid,t,"E2"); if(e1===null&&e2===null) return null; if(e1!==null&&e2!==null) return (e1+e2)/2; return e1!==null?e1:e2; };
    const trimsToAvg = trim==="ANN" ? ["T1","T2","T3"] : [trim];
    const moyEleve = (eid) => { const ms=trimsToAvg.map(t=>moyTrimEl(eid,t)).filter(m=>m!==null); return ms.length>0?ms.reduce((s,m)=>s+m,0)/ms.length:null; };
    let mAll=[], mG=[], mF=[];
    eleves.forEach(el=>{ const m=moyEleve(el.id); if(m!==null){ mAll.push(m); if(el.g==="M") mG.push(m); else mF.push(m); } });
    const avg = a => a.length>0?a.reduce((s,n)=>s+n,0)/a.length:0;
    const tauxReuss = mAll.length>0?Math.round(mAll.filter(n=>n>=10).length/mAll.length*100):null;
    const moyClasse = mAll.length>0?avg(mAll).toFixed(2):null;
    const pctG = mG.length>0?Math.round(mG.filter(n=>n>=10).length/mG.length*100):null;
    const pctF = mF.length>0?Math.round(mF.filter(n=>n>=10).length/mF.length*100):null;

    // Taux d'Assiduité (TA) — formule exacte de la légende (2) :
    // A = Volume horaire HEBDOMADAIRE × Effectif, étendu sur nbSem semaines de la période
    // B = Total des heures d'absence enregistrées sur la période
    // Présences enregistrées par DATE (indépendant du cahier de texte)
    const moisDansTrim = (dateStr, tr) => {
      if (tr === "ANN") return true;
      const mois = parseInt((dateStr||"").split("-")[1], 10);
      if (tr === "T1") return mois>=9 && mois<=12;
      if (tr === "T2") return mois>=1 && mois<=3;
      if (tr === "T3") return mois>=4 && mois<=7;
      return true;
    };
    const absPrefix = `${enseignant.id}||${cl}||`;
    const datesAvecDonnees = Object.keys(absencesIndex)
      .filter(k=>k.startsWith(absPrefix))
      .map(k=>k.slice(absPrefix.length))
      .filter(d=>moisDansTrim(d, trim));
    const heuresParSeance = lp>0 ? (meta.vh*nbSem)/lp : 0; // 1 séance ≈ (vh×semaines période)/leçons prévues
    let heuresAbsTotal = 0;
    datesAvecDonnees.forEach(d=>{
      const abs = absencesIndex[`${absPrefix}${d}`]||[];
      heuresAbsTotal += abs.length * heuresParSeance;
    });
    const seancesAvecDonnees = datesAvecDonnees.length;
    // A = capacité d'heures pour les SÉANCES RÉELLEMENT RENSEIGNÉES uniquement
    // (et non toute l'année — sinon 1 seule journée notée dilue le taux vers 100%)
    const heuresPossibles = total * heuresParSeance * seancesAvecDonnees;
    const tauxAssiduite = (seancesAvecDonnees>0 && heuresPossibles>0)
      ? Math.max(0, Math.min(100, Math.round((heuresPossibles - heuresAbsTotal) * 100 / heuresPossibles)))
      : null;
    const heuresFaitesAssidu = seancesAvecDonnees>0 ? Math.round(heuresPossibles - heuresAbsTotal) : null;

    // Couleurs taux
    const col = v => v>=75?"#16a34a":v>=50?"#92400e":"#991b1b";
    const pct = (v,sfx="%") => v!==null?`<span style="color:${col(v)};font-weight:700">${v}${sfx}</span>`:"—";
    const bgRow = idx%2===0?"#ffffff":"#f9f9f9";
    const td = `style="border:1px solid #aaa;padding:2px 3px;text-align:center;font-size:7.5pt;background:${bgRow}"`;
    const tdG = `style="border:1px solid #aaa;padding:2px 3px;text-align:center;font-size:8pt;font-weight:700;background:#f0f0f0"`;

    statsObs.push({cl, total, tauxLP, tauxTP, tauxAssiduite, pctG, pctF, tauxReuss, moyClasse});
    return `<tr>
      <td ${tdG}>${cl}</td>
      <td ${td}>${g}</td><td ${td}>${f}</td><td ${td} style="border:1px solid #aaa;padding:2px 3px;text-align:center;font-size:7.5pt;font-weight:700;background:${bgRow}">${total}</td>
      <td ${td}>${seancesAvecDonnees>0?Math.round(heuresPossibles):"—"}</td><td ${td}>${heuresFaitesAssidu!==null?heuresFaitesAssidu:"—"}</td><td ${td}>${tauxAssiduite!==null?pct(tauxAssiduite):'<span style="color:#999;font-style:italic;font-size:6.5pt">N/D</span>'}</td>
      <td ${td}>${lp}</td><td ${td}><strong>${lfTrim}</strong></td><td ${td}>${pct(tauxLP)}</td>
      <td ${td}>${ldTot}</td><td ${td}><strong>${ldFait}</strong></td><td ${td}>${pct(tauxLD)}</td>
      <td ${td}>${hDues}</td><td ${td}>${hFaites}</td><td ${td}>${pct(tauxHD)}</td>
      <td ${td}>${tpTdP}</td><td ${td}>${tpTdFT}</td><td ${td}>${pct(tauxTP)}</td>
      <td ${td}>${pctG!==null?pct(pctG):"—"}</td><td ${td}>${pctF!==null?pct(pctF):"—"}</td><td ${td}>${tauxReuss!==null?pct(tauxReuss):"—"}</td>
      <td ${td} style="border:1px solid #aaa;padding:2px 3px;text-align:center;font-size:7.5pt;font-weight:700;background:${bgRow}">${moyClasse!==null?moyClasse:"—"}</td>
    </tr>`;
  }).join("");

  // ── Constats automatiques (brouillon — à compléter par l'enseignant) ──
  function genObservations(stats) {
    const difficultes = [];
    const suggestions = [];
    stats.forEach(s => {
      if (s.tauxLP !== null && s.tauxLP < 50) {
        difficultes.push(`${s.cl} : couverture du programme à ${s.tauxLP}% — nettement sous l'objectif (75%)`);
        suggestions.push(`${s.cl} : envisager un rattrapage accéléré ou revoir le rythme de progression`);
      }
      if (s.tauxTP !== null && s.tauxTP < 40) {
        difficultes.push(`${s.cl} : faible exécution des TP/TD (${s.tauxTP}%)`);
      }
      if (s.tauxAssiduite !== null && s.tauxAssiduite < 85) {
        difficultes.push(`${s.cl} : taux d'assiduité de ${s.tauxAssiduite}% — absences à surveiller`);
      }
      if (s.pctG !== null && s.pctF !== null && Math.abs(s.pctG - s.pctF) >= 25) {
        difficultes.push(`${s.cl} : écart de réussite important entre garçons (${s.pctG}%) et filles (${s.pctF}%)`);
        suggestions.push(`${s.cl} : analyser les causes de l'écart de réussite entre genres`);
      }
      if (s.moyClasse !== null && s.moyClasse < 8) {
        difficultes.push(`${s.cl} : moyenne générale faible (${s.moyClasse}/20)`);
      }
    });
    return { difficultes: difficultes.slice(0,5), suggestions: suggestions.slice(0,4) };
  }
  const obs = genObservations(statsObs);
  const ligneObs = (txt) => `&nbsp;&nbsp;&nbsp;—&nbsp;&nbsp;${txt}<br>`;
  const blancObs = () => `&nbsp;&nbsp;&nbsp;—&nbsp;&nbsp;&nbsp;<span style="color:#bbb">_______________________________________________</span><br>`;
  const difficultesHTML = obs.difficultes.length>0
    ? obs.difficultes.map(ligneObs).join("") + blancObs()
    : blancObs() + blancObs();
  const suggestionsHTML = obs.suggestions.length>0
    ? obs.suggestions.map(ligneObs).join("") + blancObs()
    : blancObs() + blancObs();

  // ── Infos heures hebdo ──────────────────────────────────────────
  const hHebdoTotal = classes.reduce((sum,cl) => { const code=resolveProgCode(cl); return sum + (code?(PROG_META[code]?.vh||0):0); }, 0);
  const hHebdoStr = `${hHebdoTotal}h / semaine`;

  return `<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8">
<style>
  @page { size: A4 landscape; margin: 9mm 10mm 9mm 10mm; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Times New Roman", Times, serif; font-size: 9pt; color: #000; background: #fff; }

  /* ── EN-TÊTE BILINGUE ─────────────────────────────── */
  .entete { display: flex; justify-content: space-between; align-items: flex-start; padding-bottom: 4px; margin-bottom: 0; }
  .col-fr, .col-en { width: 42%; font-size: 8pt; line-height: 1.55; text-align: center; }
  .pays { font-weight: 900; font-size: 9.5pt; text-transform: uppercase; letter-spacing: .03em; }
  .devise { font-style: italic; font-size: 8pt; }
  .logo-zone { width: 14%; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 3px; }
  .logo-zone img { width: 64px; height: 64px; object-fit: contain; }

  /* ── BLOC TITRE (tableau grisé) ───────────────────── */
  .titre-tbl { width: 100%; border-collapse: collapse; margin-top: 4px; }
  .titre-tbl td { background: #bbbbbb; border: 1.5px solid #999; text-align: center; padding: 5px 8px 4px; }
  .titre-principal { font-size: 13pt; font-weight: 900; text-transform: uppercase; letter-spacing: .05em; text-decoration: underline; color: #333333; }
  .titre-annee { font-size: 10pt; font-weight: 700; letter-spacing: .06em; margin-top: 2px; color: #333333; }
  .titre-minesec { font-size: 8pt; font-style: italic; font-weight: 700; letter-spacing: .04em; margin-top: 2px; color: #333333; }

  /* ── TABLEAU IDENTITÉ ─────────────────────────────── */
  .ident-tbl { width: 100%; border-collapse: collapse; margin-top: 6px; }
  .ident-tbl td { border: 1px solid #bbbbbb; background: #f4f4f4; padding: 5px 7px; font-size: 8.5pt; vertical-align: middle; }
  .il { font-style: italic; color: #666666; }
  .iv { font-weight: 700; color: #333333; border-bottom: 1px solid #777; display: inline-block; min-width: 40px; }

  /* ── TABLEAU PRINCIPAL ────────────────────────────── */
  .main-tbl { width: 100%; border-collapse: collapse; margin-top: 6px; font-family: Arial, Helvetica, sans-serif; font-size: 7pt; }
  .main-tbl th, .main-tbl td { border: 1px solid #aaa; padding: 2px 3px; text-align: center; vertical-align: middle; line-height: 1.25; }
  .th-l1 { background: #bbbbbb; font-size: 7pt; font-weight: 700; color: #333; }
  .th-l2 { background: #cecece; font-size: 6.5pt; font-weight: 700; color: #333; }
  .th-l3 { background: #e4e4e4; font-size: 6pt; font-weight: 700; color: #000; }

  /* ── LÉGENDE ──────────────────────────────────────── */
  .legende-tbl { width: 100%; border-collapse: collapse; margin-top: 4px; }
  .legende-tbl td { border: 1px solid #aaa; padding: 4px 8px; font-family: Arial, sans-serif; font-size: 6.5pt; line-height: 1.7; vertical-align: top; }

  /* ── OBSERVATIONS ─────────────────────────────────── */
  .obs-tbl { width: 100%; border-collapse: collapse; margin-top: 0; }
  .obs-tbl td { border: 1px solid #aaa; padding: 5px 9px; font-family: Arial, sans-serif; font-size: 7.5pt; }
  .obs-label { font-weight: 700; text-decoration: underline; color: #666; }

  /* ── SIGNATURES ───────────────────────────────────── */
  .sigs-tbl { width: 100%; border-collapse: collapse; margin-top: 0; }
  .sigs-tbl td { border: 1px solid #aaa; padding: 8px 12px 12px; font-family: Arial, sans-serif; font-size: 8pt; }
  .sig-titre { font-weight: 700; text-transform: uppercase; color: #666; text-align: center; }
  .sig-nom { font-weight: 700; text-align: center; margin-top: 2px; }
  .sig-lieu { font-style: italic; font-size: 7.5pt; text-align: right; margin-bottom: 2px; }
</style></head><body>

<!-- EN-TÊTE BILINGUE -->
<div class="entete">
  <div class="col-fr">
    <div class="pays">République du Cameroun</div>
    <div class="devise">Paix &ndash; Travail &ndash; Patrie</div>
  </div>
  <div class="logo-zone">
    <img src="${LOGO_LYCEE_B64}" alt=""/>
  </div>
  <div class="col-en">
    <div class="pays">Republic of Cameroon</div>
    <div class="devise">Peace &ndash; Work &ndash; Fatherland</div>
  </div>
</div>

<!-- BLOC TITRE -->
<table class="titre-tbl">
  <tr><td>
    <div class="titre-principal">Fiche de Suivi Pédagogique de l'Enseignant(e)</div>
    <div class="titre-annee">Année Scolaire : 2025 / 2026</div>
    <div class="titre-minesec">MINESEC / DRES / EXTRÊME-NORD</div>
  </td></tr>
</table>

<!-- TABLEAU IDENTITÉ -->
<table class="ident-tbl">
  <tr>
    <td style="width:33%"><span class="il">Établissement : </span><span class="iv">LYCÉE DE KAKATARE</span></td>
    <td style="width:33%">
      <span class="il">Discipline : </span><span class="iv">${deptNom}</span>
      <span class="il" style="margin-left:8px">Grade : </span><span class="iv">PLEG</span>
    </td>
    <td style="width:34%">
      <span class="il">Ancienneté : </span><span class="iv" style="min-width:30px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
      <span class="il" style="margin-left:6px">Qualité : </span><span class="iv">Fonctionnaire</span>
    </td>
  </tr>
  <tr>
    <td><span class="il">Nom de l'enseignant(e) : </span><span class="iv">${enseignant.nom}</span></td>
    <td><span class="il">Animateur pédagogique : </span><span class="iv">${animateurNom}</span></td>
    <td><span class="il">Nb h. hebdomadaires : </span><span class="iv">${hHebdoStr}</span></td>
  </tr>
  <tr>
    <td><span class="il">Période : </span><span class="iv">${periode}</span></td>
    <td><span class="il">Évaluation : </span><span class="iv">${evalLabel}</span></td>
    <td><span class="il">Tél. : </span><span class="iv" style="min-width:80px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
</table>

<!-- TABLEAU PRINCIPAL MINESEC (23 colonnes) -->
<table class="main-tbl">
  <colgroup>
    <col style="width:54px"><!-- classe -->
    <col style="width:17px"><col style="width:17px"><col style="width:18px"><!-- effectifs -->
    <col style="width:20px"><col style="width:20px"><col style="width:20px"><!-- assiduité -->
    <col style="width:18px"><col style="width:18px"><col style="width:18px"><!-- présentiel -->
    <col style="width:18px"><col style="width:18px"><col style="width:18px"><!-- digitalisé -->
    <col style="width:20px"><col style="width:20px"><col style="width:20px"><!-- heures -->
    <col style="width:19px"><col style="width:19px"><col style="width:19px"><!-- TP/TD -->
    <col style="width:18px"><col style="width:18px"><col style="width:20px"><!-- réussite -->
    <col style="width:28px"><!-- moy. gén. -->
  </colgroup>
  <thead>
    <tr>
      <th class="th-l1" rowspan="3">Classes<br>tenues</th>
      <th class="th-l1" colspan="3">Effectifs des élèves</th>
      <th class="th-l1" colspan="3">Assiduité des élèves <sup>(²)</sup></th>
      <th class="th-l1" colspan="6">Couverture des programmes/référentiels de formation (par rapport à l'année)</th>
      <th class="th-l1" colspan="3">Couverture des heures<br>d'enseignement / année</th>
      <th class="th-l1" colspan="3">Taux d'exécution <sup>(³)</sup> des TPI/TD/DOSSIERS / année</th>
      <th class="th-l1" colspan="3">Taux de réussite <sup>(⁴)</sup> des élèves (M ≥ 10)</th>
      <th class="th-l1" rowspan="3">Moyenne Générale<br>de la classe <sup>(⁵)</sup><br>/20</th>
    </tr>
    <tr>
      <th class="th-l2">Garç.</th><th class="th-l2">Filles</th><th class="th-l2">Total</th>
      <th class="th-l2">Heures<br>Prévues</th><th class="th-l2">Heures<br>Faites</th><th class="th-l2">Taux<br>(%)</th>
      <th class="th-l2" colspan="3">En présentiel</th>
      <th class="th-l2" colspan="3">Utilisation des ressources<br>digitalisées (Distance Éd. ou non)</th>
      <th class="th-l2">Heures<br>Dues</th><th class="th-l2">Heures<br>Faites</th><th class="th-l2">Taux<br>(%)</th>
      <th class="th-l2">Nbre<br>Prévu</th><th class="th-l2">Nbre<br>Fait</th><th class="th-l2">Taux<br>(%)</th>
      <th class="th-l2">%<br>Garç.</th><th class="th-l2">%<br>Filles</th><th class="th-l2">%<br>Total</th>
    </tr>
    <tr>
      <th class="th-l3"></th><th class="th-l3"></th><th class="th-l3"></th>
      <th class="th-l3"></th><th class="th-l3"></th><th class="th-l3"></th>
      <th class="th-l3">Leçons<br>Prévues</th><th class="th-l3">Leçons<br>Faites</th><th class="th-l3">Taux (%)</th>
      <th class="th-l3">Leçons<br>Prévues</th><th class="th-l3">Leçons<br>Faites</th><th class="th-l3">Taux (%)</th>
      <th class="th-l3"></th><th class="th-l3"></th><th class="th-l3"></th>
      <th class="th-l3"></th><th class="th-l3"></th><th class="th-l3"></th>
      <th class="th-l3"></th><th class="th-l3"></th><th class="th-l3"></th>
    </tr>
  </thead>
  <tbody>${rows}</tbody>
</table>

<!-- LÉGENDE -->
<table class="legende-tbl">
  <tr>
    <td style="width:50%">
      (1) Rayer la mention inutile.<br>
      <strong>(2) Calcul du Taux d'Assiduité (TA) :</strong><br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;TA = (A &minus; B) &times; 100 / A<br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;A = Volume horaire hebdomadaire &times; Effectif de la classe<br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;B = Total des heures d'absence enregistrées dans la semaine<br>
      <em>NB : Ce taux peut être calculé à l'échelle du trimestre (12 semaines).</em>
    </td>
    <td style="width:50%">
      <strong>(3) Taux d'exécution des TD/TP/Dossiers :</strong><br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(Nbre fait de TD+TP+Dossiers) &times; 100 / Nbre prévu de TD/TP/Dossiers<br><br>
      <strong>(4) Taux de réussite :</strong><br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(Nbre d'élèves ayant M ≥ 10 par genre) &times; 100 / Effectif par genre évalué<br><br>
      <strong>(5) Moyenne générale de la classe :</strong><br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Somme des notes de chaque élève / Nombre d'élèves évalués
    </td>
  </tr>
</table>

<!-- OBSERVATIONS -->
<table class="obs-tbl">
  <tr><td>
    <span class="obs-label">Observations générales (difficultés rencontrées + suggestions pour améliorer les résultats) :</span>
    <span style="font-size:6.5pt;color:#0369a1;font-style:italic">— constats automatiques générés à partir des données ci-dessus, à vérifier et compléter —</span><br><br>
    <strong>Difficultés rencontrées :</strong><br>
    ${difficultesHTML}<br>
    <strong>Suggestions pour améliorer les résultats :</strong><br>
    ${suggestionsHTML}
  </td></tr>
</table>

<!-- SIGNATURES -->
<table class="sigs-tbl">
  <tr>
    <td style="width:50%">
      <div class="sig-titre">Visa de l'Animateur Pédagogique</div>
      <div class="sig-nom">AÏSSATOU SYLVIE</div>
    </td>
    <td style="width:50%">
      <div class="sig-lieu">Fait à Maroua, le _________________________</div>
      <div class="sig-titre">Signature de l'Enseignant(e)</div>
      <div class="sig-nom">${enseignant.nom}</div>
    </td>
  </tr>
</table>

<script>window.onload=()=>{window.print();}</script>
</body></html>`;
}




// Fiche de compilation par niveau
// Regroupement par niveau — structure exacte de la maquette de référence
// Utilisé par genCompilation (impression) ET exportCompilationExcel (export Excel)
const NIVEAUX_COMPIL = [
  {label:"6ème",     salles:3, classes:["6ème 1","6ème 2","6ème 3"]},
  {label:"5ème",     salles:3, classes:["5ème 1","5ème 2","5ème 3"]},
  {label:"4ème",     salles:3, classes:["4ème ALL","4ème ARB","4ème CHN","4ème ESP","4ème ITA"]},
  {label:"3ème",     salles:3, classes:["3ème ALL","3ème ARB","3ème CHN","3ème ESP","3ème ITA"]},
  {label:"2nde A4",  salles:2, classes:["2nde ALL","2nde ARB","2nde CHN","2nde ITA","2nde ESP"]},
  {label:"2nde C",   salles:1, classes:["2nde C"]},
  {label:"1ère A4",  salles:3, classes:["1ère A4 ALL","1ère A4 ARB","1ère A4 ESP","1ère CHN","1ère ITA"]},
  {label:"1ère S",   salles:1, classes:["1ère C","1ère Ti"]},
  {label:"1ère D",   salles:1, classes:["1ère D"]},
  {label:"Tle A4",   salles:2, classes:["Tle A4 ALL","Tle A4 ARB","Tle A4 CHN","Tle A4 ESP","Tle A4 ITA"]},
  {label:"Tle S",    salles:1, classes:["Tle C","Tle Ti"]},
  {label:"Tle D",    salles:1, classes:["Tle D"]},
];

// Calcule les statistiques complètes d'une classe pour la Compilation
// (programme, digitalisation, heures, TP, réussite) — partagé entre impression et export Excel
function calcClasseCompilation(cl, data, trim, nbSemP) {
  const prog = data?.prog||{};
  const notesIndex = data?.notes||{};
    const code = resolveProgCode(cl);
    const meta = code ? PROG_META[code] : null;
    const eleves = ELEVES_DB[cl]||[];
    const total = eleves.length;
    const g = eleves.filter(e=>e.g==="M").length;
    const f = total - g;

    if (!code || !meta) {
      return { cl, g, f, total, lp:0, lf:0, tauxLP:null, ldTot:0, ldFait:0, tauxLD:null,
        hd:0, hf:0, tauxHD:null, tpP:0, tpF:0, tauxTP:null, pctG:null, pctF:null, pctT:null, moy:null };
    }

    // Leçons faites — agrégées tous enseignants confondus pour cette classe
    const allFaites = new Set();
    const allDig = new Set();
    Object.entries(prog).forEach(([key,arr])=>{
      if (key.endsWith("||"+cl)) (arr||[]).forEach(n=>allFaites.add(n));
      if (key.endsWith("||"+cl+"||dig")) (arr||[]).forEach(n=>allDig.add(n));
    });

    const range = trim!=="ANN" ? getTrimRange(code,trim) : null;
    const leconsCode = LECONS_DATA[code]||[];
    const leconsPeriode = range ? leconsCode.filter(l=>l.n>=range[0]&&l.n<=range[1]) : leconsCode;

    const lp = leconsPeriode.length || meta.lpRef;
    const lf = [...allFaites].filter(n=>!range || (n>=range[0]&&n<=range[1])).length;
    const tauxLP = lp>0 ? Math.min(100, Math.round(lf/lp*100)) : null;

    const leconsDig = leconsPeriode.filter(l=>l.d===1);
    const ldTot = leconsDig.length;
    const ldFait = leconsDig.filter(l=>allDig.has(l.n)).length;
    const tauxLD = ldTot>0 ? Math.min(100, Math.round(ldFait/ldTot*100)) : null;

    const hd = trim==="ANN" ? meta.hd : Math.round(meta.hd*nbSemP/36);
    const hf = lp>0 ? Math.round(lf*(meta.hd/meta.lpRef)) : 0;
    const tauxHD = hd>0 ? Math.min(100, Math.round(hf/hd*100)) : null;

    const tpAll = (meta.tp||[]);
    const tpPeriode = range ? tpAll.filter(n=>n>=range[0]&&n<=range[1]) : tpAll;
    const tpP = tpPeriode.length;
    const tpF = tpPeriode.filter(n=>allFaites.has(n)).length;
    const tauxTP = tpP>0 ? Math.min(100, Math.round(tpF/tpP*100)) : null;

    // Réussite — toutes évaluations de cette classe sur la période
    const trimsToAvg = trim==="ANN" ? ["T1","T2","T3"] : [trim];
    const getN = (eid,t,e) => { const v=(notesIndex[`${cl}||${t}-${e}`]||{})[eid]; return (v===undefined||v===""||v===null)?null:+v; };
    const moyTrimEl = (eid,t) => { const e1=getN(eid,t,"E1"), e2=getN(eid,t,"E2"); if(e1===null&&e2===null) return null; if(e1!==null&&e2!==null) return (e1+e2)/2; return e1!==null?e1:e2; };
    const moyEleve = (eid) => { const ms=trimsToAvg.map(t=>moyTrimEl(eid,t)).filter(m=>m!==null); return ms.length>0?ms.reduce((s,m)=>s+m,0)/ms.length:null; };
    let mG=[], mF=[], mAll=[];
    eleves.forEach(el=>{ const m=moyEleve(el.id); if(m!==null){ mAll.push(m); if(el.g==="M") mG.push(m); else mF.push(m); } });
    const pctReuss = a => a.length>0 ? Math.round(a.filter(n=>n>=10).length/a.length*100) : null;
    const pctG = pctReuss(mG), pctF = pctReuss(mF), pctT = pctReuss(mAll);
    const moy = mAll.length>0 ? (mAll.reduce((s,n)=>s+n,0)/mAll.length).toFixed(2) : null;

    return { cl, g, f, total, lp, lf, tauxLP, ldTot, ldFait, tauxLD, hd, hf, tauxHD, tpP, tpF, tauxTP, pctG, pctF, pctT, moy };
}

// Export Excel de la Compilation — réutilise exactement le même calcul que l'impression
function exportCompilationExcel(data, trim = "ANN") {
  const nbSemP = trim === "ANN" ? 36 : 12;
  const exportRows = [];
  NIVEAUX_COMPIL.forEach(niv => {
    niv.classes.forEach(cl => {
      const r = calcClasseCompilation(cl, data, trim, nbSemP);
      exportRows.push({
        "Niveau": niv.label, "Classe": r.cl, "Salles niveau": niv.salles,
        "Garçons": r.g, "Filles": r.f, "Total élèves": r.total,
        "Couv. présentiel (%)": r.tauxLP, "Couv. digitalisé (%)": r.tauxLD,
        "Couv. heures (%)": r.tauxHD, "TP présentiel (%)": r.tauxTP,
        "Réussite Garçons (%)": r.pctG, "Réussite Filles (%)": r.pctF, "Réussite Totale (%)": r.pctT,
        "Moyenne générale": r.moy,
      });
    });
  });
  return exportToExcel(`Compilation_${TRIM_LABELS[trim]||"annuelle"}`, "Compilation", exportRows);
}

function genCompilation(data, trim = "ANN") {
  const periode = TRIM_LABELS[trim] || "Année complète";
  const nbSemP  = trim === "ANN" ? 36 : 12;

  const tauxColor = v => v===null ? "#94a3b8" : v>=75?"#16a34a":v>=50?"#e67e22":"#ef4444";
  const fmtPct = v => v===null ? "N/D" : `${v}%`;
  const fmtMoy = v => v===null ? "—" : v;

  let grandG=0, grandF=0, grandT=0, grandSalles=0;

  const niveauxHTML = NIVEAUX_COMPIL.map(niv => {
    const rows = niv.classes.map(cl => calcClasseCompilation(cl, data, trim, nbSemP));
    const sumG = rows.reduce((s,r)=>s+r.g,0), sumF = rows.reduce((s,r)=>s+r.f,0), sumT = sumG+sumF;
    grandG+=sumG; grandF+=sumF; grandT+=sumT; grandSalles+=niv.salles;

    const avgTaux = (key) => { const vals=rows.map(r=>r[key]).filter(v=>v!==null); return vals.length>0?Math.round(vals.reduce((s,v)=>s+v,0)/vals.length):null; };
    const tLP=avgTaux("tauxLP"), tLD=avgTaux("tauxLD"), tHD=avgTaux("tauxHD"), tTP=avgTaux("tauxTP"),
          tG=avgTaux("pctG"), tF=avgTaux("pctF"), tT=avgTaux("pctT");
    const moyVals = rows.map(r=>r.moy).filter(v=>v!==null).map(Number);
    const moyNiv = moyVals.length>0 ? (moyVals.reduce((s,v)=>s+v,0)/moyVals.length).toFixed(2) : "—";

    const td = "padding:4px 6px;border:1px solid #ccc;text-align:center;font-size:8px";
    const rowsHTML = rows.map(r => `
      <tr>
        <td style="${td};text-align:left;font-weight:600">${r.cl}</td>
        <td style="${td}">${r.g}</td><td style="${td}">${r.f}</td><td style="${td};font-weight:700">${r.total}</td>
        <td style="${td};color:${tauxColor(r.tauxLP)};font-weight:700">${fmtPct(r.tauxLP)}</td>
        <td style="${td};color:${tauxColor(r.tauxLD)};font-weight:700">${fmtPct(r.tauxLD)}</td>
        <td style="${td};color:${tauxColor(r.tauxHD)};font-weight:700">${fmtPct(r.tauxHD)}</td>
        <td style="${td};color:${tauxColor(r.tauxTP)};font-weight:700">${fmtPct(r.tauxTP)}</td>
        <td style="${td};color:#94a3b8;font-style:italic">N/D</td>
        <td style="${td}">${r.pctG===null?"—":r.pctG+"%"}</td>
        <td style="${td}">${r.pctF===null?"—":r.pctF+"%"}</td>
        <td style="${td};font-weight:700">${r.pctT===null?"—":r.pctT+"%"}</td>
        <td style="${td};font-weight:700">${fmtMoy(r.moy)}</td>
      </tr>`).join("");

    return `
    <h3 style="font-size:10px;font-weight:bold;margin:12px 0 3px;padding:5px 10px;background:#1a6b3c;color:#fff">
      ${niv.label} — ${niv.salles} salle${niv.salles>1?"s":""} fonctionnelle${niv.salles>1?"s":""}
    </h3>
    <table style="width:100%;border-collapse:collapse;font-size:8px">
      <thead>
        <tr style="background:#dcfce7">
          <th rowspan="2" style="${td};background:#166534;color:#fff">Classe</th>
          <th colspan="3" style="${td};background:#166534;color:#fff">Effectifs</th>
          <th style="${td};background:#166534;color:#fff">Couv.<br>présentiel</th>
          <th style="${td};background:#166534;color:#fff">Couv.<br>digitalisé</th>
          <th style="${td};background:#166534;color:#fff">Couv.<br>heures</th>
          <th style="${td};background:#166534;color:#fff">TP<br>présentiel</th>
          <th style="${td};background:#166534;color:#fff">TP<br>digitalisé</th>
          <th colspan="3" style="${td};background:#166534;color:#fff">Réussite (%)</th>
          <th style="${td};background:#166534;color:#fff">Moy.<br>Gén.</th>
        </tr>
        <tr style="background:#dcfce7">
          <th style="${td}">G</th><th style="${td}">F</th><th style="${td}">T</th>
          <th style="${td}"></th><th style="${td}"></th><th style="${td}"></th><th style="${td}"></th><th style="${td}"></th>
          <th style="${td}">G</th><th style="${td}">F</th><th style="${td}">T</th>
          <th style="${td}"></th>
        </tr>
      </thead>
      <tbody>${rowsHTML}
        <tr style="background:#f1f5f9;font-weight:800">
          <td style="${td};text-align:left">TOTAL ${niv.label}</td>
          <td style="${td}">${sumG}</td><td style="${td}">${sumF}</td><td style="${td}">${sumT}</td>
          <td style="${td};color:${tauxColor(tLP)}">${fmtPct(tLP)}</td>
          <td style="${td};color:${tauxColor(tLD)}">${fmtPct(tLD)}</td>
          <td style="${td};color:${tauxColor(tHD)}">${fmtPct(tHD)}</td>
          <td style="${td};color:${tauxColor(tTP)}">${fmtPct(tTP)}</td>
          <td style="${td};color:#94a3b8;font-style:italic">N/D</td>
          <td style="${td}">${tG===null?"—":tG+"%"}</td>
          <td style="${td}">${tF===null?"—":tF+"%"}</td>
          <td style="${td}">${tT===null?"—":tT+"%"}</td>
          <td style="${td}">${moyNiv}</td>
        </tr>
      </tbody>
    </table>`;
  }).join("");

  return `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8">
<style>
  body { font-family: Arial, sans-serif; font-size: 10px; margin: 16px; color: #111; }
  table { width: 100%; border-collapse: collapse; }
  @media print { body { margin: 8px; } }
</style>
</head>
<body>
  ${enteteOfficiel("FICHE DE COMPILATION PAR NIVEAU", `Département SVTEEHB &nbsp;|&nbsp; ${periode}`)}
  ${niveauxHTML}
  <table style="margin-top:10px;font-size:9px">
    <tr style="background:#166534;color:#fff;font-weight:800">
      <td style="padding:6px 10px;border:1px solid #ccc">TOTAL GÉNÉRAL</td>
      <td style="padding:6px 10px;border:1px solid #ccc;text-align:center">${grandSalles} salles</td>
      <td style="padding:6px 10px;border:1px solid #ccc;text-align:center">${grandG} G</td>
      <td style="padding:6px 10px;border:1px solid #ccc;text-align:center">${grandF} F</td>
      <td style="padding:6px 10px;border:1px solid #ccc;text-align:center">${grandT} élèves</td>
    </tr>
  </table>
  <div style="margin-top:8px;font-size:7.5px;color:#666;font-style:italic">
    "TP digitalisé" : N/D — aucune donnée distincte entre TP en présentiel et TP digitalisé n'est actuellement collectée dans l'application.
  </div>
  <div style="margin-top:16px;display:grid;grid-template-columns:1fr 1fr;gap:30px;font-size:9px">
    <div><div style="font-weight:bold;margin-bottom:10px">L'Animateur(trice) Pédagogique</div><div style="border-top:1px solid #999;margin-top:30px;padding-top:4px">Signature &amp; Cachet</div></div>
    <div><div style="font-weight:bold;margin-bottom:10px">Le Proviseur</div><div style="border-top:1px solid #999;margin-top:30px;padding-top:4px">Signature &amp; Cachet</div></div>
  </div>
  <script>window.onload=()=>window.print();</script>
</body>
</html>`;
}
// ═══════════════════════════════════════════════════
// IMPRESSION VIA IFRAME CACHÉ (méthode anti popup-blocker)
// ═══════════════════════════════════════════════════
// printGuard: déclaré globalement
function imprimerHTML(html) {
  if (printGuard) return;
  printGuard = true;
  const iframe = document.createElement("iframe");
  iframe.style.cssText = "position:fixed;top:-9999px;left:-9999px;width:1px;height:1px;";
  document.body.appendChild(iframe);
  const doc = iframe.contentDocument || iframe.contentWindow.document;
  doc.open(); doc.write(html); doc.close();
  setTimeout(() => {
    iframe.contentWindow.focus();
    iframe.contentWindow.print();
    setTimeout(() => { document.body.removeChild(iframe); printGuard = false; }, 2000);
  }, 400);
}

// ═══════════════════════════════════════════════════
// COMPOSANTS UI
// ═══════════════════════════════════════════════════
const DocSp = ({ size=20, color=C.green }) => (
  <div style={{ width:size,height:size,border:`2px solid ${color}30`,borderTopColor:color,borderRadius:"50%",animation:"spin .7s linear infinite" }}/>
);

const DocSk = ({ h=16, w="100%", br=6 }) => (
  <div style={{ height:h,width:w,borderRadius:br,background:"linear-gradient(90deg,#e2e8f0 25%,#f1f5f9 50%,#e2e8f0 75%)",backgroundSize:"200% 100%",animation:"shimmer 1.4s infinite" }}/>
);

const DocPi = ({ ch, color=C.green }) => (
  <span style={{ display:"inline-flex",alignItems:"center",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700,background:color+"18",color,whiteSpace:"nowrap" }}>{ch}</span>
);

const DocTaux = ({ taux }) => {
  const color = taux >= 75 ? C.green : taux >= 50 ? C.orange : C.red;
  return <DocPi ch={`${taux}%`} color={color}/>;
};

// ═══════════════════════════════════════════════════
// TABLEAU DE PRÉVISUALISATION
// ═══════════════════════════════════════════════════

// ═══════════════════════════════════════════════════
// CARTE DOCUMENT
// ═══════════════════════════════════════════════════

// ═══════════════════════════════════════════════════
// PAGE PRINCIPALE
// ═══════════════════════════════════════════════════

function DocumentsPage() {
  const {isMobile} = useDevice();
  const {data, refreshData, showToast, pendingFicheEns, setPendingFicheEns} = useApp();
  const [refreshing, setRefreshing] = useState(false);
  const [docTab, setDocTab]     = useState("compilation"); // compilation | fiches | bilan
  const [trim, setTrim]         = useState("ANN");
  const [selEns, setSelEns]     = useState("all");
  const [previewHtml, setPreviewHtml] = useState(null);  // HTML de l'aperçu
  const [previewLabel, setPreviewLabel] = useState("");
  const iframeRef = useRef(null);
  const previewMainRef = useRef(null);


  // Navigation ciblée depuis le Tableau de bord ("voir la fiche de cet enseignant")
  useEffect(() => {
    if (!pendingFicheEns || !data) return;
    setDocTab("fiches");
    aperçuFicheEns(pendingFicheEns);
    setPendingFicheEns(null);
  }, [pendingFicheEns, data]);

  const refreshDocsData = async () => {
    setRefreshing(true);
    try {
      await refreshData();
      showToast("✓ Données actualisées — fiches à jour");
    } catch {
      showToast("⚠ Actualisation impossible — vérifiez la connexion", false);
    } finally {
      setRefreshing(false);
    }
  };

  const enseignants = data ? Object.values(data?.users||{}).filter(u => u.role !== "proviseur") : [];
  const getDeptInfo = (e) => {
    const dId = e?.departement_id || 1;
    const deptNom = DEPARTEMENTS_LIST.find(d=>d.id===dId)?.nom || "SVTEEHB";
    const anim = Object.values(data?.users||{}).find(u=>(u.role==="animateur"||u.role==="animatrice") && (u.departement_id||1)===dId);
    return { deptNom, animateurNom: anim?.nom || "—" };
  };

  // ── Aperçu dans iframe ────────────────────────────────────────────
  function afficherApercu(html, label) {
    setPreviewHtml(stripAutoPrint(html));
    setPreviewLabel(label);
    // Injecter dans l'iframe après le render
    setTimeout(() => {
      if (iframeRef.current) {
        const doc = iframeRef.current.contentDocument || iframeRef.current.contentWindow.document;
        // Retirer le script d'impression auto pour l'aperçu
        const htmlSansScript = html.replace(/<script>window\.onload.*?<\/script>/gs, "");
        doc.open(); doc.write(htmlSansScript); doc.close();
      }
      // Sur mobile, faire défiler jusqu'à l'aperçu (situé en dessous de la liste)
      if (isMobile && previewMainRef.current) {
        previewMainRef.current.scrollIntoView({ behavior:"smooth", block:"start" });
      }
    }, 50);
  }

  const [selClasseParEns, setSelClasseParEns] = useState({}); // {ensId: "all" | code_classe}

  function aperçuFicheEns(ensId) {
    if (!data) return;
    const ens = (data?.users||{})[ensId];
    if (!ens) return;
    const classeSel = selClasseParEns[ensId] || "all";
    const classesAGenerer = classeSel==="all" ? (ens.classes||[]) : [classeSel];
    const html = genFicheSuivi(ens, classesAGenerer, (data?.prog||{}), trim, (data?.notes||{}), (data?.absences||{}), ...Object.values(getDeptInfo(ens)));
    afficherApercu(html, `Fiche — ${ens.nom}${classeSel==="all"?"":` — ${classeSel}`} — ${TRIM_LABELS[trim]||"Année"}`);
  }

  function aperçuCompilation() {
    if (!data) return;
    const html = genCompilation(data, trim);
    afficherApercu(html, `Compilation globale — ${TRIM_LABELS[trim]||"Année"}`);
  }

  function aperçuBilan() {
    if (!data) return;
    const html = genBilanTrimestre(trim, data);
    afficherApercu(html, `Bilan trimestriel — ${TRIM_LABELS[trim]||"Année"}`);
  }

  // ── Impression ────────────────────────────────────────────────────
  function imprimerApercu() {
    if (!previewHtml) return;
    imprimerHTML(previewHtml);
  }

  function imprimerFicheEns(ensId) {
    if (!data) return;
    const ens = (data?.users||{})[ensId];
    if (!ens) return;
    const classeSel = selClasseParEns[ensId] || "all";
    const classesAGenerer = classeSel==="all" ? (ens.classes||[]) : [classeSel];
    imprimerHTML(genFicheSuivi(ens, classesAGenerer, (data?.prog||{}), trim, (data?.notes||{}), (data?.absences||{}), ...Object.values(getDeptInfo(ens))));
  }

  function imprimerToutesFiches() {
    if (!data) return;
    enseignants.forEach((ens, i) => {
      setTimeout(() => {
        imprimerHTML(genFicheSuivi(ens, ens.classes||[], (data?.prog||{}), trim, (data?.notes||{}), (data?.absences||{}), ...Object.values(getDeptInfo(ens))));
      }, i * 2200);
    });
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", flex:1, minHeight:0, overflow:"hidden" }}>

      {/* ── Barre de contrôle Documents (remplace l'ancien header propre) ── */}
      <div style={{ background:C.white, borderBottom:`1px solid ${C.border}`, padding: isMobile?"10px 14px":"8px 20px", display:"flex", flexDirection: isMobile?"column":"row", alignItems: isMobile?"stretch":"center", gap:10, flexShrink:0 }}>
        <button onClick={refreshDocsData} disabled={refreshing}
          style={{ display:"flex", alignItems:"center", justifyContent:"center", gap:6,
            padding: isMobile?"10px 12px":"6px 12px", background:"#eff6ff", border:"1px solid #bfdbfe",
            borderRadius:8, fontSize: isMobile?13:12, fontWeight:700, color:"#1e40af",
            cursor:refreshing?"not-allowed":"pointer", fontFamily:"inherit", flexShrink:0 }}>
          {refreshing ? <><Spinner size={13} color="#1e40af"/> Actualisation…</> : <>🔄 Actualiser les données</>}
        </button>
        {!isMobile && <div style={{ flex:1 }}/>}
        {/* Période */}
        <div style={{ display:"flex", flexDirection: isMobile?"column":"row", gap:8, alignItems: isMobile?"stretch":"center" }}>
          {!isMobile && <span style={{ fontSize:11, color:C.txtMuted }}>Période :</span>}
          <select value={trim} onChange={e=>setTrim(e.target.value)}
            style={{ padding: isMobile?"10px 12px":"5px 10px", border:`1px solid ${C.border}`, borderRadius:8, fontSize: isMobile?13:12, color:C.txt, background:C.white }}>
            <option value="ANN">📅 Année complète</option>
            <option value="T1">Trimestre 1</option>
            <option value="T2">Trimestre 2</option>
            <option value="T3">Trimestre 3</option>
          </select>

          {/* Aperçu si présent : bouton imprimer */}
          {previewHtml && (
            <button onClick={imprimerApercu}
              style={{ padding: isMobile?"11px 14px":"6px 14px", background:`linear-gradient(135deg,${C.greenDark},${C.green})`, color:"#fff", border:"none", borderRadius:8, fontSize: isMobile?13:12, fontWeight:700, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", gap:6, whiteSpace:"nowrap" }}>
              🖨️ {isMobile?"Imprimer":"Imprimer ce document"}
            </button>
          )}
        </div>
      </div>

      {/* ── Layout 2 colonnes (desktop) / empilé (mobile) ──────────────── */}
      <div style={{ display:"flex", flexDirection: isMobile?"column":"row", flex:1, minHeight:0, overflow:"hidden" }}>

        {/* ── Panneau gauche — sélection ─────────────────────────────── */}
        <aside style={{
          width: isMobile?"100%":280, minWidth: isMobile?"auto":280,
          maxHeight: isMobile?"45vh":undefined,
          borderRight: isMobile?"none":`1px solid ${C.border}`,
          borderBottom: isMobile?`1px solid ${C.border}`:"none",
          overflowY:"auto", minHeight:0, padding: isMobile?12:16,
          display:"flex", flexDirection:"column", gap:12, background:C.white, flexShrink:0 }}>

          <h2 style={{ fontSize:13, fontWeight:700, color:C.txt }}>📋 Choisir un document</h2>

          {/* Onglets de catégorie de document */}
          <div style={{ display:"flex", flexDirection: isMobile?"row":"column", gap:6,
            overflowX: isMobile?"auto":"visible" }}>
            {[
              {id:"compilation", label:"📊 Compilation annuelle", sub:"Vue par niveau, toutes classes"},
              {id:"fiches",      label:"👤 Fiches individuelles", sub:"Une fiche par enseignant"},
              {id:"bilan",       label:"📋 Bilan trimestriel",    sub:"Classement par taux de couverture"},
            ].map(t=>(
              <button key={t.id} onClick={()=>{ setDocTab(t.id); setPreviewHtml(null); setPreviewLabel(""); }}
                style={{
                  flexShrink:0, textAlign:"left", padding: isMobile?"10px 12px":"9px 12px",
                  borderRadius:10, border:`1.5px solid ${docTab===t.id?C.green:C.border}`,
                  background:docTab===t.id?C.greenPale:C.white, cursor:"pointer", fontFamily:"inherit",
                  minWidth: isMobile?160:"auto" }}>
                <div style={{ fontSize:12.5, fontWeight:800, color:docTab===t.id?C.greenDark:C.txt }}>{t.label}</div>
                {!isMobile && <div style={{ fontSize:10, color:C.txtMuted, marginTop:2 }}>{t.sub}</div>}
              </button>
            ))}
          </div>

          {!data ? (
            <div style={{ textAlign:"center", padding:30, color:C.txtMuted }}>
              <Spinner size={20} color={C.green}/>
              <div style={{ marginTop:8, fontSize:12 }}>Chargement…</div>
            </div>
          ) : (
            <>
              {/* ── Onglet : Compilation annuelle ─────────────────── */}
              {docTab==="compilation" && (
                <div style={{ background:C.greenPale, border:`1.5px solid ${C.greenBorder}`, borderRadius: isMobile?14:10, padding: isMobile?"14px 16px":"12px 14px" }}>
                  <div style={{ fontSize: isMobile?14:13, fontWeight:800, color:C.green, marginBottom:4 }}>📊 Compilation globale</div>
                  <div style={{ fontSize: isMobile?11.5:11, color:C.txtMuted, marginBottom:10, lineHeight:1.5 }}>
                    Toutes les classes · Tous les enseignants · Par niveau
                  </div>
                  <div style={{ display:"flex", gap:8 }}>
                    <button onClick={aperçuCompilation}
                      style={{ flex:1, padding: isMobile?"11px":"7px", background:C.white, border:`1px solid ${C.greenBorder}`, borderRadius: isMobile?10:8, fontSize: isMobile?12.5:11, fontWeight:700, cursor:"pointer", color:C.green }}>
                      👁 Aperçu
                    </button>
                    <button onClick={()=>imprimerHTML(genCompilation(data,trim))}
                      style={{ flex:1, padding: isMobile?"11px":"7px", background:C.green, border:"none", borderRadius: isMobile?10:8, fontSize: isMobile?12.5:11, fontWeight:700, cursor:"pointer", color:"#fff" }}>
                      🖨️ Imprimer
                    </button>
                  </div>
                  <button onClick={()=>{
                      const ok = exportCompilationExcel(data, trim);
                      showToast(ok ? "✓ Fichier Excel téléchargé" : "⚠ Échec de l'export", ok);
                    }}
                    style={{ width:"100%", marginTop:8, padding: isMobile?"10px":"7px", background:"#f0fdf4", border:"1.5px solid #15803d", borderRadius: isMobile?10:8, fontSize: isMobile?12.5:11, fontWeight:700, cursor:"pointer", color:"#15803d" }}>
                    📥 Exporter Excel
                  </button>
                </div>
              )}

              {/* ── Onglet : Fiches individuelles ────────────────── */}
              {docTab==="fiches" && (
                <div style={{ background:C.bluePale, border:`1.5px solid ${C.blue}30`, borderRadius:10, padding:"12px 14px" }}>
                  <div style={{ fontSize:13, fontWeight:800, color:C.blue, marginBottom:4 }}>👨‍🏫 Fiches individuelles</div>
                  <div style={{ fontSize:11, color:C.txtMuted, marginBottom:10 }}>
                    Sélectionner un enseignant pour aperçu ou impression directe
                  </div>
                  <div style={{ display:"flex", flexDirection:"column", gap: isMobile?8:6 }}>
                    {enseignants.map(ens => (
                      <div key={ens.id} style={{
                        display:"flex", flexDirection:"column", alignItems:"stretch",
                        gap: isMobile?8:6, padding: isMobile?"10px 12px":"8px 10px",
                        background:previewLabel.includes(ens.nom)?C.greenPale:C.white,
                        border:`1px solid ${previewLabel.includes(ens.nom)?C.greenBorder:C.border}`,
                        borderRadius: isMobile?12:8, transition:"all .15s" }}>
                        <div style={{display:"flex", alignItems:"center", gap:8, minWidth:0}}>
                          <Avatar ens={ens} size={isMobile?26:22} fontSize={isMobile?9:8}/>
                          <span style={{ flex:1, minWidth:0, fontSize: isMobile?13:11, fontWeight:600, color:C.txt, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>
                            {ens.nom}
                          </span>
                          <span style={{ fontSize: isMobile?10.5:9, color:C.txtMuted, flexShrink:0 }}>
                            {(ens.classes||[]).length} classe{(ens.classes||[]).length>1?"s":""}
                          </span>
                        </div>
                        {(ens.classes||[]).length>1 && (
                          <select value={selClasseParEns[ens.id]||"all"}
                            onChange={e=>setSelClasseParEns(prev=>({...prev,[ens.id]:e.target.value}))}
                            style={{ width:"100%", padding: isMobile?"7px 10px":"5px 8px", border:`1px solid ${C.border}`, borderRadius:7, fontSize: isMobile?12:11, fontFamily:"inherit", background:"#f8fafc", color:C.txt }}>
                            <option value="all">Toutes les classes ({(ens.classes||[]).length})</option>
                            {(ens.classes||[]).map(cl=><option key={cl} value={cl}>{cl}</option>)}
                          </select>
                        )}
                        <div style={{display:"flex", gap:8}}>
                          <button onClick={() => aperçuFicheEns(ens.id)}
                            style={{ flex:1, padding: isMobile?"9px 10px":"6px 8px", background:C.greenPale, border:`1px solid ${C.greenBorder}`, borderRadius: isMobile?9:6, fontSize: isMobile?12:11, fontWeight:700, cursor:"pointer", color:C.green, display:"flex", alignItems:"center", justifyContent:"center", gap:5 }}>
                            👁 Aperçu
                          </button>
                          <button onClick={() => imprimerFicheEns(ens.id)}
                            style={{ flex:1, padding: isMobile?"9px 10px":"6px 8px", background:C.green, border:"none", borderRadius: isMobile?9:6, fontSize: isMobile?12:11, fontWeight:700, cursor:"pointer", color:"#fff", display:"flex", alignItems:"center", justifyContent:"center", gap:5 }}>
                            🖨️ Imprimer
                          </button>
                        </div>
                      </div>
                    ))}
                    {enseignants.length===0 && (
                      <div style={{textAlign:"center", padding:20, color:C.txtLight, fontSize:12}}>Aucun enseignant trouvé</div>
                    )}
                  </div>
                  <button onClick={imprimerToutesFiches}
                    style={{ marginTop:10, width:"100%", padding:"8px", background:`linear-gradient(135deg,${C.greenDark},${C.green})`, color:"#fff", border:"none", borderRadius:8, fontSize:11, fontWeight:700, cursor:"pointer" }}>
                    🖨️ Imprimer toutes les fiches
                  </button>
                </div>
              )}

              {/* ── Onglet : Bilan trimestriel ────────────────────── */}
              {docTab==="bilan" && (
                <div style={{ background:"#fef3c7", border:`1.5px solid #fcd34d`, borderRadius: isMobile?14:10, padding: isMobile?"14px 16px":"12px 14px" }}>
                  <div style={{ fontSize: isMobile?14:13, fontWeight:800, color:"#92400e", marginBottom:4 }}>📋 Bilan trimestriel</div>
                  <div style={{ fontSize: isMobile?11.5:11, color:C.txtMuted, marginBottom:10, lineHeight:1.5 }}>
                    Classement des enseignants par taux de couverture — alertes et objectifs atteints
                  </div>
                  <div style={{ display:"flex", gap:8 }}>
                    <button onClick={aperçuBilan}
                      style={{ flex:1, padding: isMobile?"11px":"7px", background:C.white, border:`1px solid #fcd34d`, borderRadius: isMobile?10:8, fontSize: isMobile?12.5:11, fontWeight:700, cursor:"pointer", color:"#92400e" }}>
                      👁 Aperçu
                    </button>
                    <button onClick={()=>imprimerHTML(genBilanTrimestre(trim,data))}
                      style={{ flex:1, padding: isMobile?"11px":"7px", background:"#d97706", border:"none", borderRadius: isMobile?10:8, fontSize: isMobile?12.5:11, fontWeight:700, cursor:"pointer", color:"#fff" }}>
                      🖨️ Imprimer
                    </button>
                  </div>
                </div>
              )}

              {/* Note */}
              <div style={{ background:C.goldPale, border:`1px solid ${C.gold}40`, borderRadius:8, padding:"10px 12px", fontSize:10, color:C.txtMuted, lineHeight:1.6 }}>
                ℹ️ {isMobile
                  ? <>L'aperçu s'affiche <strong>juste en dessous</strong>. Cliquez <strong>👁 Aperçu</strong> pour visualiser avant d'imprimer.</>
                  : <>L'aperçu s'affiche à droite. Cliquez <strong>👁 Aperçu</strong> pour visualiser avant d'imprimer.</>}
              </div>
            </>
          )}
        </aside>

        {/* ── Panneau droit — aperçu iframe ──────────────────────────── */}
        <main ref={previewMainRef} style={{ flex:1, display:"flex", flexDirection:"column", overflow:"hidden" }}>

          {/* Header aperçu */}
          <div style={{ padding: isMobile?"10px 14px":"10px 20px", borderBottom:`1px solid ${C.border}`, display:"flex", flexWrap:"wrap", alignItems:"center", gap:10, flexShrink:0, background:C.white }}>
            <div style={{ flex:1, minWidth:"60%" }}>
              <h3 style={{ margin:0, fontSize: isMobile?12.5:13, fontWeight:700, color:C.txt }}>
                {previewHtml ? `📄 ${previewLabel}` : "📄 Aperçu du document"}
              </h3>
              <p style={{ margin:"2px 0 0", fontSize: isMobile?10.5:11, color:C.txtMuted }}>
                {previewHtml
                  ? (isMobile ? "Touchez 🖨️ pour imprimer" : "Cliquez 🖨️ dans la topbar pour imprimer ce document")
                  : (isMobile ? "Sélectionnez un document ci-dessus" : "Sélectionnez un document dans le panneau gauche")}
              </p>
            </div>
            {previewHtml && (
              <div style={{ display:"flex", gap:8 }}>
                <button onClick={imprimerApercu}
                  style={{ padding: isMobile?"8px 13px":"6px 14px", background:`linear-gradient(135deg,${C.greenDark},${C.green})`, color:"#fff", border:"none", borderRadius:8, fontSize: isMobile?12:12, fontWeight:700, cursor:"pointer" }}>
                  🖨️ Imprimer
                </button>
                <button onClick={()=>{ setPreviewHtml(null); setPreviewLabel(""); }}
                  style={{ padding: isMobile?"8px 11px":"6px 12px", background:C.white, border:`1px solid ${C.border}`, borderRadius:8, fontSize:11, color:C.txtMuted, cursor:"pointer" }}>
                  ✕ Fermer
                </button>
              </div>
            )}
          </div>

          {/* Zone iframe */}
          <div style={{ flex:1, minHeight: isMobile?"60vh":undefined, overflow:"hidden", background:"#e5e7eb", padding:previewHtml?(isMobile?10:16):0, display:"flex", alignItems:"center", justifyContent:"center" }}>
            {previewHtml ? (
              <iframe
                ref={iframeRef}
                title="Aperçu document MINESEC"
                style={{
                  width:"100%", height:"100%",
                  border:"none", borderRadius:8,
                  background:"#fff",
                  boxShadow:"0 4px 24px rgba(0,0,0,.15)",
                }}
              />
            ) : (
              <div style={{ textAlign:"center", color:C.txtLight }}>
                <div style={{ fontSize:56, marginBottom:16 }}>📄</div>
                <div style={{ fontSize:15, fontWeight:700, color:C.txt, marginBottom:8 }}>
                  Aucun aperçu sélectionné
                </div>
                <div style={{ fontSize:12, color:C.txtMuted, maxWidth:320 }}>
                  Cliquez sur <strong>👁 Aperçu</strong> à côté d'un document pour le visualiser ici avant de l'imprimer.
                </div>
              </div>
            )}
          </div>
        </main>
      </div>
    </div>
  );
}

// getNowInfo: définie globalement

// ═══════════════════════════════════════════════════
// PALETTE
// ═══════════════════════════════════════════════════
// const C: référence au bloc global ci-dessus

// ═══════════════════════════════════════════════════
// COMPOSANTS UI
// ═══════════════════════════════════════════════════
const EdtSp = ({ size=18, color=C.green }) => (
  <span style={{ width:size,height:size,border:`2px solid ${color}30`,borderTopColor:color,borderRadius:"50%",display:"inline-block",animation:"spin .7s linear infinite",flexShrink:0 }}/>
);

const EdtSk = ({ h=16, w="100%", br=6 }) => (
  <div style={{ height:h,width:w,borderRadius:br,background:"linear-gradient(90deg,#e2e8f0 25%,#f1f5f9 50%,#e2e8f0 75%)",backgroundSize:"200% 100%",animation:"shimmer 1.4s infinite" }}/>
);

const EdtToast = ({ msg, ok }) => (
  <div style={{ position:"fixed",bottom:24,right:24,zIndex:9999,display:"flex",alignItems:"center",gap:10,padding:"12px 18px",borderRadius:12,background:ok===false?C.red:C.greenDark,color:"#fff",fontSize:13,fontWeight:600,boxShadow:"0 8px 24px rgba(0,0,0,.25)",animation:"fadeUp .3s ease" }}>
    {msg}
  </div>
);

// ══════════════════════════════════════════════════════════════════════
// 5. MON EMPLOI DU TEMPS — classes de l'enseignant uniquement
// ══════════════════════════════════════════════════════════════════════
function MonEdtPage() {
  const {user} = useApp();
  const {jk: nowJk, hi: nowHi} = getNowInfo();
  const ensId = user?.id;
  const edt   = EDT_REEL[ensId] || {};

  // Classes SVTEEHB de cet enseignant
  const classesDansEDT = new Set();
  Object.values(edt).forEach(jour=>jour.forEach(cl=>{if(cl)classesDansEDT.add(cl);}));

  return (
    <div style={{padding:"20px", display:"flex", flexDirection:"column", gap:16}}>

      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"14px 18px", display:"flex", justifyContent:"space-between", alignItems:"center"}}>
        <div>
          <h2 style={{fontSize:16, fontWeight:800, color:C.txt, margin:0}}>🗓 Mon emploi du temps SVTEEHB</h2>
          <p style={{fontSize:11, color:C.txtMuted, margin:"4px 0 0"}}>{user?.nom} · {(user?.classes||[]).length} classe{(user?.classes||[]).length>1?"s":""} · 2025–2026</p>
        </div>
        {nowJk && nowHi>=0 && edt[nowJk]?.[nowHi] && (
          <div style={{background:C.greenPale, border:`1px solid ${C.greenBorder}`, borderRadius:9, padding:"8px 14px", textAlign:"center"}}>
            <div style={{fontSize:9, color:C.green, fontWeight:700, marginBottom:2}}>🟢 EN COURS</div>
            <div style={{fontSize:14, fontWeight:800, color:C.green}}>{edt[nowJk]?.[nowHi]}</div>
          </div>
        )}
      </div>

      {/* Grille EDT */}
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, overflow:"auto"}}>
        <table style={{width:"100%", borderCollapse:"collapse", fontSize:11, minWidth:600}}>
          <thead>
            <tr style={{background:"#0f1f14"}}>
              <th style={{padding:"10px 12px", textAlign:"left", color:"rgba(255,255,255,.5)", fontSize:10, fontWeight:600, width:100}}>Heure</th>
              {JOURS.map((j,ji)=>(
                <th key={j} style={{padding:"10px 12px", textAlign:"center", color:nowJk===JKEYS[ji]?"#4ade80":"#fff", fontSize:11, fontWeight:nowJk===JKEYS[ji]?800:600}}>
                  {j}
                  {nowJk===JKEYS[ji] && <div style={{fontSize:8, color:"#4ade80", fontWeight:600, marginTop:1}}>Aujourd'hui</div>}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {HEURES.map((h,hi)=>(
              <tr key={hi} style={{borderBottom:`1px solid ${C.border}`}}>
                <td style={{padding:"8px 12px", fontSize:10, color:C.txtMuted, fontWeight:600, background:"#fafafa", whiteSpace:"nowrap"}}>{h}</td>
                {JKEYS.map((jk,ji)=>{
                  const cl = edt[jk]?.[hi];
                  const isNow = jk===nowJk && hi===nowHi;
                  return (
                    <td key={jk} style={{padding:"5px", verticalAlign:"top", background:isNow?"rgba(22,163,74,.05)":"transparent"}}>
                      {cl ? (
                        <div style={{padding:"8px 10px", borderRadius:8, background:isNow?C.green:getColor(ensId)+"20", color:isNow?"#fff":getColor(ensId), border:isNow?`2px solid ${C.green}`:`1px solid ${getColor(ensId)}40`, fontSize:11, fontWeight:700, lineHeight:1.3}}>
                          {cl}
                          {isNow && <div style={{fontSize:9, fontWeight:600, marginTop:3, opacity:.8}}>En cours</div>}
                        </div>
                      ) : (
                        <div style={{padding:"8px", minHeight:40}}/>
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Résumé des classes SVTEEHB */}
      {classesDansEDT.size > 0 && (
        <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"14px 18px"}}>
          <div style={{fontSize:12, fontWeight:700, color:C.txt, marginBottom:10}}>📚 Classes SVTEEHB enseignées</div>
          <div style={{display:"flex", gap:8, flexWrap:"wrap"}}>
            {[...classesDansEDT].map(cl=>(
              <div key={cl} style={{padding:"6px 12px", borderRadius:20, background:getColor(ensId)+"15", border:`1px solid ${getColor(ensId)}40`, color:getColor(ensId), fontSize:12, fontWeight:700}}>
                {cl}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}



// ═══════════════════════════════════════════════════
// GRILLE EDT D'UN ENSEIGNANT (admin, avec édition au clic)
// ═══════════════════════════════════════════════════
function buildEdtRuntime(exceptions = {}, edtBase = {}) {
  const copy = {};
  const useBase = Object.keys(edtBase).length > 0;
  const ensIds = useBase ? Object.keys(edtBase) : Object.keys(EDT_REEL);
  ensIds.forEach(ensId => {
    copy[ensId] = {};
    JKEYS.forEach(jk => { copy[ensId][jk] = HEURES.map(()=>null); });
    if (useBase) {
      Object.entries(edtBase[ensId]||{}).forEach(([slot, lbl]) => {
        const [jk, hiStr] = slot.split("-");
        const hi = parseInt(hiStr, 10);
        if (copy[ensId][jk]) copy[ensId][jk][hi] = lbl || null;
      });
    } else {
      JKEYS.forEach(jk => { copy[ensId][jk] = [...(EDT_REEL[ensId][jk] || HEURES.map(()=>null))]; });
    }
  });
  Object.entries(exceptions).forEach(([ensId, slots]) => {
    if (!copy[ensId]) { copy[ensId] = {}; JKEYS.forEach(jk => { copy[ensId][jk] = HEURES.map(()=>null); }); }
    Object.entries(slots).forEach(([slot, lbl]) => {
      const [jk, hiStr] = slot.split("-");
      const hi = parseInt(hiStr, 10);
      if (!copy[ensId][jk]) copy[ensId][jk] = HEURES.map(()=>null);
      copy[ensId][jk][hi] = lbl || null;
    });
  });
  return copy;
}

function getClassesFromEdt(ensId, edtRt) {
  const set = new Set();
  Object.values(edtRt[ensId]||{}).forEach(jour=>(jour||[]).forEach(cl=>{ if(cl) set.add(cl); }));
  return [...set].sort();
}

function EdtGrid({ ensId, ens, edtRt, isAdmin, onCellClick }) {
  const { jk: nowJk, hi: nowHi } = getNowInfo();
  const edt = edtRt[ensId] || {};
  const col = getColor(ensId);
  return (
    <div style={{ background:C.white, borderRadius:12, border:`1px solid ${C.border}`, overflow:"auto" }}>
      <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11, minWidth:600 }}>
        <thead>
          <tr style={{ background:"#0f1f14" }}>
            <th style={{ padding:"10px 12px", textAlign:"left", color:"rgba(255,255,255,.5)", fontSize:10, fontWeight:600, width:100 }}>Heure</th>
            {JOURS.map((j,ji)=>(
              <th key={j} style={{ padding:"10px 12px", textAlign:"center", color:nowJk===JKEYS[ji]?"#4ade80":"#fff", fontSize:11, fontWeight:nowJk===JKEYS[ji]?800:600 }}>
                {j}
                {nowJk===JKEYS[ji] && <div style={{ fontSize:8, color:"#4ade80", fontWeight:600, marginTop:1 }}>Aujourd'hui</div>}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {HEURES.map((h,hi)=>(
            <tr key={hi} style={{ borderBottom:`1px solid ${C.border}` }}>
              <td style={{ padding:"8px 12px", fontSize:10, color:C.txtMuted, fontWeight:600, background:"#fafafa", whiteSpace:"nowrap" }}>{h}</td>
              {JKEYS.map((jk)=>{
                const cl = edt[jk]?.[hi];
                const isNow = jk===nowJk && hi===nowHi;
                return (
                  <td key={jk} onClick={()=>isAdmin && onCellClick && onCellClick(ensId, jk, hi, cl)}
                    style={{ padding:"5px", verticalAlign:"top", background:isNow?"rgba(22,163,74,.05)":"transparent", cursor:isAdmin?"pointer":"default" }}
                    onMouseEnter={e=>isAdmin && (e.currentTarget.style.background="#f1f5f9")}
                    onMouseLeave={e=>isAdmin && (e.currentTarget.style.background=isNow?"rgba(22,163,74,.05)":"transparent")}>
                    {cl ? (
                      <div style={{ padding:"8px 10px", borderRadius:8, background:isNow?C.green:col+"20", color:isNow?"#fff":col, border:isNow?`2px solid ${C.green}`:`1px solid ${col}40`, fontSize:11, fontWeight:700, lineHeight:1.3 }}>
                        {cl}
                        {isNow && <div style={{ fontSize:9, fontWeight:600, marginTop:3, opacity:.8 }}>En cours</div>}
                      </div>
                    ) : (
                      <div style={{ padding:"8px", minHeight:40, display:"flex", alignItems:"center", justifyContent:"center" }}>
                        {isAdmin && <span style={{ fontSize:14, color:"#cbd5e1" }}>+</span>}
                      </div>
                    )}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function genEdtHTML(ens, edtRt) {
  const rt = edtRt[ens.id] || {};
  let rows = "";
  HEURES.forEach((h, hi) => {
    rows += `<tr><td style="padding:6px;border:1px solid #ccc;font-size:9px;font-weight:700;white-space:nowrap">${h}</td>`;
    JKEYS.forEach(jk => {
      const val = (rt[jk] || [])[hi] || "";
      rows += `<td style="padding:6px;border:1px solid #ccc;font-size:9px;text-align:center">${val}</td>`;
    });
    rows += `</tr>`;
  });
  const headerCols = JOURS.map(j => `<th style="padding:6px;border:1px solid #ccc;background:#1a6b3c;color:#fff;font-size:9px">${j}</th>`).join("");
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>EDT ${ens.nom}</title></head>
  <body style="font-family:Arial,sans-serif;padding:20px">
    ${enteteOfficiel("EMPLOI DU TEMPS", ens.nom, DEPARTEMENTS_LIST.find(d=>d.id===(ens.departement_id||1))?.nom)}
    <h3 style="text-align:center;margin:14px 0;font-size:13px">Emploi du temps — ${ens.nom}</h3>
    <table style="width:100%;border-collapse:collapse">
      <thead><tr><th style="padding:6px;border:1px solid #ccc;background:#1a6b3c;color:#fff;font-size:9px">Heures</th>${headerCols}</tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </body></html>`;
}

function EdtPage() {
  const {data, user, showToast, refreshData} = useApp();
  const {isMobile} = useDevice();
  const [onglet, setOnglet] = useState("maintenant"); // maintenant | parEnseignant | parClasse
  const [selEns, setSelEns] = useState(()=> Object.keys(EDT_REEL)[0]||null);
  const [selCl, setSelCl] = useState(null);
  const [editMode, setEditMode] = useState(false);
  const [edtGrid, setEdtGrid] = useState(null); // copie de travail {L:[8],Ma:[8],Me:[8],J:[8],V:[8]} pendant l'édition
  const [edtJourMobile, setEdtJourMobile] = useState("L");
  const [saving, setSaving] = useState(false);

  const [edtRt, setEdtRt] = useState(()=> buildEdtRuntime(data?.exceptions||{}, data?.edtBase||{}));
  useEffect(()=>{ setEdtRt(buildEdtRuntime(data?.exceptions||{}, data?.edtBase||{})); }, [data?.exceptions]);

  const { jk: nowJk, hi: nowHi } = getNowInfo();

  const enseignants = useMemo(()=>{
    const users = Object.values(data?.users||{}).filter(u=>u.role!=="proviseur");
    const source = (users.length>0 || data?.deptFilterActive) ? users : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant");
    return source.map(u=>({...u, col:getColor(u.id), ini:getIni(u.nom)}));
  }, [data]);

  useEffect(()=>{ if (enseignants.length>0 && !enseignants.some(e=>e.id===selEns)) setSelEns(enseignants[0].id); }, [enseignants]);

  const toutesClasses = useMemo(()=>{
    const set = new Set();
    Object.keys(edtRt).forEach(ensId=>getClassesFromEdt(ensId, edtRt).forEach(cl=>set.add(cl)));
    return [...set].sort();
  }, [edtRt]);

  useEffect(()=>{ if (toutesClasses.length>0 && !toutesClasses.includes(selCl)) setSelCl(toutesClasses[0]); }, [toutesClasses]);

  const entrerModeEdition = () => {
    if (!selEns) return;
    const ensEdt = edtRt[selEns] || {};
    const grid = {};
    JKEYS.forEach(jk => { grid[jk] = [...(ensEdt[jk] || HEURES.map(()=>null))]; });
    setEdtGrid(grid);
    setEditMode(true);
  };

  const annulerEdition = () => { setEditMode(false); setEdtGrid(null); };

  const setEdtCell = (jk, hi, value) => {
    setEdtGrid(prev => {
      const next = {...prev};
      next[jk] = [...next[jk]];
      next[jk][hi] = value || null;
      return next;
    });
  };

  // Retire une classe de TOUS ses créneaux en un clic (pas besoin de chercher case par case)
  const retirerClasseDeEdt = (classe) => {
    setEdtGrid(prev => {
      const next = {};
      JKEYS.forEach(jk => { next[jk] = prev[jk].map(v => v===classe ? null : v); });
      return next;
    });
  };

  const classesDansEdtGrid = edtGrid ? [...new Set(JKEYS.flatMap(jk=>edtGrid[jk]).filter(Boolean))].sort() : [];

  const handleSaveEdt = async () => {
    if (!selEns || !edtGrid) return;

    const classesActuelles = ensActuel?.classes || [];
    const classesDansEdt = [...new Set(JKEYS.flatMap(jk=>edtGrid[jk]).filter(Boolean))];
    const classesAjoutees = classesDansEdt.filter(cl => !classesActuelles.includes(cl));
    const classesRetirees = classesActuelles.filter(cl => !classesDansEdt.includes(cl));

    // Détection de conflit : une classe ajoutée est-elle déjà assignée à un AUTRE enseignant ?
    // (évite qu'une même classe se retrouve chez deux enseignants à la fois)
    const conflits = [];
    classesAjoutees.forEach(cl => {
      enseignants.forEach(autre => {
        if (autre.id !== selEns && (autre.classes||[]).includes(cl)) conflits.push({ classe:cl, autreEns:autre });
      });
    });
    if (conflits.length > 0) {
      const liste = conflits.map(c => `• ${c.classe} (actuellement chez ${getNomCourt(c.autreEns.nom)})`).join("\n");
      const continuer = window.confirm(
        `${conflits.length>1?"Ces classes sont":"Cette classe est"} déjà assignée(s) à un autre enseignant :\n\n${liste}\n\n`+
        `Continuer va la/les retirer de l'autre enseignant (classes assignées + son EDT) pour la/les transférer à ${getNomCourt(ensActuel?.nom)}.\n\nContinuer ?`
      );
      if (!continuer) return;
    }

    setSaving(true);
    // Remplacement complet et propre : admin_set_edt_slots réécrit (upsert) les 40 créneaux
    // en une fois — plus besoin de purge préalable, l'upsert couvre déjà toutes les positions.
    const rows = [];
    JKEYS.forEach(jk => { (edtGrid[jk]||[]).forEach((lbl, hi) => { rows.push({ slot:`${jk}-${hi}`, lbl: lbl||"" }); }); });
    const ok = await sb.rpc("admin_set_edt_slots", { p_ens_id: selEns, p_slots: rows });

    let cascadeOk = true;

    // Nettoyage chez les enseignants en conflit : retirer la classe transférée de leurs
    // classes assignées ET de tous leurs créneaux EDT correspondants.
    if (ok && conflits.length > 0) {
      const parAutreEns = {};
      conflits.forEach(c => { (parAutreEns[c.autreEns.id] ||= []).push(c.classe); });
      for (const [autreId, classesAEnlever] of Object.entries(parAutreEns)) {
        const autre = enseignants.find(e=>e.id===autreId);
        const nouvellesClassesAutre = (autre?.classes||[]).filter(cl=>!classesAEnlever.includes(cl));
        const rtAutre = buildEdtRuntime(data?.exceptions||{}, data?.edtBase||{})[autreId] || {};
        const slotsAEffacer = [];
        JKEYS.forEach(jk => { (rtAutre[jk]||[]).forEach((lbl,hi)=>{ if (lbl && classesAEnlever.includes(lbl)) slotsAEffacer.push(`${jk}-${hi}`); }); });
        if (slotsAEffacer.length > 0) {
          await sb.rpc("admin_set_edt_slots", { p_ens_id: autreId, p_slots: slotsAEffacer.map(slot=>({ slot, lbl:"" })) });
        }
        // Cascade identique aux autres points de retrait de classe (cf. resoudreConflit,
        // ModalEnsForm) : pas de progression/suivi digital fantôme chez le perdant du conflit.
        // Nettoyée seulement si le retrait de classe a réellement réussi (rOk).
        const rOk = await sb.rpc("admin_set_teacher_classes", { p_id: autreId, p_classes: nouvellesClassesAutre });
        if (rOk) {
          for (const cl of classesAEnlever) {
            await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: autreId, p_classe: cl });
            await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: autreId, p_classe: cl+"||dig" });
          }
        } else {
          cascadeOk = false;
        }
      }
    }

    // Cascade bidirectionnelle habituelle : les "classes assignées" de selEns deviennent
    // exactement le reflet du contenu de son EDT (ajouts ET retraits).
    // sb.patchRow (PATCH) et non sb.upsert : on ne modifie QUE la colonne "classes", sans
    // exiger les autres colonnes NOT NULL (nom, role) qu'un upsert/INSERT imposerait.
    if (ok && (classesAjoutees.length > 0 || classesRetirees.length > 0)) {
      const rOk = await sb.rpc("admin_set_teacher_classes", { p_id: selEns, p_classes: classesDansEdt });
      if (!rOk) cascadeOk = false;
      // Cascade prog_suivi pour selEns lui-même : une classe retirée de son EDT (donc de ses
      // classes assignées) ne doit pas garder de progression/suivi digital fantôme.
      // Gardée derrière "ok" : si la sauvegarde EDT elle-même a échoué, le retrait de classe
      // n'a pas réellement eu lieu, donc rien à nettoyer côté progression.
      if (rOk) {
        for (const cl of classesRetirees) {
          await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: selEns, p_classe: cl });
          await sb.rpc("admin_delete_prog_by_classe", { p_ens_id: selEns, p_classe: cl+"||dig" });
        }
      }
    }

    setSaving(false);
    const messages = [];
    if (classesAjoutees.length>0) messages.push(`${classesAjoutees.join(", ")} ajoutée(s)`);
    if (classesRetirees.length>0) messages.push(`${classesRetirees.join(", ")} retirée(s)`);
    if (conflits.length>0) messages.push(`transférée(s) depuis ${[...new Set(conflits.map(c=>getNomCourt(c.autreEns.nom)))].join(", ")}`);
    if (ok && cascadeOk) {
      showToast(messages.length>0
        ? `✓ Emploi du temps enregistré · ${messages.join(" · ")}`
        : "✓ Emploi du temps enregistré", true);
      setEditMode(false); setEdtGrid(null);
      await refreshData?.();
    } else if (ok && !cascadeOk) {
      showToast(`⚠ EDT enregistré, mais la mise à jour des classes assignées a échoué — vérifie manuellement dans "Gérer enseignants"`, false);
      setEditMode(false); setEdtGrid(null);
      await refreshData?.();
    } else {
      showToast("⚠ Échec de l'enregistrement — vérifiez la connexion", false);
    }
  };

  if (!data) return (
    <div style={{ padding:"60px", textAlign:"center", color:C.txtMuted }}>
      <Spinner size={28} color={C.green}/>
      <div style={{ marginTop:12, fontSize:13 }}>Chargement…</div>
    </div>
  );

  const ensActuel = enseignants.find(e=>e.id===selEns);

  return (
    <div style={{ padding: isMobile?14:20, display:"flex", flexDirection:"column", gap:16 }}>
      {saving && <Pill ch="🔄 Sauvegarde…" color={C.blue}/>}

      {/* Onglets */}
      <div style={{ display:"flex", gap: isMobile?5:8, background:"#f1f5f9", padding:5, borderRadius:10, overflowX:"auto" }}>
        {[
          {id:"maintenant",    label:"🟢 Maintenant"},
          {id:"parEnseignant", label:"👤 Par enseignant"},
          {id:"parClasse",     label:"🏫 Par classe"},
        ].map(o=>(
          <button key={o.id} onClick={()=>setOnglet(o.id)}
            style={{ flex:1, padding: isMobile?"9px 6px":"10px", borderRadius:7, border:"none", whiteSpace:"nowrap",
              background:onglet===o.id?C.white:"transparent", fontSize: isMobile?11.5:13, fontWeight:700,
              color:onglet===o.id?C.green:C.txtMuted, cursor:"pointer", fontFamily:"inherit",
              boxShadow:onglet===o.id?"0 1px 3px rgba(0,0,0,.1)":"none" }}>
            {o.label}
          </button>
        ))}
      </div>

      {/* ── Onglet Maintenant : qui enseigne en ce moment ── */}
      {onglet==="maintenant" && (
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          {!nowJk ? (
            <div style={{ background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:30, textAlign:"center", color:C.txtMuted }}>
              📅 Pas de cours le week-end
            </div>
          ) : enseignants.filter(ens=>edtRt[ens.id]?.[nowJk]?.[nowHi]).length===0 ? (
            <div style={{ background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:30, textAlign:"center", color:C.txtMuted }}>
              🕐 Aucun cours SVTEEHB en ce moment
            </div>
          ) : enseignants.filter(ens=>edtRt[ens.id]?.[nowJk]?.[nowHi]).map(ens=>(
            <div key={ens.id} style={{ display:"flex", alignItems:"center", gap:12, background:C.white, borderRadius:12, border:`1.5px solid ${C.greenBorder}`, padding:"12px 16px" }}>
              <Avatar ens={ens} size={36} fontSize={12}/>
              <div style={{ flex:1 }}>
                <div style={{ fontSize:13, fontWeight:700, color:C.txt }}>{getNomCourt(ens.nom)}</div>
                <div style={{ fontSize:11, color:C.txtMuted }}>{HEURES[nowHi]}</div>
              </div>
              <span style={{ padding:"6px 12px", borderRadius:20, background:C.greenPale, color:C.green, fontSize:12, fontWeight:800 }}>
                {edtRt[ens.id][nowJk][nowHi]}
              </span>
            </div>
          ))}
        </div>
      )}

      {/* ── Onglet Par enseignant ── */}
      {onglet==="parEnseignant" && (
        <>
          <div style={{ display:"flex", gap:6, overflowX:"auto", paddingBottom:4 }}>
            {enseignants.map(ens=>(
              <button key={ens.id} onClick={()=>{ setSelEns(ens.id); annulerEdition(); }}
                style={{ flexShrink:0, display:"flex", alignItems:"center", gap:7, padding:"7px 12px", borderRadius:20, cursor:"pointer", fontFamily:"inherit",
                  border:`1.5px solid ${selEns===ens.id?ens.col:C.border}`, background:selEns===ens.id?ens.col+"15":C.white }}>
                <Avatar ens={ens} size={20} fontSize={8}/>
                <span style={{ fontSize:12, fontWeight:700, color:selEns===ens.id?ens.col:C.txt, whiteSpace:"nowrap" }}>{getNomCourt(ens.nom)}</span>
              </button>
            ))}
          </div>
          {ensActuel && !editMode && (
            <>
              <EdtGrid ensId={selEns} ens={ensActuel} edtRt={edtRt} isAdmin={false}/>
              <button onClick={entrerModeEdition}
                style={{ alignSelf:"flex-start", padding:"9px 16px", background:C.greenPale, border:`1.5px solid ${C.greenBorder}`, borderRadius:9,
                  fontSize:12.5, fontWeight:700, color:C.green, cursor:"pointer", fontFamily:"inherit" }}>
                ✎ Modifier l'emploi du temps de {getNomCourt(ensActuel.nom)}
              </button>
            </>
          )}
          {ensActuel && editMode && (
            <div style={{ background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"16px 18px", display:"flex", flexDirection:"column", gap:12 }}>
              <p style={{ fontSize:11.5, color:C.txtMuted, margin:0, lineHeight:1.5 }}>
                Cette modification écrase entièrement l'EDT de {getNomCourt(ensActuel.nom)} dans Supabase — plus besoin de toucher au code pour une nouvelle année scolaire.
              </p>
              {isMobile ? (
                <>
                  <div style={{ display:"flex", gap:6, overflowX:"auto" }}>
                    {JKEYS.map((jk,i)=>(
                      <button key={jk} onClick={()=>setEdtJourMobile(jk)} type="button"
                        style={{ flexShrink:0, padding:"8px 14px", borderRadius:9, border:`1.5px solid ${edtJourMobile===jk?C.green:C.border}`, background:edtJourMobile===jk?C.greenPale:C.white, color:edtJourMobile===jk?C.greenDark:C.txtMuted, fontSize:12.5, fontWeight:700, cursor:"pointer", fontFamily:"inherit" }}>
                        {JOURS[i]}
                      </button>
                    ))}
                  </div>
                  {HEURES.map((h,hi)=>(
                    <div key={hi} style={{ display:"flex", flexDirection:"column", gap:4 }}>
                      <label style={{ fontSize:10.5, fontWeight:700, color:C.txtMuted }}>{h}</label>
                      <select value={edtGrid[edtJourMobile][hi]||""} onChange={e=>setEdtCell(edtJourMobile,hi,e.target.value)}
                        style={{ width:"100%", padding:"10px 12px", border:`1.5px solid ${C.border}`, borderRadius:9, fontSize:13, fontFamily:"inherit", background:C.white }}>
                        <option value="">— Libre —</option>
                        {toutesClasses.map(cl=><option key={cl} value={cl}>{cl}</option>)}
                      </select>
                    </div>
                  ))}
                </>
              ) : (
                <div style={{ overflowX:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11.5 }}>
                    <thead>
                      <tr>
                        <th style={{ padding:"6px 8px", textAlign:"left", color:C.txtMuted, fontWeight:700 }}>Heure</th>
                        {JOURS.map(j=><th key={j} style={{ padding:"6px 8px", textAlign:"left", color:C.txtMuted, fontWeight:700 }}>{j}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {HEURES.map((h,hi)=>(
                        <tr key={hi} style={{ borderTop:`1px solid ${C.border}` }}>
                          <td style={{ padding:"6px 8px", color:C.txtLight, fontSize:10.5, whiteSpace:"nowrap" }}>{h}</td>
                          {JKEYS.map(jk=>(
                            <td key={jk} style={{ padding:"4px 6px" }}>
                              <select value={edtGrid[jk][hi]||""} onChange={e=>setEdtCell(jk,hi,e.target.value)}
                                style={{ width:"100%", padding:"6px 8px", border:`1px solid ${C.border}`, borderRadius:7, fontSize:11, fontFamily:"inherit", background:C.white }}>
                                <option value="">—</option>
                                {toutesClasses.map(cl=><option key={cl} value={cl}>{cl}</option>)}
                              </select>
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {classesDansEdtGrid.length > 0 && (
                <div>
                  <label style={{ display:"block", fontSize:11, fontWeight:700, color:C.txtMuted, marginBottom:6 }}>Classes actuellement dans son EDT</label>
                  <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
                    {classesDansEdtGrid.map(cl=>(
                      <span key={cl} style={{ display:"inline-flex", alignItems:"center", gap:6, padding:"5px 10px", borderRadius:20, background:C.redPale, border:`1px solid ${C.redBorder}`, fontSize:11.5, fontWeight:700, color:C.red }}>
                        {cl}
                        <button onClick={()=>retirerClasseDeEdt(cl)} title={`Retirer ${cl} de tous ses créneaux`} type="button"
                          style={{ background:"none", border:"none", cursor:"pointer", color:C.red, fontSize:14, padding:0, lineHeight:1, fontWeight:900 }}>×</button>
                      </span>
                    ))}
                  </div>
                </div>
              )}

              <div style={{ display:"flex", gap:10 }}>
                <button onClick={annulerEdition} disabled={saving}
                  style={{ flex:1, padding:"10px 18px", background:C.white, border:`1px solid ${C.border}`, borderRadius:9, fontSize:13, fontWeight:700, color:C.txtMuted, cursor:"pointer", fontFamily:"inherit" }}>
                  Annuler
                </button>
                <button onClick={handleSaveEdt} disabled={saving}
                  style={{ flex:2, padding:"10px 18px", background:saving?"#94a3b8":`linear-gradient(135deg,${C.greenDark},${C.green})`,
                    border:"none", borderRadius:9, fontSize:13, fontWeight:700, color:"#fff", cursor:saving?"not-allowed":"pointer", fontFamily:"inherit" }}>
                  {saving ? "Enregistrement…" : "✓ Enregistrer l'emploi du temps"}
                </button>
              </div>
            </div>
          )}
        </>
      )}

      {/* ── Onglet Par classe ── */}
      {onglet==="parClasse" && (
        <>
          <div style={{ display:"flex", gap:6, overflowX:"auto", paddingBottom:4 }}>
            {toutesClasses.map(cl=>(
              <button key={cl} onClick={()=>setSelCl(cl)}
                style={{ flexShrink:0, padding:"7px 14px", borderRadius:20, cursor:"pointer", fontFamily:"inherit", fontSize:12, fontWeight:700,
                  border:`1.5px solid ${selCl===cl?C.green:C.border}`, background:selCl===cl?C.greenPale:C.white, color:selCl===cl?C.green:C.txt }}>
                {cl}
              </button>
            ))}
          </div>
          {selCl && (
            <div style={{ background:C.white, borderRadius:12, border:`1px solid ${C.border}`, overflow:"hidden" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead>
                  <tr style={{ background:"#fafbfc" }}>
                    <th style={{ padding:"9px 14px", textAlign:"left", fontSize:10, color:"#64748b", fontWeight:700, textTransform:"uppercase" }}>Créneau</th>
                    <th style={{ padding:"9px 14px", textAlign:"left", fontSize:10, color:"#64748b", fontWeight:700, textTransform:"uppercase" }}>Enseignant</th>
                  </tr>
                </thead>
                <tbody>
                  {enseignants.flatMap(ens=>
                    JKEYS.flatMap((jk,ji)=>HEURES.map((h,hi)=>({ens,jk,ji,hi,h})).filter(x=>edtRt[ens.id]?.[jk]?.[x.hi]===selCl))
                  ).map((x,i)=>(
                    <tr key={i} style={{ borderTop:`1px solid ${C.border}` }}>
                      <td style={{ padding:"9px 14px" }}>{JOURS[x.ji]} · {x.h}</td>
                      <td style={{ padding:"9px 14px", display:"flex", alignItems:"center", gap:8 }}>
                        <Avatar ens={x.ens} size={20} fontSize={8}/>
                        {getNomCourt(x.ens.nom)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}

    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// NAVIGATION & LAYOUT — Composants partagés
// ═══════════════════════════════════════════════════════════════════

const NAV_ADMIN = [
  {id:"dashboard",   emoji:"🏠", label:"Tableau de bord"},
  {id:"enseignants", emoji:"👥", label:"Enseignants"},
  {id:"eleves",      emoji:"🎓", label:"Élèves"},
  {id:"programme",   emoji:"📊", label:"Suivi programme",  sub:"Vue synthétique"},
  {id:"epreuves",    emoji:"📋", label:"Épreuves"},
  {id:"edt",         emoji:"📅", label:"Emploi du temps"},
  {id:"documents",   emoji:"📄", label:"Documents"},
  {id:"gestion-annuelle", emoji:"🔄", label:"Gestion annuelle", sub:"Import élèves"},
];

const NAV_TEACHER = [
  {id:"dashboard",    emoji:"🏠", label:"Tableau de bord"},
  {id:"mes-classes",  emoji:"👨‍🏫", label:"Mes classes"},
  {id:"cahier",       emoji:"📖", label:"Cahier de texte"},
  {id:"programme",    emoji:"📊", label:"Mon programme"},
  {id:"epreuves",     emoji:"📋", label:"Épreuves"},
  {id:"edt-teacher",  emoji:"📅", label:"Mon emploi du temps"},
];

const NAV_PROVISEUR = [...NAV_ADMIN, {id:"departements", emoji:"🏛️", label:"Départements", sub:"Matières · Animateurs"}];
const NAV_CENSEUR = NAV_ADMIN.filter(n=>n.id!=="enseignants" && n.id!=="gestion-annuelle");

// ── Structures groupées ──────────────────────────────────────────
const NAV_PROVISEUR_GROUPS = [
  { section:"", items:[{id:"dashboard",emoji:"🏠",label:"Tableau de bord"}] },
  { section:"ACTEURS & PÉDAGOGIE", items:[
    {id:"enseignants", emoji:"👥", label:"Enseignants"},
    {id:"eleves",      emoji:"🎓", label:"Élèves"},
    {id:"departements",emoji:"🏛️", label:"Départements & Matières", expandable:true,
      sub:[{emoji:"🌿",label:"SVT"},{emoji:"📐",label:"Mathématiques"},{emoji:"⚗️",label:"Sciences Physiques"},{emoji:"📚",label:"Lettres"},{emoji:"🌍",label:"Sciences Humaines"},{emoji:"🗣️",label:"Langues Vivantes"}]},
    {id:"documents",   emoji:"📄", label:"Documents"},
  ]},
  { section:"SUIVI & PLANNINGS", items:[
    {id:"edt",      emoji:"📅", label:"Emploi du temps"},
    {id:"epreuves", emoji:"📋", label:"Épreuves & Évaluations"},
    {id:"bulletins",emoji:"📒", label:"Bulletins de notes"},
    {id:"programme",emoji:"📊", label:"Suivi programme"},
  ]},
  { section:"CYCLE ANNUEL", items:[
    {id:"gestion-annuelle",emoji:"🔄",label:"Gestion annuelle"},
  ]},
];

const NAV_CENSEUR_GROUPS = [
  { section:"", items:[{id:"dashboard",emoji:"🏠",label:"Tableau de bord"}] },
  { section:"ACTEURS & PÉDAGOGIE", items:[
    {id:"eleves",      emoji:"🎓", label:"Élèves"},
    {id:"departements",emoji:"🏛️", label:"Départements & Matières", expandable:true,
      sub:[{emoji:"🌿",label:"SVT"},{emoji:"📐",label:"Mathématiques"},{emoji:"⚗️",label:"Sciences Physiques"},{emoji:"📚",label:"Lettres"},{emoji:"🌍",label:"Sciences Humaines"},{emoji:"🗣️",label:"Langues Vivantes"}]},
  ]},
  { section:"SUIVI & PLANNINGS", items:[
    {id:"edt",      emoji:"📅", label:"Emploi du temps"},
    {id:"epreuves", emoji:"📋", label:"Épreuves & Évaluations"},
    {id:"bulletins",emoji:"📒", label:"Bulletins de notes"},
    {id:"programme",emoji:"📊", label:"Suivi programme"},
  ]},
  { section:"DISCIPLINE & VIE SCOLAIRE", items:[
    {id:"sanctions",emoji:"⚠️", label:"Sanctions"},
    {id:"rapports", emoji:"📊", label:"Rapports disciplinaires"},
  ]},
];

const NAV_ANIMATEUR_GROUPS = [
  { section:"", items:[{id:"dashboard",emoji:"🏠",label:"Tableau de bord"}] },
  { section:"MON DÉPARTEMENT", items:[
    {id:"programme",  emoji:"📊",label:"Enseignants"},
    {id:"epreuves",   emoji:"📋",label:"Épreuves"},
    {id:"suivi-prog-dept",emoji:"📈",label:"Suivi programme"},
  ]},
  { section:"RÉUNIONS & RAPPORTS", items:[
    {id:"fiche-inspection",emoji:"🔍",label:"Fiches d'inspection"},
    {id:"documents-ap",emoji:"📄",label:"Documents"},
  ]},
  { section:"PLANNINGS", items:[
    {id:"edt-teacher",emoji:"📅",label:"Mon emploi du temps"},
  ]},
];
const NAV_SURVEILLANCE = [{id:"dashboard", emoji:"🏠", label:"Tableau de bord"}];

const PAGE_TITLES = {
  departements:"Départements",
  dashboard:"Tableau de bord", "mes-classes":"Mes classes",
  programme:"Mon programme", cahier:"Cahier de texte",
  epreuves:"Épreuves", "edt-teacher":"Mon emploi du temps",
  edt:"Emploi du temps", documents:"Documents MINESEC",
  eleves:"Élèves", enseignants:"Enseignants",
};

const KpiCard = ({label,value,sub,iconEmoji,bg,subColor}) => (
  <div style={{background:bg||C.greenPale,borderRadius:12,border:`1px solid ${C.border}`,padding:"16px 18px",flex:1,minWidth:140}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
      <span style={{fontSize:11,fontWeight:600,color:C.txtMuted,lineHeight:1.4}}>{label}</span>
      <span style={{fontSize:20}}>{iconEmoji}</span>
    </div>
    <div style={{fontSize:28,fontWeight:900,color:C.txt,lineHeight:1}}>{value}</div>
    {sub && <div style={{fontSize:11,marginTop:5,color:subColor||C.txtMuted,fontWeight:500}}>{sub}</div>}
  </div>
);

function displayCl(cl){ return CLASS_DISPLAY[cl]||cl; }

// ─── Dashboard Admin ────────────────────────────────────────────────
function DashboardAdmin() {
  const {data,user,setPage,setPendingFicheEns,refreshData} = useApp();
  const {isMobile} = useDevice();
  const [loading,setLoading] = useState(true);
  const [refreshing,setRefreshing] = useState(false);
  const [stats,setStats] = useState({nbEns:0,nbClasses:0,nbEleves:0,tauxMoyen:0,tauxParEns:[]});
  useEffect(()=>{
    if(!data)return;
    const ens=Object.values(data.users||{}).filter(u=>u.role!=="proviseur");
    const nbEleves=CLASSES_REELLES.reduce((s,c)=>s+c.effectif,0);
    const tauxParEns=ens.map(e=>{
      let tf=0,tr=0;
      (e.classes||[]).forEach(cl=>{const k=`${e.id}||${cl}`;const f=((data.prog||{})[k]||[]).length;const code=resolveProgCode(cl);const meta=code?PROG_META[code]:null;if(meta){tf+=f;tr+=meta.lpRef;}});
      return{id:e.id,nom:e.nom,photo:e.photo,col:getColor(e.id),ini:getIni(e.nom),classes:(e.classes||[]).length,taux:tr>0?Math.min(100, Math.round(tf/tr*100)):0,tf,tr};
    }).sort((a,b)=>a.taux-b.taux);
    const tauxMoyen=tauxParEns.length?Math.round(tauxParEns.reduce((s,e)=>s+e.taux,0)/tauxParEns.length):0;
    setStats({nbEns:ens.length,nbClasses:CLASSES_REELLES.length,nbEleves,tauxMoyen,tauxParEns});
    setLoading(false);
  },[data]);
  const tauCol=t=>t>=75?C.green:t>=50?C.amber:C.red;
  return(
    <div style={{padding:"20px 20px 40px",display:"flex",flexDirection:"column",gap:18}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div><h2 style={{fontSize:20,fontWeight:800,color:C.txt,margin:0}}>Bonjour, Administration 👋</h2><p style={{color:C.txtMuted,margin:"3px 0 0",fontSize:12}}>{new Date().toLocaleDateString("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</p></div>
        <div style={{textAlign:"right"}}><div style={{fontSize:11,color:C.txtMuted}}>2025–2026</div><div style={{fontSize:13,fontWeight:700,color:C.green}}>Données en direct ↗</div></div>
      </div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <KpiCard label="Enseignants" value={stats.nbEns} sub="SVTEEHB Kakatare" iconEmoji="👥" bg={C.greenPale} loading={loading} delay={0}/>
        <KpiCard label="Classes" value={stats.nbClasses} sub="Toutes séries" iconEmoji="📚" bg={C.bluePale} subColor={C.blue} loading={loading} delay={0.05}/>
        <KpiCard label="Élèves" value={stats.nbEleves} sub="Total effectifs" iconEmoji="🎓" bg={C.amberPale} subColor={C.amber} loading={loading} delay={0.1}/>
        <KpiCard label="Couverture" value={`${stats.tauxMoyen}%`} sub={stats.tauxMoyen>=75?"✓ Objectif atteint":"⚠ Sous objectif"} subColor={tauCol(stats.tauxMoyen)} iconEmoji="📊" bg={C.greenPale} loading={loading} delay={0.15}/>
      </div>
      {!loading && stats.tauxParEns.length>0 && (() => {
        const nbAlerte = stats.tauxParEns.filter(e=>e.taux<50).length;
        if (nbAlerte===0) return null;
        return (
          <div style={{display:"flex",alignItems:"center",gap:14,background:"#fef2f2",border:"1px solid #fecaca",borderRadius:12,padding:"14px 18px"}}>
            <span style={{fontSize:26,flexShrink:0}}>⚠️</span>
            <div>
              <div style={{fontWeight:700,color:"#b91c1c",fontSize:13}}>Couverture du programme en retard</div>
              <div style={{fontSize:12,color:"#7f1d1d",marginTop:2}}>
                {nbAlerte} enseignant{nbAlerte>1?"s":""} sur {stats.nbEns} {nbAlerte>1?"sont":"est"} en dessous de 50% de couverture — une relance pourrait être utile.
              </div>
            </div>
          </div>
        );
      })()}
      <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:18}}>
        <div style={{display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:10}}>
          <div>
            <h3 style={{margin:"0 0 2px",fontSize:13,fontWeight:700,color:C.txt}}>📊 Couverture programme — tous les enseignants</h3>
            <p style={{margin:0,fontSize:10.5,color:C.txtMuted}}>Triés du plus urgent au plus avancé</p>
          </div>
          <button onClick={async()=>{ setRefreshing(true); await refreshData(); setRefreshing(false); }}
            disabled={refreshing}
            style={{display:"flex", alignItems:"center", gap:6, padding: isMobile?"7px":"6px 12px", borderRadius:8, border:`1px solid ${C.border}`, background:C.white, color:C.txtMuted, fontSize:11, fontWeight:700, cursor:refreshing?"not-allowed":"pointer", fontFamily:"inherit", flexShrink:0}}>
            {refreshing ? <Spinner size={11} color={C.txtMuted}/> : "🔄"} {!isMobile && "Actualiser"}
          </button>
        </div>
        <div style={{marginTop:14}}/>
        {loading?[1,2,3,4,5].map(i=><Sk key={i} h={52} br={8} style={{marginBottom:8}}/>):(
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {stats.tauxParEns.map((e,i)=>{
              const statut = e.taux<50
                ? {label:"Alerte", bg:"#fef2f2", fg:"#b91c1c", dot:"#ef4444"}
                : e.taux>=75
                ? {label:"Objectif", bg:"#f0fdf4", fg:"#166534", dot:"#16a34a"}
                : {label:"En cours", bg:"#fffbeb", fg:"#92400e", dot:"#f59e0b"};
              return (
              <div key={e.id} onClick={()=>{ setPendingFicheEns(e.id); setPage("documents"); }}
                style={{display:"flex",alignItems:"center",gap:12,padding:"10px 12px",background:"#f8fafc",borderRadius:9,border:`1px solid ${C.border}`,borderLeft:`3px solid ${tauCol(e.taux)}`,cursor:"pointer",transition:"all .15s"}}
                onMouseEnter={ev=>{ev.currentTarget.style.background="#f1f5f9";ev.currentTarget.style.boxShadow="0 2px 8px rgba(0,0,0,.06)";}}
                onMouseLeave={ev=>{ev.currentTarget.style.background="#f8fafc";ev.currentTarget.style.boxShadow="none";}}>
                <Avatar ens={e} size={32} fontSize={10}/>
                <div style={{flex:1,minWidth:0}}><div style={{fontSize:12,fontWeight:700,color:C.txt,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{getNomCourt(e.nom)}</div><div style={{fontSize:10,color:C.txtMuted}}>{e.classes} classe{e.classes>1?"s":""} · {e.tf}/{e.tr} leçons</div></div>
                <div style={{width:130,flexShrink:0}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:10,color:C.txtMuted}}>Couverture</span><span style={{fontSize:11,fontWeight:800,color:tauCol(e.taux)}}>{e.taux}%</span></div>
                  <ProgBar value={e.taux}/>
                </div>
                <span style={{display:"inline-flex", alignItems:"center", gap:5, padding:"4px 10px", borderRadius:20,
                  background:statut.bg, color:statut.fg, fontSize:10, fontWeight:700, flexShrink:0}}>
                  <span style={{width:5,height:5,borderRadius:"50%",background:statut.dot,flexShrink:0}}/>
                  {statut.label}
                </span>
                <span style={{fontSize:13,color:"#cbd5e1",flexShrink:0}}>›</span>
              </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Dashboard Enseignant ─────────────────────────────────────────

// ─── Dashboard Proviseur (vue tous départements) ───────────────────

function DashboardProviseur() {
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
      (e.classes||[]).forEach(cl=>{const k=`${e.id}||${cl}`;const f=((data.prog||{})[k]||[]).length;const code=resolveProgCode(cl);const meta=code?PROG_META[code]:null;if(meta){tf+=f;tr+=meta.lpRef;}});
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
          const prog=(data.prog||{})[`${e.id}||${cl}`]||[];
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
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
        <g transform={`rotate(-90 ${size/2} ${size/2})`}>
          {segments.map((s,i)=>{
            const frac = s.value/total;
            const dash = frac*circ;
            const el = (
              <circle key={i} cx={size/2} cy={size/2} r={r} fill="none" stroke={s.color} strokeWidth={thickness}
                strokeDasharray={`${dash} ${circ-dash}`} strokeDashoffset={-acc}/>
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
        <KpiCard label="Couverture" value={`${stats.tauxMoyen}%`} sub={stats.tauxMoyen>=75?"✓ Objectif atteint":"⚠ Sous objectif"} subColor={tauCol(stats.tauxMoyen)} iconEmoji="📊" bg={C.greenPale} loading={loading} delay={0.15}/>
      </div>

      <div style={{display:"grid", gridTemplateColumns: isMobile?"1fr":"1fr 1fr 1fr", gap:14}}>
        <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:16}}>
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

        <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:16}}>
          <h3 style={{margin:"0 0 12px",fontSize:12.5,fontWeight:700,color:C.txt}}>📈 Évolution de la couverture</h3>
          {loading ? <Sk h={130} br={8}/> : <EvolutionChartLarge series={stats.evolution} height={130}/>}
        </div>

        <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:16}}>
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

      <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:18}}>
        <div style={{display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:10}}>
          <div>
            <h3 style={{margin:"0 0 2px",fontSize:13,fontWeight:700,color:C.txt}}>🏛️ Vue par département</h3>
            <p style={{margin:0,fontSize:10.5,color:C.txtMuted}}>Couverture moyenne et effectifs enseignants</p>
          </div>
          <button onClick={async()=>{ setRefreshing(true); await refreshData(); setRefreshing(false); }}
            disabled={refreshing}
            style={{display:"flex", alignItems:"center", gap:6, padding: isMobile?"7px":"6px 12px", borderRadius:8, border:`1px solid ${C.border}`, background:C.white, color:C.txtMuted, fontSize:11, fontWeight:700, cursor:refreshing?"not-allowed":"pointer", fontFamily:"inherit", flexShrink:0}}>
            {refreshing ? <Spinner size={11} color={C.txtMuted}/> : "🔄"} {!isMobile && "Actualiser"}
          </button>
        </div>
        <div style={{marginTop:14, display:"grid", gridTemplateColumns: isMobile?"1fr":"repeat(auto-fill, minmax(220px, 1fr))", gap:10}}>
          {loading?[1,2,3,4].map(i=><Sk key={i} h={78} br={10}/>):(
            stats.parDept.map(d=>(
              <div key={d.id} style={{padding:"14px",background: d.nbEns>0 ? "#f8fafc":"#fafafa",borderRadius:10,border:`1px solid ${C.border}`,borderLeft:`3px solid ${d.nbEns>0?tauCol(d.taux):C.border}`}}>
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
        <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:18}}>
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

function genererBilanTrimestriel(stats, vieSco, sgClasses, niveauLabel, data) {
  const trimestre = prompt("Numéro du trimestre (1, 2 ou 3) :", "1");
  if (!trimestre) return;
  const annee = "2025-2026";

  const eleveMap = {};
  stats.parEleve.forEach(e => {
    eleveMap[e.id] = {...e, retards:0, sanctions:0, incidents:0};
  });
  vieSco.forEach(v => {
    if (!eleveMap[v.eleve_id]) eleveMap[v.eleve_id] = {id:v.eleve_id,classe:v.classe,nom:(ELEVES_DB[v.classe]||[]).find(x=>x.id===v.eleve_id)?.nom||v.eleve_id,count:0,heures:0,retards:0,sanctions:0,incidents:0};
    if(v.type==="retard") eleveMap[v.eleve_id].retards++;
    if(v.type==="sanction") eleveMap[v.eleve_id].sanctions++;
    if(v.type==="incident") eleveMap[v.eleve_id].incidents++;
  });

  const classes = sgClasses || CLASSES_REELLES.map(c=>c.code);
  let rows = "";
  classes.forEach(cl => {
    const eleves = Object.values(eleveMap).filter(e=>e.classe===cl).sort((a,b)=>a.nom.localeCompare(b.nom));
    if (eleves.length === 0) return;
    rows += '<tr style="background:#0B4D2C;color:#fff"><td colspan="6" style="padding:8px 12px;font-weight:700;font-size:14px;">'+cl+'</td></tr>';
    eleves.forEach((e,i) => {
      rows += '<tr style="background:'+(i%2===0?"#fff":"#f8fafc")+'">'
        + '<td style="padding:7px 12px;border-bottom:1px solid #e2e8f0">'+(i+1)+'</td>'
        + '<td style="padding:7px 12px;border-bottom:1px solid #e2e8f0;font-weight:600">'+e.nom+'</td>'
        + '<td style="padding:7px 12px;border-bottom:1px solid #e2e8f0;text-align:center;color:'+(e.count>=3?"#b91c1c":"#374151")+'">'+e.count+' séances</td>'
        + '<td style="padding:7px 12px;border-bottom:1px solid #e2e8f0;text-align:center;font-weight:700;color:'+(e.heures>=6?"#b91c1c":e.heures>=3?"#d97706":"#374151")+'">'+( e.heures||e.count)+'h</td>'
        + '<td style="padding:7px 12px;border-bottom:1px solid #e2e8f0;text-align:center">'+(e.retards>0?'⏰ '+e.retards:'—')+'</td>'
        + '<td style="padding:7px 12px;border-bottom:1px solid #e2e8f0;text-align:center">'+(e.sanctions>0?'⚠️ '+e.sanctions:e.incidents>0?'🚨 '+e.incidents:'—')+'</td>'
        + '</tr>';
    });
  });

  const html = '<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">'
    + '<title>Bilan Trimestriel T'+trimestre+' — '+niveauLabel+'</title>'
    + '<style>body{font-family:sans-serif;padding:24px;color:#1f2937}h1{color:#0B4D2C}table{width:100%;border-collapse:collapse}th{background:#0B4D2C;color:#fff;padding:10px 12px;text-align:left;font-size:12px;letter-spacing:.05em}@media print{.no-print{display:none}}</style>'
    + '</head><body>'
    + '<div class="no-print" style="margin-bottom:16px"><button onclick="window.print()" style="padding:10px 20px;background:#0B4D2C;color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:14px;font-weight:700">🖨️ Imprimer / Exporter PDF</button></div>'
    + '<h1>Bilan Trimestriel — Trimestre '+trimestre+'</h1>'
    + '<p style="color:#6b7280">Année scolaire '+annee+' · Niveau : '+niveauLabel+' · Généré le '+new Date().toLocaleDateString("fr-FR")+'</p>'
    + '<table><thead><tr><th>#</th><th>Élève</th><th>Séances abs.</th><th>Heures abs.</th><th>Retards</th><th>Sanctions/Incidents</th></tr></thead>'
    + '<tbody>'+rows+'</tbody></table>'
    + '</body></html>';

  const w = window.open("","_blank");
  w.document.write(html);
  w.document.close();
}

function DashboardSurveillance() {
  const {rawData:data, user} = useApp();
  const {isMobile} = useDevice();

  const sgClasses = user?.classes?.length > 0 ? user.classes : null;
  const niveauLabel = sgClasses
    ? (sgClasses[0]?.startsWith('6')?'6ème':sgClasses[0]?.startsWith('5')?'5ème':sgClasses[0]?.startsWith('4')?'4ème':sgClasses[0]?.startsWith('3')?'3ème':sgClasses[0]?.startsWith('2')?'2nde':'1ère & Tle')
    : "Toute l'école";

  const [loading,setLoading]     = useState(true);
  const [tab,setTab]             = useState(()=>{ const t=window.__sgTab; window.__sgTab=null; return t||"vue"; });
  useEffect(()=>{
    const handler = (e) => { setTab(e.detail); window.__sgTab = null; };
    window.addEventListener("sg:tab", handler);
    return () => window.removeEventListener("sg:tab", handler);
  }, []);
  const [vieSco,setVieSco]       = useState([]);
  const [vieLoading,setVieLoading] = useState(true);
  const [ficheEleve,setFicheEleve] = useState(null);
  const [stats,setStats]         = useState({total:0,parEleve:[],parDept:[],nbAlerte:0,hebdo:[],scoreDisc:[]});
  const [showForm,setShowForm]   = useState(false);
  const [saving,setSaving]       = useState(false);
  const [form,setForm]           = useState({eleve_id:"",classe:"",motif:"",details:"",gravite:"faible"});
  const [selClasse,setSelClasse] = useState("");
  const [elevesSel,setElevesSel] = useState([]);
  const [formErr,setFormErr]     = useState("");
  const [selTrim,setSelTrim]     = useState(1); // bilan trimestriel

  const getWeekKey = (d) => {
    const dt = new Date(d); if(isNaN(dt)) return d;
    const day = dt.getDay()||7;
    const lun = new Date(dt); lun.setDate(dt.getDate()-day+1);
    return lun.toISOString().slice(0,10);
  };
  const weekLabel = (wk) => {
    const d = new Date(wk), fin = new Date(wk);
    fin.setDate(fin.getDate()+6);
    return 'Sem. '+d.toLocaleDateString('fr-FR',{day:'2-digit',month:'short'})+' – '+fin.toLocaleDateString('fr-FR',{day:'2-digit',month:'short'});
  };

  useEffect(()=>{
    if(!data) return;
    const parEleveMap={}, hebdoMap={};
    let total=0;
    Object.entries(data.absences||{}).forEach(([k,absents])=>{
      const [,classe,seance]=k.split("||");
      if(sgClasses && !sgClasses.includes(classe)) return;
      (absents||[]).forEach(id=>{
        total++;
        const wk=getWeekKey(seance);
        if(!parEleveMap[id]) parEleveMap[id]={id,classe,count:0,dernier:seance,semaines:{}};
        parEleveMap[id].count++;
        parEleveMap[id].semaines[wk]=(parEleveMap[id].semaines[wk]||0)+1;
        if(seance>parEleveMap[id].dernier) parEleveMap[id].dernier=seance;
        if(!hebdoMap[wk]) hebdoMap[wk]={wk,total:0,nbEleves:new Set()};
        hebdoMap[wk].total++; hebdoMap[wk].nbEleves.add(id);
      });
    });
    const parEleve=Object.values(parEleveMap)
      .map(e=>({...e,nom:(ELEVES_DB[e.classe]||[]).find(x=>x.id===e.id)?.nom||e.id}))
      .sort((a,b)=>b.count-a.count).slice(0,30);
    const nbAlerte=Object.values(parEleveMap).filter(e=>e.count>=3).length;
    const hebdo=Object.values(hebdoMap)
      .map(w=>({...w,nbEleves:w.nbEleves.size}))
      .sort((a,b)=>b.wk.localeCompare(a.wk)).slice(0,8);
    const deptOf={}; Object.values(data.users||{}).forEach(u=>{deptOf[u.id]=u.departement_id||1;});
    const absParDept={};
    if(!sgClasses){ Object.entries(data.absences||{}).forEach(([k,abs])=>{ const d=deptOf[k.split("||")[0]]||1; absParDept[d]=(absParDept[d]||0)+(abs?abs.length:0); }); }
    const parDept=DEPARTEMENTS_LIST.map(d=>({...d,total:absParDept[d.id]||0})).sort((a,b)=>b.total-a.total);
    setStats({total,parEleve,parDept,nbAlerte,hebdo,scoreDisc:parEleve});
    setLoading(false);
  },[data]);

  const loadVieSco = async()=>{
    setVieLoading(true);
    const rows = await sb.get("vie_scolaire","?select=*&order=date.desc,created_at.desc&limit=500");
    const filtered = sgClasses ? (rows||[]).filter(v=>sgClasses.includes(v.classe)) : (rows||[]);
    setVieSco(filtered);
    setVieLoading(false);
  };
  useEffect(()=>{ loadVieSco(); },[]);
  useEffect(()=>{
    if(!selClasse){setElevesSel([]);setForm(f=>({...f,eleve_id:"",classe:""}));return;}
    setElevesSel(ELEVES_DB[selClasse]||[]);
    setForm(f=>({...f,classe:selClasse,eleve_id:""}));
  },[selClasse]);

  const typeMap={retards:"retard",sanctions:"sanction",incidents:"incident"};

  const saveEntry = async()=>{
    if(!form.eleve_id||!form.classe){setFormErr("Sélectionnez une classe et un élève.");return;}
    setFormErr("");setSaving(true);
    const payload={type:typeMap[tab]||"retard",eleve_id:form.eleve_id,classe:form.classe,
          motif:form.motif||null,details:form.details||null,
          gravite:(tab==="retards")?"faible":form.gravite||"faible",enregistre_par:user?.id||"sg"};
    const ok=await sb.rpc("submit_vie_scolaire",{p_type:payload.type,p_eleve_id:payload.eleve_id,p_classe:payload.classe,p_motif:payload.motif,p_details:payload.details,p_gravite:payload.gravite});
    if(ok){await loadVieSco();setShowForm(false);}
    else setFormErr("Erreur d'enregistrement.");
    setSaving(false);
  };

  // Score discipline = absences + retards*0.5 + sanctions*2 + incidents*3
  const scoreParEleve = (() => {
    const map={};
    stats.parEleve.forEach(e=>{ map[e.id]={...e,retards:0,sanctions:0,incidents:0}; });
    vieSco.forEach(v=>{
      if(!map[v.eleve_id]) map[v.eleve_id]={id:v.eleve_id,classe:v.classe,nom:(ELEVES_DB[v.classe]||[]).find(x=>x.id===v.eleve_id)?.nom||v.eleve_id,count:0,retards:0,sanctions:0,incidents:0};
      if(v.type==="retard") map[v.eleve_id].retards++;
      if(v.type==="sanction") map[v.eleve_id].sanctions++;
      if(v.type==="incident") map[v.eleve_id].incidents++;
    });
    return Object.values(map).map(e=>({...e,score:e.count+e.retards*0.5+e.sanctions*2+e.incidents*3}))
      .sort((a,b)=>b.score-a.score).slice(0,20);
  })();

  const filteredVie = vieSco.filter(v=>v.type===(typeMap[tab]||"retard"));
  const classesSG = sgClasses || CLASSES_REELLES.map(c=>c.code);

  const TABS=[
    {id:"vue",       label:"Vue d'ensemble",emoji:"🎯"},
    {id:"absences",  label:"Absences",      emoji:"📋"},
    {id:"hebdo",     label:"Par semaine",   emoji:"📅"},
    {id:"bilan_trim", label:"Bilan trimestriel", emoji:"📊"},
    {id:"retards",   label:"Retards",       emoji:"⏰"},
    {id:"sanctions", label:"Sanctions",     emoji:"⚠️"},
    {id:"incidents", label:"Incidents",     emoji:"🚨"},
  ];

  const GravBadge=({g})=>{
    const c={faible:{bg:"#fefce8",fg:"#854d0e"},moyen:{bg:"#fff7ed",fg:"#c2410c"},grave:{bg:"#fef2f2",fg:"#b91c1c"}}[g]||{bg:"#fefce8",fg:"#854d0e"};
    return React.createElement("span",{style:{fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:10,background:c.bg,color:c.fg}},g||"faible");
  };
  const NomEleve=({eleveId,classe})=>React.createElement("span",null,(ELEVES_DB[classe]||[]).find(x=>x.id===eleveId)?.nom||eleveId);

  // Fiche élève (panel latéral)
  const FichePanel = ficheEleve ? (()=>{
    const absEleve = stats.parEleve.find(e=>e.id===ficheEleve.id);
    const vieScoEleve = vieSco.filter(v=>v.eleve_id===ficheEleve.id);
    return(
      <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.45)",zIndex:200,display:"flex",alignItems:"flex-end",justifyContent:"flex-end"}}
        onClick={e=>{if(e.target===e.currentTarget)setFicheEleve(null);}}>
        <div style={{width:isMobile?"100%":"420px",height:isMobile?"85vh":"100vh",background:C.white,borderRadius:isMobile?"16px 16px 0 0":0,overflow:"auto",padding:22,display:"flex",flexDirection:"column",gap:16}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div>
              <h3 style={{fontSize:16,fontWeight:800,color:C.txt,margin:0}}>{ficheEleve.nom}</h3>
              <p style={{fontSize:11,color:C.txtMuted,margin:"2px 0 0"}}>{ficheEleve.classe}</p>
            </div>
            <button onClick={()=>setFicheEleve(null)} style={{background:"none",border:"none",fontSize:20,cursor:"pointer",color:C.txtMuted}}>✕</button>
          </div>
          <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
            {[
              {label:"Absences",val:absEleve?.count||0,col:C.red},
              {label:"Retards", val:vieScoEleve.filter(v=>v.type==="retard").length, col:C.amber},
              {label:"Sanctions",val:vieScoEleve.filter(v=>v.type==="sanction").length, col:C.red},
              {label:"Incidents",val:vieScoEleve.filter(v=>v.type==="incident").length, col:"#7c3aed"},
            ].map(({label,val,col})=>(
              <div key={label} style={{flex:1,minWidth:70,background:C.bg,borderRadius:8,padding:"10px 12px",textAlign:"center"}}>
                <div style={{fontSize:20,fontWeight:800,color:col}}>{val}</div>
                <div style={{fontSize:10,color:C.txtMuted}}>{label}</div>
              </div>
            ))}
          </div>
          <div>
            <h4 style={{fontSize:12,fontWeight:700,color:C.txt,marginBottom:10}}>Historique vie scolaire</h4>
            {vieScoEleve.length===0
              ? <p style={{fontSize:12,color:C.txtLight,textAlign:"center",padding:"20px 0"}}>Aucun incident enregistré</p>
              : vieScoEleve.map(v=>(
                <div key={v.id} style={{display:"flex",gap:10,padding:"8px 0",borderBottom:"1px solid "+C.border,alignItems:"flex-start"}}>
                  <span style={{fontSize:11,color:C.txtMuted,whiteSpace:"nowrap",width:55}}>{new Date(v.date).toLocaleDateString("fr-FR",{day:"2-digit",month:"short"})}</span>
                  <div style={{flex:1}}>
                    <span style={{fontSize:11,fontWeight:700,color:v.type==="incident"?"#7c3aed":v.type==="sanction"?C.red:C.amber,textTransform:"capitalize"}}>{v.type}</span>
                    {v.motif && <span style={{fontSize:11,color:C.txt}}> · {v.motif}</span>}
                    {v.gravite && v.gravite!=="faible" && <GravBadge g={v.gravite}/>}
                  </div>
                </div>
              ))
            }
          </div>
        </div>
      </div>
    );
  })() : null;

  return(
    <div style={{padding:"20px 20px 40px",display:"flex",flexDirection:"column",gap:18}}>
      {FichePanel}

      {/* En-tête */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:10}}>
        <div>
          <h2 style={{fontSize:20,fontWeight:800,color:C.txt,margin:0}}>Surveillance générale 🛡️</h2>
          <p style={{color:C.txtMuted,margin:"3px 0 0",fontSize:12}}>
            {new Date().toLocaleDateString("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}
            {" · "}<strong style={{color:C.green}}>{niveauLabel}</strong>
          </p>
        </div>
        <button onClick={()=>genererBilanTrimestriel(stats,vieSco,sgClasses,niveauLabel,data)}
          style={{padding:"9px 16px",borderRadius:10,border:"1px solid "+C.border,background:C.white,color:C.txt,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:6}}>
          📄 Bilan trimestriel
        </button>
        {(tab==="retards"||tab==="sanctions"||tab==="incidents") && (
          <button onClick={()=>{setFormErr("");setForm({eleve_id:"",classe:"",motif:"",details:"",gravite:"faible"});setSelClasse("");setShowForm(!showForm);}}
            style={{padding:"9px 18px",borderRadius:10,border:"none",background:showForm?C.border:C.green,color:showForm?C.txt:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>
            {showForm?"✕ Annuler":"➕ Enregistrer"}
          </button>
        )}
      </div>

      {/* KPI */}
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <KpiCard label="Heures d'absence" value={(stats.totalHeures||stats.total)+"h"} sub={niveauLabel+" · "+stats.total+" séances"} iconEmoji="📋" bg={C.bluePale} subColor={C.blue} loading={loading} delay={0}/>
        <KpiCard label="Élèves en alerte" value={stats.nbAlerte} sub="3+ absences" iconEmoji="⚠️" bg={C.redPale} subColor={C.red} loading={loading} delay={0.05}/>
        <KpiCard label="Retards" value={vieSco.filter(v=>v.type==="retard").length} sub="Enregistrés" iconEmoji="⏰" bg={C.amberPale} subColor={C.amber} loading={vieLoading} delay={0.1}/>
        <KpiCard label="Sanctions" value={vieSco.filter(v=>v.type==="sanction").length} sub="Enregistrées" iconEmoji="⚠️" bg={C.redPale} subColor={C.red} loading={vieLoading} delay={0.15}/>
      </div>

      {/* Onglets */}
      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>{setTab(t.id);setShowForm(false);}}
            style={{padding:"8px 14px",borderRadius:20,border:"1.5px solid "+(tab===t.id?C.green:C.border),
              background:tab===t.id?C.greenPale:C.white,color:tab===t.id?C.green:C.txtMuted,
              fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit",transition:"all .15s",whiteSpace:"nowrap"}}>
            {t.emoji} {t.label}
          </button>
        ))}
      </div>

      {/* Formulaire */}
      {showForm && (
        <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:18,display:"flex",flexDirection:"column",gap:12}}>
          <h3 style={{margin:0,fontSize:13,fontWeight:700,color:C.txt}}>
            {tab==="retards"?"⏰ Retard":tab==="sanctions"?"⚠️ Sanction":"🚨 Incident"}
          </h3>
          {formErr&&<div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:8,padding:"8px 12px",fontSize:12.5,color:"#b91c1c"}}>{formErr}</div>}
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:10}}>
            <div>
              <label style={{fontSize:11,fontWeight:600,color:C.txtMuted,display:"block",marginBottom:4}}>Classe *</label>
              <select value={selClasse} onChange={e=>setSelClasse(e.target.value)}
                style={{width:"100%",padding:"9px 12px",border:"1.5px solid "+C.border,borderRadius:8,fontSize:13,fontFamily:"inherit",background:"#f8fafc"}}>
                <option value="">— Sélectionner —</option>
                {classesSG.map(c=><option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            <div>
              <label style={{fontSize:11,fontWeight:600,color:C.txtMuted,display:"block",marginBottom:4}}>Élève *</label>
              <select value={form.eleve_id} onChange={e=>setForm(f=>({...f,eleve_id:e.target.value}))}
                disabled={!selClasse}
                style={{width:"100%",padding:"9px 12px",border:"1.5px solid "+C.border,borderRadius:8,fontSize:13,fontFamily:"inherit",background:selClasse?"#f8fafc":"#f1f5f9",opacity:selClasse?1:.6}}>
                <option value="">— Sélectionner —</option>
                {elevesSel.map(e=><option key={e.id} value={e.id}>{e.nom}</option>)}
              </select>
            </div>
            <div>
              <label style={{fontSize:11,fontWeight:600,color:C.txtMuted,display:"block",marginBottom:4}}>Motif</label>
              <input type="text" value={form.motif} onChange={e=>setForm(f=>({...f,motif:e.target.value}))}
                placeholder="Motif..." style={{width:"100%",padding:"9px 12px",border:"1.5px solid "+C.border,borderRadius:8,fontSize:13,fontFamily:"inherit",background:"#f8fafc",boxSizing:"border-box"}}/>
            </div>
            {tab!=="retards"&&(
              <div>
                <label style={{fontSize:11,fontWeight:600,color:C.txtMuted,display:"block",marginBottom:4}}>Gravité</label>
                <select value={form.gravite} onChange={e=>setForm(f=>({...f,gravite:e.target.value}))}
                  style={{width:"100%",padding:"9px 12px",border:"1.5px solid "+C.border,borderRadius:8,fontSize:13,fontFamily:"inherit",background:"#f8fafc"}}>
                  <option value="faible">Faible</option>
                  <option value="moyen">Moyen</option>
                  <option value="grave">Grave</option>
                </select>
              </div>
            )}
            <div style={{gridColumn:"1 / -1"}}>
              <label style={{fontSize:11,fontWeight:600,color:C.txtMuted,display:"block",marginBottom:4}}>Détails</label>
              <textarea value={form.details} onChange={e=>setForm(f=>({...f,details:e.target.value}))}
                placeholder="Détails..." style={{width:"100%",padding:"9px 12px",border:"1.5px solid "+C.border,borderRadius:8,fontSize:13,fontFamily:"inherit",background:"#f8fafc",resize:"vertical",minHeight:56,boxSizing:"border-box"}}/>
            </div>
          </div>
          <button onClick={saveEntry} disabled={saving}
            style={{alignSelf:"flex-end",padding:"10px 24px",borderRadius:10,border:"none",background:saving?"#94a3b8":C.green,color:"#fff",fontWeight:700,fontSize:13,cursor:saving?"not-allowed":"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:8}}>
            {saving?<><Spinner size={12} color="#fff"/> Enregistrement...</>:"✓ Enregistrer"}
          </button>
        </div>
      )}

      {/* Vue d'ensemble */}
      {tab==="vue" && (
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1.4fr 1fr",gap:14}}>
          <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:16}}>
            <h3 style={{margin:"0 0 4px",fontSize:12.5,fontWeight:700,color:C.txt}}>🎯 Score discipline — élèves à surveiller</h3>
            <p style={{margin:"0 0 12px",fontSize:10,color:C.txtMuted}}>Abs + retards + sanctions + incidents — cliquer pour la fiche</p>
            {(loading||vieLoading)?<Sk h={200} br={8}/>:scoreParEleve.length>0?(
              <div style={{display:"flex",flexDirection:"column",gap:0,maxHeight:420,overflowY:"auto"}}>
                {scoreParEleve.map((e,i)=>(
                  <div key={e.id} onClick={()=>setFicheEleve(e)}
                    style={{display:"flex",alignItems:"center",gap:10,padding:"9px 6px",borderBottom:i<scoreParEleve.length-1?"1px solid "+C.border:"none",cursor:"pointer",borderRadius:6,transition:"background .1s"}}
                    onMouseEnter={ev=>ev.currentTarget.style.background=C.greenPaleA60||"#f0fdf4"}
                    onMouseLeave={ev=>ev.currentTarget.style.background="transparent"}>
                    <span style={{fontSize:11,color:C.txtMuted,width:22,flexShrink:0,fontWeight:600}}>{i+1}</span>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:12.5,fontWeight:700,color:C.txt,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{e.nom}</div>
                      <div style={{fontSize:10,color:C.txtMuted,marginTop:1,display:"flex",gap:8}}>
                        <span>{e.count} abs</span>
                        {e.retards>0&&<span style={{color:C.amber}}>{e.retards} ret.</span>}
                        {e.sanctions>0&&<span style={{color:C.red}}>{e.sanctions} sanc.</span>}
                        {e.incidents>0&&<span style={{color:"#7c3aed"}}>{e.incidents} inc.</span>}
                      </div>
                    </div>
                    <div style={{textAlign:"right",flexShrink:0}}>
                      <div style={{fontSize:13,fontWeight:800,color:e.score>=5?C.red:e.score>=2?C.amber:C.green}}>{e.score.toFixed(1)}</div>
                      <div style={{fontSize:9,color:C.txtMuted}}>score</div>
                    </div>
                    <span style={{fontSize:12,color:C.txtMuted}}>›</span>
                  </div>
                ))}
              </div>
            ):<div style={{fontSize:11,color:C.txtLight,textAlign:"center",padding:"30px 0"}}>Aucun incident enregistré</div>}
          </div>

          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            {/* Stats discipline */}
            <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:16}}>
              <h3 style={{margin:"0 0 12px",fontSize:12.5,fontWeight:700,color:C.txt}}>📊 Répartition vie scolaire</h3>
              {vieLoading?<Sk h={100} br={8}/>:(()=>{
                const tot=vieSco.length||1;
                const nb={retard:vieSco.filter(v=>v.type==="retard").length,sanction:vieSco.filter(v=>v.type==="sanction").length,incident:vieSco.filter(v=>v.type==="incident").length};
                return tot===1&&vieSco.length===0
                  ?<div style={{fontSize:11,color:C.txtLight,textAlign:"center",padding:"20px 0"}}>Aucune donnée</div>
                  :<div style={{display:"flex",flexDirection:"column",gap:10}}>
                    {[{label:"Retards",val:nb.retard,col:C.amber},{label:"Sanctions",val:nb.sanction,col:C.red},{label:"Incidents",val:nb.incident,col:"#7c3aed"}].map(({label,val,col})=>(
                      <div key={label}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                          <span style={{fontSize:12,color:C.txt,fontWeight:600}}>{label}</span>
                          <span style={{fontSize:12,fontWeight:800,color:col}}>{val}</span>
                        </div>
                        <div style={{height:6,borderRadius:3,background:"#e2e8f0"}}>
                          <div style={{height:"100%",borderRadius:3,background:col,width:Math.round(val/Math.max(vieSco.length,1)*100)+"%",transition:"width .4s"}}/>
                        </div>
                      </div>
                    ))}
                  </div>;
              })()}
            </div>

            {/* Absences par classe */}
            <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:16}}>
              <h3 style={{margin:"0 0 12px",fontSize:12.5,fontWeight:700,color:C.txt}}>📋 Absences par classe</h3>
              {loading?<Sk h={100} br={8}/>:(()=>{
                const absMap={};
                Object.entries(data?.absences||{}).forEach(([k,abs])=>{
                  const [,cl]=k.split("||");
                  if(sgClasses&&!sgClasses.includes(cl))return;
                  absMap[cl]=(absMap[cl]||0)+(abs?abs.length:0);
                });
                const rows=Object.entries(absMap).sort((a,b)=>b[1]-a[1]).slice(0,8);
                return rows.length>0
                  ?<div style={{display:"flex",flexDirection:"column",gap:7}}>
                    {rows.map(([cl,n])=>(
                      <div key={cl} style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontSize:11,color:C.txt,flex:1,fontWeight:n>5?700:400,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{cl}</span>
                        <span style={{fontSize:11,fontWeight:800,color:n>10?C.red:n>5?C.amber:C.green,flexShrink:0}}>{n}</span>
                      </div>
                    ))}
                  </div>
                  :<div style={{fontSize:11,color:C.txtLight,textAlign:"center",padding:"16px 0"}}>Aucune absence</div>;
              })()}
            </div>
          </div>
        </div>
      )}

      {/* Absences */}
      {tab==="absences" && (
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1.4fr 1fr",gap:14}}>
          <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:16}}>
            <h3 style={{margin:"0 0 4px",fontSize:12.5,fontWeight:700,color:C.txt}}>🔍 Élèves les plus absents</h3>
            <p style={{margin:"0 0 12px",fontSize:10,color:C.txtMuted}}>Top 30 — cliquer pour la fiche</p>
            {loading?<Sk h={200} br={8}/>:stats.parEleve.length>0?(
              <div style={{display:"flex",flexDirection:"column",gap:0,maxHeight:440,overflowY:"auto"}}>
                {stats.parEleve.map((e,i)=>(
                  <div key={e.id} onClick={()=>setFicheEleve(e)}
                    style={{display:"flex",alignItems:"center",gap:10,padding:"8px 6px",borderBottom:i<stats.parEleve.length-1?"1px solid "+C.border:"none",cursor:"pointer",borderRadius:6}}
                    onMouseEnter={ev=>ev.currentTarget.style.background="#f0fdf4"}
                    onMouseLeave={ev=>ev.currentTarget.style.background="transparent"}>
                    <span style={{fontSize:11,color:C.txtMuted,width:22,flexShrink:0}}>{i+1}</span>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:12.5,fontWeight:700,color:C.txt,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{e.nom}</div>
                      <div style={{fontSize:10,color:C.txtMuted}}>{e.classe}</div>
                    </div>
                    <span style={{fontSize:12,fontWeight:800,color:e.count>=3?C.red:C.amber,flexShrink:0}}>{e.count} abs.</span>
                    <span style={{fontSize:12,color:C.txtMuted}}>›</span>
                  </div>
                ))}
              </div>
            ):<div style={{fontSize:11,color:C.txtLight,textAlign:"center",padding:"30px 0"}}>Aucune absence</div>}
          </div>
          <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:16}}>
            <h3 style={{margin:"0 0 12px",fontSize:12.5,fontWeight:700,color:C.txt}}>{sgClasses?"📚 Par classe":"🏛️ Par département"}</h3>
            {loading?<Sk h={150} br={8}/>:(()=>{
              const map={};
              Object.entries(data?.absences||{}).forEach(([k,abs])=>{
                const [,cl]=k.split("||");
                if(sgClasses&&!sgClasses.includes(cl))return;
                map[cl]=(map[cl]||0)+(abs?abs.length:0);
              });
              const rows=Object.entries(map).sort((a,b)=>b[1]-a[1]).slice(0,10);
              return rows.some(r=>r[1]>0)
                ?<div style={{display:"flex",flexDirection:"column",gap:8}}>
                  {rows.map(([cl,n])=>(
                    <div key={cl} style={{display:"flex",alignItems:"center",gap:8}}>
                      <span style={{fontSize:11,color:C.txt,flex:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{cl}</span>
                      <span style={{fontSize:11,fontWeight:800,color:n>0?C.red:C.txtMuted}}>{n}</span>
                    </div>
                  ))}
                </div>
                :<div style={{fontSize:11,color:C.txtLight,textAlign:"center",padding:"30px 0"}}>Aucune absence</div>;
            })()}
          </div>
        </div>
      )}

      {/* Par semaine */}
      {tab==="hebdo" && (
        <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:16}}>
          <h3 style={{margin:"0 0 4px",fontSize:12.5,fontWeight:700,color:C.txt}}>📅 Absences par semaine</h3>
          <p style={{margin:"0 0 16px",fontSize:10,color:C.txtMuted}}>8 dernières semaines · {niveauLabel}</p>
          {loading?<Sk h={200} br={8}/>:stats.hebdo.length>0?(
            <div style={{display:"flex",flexDirection:"column",gap:0}}>
              <div style={{display:"grid",gridTemplateColumns:"1fr auto auto",gap:12,padding:"8px 12px",background:"#f8fafc",borderRadius:"8px 8px 0 0",borderBottom:"1px solid "+C.border}}>
                <span style={{fontSize:10,fontWeight:700,color:C.txtMuted}}>SEMAINE</span>
                <span style={{fontSize:10,fontWeight:700,color:C.txtMuted,textAlign:"center"}}>ÉLÈVES</span>
                <span style={{fontSize:10,fontWeight:700,color:C.txtMuted,textAlign:"right"}}>ABSENCES</span>
              </div>
              {stats.hebdo.map((w,i)=>{
                const pct=stats.hebdo[0]?.total>0?Math.round(w.total/stats.hebdo[0].total*100):0;
                return(
                  <div key={w.wk} style={{display:"grid",gridTemplateColumns:"1fr auto auto",gap:12,padding:"11px 12px",borderBottom:i<stats.hebdo.length-1?"1px solid "+C.border:"none",alignItems:"center"}}>
                    <div>
                      <div style={{fontSize:12.5,fontWeight:600,color:C.txt}}>{weekLabel(w.wk)}</div>
                      <div style={{height:4,borderRadius:2,background:"#e2e8f0",marginTop:5,width:"100%",maxWidth:160}}>
                        <div style={{height:"100%",borderRadius:2,background:w.total>10?C.red:w.total>5?C.amber:C.green,width:pct+"%"}}/>
                      </div>
                    </div>
                    <span style={{fontSize:12,fontWeight:700,color:C.blue,textAlign:"center"}}>{w.nbEleves}</span>
                    <span style={{fontSize:13,fontWeight:800,color:w.total>10?C.red:w.total>5?C.amber:C.green,textAlign:"right"}}>{w.total}</span>
                  </div>
                );
              })}
            </div>
          ):<div style={{fontSize:11,color:C.txtLight,textAlign:"center",padding:"30px 0"}}>Aucune donnée</div>}
        </div>
      )}

      {/* ══ Bilan Trimestriel ══════════════════════════════════════ */}
      {tab==="bilan_trim" && (()=>{
        // Agrégation absences → élève × semaine trimestre
        const eleveMap = {};
        Object.entries(data?.absences||{}).forEach(([k,absents])=>{
          const [,classe,dateStr]=k.split("||");
          if(sgClasses&&!sgClasses.includes(classe))return;
          (absents||[]).forEach(id=>{
            if(!eleveMap[id]) eleveMap[id]={id,classe,nom:(ELEVES_DB[classe]||[]).find(x=>x.id===id)?.nom||id,sems:{}};
            const pos=getSemaineTrimestre(dateStr);
            if(!pos)return;
            const k2=pos.trim+"_"+pos.sem;
            eleveMap[id].sems[k2]=(eleveMap[id].sems[k2]||0)+getDureeSVT(classe);
          });
        });
        const rows=Object.values(eleveMap).map(e=>{
          const totalTrim=[1,2,3].map(t=>[1,2,3,4,5,6].reduce((a,s)=>a+(e.sems[t+"_"+s]||0),0));
          return {...e,totalTrim};
        }).filter(e=>e.totalTrim.some(t=>t>0)).sort((a,b)=>b.totalTrim[selTrim-1]-a.totalTrim[selTrim-1]);
        const aC=(h)=>h>=15?"#b91c1c":h>=6?"#d97706":"#15803d";
        const aB=(h)=>h>=15?"#fef2f2":h>=6?"#fffbeb":"#f0fdf4";
        const aL=(h)=>h>=15?"🔴 Blâme":h>=6?"🟠 Avertissement":"🟢 Assidu";
        return(
          <div>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12,flexWrap:"wrap"}}>
              {[1,2,3].map(t=>(
                <button key={t} onClick={()=>setSelTrim(t)}
                  style={{padding:"7px 16px",borderRadius:20,border:"1.5px solid "+(selTrim===t?C.green:C.border),
                    background:selTrim===t?C.greenPale:C.white,color:selTrim===t?C.green:C.txtMuted,
                    fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>
                  Trimestre {t}
                </button>
              ))}
              <button onClick={()=>genererBilanTrimestriel(stats,vieSco,sgClasses,niveauLabel,data)}
                style={{marginLeft:"auto",padding:"7px 14px",borderRadius:10,border:"none",
                  background:"#D4AF37",color:"#0B3D20",fontWeight:800,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>
                📄 PDF
              </button>
            </div>
            <div style={{display:"flex",gap:6,marginBottom:10,flexWrap:"wrap"}}>
              {[{h:0,l:"< 6h · Assidu"},{h:6,l:"6–14h · Avertissement"},{h:15,l:"≥ 15h · Blâme"}].map((s,i)=>(
                <span key={i} style={{fontSize:10,fontWeight:700,padding:"3px 9px",borderRadius:8,background:aB(s.h),color:aC(s.h)}}>{s.l}</span>
              ))}
            </div>
            <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,overflow:"hidden"}}>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:600}}>
                  <thead>
                    <tr style={{background:"#f8fafc",borderBottom:"1px solid "+C.border}}>
                      <th style={{padding:"8px 12px",textAlign:"left",fontSize:10,fontWeight:700,color:C.txtMuted}}>ÉLÈVE</th>
                      <th style={{padding:"8px 8px",textAlign:"center",fontSize:10,fontWeight:700,color:C.txtMuted}}>CLASSE</th>
                      {[1,2,3,4,5,6].map(s=><th key={s} style={{padding:"8px 6px",textAlign:"center",fontSize:10,fontWeight:700,color:C.txtMuted}}>S{s}</th>)}
                      <th style={{padding:"8px 10px",textAlign:"center",fontSize:10,fontWeight:700,color:C.txtMuted}}>TOTAL</th>
                      <th style={{padding:"8px 10px",textAlign:"center",fontSize:10,fontWeight:700,color:C.txtMuted}}>STATUT</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.length===0?<tr><td colSpan={10} style={{padding:"32px 0",textAlign:"center",color:C.txtLight,fontSize:11}}>Aucune absence pour T{selTrim}</td></tr>
                    :rows.map((e,i)=>{
                      const total=e.totalTrim[selTrim-1];
                      return(
                        <tr key={e.id} onClick={()=>setFicheEleve(e)}
                          style={{borderBottom:"1px solid "+C.border,background:i%2===0?C.white:"#fafafa",cursor:"pointer"}}>
                          <td style={{padding:"9px 12px",fontWeight:700,color:C.txt}}>{e.nom}</td>
                          <td style={{padding:"9px 8px",textAlign:"center",fontSize:10,color:C.txtMuted}}>{e.classe}</td>
                          {[1,2,3,4,5,6].map(s=>{
                            const h=e.sems[selTrim+"_"+s]||0;
                            return <td key={s} style={{padding:"9px 6px",textAlign:"center",fontWeight:h>0?700:400,
                              fontSize:12,color:h>=4?"#b91c1c":h>=2?"#d97706":h>0?"#374151":C.txtLight}}>
                              {h>0?h+"h":"—"}
                            </td>;
                          })}
                          <td style={{padding:"9px 10px",textAlign:"center",fontWeight:800,fontSize:13,color:aC(total)}}>{total}h</td>
                          <td style={{padding:"9px 10px",textAlign:"center"}}>
                            <span style={{fontSize:9,fontWeight:700,padding:"3px 7px",borderRadius:8,background:aB(total),color:aC(total)}}>{aL(total)}</span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        );
      })()}

      {/* Retards / Sanctions / Incidents */}
      {(tab==="retards"||tab==="sanctions"||tab==="incidents") && (
        <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12.5,minWidth:440}}>
              <thead>
                <tr style={{background:"#f8fafc",borderBottom:"1px solid "+C.border}}>
                  <th style={{padding:"10px 12px",textAlign:"left",fontSize:10,fontWeight:700,color:C.txtMuted}}>Date</th>
                  <th style={{padding:"10px 12px",textAlign:"left",fontSize:10,fontWeight:700,color:C.txtMuted}}>Élève</th>
                  <th style={{padding:"10px 12px",textAlign:"left",fontSize:10,fontWeight:700,color:C.txtMuted}}>Classe</th>
                  <th style={{padding:"10px 12px",textAlign:"left",fontSize:10,fontWeight:700,color:C.txtMuted}}>Motif</th>
                  {tab!=="retards"&&<th style={{padding:"10px 12px",textAlign:"center",fontSize:10,fontWeight:700,color:C.txtMuted}}>Gravité</th>}
                </tr>
              </thead>
              <tbody>
                {vieLoading?<tr><td colSpan={5} style={{padding:24,textAlign:"center",color:C.txtLight}}>Chargement...</td></tr>
                :filteredVie.length===0?<tr><td colSpan={5} style={{padding:32,textAlign:"center",color:C.txtLight}}>
                  <div style={{fontSize:24,marginBottom:6}}>📭</div>Aucun enregistrement
                </td></tr>
                :filteredVie.map((v,i)=>(
                  <tr key={v.id} style={{borderBottom:"1px solid "+C.border,background:i%2===0?C.white:"#fafafa",cursor:"pointer"}}
                    onClick={()=>setFicheEleve({id:v.eleve_id,classe:v.classe,nom:(ELEVES_DB[v.classe]||[]).find(x=>x.id===v.eleve_id)?.nom||v.eleve_id})}>
                    <td style={{padding:"10px 12px",color:C.txtMuted,whiteSpace:"nowrap"}}>{new Date(v.date).toLocaleDateString("fr-FR",{day:"2-digit",month:"short"})}</td>
                    <td style={{padding:"10px 12px",fontWeight:600,color:C.txt}}><NomEleve eleveId={v.eleve_id} classe={v.classe}/></td>
                    <td style={{padding:"10px 12px",color:C.txtMuted}}>{v.classe}</td>
                    <td style={{padding:"10px 12px",color:C.txt,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{v.motif||"—"}</td>
                    {tab!=="retards"&&<td style={{padding:"10px 12px",textAlign:"center"}}><GravBadge g={v.gravite}/></td>}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}


// ── Calendrier scolaire 2025-2026 : date → {trim, sem} ──────────
// Trimestres chargés depuis Supabase (table annee_scolaire), fallback 2025-2026
let TRIMESTRES_DYNAMIQUES = [
  {trim:1, debut:new Date("2025-10-06")},
  {trim:2, debut:new Date("2026-01-05")},
  {trim:3, debut:new Date("2026-04-06")},
];
async function loadTrimestres() {
  try {
    const rows = await sb.get("annee_scolaire","?select=trim,debut,fin&active=eq.true&order=trim.asc");
    if (rows && rows.length >= 3) {
      TRIMESTRES_DYNAMIQUES = rows.map(r=>({trim:r.trim, debut:new Date(r.debut), fin:new Date(r.fin)}));
    }
  } catch(e) { /* fallback 2025-2026 */ }
}
function getSemaineTrimestre(dateStr) {
  const dt = new Date(dateStr);
  if (isNaN(dt)) return null;
  for (const {trim, debut} of TRIMESTRES_DYNAMIQUES) {
    const ms = dt - debut;
    if (ms < 0) continue;
    const sem = Math.floor(ms / (7 * 86400000)) + 1;
    if (sem >= 1 && sem <= 6) return { trim, sem };
  }
  return null;
}
// ════════════════════════════════════════════════════════════════════
// SYSTÈME BULLETINS — EduPilot Cameroun / MINESEC
// Format : bulletin de notes par séquence avec coefficients officiels
// ════════════════════════════════════════════════════════════════════

// Chargement des coefficients depuis Supabase
let COEFFICIENTS_DB = [];
async function loadCoefficients() {
  try {
    const rows = await sb.get("coefficients","?select=niveau,serie,matiere,coef&order=niveau.asc,matiere.asc");
    if (rows && rows.length > 0) COEFFICIENTS_DB = rows;
  } catch(e) {}
}

// Résoudre niveau+série depuis nom de classe
function parseClasseNiveauSerie(classe) {
  const c = (classe||"").trim();
  let niveau = "6ème", serie = "";
  if (/^6/i.test(c)) niveau = "6ème";
  else if (/^5/i.test(c)) niveau = "5ème";
  else if (/^4/i.test(c)) {
    niveau = "4ème";
    if (/ALL|Alle/i.test(c)) serie = "ALL";
    else if (/ESP|Esp/i.test(c)) serie = "ESP";
    else if (/ARB|Ara/i.test(c)) serie = "ARB";
    else if (/ITA|Ita/i.test(c)) serie = "ITA";
    else if (/CHI|Chi/i.test(c)) serie = "CHI";
  }
  else if (/^3/i.test(c)) {
    niveau = "3ème";
    if (/ALL|Alle/i.test(c)) serie = "ALL";
    else if (/ESP|Esp/i.test(c)) serie = "ESP";
    else if (/ARB|Ara/i.test(c)) serie = "ARB";
    else if (/ITA|Ita/i.test(c)) serie = "ITA";
    else if (/CHI|Chi/i.test(c)) serie = "CHI";
  }
  else if (/^2nde|^2de|^2nd/i.test(c)) {
    niveau = "2nde";
    if (/ALL|Alle/i.test(c)) serie = "ALL";
    else if (/ESP|Esp/i.test(c)) serie = "ESP";
    else if (/ARA|Ara/i.test(c)) serie = "ARA";
    else if (/ITA|Ita/i.test(c)) serie = "ITA";
    else if (/CHI|Chi/i.test(c)) serie = "CHI";
    else serie = "C";
  }
  else if (/^1/i.test(c)) {
    niveau = "1ère";
    if (/A4/i.test(c)) {
      if (/ALL|Alle/i.test(c)) serie = "ALL";
      else if (/ESP|Esp/i.test(c)) serie = "ESP";
      else if (/ARA|Ara/i.test(c)) serie = "ARA";
      else if (/ITA|Ita/i.test(c)) serie = "ITA";
      else if (/CHI|Chi/i.test(c)) serie = "CHI";
      else serie = "A4";
    }
    else if (/Ti/i.test(c)) serie = "Ti";
    else if (/C/.test(c)) serie = "C";
    else if (/D/.test(c)) serie = "D";
    else serie = "A4";
  }
  else if (/^Tle|^Ter/i.test(c)) {
    niveau = "Tle";
    if (/A4/i.test(c)) {
      if (/ALL|Alle/i.test(c)) serie = "ALL";
      else if (/ESP|Esp/i.test(c)) serie = "ESP";
      else if (/ARA|Ara/i.test(c)) serie = "ARA";
      else if (/ITA|Ita/i.test(c)) serie = "ITA";
      else if (/CHI|Chi/i.test(c)) serie = "CHI";
      else serie = "A4";
    }
    else if (/Ti/i.test(c)) serie = "Ti";
    else if (/C/.test(c)) serie = "C";
    else if (/D/.test(c)) serie = "D";
    else serie = "A4";
  }
  return { niveau, serie };
}

function getCoefsForClasse(classe) {
  const { niveau, serie } = parseClasseNiveauSerie(classe);
  return COEFFICIENTS_DB.filter(r => r.niveau === niveau && r.serie === (serie||""));
}

function calcMoyClasse(classe, sequence, matiere, notesIndex, elevesClasse) {
  const key = `${classe}||${matiere}-S${sequence}`;
  const notes = elevesClasse.map(e => (notesIndex[key]||{})[e.id]).filter(n => n !== undefined && n !== null && n !== "").map(Number);
  if (!notes.length) return null;
  return Math.round((notes.reduce((a,b)=>a+b,0)/notes.length)*100)/100;
}

function calcRangsClasse(classe, sequence, notesIndex, elevesClasse) {
  const coefs = getCoefsForClasse(classe);
  const moyennes = elevesClasse.map(e => {
    let tp=0, tc=0;
    coefs.forEach(({matiere,coef}) => {
      const k=`${classe}||${matiere}-S${sequence}`;
      const n=(notesIndex[k]||{})[e.id];
      if(n!==undefined&&n!==null&&n!==""){tp+=+n*coef;tc+=coef;}
    });
    return {id:e.id, nom:e.nom, moyenne: tc>0?Math.round(tp/tc*100)/100:null};
  }).filter(e=>e.moyenne!==null).sort((a,b)=>b.moyenne-a.moyenne);
  const rangs={};
  moyennes.forEach((e,i)=>{rangs[e.id]=i+1;});
  return {rangs, classees:moyennes};
}

function getMention(m) {
  if(m===null) return "";
  if(m>=16) return "Bien";
  if(m>=14) return "Assez Bien";
  if(m>=10) return "Passable";
  return "Insuffisant";
}

function getAppreciation(n) {
  if(n===null) return "—";
  if(n>=18) return "Excellent"; if(n>=16) return "Très Bien";
  if(n>=14) return "Bien"; if(n>=12) return "Assez Bien";
  if(n>=10) return "Passable"; if(n>=8) return "Insuffisant";
  return "Très Insuffisant";
}

function genBulletin(opts) {
  const {
    eleve, classe, sequence, annee="2025-2026",
    notesIndex={}, absencesIndex={}, elevesClasse=[],
    appreciation="", decision="", mention="",
    profPrincipalNom="À définir",
    conduite=null, retards=0,
    exclusionsH=0, exclusionsJ=0,
    consignesH=0, consignesJ=0,
    blameTravail="Aucun", blameConduite="Aucun"
  } = opts;
  const coefs = getCoefsForClasse(classe);
  const {niveau,serie} = parseClasseNiveauSerie(classe);
  const seq=parseInt(sequence);
  const trim=seq<=2?1:seq<=4?2:3;
  const G="#0B4D2C", gold="#D4AF37", rouge="#dc2626";

  // Calcul notes par matière pour toutes les séquences disponibles (S1 à Sseq)
  // On affiche S1..S6 avec "—" si pas encore saisie
  const SEQUENCES=[1,2,3,4,5,6];

  let totalPts=0,totalCoef=0,nbMatieres=0;
  const lignes=coefs.map(({matiere,coef})=>{
    const notesSeq=SEQUENCES.map(s=>{
      const k=`${classe}||${matiere}-S${s}`;
      const n=(notesIndex[k]||{})[eleve.id];
      return (n!==undefined&&n!==null&&n!=="")?+n:null;
    });
    // Moyenne sur séquences saisies jusqu'à seq
    const notesSaisies=notesSeq.slice(0,seq).filter(n=>n!==null);
    const moy=notesSaisies.length>0?Math.round(notesSaisies.reduce((a,b)=>a+b,0)/notesSaisies.length*100)/100:null;
    const pts=moy!==null?Math.round(moy*coef*100)/100:null;

    // Stats classe pour cette matière à cette séquence
    const keySeq=`${classe}||${matiere}-S${seq}`;
    const toutes=elevesClasse.map(e=>(notesIndex[keySeq]||{})[e.id]).filter(n=>n!==undefined&&n!==null&&n!=="").map(Number);
    const moyC=toutes.length?Math.round(toutes.reduce((a,b)=>a+b,0)/toutes.length*100)/100:null;
    const minC=toutes.length?Math.min(...toutes):null;
    const maxC=toutes.length?Math.max(...toutes):null;

    // Rang pour cette matière
    const noteEleve=(notesIndex[keySeq]||{})[eleve.id];
    const nEleve=(noteEleve!==undefined&&noteEleve!==null&&noteEleve!=="")?+noteEleve:null;
    const rang=nEleve!==null?toutes.filter(n=>n>nEleve).length+1:null;

    if(moy!==null){totalPts+=(pts||0);totalCoef+=coef;nbMatieres++;}

    const col=moy===null?"#9ca3af":moy>=10?"#15803d":"#dc2626";
    const appTxt=moy===null?"—":moy>=16?"Très Bien":moy>=14?"Bien":moy>=12?"Assez Bien":moy>=10?"Passable":moy>=8?"Insuffisant":"Faible";
    const appCol=moy===null?"#9ca3af":moy>=14?"#16a34a":moy>=10?"#d97706":"#dc2626";
    const appDot=moy===null?"#9ca3af":moy>=14?"#16a34a":moy>=10?"#d97706":"#dc2626";

    const seqCols=notesSeq.map((n,i)=>{
      if(i>=seq) return `<td style="background:#f9fafb;color:#d1d5db;text-align:center;padding:4px;">—</td>`;
      const nc=n===null?"#9ca3af":n>=10?"#15803d":"#dc2626";
      return `<td style="text-align:center;padding:4px;color:${nc};font-weight:${n!==null?"600":"400"}">${n!==null?n.toFixed(2):"—"}</td>`;
    }).join("");

    return `<tr style="border-bottom:1px solid #f0f0f0;">
      <td style="padding:5px 8px;font-size:11px;">${matiere}</td>
      ${seqCols}
      <td style="text-align:center;padding:4px;font-weight:700;color:${col};">${moy!==null?moy.toFixed(2):"—"}</td>
      <td style="text-align:center;padding:4px;font-size:11px;">${coef}</td>
      <td style="text-align:center;padding:4px;font-size:11px;">${pts!==null?pts.toFixed(2):"—"}</td>
      <td style="text-align:center;padding:4px;font-size:11px;">${rang!==null?rang:"—"}</td>
      <td style="text-align:center;padding:4px;font-size:10px;color:#6b7280;">${minC!==null?minC.toFixed(1):"—"}</td>
      <td style="text-align:center;padding:4px;font-size:10px;color:#6b7280;">${moyC!==null?moyC.toFixed(1):"—"}</td>
      <td style="text-align:center;padding:4px;font-size:10px;color:#6b7280;">${maxC!==null?maxC.toFixed(1):"—"}</td>
      <td style="padding:4px 8px;font-size:10px;color:${appCol};">${appTxt} <span style="color:${appDot};">●</span></td>
    </tr>`;
  });

  const moyenne=totalCoef>0?Math.round(totalPts/totalCoef*100)/100:null;
  const {rangs}=calcRangsClasse(classe,seq,notesIndex,elevesClasse);
  const rang=rangs[eleve.id]||"—";
  const effectif=elevesClasse.length;
  const mentionAff=mention||getMention(moyenne);
  const decColor=decision==="Passage"||decision?.includes("ADMIS")?"#16a34a":decision==="Redoublement"?"#dc2626":"#d97706";

  // Grouper matières
  const groupes=[
    {title:"MATIÈRES SCIENTIFIQUES", matieres:["Mathématiques","PCT","SVT","Sciences Physiques","Physique-Chimie","Informatique","SVTEEHB Théorique","Physique-Chimie-Tech."]},
    {title:"MATIÈRES LITTÉRAIRES", matieres:["Français","Anglais","Histoire-Géographie","Philosophie","ECM","LV2","Allemand","Espagnol","Arabe","Italien","Chinois","Étude de texte","Expression orale","Éducation à la citoyenneté et à la morale"]},
  ];
  const autresMatieres=coefs.map(c=>c.matiere).filter(m=>!groupes.flatMap(g=>g.matieres).some(gm=>m.toLowerCase().includes(gm.toLowerCase())));

  const renderGroupe=(groupLabel, matiereList, allLignes)=>{
    const matching=allLignes.filter((_,i)=>{
      const m=coefs[i]?.matiere||"";
      return matiereList.some(gm=>m.toLowerCase().includes(gm.toLowerCase()));
    });
    if(!matching.length) return "";
    return `<tr style="background:#1f2937;"><td colspan="15" style="padding:5px 8px;font-size:10px;font-weight:900;color:#fff;text-transform:uppercase;letter-spacing:.8px;">${groupLabel}</td></tr>${matching.join("")}`;
  };

  const listeAutres=lignes.filter((_,i)=>autresMatieres.includes(coefs[i]?.matiere));

  const armoiries=`<img src="data:image/jpeg;base64,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" width="80" height="80" style="border-radius:50%;object-fit:cover;" alt="Logo Lycée de Kakatare"/>`;

  return `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:"Arial Narrow",Arial,sans-serif;color:#1f2937;font-size:11px;padding:10px;background:#fff;}
.page{max-width:210mm;margin:0 auto;}
.top-header{display:grid;grid-template-columns:120px 1fr 120px;gap:8px;align-items:center;border-bottom:4px solid ${G};padding-bottom:8px;margin-bottom:8px;}
.top-left{font-size:9px;color:#374151;line-height:1.7;}
.top-center{text-align:center;}
.etablissement{font-size:20px;font-weight:900;color:${G};text-transform:uppercase;letter-spacing:2px;margin:4px 0;}
.sous-etablissement{font-size:9px;color:#6b7280;}
.bulletin-title{font-size:22px;font-weight:900;letter-spacing:4px;text-align:center;color:#fff;background:${G};padding:6px 0;margin:6px 0;border-radius:2px;}
.top-right{text-align:right;font-size:9px;color:#374151;line-height:1.7;}
.annee-box{border:2px solid ${G};border-radius:6px;padding:6px 12px;text-align:center;font-size:10px;font-weight:700;}
.annee-num{font-size:16px;font-weight:900;color:${G};}
.trim-box{border:2px solid ${gold};border-radius:6px;padding:4px 10px;text-align:center;margin-top:4px;font-size:10px;}
.trim-num{font-size:22px;font-weight:900;color:${gold};}

.info-eleve{display:grid;grid-template-columns:80px 1fr auto;gap:10px;background:#f8fafc;border:1px solid #e5e7eb;border-radius:8px;padding:10px;margin-bottom:8px;align-items:start;}
.photo-placeholder{width:70px;height:85px;background:#e5e7eb;border-radius:4px;display:flex;align-items:center;justify-content:center;font-size:28px;color:#9ca3af;}
.info-grid{display:grid;grid-template-columns:1fr 1fr;gap:3px;}
.info-line{font-size:10px;padding:2px 0;border-bottom:1px dotted #e5e7eb;}
.info-label{color:#6b7280;font-weight:600;}
.info-val{font-weight:700;color:#1f2937;}
.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:6px;}
.kpi{background:#fff;border:1.5px solid ${G};border-radius:6px;padding:6px;text-align:center;}
.kpi-num{font-size:16px;font-weight:900;color:${G};}
.kpi-den{font-size:10px;color:#6b7280;}
.kpi-label{font-size:8px;color:#6b7280;margin-top:2px;text-transform:uppercase;}

.section-title{background:${G};color:#fff;padding:4px 8px;font-size:10px;font-weight:900;text-transform:uppercase;letter-spacing:.8px;margin:8px 0 4px;border-radius:2px;}

table{width:100%;border-collapse:collapse;font-size:10px;margin-bottom:8px;}
th{background:#374151;color:#fff;padding:4px 6px;font-size:9px;font-weight:700;text-align:center;}
th:first-child{text-align:left;}
tr:nth-child(even){background:#f9fafb;}

.bottom-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-top:8px;}
.bottom-card{border:1px solid #e5e7eb;border-radius:6px;padding:8px;}
.card-title{font-size:9px;font-weight:900;color:${G};text-transform:uppercase;margin-bottom:6px;border-bottom:1px solid ${G};padding-bottom:3px;}
.card-row{display:flex;justify-content:space-between;font-size:9px;padding:2px 0;border-bottom:1px dotted #f0f0f0;}
.card-label{color:#6b7280;}
.card-val{font-weight:700;}

.decision-box{background:${decColor}15;border:2px solid ${decColor};border-radius:8px;padding:8px 14px;text-align:center;margin:8px 0;}
.decision-label{font-size:9px;color:#6b7280;margin-bottom:2px;}
.decision-text{font-size:16px;font-weight:900;color:${decColor};text-transform:uppercase;}
.decision-stats{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;margin-top:6px;font-size:9px;}
.ds-item{text-align:center;}.ds-label{color:#6b7280;}.ds-val{font-weight:700;font-size:11px;}

.sig-grid{display:grid;grid-template-columns:1fr 1fr;gap:30px;margin-top:12px;}
.sig-box{text-align:center;border-top:1px solid #e5e7eb;padding-top:8px;font-size:10px;}
.sig-name{font-weight:700;font-size:11px;margin-top:4px;color:${G};}
.footer-bar{background:${G};color:${gold};text-align:center;padding:5px;font-size:10px;font-weight:700;letter-spacing:3px;margin-top:12px;border-radius:2px;}
@media print{body{padding:4px;}@page{margin:8mm;size:A4;}table{page-break-inside:avoid;}}
</style>
<script>window.onload=()=>window.print();</script>
</head><body><div class="page">

<!-- EN-TÊTE -->
<div class="top-header">
  <div class="top-left">
    <strong>RÉGION DE L'EXTRÊME-NORD</strong><br>
    DÉPARTEMENT DU DIAMARÉ<br>
    ARRONDISSEMENT DE MAROUA II
  </div>
  <div class="top-center">
    <div style="font-size:9px;color:#374151;font-weight:700;">RÉPUBLIQUE DU CAMEROUN</div>
    <div style="font-size:8px;color:${gold};font-style:italic;">Paix – Travail – Patrie</div>
    <div style="margin:4px 0;">${armoiries}</div>
    <div class="etablissement">LYCÉE DE KAKATARE-MAROUA</div>
    <div class="sous-etablissement">BP 162 Maroua – Tél. 222 29 21 63 – Mle 0CJ1GSF8111231106</div>
  </div>
  <div class="top-right">
    <div class="annee-box">
      ANNÉE SCOLAIRE<br>
      <div class="annee-num">${annee}</div>
    </div>
    <div class="trim-box">
      TRIMESTRE<br>
      <div class="trim-num">${trim}</div>
    </div>
  </div>
</div>

<div class="bulletin-title">✦ BULLETIN SCOLAIRE ✦</div>

<!-- INFOS ÉLÈVE -->
<div class="info-eleve">
  <div class="photo-placeholder">👤</div>
  <div>
    <div style="font-size:15px;font-weight:900;color:${G};margin-bottom:6px;">${eleve.nom}</div>
    <div class="info-grid">
      <div class="info-line"><span class="info-label">Classe : </span><span class="info-val">${classe}</span></div>
      <div class="info-line"><span class="info-label">Matricule : </span><span class="info-val">${eleve.numero||"—"}</span></div>
      <div class="info-line"><span class="info-label">Née le : </span><span class="info-val">—</span></div>
      <div class="info-line"><span class="info-label">Sexe : </span><span class="info-val">${eleve.sexe==="G"||eleve.sexe==="M"?"Masculin":"Féminin"}</span></div>
      <div class="info-line"><span class="info-label">Statut : </span><span class="info-val">Non redoublant(e)</span></div>
      <div class="info-line"><span class="info-label">Prof. principal : </span><span class="info-val">${profPrincipalNom}</span></div>
    </div>
  </div>
  <div class="kpis">
    <div class="kpi">
      <div class="kpi-num" style="font-size:13px;">${moyenne!==null?moyenne.toFixed(2):"—"}</div>
      <div class="kpi-den">/20</div>
      <div class="kpi-label">Moyenne générale</div>
    </div>
    <div class="kpi">
      <div class="kpi-num">${rang}</div>
      <div class="kpi-den">/${effectif}</div>
      <div class="kpi-label">Rang</div>
    </div>
    <div class="kpi">
      <div class="kpi-num">${nbMatieres}</div>
      <div class="kpi-den">/${coefs.length}</div>
      <div class="kpi-label">Matières évaluées</div>
    </div>
    <div class="kpi">
      <div class="kpi-num" style="font-size:12px;">${totalPts.toFixed(1)}</div>
      <div class="kpi-label">Total points</div>
    </div>
    <div class="kpi">
      <div class="kpi-num">${conduite!==null?conduite:"—"}</div>
      <div class="kpi-den">/20</div>
      <div class="kpi-label">Conduite</div>
    </div>
    <div class="kpi">
      <div class="kpi-num" style="color:${mentionAff==="Bien"?"#16a34a":mentionAff==="Assez Bien"?"#2563eb":"#d97706"};font-size:10px;">${mentionAff||"—"}</div>
      <div class="kpi-label">Mention</div>
    </div>
  </div>
</div>

<!-- TABLEAU RÉSULTATS -->
<div class="section-title">📊 Résultats Académiques</div>
<table>
  <thead>
    <tr>
      <th style="width:22%;text-align:left;">MATIÈRES</th>
      <th>S1</th><th>S2</th><th>S3</th><th>S4</th><th>S5</th><th>S6</th>
      <th>MOY.</th><th>COEF.</th><th>POINTS</th><th>RANG</th>
      <th>MIN</th><th>MOY.CL</th><th>MAX</th>
      <th style="width:14%;">APPRÉCIATION</th>
    </tr>
  </thead>
  <tbody>
    ${renderGroupe("MATIÈRES SCIENTIFIQUES",groupes[0].matieres,lignes)}
    ${renderGroupe("MATIÈRES LITTÉRAIRES",groupes[1].matieres,lignes)}
    ${listeAutres.length?`<tr style="background:#1f2937;"><td colspan="15" style="padding:5px 8px;font-size:10px;font-weight:900;color:#fff;text-transform:uppercase;letter-spacing:.8px;">AUTRES MATIÈRES</td></tr>${listeAutres.join("")}`:""}
  </tbody>
</table>

<!-- BAS DE PAGE : 3 colonnes -->
<div class="bottom-grid">
  <!-- Assiduité -->
  <div class="bottom-card">
    <div class="card-title">⏱ Assiduité</div>
    <table style="width:100%;font-size:9px;margin:0;">
      <tr><th></th><th>TOTAL</th><th>JUSTIFIÉES</th><th>NON JUST.</th></tr>
      <tr><td class="card-label">Absences</td><td style="text-align:center;font-weight:700;">— h</td><td style="text-align:center;">0 h</td><td style="text-align:center;color:#dc2626;">— h</td></tr>
    </table>
    <div class="card-row" style="margin-top:4px;"><span class="card-label">Retards :</span><span class="card-val">${retards}</span></div>
  </div>

  <!-- Conduite & Vie scolaire -->
  <div class="bottom-card">
    <div class="card-title">🛡 Conduite & Vie scolaire</div>
    <div class="card-row"><span class="card-label">Note de conduite :</span><span class="card-val">${conduite!==null?conduite+"/20":"—"}</span></div>
    <div class="card-row"><span class="card-label">Retards :</span><span class="card-val">${retards}</span></div>
    <div class="card-row"><span class="card-label">Exclusions (heures) :</span><span class="card-val">${exclusionsH} h</span></div>
    <div class="card-row"><span class="card-label">Exclusions (jours) :</span><span class="card-val">${exclusionsJ} j</span></div>
    <div class="card-row"><span class="card-label">Consignes (heures) :</span><span class="card-val">${consignesH} h</span></div>
    <div class="card-row"><span class="card-label">Blâme travail :</span><span class="card-val">${blameTravail}</span></div>
    <div class="card-row"><span class="card-label">Blâme conduite :</span><span class="card-val">${blameConduite}</span></div>
  </div>

  <!-- Conseil de Classe -->
  <div class="bottom-card">
    <div class="card-title">👥 Conseil de Classe</div>
    <div style="font-size:9px;font-weight:700;margin-bottom:3px;">Appréciation générale :</div>
    <div style="font-size:9px;color:#374151;font-style:italic;min-height:35px;">${appreciation||"—"}</div>
    ${decision?`<div style="margin-top:6px;background:${decColor}15;border:1px solid ${decColor};border-radius:4px;padding:4px 8px;text-align:center;"><div style="font-size:8px;color:#6b7280;">ENCOURAGEMENTS / DÉCISION</div><div style="font-size:12px;font-weight:900;color:${decColor};">${decision}</div></div>`:""}
  </div>
</div>

<!-- SIGNATURES -->
<div class="sig-grid">
  <div class="sig-box">
    <div style="height:40px;"></div>
    <div>LE PROFESSEUR PRINCIPAL</div>
    <div class="sig-name">${profPrincipalNom}</div>
  </div>
  <div class="sig-box">
    <div style="height:40px;"></div>
    <div>LE CHEF D'ÉTABLISSEMENT</div>
    <div class="sig-name">Le Proviseur</div>
  </div>
</div>

<div class="footer-bar">✦ DISCIPLINE – TRAVAIL – RÉUSSITE ✦</div>
</div></body></html>`;

function BulletinsPage() {
  const {user,data} = useApp();
  const {isMobile} = useDevice();
  const isAdm = isAdminRole(user?.role)||user?.role==="censeur";
  // Classes disponibles selon rôle
  const classes = useMemo(()=>{
    if(!data) return [];
    const set=new Set();
    if(isAdm){
      Object.values(data.users||{}).forEach(u=>(u.classes||[]).forEach(c=>set.add(c)));
    } else {
      (user?.classes||[]).forEach(c=>set.add(c));
    }
    return [...set].sort();
  },[data,user,isAdm]);

  const [selClasse,setSelClasse]=useState("");
  const [selSeq,setSelSeq]=useState(1);
  const [selEleve,setSelEleve]=useState(null);
  const [previewHtml,setPreviewHtml]=useState(null);
  const [searchTerm,setSearchTerm]=useState("");

  const elevesClasse = useMemo(()=>{
    if(!selClasse||!data) return [];
    // Chercher dans ELEVES_DB (frontend) — fallback si table eleves pas encore chargée
    return (ELEVES_DB[selClasse]||[]).map(e=>({id:e.id, nom:e.n||e.nom||"", sexe:e.g||e.sexe||"G", numero:e.num||e.numero||""}));
  },[selClasse,data]);

  const filteredEleves = useMemo(()=>{
    if(!searchTerm) return elevesClasse;
    return elevesClasse.filter(e=>(e.nom||"").toLowerCase().includes(searchTerm.toLowerCase()));
  },[elevesClasse,searchTerm]);

  const {rangs} = useMemo(()=>{
    if(!selClasse||!data) return {rangs:{}};
    return calcRangsClasse(selClasse,selSeq,data.notes||{},elevesClasse);
  },[selClasse,selSeq,data,elevesClasse]);

  const handlePreview = (eleve) => {
    const html = genBulletin({
      eleve, classe:selClasse, sequence:selSeq,
      notesIndex:data?.notes||{}, absencesIndex:data?.absences||{},
      elevesClasse
    });
    setPreviewHtml(stripAutoPrint(html));
    setSelEleve(eleve);
  };

  const handlePrintAll = () => {
    if(!filteredEleves.length) return;
    const htmls = filteredEleves.map(e=>genBulletin({
      eleve:e, classe:selClasse, sequence:selSeq,
      notesIndex:data?.notes||{}, absencesIndex:data?.absences||{},
      elevesClasse
    })).join('<div style="page-break-after:always;"></div>');
    imprimerHTML(htmls);
  };

  const inp={border:"1.5px solid #e5e7eb",borderRadius:8,padding:"8px 12px",fontSize:13,fontFamily:"inherit",outline:"none",background:"#fff",width:"100%",boxSizing:"border-box"};

  return (
    <div style={{padding:isMobile?"12px":"24px",maxWidth:1000,margin:"0 auto"}}>
      {/* En-tête */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:10}}>
        <div>
          <div style={{fontSize:isMobile?16:20,fontWeight:900,color:"#0B4D2C"}}>📋 Bulletins de notes</div>
          <div style={{fontSize:12,color:"#6b7280",marginTop:2}}>Génération · Séquences S1 à S6</div>
        </div>
        {selClasse && filteredEleves.length>0 && (
          <button onClick={handlePrintAll}
            style={{padding:"10px 18px",borderRadius:10,border:"none",background:"#0B4D2C",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:6}}>
            🖨 Imprimer tout ({filteredEleves.length})
          </button>
        )}
      </div>

      {/* Filtres */}
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr 1fr",gap:12,marginBottom:16}}>
        <div>
          <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:4,textTransform:"uppercase",letterSpacing:".5px"}}>Classe</div>
          <select value={selClasse} onChange={e=>{setSelClasse(e.target.value);setSelEleve(null);}} style={inp}>
            <option value="">— Choisir une classe —</option>
            {classes.map(c=><option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div>
          <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:4,textTransform:"uppercase",letterSpacing:".5px"}}>Séquence</div>
          <select value={selSeq} onChange={e=>setSelSeq(+e.target.value)} style={inp}>
            {[1,2,3,4,5,6].map(s=><option key={s} value={s}>Séquence {s} — Trimestre {s<=2?1:s<=4?2:3}</option>)}
          </select>
        </div>
        <div>
          <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:4,textTransform:"uppercase",letterSpacing:".5px"}}>Rechercher</div>
          <input value={searchTerm} onChange={e=>setSearchTerm(e.target.value)} style={inp} placeholder="Nom de l'élève…"/>
        </div>
      </div>

      {/* Liste élèves */}
      {!selClasse ? (
        <div style={{textAlign:"center",padding:60,color:"#9ca3af"}}>
          <div style={{fontSize:40,marginBottom:12}}>📋</div>
          <div style={{fontSize:15,fontWeight:700}}>Choisissez une classe pour commencer</div>
        </div>
      ) : filteredEleves.length===0 ? (
        <div style={{textAlign:"center",padding:40,color:"#9ca3af"}}>Aucun élève dans cette classe</div>
      ) : (
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          <div style={{display:"grid",gridTemplateColumns:"auto 1fr auto auto auto",gap:10,padding:"8px 14px",background:"#f0fdf4",borderRadius:8,fontSize:11,fontWeight:700,color:"#0B4D2C",textTransform:"uppercase",letterSpacing:".5px"}}>
            <div>#</div><div>Nom</div><div>Moy.</div><div>Rang</div><div></div>
          </div>
          {filteredEleves.map((e,i)=>{
            // Calculer moyenne à la volée
            const coefs=getCoefsForClasse(selClasse);
            let tp=0,tc=0;
            coefs.forEach(({matiere,coef})=>{
              const k=`${selClasse}||${matiere}-S${selSeq}`;
              const n=(data?.notes?.[k]||{})[e.id];
              if(n!==undefined&&n!==null&&n!==""){tp+=+n*coef;tc+=coef;}
            });
            const moy=tc>0?Math.round(tp/tc*100)/100:null;
            const rang=rangs[e.id]||"—";
            const moyCol=moy===null?"#9ca3af":moy>=10?"#15803d":"#dc2626";
            return (
              <div key={e.id} style={{background:"#fff",borderRadius:10,border:"1px solid #e5e7eb",padding:"12px 16px",display:"grid",gridTemplateColumns:"40px 1fr 70px 50px auto",alignItems:"center",gap:10}}>
                <div style={{width:36,height:36,borderRadius:"50%",background:"#0B4D2C",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:12,flexShrink:0}}>
                  {(e.nom||"?")[0]}
                </div>
                <div>
                  <div style={{fontWeight:700,fontSize:13,color:"#1f2937"}}>{e.nom}</div>
                  <div style={{fontSize:11,color:"#6b7280"}}>{e.sexe==="G"||e.sexe==="M"?"Garçon":"Fille"}</div>
                </div>
                <div style={{fontWeight:900,fontSize:15,color:moyCol}}>
                  {moy!==null?moy.toFixed(2):"—"}<span style={{fontSize:10,color:"#9ca3af"}}>/20</span>
                </div>
                <div style={{fontSize:13,fontWeight:700,color:"#6b7280"}}>{rang}</div>
                <button onClick={()=>handlePreview(e)}
                  style={{padding:"6px 14px",borderRadius:8,border:"1px solid #0B4D2C",background:"#fff",color:"#0B4D2C",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>
                  👁 Bulletin
                </button>
              </div>
            );
          })}
        </div>
      )}

      {/* Aperçu PDF */}
      {previewHtml && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.6)",zIndex:2100,display:"flex",flexDirection:"column",padding:16}}>
          <div style={{background:"#fff",borderRadius:12,flex:1,display:"flex",flexDirection:"column",overflow:"hidden",maxWidth:900,margin:"0 auto",width:"100%"}}>
            <div style={{padding:"12px 18px",borderBottom:"1px solid #e5e7eb",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
              <div style={{fontSize:13,fontWeight:800,color:"#0B4D2C"}}>📋 Bulletin — {selEleve?.nom} · {selClasse} · S{selSeq}</div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>imprimerHTML(previewHtml)}
                  style={{padding:"7px 14px",borderRadius:8,border:"none",background:"#0B4D2C",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Imprimer</button>
                <button onClick={()=>setPreviewHtml(null)}
                  style={{padding:"7px 14px",borderRadius:8,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#374151",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Fermer</button>
              </div>
            </div>
            <div style={{flex:1,overflow:"hidden",background:"#e5e7eb",padding:12}}>
              <iframe srcDoc={previewHtml} title="bulletin" style={{width:"100%",height:"100%",border:"none",borderRadius:8,background:"#fff"}}/>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function DashboardCenseur() {
  const {rawData:data, setPage} = useApp();
  const {isMobile} = useDevice();
  const [loading, setLoading] = useState(true);
  const [stats, setStats] = useState({
    nbEns:0, tauxMoyen:0, nbAlerteProg:0, nbAbsTotal:0,
    tauxParEns:[], absParClasse:[], edtCoverage:{oui:0,non:0},
    classesAlertes:[]
  });

  useEffect(()=>{
    if(!data) return;
    const ens = Object.values(data.users||{}).filter(u=>u.role!=="proviseur"&&u.role!=="surveillant_general"&&u.role!=="censeur");

    // Couverture programme par enseignant
    const tauxParEns = ens.map(e=>{
      let tf=0,tr=0;
      (e.classes||[]).forEach(cl=>{
        const k=e.id+"||"+cl;
        const f=((data.prog||{})[k]||[]).length;
        const code=resolveProgCode(cl);
        const meta=code?PROG_META[code]:null;
        if(meta){tf+=f;tr+=meta.lpRef;}
      });
      return{id:e.id,nom:e.nom,col:getColor(e.id),ini:getIni(e.nom),
        classes:(e.classes||[]).length,taux:tr>0?Math.min(100,Math.round(tf/tr*100)):0,tf,tr};
    }).sort((a,b)=>a.taux-b.taux);
    const tauxMoyen = tauxParEns.length ? Math.round(tauxParEns.reduce((s,e)=>s+e.taux,0)/tauxParEns.length) : 0;
    const nbAlerteProg = tauxParEns.filter(e=>e.taux<50).length;

    // Absences par classe
    const absMap={};
    Object.entries(data.absences||{}).forEach(([k,abs])=>{
      const [,cl]=k.split("||");
      absMap[cl]=(absMap[cl]||0)+(abs?abs.length:0);
    });
    const nbAbsTotal = Object.values(absMap).reduce((s,n)=>s+n,0);
    const absParClasse = Object.entries(absMap)
      .map(([cl,n])=>({cl,n}))
      .sort((a,b)=>b.n-a.n).slice(0,12);

    // Couverture EDT (classes avec au moins 1 créneau)
    const classesAvecEdt = new Set();
    Object.values(data.edtBase||{}).forEach(slots=>{
      Object.keys(slots||{}).forEach(cl=>classesAvecEdt.add(cl));
    });
    const totalClasses = CLASSES_REELLES.length;
    const edtOui = Math.min(classesAvecEdt.size, totalClasses);
    const edtCoverage = {oui:edtOui, non:totalClasses-edtOui, total:totalClasses};

    // Classes avec aucun programme saisi
    const classesAlertes = CLASSES_REELLES.filter(c=>{
      return !Object.keys(data.prog||{}).some(k=>k.endsWith("||"+c.code));
    }).map(c=>c.code).slice(0,10);

    setStats({nbEns:ens.length,tauxMoyen,nbAlerteProg,nbAbsTotal,tauxParEns,absParClasse,edtCoverage,classesAlertes});
    setLoading(false);
  },[data]);

  const tauCol = t => t>=75?C.green:t>=50?C.amber:C.red;

  return(
    <div style={{padding:"20px 20px 40px",display:"flex",flexDirection:"column",gap:18}}>
      {/* En-tête */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:10}}>
        <div>
          <h2 style={{fontSize:20,fontWeight:800,color:C.txt,margin:0}}>Tableau de bord Censeur 📐</h2>
          <p style={{color:C.txtMuted,margin:"3px 0 0",fontSize:12}}>{new Date().toLocaleDateString("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric"})} · Vue pédagogique globale</p>
        </div>
        <div style={{display:"flex",gap:8}}>
          {[
            {label:"Programmes",page:"programme",emoji:"📖"},
            {label:"EDT",page:"edt",emoji:"📅"},
            {label:"Élèves",page:"eleves",emoji:"👥"},
          ].map(({label,page,emoji})=>(
            <button key={page} onClick={()=>setPage(page)}
              style={{padding:"8px 12px",borderRadius:10,border:"1px solid "+C.border,background:C.white,color:C.txt,fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:5}}>
              {emoji} {label}
            </button>
          ))}
        </div>
      </div>

      {/* KPI */}
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <KpiCard label="Couverture programme" value={stats.tauxMoyen+"%"} sub={stats.tauxMoyen>=75?"Objectif atteint ✓":"Sous l'objectif"} subColor={tauCol(stats.tauxMoyen)} iconEmoji="📊" bg={C.greenPale} loading={loading} delay={0}/>
        <KpiCard label="Enseignants en retard" value={stats.nbAlerteProg} sub="< 50% couverture" iconEmoji="⚠️" bg={C.redPale} subColor={C.red} loading={loading} delay={0.05}/>
        <KpiCard label="EDT configuré" value={stats.edtCoverage.oui+"/"+stats.edtCoverage.total} sub="classes avec créneaux" iconEmoji="📅" bg={C.bluePale} subColor={C.blue} loading={loading} delay={0.1}/>
        <KpiCard label="Absences totales" value={stats.nbAbsTotal} sub="Toutes classes" iconEmoji="📋" bg={C.amberPale} subColor={C.amber} loading={loading} delay={0.15}/>
      </div>

      {/* Alerte programme */}
      {!loading && stats.nbAlerteProg > 0 && (
        <div style={{display:"flex",alignItems:"center",gap:14,background:"#fef2f2",border:"1px solid #fecaca",borderRadius:12,padding:"14px 18px"}}>
          <span style={{fontSize:26,flexShrink:0}}>⚠️</span>
          <div>
            <div style={{fontWeight:700,color:"#b91c1c",fontSize:13}}>Programme en retard critique</div>
            <div style={{fontSize:12,color:"#7f1d1d",marginTop:2}}>
              {stats.nbAlerteProg} enseignant{stats.nbAlerteProg>1?"s":""} sous 50% de couverture — action requise.
            </div>
          </div>
          <button onClick={()=>setPage("programme")} style={{marginLeft:"auto",padding:"8px 14px",borderRadius:8,border:"1px solid #fecaca",background:"#fff",color:"#b91c1c",fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit",flexShrink:0}}>
            Voir programmes →
          </button>
        </div>
      )}

      {/* Grille principale */}
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1.4fr 1fr",gap:14}}>

        {/* Couverture par enseignant */}
        <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:16}}>
          <h3 style={{margin:"0 0 4px",fontSize:12.5,fontWeight:700,color:C.txt}}>📊 Couverture programme — enseignants</h3>
          <p style={{margin:"0 0 12px",fontSize:10,color:C.txtMuted}}>Triés du plus urgent au plus avancé</p>
          {loading?<Sk h={200} br={8}/>:stats.tauxParEns.length>0?(
            <div style={{display:"flex",flexDirection:"column",gap:10,maxHeight:380,overflowY:"auto"}}>
              {stats.tauxParEns.map(e=>(
                <div key={e.id} style={{display:"flex",alignItems:"center",gap:10}}>
                  <div style={{width:32,height:32,borderRadius:"50%",background:e.col,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800,color:"#fff",flexShrink:0}}>{e.ini}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:12.5,fontWeight:600,color:C.txt,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{e.nom}</div>
                    <div style={{height:5,borderRadius:3,background:"#e2e8f0",marginTop:4}}>
                      <div style={{height:"100%",borderRadius:3,background:tauCol(e.taux),width:e.taux+"%",transition:"width .4s ease"}}/>
                    </div>
                  </div>
                  <span style={{fontSize:12,fontWeight:800,color:tauCol(e.taux),flexShrink:0,width:36,textAlign:"right"}}>{e.taux}%</span>
                </div>
              ))}
            </div>
          ):<div style={{textAlign:"center",padding:"30px 0",color:C.txtLight,fontSize:11}}>Aucune donnée</div>}
        </div>

        {/* Colonne droite */}
        <div style={{display:"flex",flexDirection:"column",gap:14}}>

          {/* Absences par classe */}
          <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:16}}>
            <h3 style={{margin:"0 0 12px",fontSize:12.5,fontWeight:700,color:C.txt}}>📋 Absences par classe</h3>
            {loading?<Sk h={120} br={8}/>:stats.absParClasse.length>0?(
              <div style={{display:"flex",flexDirection:"column",gap:7}}>
                {stats.absParClasse.map(({cl,n})=>(
                  <div key={cl} style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:11,color:C.txt,flex:1,fontWeight:n>5?700:400}}>{cl}</span>
                    <div style={{width:60,height:4,borderRadius:2,background:"#e2e8f0"}}>
                      <div style={{height:"100%",borderRadius:2,background:n>10?C.red:n>5?C.amber:C.green,width:Math.min(100,n*6)+"%"}}/>
                    </div>
                    <span style={{fontSize:11,fontWeight:800,color:n>10?C.red:n>5?C.amber:C.green,width:20,textAlign:"right"}}>{n}</span>
                  </div>
                ))}
              </div>
            ):<div style={{textAlign:"center",padding:"20px 0",color:C.txtLight,fontSize:11}}>Aucune absence enregistrée</div>}
          </div>

          {/* Classes sans programme */}
          {!loading && stats.classesAlertes.length>0 && (
            <div style={{background:"#fffbeb",borderRadius:12,border:"1px solid #fde68a",padding:14}}>
              <h3 style={{margin:"0 0 8px",fontSize:12,fontWeight:700,color:"#92400e"}}>📭 Classes sans programme saisi</h3>
              <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                {stats.classesAlertes.map(cl=>(
                  <span key={cl} style={{fontSize:11,padding:"3px 10px",borderRadius:12,background:"#fef3c7",color:"#78350f",fontWeight:600}}>{cl}</span>
                ))}
              </div>
            </div>
          )}

          {/* EDT coverage */}
          <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:16}}>
            <h3 style={{margin:"0 0 12px",fontSize:12.5,fontWeight:700,color:C.txt}}>📅 Couverture EDT</h3>
            {loading?<Sk h={60} br={8}/>:(
              <div>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                  <span style={{fontSize:12,color:C.txt}}>{stats.edtCoverage.oui} classe{stats.edtCoverage.oui>1?"s":""} configurée{stats.edtCoverage.oui>1?"s":""}</span>
                  <span style={{fontSize:12,fontWeight:700,color:stats.edtCoverage.non>0?C.red:C.green}}>
                    {stats.edtCoverage.non>0?stats.edtCoverage.non+" sans EDT":"✓ Complet"}
                  </span>
                </div>
                <div style={{height:8,borderRadius:4,background:"#e2e8f0"}}>
                  <div style={{height:"100%",borderRadius:4,background:stats.edtCoverage.non>0?C.amber:C.green,width:(stats.edtCoverage.oui/Math.max(stats.edtCoverage.total,1)*100)+"%",transition:"width .5s ease"}}/>
                </div>
              </div>
            )}
          </div>

        </div>
      </div>
    </div>
  );
}

function genRapportDept(stats, user, trim) {
  const deptNom=DEPARTEMENTS_LIST.find(d=>d.id===user.departement_id)?.nom||"SVTEEHB";
  const periode={ANN:"Année 2025-2026",T1:"Trimestre 1",T2:"Trimestre 2",T3:"Trimestre 3"}[trim]||trim;
  const date=new Date().toLocaleDateString("fr-FR",{day:"2-digit",month:"long",year:"numeric"});
  const rows=stats.enseignants.map(e=>{
    const c=e.taux>=75?"#15803d":e.taux>=50?"#92400e":"#b91c1c";
    const bg=e.taux>=75?"#f0fdf4":e.taux>=50?"#fffbeb":"#fef2f2";
    const lbl=e.taux>=75?"✓ Objectif":e.taux>=50?"En cours":"⚠ Retard";
    return `<tr><td style='padding:8px 12px;font-weight:600'>${e.nom}</td>
      <td style='padding:8px 12px;text-align:center'>${(e.classes||[]).join(", ")||"—"}</td>
      <td style='padding:8px 12px;text-align:center'>${e.fait}</td>
      <td style='padding:8px 12px;text-align:center'>${e.ref}</td>
      <td style='padding:8px 12px;text-align:center;font-weight:800;color:${c}'>${e.taux}%</td>
      <td style='padding:8px 12px;text-align:center'><span style='background:${bg};color:${c};padding:3px 8px;border-radius:8px;font-size:11px;font-weight:700'>${lbl}</span></td>
    </tr>`;
  }).join("");
  const html=`<!DOCTYPE html><html lang='fr'><head><meta charset='UTF-8'><title>Rapport ${deptNom}</title>
  <style>body{font-family:Georgia,serif;max-width:800px;margin:32px auto;color:#1f2937;font-size:13px}
  h1{font-size:18px;font-weight:800;color:#0B4D2C;border-bottom:3px solid #0B4D2C;padding-bottom:8px}
  table{width:100%;border-collapse:collapse;margin-top:16px}
  th{background:#0B4D2C;color:#fff;padding:9px 12px;text-align:left;font-size:11px}
  tr:nth-child(even){background:#f8fafc}
  .kpi{display:inline-block;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:10px 20px;margin:0 8px 8px 0;text-align:center}
  .kv{font-size:22px;font-weight:800;color:#0B4D2C}.kl{font-size:10px;color:#6b7280;margin-top:2px}
  @media print{body{margin:16px}}</style></head><body>
  <h1>Rapport pédagogique — Département ${deptNom}</h1>
  <div style='color:#6b7280;font-size:11px;margin-bottom:20px'>Lycée de Kakatare-Maroua · ${periode} · ${date}<br>Animateur : ${user.nom}</div>
  <div>
    <div class='kpi'><div class='kv'>${stats.enseignants.length}</div><div class='kl'>Enseignants</div></div>
    <div class='kpi'><div class='kv'>${stats.tauxMoyen}%</div><div class='kl'>Couverture moy.</div></div>
    <div class='kpi'><div class='kv'>${stats.totalFait}/${stats.totalRef}</div><div class='kl'>Leçons faites</div></div>
    <div class='kpi'><div class='kv'>${stats.enseignants.filter(e=>e.taux<50).length}</div><div class='kl'>En retard</div></div>
  </div>
  <table><thead><tr><th>Enseignant</th><th style='text-align:center'>Classes</th><th style='text-align:center'>Faites</th><th style='text-align:center'>Prévues</th><th style='text-align:center'>Taux</th><th style='text-align:center'>Statut</th></tr></thead>
  <tbody>${rows}</tbody></table>
  <div style='margin-top:48px;display:grid;grid-template-columns:1fr 1fr;gap:40px'>
    <div style='border-top:1px solid #d1d5db;padding-top:8px;text-align:center;font-size:12px'>Le Proviseur<br><br><br>Signature &amp; Cachet :</div>
    <div style='border-top:1px solid #d1d5db;padding-top:8px;text-align:center;font-size:12px'>L\'Animateur Pédagogique<br><br><br>Signature :</div>
  </div></body></html>`;
  return html;
}

function genPVReunion(dept, user, ordreJour, presents, absents, decisions, stats) {
  const date = new Date().toLocaleDateString("fr-FR",{day:"2-digit",month:"long",year:"numeric"});
  const presentsHtml = presents.length>0
    ? presents.map(function(n){return "<li>"+n+"</li>";}).join("")
    : "<li style=\"color:#9ca3af\">Aucun</li>";
  const absentsHtml = absents.length>0
    ? absents.map(function(n){return "<li>"+n+"</li>";}).join("")
    : "<li style=\"color:#9ca3af\">Aucun</li>";
  const ordreJourHtml = ordreJour.split(String.fromCharCode(10)).filter(function(l){return l.trim();})
    .map(function(l){return "<li>"+l+"</li>";}).join("");
  const decisionsHtml = decisions.split(String.fromCharCode(10)).filter(function(l){return l.trim();})
    .map(function(l){return "<li>"+l+"</li>";}).join("");
  const enRetard = stats.enseignants.filter(function(e){return e.taux<50;});
  const enRetardHtml = enRetard.length>0
    ? enRetard.map(function(e){return "<li>"+e.nom+" - "+e.taux+"% ("+e.fait+"/"+e.ref+" lecons)</li>";}).join("")
    : "<li style=\"color:#15803d\">Aucun enseignant en retard</li>";
  const html = "<!DOCTYPE html><html lang=\"fr\"><head><meta charset=\"UTF-8\"><title>PV Reunion</title>"
    + "<style>body{font-family:Georgia,serif;max-width:800px;margin:32px auto;color:#1f2937;font-size:13px;line-height:1.6}"
    + "h1{font-size:17px;font-weight:800;color:#0B4D2C;border-bottom:3px solid #0B4D2C;padding-bottom:8px;text-align:center}"
    + "h2{font-size:13px;color:#0B4D2C;margin:22px 0 8px;border-left:4px solid #D4AF37;padding-left:8px}"
    + ".meta{text-align:center;color:#6b7280;font-size:11px;margin-bottom:20px}"
    + "ul{margin:4px 0;padding-left:20px}"
    + ".cols{display:grid;grid-template-columns:1fr 1fr;gap:24px}"
    + ".sign{display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:44px;font-size:12px}"
    + ".sign div{border-top:1px solid #d1d5db;padding-top:8px;text-align:center}"
    + "@media print{body{margin:16px}}</style></head><body>"
    + "<h1>PROCES-VERBAL DE REUNION DE DEPARTEMENT</h1>"
    + "<div class=\"meta\">Lycee de Kakatare-Maroua | Departement " + dept + " | Le " + date + "</div>"
    + "<h2>1. Ordre du jour</h2><ul>" + (ordreJourHtml||"<li style=\"color:#9ca3af\">Non precise</li>") + "</ul>"
    + "<h2>2. Presences</h2><div class=\"cols\">"
    + "<div><b>Presents</b><ul>" + presentsHtml + "</ul></div>"
    + "<div><b>Absents</b><ul>" + absentsHtml + "</ul></div></div>"
    + "<h2>3. Point sur la progression pedagogique</h2>"
    + "<p>Couverture moyenne du departement : <b>" + stats.tauxMoyen + "%</b> (" + stats.totalFait + "/" + stats.totalRef + " lecons faites).</p>"
    + "<p><b>Enseignants en retard (couverture &lt; 50%) :</b></p><ul>" + enRetardHtml + "</ul>"
    + "<h2>4. Decisions et recommandations</h2><ul>" + (decisionsHtml||"<li style=\"color:#9ca3af\">Aucune</li>") + "</ul>"
    + "<div class=\"sign\"><div>L'Animateur Pedagogique<br><br><br>Signature :</div>"
    + "<div>Secretaire de seance<br><br><br>Signature :</div></div>"
    + "</body></html>";
  return html;
}

function PVReunionModal({stats, user, onClose, onGenerate}) {
  const deptNom = DEPARTEMENTS_LIST.find(function(d){return d.id===user.departement_id;}).nom;
  const [ordreJour, setOrdreJour] = useState("Point sur la progression pedagogique" + String.fromCharCode(10) + "Difficultes rencontrees" + String.fromCharCode(10) + "Preparation des evaluations");
  const [presentsMap, setPresentsMap] = useState(function(){
    const m = {};
    stats.enseignants.forEach(function(e){ m[e.id] = true; });
    return m;
  });
  const [decisions, setDecisions] = useState("");
  const togglePresent = function(id){ setPresentsMap(function(s){ const n={...s}; n[id]=!n[id]; return n; }); };
  const generer = function(){
    const presents = stats.enseignants.filter(function(e){return presentsMap[e.id];}).map(function(e){return e.nom;});
    const absents = stats.enseignants.filter(function(e){return !presentsMap[e.id];}).map(function(e){return e.nom;});
    const html = genPVReunion(deptNom, user, ordreJour, presents, absents, decisions, stats);
    onGenerate(html);
    onClose();
  };
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"#fff",borderRadius:16,maxWidth:560,width:"100%",maxHeight:"90vh",display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <div style={{padding:"16px 20px",borderBottom:"1px solid #e5e7eb",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
          <div>
            <div style={{fontSize:14,fontWeight:800,color:"#0B4D2C"}}>PV de reunion - {deptNom}</div>
            <div style={{fontSize:10,color:"#6b7280"}}>Reunion mensuelle du departement</div>
          </div>
          <button onClick={onClose} style={{background:"#f3f4f6",border:"none",borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:14}}>X</button>
        </div>
        <div style={{flex:1,overflowY:"auto",padding:"18px 20px",display:"flex",flexDirection:"column",gap:16}}>
          <div>
            <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:6}}>Ordre du jour (un point par ligne)</div>
            <textarea value={ordreJour} onChange={function(e){setOrdreJour(e.target.value);}}
              rows={4} style={{width:"100%",border:"1px solid #d1d5db",borderRadius:8,padding:"8px 10px",fontSize:12,fontFamily:"inherit",resize:"vertical"}}/>
          </div>
          <div>
            <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:6}}>
              Presences ({Object.values(presentsMap).filter(Boolean).length}/{stats.enseignants.length} presents)
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:4,maxHeight:180,overflowY:"auto",border:"1px solid #e5e7eb",borderRadius:8,padding:8}}>
              {stats.enseignants.map(function(e){
                return(
                  <label key={e.id} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 6px",cursor:"pointer",fontSize:12,borderRadius:6,background:presentsMap[e.id]?"#f0fdf4":"transparent"}}>
                    <input type="checkbox" checked={!!presentsMap[e.id]} onChange={function(){togglePresent(e.id);}}/>
                    <span style={{color:presentsMap[e.id]?"#166534":"#6b7280"}}>{e.nom}</span>
                  </label>
                );
              })}
            </div>
          </div>
          <div style={{background:"#f8fafc",borderRadius:10,padding:"10px 12px",border:"1px solid #e5e7eb"}}>
            <div style={{fontSize:10,fontWeight:700,color:"#6b7280",marginBottom:4}}>PROGRESSION (auto-rempli)</div>
            <div style={{fontSize:12,color:"#374151"}}>
              Couverture moyenne : <b style={{color:stats.tauxMoyen>=75?"#15803d":stats.tauxMoyen>=50?"#d97706":"#b91c1c"}}>{stats.tauxMoyen}%</b>
              {" - "}{stats.enseignants.filter(function(e){return e.taux<50;}).length} enseignant(s) en retard
            </div>
          </div>
          <div>
            <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:6}}>Decisions et recommandations</div>
            <textarea value={decisions} onChange={function(e){setDecisions(e.target.value);}}
              rows={4} placeholder="Un point par ligne..."
              style={{width:"100%",border:"1px solid #d1d5db",borderRadius:8,padding:"8px 10px",fontSize:12,fontFamily:"inherit",resize:"vertical"}}/>
          </div>
        </div>
        <div style={{padding:"14px 20px",borderTop:"1px solid #e5e7eb",flexShrink:0}}>
          <button onClick={generer}
            style={{width:"100%",padding:"11px 0",background:"#0B4D2C",color:"#fff",border:"none",borderRadius:10,fontWeight:800,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>
            Generer le PV
          </button>
          <button onClick={()=>{
            const presents = stats.enseignants.filter(function(e){return presentsMap[e.id];}).map(function(e){return e.nom;});
            const absents = stats.enseignants.filter(function(e){return !presentsMap[e.id];}).map(function(e){return e.nom;});
            const html = genPVReunion(deptNom, user, ordreJour, presents, absents, decisions, stats);
            onGenerate(html);
          }}
            style={{width:"100%",padding:"9px 0",marginTop:8,background:"#f9fafb",color:"#374151",border:"1px solid #e5e7eb",borderRadius:10,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>
            👁️ Aperçu avant impression
          </button>
        </div>
      </div>
    </div>
  );
}

// Retire le script d'auto-impression avant affichage en apercu (evite le print automatique)
// ════════════════════════════════════════════════════════════════
// FICHE D'INSPECTION — Générateur HTML imprimable
// ════════════════════════════════════════════════════════════════
function genFicheInspection(data) {
  const {
    etablissement="Lycée de Kakatare-Maroua", animateur="", enseignant="",
    classe="", matiere="", dateVisite="", heureDebut="", heureFin="",
    effectifPresent="",
    obs={}, doc={},
    pointsForts="", pointsAmeliorer="", recommandations="",
    noteSur20="", mention=""
  } = data;

  const row = (label, val) => {
    const v = val===true?"✔ Oui":val===false?"✘ Non":"—";
    const c = val===true?"#16a34a":val===false?"#dc2626":"#6b7280";
    return `<tr><td style="padding:5px 10px;border-bottom:1px solid #f0f0f0;font-size:12px;">${label}</td>
      <td style="padding:5px 10px;border-bottom:1px solid #f0f0f0;text-align:center;font-weight:700;font-size:12px;color:${c};">${v}</td></tr>`;
  };
  const mention_color = {"Très bien":"#16a34a","Bien":"#2563eb","Assez bien":"#7c3aed","Passable":"#d97706","Insuffisant":"#dc2626"}[mention]||"#374151";

  return `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
  <style>
    body{font-family:Arial,sans-serif;color:#1f2937;margin:0;padding:20px;font-size:13px;}
    .header{text-align:center;border-bottom:3px solid #0B4D2C;padding-bottom:12px;margin-bottom:18px;}
    .logo{font-size:22px;font-weight:900;color:#0B4D2C;letter-spacing:1px;}
    .subtitle{font-size:11px;color:#6b7280;margin-top:2px;}
    .titre-fiche{font-size:16px;font-weight:900;color:#0B4D2C;margin:10px 0 4px;text-transform:uppercase;letter-spacing:1px;}
    .grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px;}
    .cell{background:#f8fafc;border-radius:6px;padding:10px 14px;}
    .cell-label{font-size:10px;color:#6b7280;text-transform:uppercase;letter-spacing:.5px;font-weight:700;}
    .cell-val{font-size:13px;font-weight:700;color:#1f2937;margin-top:2px;}
    h3{font-size:12px;font-weight:900;color:#0B4D2C;text-transform:uppercase;letter-spacing:.8px;margin:16px 0 6px;padding-bottom:4px;border-bottom:2px solid #0B4D2C;}
    table{width:100%;border-collapse:collapse;margin-bottom:10px;}
    .note-box{display:flex;align-items:center;gap:20px;background:#f0fdf4;border:2px solid #16a34a;border-radius:10px;padding:14px 20px;margin:14px 0;}
    .note-num{font-size:36px;font-weight:900;color:#0B4D2C;}
    .note-label{font-size:11px;color:#6b7280;}
    .mention{font-size:18px;font-weight:900;color:${mention_color};}
    .textarea-section{background:#fafafa;border:1px solid #e5e7eb;border-radius:8px;padding:12px;margin-bottom:10px;min-height:50px;font-size:12px;line-height:1.6;}
    .signature-area{display:grid;grid-template-columns:1fr 1fr;gap:30px;margin-top:20px;}
    .sig-box{text-align:center;border-top:1px solid #d1d5db;padding-top:10px;font-size:11px;color:#6b7280;}
    @media print{body{padding:10px;}@page{margin:15mm;}}
  </style>
  <script>window.onload=()=>{window.print();}</script>
  </head><body>
  <div class="header">
    <div class="logo">EduPilot Cameroun</div>
    <div class="subtitle">${etablissement}</div>
    <div class="titre-fiche">Fiche d'Inspection Pédagogique</div>
    <div style="font-size:11px;color:#6b7280;">Animateur Pédagogique · Visite de classe</div>
  </div>

  <div class="grid2">
    <div class="cell"><div class="cell-label">Animateur</div><div class="cell-val">${animateur}</div></div>
    <div class="cell"><div class="cell-label">Enseignant inspecté</div><div class="cell-val">${enseignant}</div></div>
    <div class="cell"><div class="cell-label">Classe</div><div class="cell-val">${classe}</div></div>
    <div class="cell"><div class="cell-label">Matière</div><div class="cell-val">${matiere}</div></div>
    <div class="cell"><div class="cell-label">Date de visite</div><div class="cell-val">${dateVisite}</div></div>
    <div class="cell"><div class="cell-label">Horaire · Effectif présent</div><div class="cell-val">${heureDebut} – ${heureFin} &nbsp;|&nbsp; ${effectifPresent} élèves</div></div>
  </div>

  <h3>I. Observation de la séance de cours</h3>
  <table>
    <tr style="background:#f0fdf4;"><th style="text-align:left;padding:6px 10px;font-size:11px;">Critère</th><th style="width:80px;text-align:center;font-size:11px;">Appréciation</th></tr>
    ${row("Tenue vestimentaire correcte", obs.tenue_correcte)}
    ${row("Tableau structuré en 3 parties", obs.tableau_structure)}
    ${row("Plan du cours visible", obs.plan_cours_visible)}
    ${row("Titre de leçon encadré", obs.titre_encadre)}
    ${row("Écriture lisible", obs.ecriture_lisible)}
    ${row("Voix audible par toute la classe", obs.voix_audible)}
    ${row("Niveau de langue adapté aux élèves", obs.langue_adaptee)}
    ${row("Élèves interrogés de façon nominative", obs.eleves_interroges)}
    ${row("Transitions bien menées entre parties", obs.transitions_menees)}
    ${row("Situation d'apprentissage APC présentée", obs.situation_apc)}
    ${row("Tâches traduisant les habiletés/contenus", obs.taches_habiletes)}
    ${row("Trace écrite conforme au programme", obs.trace_ecrite_conforme)}
    ${row("Exercices corrigés avec participation classe", obs.exercices_corriges)}
    ${row("Classe impliquée activement", obs.classe_impliquee)}
  </table>

  <h3>II. Documents administratifs de classe</h3>
  <table>
    <tr style="background:#f0fdf4;"><th style="text-align:left;padding:6px 10px;font-size:11px;">Document</th><th style="width:80px;text-align:center;font-size:11px;">État</th></tr>
    ${row("Cahier de textes tenu à jour", doc.cahier_texte_tenu)}
    ${row("Progression annuelle collée dans le cahier", doc.progression_collee)}
    ${row("Progression respectée", doc.progression_respectee)}
    ${row("Registre de notes tenu correctement", doc.registre_notes_tenu)}
    ${row("Absences renseignées", doc.absences_renseignees)}
    ${row("Fiche pédagogique de leçon présente", doc.fiche_pedago_presente)}
  </table>

  <h3>III. Appréciation globale</h3>
  <div class="note-box">
    <div>
      <div class="note-label">Note attribuée</div>
      <div class="note-num">${noteSur20} <span style="font-size:18px;font-weight:400;color:#6b7280;">/ 20</span></div>
    </div>
    <div>
      <div class="note-label">Mention</div>
      <div class="mention">${mention}</div>
    </div>
  </div>

  <div style="margin-bottom:10px;">
    <div style="font-size:11px;font-weight:700;color:#0B4D2C;margin-bottom:4px;">✅ Points forts</div>
    <div class="textarea-section">${pointsForts||"—"}</div>
  </div>
  <div style="margin-bottom:10px;">
    <div style="font-size:11px;font-weight:700;color:#d97706;margin-bottom:4px;">⚠️ Points à améliorer</div>
    <div class="textarea-section">${pointsAmeliorer||"—"}</div>
  </div>
  <div style="margin-bottom:10px;">
    <div style="font-size:11px;font-weight:700;color:#2563eb;margin-bottom:4px;">📌 Recommandations</div>
    <div class="textarea-section">${recommandations||"—"}</div>
  </div>

  <div class="signature-area">
    <div class="sig-box">
      <div style="height:40px;"></div>
      <div>L'Animateur Pédagogique</div>
      <div style="font-weight:700;margin-top:4px;">${animateur}</div>
    </div>
    <div class="sig-box">
      <div style="height:40px;"></div>
      <div>L'Enseignant</div>
      <div style="font-weight:700;margin-top:4px;">${enseignant}</div>
    </div>
  </div>
  </body></html>`;
}

// ════════════════════════════════════════════════════════════════
// PAGE FICHES D'INSPECTION
// ════════════════════════════════════════════════════════════════
function FicheInspectionPage() {
  const {user, data} = useApp();
  const {isMobile} = useDevice();
  const [fiches, setFiches] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [editFiche, setEditFiche] = useState(null);
  const [previewHtml, setPreviewHtml] = useState(null);
  const [saving, setSaving] = useState(false);
  const [form, setForm] = useState({
    enseignant_id:"", classe:"", matiere:"", date_visite:"", heure_debut:"", heure_fin:"", effectif_present:"",
    obs:{tenue_correcte:null,tableau_structure:null,plan_cours_visible:null,titre_encadre:null,ecriture_lisible:null,
      voix_audible:null,langue_adaptee:null,eleves_interroges:null,transitions_menees:null,
      situation_apc:null,taches_habiletes:null,trace_ecrite_conforme:null,exercices_corriges:null,classe_impliquee:null},
    doc:{cahier_texte_tenu:null,progression_collee:null,progression_respectee:null,registre_notes_tenu:null,absences_renseignees:null,fiche_pedago_presente:null},
    points_forts:"", points_ameliorer:"", recommandations:"", note_sur_20:"", mention:""
  });

  const deptId = user?.departement_id;
  const enseignantsDept = Object.values(data?.users||{}).filter(u=>u.role==="enseignant"&&u.departement_id===deptId);

  useEffect(()=>{
    sb.get("fiches_inspection","?order=date_visite.desc&limit=50").then(rows=>{
      setFiches((rows||[]).filter(r=>r.animateur_id===user?.id));
      setLoading(false);
    }).catch(()=>setLoading(false));
  },[]);

  const resetForm = ()=>setForm({
    enseignant_id:"", classe:"", matiere:"", date_visite:"", heure_debut:"", heure_fin:"", effectif_present:"",
    obs:{tenue_correcte:null,tableau_structure:null,plan_cours_visible:null,titre_encadre:null,ecriture_lisible:null,
      voix_audible:null,langue_adaptee:null,eleves_interroges:null,transitions_menees:null,
      situation_apc:null,taches_habiletes:null,trace_ecrite_conforme:null,exercices_corriges:null,classe_impliquee:null},
    doc:{cahier_texte_tenu:null,progression_collee:null,progression_respectee:null,registre_notes_tenu:null,absences_renseignees:null,fiche_pedago_presente:null},
    points_forts:"", points_ameliorer:"", recommandations:"", note_sur_20:"", mention:""
  });

  const openNew = ()=>{ resetForm(); setEditFiche(null); setShowForm(true); };
  const openEdit = (f)=>{
    setForm({
      enseignant_id:f.enseignant_id, classe:f.classe, matiere:f.matiere,
      date_visite:f.date_visite, heure_debut:f.heure_debut||"", heure_fin:f.heure_fin||"",
      effectif_present:f.effectif_present||"",
      obs:{tenue_correcte:f.obs_tenue_correcte,tableau_structure:f.obs_tableau_structure,
        plan_cours_visible:f.obs_plan_cours_visible,titre_encadre:f.obs_titre_encadre,
        ecriture_lisible:f.obs_ecriture_lisible,voix_audible:f.obs_voix_audible,
        langue_adaptee:f.obs_langue_adaptee,eleves_interroges:f.obs_eleves_interroges,
        transitions_menees:f.obs_transitions_menees,situation_apc:f.obs_situation_apc,
        taches_habiletes:f.obs_taches_habiletes,trace_ecrite_conforme:f.obs_trace_ecrite_conforme,
        exercices_corriges:f.obs_exercices_corriges,classe_impliquee:f.obs_classe_impliquee},
      doc:{cahier_texte_tenu:f.doc_cahier_texte_tenu,progression_collee:f.doc_progression_collée,
        progression_respectee:f.doc_progression_respectee,registre_notes_tenu:f.doc_registre_notes_tenu,
        absences_renseignees:f.doc_absences_renseignées,fiche_pedago_presente:f.doc_fiche_pedago_presente},
      points_forts:f.points_forts||"", points_ameliorer:f.points_ameliorer||"",
      recommandations:f.recommandations||"", note_sur_20:f.note_sur_20||"", mention:f.mention||""
    });
    setEditFiche(f.id); setShowForm(true);
  };

  const handleSave = async ()=>{
    if(!form.enseignant_id||!form.date_visite||!form.classe){alert("Enseignant, classe et date sont obligatoires.");return;}
    setSaving(true);
    const token = window.__svtSessionToken;
    const res = await sb.rpc("submit_fiche_inspection",{
      p_token:token, p_enseignant_id:form.enseignant_id, p_classe:form.classe,
      p_matiere:form.matiere, p_date_visite:form.date_visite,
      p_heure_debut:form.heure_debut, p_heure_fin:form.heure_fin,
      p_effectif_present:parseInt(form.effectif_present)||0,
      p_obs:form.obs, p_doc:form.doc,
      p_points_forts:form.points_forts, p_points_ameliorer:form.points_ameliorer,
      p_recommandations:form.recommandations,
      p_note_sur_20:parseFloat(form.note_sur_20)||null,
      p_mention:form.mention||null,
      p_fiche_id:editFiche||null
    });
    setSaving(false);
    if(res?.ok){
      const rows = await sb.get("fiches_inspection","?order=date_visite.desc&limit=50");
      setFiches((rows||[]).filter(r=>r.animateur_id===user?.id));
      setShowForm(false); resetForm(); setEditFiche(null);
    } else { alert("Erreur : "+(res?.error||"inconnue")); }
  };

  const handlePreview = (f)=>{
    const ens = data?.users?.[f.enseignant_id];
    // Calcul effectif présent depuis absences de l'enseignant ce jour-là
    const totalClasse = (ELEVES_DB[f.classe]||[]).length;
    const absKey = f.enseignant_id+"||"+f.classe;
    const absSeance = (data?.absences?.[absKey]||[]).find(a=>a.seance===f.date_visite);
    const nbAbsents = absSeance ? (absSeance.absents||[]).length : 0;
    const effectifCalcule = totalClasse > 0 ? (totalClasse - nbAbsents) : (f.effectif_present||"");
    const html = genFicheInspection({
      animateur:user?.nom||"", enseignant:ens?.nom||f.enseignant_id,
      classe:f.classe, matiere:f.matiere,
      dateVisite:f.date_visite, heureDebut:f.heure_debut||"", heureFin:f.heure_fin||"",
      effectifPresent:effectifCalcule,
      obs:{tenue_correcte:f.obs_tenue_correcte,tableau_structure:f.obs_tableau_structure,
        plan_cours_visible:f.obs_plan_cours_visible,titre_encadre:f.obs_titre_encadre,
        ecriture_lisible:f.obs_ecriture_lisible,voix_audible:f.obs_voix_audible,
        langue_adaptee:f.obs_langue_adaptee,eleves_interroges:f.obs_eleves_interroges,
        transitions_menees:f.obs_transitions_menees,situation_apc:f.obs_situation_apc,
        taches_habiletes:f.obs_taches_habiletes,trace_ecrite_conforme:f.obs_trace_ecrite_conforme,
        exercices_corriges:f.obs_exercices_corriges,classe_impliquee:f.obs_classe_impliquee},
      doc:{cahier_texte_tenu:f.doc_cahier_texte_tenu,progression_collee:f.doc_progression_collée,
        progression_respectee:f.doc_progression_respectee,registre_notes_tenu:f.doc_registre_notes_tenu,
        absences_renseignees:f.doc_absences_renseignées,fiche_pedago_presente:f.doc_fiche_pedago_presente},
      pointsForts:f.points_forts, pointsAmeliorer:f.points_ameliorer,
      recommandations:f.recommandations, noteSur20:f.note_sur_20, mention:f.mention
    });
    setPreviewHtml(stripAutoPrint(html));
  };

  const OuiNonBtn = ({val, onChange})=>(
    <div style={{display:"flex",gap:6}}>
      {[["Oui",true],["Non",false],["N/A",null]].map(([lbl,v])=>(
        <button key={lbl} onClick={()=>onChange(v)}
          style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",
            borderColor:val===v?"#0B4D2C":"#e5e7eb",
            background:val===v?(v===true?"#dcfce7":v===false?"#fee2e2":"#f3f4f6"):"#fff",
            color:val===v?(v===true?"#15803d":v===false?"#dc2626":"#374151"):"#9ca3af"}}>
          {lbl}
        </button>
      ))}
    </div>
  );

  const obsLabels = [
    ["tenue_correcte","Tenue vestimentaire correcte"],
    ["tableau_structure","Tableau structuré en 3 parties"],
    ["plan_cours_visible","Plan du cours visible"],
    ["titre_encadre","Titre de leçon encadré"],
    ["ecriture_lisible","Écriture lisible"],
    ["voix_audible","Voix audible par toute la classe"],
    ["langue_adaptee","Niveau de langue adapté"],
    ["eleves_interroges","Élèves interrogés nominativement"],
    ["transitions_menees","Transitions bien menées"],
    ["situation_apc","Situation APC présentée"],
    ["taches_habiletes","Tâches/habiletés respectées"],
    ["trace_ecrite_conforme","Trace écrite conforme au programme"],
    ["exercices_corriges","Exercices corrigés avec la classe"],
    ["classe_impliquee","Classe activement impliquée"],
  ];
  const docLabels = [
    ["cahier_texte_tenu","Cahier de textes à jour"],
    ["progression_collee","Progression collée dans le cahier"],
    ["progression_respectee","Progression respectée"],
    ["registre_notes_tenu","Registre de notes tenu"],
    ["absences_renseignees","Absences renseignées"],
    ["fiche_pedago_presente","Fiche pédagogique présente"],
  ];

  const mentions = ["Très bien","Bien","Assez bien","Passable","Insuffisant"];
  const mention_colors = {"Très bien":"#16a34a","Bien":"#2563eb","Assez bien":"#7c3aed","Passable":"#d97706","Insuffisant":"#dc2626"};
  const inp = {width:"100%",border:"1.5px solid #e5e7eb",borderRadius:8,padding:"8px 12px",fontSize:13,fontFamily:"inherit",outline:"none",boxSizing:"border-box"};
  const label = {fontSize:11,fontWeight:700,color:"#374151",textTransform:"uppercase",letterSpacing:".5px",marginBottom:4,display:"block"};

  return (
    <div style={{padding:isMobile?"12px":"24px",maxWidth:900,margin:"0 auto"}}>
      {/* En-tête */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <div>
          <div style={{fontSize:isMobile?16:20,fontWeight:900,color:"#0B4D2C"}}>🔍 Fiches d'inspection</div>
          <div style={{fontSize:12,color:"#6b7280",marginTop:2}}>{fiches.length} visite{fiches.length!==1?"s":""} enregistrée{fiches.length!==1?"s":""}</div>
        </div>
        <button onClick={openNew}
          style={{padding:"10px 18px",borderRadius:10,border:"none",background:"#0B4D2C",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:6}}>
          ＋ Nouvelle inspection
        </button>
      </div>

      {/* Liste des fiches */}
      {loading ? <div style={{textAlign:"center",padding:40,color:"#9ca3af"}}>Chargement…</div> :
      fiches.length===0 ? (
        <div style={{textAlign:"center",padding:60,color:"#9ca3af"}}>
          <div style={{fontSize:40,marginBottom:12}}>🔍</div>
          <div style={{fontSize:15,fontWeight:700}}>Aucune fiche d'inspection</div>
          <div style={{fontSize:13,marginTop:6}}>Cliquez sur "Nouvelle inspection" pour commencer</div>
        </div>
      ) : (
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {fiches.map(f=>{
            const ens = data?.users?.[f.enseignant_id];
            const mc = mention_colors[f.mention]||"#6b7280";
            return (
              <div key={f.id} style={{background:"#fff",borderRadius:12,border:"1px solid #e5e7eb",padding:"14px 18px",display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
                <div style={{width:40,height:40,borderRadius:"50%",background:"#0B4D2C",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:15,flexShrink:0}}>
                  {(ens?.nom||"?").split(" ").map(w=>w[0]).join("").slice(0,2).toUpperCase()}
                </div>
                <div style={{flex:1,minWidth:120}}>
                  <div style={{fontWeight:800,fontSize:14,color:"#1f2937"}}>{ens?.nom||f.enseignant_id}</div>
                  <div style={{fontSize:12,color:"#6b7280"}}>{f.classe} · {f.matiere} · {f.date_visite}</div>
                </div>
                {f.mention && <span style={{background:mc+"22",color:mc,borderRadius:20,padding:"3px 12px",fontSize:11,fontWeight:700}}>{f.mention}</span>}
                {f.note_sur_20!=null && <span style={{fontWeight:900,color:"#0B4D2C",fontSize:16}}>{f.note_sur_20}<span style={{fontSize:12,fontWeight:400,color:"#9ca3af"}}>/20</span></span>}
                <div style={{display:"flex",gap:6}}>
                  <button onClick={()=>handlePreview(f)}
                    style={{padding:"6px 12px",borderRadius:8,border:"1px solid #e5e7eb",background:"#f9fafb",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                    👁 Aperçu
                  </button>
                  <button onClick={()=>openEdit(f)}
                    style={{padding:"6px 12px",borderRadius:8,border:"1px solid #0B4D2C",background:"#fff",color:"#0B4D2C",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                    Modifier
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* FORMULAIRE MODAL */}
      {showForm && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.6)",zIndex:2000,overflowY:"auto",padding:isMobile?"0":"20px"}}>
          <div style={{background:"#fff",borderRadius:isMobile?0:16,maxWidth:680,margin:"0 auto",padding:isMobile?"16px":"28px",minHeight:"100vh"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
              <div style={{fontSize:16,fontWeight:900,color:"#0B4D2C"}}>🔍 {editFiche?"Modifier":"Nouvelle"} fiche d'inspection</div>
              <button onClick={()=>{setShowForm(false);resetForm();setEditFiche(null);}}
                style={{background:"none",border:"none",fontSize:22,cursor:"pointer",color:"#6b7280"}}>✕</button>
            </div>

            {/* Infos générales */}
            <div style={{background:"#f8fafc",borderRadius:10,padding:16,marginBottom:16}}>
              <div style={{fontSize:12,fontWeight:900,color:"#0B4D2C",marginBottom:12}}>INFORMATIONS GÉNÉRALES</div>
              <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:12}}>
                <div>
                  <span style={label}>Enseignant *</span>
                  <select value={form.enseignant_id} onChange={e=>setForm(p=>({...p,enseignant_id:e.target.value}))} style={inp}>
                    <option value="">— Choisir —</option>
                    {enseignantsDept.map(e=><option key={e.id} value={e.id}>{e.nom}</option>)}
                  </select>
                </div>
                <div>
                  <span style={label}>Classe *</span>
                  <select value={form.classe} onChange={e=>{
                    const ens = data?.users?.[form.enseignant_id];
                    setForm(p=>({...p,classe:e.target.value}));
                  }} style={inp}>
                    <option value="">— Choisir —</option>
                    {(data?.users?.[form.enseignant_id]?.classes||[]).map(c=><option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
                <div>
                  <span style={label}>Matière</span>
                  <input value={form.matiere} onChange={e=>setForm(p=>({...p,matiere:e.target.value}))} style={inp} placeholder="ex: SVT"/>
                </div>
                <div>
                  <span style={label}>Date de visite *</span>
                  <input type="date" value={form.date_visite} onChange={e=>setForm(p=>({...p,date_visite:e.target.value}))} style={inp}/>
                </div>
                <div>
                  <span style={label}>Heure début</span>
                  <input type="time" value={form.heure_debut} onChange={e=>setForm(p=>({...p,heure_debut:e.target.value}))} style={inp}/>
                </div>
                <div>
                  <span style={label}>Heure fin</span>
                  <input type="time" value={form.heure_fin} onChange={e=>setForm(p=>({...p,heure_fin:e.target.value}))} style={inp}/>
                </div>
                <div>
                  <span style={label}>Effectif présent <span style={{fontWeight:400,color:"#9ca3af"}}>(calculé auto si vide)</span></span>
                  <input type="number" value={form.effectif_present} onChange={e=>setForm(p=>({...p,effectif_present:e.target.value}))} style={inp} placeholder="Auto depuis absences"/>
                </div>
              </div>
            </div>

            {/* Observation séance */}
            <div style={{background:"#f8fafc",borderRadius:10,padding:16,marginBottom:16}}>
              <div style={{fontSize:12,fontWeight:900,color:"#0B4D2C",marginBottom:12}}>I. OBSERVATION DE LA SÉANCE</div>
              {obsLabels.map(([key,lbl])=>(
                <div key={key} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:"1px solid #f0f0f0"}}>
                  <span style={{fontSize:12,color:"#374151",flex:1,marginRight:12}}>{lbl}</span>
                  <OuiNonBtn val={form.obs[key]} onChange={v=>setForm(p=>({...p,obs:{...p.obs,[key]:v}}))}/>
                </div>
              ))}
            </div>

            {/* Documents */}
            <div style={{background:"#f8fafc",borderRadius:10,padding:16,marginBottom:16}}>
              <div style={{fontSize:12,fontWeight:900,color:"#0B4D2C",marginBottom:12}}>II. DOCUMENTS ADMINISTRATIFS</div>
              {docLabels.map(([key,lbl])=>(
                <div key={key} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:"1px solid #f0f0f0"}}>
                  <span style={{fontSize:12,color:"#374151",flex:1,marginRight:12}}>{lbl}</span>
                  <OuiNonBtn val={form.doc[key]} onChange={v=>setForm(p=>({...p,doc:{...p.doc,[key]:v}}))}/>
                </div>
              ))}
            </div>

            {/* Appréciation */}
            <div style={{background:"#f8fafc",borderRadius:10,padding:16,marginBottom:16}}>
              <div style={{fontSize:12,fontWeight:900,color:"#0B4D2C",marginBottom:12}}>III. APPRÉCIATION GLOBALE</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:12}}>
                <div>
                  <span style={label}>Note / 20</span>
                  <input type="number" min="0" max="20" step="0.5" value={form.note_sur_20}
                    onChange={e=>setForm(p=>({...p,note_sur_20:e.target.value}))} style={inp} placeholder="ex: 14.5"/>
                </div>
                <div>
                  <span style={label}>Mention</span>
                  <select value={form.mention} onChange={e=>setForm(p=>({...p,mention:e.target.value}))} style={inp}>
                    <option value="">— Mention —</option>
                    {mentions.map(m=><option key={m} value={m}>{m}</option>)}
                  </select>
                </div>
              </div>
              <div style={{marginBottom:10}}>
                <span style={label}>✅ Points forts</span>
                <textarea value={form.points_forts} onChange={e=>setForm(p=>({...p,points_forts:e.target.value}))}
                  style={{...inp,height:70,resize:"vertical"}} placeholder="Points positifs observés…"/>
              </div>
              <div style={{marginBottom:10}}>
                <span style={label}>⚠️ Points à améliorer</span>
                <textarea value={form.points_ameliorer} onChange={e=>setForm(p=>({...p,points_ameliorer:e.target.value}))}
                  style={{...inp,height:70,resize:"vertical"}} placeholder="Aspects à améliorer…"/>
              </div>
              <div>
                <span style={label}>📌 Recommandations</span>
                <textarea value={form.recommandations} onChange={e=>setForm(p=>({...p,recommandations:e.target.value}))}
                  style={{...inp,height:70,resize:"vertical"}} placeholder="Recommandations à l'enseignant…"/>
              </div>
            </div>

            <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
              <button onClick={()=>{setShowForm(false);resetForm();setEditFiche(null);}}
                style={{padding:"10px 20px",borderRadius:10,border:"1px solid #e5e7eb",background:"#f9fafb",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                Annuler
              </button>
              <button onClick={handleSave} disabled={saving}
                style={{padding:"10px 24px",borderRadius:10,border:"none",background:"#0B4D2C",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit",opacity:saving?.6:1}}>
                {saving?"Enregistrement…":"💾 Enregistrer"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* APERÇU PDF */}
      {previewHtml && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.6)",zIndex:2100,display:"flex",flexDirection:"column",padding:16}}>
          <div style={{background:"#fff",borderRadius:12,flex:1,display:"flex",flexDirection:"column",overflow:"hidden",maxWidth:900,margin:"0 auto",width:"100%"}}>
            <div style={{padding:"12px 18px",borderBottom:"1px solid #e5e7eb",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
              <div style={{fontSize:13,fontWeight:800,color:"#0B4D2C"}}>📄 Fiche d'inspection</div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>imprimerHTML(previewHtml)}
                  style={{padding:"7px 14px",borderRadius:8,border:"none",background:"#0B4D2C",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                  Imprimer
                </button>
                <button onClick={()=>setPreviewHtml(null)}
                  style={{padding:"7px 14px",borderRadius:8,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#374151",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                  Fermer
                </button>
              </div>
            </div>
            <div style={{flex:1,overflow:"hidden",background:"#e5e7eb",padding:12}}>
              <iframe srcDoc={previewHtml} title="apercu-inspection"
                style={{width:"100%",height:"100%",border:"none",borderRadius:8,background:"#fff"}}/>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function stripAutoPrint(html) {
  let h = (html||"").replace(/<script>window\.onload=\(\)=>\{?window\.print\(\);\}?<\/script>/g, "");
  // Injecte un script de mise a l'echelle automatique pour lisibilite sur mobile
  const fitScript = "<script>(function(){function fit(){var b=document.body;if(!b)return;b.style.transform=\"\";b.style.width=\"\";var w=b.scrollWidth;var vw=window.innerWidth;if(w>vw){var s=vw/w;b.style.transformOrigin=\"top left\";b.style.transform=\"scale(\"+s+\")\";b.style.width=(100/s)+\"%\";}}window.addEventListener(\"load\",fit);window.addEventListener(\"resize\",fit);setTimeout(fit,50);})();</script>";
  if (h.includes("</body>")) { h = h.replace("</body>", fitScript + "</body>"); }
  else { h += fitScript; }
  return h;
}

function DocumentsAnimateurPage() {
  const {user, data} = useApp();
  const {isMobile} = useDevice();
  const [stats, setStats] = useState({enseignants:[],tauxMoyen:0,totalFait:0,totalRef:0});
  const [loading, setLoading] = useState(true);
  const [selTrimAnim, setSelTrimAnim] = useState("ANN");
  const [showPV, setShowPV] = useState(false);
  const [previewHtml, setPreviewHtml] = useState(null);
  const [previewLabel, setPreviewLabel] = useState("");

  useEffect(()=>{
    if(!data||!user) return;
    const deptId = user.departement_id;
    const enseignants = Object.values(data.users||{}).filter(u=>
      u.departement_id===deptId && u.role==="enseignant"
    );
    const ensAvecStats = enseignants.map(u=>{
      const classes = (u.classes||[]).filter(Boolean);
      let fait=0, ref=0;
      classes.forEach(cl=>{
        const key=u.id+"||"+cl;
        const done=(data.prog?.[key]||[]).length;
        const code=resolveProgCode(cl);
        const meta=code?PROG_META[code]:null;
        fait+=done; ref+=meta?.lpRef||0;
      });
      const taux=ref>0?Math.min(100,Math.round(fait/ref*100)):0;
      return {...u,fait,ref,taux,nbClasses:classes.length};
    }).sort((a,b)=>a.taux-b.taux);
    const tauxMoyen = ensAvecStats.length>0
      ? Math.round(ensAvecStats.reduce((s,e)=>s+e.taux,0)/ensAvecStats.length)
      : 0;
    const totalFait = ensAvecStats.reduce((s,e)=>s+e.fait,0);
    const totalRef  = ensAvecStats.reduce((s,e)=>s+e.ref,0);
    setStats({enseignants:ensAvecStats,tauxMoyen,totalFait,totalRef});
    setLoading(false);
  },[data,user]);

  return(
    <div style={{padding:"20px 20px 40px",display:"flex",flexDirection:"column",gap:18}}>
      <div>
        <h2 style={{fontSize:18,fontWeight:800,color:C.txt,margin:0}}>Documents à produire</h2>
        <p style={{color:C.txtMuted,margin:"3px 0 0",fontSize:12}}>Fiches, rapports et PV de réunion du département</p>
      </div>
      <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:18}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:8}}>
          <h3 style={{margin:0,fontSize:13,fontWeight:700,color:C.txt}}>
            📄 Documents à produire
          </h3>
          <div style={{display:"flex",gap:6}}>
            {["ANN","T1","T2","T3"].map(t=>(
              <button key={t} onClick={()=>setSelTrimAnim(t)}
                style={{padding:"5px 12px",borderRadius:8,border:`1.5px solid ${selTrimAnim===t?C.green:C.border}`,
                  background:selTrimAnim===t?C.greenPale:C.white,color:selTrimAnim===t?C.green:C.txtMuted,
                  fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                {t==="ANN"?"Année":t}
              </button>
            ))}
          </div>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {stats.enseignants.map(e=>{
            const deptNom=DEPARTEMENTS_LIST.find(d=>d.id===user.departement_id)?.nom||"SVTEEHB";
            return(
              <div key={e.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",
                background:"#f8fafc",borderRadius:10,border:`1px solid ${C.border}`,flexWrap:"wrap"}}>
                <EnsAvatarInfo e={e} subtitle={`${e.nbClasses} classe${e.nbClasses>1?"s":""} · ${e.taux}% couverture`}/>
                <div style={{display:"flex",gap:6,flexShrink:0,
                  width:isMobile?"100%":"auto",justifyContent:isMobile?"flex-end":"flex-start",
                  marginTop:isMobile?6:0}}>
                  <button onClick={()=>{
                    const html=genFicheSuivi(e,e.classes||[],(data?.prog||{}),selTrimAnim,(data?.notes||{}),(data?.absences||{}),deptNom,user.nom||"—");
                    imprimerHTML(html);
                  }}
                    style={{padding:"7px 12px",borderRadius:8,border:"none",background:C.green,color:"#fff",
                      fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flexShrink:0,whiteSpace:"nowrap"}}>
                    📄 Fiche suivi
                  </button>
                  <button onClick={()=>{
                    const html=genFicheSuivi(e,e.classes||[],(data?.prog||{}),selTrimAnim,(data?.notes||{}),(data?.absences||{}),deptNom,user.nom||"—");
                    setPreviewHtml(stripAutoPrint(html)); setPreviewLabel("Fiche suivi — "+e.nom);
                  }}
                    style={{padding:"7px 10px",borderRadius:8,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#374151",
                      fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flexShrink:0}}>
                    👁️
                  </button>
                </div>
              </div>
            );
          })}
          <div style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",
            background:"#f0fdf4",borderRadius:10,border:`1px solid ${C.greenBorder}`,marginTop:4}}>
            <span style={{fontSize:22,flexShrink:0}}>📊</span>
            <div style={{flex:1}}>
              <div style={{fontSize:12.5,fontWeight:700,color:C.txt}}>Rapport trimestriel département</div>
              <div style={{fontSize:10,color:C.txtMuted}}>
                {stats.enseignants.length} enseignants · {stats.tauxMoyen}% moy. · {selTrimAnim==="ANN"?"Année":selTrimAnim}
              </div>
            </div>
            <button onClick={()=>{ const html=genRapportDept(stats,user,selTrimAnim); imprimerHTML(html); }}
              style={{padding:"7px 12px",borderRadius:8,border:"none",background:"#D4AF37",color:"#0B3D20",
                fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flexShrink:0}}>
              📥 Exporter PDF
            </button>
            <button onClick={()=>{ const html=genRapportDept(stats,user,selTrimAnim); setPreviewHtml(stripAutoPrint(html)); setPreviewLabel("Rapport trimestriel département"); }}
              style={{padding:"7px 10px",borderRadius:8,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#374151",
                fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flexShrink:0}}>
              👁️
            </button>
          </div>

          {/* PV de reunion */}
          <div style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",
            background:"#eff6ff",borderRadius:10,border:"1px solid #bfdbfe",marginTop:8}}>
            <span style={{fontSize:22,flexShrink:0}}>📝</span>
            <div style={{flex:1}}>
              <div style={{fontSize:12.5,fontWeight:700,color:C.txt}}>PV de reunion de departement</div>
              <div style={{fontSize:10,color:C.txtMuted}}>Reunion mensuelle - ordre du jour, presents, progression, decisions</div>
            </div>
            <button onClick={()=>setShowPV(true)}
              style={{padding:"7px 14px",borderRadius:8,border:"none",background:"#3b82f6",color:"#fff",
                fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flexShrink:0}}>
              Rediger
            </button>
          </div>
        </div>
      </div>

            {showPV && (
        <PVReunionModal stats={stats} user={user} onClose={()=>setShowPV(false)}
          onGenerate={(html)=>{ setPreviewHtml(stripAutoPrint(html)); setPreviewLabel("PV de réunion de département"); }}/>
      )}
      {previewHtml && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.6)",zIndex:2100,display:"flex",flexDirection:"column",padding:16}}>
          <div style={{background:"#fff",borderRadius:12,flex:1,display:"flex",flexDirection:"column",overflow:"hidden",maxWidth:900,margin:"0 auto",width:"100%"}}>
            <div style={{padding:"12px 18px",borderBottom:"1px solid #e5e7eb",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
              <div style={{fontSize:13,fontWeight:800,color:"#0B4D2C"}}>📄 {previewLabel}</div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>imprimerHTML(previewHtml)}
                  style={{padding:"7px 14px",borderRadius:8,border:"none",background:"#0B4D2C",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                  Imprimer
                </button>
                <button onClick={()=>{setPreviewHtml(null);setPreviewLabel("");}}
                  style={{padding:"7px 14px",borderRadius:8,border:"1px solid #e5e7eb",background:"#f9fafb",color:"#374151",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                  Fermer
                </button>
              </div>
            </div>
            <div style={{flex:1,overflow:"hidden",background:"#e5e7eb",padding:12}}>
              <iframe srcDoc={previewHtml} title="apercu"
                style={{width:"100%",height:"100%",border:"none",borderRadius:8,background:"#fff"}}/>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// Composant partage : avatar + nom + sous-titre, mobile-safe (troncature garantie)
function EnsAvatarInfo({e, subtitle, size}) {
  return (
    <React.Fragment>
      <Avatar ens={e} size={size||32} fontSize={11}/>
      <div style={{flex:1,minWidth:0,overflow:"hidden"}}>
        <div style={{fontSize:12,fontWeight:700,color:C.txt,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{e.nom}</div>
        <div style={{fontSize:9.5,color:C.txtMuted,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{subtitle}</div>
      </div>
    </React.Fragment>
  );
}

function DashboardAnimateur() {
  const {user,data} = useApp();
  const [stats, setStats] = useState({enseignants:[],tauxMoyen:0,epAttente:0,absWeek:0,totalFait:0,totalRef:0});
  const [selTrimAnim, setSelTrimAnim] = useState("ANN");
  const [showPV, setShowPV] = useState(false);
  const [loading, setLoading] = useState(true);

  useEffect(()=>{
    if(!data||!user) return;
    const deptId = user.departement_id;
    const deptNom = DEPARTEMENTS_LIST.find(d=>d.id===deptId)?.nom||"";

    // Enseignants du département
    const enseignants = Object.values(data.users||{}).filter(u=>
      u.departement_id===deptId && u.role==="enseignant"
    );

    // Taux couverture par enseignant
    const ensAvecStats = enseignants.map(u=>{
      const classes = (u.classes||[]).filter(Boolean);
      let fait=0, ref=0;
      classes.forEach(cl=>{
        const key=u.id+"||"+cl;
        const done=(data.prog?.[key]||[]).length;
        const code=resolveProgCode(cl);
        const meta=code?PROG_META[code]:null;
        fait+=done; ref+=meta?.lpRef||0;
      });
      const taux=ref>0?Math.min(100,Math.round(fait/ref*100)):0;
      return {...u,fait,ref,taux,nbClasses:classes.length};
    }).sort((a,b)=>a.taux-b.taux);

    const tauxMoyen = ensAvecStats.length>0
      ? Math.round(ensAvecStats.reduce((s,e)=>s+e.taux,0)/ensAvecStats.length)
      : 0;
    const totalFait = ensAvecStats.reduce((s,e)=>s+e.fait,0);
    const totalRef  = ensAvecStats.reduce((s,e)=>s+e.ref,0);

    // Épreuves en attente de validation du département
    const epAttente = (data.epreuves||[]).filter(e=>
      enseignants.some(u=>u.id===e.ens_id) && e.statut==="attente"
    ).length;

    // Absences cette semaine dans les classes du département
    const today = new Date();
    const lun = new Date(today); lun.setDate(today.getDate()-(today.getDay()||7)+1);
    const lunStr = lun.toISOString().slice(0,10);
    let absWeek=0;
    Object.entries(data.absences||{}).forEach(([k,abs])=>{
      const [,cl,date]=k.split("||");
      if(date>=lunStr && enseignants.some(u=>(u.classes||[]).includes(cl)))
        absWeek+=(abs||[]).length;
    });

    setStats({enseignants:ensAvecStats,tauxMoyen,epAttente,absWeek,totalFait,totalRef,deptNom});
    setLoading(false);
  },[data,user]);

  const tauCol=t=>t>=75?C.green:t>=50?C.amber:C.red;
  const tauBg=t=>t>=75?C.greenPale:t>=50?"#fffbeb":"#fef2f2";

  return(
    <div style={{padding:"20px 20px 40px",display:"flex",flexDirection:"column",gap:18}}>

      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:8}}>
        <div>
          <h2 style={{fontSize:20,fontWeight:800,color:C.txt,margin:0}}>
            Bonjour, {getNomCourt(user?.nom)} 👋
          </h2>
          <p style={{color:C.txtMuted,margin:"3px 0 0",fontSize:12}}>
            Animateur pédagogique · Département {stats.deptNom||""}
          </p>
        </div>
        <div style={{display:"inline-flex",alignItems:"center",gap:5,background:C.greenPale,border:`1px solid ${C.greenBorder}`,borderRadius:7,padding:"4px 10px",fontSize:11,fontWeight:600,color:C.green}}>
          <span style={{width:5,height:5,borderRadius:"50%",background:C.green}}/>Synchronisé
        </div>
      </div>

      {/* KPIs */}
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <KpiCard label="Enseignants" value={stats.enseignants.length} sub="dans mon département" iconEmoji="👥" bg={C.greenPale} loading={loading} delay={0}/>
        <KpiCard label="Couverture moyenne" value={`${stats.tauxMoyen}%`} sub={stats.tauxMoyen>=75?"Objectif atteint ✓":"Sous l'objectif"} subColor={tauCol(stats.tauxMoyen)} iconEmoji="📊" bg={tauBg(stats.tauxMoyen)} loading={loading} delay={0.05}/>
        <KpiCard label="Leçons faites" value={stats.totalFait} sub={`sur ${stats.totalRef} prévues`} iconEmoji="✅" bg={C.bluePale} subColor={C.blue} loading={loading} delay={0.1}/>
        <KpiCard label="Épreuves en attente" value={stats.epAttente} sub="à valider" iconEmoji="📋" bg={stats.epAttente>0?"#fff7ed":C.greenPale} subColor={stats.epAttente>0?C.amber:C.green} loading={loading} delay={0.15}/>
        <KpiCard label="Absences cette semaine" value={stats.absWeek} sub="dans le département" iconEmoji="⚠️" bg={stats.absWeek>5?"#fef2f2":C.greenPale} subColor={stats.absWeek>5?C.red:C.green} loading={loading} delay={0.2}/>
      </div>

      {/* Tableau enseignants */}
      <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:18}}>
        <h3 style={{margin:"0 0 14px",fontSize:13,fontWeight:700,color:C.txt}}>
          📚 Enseignants du département — suivi programme
        </h3>
        {loading?[1,2,3].map(i=><Sk key={i} h={72} br={9} style={{marginBottom:8}}/>)
        :stats.enseignants.length===0?(
          <div style={{textAlign:"center",padding:"32px 0",color:C.txtLight}}>
            <div style={{fontSize:32,marginBottom:8}}>📭</div>
            Aucun enseignant dans ce département
          </div>
        ):(
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {stats.enseignants.map((e,i)=>(
              <div key={e.id} style={{display:"flex",alignItems:"center",gap:12,padding:"12px 14px",
                background:i%2===0?"#f8fafc":C.white,borderRadius:10,border:`1px solid ${C.border}`}}>
                <Avatar ens={e} size={36} fontSize={12}/>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,marginBottom:5}}>
                    <div style={{minWidth:0,overflow:"hidden",flex:1}}>
                      <div style={{fontSize:12,fontWeight:700,color:C.txt,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{e.nom}</div>
                      <div style={{fontSize:9.5,color:C.txtMuted,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{e.nbClasses} cl. · {e.fait}/{e.ref} leç.</div>
                    </div>
                    <div style={{textAlign:"right",flexShrink:0}}>
                      <div style={{fontSize:16,fontWeight:800,color:tauCol(e.taux)}}>{e.taux}%</div>
                      <span style={{fontSize:9,fontWeight:700,padding:"2px 7px",borderRadius:8,
                        background:tauBg(e.taux),color:tauCol(e.taux)}}>
                        {e.taux>=75?"✓ Objectif":e.taux>=50?"En cours":"⚠ Retard"}
                      </span>
                    </div>
                  </div>
                  <ProgBar value={e.taux}/>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {showPV && (
        <PVReunionModal stats={stats} user={user} onClose={()=>setShowPV(false)}/>
      )}
    </div>
  );
}

function DashboardTeacher() {
  const {user,data} = useApp();
  const [loading,setLoading] = useState(true);
  const [myStats,setMyStats] = useState({classes:[],tauxGlobal:0,totalFait:0,totalRef:0});
  useEffect(()=>{
    if(!data||!user)return;
    const prog=data.prog||{};
    const classes=(user.classes||[]).filter(Boolean).map(cl=>{
      const key=`${user.id}||${cl}`;const done=(prog[key]||[]).length;
      const code=resolveProgCode(cl);const meta=code?PROG_META[code]:null;
      const ref=meta?.lpRef||0;const taux=ref>0?Math.min(100, Math.round(done/ref*100)):0;
      const ef=CLASSES_REELLES.find(c=>c.code===cl)?.effectif || (data.classes||[]).find(c=>c.code===cl)?.effectif || 0;
      return{cl,clDisplay:displayCl(cl),done,ref,taux,ef,vh:meta?.vh||0};
    });
    const totalFait=classes.reduce((s,c)=>s+c.done,0);
    const totalRef=classes.reduce((s,c)=>s+c.ref,0);
    setMyStats({classes,tauxGlobal:totalRef>0?Math.min(100, Math.round(totalFait/totalRef*100)):0,totalFait,totalRef});
    setLoading(false);
  },[data,user]);
  const tauCol=t=>t>=75?C.green:t>=50?C.amber:C.red;
  return(
    <div style={{padding:"20px 20px 40px",display:"flex",flexDirection:"column",gap:18}}>
      <div style={{display:"flex",justifyContent:"space-between"}}>
        <div><h2 style={{fontSize:20,fontWeight:800,color:C.txt,margin:0}}>Bonjour, {getNomCourt(user?.nom)} 👋</h2><p style={{color:C.txtMuted,margin:"3px 0 0",fontSize:12}}>{new Date().toLocaleDateString("fr-FR",{weekday:"long",day:"numeric",month:"long"})}</p></div>
        <div style={{display:"inline-flex",alignItems:"center",gap:5,background:C.greenPale,border:`1px solid ${C.greenBorder}`,borderRadius:7,padding:"4px 10px",fontSize:11,fontWeight:600,color:C.green,height:"fit-content"}}><span style={{width:5,height:5,borderRadius:"50%",background:C.green}}/>Synchronisé</div>
      </div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <KpiCard label="Mes classes" value={myStats.classes.length} sub={`${myStats.classes.reduce((s,c)=>s+c.ef,0)} élèves`} iconEmoji="📚" bg={C.greenPale} loading={loading} delay={0}/>
        <KpiCard label="Leçons dispensées" value={myStats.totalFait} sub={`sur ${myStats.totalRef} prévues`} iconEmoji="✅" bg={C.bluePale} subColor={C.blue} loading={loading} delay={0.05}/>
        <KpiCard label="Couverture globale" value={`${myStats.tauxGlobal}%`} sub={myStats.tauxGlobal>=75?"Objectif atteint ✓":"Sous l'objectif"} subColor={tauCol(myStats.tauxGlobal)} iconEmoji="📊" bg={C.greenPale} loading={loading} delay={0.1}/>
      </div>
      <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:18}}>
        <h3 style={{margin:"0 0 14px",fontSize:13,fontWeight:700,color:C.txt}}>📚 Mes classes — suivi en direct</h3>
        {loading?[1,2,3].map(i=><Sk key={i} h={72} br={9} style={{marginBottom:8}}/>):myStats.classes.length===0?(
          <div style={{textAlign:"center",padding:"32px 0",color:C.txtLight}}><div style={{fontSize:32,marginBottom:8}}>📭</div>Aucune classe assignée</div>
        ):(
          <div style={{display:"flex",flexDirection:"column",gap:9}}>
            {myStats.classes.map((c,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:12,padding:"13px 14px",background:"#f8fafc",borderRadius:10,border:`1px solid ${C.border}`,cursor:"pointer"}}
                onMouseEnter={e=>{e.currentTarget.style.borderColor=C.green;e.currentTarget.style.boxShadow=`0 2px 8px ${C.green}18`;}} onMouseLeave={e=>{e.currentTarget.style.borderColor=C.border;e.currentTarget.style.boxShadow="none";}}>
                <div style={{width:40,height:40,borderRadius:10,background:getColor(user?.id),display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:10,fontWeight:800,flexShrink:0}}>{c.cl.replace(/[àáâäèéêëìíîïòóôöùúûü]/gi,"").substring(0,3).toUpperCase()}</div>
                <div style={{flex:1}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                    <div><div style={{fontSize:13,fontWeight:700,color:C.txt}}>{c.clDisplay||c.cl}</div><div style={{fontSize:10,color:C.txtMuted}}>{c.ef} élèves · {c.vh}h/sem</div></div>
                    <div style={{textAlign:"right"}}><div style={{fontSize:18,fontWeight:800,color:tauCol(c.taux)}}>{c.taux}%</div><div style={{fontSize:9,color:C.txtMuted}}>{c.done}/{c.ref}</div></div>
                  </div>
                  <ProgBar value={c.taux}/>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Page placeholder ─────────────────────────────────────────────
const PlaceholderPage = ({title,emoji}) => (
  <div style={{padding:"40px 24px",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",textAlign:"center",minHeight:"60vh"}}>
    <div style={{fontSize:48,marginBottom:16}}>{emoji}</div>
    <h2 style={{fontSize:18,fontWeight:800,color:C.txt,margin:"0 0 8px"}}>{title}</h2>
    <p style={{color:C.txtMuted,fontSize:13,maxWidth:360,lineHeight:1.6}}>Cette section est disponible via les fichiers individuels de chaque module.</p>
  </div>
);

// ── Sidebar ─────────────────────────────────────────────────────────
const SidebarSG = ({collapsed, setCollapsed}) => {
  const {user, page, setPage, data} = useApp();
  const {isMobile, mobileLandscape} = useDevice();
  const effectiveCollapsed = mobileLandscape ? true : collapsed;

  const sgClasses = user?.classes?.length > 0 ? user.classes : null;
  const niveauLabel = !sgClasses ? "Toute l'école"
    : sgClasses[0]?.startsWith('6') ? '6ème'
    : sgClasses[0]?.startsWith('5') ? '5ème'
    : sgClasses[0]?.startsWith('4') ? '4ème'
    : sgClasses[0]?.startsWith('3') ? '3ème'
    : sgClasses[0]?.startsWith('2') ? '2nde' : '1ère & Tle';

  const niveauColor = sgClasses
    ? (sgClasses[0]?.startsWith('6')?"#3b82f6":sgClasses[0]?.startsWith('5')?"#8b5cf6":sgClasses[0]?.startsWith('4')?"#f59e0b":sgClasses[0]?.startsWith('3')?"#10b981":sgClasses[0]?.startsWith('2')?"#ec4899":"#f97316")
    : "#D4AF37";

  const [classesOpen, setClassesOpen] = useState(false);

  // Badges — calcul depuis data disponible
  const today = new Date().toISOString().slice(0,10);
  let absToday = 0;
  Object.entries(data?.absences||{}).forEach(([k,abs])=>{
    const [,cl,date]=k.split("||");
    if(sgClasses&&!sgClasses.includes(cl))return;
    if(date===today) absToday+=(abs?.length||0);
  });

  const go = (tabId) => {
    if (tabId === null) { window.__sgTab = null; setPage("dashboard"); return; }
    window.__sgTab = tabId;
    window.dispatchEvent(new CustomEvent("sg:tab", { detail: tabId }));
    setPage("dashboard");
  };

  const NAV = [
    {id:"dashboard",   emoji:"🏠", label:"Tableau de bord",      tab:null},
    {id:"absences-sg", emoji:"📋", label:"Appels & Absences",     tab:"absences", badge:absToday>0?absToday:null},
    {id:"retards-sg",  emoji:"⏱️", label:"Billets de retard",     tab:"retards"},
    {id:"sanctions-sg",emoji:"⚠️", label:"Discipline & Sanctions",tab:"sanctions"},
    {id:"rapports-sg", emoji:"📊", label:"Rapports & Synthèses",  tab:"vue"},
  ];

  const activePage = page;
  const [activeTab, setActiveTab_sg] = useState(window.__sgTab||null);
  useEffect(()=>{
    const h = (e) => setActiveTab_sg(e.detail);
    window.addEventListener("sg:tab", h);
    return () => window.removeEventListener("sg:tab", h);
  },[]);
  const isNavActive = (item) => {
    if(item.id==="dashboard") return activePage==="dashboard" && !activeTab;
    return activePage==="dashboard" && activeTab===item.tab;
  };

  const G = "#D4AF37"; // gold accent

  return(
    <aside style={{
      width: effectiveCollapsed ? 56 : 240,
      minWidth: effectiveCollapsed ? 56 : 240,
      background:"#0B3D20",
      display:"flex", flexDirection:"column",
      transition:"width .25s, min-width .25s",
      overflow:"hidden", flexShrink:0,
      borderRight:"1px solid rgba(212,175,55,.12)",
    }}>
      {/* ── HEADER ── */}
      <div style={{padding: effectiveCollapsed?"14px 0":"16px 14px 12px", borderBottom:"1px solid rgba(212,175,55,.15)", flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <img src={LOGO_LYCEE_B64} alt="" width={30} height={30} style={{flexShrink:0,objectFit:"contain",borderRadius:"50%",border:"1.5px solid "+G}}/>
          {!effectiveCollapsed&&(
            <div style={{minWidth:0}}>
              <div style={{fontSize:12.5,fontWeight:800,color:"#fff",lineHeight:1.2}}>Lykama</div>
              <div style={{fontSize:9,color:"rgba(255,255,255,.45)",marginTop:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>Lycée de Kakatare · Maroua</div>
            </div>
          )}
        </div>
        {!effectiveCollapsed&&(
          <div style={{marginTop:10,display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
            <div style={{fontSize:9,fontWeight:700,color:"rgba(255,255,255,.35)",textTransform:"uppercase",letterSpacing:".1em"}}>Année scolaire</div>
            <div style={{fontSize:10,fontWeight:700,color:G,background:"rgba(212,175,55,.12)",borderRadius:6,padding:"3px 8px",border:"1px solid rgba(212,175,55,.25)"}}>2025–2026</div>
          </div>
        )}
      </div>

      {/* ── BADGE NIVEAU ── */}
      {!effectiveCollapsed&&(
        <div style={{padding:"10px 14px",borderBottom:"1px solid rgba(212,175,55,.15)",flexShrink:0,background:"rgba(0,0,0,.15)"}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <div style={{width:8,height:8,borderRadius:"50%",background:niveauColor,flexShrink:0,boxShadow:"0 0 6px "+niveauColor}}/>
            <div style={{flex:1}}>
              <div style={{fontSize:9,color:"rgba(255,255,255,.4)",textTransform:"uppercase",letterSpacing:".1em",fontWeight:700}}>Niveau supervisé</div>
              <div style={{fontSize:14,fontWeight:900,color:"#fff",marginTop:1}}>{niveauLabel}</div>
            </div>
            <div style={{fontSize:9,fontWeight:800,color:niveauColor,background:"rgba(212,175,55,.12)",border:"1px solid rgba(212,175,55,.3)",borderRadius:12,padding:"2px 8px",whiteSpace:"nowrap"}}>{sgClasses?.length||0} classes</div>
          </div>
        </div>
      )}

      {/* ── NAVIGATION ── */}
      <nav style={{flex:1,overflowY:"auto",scrollbarWidth:"none",padding:"8px 0"}}>

        {/* Tableau de bord */}
        <NavItemSG active={activePage==="dashboard"&&!window.__sgTab} collapsed={effectiveCollapsed}
          onClick={()=>{window.__sgTab=null;setActiveTab_sg(null);setPage("dashboard");}}
          emoji="🏠" label="Tableau de bord" gold={G}/>

        {/* Classes (expandable) */}
        {!effectiveCollapsed&&sgClasses&&(
          <div>
            <div onClick={()=>setClassesOpen(o=>!o)}
              style={{display:"flex",alignItems:"center",gap:10,padding:"9px 16px",cursor:"pointer",color:"rgba(255,255,255,.6)",fontSize:13,transition:"all .15s"}}
              onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.06)"}
              onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              <span style={{fontSize:16,flexShrink:0}}>📁</span>
              <span style={{flex:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>Mes classes</span>
              <span style={{fontSize:10,transition:"transform .2s",transform:classesOpen?"rotate(90deg)":"rotate(0deg)"}}>›</span>
            </div>
            {classesOpen&&(
              <div style={{paddingLeft:42,paddingBottom:4}}>
                {sgClasses.map(cl=>(
                  <div key={cl} onClick={()=>setPage("eleves")}
                    style={{padding:"6px 16px 6px 0",fontSize:12,color:"rgba(255,255,255,.55)",cursor:"pointer",borderLeft:"2px solid rgba(212,175,55,.25)",marginLeft:6,paddingLeft:10,marginBottom:2,borderRadius:"0 6px 6px 0",transition:"all .12s"}}
                    onMouseEnter={e=>{e.currentTarget.style.color="#fff";e.currentTarget.style.borderLeftColor=G;}}
                    onMouseLeave={e=>{e.currentTarget.style.color="rgba(255,255,255,.55)";e.currentTarget.style.borderLeftColor="rgba(212,175,55,.25)";}}>
                    {cl}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
        {effectiveCollapsed&&(
          <div title="Mes classes" onClick={()=>setPage("eleves")} style={{display:"flex",alignItems:"center",justifyContent:"center",padding:"10px 0",cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.06)"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
            <span style={{fontSize:16}}>📁</span>
          </div>
        )}

        {/* Séparateur */}
        {!effectiveCollapsed&&<div style={{margin:"6px 14px",height:1,background:"rgba(255,255,255,.08)"}}/>}

        {/* Autres items */}
        {NAV.filter(n=>n.id!=="dashboard").map(item=>(
          <NavItemSG key={item.id}
            active={activePage==="dashboard"&&window.__sgTab===item.tab}
            collapsed={effectiveCollapsed}
            onClick={()=>go(item.tab)}
            emoji={item.emoji} label={item.label} badge={item.badge} gold={G}/>
        ))}
      </nav>

      {/* ── PROFIL / PIED ── */}
      <div style={{borderTop:"1px solid rgba(212,175,55,.15)",flexShrink:0}}>
        {!effectiveCollapsed&&(
          <div style={{padding:"12px 14px"}}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
              <Avatar ens={user} size={34} fontSize={12}/>
              <div style={{minWidth:0,flex:1}}>
                <div style={{fontSize:12,fontWeight:700,color:"#fff",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{user?.nom}</div>
                <div style={{fontSize:9,color:"rgba(255,255,255,.45)",marginTop:1}}>Surveillance Générale · {niveauLabel}</div>
              </div>
            </div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>setPage("settings")}
                style={{flex:1,padding:"7px 0",borderRadius:8,border:"1px solid rgba(255,255,255,.12)",background:"rgba(255,255,255,.06)",color:"rgba(255,255,255,.6)",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>
                ⚙️ Paramètres
              </button>
              <button onClick={()=>{localStorage.removeItem("svt_user");window.location.reload();}}
                style={{flex:1,padding:"7px 0",borderRadius:8,border:"1px solid rgba(180,71,46,.4)",background:"rgba(180,71,46,.12)",color:"#f87171",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>
                🚪 Déconnexion
              </button>
            </div>
          </div>
        )}
        {effectiveCollapsed&&(
          <div style={{padding:"10px 0",display:"flex",flexDirection:"column",alignItems:"center",gap:8}}>
            <div onClick={()=>setPage("settings")} style={{cursor:"pointer",padding:6,borderRadius:8}} title="Paramètres" onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.08)"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              <span style={{fontSize:16}}>⚙️</span>
            </div>
            <Avatar ens={user} size={28} fontSize={10}/>
          </div>
        )}
      </div>

      {/* Toggle collapse */}
      <button onClick={()=>setCollapsed(c=>!c)}
        style={{position:"absolute",top:18,right:effectiveCollapsed?-12:-12,width:22,height:22,borderRadius:"50%",background:"#0B3D20",border:"1.5px solid rgba(212,175,55,.4)",color:G,cursor:"pointer",fontSize:10,display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"0 2px 8px rgba(0,0,0,.3)",zIndex:10}}>
        {effectiveCollapsed?"›":"‹"}
      </button>
    </aside>
  );
};

const NavItemSG = ({active, collapsed, onClick, emoji, label, badge, gold="#D4AF37"}) => (
  <div onClick={onClick} title={collapsed?label:""}
    style={{
      display:"flex", alignItems:"center", gap:10,
      padding: collapsed?"10px 0":"9px 16px",
      justifyContent: collapsed?"center":"flex-start",
      cursor:"pointer",
      background: active?"rgba(212,175,55,.14)":"transparent",
      borderLeft: active?"3px solid "+gold:"3px solid transparent",
      color: active?gold:"rgba(255,255,255,.58)",
      fontSize:13, fontWeight: active?700:400,
      transition:"all .15s",
    }}
    onMouseEnter={e=>{if(!active){e.currentTarget.style.background="rgba(255,255,255,.06)";e.currentTarget.style.color="#fff";}}}
    onMouseLeave={e=>{if(!active){e.currentTarget.style.background="transparent";e.currentTarget.style.color="rgba(255,255,255,.58)";}}}
  >
    <span style={{fontSize:16,flexShrink:0}}>{emoji}</span>
    {!collapsed&&<span style={{flex:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{label}</span>}
    {!collapsed&&badge>0&&(
      <span style={{marginLeft:"auto",fontSize:9,fontWeight:800,background:"#dc2626",color:"#fff",borderRadius:20,padding:"2px 7px",flexShrink:0,minWidth:18,textAlign:"center"}}>{badge}</span>
    )}
  </div>
);

// ── Composant sidebar générique groupé (Proviseur, Censeur, Animateur) ──
const SidebarGrouped = ({groups, role, roleLabel, collapsed, setCollapsed, effectiveCollapsed, nbEpAttente}) => {
  const {user, page, setPage} = useApp();
  const [expandedGroups, setExpandedGroups] = useState({departements:false});
  const toggleExpand = (id) => setExpandedGroups(s=>({...s,[id]:!s[id]}));
  const isActive = (id) => page === id;

  const renderItem = (item) => {
    if (item.expandable) {
      const open = expandedGroups[item.id]||false;
      const groupActive = isActive(item.id);
      return (
        <div key={item.id}>
          <div onClick={()=>toggleExpand(item.id)}
            style={{display:"flex",alignItems:"center",gap:10,
              padding:effectiveCollapsed?"10px 0":"9px 16px",
              justifyContent:effectiveCollapsed?"center":"flex-start",
              cursor:"pointer",fontSize:13,transition:"all .15s",
              color:open||groupActive?"#4ade80":"rgba(255,255,255,.55)",
              fontWeight:open||groupActive?700:400,
              background:groupActive?"rgba(34,197,94,.15)":"transparent",
              borderLeft:groupActive?("3px solid "+C.green):"3px solid transparent"}}
            onMouseEnter={e=>{if(!groupActive)e.currentTarget.style.background="rgba(255,255,255,.06)";}}
            onMouseLeave={e=>{if(!groupActive)e.currentTarget.style.background=groupActive?"rgba(34,197,94,.15)":"transparent";}}>
            <span style={{fontSize:16,flexShrink:0}}>{item.emoji}</span>
            {!effectiveCollapsed&&<>
              <span style={{flex:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{item.label}</span>
              <span style={{fontSize:10,transition:"transform .2s",transform:open?"rotate(90deg)":"rotate(0deg)",opacity:.5}}>›</span>
            </>}
          </div>
          {open&&!effectiveCollapsed&&item.sub&&(
            <div style={{paddingBottom:4}}>
              {item.sub.map((s,i)=>(
                <div key={i} onClick={()=>{window.__deptFilter=s.label;window.dispatchEvent(new CustomEvent("dept:open",{detail:s.label}));setPage(item.id);}}
                  style={{display:"flex",alignItems:"center",gap:8,padding:"6px 16px 6px 36px",
                    fontSize:12.5,cursor:"pointer",color:"rgba(255,255,255,.5)",transition:"all .12s",
                    borderLeft:"2px solid transparent"}}
                  onMouseEnter={e=>{e.currentTarget.style.color="#fff";e.currentTarget.style.borderLeftColor=C.green;}}
                  onMouseLeave={e=>{e.currentTarget.style.color="rgba(255,255,255,.5)";e.currentTarget.style.borderLeftColor="transparent";}}>
                  <span style={{fontSize:13}}>{s.emoji}</span>
                  <span>{s.label}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      );
    }
    const active = isActive(item.id);
    return (
      <div key={item.id} onClick={()=>setPage(item.id)} title={effectiveCollapsed?item.label:""}
        style={{display:"flex",alignItems:"center",gap:10,
          padding:effectiveCollapsed?"10px 0":"9px 16px",
          justifyContent:effectiveCollapsed?"center":"flex-start",
          cursor:"pointer",fontSize:13,transition:"all .15s",
          color:active?"#4ade80":"rgba(255,255,255,.55)",
          fontWeight:active?700:400,
          background:active?"rgba(34,197,94,.15)":"transparent",
          borderLeft:active?("3px solid "+C.green):"3px solid transparent"}}
        onMouseEnter={e=>{if(!active)e.currentTarget.style.background="rgba(255,255,255,.06)";}}
        onMouseLeave={e=>{if(!active)e.currentTarget.style.background="transparent";}}>
        <span style={{fontSize:16,flexShrink:0}}>{item.emoji}</span>
        {!effectiveCollapsed&&<span style={{flex:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{item.label}</span>}
        {!effectiveCollapsed&&item.id==="epreuves"&&nbEpAttente>0&&(
          <span style={{marginLeft:"auto",fontSize:9,fontWeight:800,background:C.red,color:"#fff",borderRadius:20,padding:"1px 6px"}}>{nbEpAttente}</span>
        )}
      </div>
    );
  };

  return (
    <aside style={{width:effectiveCollapsed?56:220,minWidth:effectiveCollapsed?56:220,
      background:C.sidebar,display:"flex",flexDirection:"column",
      transition:"width .25s, min-width .25s",overflow:"hidden",flexShrink:0,
      borderRight:"1px solid rgba(255,255,255,.08)",position:"relative"}}>

      {/* Logo */}
      <div style={{padding:effectiveCollapsed?"14px 0":"14px 16px",display:"flex",alignItems:"center",gap:10,borderBottom:"1px solid rgba(255,255,255,.08)",flexShrink:0}}>
        <img src={LOGO_LYCEE_B64} alt="" width={30} height={30} style={{flexShrink:0,objectFit:"contain",borderRadius:"50%"}}/>
        {!effectiveCollapsed&&(
          <div>
            <div style={{fontSize:12,fontWeight:800,color:"#fff",lineHeight:1}}>Lykama</div>
            <div style={{fontSize:9,color:"rgba(255,255,255,.4)",marginTop:2}}>Lycée de Kakatare · Maroua</div>
          </div>
        )}
      </div>

      {/* Année scolaire */}
      {!effectiveCollapsed&&(
        <div style={{padding:"8px 16px 10px",borderBottom:"1px solid rgba(255,255,255,.08)",flexShrink:0}}>
          <div style={{fontSize:9,fontWeight:700,color:"rgba(255,255,255,.3)",textTransform:"uppercase",letterSpacing:".1em",marginBottom:3}}>VUE D'ENSEMBLE</div>
          <div style={{fontSize:12,fontWeight:700,color:"rgba(255,255,255,.7)",background:"rgba(255,255,255,.08)",borderRadius:7,padding:"5px 10px"}}>2025 – 2026 ▾</div>
        </div>
      )}

      {/* Navigation groupée */}
      <nav style={{flex:1,overflowY:"auto",scrollbarWidth:"none",padding:"4px 0"}}>
        {groups.map((group,gi)=>(
          <div key={gi}>
            {!effectiveCollapsed&&group.section&&(
              <>
                {gi>0&&<div style={{margin:"6px 14px 2px",height:1,background:"rgba(255,255,255,.07)"}}/>}
                <div style={{padding:"8px 16px 3px",fontSize:8.5,fontWeight:800,
                  color:"rgba(255,255,255,.25)",letterSpacing:".12em",textTransform:"uppercase"}}>
                  {group.section}
                </div>
              </>
            )}
            {group.items.map(item=>renderItem(item))}
          </div>
        ))}
      </nav>

      {/* Profil */}
      <div onClick={()=>setPage("settings")}
        style={{borderTop:"1px solid rgba(255,255,255,.08)",padding:effectiveCollapsed?"10px 0":"12px 14px",
          display:"flex",alignItems:"center",gap:9,flexShrink:0,
          justifyContent:effectiveCollapsed?"center":"flex-start",cursor:"pointer",transition:"background .15s"}}
        onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.05)"}
        onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
        <Avatar ens={user} size={32} fontSize={11}/>
        {!effectiveCollapsed&&(
          <div style={{overflow:"hidden"}}>
            <div style={{fontSize:12,fontWeight:700,color:"#fff",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{user?.nom}</div>
            <div style={{fontSize:10,color:"rgba(255,255,255,.4)"}}>{roleLabel} · ⚙️ Paramètres</div>
          </div>
        )}
      </div>

      {/* Toggle collapse */}
      <button onClick={()=>setCollapsed(c=>!c)}
        style={{position:"absolute",top:18,right:-12,width:22,height:22,borderRadius:"50%",
          background:C.sidebar,border:"1.5px solid rgba(255,255,255,.2)",color:"rgba(255,255,255,.6)",
          cursor:"pointer",fontSize:10,display:"flex",alignItems:"center",justifyContent:"center",
          boxShadow:"0 2px 8px rgba(0,0,0,.3)",zIndex:10}}>
        {effectiveCollapsed?"›":"‹"}
      </button>
    </aside>
  );
};

const SidebarProviseur = ({collapsed, setCollapsed, effectiveCollapsed, nbEpAttente}) => {
  const {user, page, setPage} = useApp();
  const [deptOpen, setDeptOpen] = useState(false);
  const G = C.green;
  const isActive = (id) => page === id;

  const NavItemProv = ({item}) => {
    const active = isActive(item.id);
    if (item.expandable) {
      return (
        <div>
          <div onClick={()=>setDeptOpen(o=>!o)}
            style={{display:"flex",alignItems:"center",gap:10,
              padding:effectiveCollapsed?"10px 0":"9px 16px",
              justifyContent:effectiveCollapsed?"center":"flex-start",
              cursor:"pointer",fontSize:13,
              color:deptOpen||page.startsWith("dept-")?"#4ade80":"rgba(255,255,255,.55)",
              fontWeight:deptOpen||page.startsWith("dept-")?700:400,
              background:deptOpen||page.startsWith("dept-")?"rgba(34,197,94,.15)":"transparent",
              borderLeft:deptOpen||page.startsWith("dept-")?("3px solid "+G):"3px solid transparent",
              transition:"all .15s"}}
            onMouseEnter={e=>{if(!deptOpen)e.currentTarget.style.background="rgba(255,255,255,.06)";}}
            onMouseLeave={e=>{if(!deptOpen)e.currentTarget.style.background=page.startsWith("dept-")?"rgba(34,197,94,.15)":"transparent";}}>
            <span style={{fontSize:16,flexShrink:0}}>{item.emoji}</span>
            {!effectiveCollapsed&&<>
              <span style={{flex:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{item.label}</span>
              <span style={{fontSize:10,transition:"transform .2s",transform:deptOpen?"rotate(90deg)":"rotate(0deg)",opacity:.6,flexShrink:0}}>›</span>
            </>}
          </div>
          {deptOpen&&!effectiveCollapsed&&(
            <div style={{paddingBottom:4}}>
              {item.sub.map(s=>(
                <div key={s.id} onClick={()=>setPage("departements")}
                  style={{display:"flex",alignItems:"center",gap:8,
                    padding:"6px 16px 6px 36px",fontSize:12.5,cursor:"pointer",
                    color:isActive(s.id)?"#4ade80":"rgba(255,255,255,.5)",
                    fontWeight:isActive(s.id)?700:400,
                    borderLeft:isActive(s.id)?("2px solid "+G):"2px solid transparent",
                    transition:"all .12s"}}
                  onMouseEnter={e=>e.currentTarget.style.color="#fff"}
                  onMouseLeave={e=>e.currentTarget.style.color=isActive(s.id)?"#4ade80":"rgba(255,255,255,.5)"}>
                  <span style={{fontSize:13}}>{s.emoji}</span>
                  <span>{s.label}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      );
    }
    return (
      <div onClick={()=>setPage(item.id)}
        title={effectiveCollapsed?item.label:""}
        style={{display:"flex",alignItems:"center",gap:10,
          padding:effectiveCollapsed?"10px 0":"9px 16px",
          justifyContent:effectiveCollapsed?"center":"flex-start",
          cursor:"pointer",fontSize:13,
          color:active?"#4ade80":"rgba(255,255,255,.55)",
          fontWeight:active?700:400,
          background:active?"rgba(34,197,94,.15)":"transparent",
          borderLeft:active?("3px solid "+G):"3px solid transparent",
          transition:"all .15s"}}
        onMouseEnter={e=>{if(!active)e.currentTarget.style.background="rgba(255,255,255,.06)";}}
        onMouseLeave={e=>{if(!active)e.currentTarget.style.background="transparent";}}>
        <span style={{fontSize:16,flexShrink:0}}>{item.emoji}</span>
        {!effectiveCollapsed&&<span style={{flex:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{item.label}</span>}
        {!effectiveCollapsed&&item.id==="epreuves"&&nbEpAttente>0&&(
          <span style={{marginLeft:"auto",fontSize:9,fontWeight:800,background:C.red,color:"#fff",borderRadius:20,padding:"1px 6px",flexShrink:0}}>{nbEpAttente}</span>
        )}
      </div>
    );
  };

  return (
    <aside style={{
      width:effectiveCollapsed?56:220,minWidth:effectiveCollapsed?56:220,
      background:C.sidebar,display:"flex",flexDirection:"column",
      transition:"width .25s, min-width .25s",overflow:"hidden",flexShrink:0,
      borderRight:"1px solid rgba(255,255,255,.08)",position:"relative"
    }}>
      {/* Logo */}
      <div style={{padding:effectiveCollapsed?"14px 0":"14px 16px",display:"flex",alignItems:"center",gap:10,borderBottom:"1px solid rgba(255,255,255,.08)",flexShrink:0}}>
        <img src={LOGO_LYCEE_B64} alt="" width={30} height={30} style={{flexShrink:0,objectFit:"contain",borderRadius:"50%"}}/>
        {!effectiveCollapsed&&(
          <div>
            <div style={{fontSize:12,fontWeight:800,color:"#fff",lineHeight:1}}>Lykama</div>
            <div style={{fontSize:9,color:"rgba(255,255,255,.4)",marginTop:2}}>Lycée de Kakatare · Maroua</div>
          </div>
        )}
      </div>

      {/* Année scolaire */}
      {!effectiveCollapsed&&(
        <div style={{padding:"8px 16px",borderBottom:"1px solid rgba(255,255,255,.08)",flexShrink:0}}>
          <div style={{fontSize:9,fontWeight:700,color:"rgba(255,255,255,.35)",textTransform:"uppercase",letterSpacing:".1em",marginBottom:3}}>VUE D'ENSEMBLE</div>
          <div style={{fontSize:12,fontWeight:700,color:"rgba(255,255,255,.7)",background:"rgba(255,255,255,.08)",borderRadius:7,padding:"5px 10px"}}>2025 – 2026 ▾</div>
        </div>
      )}

      {/* Navigation groupée */}
      <nav style={{flex:1,overflowY:"auto",scrollbarWidth:"none",padding:"6px 0"}}>
        {NAV_PROVISEUR_GROUPS.map((group,gi)=>(
          <div key={gi}>
            {!effectiveCollapsed&&gi>0&&(
              <div style={{padding:"10px 16px 4px",fontSize:8.5,fontWeight:800,
                color:"rgba(255,255,255,.25)",letterSpacing:".12em",textTransform:"uppercase"}}>
                {group.section}
              </div>
            )}
            {gi>0&&!effectiveCollapsed&&<div style={{margin:"2px 14px 4px",height:1,background:"rgba(255,255,255,.07)"}}/>}
            {group.items.map(item=><NavItemProv key={item.id} item={item}/>)}
          </div>
        ))}
      </nav>

      {/* Profil */}
      <div onClick={()=>setPage("settings")}
        style={{borderTop:"1px solid rgba(255,255,255,.08)",padding:effectiveCollapsed?"10px 0":"12px 14px",
          display:"flex",alignItems:"center",gap:9,flexShrink:0,
          justifyContent:effectiveCollapsed?"center":"flex-start",cursor:"pointer",transition:"background .15s"}}
        onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.05)"}
        onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
        <Avatar ens={user} size={32} fontSize={11}/>
        {!effectiveCollapsed&&(
          <div style={{overflow:"hidden"}}>
            <div style={{fontSize:12,fontWeight:700,color:"#fff",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{user?.nom}</div>
            <div style={{fontSize:10,color:"rgba(255,255,255,.4)"}}>Proviseur · ⚙️ Paramètres</div>
          </div>
        )}
      </div>

      {/* Toggle */}
      <button onClick={()=>setCollapsed(c=>!c)}
        style={{position:"absolute",top:18,right:-12,width:22,height:22,borderRadius:"50%",
          background:C.sidebar,border:"1.5px solid rgba(255,255,255,.2)",color:"rgba(255,255,255,.6)",
          cursor:"pointer",fontSize:10,display:"flex",alignItems:"center",justifyContent:"center",
          boxShadow:"0 2px 8px rgba(0,0,0,.3)",zIndex:10}}>
        {effectiveCollapsed?"›":"‹"}
      </button>
    </aside>
  );
};

const Sidebar = ({collapsed, setCollapsed}) => {
  const {user, page, setPage, data, t} = useApp();
  const {isMobile, mobileLandscape, isTablet} = useDevice();
  // Forcer collapsed sur mobile/tablette en paysage pour libérer l'espace
  // Le repli initial sur petits écrans est déjà géré par l'état initial de `collapsed`
  // (voir useState(()=>window.innerWidth<1024) dans AppLayout) — ici on ne force PLUS
  // un repli permanent sur tablette, sinon le bouton de la sidebar devient inopérant.
  const effectiveCollapsed = mobileLandscape ? true : collapsed;
  const isAdmin = isAdminRole(user?.role);
  const nav = user?.role==="proviseur" ? NAV_PROVISEUR : user?.role==="surveillant_general" ? NAV_SURVEILLANCE : user?.role==="censeur" ? NAV_CENSEUR : isAdmin ? NAV_ADMIN : NAV_TEACHER;
  // Compter les épreuves en attente pour le badge
  const nbEpAttente = isAdmin
    ? (data?.epreuves||[]).filter(e=>e.statut==="attente").length
    : (data?.epreuves||[]).filter(e=>e.ens_id===user?.id&&e.statut==="attente").length;

  if (user?.role === "proviseur") return <SidebarGrouped groups={NAV_PROVISEUR_GROUPS} role="proviseur" roleLabel="Proviseur" collapsed={collapsed} setCollapsed={setCollapsed} effectiveCollapsed={effectiveCollapsed} nbEpAttente={nbEpAttente}/>;
  if (user?.role === "censeur")   return <SidebarGrouped groups={NAV_CENSEUR_GROUPS}   role="censeur"   roleLabel="Censeur"   collapsed={collapsed} setCollapsed={setCollapsed} effectiveCollapsed={effectiveCollapsed} nbEpAttente={nbEpAttente}/>;
  if (user?.role === "animateur"||user?.role === "animatrice") return <SidebarGrouped groups={NAV_ANIMATEUR_GROUPS} role={user?.role} roleLabel="Animateur Pédagogique" collapsed={collapsed} setCollapsed={setCollapsed} effectiveCollapsed={effectiveCollapsed} nbEpAttente={nbEpAttente}/>;
  if (user?.role === "proviseur") {
    return (
      <SidebarProviseur collapsed={collapsed} setCollapsed={setCollapsed}
        effectiveCollapsed={effectiveCollapsed}
        nbEpAttente={nbEpAttente}/>
    );
  }
  if (user?.role === "surveillant_general") return <SidebarSG collapsed={collapsed} setCollapsed={setCollapsed}/>;

  return (
    <aside style={{
      width: effectiveCollapsed ? 56 : 220,
      minWidth: effectiveCollapsed ? 56 : 220,
      background: C.sidebar,
      display: "flex",
      flexDirection: "column",
      transition: "width .25s, min-width .25s",
      overflow: "hidden",
      flexShrink: 0,
    }}>
      {/* Logo */}
      <div style={{padding: collapsed ? "14px 0" : "14px 16px", display:"flex", alignItems:"center", gap:10, borderBottom:"1px solid rgba(255,255,255,.08)", flexShrink:0}}>
        <img src={LOGO_LYCEE_B64} alt="SVTEEHB" width={30} height={30} style={{flexShrink:0, objectFit:"contain", borderRadius:"50%"}}/>
        {!effectiveCollapsed && (
          <div>
            <div style={{fontSize:12, fontWeight:800, color:"#fff", lineHeight:1}}>Lykama</div>
            <div style={{fontSize:9, color:"rgba(255,255,255,.4)", marginTop:2}}>Lycée de Kakatare · Maroua</div>
          </div>
        )}
      </div>

      {/* Année scolaire */}
      {!effectiveCollapsed && (
        <div style={{padding:"10px 16px", borderBottom:"1px solid rgba(255,255,255,.08)", flexShrink:0}}>
          <div style={{fontSize:9, fontWeight:700, color:"rgba(255,255,255,.35)", textTransform:"uppercase", letterSpacing:".1em", marginBottom:4}}>Année scolaire</div>
          <div style={{fontSize:12, fontWeight:700, color:"rgba(255,255,255,.7)", background:"rgba(255,255,255,.08)", borderRadius:7, padding:"5px 10px"}}>2025 – 2026 ▾</div>
        </div>
      )}

      {/* Navigation */}
      <nav style={{flex:1, overflowY:"auto", scrollbarWidth:"none", padding:"8px 0"}}>
        {nav.map(item => {
          const isActive = page === item.id;
          return (
            <div key={item.id} onClick={()=>setPage(item.id)}
              title={effectiveCollapsed ? t(item.label) : ""}
              style={{
                display:"flex", alignItems:"center", gap:10,
                padding: effectiveCollapsed ? "10px 0" : "9px 16px",
                justifyContent: effectiveCollapsed ? "center" : "flex-start",
                cursor:"pointer",
                background: isActive ? "rgba(34,197,94,.15)" : "transparent",
                borderLeft: isActive ? `3px solid ${C.green}` : "3px solid transparent",
                color: isActive ? "#4ade80" : "rgba(255,255,255,.55)",
                fontSize: 13,
                fontWeight: isActive ? 700 : 400,
                transition: "all .15s",
              }}
              onMouseEnter={e=>{if(!isActive)e.currentTarget.style.background="rgba(255,255,255,.06)";}}
              onMouseLeave={e=>{if(!isActive)e.currentTarget.style.background="transparent";}}>
              <span style={{fontSize:16, flexShrink:0}}>{item.emoji}</span>
              {!collapsed && <span style={{whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis"}}>{t(item.label)}</span>}
              {!collapsed && item.id==="epreuves" && nbEpAttente>0 && (
                <span style={{marginLeft:"auto",fontSize:9,fontWeight:800,background:C.red,color:"#fff",borderRadius:20,padding:"1px 6px",flexShrink:0}}>{nbEpAttente}</span>
              )}
            </div>
          );
        })}
      </nav>

      {/* Profil utilisateur — cliquable vers Paramètres (photo, mot de passe) */}
      <div onClick={()=>setPage("settings")} title="Paramètres — photo de profil, mot de passe"
        style={{borderTop:"1px solid rgba(255,255,255,.08)", padding: collapsed ? "10px 0" : "12px 14px", display:"flex", alignItems:"center", gap:9, flexShrink:0, justifyContent: effectiveCollapsed ? "center" : "flex-start", cursor:"pointer", transition:"background .15s"}}
        onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.05)"}
        onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
        <Avatar ens={user} size={32} fontSize={11}/>
        {!effectiveCollapsed && (
          <div style={{overflow:"hidden"}}>
            <div style={{fontSize:12, fontWeight:700, color:"#fff", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis"}}>{user?.nom}</div>
            <div style={{fontSize:10, color:"rgba(255,255,255,.4)", textTransform:"capitalize"}}>{user?.role==="proviseur"?"Proviseur":user?.role==="surveillant_general"?"Surveillance générale":user?.role==="censeur"?"Censeur":(user?.role==="animateur"||user?.role==="animatrice")?"Animateur pédagogique":"Enseignant"} · ⚙️ Paramètres</div>
          </div>
        )}
      </div>
    </aside>
  );
};


const DarkModeToggle = () => {
  const [dark, setDark] = useDarkMode();
  return (
    <button onClick={()=>setDark(!dark)}
      title={dark?"Mode clair":"Mode sombre"}
      style={{width:32,height:32,borderRadius:8,border:`1px solid ${C.border}`,
        background:C.white,cursor:"pointer",fontSize:14,display:"flex",
        alignItems:"center",justifyContent:"center"}}>
      {dark?"☀️":"🌙"}
    </button>
  );
};

// ── Topbar ──────────────────────────────────────────────────────────
// ── Recherche globale (élèves, enseignants, classes) — admin uniquement ──
const GlobalSearch = () => {
  const {data, setPage, setPendingFicheEns, setPendingClasseSelect, mobileSearchOpen, setMobileSearchOpen} = useApp();
  const {isMobile} = useDevice();
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);       // desktop : dropdown ouvert
  const boxRef = useRef(null);
  const mobileInputRef = useRef(null);

  const openMobileSearch = () => {
    window.history.pushState({modal:"search"}, "", "");
    setMobileSearchOpen(true);
  };
  const closeMobileSearch = () => {
    setQ("");
    window.history.back(); // déclenche popstate → ferme via le handler central
  };

  useEffect(() => {
    const onClickOutside = (e) => { if (boxRef.current && !boxRef.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", onClickOutside);
    return () => document.removeEventListener("mousedown", onClickOutside);
  }, []);

  useEffect(() => {
    if (mobileSearchOpen && mobileInputRef.current) {
      setTimeout(()=>mobileInputRef.current?.focus(), 80);
    }
  }, [mobileSearchOpen]);

  const query = q.trim().toLowerCase();
  const results = useMemo(() => {
    if (query.length < 2) return { eleves:[], enseignants:[], classes:[] };
    const eleves = [];
    Object.keys(ELEVES_DB).forEach(cl => {
      (ELEVES_DB[cl]||[]).forEach(e => {
        if ((e.nom||"").toLowerCase().includes(query)) eleves.push({...e, classe:cl});
      });
    });
    const enseignants = Object.values(data?.users||{})
      .filter(u=>u.role!=="proviseur" && (u.nom||"").toLowerCase().includes(query));
    const classes = Object.keys(ELEVES_DB).filter(cl => cl.toLowerCase().includes(query));
    return { eleves: eleves.slice(0,6), enseignants: enseignants.slice(0,5), classes: classes.slice(0,5) };
  }, [query, data]);

  const totalResults = results.eleves.length + results.enseignants.length + results.classes.length;

  const closeAll = () => {
    setOpen(false);
    setQ("");
    if (mobileSearchOpen) window.history.back(); // dépile l'entrée poussée à l'ouverture
  };
  const goToClasse = (cl) => { setPendingClasseSelect(cl); setPage("eleves"); closeAll(); };
  const goToEns    = (ensId) => { setPendingFicheEns(ensId); setPage("documents"); closeAll(); };

  // Liste de résultats — identique en contenu pour desktop (dropdown) et mobile (plein écran),
  // seule la taille des zones tactiles change.
  const renderResults = (touchSize) => {
    if (query.length < 2) return null;
    if (totalResults===0) return (
      <div style={{padding:"24px", textAlign:"center", color:"#94a3b8", fontSize:touchSize?13:12}}>Aucun résultat pour "{q}"</div>
    );
    const rowPad = touchSize ? "13px 16px" : "8px 14px";
    const rowFont = touchSize ? 14 : 12.5;
    const labelFont = touchSize ? 11 : 9.5;
    return (
      <>
        {results.classes.length>0 && (
          <div>
            <div style={{padding:`10px 16px 4px`, fontSize:labelFont, fontWeight:800, color:"#94a3b8", textTransform:"uppercase"}}>Classes</div>
            {results.classes.map(cl=>(
              <div key={cl} onClick={()=>goToClasse(cl)} style={{padding:rowPad, cursor:"pointer", display:"flex", alignItems:"center", gap:10, fontSize:rowFont, borderBottom:touchSize?`1px solid #f1f5f9`:"none"}}
                onMouseEnter={e=>!touchSize&&(e.currentTarget.style.background="#f8fafc")} onMouseLeave={e=>!touchSize&&(e.currentTarget.style.background="transparent")}>
                <span>🏫</span><span style={{fontWeight:600, color:C.txt}}>{cl}</span>
                <span style={{fontSize:rowFont-2, color:"#94a3b8"}}>· {(ELEVES_DB[cl]||[]).length} élèves</span>
              </div>
            ))}
          </div>
        )}
        {results.enseignants.length>0 && (
          <div>
            <div style={{padding:`10px 16px 4px`, fontSize:labelFont, fontWeight:800, color:"#94a3b8", textTransform:"uppercase"}}>Enseignants</div>
            {results.enseignants.map(ens=>(
              <div key={ens.id} onClick={()=>goToEns(ens.id)} style={{padding:rowPad, cursor:"pointer", display:"flex", alignItems:"center", gap:10, fontSize:rowFont, borderBottom:touchSize?`1px solid #f1f5f9`:"none"}}
                onMouseEnter={e=>!touchSize&&(e.currentTarget.style.background="#f8fafc")} onMouseLeave={e=>!touchSize&&(e.currentTarget.style.background="transparent")}>
                <Avatar ens={ens} size={touchSize?26:22} fontSize={touchSize?9.5:8}/>
                <span style={{fontWeight:600, color:C.txt}}>{getNomCourt(ens.nom)}</span>
                <span style={{fontSize:rowFont-2, color:"#94a3b8"}}>· {(ens.classes||[]).length} classes</span>
              </div>
            ))}
          </div>
        )}
        {results.eleves.length>0 && (
          <div>
            <div style={{padding:`10px 16px 4px`, fontSize:labelFont, fontWeight:800, color:"#94a3b8", textTransform:"uppercase"}}>Élèves</div>
            {results.eleves.map(e=>(
              <div key={e.id} onClick={()=>goToClasse(e.classe)} style={{padding:rowPad, cursor:"pointer", display:"flex", alignItems:"center", gap:10, fontSize:rowFont, borderBottom:touchSize?`1px solid #f1f5f9`:"none"}}
                onMouseEnter={ev=>!touchSize&&(ev.currentTarget.style.background="#f8fafc")} onMouseLeave={ev=>!touchSize&&(ev.currentTarget.style.background="transparent")}>
                <span>{e.g==="F"?"👧":"👦"}</span><span style={{fontWeight:600, color:C.txt}}>{e.nom}</span>
                <span style={{fontSize:rowFont-2, color:"#94a3b8"}}>· {e.classe}</span>
              </div>
            ))}
          </div>
        )}
      </>
    );
  };

  // ── Mobile : icône seule dans le Topbar + overlay plein écran ──────
  if (isMobile) {
    return (
      <>
        <button onClick={openMobileSearch}
          style={{width:32, height:32, borderRadius:8, border:`1px solid ${C.border}`, background:C.white, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", fontSize:14, flexShrink:0}}>
          🔍
        </button>
        {mobileSearchOpen && (
          <div style={{position:"fixed", inset:0, background:"#fff", zIndex:300, display:"flex", flexDirection:"column"}}>
            <div style={{display:"flex", alignItems:"center", gap:10, padding:"12px 14px", borderBottom:`1px solid ${C.border}`, flexShrink:0}}>
              <button onClick={closeMobileSearch} style={{width:34, height:34, borderRadius:8, border:"none", background:"#f1f5f9", fontSize:16, color:C.txtMuted, flexShrink:0}}>←</button>
              <input ref={mobileInputRef} value={q} onChange={e=>setQ(e.target.value)}
                placeholder="Élève, enseignant, classe…"
                style={{flex:1, border:"none", outline:"none", fontSize:15, fontFamily:"inherit", color:C.txt, background:"transparent"}}/>
              {q && <button onClick={()=>setQ("")} style={{border:"none", background:"transparent", color:"#94a3b8", fontSize:16}}>✕</button>}
            </div>
            <div style={{flex:1, overflowY:"auto"}}>
              {query.length<2 ? (
                <div style={{padding:"40px 20px", textAlign:"center", color:"#94a3b8", fontSize:13}}>Tape au moins 2 lettres pour chercher</div>
              ) : renderResults(true)}
            </div>
          </div>
        )}
      </>
    );
  }

  // ── Desktop : champ inline + dropdown ───────────────────────────────
  return (
    <div ref={boxRef} style={{position:"relative", width: open?320:220, transition:"width .2s"}}>
      <div style={{display:"flex", alignItems:"center", gap:7, background:"#f1f5f9", borderRadius:9, padding:"7px 12px", border:`1px solid ${open?C.green:"transparent"}`}}>
        <span style={{fontSize:13, color:"#94a3b8"}}>🔍</span>
        <input value={q} onChange={e=>setQ(e.target.value)} onFocus={()=>setOpen(true)}
          placeholder="Élève, enseignant, classe…"
          style={{flex:1, border:"none", background:"transparent", outline:"none", fontSize:12.5, fontFamily:"inherit", color:C.txt}}/>
        {q && <span onClick={()=>setQ("")} style={{cursor:"pointer", color:"#94a3b8", fontSize:13}}>✕</span>}
      </div>

      {open && query.length>=2 && (
        <div style={{position:"absolute", top:"calc(100% + 6px)", left:0, width:340, maxHeight:380, overflowY:"auto",
          background:"#fff", borderRadius:12, border:`1px solid ${C.border}`, boxShadow:"0 8px 24px rgba(0,0,0,.12)", zIndex:100}}>
          {renderResults(false)}
        </div>
      )}
    </div>
  );
};

const Topbar = ({title, onLogout, collapsed, setCollapsed}) => {
  const {user, realtimeStatus, viewDeptId, setViewDeptId, lang, setLang, t} = useApp();
  const {isMobile} = useDevice();
  const [installPrompt, setInstallPrompt] = useState(null);
  const [installed, setInstalled]         = useState(false);

  const REALTIME_CFG = {
    connected:    { dot:"#22c55e", label:"Synchro en direct" },
    connecting:   { dot:"#f59e0b", label:"Connexion…" },
    error:        { dot:"#ef4444", label:"Synchro indisponible" },
    disconnected: { dot:"#94a3b8", label:"Synchro hors ligne" },
  };
  const rtCfg = REALTIME_CFG[realtimeStatus] || REALTIME_CFG.disconnected;

  useEffect(()=>{
    const handler = e => { e.preventDefault(); setInstallPrompt(e); };
    window.addEventListener('beforeinstallprompt', handler);
    window.addEventListener('appinstalled', ()=>setInstalled(true));
    return()=>window.removeEventListener('beforeinstallprompt', handler);
  },[]);

  const handleInstall = async() => {
    if(!installPrompt) return;
    installPrompt.prompt();
    const result = await installPrompt.userChoice;
    if(result.outcome==='accepted') { setInstalled(true); setInstallPrompt(null); }
  };
  return (
    <header style={{height:52, background:"rgba(255,255,255,.95)", backdropFilter:"blur(12px)", borderBottom:`1px solid ${C.border}`, display:"flex", alignItems:"center", gap:12, padding:"0 18px", position:"sticky", top:0, zIndex:30, flexShrink:0, overflowX:"auto", scrollbarWidth:"none"}}>
      <button onClick={()=>setCollapsed(!collapsed)}
        style={{width:32, height:32, borderRadius:8, border:`1px solid ${C.border}`, background:C.white, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", fontSize:14, color:C.txtMuted}}>
        {collapsed ? "→" : "←"}
      </button>
      <h1 style={{fontSize:14, fontWeight:700, color:C.txt, margin:0}}>{title}</h1>
      {user?.role==="animatrice" && <GlobalSearch/>}
      <div style={{flex:1}}/>
      {!isMobile && (
        <div title={rtCfg.label} style={{display:"flex", alignItems:"center", gap:6, fontSize:10.5, color:C.txtMuted}}>
          <span style={{width:7, height:7, borderRadius:"50%", background:rtCfg.dot, flexShrink:0,
            animation:realtimeStatus==="connecting"?"pulse 1.2s infinite":"none"}}/>
          {rtCfg.label}
        </div>
      )}
      <div style={{fontSize:11, color:C.txtMuted}}>
        {new Date().toLocaleDateString("fr-FR",{weekday:"short",day:"numeric",month:"short"})}
      </div>
      {installPrompt && !installed && (
        <button onClick={handleInstall}
          style={{padding:"5px 12px",borderRadius:8,border:`1px solid ${C.green}`,
            background:C.greenPale,fontSize:11,fontWeight:700,cursor:"pointer",
            color:C.green,fontFamily:"inherit",display:"flex",alignItems:"center",gap:5}}>
          📲 Installer
        </button>
      )}
      {user?.role==="proviseur" && (
        <select value={viewDeptId||""} onChange={e=>setViewDeptId(e.target.value?parseInt(e.target.value):null)}
          title="Filtrer par département"
          style={{padding:"5px 8px", borderRadius:8, border:`1px solid ${C.border}`, background:C.white, fontSize:11, fontWeight:700, color:C.txt, fontFamily:"inherit", cursor:"pointer", maxWidth:130, flexShrink:0}}>
          <option value="">🏛️ Tous les départements</option>
          {DEPARTEMENTS_LIST.map(d=><option key={d.id} value={d.id}>{d.emoji} {d.nom}</option>)}
        </select>
      )}
      <button onClick={()=>setLang(lang==="fr"?"en":"fr")} title="Langue / Language"
        style={{padding:"5px 9px", borderRadius:8, border:`1px solid ${C.border}`, background:C.white, fontSize:11, fontWeight:700, color:C.txtMuted, fontFamily:"inherit", cursor:"pointer", flexShrink:0}}>
        {lang==="fr" ? "FR" : "EN"}
      </button>
      <DarkModeToggle/>
      <button onClick={()=>{ if(window.confirm("Se déconnecter ?")) onLogout(); }}
        style={{padding:"5px 12px", borderRadius:8, border:`1px solid ${C.border}`, background:C.white, fontSize:11, fontWeight:600, cursor:"pointer", color:C.txtMuted, fontFamily:"inherit", flexShrink:0, whiteSpace:"nowrap"}}>
        {t("Déconnexion")}
      </button>
    </header>
  );
};


// ════════════════════════════════════════════════════════════════
// GÉNÉRATION BILAN TRIMESTRIEL — PDF automatique animatrice
// ════════════════════════════════════════════════════════════════
function genBilanTrimestre(trim, data) {
  const trimLabels = {T1:"1er Trimestre",T2:"2ème Trimestre",T3:"3ème Trimestre",ANN:"Annuel"};
  const periode = trimLabels[trim] || trim;
  const dateJour = new Date().toLocaleDateString("fr-FR",{day:"2-digit",month:"long",year:"numeric"});

  // Source enseignants avec fallback
  const supabaseEns = Object.values(data?.users||{}).filter(u=>u.role!=="proviseur");
  const enseignants = supabaseEns.length > 0
    ? supabaseEns.map(u=>({...u, col:u.col||getColor(u.id), classes:(u.classes||[]).length>0?u.classes:(ENS_CLASSES_REF[u.id]||[])}))
    : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant").map(a=>({...a, classes:ENS_CLASSES_REF[a.id]||[]}));

  // Calculer les stats par enseignant
  const statsEns = enseignants.map(ens => {
    let totalFait = 0, totalRef = 0, classesStats = [];
    (ens.classes||[]).forEach(cl => {
      const code = resolveProgCode(cl);
      const meta = code ? PROG_META[code] : null;
      if(!meta) return;
      const prog = (data?.prog||{})[`${ens.id}||${cl}`]||[];
      let lp = meta.lpRef, lfTrim = prog.length;
      if(trim !== "ANN" && code) {
        const range = getTrimRange(code, trim);
        if(range) {
          lp = LECONS_DATA[code]?.filter(l=>l.n>=range[0]&&l.n<=range[1]).length||lp;
          lfTrim = prog.filter(n=>n>=range[0]&&n<=range[1]).length;
        }
      }
      const taux = lp>0?Math.min(100, Math.round(lfTrim/lp*100)):0;
      totalFait += lfTrim; totalRef += lp;
      classesStats.push({cl, lf:lfTrim, lp, taux, ef:(ELEVES_DB[cl]||[]).length});
    });
    const tauxGlobal = totalRef>0?Math.min(100, Math.round(totalFait/totalRef*100)):0;
    return {...ens, totalFait, totalRef, tauxGlobal, classesStats};
  }).sort((a,b)=>b.tauxGlobal-a.tauxGlobal);

  const tauxMoyen = statsEns.length>0?Math.round(statsEns.reduce((s,e)=>s+e.tauxGlobal,0)/statsEns.length):0;
  const enAlerte  = statsEns.filter(e=>e.tauxGlobal<50).length;
  const enObjectif= statsEns.filter(e=>e.tauxGlobal>=75).length;

  const pct = v => v>=75?"#166534":v>=50?"#92400e":"#991b1b";
  const bar = v => `<div style="height:6px;background:#e2e8f0;border-radius:3px;margin-top:3px"><div style="width:${v}%;height:100%;background:${pct(v)};border-radius:3px"></div></div>`;

  const rowsHtml = statsEns.map((ens,i) => `
    <tr style="background:${i%2?"#fafafa":"#fff"}">
      <td style="padding:8px 10px;border:1px solid #ddd;font-weight:700">${i+1}</td>
      <td style="padding:8px 10px;border:1px solid #ddd;font-weight:700">${ens.nom}</td>
      <td style="padding:8px 10px;border:1px solid #ddd;text-align:center">${(ens.classes||[]).length}</td>
      <td style="padding:8px 10px;border:1px solid #ddd;text-align:center">${ens.totalFait}</td>
      <td style="padding:8px 10px;border:1px solid #ddd;text-align:center">${ens.totalRef}</td>
      <td style="padding:8px 10px;border:1px solid #ddd;text-align:center;font-weight:800;color:${pct(ens.tauxGlobal)}">${ens.tauxGlobal}%</td>
      <td style="padding:8px 10px;border:1px solid #ddd;text-align:center">
        ${ens.tauxGlobal>=75?"✅ Objectif":ens.tauxGlobal>=50?"⚠️ En cours":"🔴 Alerte"}
      </td>
      <td style="padding:8px 10px;border:1px solid #ddd;font-size:11px">${(ens.classes||[]).join(", ")}</td>
    </tr>`).join("");

  const html = `<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8">
<style>
  @page{size:A4;margin:15mm;}
  *{box-sizing:border-box;font-family:Arial,sans-serif;font-size:11px;}
  body{color:#1a1a1a;}
  h1{font-size:18px;font-weight:900;text-align:center;margin:0 0 4px;color:#1a5276;}
  h2{font-size:13px;text-align:center;color:#555;margin:0 0 20px;font-weight:400;}
  .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px;}
  .kpi{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px;text-align:center;}
  .kpi-val{font-size:26px;font-weight:900;line-height:1;}
  .kpi-lbl{font-size:10px;color:#888;margin-top:3px;}
  table{width:100%;border-collapse:collapse;margin-top:10px;}
  th{background:#1a5276;color:#fff;padding:9px 10px;text-align:left;font-weight:700;}
  .sig{display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:40px;}
  .sig-bloc{border-top:1px solid #999;padding-top:8px;font-size:11px;}
  @media print{button{display:none;}}
</style></head>
<body>
  <div style="text-align:center;margin-bottom:16px">
    <div style="font-size:13px;font-weight:700;color:#1a5276">LYCÉE DE KAKATARE — MAROUA</div>
    <div style="font-size:10px;color:#888">Conseil d'Enseignement SVTEEHB · 2025–2026</div>
  </div>
  <h1>BILAN PÉDAGOGIQUE — ${periode.toUpperCase()}</h1>
  <h2>Couverture des programmes · Édité le ${dateJour}</h2>

  <div class="kpis">
    <div class="kpi"><div class="kpi-val" style="color:#1a5276">${statsEns.length}</div><div class="kpi-lbl">Enseignants</div></div>
    <div class="kpi"><div class="kpi-val" style="color:${pct(tauxMoyen)}">${tauxMoyen}%</div><div class="kpi-lbl">Taux moyen</div></div>
    <div class="kpi"><div class="kpi-val" style="color:#166534">${enObjectif}</div><div class="kpi-lbl">Objectif ≥75%</div></div>
    <div class="kpi"><div class="kpi-val" style="color:#991b1b">${enAlerte}</div><div class="kpi-lbl">En alerte &lt;50%</div></div>
  </div>

  <table>
    <thead>
      <tr>
        <th style="width:35px">#</th>
        <th>Nom de l'enseignant(e)</th>
        <th style="width:60px;text-align:center">Classes</th>
        <th style="width:50px;text-align:center">LF</th>
        <th style="width:50px;text-align:center">LP</th>
        <th style="width:65px;text-align:center">Taux</th>
        <th style="width:90px;text-align:center">Statut</th>
        <th>Classes enseignées</th>
      </tr>
    </thead>
    <tbody>${rowsHtml}</tbody>
  </table>

  ${enAlerte>0?`<div style="margin-top:16px;padding:10px 14px;background:#fef2f2;border:1px solid #fecaca;border-radius:8px;font-size:11px;color:#991b1b"><strong>⚠️ Enseignants en alerte :</strong> ${statsEns.filter(e=>e.tauxGlobal<50).map(e=>e.nom).join(", ")}</div>`:""}

  <div class="sig">
    <div class="sig-bloc">
      <strong>L'Animatrice Pédagogique</strong><br>
      AÏSSATOU SYLVIE — PCEG<br><br>
      Signature :
    </div>
    <div class="sig-bloc" style="text-align:right">
      Fait à Maroua, le ${dateJour}<br><br>
      <em>Document généré automatiquement par la plateforme SVTEEHB</em>
    </div>
  </div>

  <script>window.onload=()=>window.print();</script>
</body></html>`;

  // Ouvrir dans un iframe pour impression
  const iframe = document.createElement("iframe");
  iframe.style.cssText = "position:fixed;top:-9999px;left:-9999px;width:210mm;height:297mm";
  document.body.appendChild(iframe);
  iframe.contentDocument.open();
  iframe.contentDocument.write(html);
  iframe.contentDocument.close();
  iframe.onload = () => {
    setTimeout(() => {
      try { iframe.contentWindow.focus(); iframe.contentWindow.print(); }
      catch(e) { console.warn("Impression:", e); }
      setTimeout(() => document.body.removeChild(iframe), 5000);
    }, 500);
  };
}

// ══════════════════════════════════════════════════════════════════════
// SUIVI PROGRAMME SYNTHÉTIQUE — Vue animatrice : tableau global
// Tous les enseignants · Toutes les classes · En un coup d'œil
// ══════════════════════════════════════════════════════════════════════
// ── Courbe d'évolution T1→T2→T3 — SVG natif, sans dépendance externe ──────
// Mini-courbe compacte (sparkline) : utilisée dans chaque ligne de tableau
const MiniEvolutionChart = ({ values, width=70, height=28, color }) => {
  const vals = values.map(v => v===null||v===undefined ? null : v);
  const present = vals.filter(v=>v!==null);
  if (present.length < 2) return <span style={{fontSize:11, color:"#cbd5e1"}}>—</span>;
  const min = 0, max = 100;
  const n = vals.length;
  const pts = vals.map((v,i) => v===null ? null : {
    x: (i/(n-1)) * (width-8) + 4,
    y: height-4 - ((v-min)/(max-min)) * (height-8),
  });
  const ptsValides = pts.filter(Boolean);
  const path = ptsValides.map((p,i)=> (i===0?"M":"L")+p.x.toFixed(1)+","+p.y.toFixed(1)).join(" ");
  const dernierPresent = [...vals].reverse().find(v=>v!==null);
  const lineColor = color || taux2col(dernierPresent);
  return (
    <svg width={width} height={height} style={{display:"block"}}>
      <path d={path} fill="none" stroke={lineColor} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      {pts.map((p,i)=> p && <circle key={i} cx={p.x} cy={p.y} r={i===pts.length-1?3:2} fill={lineColor}/>)}
    </svg>
  );
};

// Grande courbe annotée — vue d'ensemble avec axes, points et valeurs affichées
const EvolutionChartLarge = ({ series, height=140 }) => {
  // series: [{label:"T1", value:62}, {label:"T2", value:70}, {label:"T3", value:null}, ...]
  const width = 100; // en %, viewBox responsive
  const padTop=18, padBottom=24, padX=6;
  const innerH = height - padTop - padBottom;
  const valides = series.filter(s=>s.value!==null);
  if (valides.length < 2) {
    return <div style={{padding:"30px 14px", textAlign:"center", color:C.txtLight, fontSize:12}}>Pas assez de données pour tracer une courbe (au moins 2 trimestres avec des notes saisies sont nécessaires).</div>;
  }
  const n = series.length;
  const pts = series.map((s,i) => s.value===null ? null : {
    x: padX + (i/(n-1)) * (width-2*padX),
    y: padTop + innerH - (s.value/100) * innerH,
    label: s.label, value: s.value,
  });
  const ptsValides = pts.filter(Boolean);
  const path = ptsValides.map((p,i)=>(i===0?"M":"L")+p.x.toFixed(2)+","+p.y.toFixed(2)).join(" ");
  const aireD = path + ` L${ptsValides[ptsValides.length-1].x.toFixed(2)},${(padTop+innerH).toFixed(2)} L${ptsValides[0].x.toFixed(2)},${(padTop+innerH).toFixed(2)} Z`;
  return (
    <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height} preserveAspectRatio="none">
      {/* Lignes de repère horizontales (0/25/50/75/100%) */}
      {[0,25,50,75,100].map(g=>{
        const y = padTop + innerH - (g/100)*innerH;
        return <g key={g}>
          <line x1={padX} y1={y} x2={width-padX} y2={y} stroke="#eef1f4" strokeWidth="0.4"/>
          <text x="0" y={y+1.2} fontSize="3" fill="#94a3b8">{g}</text>
        </g>;
      })}
      <path d={aireD} fill={C.green} opacity="0.07"/>
      <path d={path} fill="none" stroke={C.green} strokeWidth="0.9" strokeLinecap="round" strokeLinejoin="round"/>
      {pts.map((p,i) => p && (
        <g key={i}>
          <circle cx={p.x} cy={p.y} r="1.4" fill="#fff" stroke={taux2col(p.value)} strokeWidth="1.1"/>
          <text x={p.x} y={p.y-3.5} fontSize="3.4" fontWeight="700" textAnchor="middle" fill={taux2col(p.value)}>{p.value}%</text>
          <text x={p.x} y={height-4} fontSize="3.2" fontWeight="700" textAnchor="middle" fill="#64748b">{p.label}</text>
        </g>
      ))}
    </svg>
  );
};

function SuiviProgrammePage() {
  const {data, showToast, user} = useApp();
  const {isMobile} = useDevice();
  const [trim, setTrim] = useState(0); // 0=Annuel,1=T1,2=T2,3=T3
  const [filtre, setFiltre] = useState("tous"); // tous|alerte|objectif
  const [vueMode, setVueMode] = useState("normal"); // normal|comparaison
  const [expandedEns, setExpandedEns] = useState(() => new Set());
  const toggleExpand = (ensId) => setExpandedEns(prev => {
    const next = new Set(prev);
    next.has(ensId) ? next.delete(ensId) : next.add(ensId);
    return next;
  });

  // Calcul du taux de couverture pour un trimestre donné (1,2,3) — réutilisé pour la comparaison
  const calcTauxTrim = (ens, cl, trimNum, data) => {
    const code = resolveProgCode(cl);
    const meta = code ? PROG_META[code] : null;
    const key  = `${ens.id}||${cl}`;
    const prog = (data?.prog||{})[key]||[];
    let lp = meta?.lpRef||0;
    const tk = ["T1","T2","T3"][trimNum-1];
    const range = code ? getTrimRange(code, tk) : null;
    if (range) {
      const lecons = LECONS_DATA[code]||[];
      lp = lecons.filter(l=>l.n>=range[0]&&l.n<=range[1]).length || lp;
    }
    const lf = range ? prog.filter(n=>n>=range[0]&&n<=range[1]).length : 0;
    return lp>0 ? Math.min(100, Math.round(lf/lp*100)) : null;
  };

  if (!data) return (
    <div style={{padding:"60px",textAlign:"center",color:C.txtMuted}}>
      <Spinner size={28} color={C.green}/><div style={{marginTop:12}}>Chargement…</div>
    </div>
  );

  // Source enseignants
  const supabaseEns = Object.values(data?.users||{}).filter(u=>u.role!=="proviseur");
  const enseignants = ((supabaseEns.length>0 || data?.deptFilterActive)
    ? supabaseEns.map(u=>({...u,col:u.col||getColor(u.id),ini:u.ini||getIni(u.nom),classes:(u.classes||[]).length>0?u.classes:(ENS_CLASSES_REF[u.id]||[])}))
    : DEMO_ACCOUNTS.filter(a=>a.role==="enseignant").map(a=>({...a,col:getColor(a.id),ini:getIni(a.nom),classes:ENS_CLASSES_REF[a.id]||[]}))
  );

  // Calculer stats par enseignant × classe
  const rows = enseignants.flatMap(ens => {
    // Dédupliquer les classes avant de générer les lignes
    const classesUniques = [...new Set(ens.classes||[])];
    return classesUniques.map(cl => {
      const code  = resolveProgCode(cl);
      const meta  = code ? PROG_META[code] : null;
      const key   = `${ens.id}||${cl}`;
      const prog  = (data?.prog||{})[key]||[];

      let lp = meta?.lpRef||0, tpP = meta?.tp?.length||0;
      let range = null;
      if (trim > 0 && code) {
        const tk    = ["T1","T2","T3"][trim-1];
        range = getTrimRange(code, tk);
        if (range) {
          const lecons = LECONS_DATA[code]||[];
          lp  = lecons.filter(l=>l.n>=range[0]&&l.n<=range[1]).length||lp;
          tpP = (meta?.tp||[]).filter(n=>n>=range[0]&&n<=range[1]).length;
        }
      }

      const lfTrim = trim===0 ? prog.length : prog.filter(n=>{
        const tk=["T1","T2","T3"][trim-1]; const r=getTrimRange(code,tk);
        return r&&n>=r[0]&&n<=r[1];
      }).length;

      const taux  = lp>0 ? Math.min(100, Math.round(lfTrim/lp*100)) : 0;
      // tpFait limité à la même plage que tpP (range) — sinon des TP/TD faits en avance
      // sur un autre trimestre se comptaient dans le trimestre affiché, faussant le taux.
      const tpFait= (meta?.tp||[]).filter(n => prog.includes(n) && (trim===0 || (range && n>=range[0] && n<=range[1]))).length;
      const ef    = (ELEVES_DB[cl]||[]).length;
      // Digitalisées
      const digKey= `${ens.id}||${cl}||dig`;
      const digProg = (data?.prog?.[digKey]||[]);
      const leconsList = code ? (LECONS_DATA[code]||[]) : [];
      const ldTot = leconsList.filter(l=>l.d===1).length;
      const ldFait= digProg.filter(n=>{ const l=leconsList.find(x=>x.n===n); return l&&l.d===1; }).length;
      const tauxDig = ldTot>0 ? Math.min(100, Math.round(ldFait/ldTot*100)) : 0;

      return {ens, cl, lp, lf:lfTrim, taux, tpP, tpFait, ef, code, meta, ldTot, ldFait, tauxDig};
    });
  }).filter(r => {
    if (filtre==="alerte")  return r.taux < 50;
    if (filtre==="objectif") return r.taux >= 75;
    return true;
  }).sort((a,b)=>a.taux-b.taux);

  // KPIs globaux
  const totalRows = rows.length;
  const enAlerte  = rows.filter(r=>r.taux<50).length;
  const enObjectif= rows.filter(r=>r.taux>=75).length;
  const tauxMoyen = totalRows>0 ? Math.round(rows.reduce((s,r)=>s+r.taux,0)/totalRows) : 0;

  // Regroupement par enseignant — un seul nom affiché, dépliable
  const groupedByEns = [];
  const seenEns = new Map();
  rows.forEach(r => {
    if (!seenEns.has(r.ens.id)) {
      const grp = { ens:r.ens, classes:[] };
      seenEns.set(r.ens.id, grp);
      groupedByEns.push(grp);
    }
    seenEns.get(r.ens.id).classes.push(r);
  });
  groupedByEns.forEach(g => {
    g.avgTaux = Math.round(g.classes.reduce((s,r)=>s+r.taux,0)/g.classes.length);
    g.nbAlerte = g.classes.filter(r=>r.taux<50).length;
    g.nbObjectif = g.classes.filter(r=>r.taux>=75).length;
  });
  groupedByEns.sort((a,b)=>a.avgTaux-b.avgTaux);

  // ── Données pour la vue Comparaison T1/T2/T3 ──────────────────────
  const rowsComparaison = enseignants.flatMap(ens => {
    const classesUniques = [...new Set(ens.classes||[])];
    return classesUniques.map(cl => ({
      ens, cl,
      t1: calcTauxTrim(ens, cl, 1, data),
      t2: calcTauxTrim(ens, cl, 2, data),
      t3: calcTauxTrim(ens, cl, 3, data),
    }));
  });
  const groupedComparaison = [];
  const seenComp = new Map();
  rowsComparaison.forEach(r => {
    if (!seenComp.has(r.ens.id)) {
      const grp = { ens:r.ens, classes:[] };
      seenComp.set(r.ens.id, grp);
      groupedComparaison.push(grp);
    }
    seenComp.get(r.ens.id).classes.push(r);
  });
  const moyenneGlobale = (k) => {
    const vals = rowsComparaison.map(r=>r[k]).filter(v=>v!==null);
    return vals.length>0 ? Math.round(vals.reduce((s,v)=>s+v,0)/vals.length) : null;
  };
  const serieGlobale = [
    {label:"T1", value: moyenneGlobale("t1")},
    {label:"T2", value: moyenneGlobale("t2")},
    {label:"T3", value: moyenneGlobale("t3")},
  ];

  return (
    <div style={{padding: isMobile?12:20, display:"flex", flexDirection:"column", gap: isMobile?10:14}}>

      {/* Header */}
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding: isMobile?"12px 14px":"14px 18px", display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:10}}>
        <div>
          <h2 style={{fontSize: isMobile?14:16, fontWeight:800, color:C.txt, margin:0}}>📊 Suivi programme</h2>
          <p style={{fontSize: isMobile?10.5:12, color:C.txtMuted, margin:"4px 0 0"}}>Tous les enseignants · Toutes les classes · 2025–2026</p>
        </div>
        <div style={{display:"flex", gap:8, alignItems:"center"}}>
          <button onClick={()=>genBilanTrimestre(["ANN","T1","T2","T3"][trim], data)}
            style={{padding: isMobile?"7px 12px":"8px 16px",background:`linear-gradient(135deg,${C.greenDark},${C.green})`,
              color:"#fff",border:"none",borderRadius:10,fontSize: isMobile?11:12,fontWeight:700,
              cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
            📄 {isMobile?"Bilan":"Générer bilan"}
          </button>
        </div>
        {/* Filtre trimestre */}
        <div style={{display:"flex", gap:5}}>
          {[{l:"Annuel",v:0},{l:"T1",v:1},{l:"T2",v:2},{l:"T3",v:3}].map(t=>(
            <button key={t.v} onClick={()=>setTrim(t.v)}
              style={{padding:"6px 11px", borderRadius:8, fontSize:11, fontWeight:700, cursor:"pointer", fontFamily:"inherit", border:`1.5px solid ${trim===t.v?C.green:C.border}`, background:trim===t.v?C.green:"transparent", color:trim===t.v?"#fff":C.txtMuted, transition:"all .15s"}}>
              {t.l}
            </button>
          ))}
        </div>
      </div>

      {/* KPIs */}
      <div style={{display:"grid", gridTemplateColumns: isMobile?"repeat(2,1fr)":"repeat(4,1fr)", gap:10}}>
        {[
          {label:"Classes suivies",  val:totalRows,   col:C.blue,  bg:C.bluePale,   emoji:"📚"},
          {label:"Taux moyen",       val:`${tauxMoyen}%`, col:taux2col(tauxMoyen), bg:taux2bg(tauxMoyen), emoji:"📊"},
          {label:"Objectif ≥75%",   val:enObjectif,  col:C.green, bg:C.greenPale,  emoji:"✅"},
          {label:"En alerte <50%",  val:enAlerte,    col:enAlerte>0?C.red:C.green, bg:enAlerte>0?C.redPale:C.greenPale, emoji:"⚠️"},
        ].map((k,i)=>(
          <div key={i} style={{background:k.bg, borderRadius:11, border:`1px solid ${C.border}`, padding:"12px 14px"}}>
            <div style={{display:"flex", justifyContent:"space-between", marginBottom:5}}>
              <span style={{fontSize:10, fontWeight:600, color:C.txtMuted}}>{k.label}</span>
              <span style={{fontSize:16}}>{k.emoji}</span>
            </div>
            <div style={{fontSize:24, fontWeight:900, color:k.col}}>{k.val}</div>
          </div>
        ))}
      </div>

      {/* Filtres rapides */}
      <div style={{display:"flex", gap:8, flexWrap:"wrap", justifyContent:"space-between", alignItems:"center"}}>
        <div style={{display:"flex", gap:8}}>
          {[
            {id:"tous",     label:"Tous",           count:rows.length},
            {id:"alerte",   label:"⚠ En alerte",    count:enAlerte},
            {id:"objectif", label:"✅ Objectif atteint", count:enObjectif},
          ].map(f=>(
            <button key={f.id} onClick={()=>setFiltre(f.id)}
              style={{padding:"6px 14px", borderRadius:20, fontSize:11, fontWeight:700, cursor:"pointer", fontFamily:"inherit",
                border:`1.5px solid ${filtre===f.id?C.green:C.border}`,
                background:filtre===f.id?C.greenPale:"transparent",
                color:filtre===f.id?C.green:C.txtMuted}}>
              {f.label} <span style={{opacity:.6}}>({f.count})</span>
            </button>
          ))}
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>{
              const exportRows = vueMode==="comparaison"
                ? rowsComparaison.map(r=>({
                    "Enseignant": getNomCourt(r.ens.nom), "Classe": r.cl,
                    "T1 (%)": r.t1, "T2 (%)": r.t2, "T3 (%)": r.t3,
                    "Tendance": (r.t1!==null&&r.t3!==null) ? (r.t3-r.t1) : ""
                  }))
                : rows.map(r=>({
                    "Enseignant": getNomCourt(r.ens.nom), "Classe": r.cl, "Élèves": r.ef,
                    "Leçons faites": r.lf, "Leçons prévues": r.lp, "Taux couverture (%)": r.taux,
                    "TP fait": r.tpFait, "TP prévu": r.tpP,
                    "Digital fait": r.ldFait, "Digital total": r.ldTot, "Taux digital (%)": r.tauxDig,
                    "Statut": r.taux<50?"Alerte":r.taux>=75?"Objectif":"En cours"
                  }));
              if (exportRows.length === 0) {
                showToast("⚠ Aucune donnée à exporter pour ce filtre", false);
                return;
              }
              const ok = exportToExcel(
                `Suivi_programme_${vueMode==="comparaison"?"comparaison":TRIM_LABELS[["ANN","T1","T2","T3"][trim]]||"annuel"}`,
                "Suivi", exportRows
              );
              showToast(ok ? "✓ Fichier Excel téléchargé" : "⚠ Échec de l'export", ok);
            }}
            style={{padding:"6px 14px", borderRadius:20, fontSize:11, fontWeight:700, cursor:"pointer", fontFamily:"inherit",
              border:`1.5px solid #15803d`, background:"#f0fdf4", color:"#15803d", display:"flex", alignItems:"center", gap:6}}>
            📥 Exporter Excel
          </button>
          <button onClick={()=>setVueMode(vueMode==="normal"?"comparaison":"normal")}
          style={{padding:"6px 14px", borderRadius:20, fontSize:11, fontWeight:700, cursor:"pointer", fontFamily:"inherit",
            border:`1.5px solid ${vueMode==="comparaison"?"#7c3aed":C.border}`,
            background:vueMode==="comparaison"?"#f5f3ff":"transparent",
            color:vueMode==="comparaison"?"#7c3aed":C.txtMuted, display:"flex", alignItems:"center", gap:6}}>
          📈 {vueMode==="comparaison" ? "Vue normale" : "Comparer T1 · T2 · T3"}
        </button>
        </div>
      </div>

      {vueMode==="comparaison" ? (
      <div style={{display:"flex", flexDirection:"column", gap:14}}>
      <div style={{background:C.white, borderRadius:14, border:`1px solid ${C.border}`, padding:"16px 18px"}}>
        <h3 style={{fontSize:12.5, fontWeight:800, color:C.txt, margin:"0 0 8px"}}>📈 Évolution du taux de couverture global · Tous enseignants</h3>
        <EvolutionChartLarge series={serieGlobale}/>
      </div>
      <div style={{background:C.white, borderRadius:14, border:`1px solid ${C.border}`, overflow:"hidden", boxShadow:"0 1px 3px rgba(0,0,0,.04)"}}>
        <div style={{overflowX:"auto"}}>
        <table style={{width:"100%", minWidth: 640, borderCollapse:"collapse", fontSize:12.5}}>
          <thead>
            <tr style={{background:"#fafbfc", borderBottom:`2px solid ${C.border}`}}>
              {["Enseignant / Classe","T1","T2","T3","Évolution","Tendance"].map((h,i)=>(
                <th key={i} style={{padding:"11px 14px", textAlign:i===0?"left":"center", color:"#64748b", fontSize:10.5, fontWeight:700, textTransform:"uppercase", letterSpacing:".04em"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {groupedComparaison.map(g => {
              const isOpen = expandedEns.has(g.ens.id);
              const avgT = (k) => { const vals=g.classes.map(r=>r[k]).filter(v=>v!==null); return vals.length>0?Math.round(vals.reduce((s,v)=>s+v,0)/vals.length):null; };
              const gT1=avgT("t1"), gT2=avgT("t2"), gT3=avgT("t3");
              const cellTaux = (v) => v===null
                ? <span style={{fontSize:11,color:"#cbd5e1"}}>—</span>
                : <span style={{fontSize:12, fontWeight:800, color:taux2col(v)}}>{v}%</span>;
              const tendance = (v1,v3) => {
                if (v1===null||v3===null) return <span style={{color:"#cbd5e1"}}>—</span>;
                const diff = v3-v1;
                if (diff>5) return <span style={{color:"#16a34a",fontWeight:800}}>↑ +{diff}</span>;
                if (diff<-5) return <span style={{color:"#ef4444",fontWeight:800}}>↓ {diff}</span>;
                return <span style={{color:"#94a3b8",fontWeight:700}}>→ stable</span>;
              };
              return [
                <tr key={`h-${g.ens.id}`} onClick={()=>toggleExpand(g.ens.id)}
                  style={{cursor:"pointer", background:isOpen?"#fafbfc":"transparent", borderBottom:`1px solid #f1f5f9`}}
                  onMouseEnter={e=>e.currentTarget.style.background="#fafbfc"}
                  onMouseLeave={e=>e.currentTarget.style.background=isOpen?"#fafbfc":"transparent"}>
                  <td style={{padding:"13px 14px"}}>
                    <div style={{display:"flex", alignItems:"center", gap:10}}>
                      <span style={{fontSize:11, color:"#94a3b8", transform:isOpen?"rotate(90deg)":"none", display:"inline-block"}}>▶</span>
                      <Avatar ens={g.ens} size={28} fontSize={9}/>
                      <div>
                        <div style={{fontSize:12.5,fontWeight:700,color:"#1e293b"}}>{getNomCourt(g.ens.nom)}</div>
                        <div style={{fontSize:10,color:"#94a3b8"}}>{g.classes.length} classe{g.classes.length>1?"s":""}</div>
                      </div>
                    </div>
                  </td>
                  <td style={{textAlign:"center"}}>{cellTaux(gT1)}</td>
                  <td style={{textAlign:"center"}}>{cellTaux(gT2)}</td>
                  <td style={{textAlign:"center"}}>{cellTaux(gT3)}</td>
                  <td style={{textAlign:"center"}}><div style={{display:"flex", justifyContent:"center"}}><MiniEvolutionChart values={[gT1,gT2,gT3]}/></div></td>
                  <td style={{textAlign:"center"}}>{tendance(gT1,gT3)}</td>
                </tr>,
                ...(!isOpen ? [] : g.classes.map(r => (
                  <tr key={`${g.ens.id}-${r.cl}`} style={{borderBottom:`1px solid #f1f5f9`, background:"#fcfdfe"}}>
                    <td style={{padding:"9px 14px 9px 30px", borderLeft:`2px solid ${C.border}`, fontSize:12, color:"#1e293b"}}>{r.cl}</td>
                    <td style={{textAlign:"center"}}>{cellTaux(r.t1)}</td>
                    <td style={{textAlign:"center"}}>{cellTaux(r.t2)}</td>
                    <td style={{textAlign:"center"}}>{cellTaux(r.t3)}</td>
                    <td style={{textAlign:"center"}}><div style={{display:"flex", justifyContent:"center"}}><MiniEvolutionChart values={[r.t1,r.t2,r.t3]}/></div></td>
                    <td style={{textAlign:"center"}}>{tendance(r.t1,r.t3)}</td>
                  </tr>
                )))
              ];
            })}
          </tbody>
        </table>
        </div>
        <div style={{padding:"10px 16px", fontSize:10.5, color:"#94a3b8", borderTop:`1px solid ${C.border}`}}>
          Tendance = évolution entre T1 et T3 · ↑ amélioration &gt;5pts · ↓ recul &gt;5pts · → stable
        </div>
      </div>
      </div>
      ) : (
      <>
      {/* Tableau synthétique (desktop) / Cartes (mobile) */}
      {isMobile ? (
        <div style={{display:"flex", flexDirection:"column", gap:8}}>
          {rows.length===0 ? (
            <div style={{padding:"40px",textAlign:"center",color:C.txtLight,background:C.white,borderRadius:12,border:`1px solid ${C.border}`}}>
              <div style={{fontSize:28,marginBottom:8}}>🔍</div>Aucune classe dans ce filtre
            </div>
          ) : rows.map((r,i)=>{
            const alerte  = r.taux < 50;
            const objectif= r.taux >= 75;
            const tauxTP = r.tpP>0 ? Math.min(100, Math.round(r.tpFait/r.tpP*100)) : null;
            return (
              <div key={i} style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`,
                borderLeft:`4px solid ${alerte?C.red:objectif?C.green:C.border}`, padding:"12px 14px"}}>
                {/* En-tête : enseignant + classe */}
                <div style={{display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:8}}>
                  <div style={{display:"flex", alignItems:"center", gap:8}}>
                    <Avatar ens={r.ens} size={28} fontSize={10}/>
                    <div>
                      <div style={{fontSize:12.5,fontWeight:700,color:C.txt}}>{getNomCourt(r.ens.nom)}</div>
                      <div style={{fontSize:11,color:C.txtMuted}}>{r.cl} · {r.ef} élèves</div>
                    </div>
                  </div>
                  <span style={{fontSize:16,fontWeight:900,color:taux2col(r.taux)}}>{r.taux}%</span>
                </div>

                {/* Barre progression */}
                <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:10}}>
                  <div style={{flex:1,height:7,borderRadius:4,overflow:"hidden",background:"#e2e8f0"}}>
                    <div style={{width:`${r.taux}%`,height:"100%",background:taux2col(r.taux),transition:"width .5s"}}/>
                  </div>
                  {alerte  && <span style={{fontSize:10,color:C.red,fontWeight:700,flexShrink:0}}>⚠ Alerte</span>}
                  {objectif && <span style={{fontSize:10,color:C.green,fontWeight:700,flexShrink:0}}>✓ Objectif</span>}
                </div>

                {/* Détails en grille 2x2 */}
                <div style={{display:"grid", gridTemplateColumns:"repeat(2,1fr)", gap:8, fontSize:11}}>
                  <div style={{background:"#f8fafc",borderRadius:8,padding:"6px 9px"}}>
                    <div style={{color:C.txtMuted,fontSize:9.5}}>Leçons</div>
                    <div style={{fontWeight:700,color:C.txt}}>{r.lf} / {r.lp}</div>
                  </div>
                  <div style={{background:"#f8fafc",borderRadius:8,padding:"6px 9px"}}>
                    <div style={{color:C.txtMuted,fontSize:9.5}}>TP réalisés</div>
                    <div style={{fontWeight:700,color:C.txt}}>{r.tpFait} / {r.tpP} {tauxTP!==null && <span style={{color:taux2col(tauxTP)}}>({tauxTP}%)</span>}</div>
                  </div>
                  <div style={{background:"#e0f2fe",borderRadius:8,padding:"6px 9px", gridColumn:"span 2"}}>
                    <div style={{color:"#0369a1",fontSize:9.5}}>🖥️ Digitalisation</div>
                    <div style={{fontWeight:700,color:"#0369a1"}}>{r.ldFait} / {r.ldTot} leçons ({r.tauxDig}%)</div>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      ) : (
      <div style={{background:C.white, borderRadius:14, border:`1px solid ${C.border}`, overflow:"hidden", boxShadow:"0 1px 3px rgba(0,0,0,.04)"}}>
        <table style={{width:"100%", borderCollapse:"collapse", fontSize:12.5}}>
          <thead>
            <tr style={{background:"#fafbfc", borderBottom:`2px solid ${C.border}`}}>
              {["Enseignant","Classe","Programme","TP / TD","Digitalisation","Statut"].map((h,i)=>(
                <th key={i} style={{padding:"11px 14px", textAlign:i===0?"left":i===5?"center":"left", color:"#64748b", fontSize:10.5, fontWeight:700, textTransform:"uppercase", letterSpacing:".04em", whiteSpace:"nowrap"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {groupedByEns.length===0 ? (
              <tr><td colSpan={6} style={{padding:"48px",textAlign:"center",color:C.txtLight}}>
                <div style={{fontSize:30,marginBottom:10}}>🔍</div>Aucune classe dans ce filtre
              </td></tr>
            ) : groupedByEns.flatMap((g) => {
              const isOpen = expandedEns.has(g.ens.id);
              const statutGrp = g.nbAlerte>0
                ? {label:`${g.nbAlerte} en alerte`, bg:"#fef2f2", fg:"#b91c1c", dot:"#ef4444"}
                : g.nbObjectif===g.classes.length
                ? {label:"Tout en objectif", bg:"#f0fdf4", fg:"#166534", dot:"#16a34a"}
                : {label:"En cours", bg:"#fffbeb", fg:"#92400e", dot:"#f59e0b"};

              const headerRow = (
                <tr key={`h-${g.ens.id}`} onClick={()=>toggleExpand(g.ens.id)}
                  style={{cursor:"pointer", background:isOpen?"#fafbfc":"transparent", borderBottom:`1px solid #f1f5f9`}}
                  onMouseEnter={e=>e.currentTarget.style.background="#fafbfc"}
                  onMouseLeave={e=>e.currentTarget.style.background=isOpen?"#fafbfc":"transparent"}>
                  <td style={{padding:"13px 14px"}} colSpan={2}>
                    <div style={{display:"flex", alignItems:"center", gap:10}}>
                      <span style={{fontSize:11, color:"#94a3b8", transition:"transform .15s", display:"inline-block", transform:isOpen?"rotate(90deg)":"none"}}>▶</span>
                      <Avatar ens={g.ens} size={30} fontSize={10}/>
                      <div>
                        <div style={{fontSize:12.5,fontWeight:700,color:"#1e293b"}}>{getNomCourt(g.ens.nom)}</div>
                        <div style={{fontSize:10.5,color:"#94a3b8",marginTop:1}}>{g.classes.length} classe{g.classes.length>1?"s":""}</div>
                      </div>
                    </div>
                  </td>
                  <td style={{padding:"13px 14px"}}>
                    <span style={{fontSize:12.5, fontWeight:800, color:taux2col(g.avgTaux)}}>{g.avgTaux}%</span>
                    <span style={{fontSize:10, color:"#94a3b8", marginLeft:5}}>moy.</span>
                  </td>
                  <td colSpan={2}/>
                  <td style={{padding:"13px 14px", textAlign:"center"}}>
                    <span style={{display:"inline-flex", alignItems:"center", gap:6, padding:"5px 12px", borderRadius:20,
                      background:statutGrp.bg, color:statutGrp.fg, fontSize:11, fontWeight:700}}>
                      <span style={{width:6,height:6,borderRadius:"50%",background:statutGrp.dot,flexShrink:0}}/>
                      {statutGrp.label}
                    </span>
                  </td>
                </tr>
              );

              const classRows = !isOpen ? [] : g.classes.map((r,i) => {
                const alerte   = r.taux < 50;
                const objectif = r.taux >= 75;
                const tauxTP = r.tpP>0 ? Math.min(100, Math.round(r.tpFait/r.tpP*100)) : null;
                const badge = (val, taux, color) => (
                  <div style={{display:"inline-flex", alignItems:"center", gap:7}}>
                    <span style={{fontSize:12.5, fontWeight:600, color:"#334155"}}>{val}</span>
                    {taux!==null ? (
                      <span style={{fontSize:10.5, fontWeight:800, padding:"2px 7px", borderRadius:20, background:`${color}15`, color}}>{taux}%</span>
                    ) : (
                      <span style={{fontSize:10.5, fontWeight:600, padding:"2px 7px", borderRadius:20, background:"#f1f5f9", color:"#94a3b8"}}>—</span>
                    )}
                  </div>
                );
                const statutInfo = alerte
                  ? {label:"Alerte", bg:"#fef2f2", fg:"#b91c1c", dot:"#ef4444"}
                  : objectif
                  ? {label:"Objectif", bg:"#f0fdf4", fg:"#166534", dot:"#16a34a"}
                  : {label:"En cours", bg:"#fffbeb", fg:"#92400e", dot:"#f59e0b"};
                return (
                  <tr key={`${g.ens.id}-${r.cl}`} style={{borderBottom:`1px solid #f1f5f9`, background:"#fcfdfe"}}>
                    <td style={{padding:"10px 14px 10px 30px", borderLeft:`2px solid ${C.border}`}} colSpan={2}>
                      <div style={{fontSize:12.5,fontWeight:600,color:"#1e293b"}}>{r.cl}</div>
                      <div style={{fontSize:10.5,color:"#94a3b8",marginTop:1}}>{r.ef} élève{r.ef>1?"s":""}</div>
                    </td>
                    <td style={{padding:"10px 14px"}}>{badge(`${r.lf}/${r.lp}`, r.taux, taux2col(r.taux))}</td>
                    <td style={{padding:"10px 14px"}}>{badge(`${r.tpFait}/${r.tpP}`, tauxTP, tauxTP!==null?taux2col(tauxTP):"#94a3b8")}</td>
                    <td style={{padding:"10px 14px"}}>{badge(`${r.ldFait}/${r.ldTot}`, r.ldTot>0?r.tauxDig:null, "#0369a1")}</td>
                    <td style={{padding:"10px 14px", textAlign:"center"}}>
                      <span style={{display:"inline-flex", alignItems:"center", gap:6, padding:"4px 10px", borderRadius:20,
                        background:statutInfo.bg, color:statutInfo.fg, fontSize:10.5, fontWeight:700}}>
                        <span style={{width:5,height:5,borderRadius:"50%",background:statutInfo.dot,flexShrink:0}}/>
                        {statutInfo.label}
                      </span>
                    </td>
                  </tr>
                );
              });

              return [headerRow, ...classRows];
            })}
          </tbody>
        </table>
      </div>
      )}
      </>
      )}
    </div>
  );
}


// ════════════════════════════════════════════════════════════════
// PAGE CHANGER MOT DE PASSE
// Accessible depuis les paramètres enseignant
// ════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════
// PAGE GESTION ANNUELLE — Mise à jour des données (animatrice)
//   - Import CSV des élèves par classe
//   - Visualisation de l'état des données
//   - Préparation nouvelle année scolaire
// ════════════════════════════════════════════════════════════════
function GestionAnnuellePage() {
  const {showToast, data, setPage} = useApp();
  const {isMobile} = useDevice();
  const [selClasse, setSelClasse] = useState("6ème 1");
  const [csvText, setCsvText]     = useState("");
  const [preview, setPreview]     = useState([]);
  const [importing, setImporting] = useState(false);

  const toutesClasses = Object.keys(ELEVES_DB).sort();

  // Parser le CSV : format "NOM Prénom;M" ou "NOM Prénom,F" — une ligne par élève
  const parseCsv = (text) => {
    const lines = text.split("\n").map(l=>l.trim()).filter(Boolean);
    return lines.map((line,i) => {
      const parts = line.split(/[;,\t]/).map(p=>p.trim());
      const nom = parts[0] || "";
      let g = (parts[1]||"").toUpperCase();
      if(g!=="M" && g!=="F") g = "M"; // défaut
      return { nom, g, valid: nom.length>2 };
    }).filter(e=>e.nom);
  };

  const handlePreview = () => {
    const parsed = parseCsv(csvText);
    setPreview(parsed);
    if(parsed.length===0) showToast("Aucun élève détecté dans le texte", false);
    else showToast(`${parsed.length} élèves détectés`, true);
  };

  const handleImport = async () => {
    const valides = preview.filter(e=>e.valid);
    if(valides.length===0) return showToast("Aucun élève valide à importer", false);
    setImporting(true);

    // Générer les IDs et construire la nouvelle liste
    const clId = selClasse.replace(/[^a-zA-Z0-9]/g,"_");
    const nouveauxEleves = valides.map((e,i)=>({
      id: `${clId}_${i+1}`,
      nom: e.nom.toUpperCase(),
      g: e.g
    }));

    // Mettre à jour ELEVES_DB en mémoire
    ELEVES_DB[selClasse] = nouveauxEleves;

    // Sauvegarder dans Supabase (table eleves)
    const ok = await sb.upsert("eleves_import",
      {classe:selClasse, donnees:JSON.stringify(nouveauxEleves), updated_at:new Date().toISOString()},
      "classe");

    setImporting(false);
    if(ok) {
      showToast(`✅ ${valides.length} élèves importés dans ${selClasse}`, true);
      setCsvText(""); setPreview([]);
    } else {
      showToast("✅ Import local effectué (Supabase non configuré)", true);
      setCsvText(""); setPreview([]);
    }
  };

  return (
    <div style={{padding:20, maxWidth:720, margin:"0 auto", display:"flex", flexDirection:"column", gap:16}}>
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"18px 20px"}}>
        <h2 style={{fontSize:17, fontWeight:800, color:C.txt, margin:"0 0 4px"}}>🔄 Gestion annuelle des données</h2>
        <p style={{fontSize:12, color:C.txtMuted, margin:0}}>Importer les listes d'élèves · Préparer une nouvelle année scolaire</p>
      </div>

      {/* État actuel */}
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"16px 18px"}}>
        <h3 style={{fontSize:13, fontWeight:700, color:C.txt, margin:"0 0 10px"}}>📊 État des données</h3>
        <div style={{display:"grid", gridTemplateColumns: isMobile?"repeat(2,1fr)":"repeat(3,1fr)", gap:10}}>
          <div style={{background:C.greenPale, borderRadius:9, padding:"12px 14px", textAlign:"center"}}>
            <div style={{fontSize:24, fontWeight:900, color:C.green}}>{toutesClasses.length}</div>
            <div style={{fontSize:10, color:C.green, fontWeight:600}}>Classes</div>
          </div>
          <div style={{background:"#eff6ff", borderRadius:9, padding:"12px 14px", textAlign:"center"}}>
            <div style={{fontSize:24, fontWeight:900, color:"#1e40af"}}>{Object.values(ELEVES_DB).reduce((s,e)=>s+e.length,0)}</div>
            <div style={{fontSize:10, color:"#1e40af", fontWeight:600}}>Élèves total</div>
          </div>
          <div style={{background:"#fef3c7", borderRadius:9, padding:"12px 14px", textAlign:"center"}}>
            <div style={{fontSize:24, fontWeight:900, color:"#92400e"}}>2025-26</div>
            <div style={{fontSize:10, color:"#92400e", fontWeight:600}}>Année scolaire</div>
          </div>
        </div>
      </div>

      {/* Import CSV */}
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"16px 18px", display:"flex", flexDirection:"column", gap:12}}>
        <h3 style={{fontSize:13, fontWeight:700, color:C.txt, margin:0}}>📥 Importer une liste d'élèves</h3>

        <div>
          <label style={{display:"block", fontSize:11, fontWeight:700, color:C.txtMuted, marginBottom:5}}>Classe cible</label>
          <select value={selClasse} onChange={e=>setSelClasse(e.target.value)}
            style={{width:"100%", padding:"9px 12px", border:`1.5px solid ${C.border}`, borderRadius:9, fontSize:13, fontFamily:"inherit", background:C.white}}>
            {toutesClasses.map(cl=><option key={cl} value={cl}>{cl} ({(ELEVES_DB[cl]||[]).length} élèves actuels)</option>)}
          </select>
        </div>

        <div>
          <label style={{display:"block", fontSize:11, fontWeight:700, color:C.txtMuted, marginBottom:5}}>
            Liste des élèves (un par ligne : <code style={{background:"#f1f5f9",padding:"1px 5px",borderRadius:4}}>NOM Prénom;M</code> ou <code style={{background:"#f1f5f9",padding:"1px 5px",borderRadius:4}}>NOM Prénom;F</code>)
          </label>
          <textarea value={csvText} onChange={e=>setCsvText(e.target.value)}
            placeholder={"MBASSA Jean;M\nFOTSO Marie;F\nNGONO Paul;M"}
            rows={8}
            style={{width:"100%", padding:"10px 14px", border:`1.5px solid ${C.border}`, borderRadius:9, fontSize:13, fontFamily:"monospace", resize:"vertical"}}/>
        </div>

        <div style={{display:"flex", gap:10}}>
          <button onClick={handlePreview}
            style={{padding:"10px 18px", background:C.white, border:`1.5px solid ${C.green}`, borderRadius:9, fontSize:13, fontWeight:700, color:C.green, cursor:"pointer", fontFamily:"inherit"}}>
            👁 Aperçu
          </button>
          {preview.length>0 && (
            <button onClick={handleImport} disabled={importing}
              style={{flex:1, padding:"10px 18px", background:importing?"#94a3b8":`linear-gradient(135deg,${C.greenDark},${C.green})`, border:"none", borderRadius:9, fontSize:13, fontWeight:700, color:"#fff", cursor:importing?"not-allowed":"pointer", fontFamily:"inherit"}}>
              {importing?"Import…":`✓ Importer ${preview.filter(e=>e.valid).length} élèves dans ${selClasse}`}
            </button>
          )}
        </div>

        {/* Aperçu */}
        {preview.length>0 && (
          <div style={{border:`1px solid ${C.border}`, borderRadius:9, overflow:"hidden", maxHeight:260, overflowY:"auto"}}>
            <div style={{padding:"8px 12px", background:"#f8fafc", fontSize:11, fontWeight:700, color:C.txtMuted, position:"sticky", top:0}}>
              {preview.length} élèves · {preview.filter(e=>e.valid).length} valides
            </div>
            {preview.map((e,i)=>(
              <div key={i} style={{padding:"7px 12px", borderTop:`1px solid ${C.border}`, display:"flex", alignItems:"center", gap:10, fontSize:12, opacity:e.valid?1:0.4}}>
                <span style={{width:24, textAlign:"center", color:C.txtLight, fontSize:10}}>{i+1}</span>
                <span style={{width:24, textAlign:"center", fontSize:13}}>{e.g==="F"?"👧":"👦"}</span>
                <span style={{flex:1, fontWeight:600, color:C.txt}}>{e.nom}</span>
                {!e.valid && <span style={{fontSize:10, color:C.red}}>⚠ nom trop court</span>}
              </div>
            ))}
          </div>
        )}

        <div style={{background:"#fffbeb", border:"1px solid #fde68a", borderRadius:9, padding:"11px 14px", fontSize:11.5, color:"#92400e"}}>
          ⚠️ <strong>Attention :</strong> l'import remplace TOUTE la liste actuelle de la classe sélectionnée. Vérifiez l'aperçu avant de confirmer. Cette action est définitive.
        </div>
      </div>

      {/* Aide format */}
      <div style={{background:"#eff6ff", border:"1px solid #bfdbfe", borderRadius:10, padding:"14px 16px", fontSize:12, color:"#1e40af"}}>
        <strong>💡 Comment préparer votre liste depuis Excel :</strong>
        <ol style={{margin:"8px 0 0", paddingLeft:20, lineHeight:1.7}}>
          <li>Dans Excel : colonne A = NOM Prénom, colonne B = M ou F</li>
          <li>Sélectionnez les cellules → Copier (Ctrl+C)</li>
          <li>Collez ici dans la zone de texte ci-dessus</li>
          <li>Cliquez « Aperçu » puis « Importer »</li>
        </ol>
      </div>

      {/* ── Lien vers l'éditeur d'emploi du temps (unifié dans la page Emploi du temps) ── */}
      <div style={{background:C.greenPale, borderRadius:12, border:`1px solid ${C.greenBorder}`, padding:"16px 18px", display:"flex", alignItems:"center", justifyContent:"space-between", gap:14, flexWrap:"wrap"}}>
        <div>
          <h3 style={{fontSize:13, fontWeight:700, color:C.greenDark, margin:0}}>📅 Emploi du temps</h3>
          <p style={{fontSize:11.5, color:C.txtMuted, margin:"4px 0 0", lineHeight:1.5}}>
            La modification de l'EDT (par enseignant) se fait maintenant directement dans la page « Emploi du temps », onglet « Par enseignant ».
          </p>
        </div>
        <button onClick={()=>setPage("edt")}
          style={{flexShrink:0, padding:"10px 18px", background:`linear-gradient(135deg,${C.greenDark},${C.green})`, border:"none", borderRadius:9, fontSize:13, fontWeight:700, color:"#fff", cursor:"pointer", fontFamily:"inherit"}}>
          → Aller à l'Emploi du temps
        </button>
      </div>
    </div>
  );
}

function DepartementsPage() {
  const {data, showToast} = useApp();
  const {isMobile} = useDevice();
  const [matieres, setMatieres] = useState(null);
  const [loading, setLoading] = useState(true);
  const [openDept, setOpenDept] = useState(null);

  // Écouter l'event dept:open dispatché depuis la sidebar
  useEffect(()=>{
    const handler = (e) => {
      const nom = e.detail;
      window.__deptFilter = null;
      const dept = DEPARTEMENTS_LIST.find(d=>(d.nom||"").toLowerCase().includes(nom.toLowerCase()));
      if (dept) setOpenDept(dept.id);
    };
    window.addEventListener("dept:open", handler);
    // Appliquer aussi si __deptFilter déjà défini (navigation initiale)
    if (window.__deptFilter) {
      const nom = window.__deptFilter;
      window.__deptFilter = null;
      const dept = DEPARTEMENTS_LIST.find(d=>(d.nom||"").toLowerCase().includes(nom.toLowerCase()));
      if (dept) setOpenDept(dept.id);
    }
    return () => window.removeEventListener("dept:open", handler);
  },[]);
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
    setSavingId(`new-${deptId}`);
    const ok = await sb.rpc("admin_add_matiere", { p_nom: newMatiere.trim(), p_departement_id: deptId });
    if (ok) { showToast(`✓ ${newMatiere.trim()} ajoutée`); setNewMatiere(""); await loadMatieres(); }
    else showToast("⚠ Échec de l'ajout", false);
    setSavingId(null);
  };

  const supprimerMatiere = async (m) => {
    if (!window.confirm(`Supprimer "${m.nom}" ?`)) return;
    setSavingId(m.id);
    const ok = await sb.rpc("admin_delete_matiere", { p_id: m.id });
    if (ok) { showToast(`✓ ${m.nom} supprimée`); await loadMatieres(); }
    else showToast("⚠ Échec de la suppression", false);
    setSavingId(null);
  };

  const renommerMatiere = async (m, nom) => {
    if (!nom.trim() || nom.trim()===m.nom) { setEditingNom(null); return; }
    setSavingId(m.id);
    const ok = await sb.rpc("admin_rename_matiere", { p_id: m.id, p_nom: nom.trim() });
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

      {/* Bouton retour si filtré depuis sidebar */}
      {openDept && (
        <button onClick={()=>setOpenDept(null)}
          style={{alignSelf:"flex-start",padding:"6px 14px",borderRadius:8,border:`1px solid ${C.border}`,background:C.white,
            color:C.txtMuted,fontSize:11.5,fontWeight:600,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:6}}>
          ← Tous les départements
        </button>
      )}
      <div style={{display:"grid", gridTemplateColumns: isMobile?"1fr":"repeat(auto-fill, minmax(320px, 1fr))", gap:12}}>
        {DEPARTEMENTS_LIST.filter(d=>!openDept||openDept===d.id).map(d => {
          const deptMatieres = (matieres||[]).filter(m=>m.departement_id===d.id);
          const isOpen = openDept===d.id || (!openDept && false);
          // Enrichissement
          const enseignantsDept = Object.values(data?.users||{}).filter(u=>u.departement_id===d.id&&u.role==="enseignant");
          const animateur = Object.values(data?.users||{}).find(u=>u.departement_id===d.id&&(u.role==="animateur"||u.role==="animatrice"));
          const epDept = (data?.epreuves||[]).filter(e=>enseignantsDept.some(u=>u.id===e.ens_id));
          const epSoumises = epDept.filter(e=>e.statut!=="brouillon").length;
          // Taux couverture programme (moyenne enseignants du dept)
          const tauxProg = enseignantsDept.length===0?null:(()=>{
            let total=0,count=0;
            enseignantsDept.forEach(u=>{
              (u.classes||[]).forEach(cl=>{
                const prog=data?.prog?.[u.id]?.[cl];
                if(prog){total+=(prog.faites||0);count+=(prog.total||1);}
              });
            });
            return count>0?Math.round(total/count*100):0;
          })();
          return (
            <div key={d.id} style={{background:C.white, borderRadius:12, border:`1.5px solid ${openDept===d.id?C.green:C.border}`, padding:16,
              boxShadow:openDept===d.id?"0 4px 20px rgba(11,77,44,.10)":"none"}}>
              <div onClick={()=>setOpenDept(isOpen?null:d.id)} style={{display:"flex", alignItems:"center", gap:10, cursor:"pointer"}}>
                <span style={{fontSize:22}}>{d.emoji}</span>
                <div style={{flex:1}}>
                  <div style={{fontSize:14, fontWeight:800, color:C.txt}}>{d.nom}</div>
                  <div style={{fontSize:10.5, color:C.txtMuted}}>{enseignantsDept.length} enseignant{enseignantsDept.length>1?"s":""} · {deptMatieres.length} matière{deptMatieres.length>1?"s":""}</div>
                </div>
                {tauxProg!==null&&<span style={{fontSize:11,fontWeight:700,padding:"3px 8px",borderRadius:10,
                  background:tauxProg>=75?"#f0fdf4":tauxProg>=50?"#fefce8":"#fef2f2",
                  color:tauxProg>=75?"#15803d":tauxProg>=50?"#92400e":"#b91c1c"}}>{tauxProg}%</span>}
                <span style={{fontSize:12, color:C.txtMuted}}>{openDept===d.id?"▲":"▼"}</span>
              </div>

              {openDept===d.id && (
                <div style={{marginTop:14, paddingTop:14, borderTop:`1px solid ${C.border}`, display:"flex", flexDirection:"column", gap:14}}>

                  {/* Animateur */}
                  <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:"#f8fafc",borderRadius:10}}>
                    <Avatar ens={animateur||{nom:""}} size={32} fontSize={11}/>
                    <div>
                      <div style={{fontSize:10,fontWeight:700,color:C.txtMuted,textTransform:"uppercase",letterSpacing:".06em"}}>Animateur pédagogique</div>
                      <div style={{fontSize:12,fontWeight:700,color:C.txt}}>{animateur?.nom||"Non assigné"}</div>
                    </div>
                  </div>

                  {/* KPIs */}
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
                    {[
                      {label:"Enseignants",value:enseignantsDept.length,color:"#3b82f6"},
                      {label:"Épreuves soumises",value:epSoumises,color:C.green},
                      {label:"Couverture prog.",value:tauxProg!==null?tauxProg+"%":"—",color:tauxProg>=75?"#15803d":tauxProg>=50?"#d97706":"#b91c1c"},
                    ].map((k,i)=>(
                      <div key={i} style={{background:"#f8fafc",borderRadius:10,padding:"10px 12px",border:`1px solid ${C.border}`,textAlign:"center"}}>
                        <div style={{fontSize:18,fontWeight:800,color:k.color}}>{k.value}</div>
                        <div style={{fontSize:9,color:C.txtMuted,fontWeight:600,marginTop:2}}>{k.label}</div>
                      </div>
                    ))}
                  </div>

                  {/* Enseignants du département */}
                  {enseignantsDept.length>0&&(
                    <div>
                      <div style={{fontSize:10,fontWeight:700,color:C.txtMuted,textTransform:"uppercase",letterSpacing:".06em",marginBottom:6}}>Enseignants</div>
                      <div style={{display:"flex",flexDirection:"column",gap:6}}>
                        {enseignantsDept.map(u=>(
                          <div key={u.id} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",background:"#f8fafc",borderRadius:8}}>
                            <Avatar ens={u} size={26} fontSize={9}/>
                            <div style={{flex:1,minWidth:0}}>
                              <div style={{fontSize:11.5,fontWeight:700,color:C.txt,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{u.nom}</div>
                              <div style={{fontSize:10,color:C.txtMuted}}>{(u.classes||[]).join(", ")||"Aucune classe"}</div>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Matières + gestion */}
                  <div>
                    <div style={{fontSize:10,fontWeight:700,color:C.txtMuted,textTransform:"uppercase",letterSpacing:".06em",marginBottom:6}}>Matières</div>
                    {loading ? <Sk h={16} w="60%"/> : deptMatieres.length===0 ? (
                      <div style={{fontSize:11, color:C.txtLight, fontStyle:"italic"}}>Aucune matière</div>
                    ) : deptMatieres.map(m => (
                      <div key={m.id} style={{display:"flex", alignItems:"center", gap:8, padding:"6px 8px", background:"#f8fafc", borderRadius:7, marginBottom:4}}>
                        {editingNom===m.id ? (
                          <input autoFocus defaultValue={m.nom}
                            onBlur={e=>renommerMatiere(m, e.target.value)}
                            onKeyDown={e=>{ if(e.key==="Enter") e.target.blur(); if(e.key==="Escape") setEditingNom(null); }}
                            style={{flex:1, border:`1px solid ${C.green}`, borderRadius:5, padding:"3px 6px", fontSize:11.5, fontFamily:"inherit"}}/>
                        ) : (
                          <span onClick={()=>setEditingNom(m.id)} style={{flex:1, fontSize:11.5, color:C.txt, cursor:"pointer"}}>{m.nom}</span>
                        )}
                        <button onClick={()=>supprimerMatiere(m)} disabled={savingId===m.id}
                          style={{border:"none", background:"transparent", color:C.red, cursor:"pointer", fontSize:13, padding:2}}>
                          {savingId===m.id ? <Spinner size={11} color={C.red}/> : "🗑️"}
                        </button>
                      </div>
                    ))}
                    <div style={{display:"flex", gap:6, marginTop:6}}>
                      <input value={openDept===d.id?newMatiere:""} onChange={e=>setNewMatiere(e.target.value)}
                        onKeyDown={e=>{ if(e.key==="Enter") ajouterMatiere(d.id); }}
                        placeholder="Nouvelle matière…"
                        style={{flex:1, border:`1px solid ${C.border}`, borderRadius:6, padding:"6px 8px", fontSize:11.5, fontFamily:"inherit"}}/>
                      <button onClick={()=>ajouterMatiere(d.id)} disabled={savingId===`new-${d.id}`}
                        style={{padding:"6px 12px", borderRadius:6, border:"none", background:C.green, color:"#fff", fontSize:11.5, fontWeight:700, cursor:"pointer", fontFamily:"inherit"}}>
                        {savingId===`new-${d.id}` ? <Spinner size={11}/> : "+ Ajouter"}
                      </button>
                    </div>
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

function ChangePasswordPage() {
  const {user, setUser, data, setData, showToast} = useApp();
  const {isMobile} = useDevice();
  const [oldPw,  setOldPw]  = useState("");
  const [newPw,  setNewPw]  = useState("");
  const [conf,   setConf]   = useState("");
  const [saving, setSaving] = useState(false);
  const [done,   setDone]   = useState(false);
  const [photoFile, setPhotoFile] = useState(null);
  const [photoPreview, setPhotoPreview] = useState(null);
  const [uploadingPhoto, setUploadingPhoto] = useState(false);
  const fileInputRef = useRef(null);
  const photoActuelle = user?.photo || data?.users?.[user?.id]?.photo || null;

  const choisirPhoto = (file) => {
    if (!file) return;
    if (!file.type.startsWith("image/")) { showToast("⚠ Choisis une image (jpg, png…)", false); return; }
    if (file.size > 8*1024*1024) { showToast("⚠ Image trop lourde (max 8 Mo)", false); return; }
    setPhotoFile(file);
    setPhotoPreview(URL.createObjectURL(file));
  };

  const enregistrerPhoto = async () => {
    if (!photoFile) return;
    setUploadingPhoto(true);
    try {
      const compressed = await resizeImageFile(photoFile, 300, 0.82);
      const path = `${user.id}_${Date.now()}.jpg`;
      const ok = await sb.uploadPhoto(path, compressed);
      if (ok) {
        const upd = await sb.rpc("update_teacher_photo", { p_id: user.id, p_photo: path });
        if (upd) {
          setUser(prev => ({...prev, photo:path}));
          setData(prev => ({...prev, users:{...(prev?.users||{}), [user.id]:{...(prev?.users?.[user.id]||{}), photo:path}}}));
          setPhotoFile(null); setPhotoPreview(null);
          showToast("✓ Photo de profil mise à jour");
        } else showToast(`⚠ Non enregistrée : ${(sb.lastError||"erreur inconnue").slice(0,80)}`, false);
      } else showToast("⚠ Échec de l'envoi de la photo", false);
    } catch {
      showToast("⚠ Erreur lors du traitement de la photo", false);
    }
    setUploadingPhoto(false);
  };

  const retirerPhoto = async () => {
    setUploadingPhoto(true);
    const ok = await sb.rpc("update_teacher_photo", { p_id: user.id, p_photo: null });
    if (ok) {
      setUser(prev => ({...prev, photo:null}));
      setData(prev => ({...prev, users:{...(prev?.users||{}), [user.id]:{...(prev?.users?.[user.id]||{}), photo:null}}}));
      setPhotoFile(null); setPhotoPreview(null);
      showToast("✓ Photo retirée");
    } else showToast("⚠ Erreur lors du retrait", false);
    setUploadingPhoto(false);
  };

  const handleChange = async() => {
    if(!oldPw) return showToast("Saisissez votre mot de passe actuel", false);
    if(newPw.length < 6) return showToast("Le nouveau mot de passe doit faire au moins 6 caractères", false);
    if(newPw !== conf) return showToast("Les mots de passe ne correspondent pas", false);
    if(newPw === oldPw) return showToast("Le nouveau mot de passe doit être différent", false);

    setSaving(true);
    const ok = await sb.rpc("change_password", {p_id: user.id, p_old_mdp: oldPw, p_new_mdp: newPw});
    setSaving(false);

    if(ok) {
      setDone(true);
      showToast("✓ Mot de passe modifié avec succès");
      setOldPw(""); setNewPw(""); setConf("");
    } else {
      showToast("⚠ Mot de passe actuel incorrect ou erreur serveur", false);
    }
  };

  return (
    <div style={{padding:20, maxWidth:440, margin:"0 auto", display:"flex", flexDirection:"column", gap:16}}>
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"18px 20px"}}>
        <h2 style={{fontSize:16, fontWeight:800, color:C.txt, margin:"0 0 4px"}}>🔐 Changer mon mot de passe</h2>
        <p style={{fontSize:12, color:C.txtMuted, margin:0}}>{user?.nom} · {user?.role}</p>
      </div>

      {/* Photo de profil */}
      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding: isMobile?"18px 16px":"18px 20px"}}>
        <h2 style={{fontSize:14, fontWeight:800, color:C.txt, margin:"0 0 14px"}}>📷 Ma photo de profil</h2>
        <div style={{display:"flex", flexDirection: isMobile?"column":"row", alignItems: isMobile?"center":"center", gap: isMobile?14:16}}>
          <div style={{width: isMobile?88:72, height: isMobile?88:72, borderRadius:"50%",overflow:"hidden",flexShrink:0,background:"#f1f5f9",display:"flex",alignItems:"center",justifyContent:"center",border:`1.5px solid ${C.border}`}}>
            {photoPreview ? (
              <img src={photoPreview} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/>
            ) : photoActuelle ? (
              <img src={sb.photoUrl(photoActuelle)} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/>
            ) : (
              <span style={{fontSize: isMobile?30:26,color:"#cbd5e1"}}>👤</span>
            )}
          </div>
          <div style={{display:"flex", flexDirection:"column", gap:8, width: isMobile?"100%":"auto", alignItems: isMobile?"stretch":"flex-start"}}>
            <input ref={fileInputRef} type="file" accept="image/*" style={{display:"none"}}
              onChange={e=>choisirPhoto(e.target.files?.[0])}/>
            {photoPreview ? (
              <>
                <button type="button" disabled={uploadingPhoto} onClick={enregistrerPhoto}
                  style={{padding: isMobile?"12px 16px":"8px 16px",borderRadius:9,border:"none",background:C.green,color:"#fff",fontSize: isMobile?13.5:12.5,fontWeight:700,cursor:uploadingPhoto?"not-allowed":"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:6,width:"100%"}}>
                  {uploadingPhoto ? <><Spinner size={12} color="#fff"/> Envoi…</> : "✓ Enregistrer cette photo"}
                </button>
                <button type="button" onClick={()=>{setPhotoFile(null);setPhotoPreview(null);}}
                  style={{padding: isMobile?"10px":"6px 10px",borderRadius:8,border:"none",background:"transparent",color:C.txtMuted,fontSize: isMobile?12.5:11.5,cursor:"pointer",fontFamily:"inherit",textAlign:"center",width:"100%"}}>
                  Annuler
                </button>
              </>
            ) : (
              <>
                <button type="button" onClick={()=>fileInputRef.current?.click()}
                  style={{padding: isMobile?"12px 16px":"8px 16px",borderRadius:9,border:`1.5px solid ${C.green}`,background:C.greenPale,color:C.green,fontSize: isMobile?13.5:12.5,fontWeight:700,cursor:"pointer",fontFamily:"inherit",width:"100%",textAlign:"center"}}>
                  📷 {photoActuelle ? "Changer la photo" : "Ajouter une photo"}
                </button>
                {photoActuelle && (
                  <button type="button" disabled={uploadingPhoto} onClick={retirerPhoto}
                    style={{padding: isMobile?"10px":"6px 10px",borderRadius:8,border:"none",background:"transparent",color:C.txtMuted,fontSize: isMobile?12.5:11.5,cursor:uploadingPhoto?"not-allowed":"pointer",fontFamily:"inherit",textAlign:"center",width:"100%"}}>
                    Retirer la photo
                  </button>
                )}
              </>
            )}
          </div>
        </div>
      </div>

      {done && (
        <div style={{background:C.greenPale, border:`1px solid ${C.greenBorder}`, borderRadius:12, padding:"14px 16px", display:"flex", gap:10, alignItems:"center"}}>
          <span style={{fontSize:20}}>✅</span>
          <span style={{fontSize:13, fontWeight:600, color:C.green}}>Mot de passe modifié avec succès</span>
        </div>
      )}

      <div style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:"18px 20px", display:"flex", flexDirection:"column", gap:14}}>
        {[
          {label:"Mot de passe actuel", val:oldPw, set:setOldPw, placeholder:"Votre mot de passe actuel"},
          {label:"Nouveau mot de passe", val:newPw, set:setNewPw, placeholder:"Minimum 6 caractères"},
          {label:"Confirmer le nouveau", val:conf, set:setConf, placeholder:"Répéter le nouveau mot de passe"},
        ].map(({label, val, set, placeholder}, i) => (
          <div key={i}>
            <label style={{display:"block", fontSize:11, fontWeight:700, color:C.txtMuted, marginBottom:5, textTransform:"uppercase", letterSpacing:".06em"}}>
              {label}
            </label>
            <input type="password" value={val} onChange={e=>set(e.target.value)}
              onKeyDown={e=>e.key==="Enter"&&handleChange()}
              placeholder={placeholder} maxLength={50}
              style={{width:"100%", padding:"10px 14px", border:`1.5px solid ${C.border}`, borderRadius:9, fontSize:13, color:C.txt, fontFamily:"inherit"}}
              onFocus={e=>{e.target.style.borderColor=C.green;e.target.style.background=C.white;}}
              onBlur={e=>{e.target.style.borderColor=C.border;e.target.style.background="#f8fafc";}}/>
          </div>
        ))}

        {/* Indicateur force mdp */}
        {newPw.length > 0 && (
          <div>
            <div style={{display:"flex", gap:4, marginBottom:4}}>
              {[...Array(4)].map((_,i) => {
                const strength = newPw.length >= 8 ? 4 : newPw.length >= 6 ? 2 : 1;
                return <div key={i} style={{flex:1, height:4, borderRadius:2, background:i<strength?(strength>=4?C.green:strength>=2?C.amber:C.red):"#e2e8f0"}}/>;
              })}
            </div>
            <div style={{fontSize:10, color:C.txtMuted}}>
              {newPw.length < 6 ? "Trop court" : newPw.length < 8 ? "Acceptable" : "Fort"}
            </div>
          </div>
        )}

        <button onClick={handleChange} disabled={saving}
          style={{padding:"12px", borderRadius:10, border:"none",
            background:saving?"#94a3b8":`linear-gradient(135deg,${C.greenDark},${C.green})`,
            color:"#fff", fontSize:14, fontWeight:700, cursor:saving?"not-allowed":"pointer",
            fontFamily:"inherit", display:"flex", alignItems:"center", justifyContent:"center", gap:7}}>
          {saving ? <><Spinner size={16} color="#fff"/>&nbsp;Modification…</> : "✓ Modifier mon mot de passe"}
        </button>
      </div>

      <div style={{background:"#fef9e7", border:"1px solid #f9ca24", borderRadius:10, padding:"12px 16px", fontSize:12, color:"#7d6608"}}>
        <strong>Conseils :</strong> Utilisez un mot de passe d'au moins 8 caractères avec des chiffres et des lettres. Ne partagez jamais votre mot de passe.
      </div>
    </div>
  );
}

// ─── Placeholder ────
const AppLayout = ({onLogout}) => {
  const {user,page,data,setData,toast,showToast} = useApp();
  const [collapsed,setCollapsed] = useState(()=>window.innerWidth < 1024);
  const isAdmin = isAdminRole(user?.role);
  // Pages qui gèrent leur propre flex+scroll (retournent flex:1, overflow:hidden)
  const SELF_SCROLL = ["mes-classes","cahier","documents","eleves"];
  const renderPage = () => {
    // ── Pages AUTO-SCROLLANTES (flex:1 + overflow:hidden sur leur root) ──
    if(page==="mes-classes") return <MesClassesPage/>
    if(page==="cahier")      return <CahierDeTextePage/>
    if(page==="documents")   return isAdmin?<DocumentsPage/>:null
    if(page==="eleves")      return (isAdmin||user?.role==="censeur")?<ElevesPage/>:<MesClassesPage/>
    // ── Pages SIMPLES — enveloppées dans un scroller ───────────────────
    const W = ({children}) => (
      <div style={{flex:1, minHeight:0, overflowY:"auto"}}>
        {children}
      </div>
    );
    if(page==="dashboard")         return <W>{user?.role==="proviseur"?<DashboardProviseur/>:user?.role==="surveillant_general"?<DashboardSurveillance/>:user?.role==="censeur"?<DashboardCenseur/>:(user?.role==="animateur"||user?.role==="animatrice")?<DashboardAnimateur/>:isAdmin?<DashboardAdmin/>:<DashboardTeacher/>}</W>
    if(page==="programme")         return <W>{(isAdmin||user?.role==="censeur")?<SuiviProgrammePage/>:<MonProgrammePage/>}</W>
    if(page==="epreuves")          return <W><EpreuvesPage/></W>
    if(page==="edt-teacher")       return <W><MonEdtPage/></W>
    if(page==="edt")               return <W>{(isAdmin||user?.role==="censeur")?<EdtPage/>:<MonEdtPage/>}</W>
    if(page==="enseignants")       return <W>{isAdmin?<EnseignantsPage/>:null}</W>
    if(page==="gestion-annuelle")  return <W>{isAdmin?<GestionAnnuellePage/>:null}</W>
    if(page==="departements")      return <W>{(user?.role==="proviseur"||user?.role==="censeur"||user?.role==="animateur"||user?.role==="animatrice")?<DepartementsPage/>:null}</W>
    if(page==="bulletins")          return <W><BulletinsPage/></W>
    if(page==="suivi-prog-dept")    return <W><SuiviProgrammePage/></W>
    if(page==="fiche-inspection")   return <W>{(user?.role==="animateur"||user?.role==="animatrice")?<FicheInspectionPage/>:null}</W>
    if(page==="documents-ap")       return <W>{(user?.role==="animateur"||user?.role==="animatrice")?<DocumentsAnimateurPage/>:null}</W>
    if(page==="settings")          return <W><ChangePasswordPage/></W>
    return <W><PlaceholderPage title={PAGE_TITLES[page]||page} emoji="🚧"/></W>
  };
  return(
    <div style={{display:"flex",height:"100vh",overflow:"hidden",background:C.bg}}>
      <Sidebar collapsed={collapsed} setCollapsed={setCollapsed}/>
      <div style={{flex:1,display:"flex",flexDirection:"column",minWidth:0,overflow:"hidden"}}>
        <Topbar title={PAGE_TITLES[page]||"—"} onLogout={onLogout} collapsed={collapsed} setCollapsed={setCollapsed}/>
        <div style={{flex:1, minHeight:0, overflow:"hidden", display:"flex", flexDirection:"column"}}>
          {renderPage()}
        </div>
      </div>
      {toast && <Toast msg={toast.msg} ok={toast.ok}/>}
    </div>
  );
};

// ─── Splash ───────────────────────────────────────────────────────
function Splash({onDone}){
  const [prog,setProg]=useState(0);
  const [label,setLabel]=useState("Initialisation…");
  useEffect(()=>{
    const steps=[[300,20,"Connexion Supabase…"],[700,55,"Chargement des données…"],[1100,85,"Vérification des droits…"],[1500,100,"Prêt !"]];
    const timers=steps.map(([d,p,l])=>setTimeout(()=>{setProg(p);setLabel(l);},d));
    const done=setTimeout(onDone,1900);
    return()=>{timers.forEach(clearTimeout);clearTimeout(done);};
  },[]);
  return(
    <div style={{minHeight:"100vh",background:`linear-gradient(160deg,rgba(12,61,36,.46),rgba(22,163,74,.4) 60%,rgba(55,168,102,.36)), url(${SPLASH_BG_B64})`,backgroundSize:"cover",backgroundPosition:"center",backgroundRepeat:"no-repeat",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:28}}>
      <div style={{animation:"logoIn .7s cubic-bezier(.2,.8,.2,1)"}}><img src={LOGO_LYCEE_B64} alt="Lykama" width={100} height={100} style={{borderRadius:"50%", objectFit:"contain"}}/></div>
      <div style={{textAlign:"center"}}>
        <h1 style={{fontFamily:"'Playfair Display',Georgia,serif",fontSize:36,color:"#fff",letterSpacing:".06em",margin:"0 0 8px"}}>Lykama</h1>
        <p style={{color:"rgba(255,255,255,.55)",fontSize:13}}>Lycée de Kakatare · Maroua · Cameroun</p>
      </div>
      <div style={{width:220}}>
        <div style={{height:3,background:"rgba(255,255,255,.15)",borderRadius:3,overflow:"hidden"}}><div style={{height:"100%",width:`${prog}%`,background:C.gold,borderRadius:3,transition:"width .4s ease"}}/></div>
        <p style={{textAlign:"center",color:"rgba(255,255,255,.4)",fontSize:11,marginTop:8}}>{label}</p>
      </div>
    </div>
  );
}

// ─── Login ────────────────────────────────────────────────────────
function LoginPage({onLogin}){
  const {isMobile} = useDevice();

  /* ── Auth state (inchangé) ──────────────────────────── */
  const [id,setId]           = useState(()=>localStorage.getItem("svt_remember_id")||"");
  const [pw,setPw]           = useState("");
  const [showPw,setShowPw]   = useState(false);
  const [err,setErr]         = useState("");
  const [loading,setLoading] = useState(false);
  const [rememberMe,setRememberMe] = useState(()=>!!localStorage.getItem("svt_remember_id"));
  const [selDept,setSelDept] = useState("");

  /* ── Compteurs réels ────────────────────────────────── */
  const [counter,setCounter]   = useState({eleves:0,ens:0});
  const [ensCountReal,setEnsCountReal] = useState(null);
  useEffect(()=>{
    sb.get("utilisateurs","?select=id&role=in.(enseignant,animateur)").then(rows=>{if(rows)setEnsCountReal(rows.length);});
  },[]);
  useEffect(()=>{
    if(ensCountReal===null)return;
    const targets={eleves:getTotalEleves(),ens:ensCountReal};
    let step=0;const steps=60;
    const timer=setInterval(()=>{
      step++;const p=Math.min(step/steps,1);const e=1-Math.pow(1-p,3);
      setCounter({eleves:Math.round(targets.eleves*e),ens:Math.round(targets.ens*e)});
      if(step>=steps)clearInterval(timer);
    },1800/60);
    return()=>clearInterval(timer);
  },[ensCountReal]);

  /* ── Navigation portail ─────────────────────────────── */
  const [portalStep,setPortalStep] = useState(0); // 0=landing, 1=portail

  /* ── Profil sélectionné ─────────────────────────────── */
  const [selProfile,setSelProfile] = useState("enseignant");
  const [mobileFormOpen,setMobileFormOpen] = useState(false);
  const [selNiveau,setSelNiveau] = useState("");

  const SG_NIVEAUX = [
    {label:"6ème",        id:"sg_6eme"},
    {label:"5ème",        id:"sg_5eme"},
    {label:"4ème",        id:"sg_4eme"},
    {label:"3ème",        id:"sg_3eme"},
    {label:"2nde",        id:"sg_2nde"},
    {label:"1ère et Tle", id:"sg_lycee"},
    {label:"Toute école", id:"surveillance"},
  ];

  const PROFILES = [
    {key:"direction",  label:"Direction",             sub:"Proviseur",             desc:"Gestion globale de l'établissement", emoji:"👨🏾‍💼", role:"proviseur",          needsDept:false},
    {key:"censeur",    label:"Censeur",               sub:"Organisation pédagogique", desc:"Classes, notes, suivi discipline",  emoji:"📚",  role:"censeur",            needsDept:false},
    {key:"sg",         label:"Surveillance Générale", sub:"Vie scolaire & discipline", desc:"Absences, retards, incidents",    emoji:"🛡️", role:"surveillant_general", needsDept:false},
    {key:"animateur",  label:"Animateur Pédagogique", sub:"Suivi pédagogique",     desc:"Supervision du département & programmes", emoji:"📋", role:"animateur", needsDept:false},
    {key:"enseignant", label:"Enseignant",            sub:"Corps professoral",      desc:"Cours, cahier de textes, évaluations", emoji:"👨🏾‍🏫", role:"enseignant",   needsDept:true},
    {key:"eleve",      label:"Élève",                 sub:"Espace apprenant",       desc:"Résultats, emploi du temps",         emoji:"🎓",  role:null, soon:true},
    {key:"parent",     label:"Parent",                sub:"Suivi scolaire",         desc:"Suivi scolaire de l'élève",          emoji:"👪",  role:null, soon:true},
  ];

  const profile = PROFILES.find(p=>p.key===selProfile);

  /* ── Soumission (logique inchangée) ─────────────────── */
  const submit = async()=>{
    setErr("");
    if(!id.trim()){setErr("Veuillez saisir votre identifiant.");return;}
    if(!pw){setErr("Veuillez saisir votre mot de passe.");return;}
    if(profile.needsDept&&!selDept){setErr("Sélectionnez un département.");return;}
    setLoading(true);
    let authUser=null;
    try{ authUser=await sb.rpc("authenticate_user",{p_id:id.trim().toLowerCase(),p_mdp:pw}); }
    catch(e){ authUser=null; }
    if(!authUser){setErr("Identifiant ou mot de passe incorrect.");setLoading(false);return;}
    if(profile.role==="proviseur"          && authUser.role!=="proviseur")          {setErr("Ce compte n'est pas un compte Direction.");setLoading(false);return;}
    if(profile.role==="censeur"            && authUser.role!=="censeur")            {setErr("Ce compte n'est pas un compte Censeur.");setLoading(false);return;}
    if(profile.role==="animateur" && authUser.role!=="animateur" && authUser.role!=="animatrice"){setErr("Ce compte n'est pas un compte Animateur Pédagogique.");setLoading(false);return;}
    if(profile.role==="surveillant_general"&& authUser.role!=="surveillant_general"){setErr("Ce compte n'est pas un compte Surveillance Générale.");setLoading(false);return;}
    if(profile.role==="enseignant" && (authUser.role==="proviseur"||authUser.role==="censeur"||authUser.role==="surveillant_general")){setErr("Utilisez l'accès correspondant à ce compte.");setLoading(false);return;}
    if(profile.needsDept && authUser.departement_id && String(authUser.departement_id)!==String(selDept)){setErr("Ce compte n'appartient pas à ce département.");setLoading(false);return;}
    if(rememberMe)localStorage.setItem("svt_remember_id",id.trim().toLowerCase());
    else localStorage.removeItem("svt_remember_id");
    const demoRef=DEMO_ACCOUNTS.find(a=>a.id===authUser.id)||{};
    onLogin({...demoRef,...authUser,col:getColor(authUser.id),ini:getIni(authUser.nom),mustChangePwd:!!authUser.must_change_pwd});
  };

  /* ── Couleurs tokens ─────────────────────────────────── */
  const clr = {
    forest:"#0B4D2C", forestDark:"#083D22", forestLight:"#E8F5EE",
    gold:"#D4AF37", goldLight:"#FBF5D8",
    navy:"#0F172A", slate:"#64748B", slateLight:"#F1F5F9",
    white:"#FFFFFF", border:"#E2E8F0",
  };

  /* ════════════════════════════════════════════════
     ÉTAPE 0 : LANDING
  ════════════════════════════════════════════════ */
  if(portalStep===0) return (
    <div style={{
      minHeight:"100vh", overflowY:"auto", position:"relative",
      background:"#0B3D20",
      fontFamily:"'Plus Jakarta Sans',sans-serif",
    }}>
      <style>{`
        @keyframes fadeDown{from{opacity:0;transform:translateY(-16px);}to{opacity:1;transform:none;}}
        @keyframes fadeUp{from{opacity:0;transform:translateY(20px);}to{opacity:1;transform:none;}}
        @keyframes scaleIn{from{opacity:0;transform:scale(.92);}to{opacity:1;transform:scale(1);}}
        .lp-btn-gold:hover{filter:brightness(1.07);transform:translateY(-2px);box-shadow:0 12px 32px rgba(212,175,55,.45)!important;}
        .lp-btn-gold{transition:all .22s ease;}
        .lp-feat:hover{background:rgba(255,255,255,.07)!important;}
      `}</style>

      {/* Photo fond (droite, fondue) */}
      <div style={{
        position:"fixed",top:0,right:0,width:"55%",height:"100%",
        background:`url(${LOGIN_BG_B64}) center/cover no-repeat`,
        WebkitMaskImage:"linear-gradient(to left, rgba(0,0,0,0.45) 0%, transparent 80%)",
        maskImage:"linear-gradient(to left, rgba(0,0,0,0.45) 0%, transparent 80%)",
        zIndex:0,pointerEvents:"none",
      }}/>

      {/* Overlay vert + motif losanges */}
      <div style={{position:"fixed",inset:0,zIndex:1,pointerEvents:"none",
        background:"linear-gradient(135deg, rgba(11,61,32,.97) 0%, rgba(11,61,32,.88) 55%, rgba(11,61,32,.72) 100%)",
      }}/>
      <svg style={{position:"fixed",inset:0,width:"100%",height:"100%",zIndex:1,opacity:.12,pointerEvents:"none"}} xmlns="http://www.w3.org/2000/svg">
        <defs>
          <pattern id="diamonds" x="0" y="0" width="40" height="40" patternUnits="userSpaceOnUse">
            <polygon points="20,2 38,20 20,38 2,20" fill="none" stroke="#D4AF37" strokeWidth="0.8"/>
          </pattern>
        </defs>
        <rect width="100%" height="100%" fill="url(#diamonds)"/>
      </svg>

      {/* Contenu scrollable */}
      <div style={{position:"relative",zIndex:2,display:"flex",flexDirection:"column",alignItems:"center",padding:isMobile?"28px 20px 40px":"40px 24px 60px",textAlign:"center"}}>

        {/* En-tête République */}
        <div style={{animation:"fadeDown .6s ease",marginBottom:20}}>
          <svg viewBox="0 0 45 30" width={isMobile?32:38} height={isMobile?21:25} style={{borderRadius:3,boxShadow:"0 2px 8px rgba(0,0,0,.35)",margin:"0 auto 10px",display:"block"}}>
            <rect width="15" height="30" fill="#007A5E"/>
            <rect x="15" width="15" height="30" fill="#CE1126"/>
            <rect x="30" width="15" height="30" fill="#FCD116"/>
            <polygon points="22.5,8 23.9,12.6 28.7,12.6 24.9,15.4 26.3,20 22.5,17.2 18.7,20 20.1,15.4 16.3,12.6 21.1,12.6" fill="#FCD116"/>
          </svg>
          <p style={{fontSize:isMobile?10:11,fontWeight:800,letterSpacing:".18em",textTransform:"uppercase",color:"rgba(255,255,255,.9)",margin:"0 0 3px"}}>
            République du Cameroun
          </p>
          <p style={{fontSize:isMobile?10:11,fontWeight:500,color:"#D4AF37",margin:0,letterSpacing:".06em"}}>
            Paix • Travail • Patrie
          </p>
        </div>

        {/* Logo */}
        <div style={{
          width:isMobile?110:130, height:isMobile?110:130, borderRadius:"50%",
          border:"3px solid #D4AF37",
          boxShadow:"0 0 0 6px rgba(212,175,55,.18), 0 12px 40px rgba(0,0,0,.5)",
          overflow:"hidden", background:"#fff",
          display:"flex",alignItems:"center",justifyContent:"center",
          marginBottom:24, flexShrink:0,
          animation:"scaleIn .7s ease .1s both",
        }}>
          <img src={LOGO_LYCEE_B64} alt="Logo" style={{width:"90%",height:"90%",objectFit:"contain"}}/>
        </div>

        {/* Titre */}
        <div style={{animation:"fadeUp .7s ease .2s both",marginBottom:16}}>
          <h1 style={{
            fontSize:isMobile?"clamp(32px,8vw,42px)":"clamp(38px,5vw,54px)",
            fontFamily:"'Playfair Display',serif",
            fontWeight:800, color:"#fff",
            lineHeight:1.1, margin:"0 0 12px",
            textShadow:"0 2px 20px rgba(0,0,0,.4)",
          }}>
            Lycée de<br/>Kakatare – Maroua
          </h1>
          <p style={{fontSize:isMobile?14:16,color:"rgba(255,255,255,.8)",margin:"0 0 20px",lineHeight:1.5,maxWidth:480}}>
            Plateforme numérique de gestion<br/>et de suivi scolaire
          </p>

          {/* Pill année */}
          <div style={{display:"inline-flex",alignItems:"center",gap:8,border:"1.5px solid #D4AF37",borderRadius:30,padding:"8px 20px",marginBottom:28}}>
            <span style={{fontSize:14}}>📅</span>
            <span style={{fontSize:isMobile?11:12,fontWeight:800,letterSpacing:".1em",textTransform:"uppercase",color:"#D4AF37"}}>
              Année scolaire 2025 – 2026
            </span>
          </div>
        </div>

        {/* CTA principal */}
        <div style={{animation:"fadeUp .6s ease .35s both",width:"100%",maxWidth:420,marginBottom:18}}>
          <button
            className="lp-btn-gold"
            onClick={()=>setPortalStep(1)}
            style={{
              width:"100%", padding:isMobile?"18px 24px":"20px 32px",
              background:"linear-gradient(135deg,#D4AF37,#b8860b)",
              border:"none", borderRadius:14, cursor:"pointer",
              display:"flex", alignItems:"center", justifyContent:"center", gap:12,
              boxShadow:"0 8px 28px rgba(212,175,55,.35)",
            }}>
            <span style={{fontSize:20}}>🔒</span>
            <span style={{fontSize:isMobile?15:17,fontWeight:800,letterSpacing:".06em",textTransform:"uppercase",color:"#0B3D20"}}>
              Accéder au Portail
            </span>
            <span style={{fontSize:18,color:"#0B3D20",fontWeight:900}}>→</span>
          </button>
          <p style={{fontSize:12,color:"rgba(255,255,255,.5)",margin:"10px 0 0"}}>
            Accès réservé au personnel autorisé
          </p>
        </div>

        {/* 3 rôles */}
        <div style={{display:"flex",gap:isMobile?16:28,marginBottom:36,animation:"fadeUp .6s ease .45s both",flexWrap:"wrap",justifyContent:"center"}}>
          {[["🛡️","Administration"],["🎓","Enseignants"],["👥","Vie scolaire"]].map(([ico,lbl])=>(
            <div key={lbl} style={{display:"flex",alignItems:"center",gap:7,color:"rgba(255,255,255,.7)",fontSize:isMobile?12:13,fontWeight:600}}>
              <span style={{fontSize:16}}>{ico}</span>{lbl}
            </div>
          ))}
        </div>

        {/* Carte stats */}
        <div style={{
          width:"100%", maxWidth:600,
          background:"rgba(5,30,15,.75)",
          backdropFilter:"blur(8px)",
          border:"1px solid rgba(212,175,55,.22)",
          borderRadius:16, padding:isMobile?"20px 16px":"28px 32px",
          marginBottom:16,
          animation:"fadeUp .7s ease .55s both",
        }}>
          <p style={{fontSize:isMobile?10:11,fontWeight:800,letterSpacing:".18em",textTransform:"uppercase",color:"#D4AF37",margin:"0 0 22px"}}>
            Notre établissement en chiffres
          </p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:isMobile?12:20}}>
            {[
              {ico:"👥", val:counter.eleves||getTotalEleves(), lbl:"Élèves inscrits",   color:"#4ade80"},
              {ico:"🎓", val:counter.ens||0,                  lbl:"Enseignants",        color:"#D4AF37"},
              {ico:"🏫", val:CLASSES_REELLES.length,          lbl:"Classes",            color:"#60a5fa"},
              {ico:"🏛️", val:52,                              lbl:"Personnel\nadministratif", color:"#f472b6"},
            ].map(({ico,val,lbl,color})=>(
              <div key={lbl} style={{textAlign:"center"}}>
                <div style={{fontSize:isMobile?22:28,marginBottom:4}}>{ico}</div>
                <div style={{fontSize:isMobile?20:26,fontWeight:900,color,lineHeight:1,fontFamily:"'Playfair Display',serif"}}>{val}</div>
                <div style={{fontSize:isMobile?9:10,color:"rgba(255,255,255,.55)",marginTop:4,whiteSpace:"pre-line",lineHeight:1.3}}>{lbl}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Carte fonctionnalités */}
        <div style={{
          width:"100%", maxWidth:600,
          background:"rgba(5,30,15,.75)",
          backdropFilter:"blur(8px)",
          border:"1px solid rgba(212,175,55,.22)",
          borderRadius:16, padding:isMobile?"20px 16px":"28px 32px",
          marginBottom:28,
          animation:"fadeUp .7s ease .65s both",
        }}>
          <p style={{fontSize:isMobile?10:11,fontWeight:800,letterSpacing:".18em",textTransform:"uppercase",color:"#D4AF37",margin:"0 0 20px"}}>
            Une plateforme pour mieux gérer
          </p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:isMobile?12:16}}>
            {[
              {ico:"📖","lbl":"Suivi\npédagogique"},
              {ico:"📝","lbl":"Notes &\névaluations"},
              {ico:"📅","lbl":"Emplois du\ntemps"},
              {ico:"⏰","lbl":"Absences &\nretards"},
              {ico:"📚","lbl":"Cahier de texte\nnumérique"},
              {ico:"🔔","lbl":"Communications"},
            ].map(({ico,lbl})=>(
              <div key={lbl} className="lp-feat" style={{padding:"12px 8px",borderRadius:10,transition:"all .2s",cursor:"default",background:"rgba(255,255,255,.04)"}}>
                <div style={{fontSize:isMobile?22:26,marginBottom:6}}>{ico}</div>
                <div style={{fontSize:isMobile?10:11,color:"rgba(255,255,255,.7)",whiteSpace:"pre-line",lineHeight:1.35,fontWeight:500}}>{lbl}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Footer */}
        <div style={{animation:"fadeUp .6s ease .75s both",textAlign:"center"}}>
          <p style={{fontSize:12,color:"rgba(255,255,255,.4)",margin:"0 0 4px",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
            🔒 Accès sécurisé • Données protégées
          </p>
          <p style={{fontSize:11,color:"rgba(255,255,255,.25)",margin:0}}>
            © Lycée de Kakatare – Maroua
          </p>
        </div>

      </div>
    </div>
  );

  /* ════════════════════════════════════════════════
     ÉTAPE 1 : PORTAIL (sélection profil + formulaire)
  ════════════════════════════════════════════════ */

  const formPanelJSX = (
    <div style={{background:clr.white,borderRadius:20,border:`1px solid ${clr.border}`,boxShadow:"0 8px 32px rgba(0,0,0,.1)",padding:"28px 24px",position:"relative",overflow:"hidden"}}>
      {/* Liseré dégradé */}
      <div style={{position:"absolute",top:0,left:0,right:0,height:4,background:`linear-gradient(90deg,${clr.forest},${clr.gold},${clr.navy})`}}/>

      <h3 style={{fontSize:17,fontWeight:800,color:clr.navy,margin:"0 0 4px",fontFamily:"'Playfair Display',serif"}}>
        Connexion {PROFILES.find(p=>p.key===selProfile)?.label||""}
      </h3>
      <p style={{fontSize:12,color:clr.slate,margin:"0 0 20px",fontFamily:"'Plus Jakarta Sans',sans-serif"}}>
        Saisissez vos identifiants pour accéder à votre espace
      </p>

      {err && (
        <div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:10,padding:"10px 14px",marginBottom:16,fontSize:13,color:"#b91c1c",fontFamily:"'Plus Jakarta Sans',sans-serif"}}>
          ⚠️ {err}
        </div>
      )}

      {/* Niveau SG */}
      {selProfile==="sg" && (
        <div style={{marginBottom:14}}>
          <label style={{display:"block",fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",color:clr.slate,marginBottom:6,fontFamily:"'Plus Jakarta Sans',sans-serif"}}>
            Niveau supervisé
          </label>
          <select
            value={selNiveau}
            onChange={e=>{setSelNiveau(e.target.value);const sg=SG_NIVEAUX.find(n=>n.label===e.target.value);if(sg)setId(sg.id);}}
            style={{width:"100%",padding:"11px 14px",border:"1.5px solid "+clr.border,borderRadius:10,fontSize:14,fontFamily:"'Plus Jakarta Sans',sans-serif",background:clr.slateLight,color:selNiveau?clr.navy:"#94a3b8",transition:"all .2s",boxSizing:"border-box",appearance:"none"}}>
            <option value="">— Sélectionner votre niveau —</option>
            {SG_NIVEAUX.map(n=>(
              <option key={n.id} value={n.label}>{n.label}</option>
            ))}
          </select>
        </div>
      )}

      {/* Identifiant */}
      <div style={{marginBottom:14}}>
        <label style={{display:"block",fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",color:clr.slate,marginBottom:6,fontFamily:"'Plus Jakarta Sans',sans-serif"}}>
          Matricule / Identifiant
        </label>
        <input
          className="lp-input"
          type="text" value={id} onChange={e=>setId(e.target.value)}
          onKeyDown={e=>{if(e.key==="Enter")submit();}}
          placeholder="Votre identifiant"
          style={{width:"100%",padding:"11px 14px",border:`1.5px solid ${clr.border}`,borderRadius:10,fontSize:14,fontFamily:"'Plus Jakarta Sans',sans-serif",background:clr.slateLight,color:clr.navy,transition:"all .2s",boxSizing:"border-box"}}
        />
      </div>

      {/* Département (seulement Enseignant) */}
      {profile.needsDept && (
        <div style={{marginBottom:14}}>
          <label style={{display:"block",fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",color:clr.slate,marginBottom:6,fontFamily:"'Plus Jakarta Sans',sans-serif"}}>
            Département
          </label>
          <select
            className="lp-input"
            value={selDept} onChange={e=>setSelDept(e.target.value)}
            style={{width:"100%",padding:"11px 14px",border:`1.5px solid ${clr.border}`,borderRadius:10,fontSize:14,fontFamily:"'Plus Jakarta Sans',sans-serif",background:clr.slateLight,color:selDept?clr.navy:"#94a3b8",transition:"all .2s",boxSizing:"border-box",appearance:"none"}}>
            <option value="">— Sélectionner —</option>
            {DEPARTEMENTS_LIST.map(d=>(
              <option key={d.id} value={d.id}>{d.emoji} {d.nom}</option>
            ))}
          </select>
        </div>
      )}

      {/* Mot de passe */}
      <div style={{marginBottom:20}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
          <label style={{fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",color:clr.slate,fontFamily:"'Plus Jakarta Sans',sans-serif"}}>
            Mot de passe
          </label>
        </div>
        <div style={{position:"relative"}}>
          <input
            className="lp-input"
            type={showPw?"text":"password"} value={pw} onChange={e=>setPw(e.target.value)}
            onKeyDown={e=>{if(e.key==="Enter")submit();}}
            placeholder="Votre mot de passe"
            style={{width:"100%",padding:"11px 42px 11px 14px",border:`1.5px solid ${clr.border}`,borderRadius:10,fontSize:14,fontFamily:"'Plus Jakarta Sans',sans-serif",background:clr.slateLight,color:clr.navy,transition:"all .2s",boxSizing:"border-box"}}
          />
          <button onClick={()=>setShowPw(!showPw)} type="button" style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:14,color:clr.slate}}>
            {showPw?"🙈":"👁️"}
          </button>
        </div>
      </div>

      {/* Se souvenir */}
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:18}}>
        <input type="checkbox" id="remember" checked={rememberMe} onChange={e=>setRememberMe(e.target.checked)} style={{accentColor:clr.forest,width:15,height:15,cursor:"pointer"}}/>
        <label htmlFor="remember" style={{fontSize:12,color:clr.slate,cursor:"pointer",fontFamily:"'Plus Jakarta Sans',sans-serif",userSelect:"none"}}>Se souvenir de moi</label>
      </div>

      {/* Bouton */}
      <button
        className="lp-btn"
        onClick={submit} disabled={loading}
        style={{
          width:"100%",padding:"13px",borderRadius:12,border:"none",cursor:loading?"not-allowed":"pointer",
          background:loading?"#94a3b8":`linear-gradient(135deg,${clr.forest},#125c34)`,
          color:"#fff",fontWeight:700,fontSize:15,fontFamily:"'Plus Jakarta Sans',sans-serif",
          boxShadow:loading?"none":`0 4px 16px ${clr.forest}40`,
          transition:"all .2s ease", display:"flex",alignItems:"center",justifyContent:"center",gap:8,
        }}>
        {loading?<><Spinner size={14} color="#fff"/> Connexion en cours…</>:"Se connecter →"}
      </button>

      <p style={{textAlign:"center",marginTop:14,fontSize:11,color:clr.slate,display:"flex",alignItems:"center",justifyContent:"center",gap:5,fontFamily:"'Plus Jakarta Sans',sans-serif"}}>
        🔒 Connexion sécurisée • Lycée de Kakatare
      </p>

      {/* Retour */}
      <button onClick={()=>{if(isMobile&&mobileFormOpen){setMobileFormOpen(false);}else{setPortalStep(0);}}} style={{display:"block",margin:"12px auto 0",background:"none",border:"none",cursor:"pointer",fontSize:12,color:clr.slate,fontFamily:"'Plus Jakarta Sans',sans-serif",fontWeight:600}}>
        ← Retour
      </button>
    </div>
  );

  /* ── Carte profil ───────────────────────────────────── */
  const ProfileCard = ({p})=>{
    const isSelected = selProfile===p.key;
    return (
      <div
        className={p.soon?"":"lp-card-hover"}
        tabIndex={p.soon?-1:0}
        onClick={()=>{
          if(p.soon)return;
          setSelProfile(p.key); setErr(""); setSelNiveau(""); setSelNiveau("");
          if(isMobile)setMobileFormOpen(true);
        }}
        onKeyDown={e=>{if((e.key==="Enter"||e.key===" ")&&!p.soon){setSelProfile(p.key);setErr("");if(isMobile)setMobileFormOpen(true);}}}
        style={{
          background:isSelected?clr.forestLight:clr.white,
          border:`2px solid ${isSelected?clr.forest:clr.border}`,
          borderRadius:16, padding:"16px 14px",
          cursor:p.soon?"not-allowed":"pointer",
          position:"relative", transition:"all .2s ease",
          opacity:p.soon?.5:1, userSelect:"none",
          boxShadow:isSelected?"0 4px 16px rgba(11,77,44,.15)":"none",
        }}>
        {/* Badge sélectionné */}
        {isSelected && !p.soon && (
          <div style={{position:"absolute",top:10,right:10,width:22,height:22,borderRadius:"50%",background:clr.forest,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"#fff",fontWeight:800}}>✓</div>
        )}
        {/* Bientôt */}
        {p.soon && (
          <span style={{position:"absolute",top:8,right:8,fontSize:9,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",background:"#f1f5f9",color:clr.slate,padding:"2px 7px",borderRadius:8,fontFamily:"'Plus Jakarta Sans',sans-serif"}}>Bientôt</span>
        )}
        <div style={{width:44,height:44,borderRadius:12,background:isSelected?clr.forest:"#f1f5f9",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,marginBottom:10,transition:"all .2s"}}>
          {p.emoji}
        </div>
        <div style={{fontSize:14,fontWeight:700,color:isSelected?clr.forestDark:clr.navy,marginBottom:2,fontFamily:"'Plus Jakarta Sans',sans-serif"}}>{p.label}</div>
        <div style={{fontSize:11,fontWeight:600,color:clr.forest,marginBottom:4,fontFamily:"'Plus Jakarta Sans',sans-serif"}}>{p.sub}</div>
        <div style={{fontSize:11,color:clr.slate,lineHeight:1.45,fontFamily:"'Plus Jakarta Sans',sans-serif"}}>{p.desc}</div>
      </div>
    );
  };

  /* ── Layout desktop / mobile ───────────────────────── */
  const isDesktop = !isMobile;

  return (
    <div style={{
      minHeight:"100vh",
      background:`linear-gradient(180deg,${clr.forestLight} 0%, #fff 40%)`,
      padding:isDesktop?"48px 24px":"20px 16px 40px",
      fontFamily:"'Plus Jakarta Sans',sans-serif",
    }}>
      <style>{`
        @keyframes fadeUp{from{opacity:0;transform:translateY(20px);}to{opacity:1;transform:none;}}
        @keyframes scaleIn{from{opacity:0;transform:scale(.95);}to{opacity:1;transform:scale(1);}}
        .lp-btn{transition:all .2s ease;}
        .lp-btn:hover{filter:brightness(1.06);transform:translateY(-1px);}
        .lp-card-hover{transition:all .2s ease!important;}
        .lp-card-hover:hover{border-color:${clr.forest}!important;box-shadow:0 6px 20px rgba(11,77,44,.14)!important;transform:translateY(-2px);}
        .lp-input:focus{outline:none!important;border-color:${clr.forest}!important;box-shadow:0 0 0 3px rgba(11,77,44,.12)!important;}
        select.lp-input{background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2364748b' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center;padding-right:36px;}
        @media(max-width:768px){.portal-grid{grid-template-columns:1fr!important;}}
      `}</style>

      {/* En-tête portail */}
      <div style={{textAlign:"center",marginBottom:isDesktop?36:24,animation:"fadeUp .5s ease"}}>
        <button onClick={()=>setPortalStep(0)} style={{background:"none",border:"none",cursor:"pointer",marginBottom:12,display:"inline-flex",alignItems:"center",gap:6,color:clr.slate,fontSize:13,fontWeight:600,padding:"6px 12px",borderRadius:8}}>
          ← Accueil
        </button>
        <h2 style={{fontSize:isDesktop?"clamp(24px,3vw,32px)":"clamp(20px,6vw,26px)",fontWeight:800,color:clr.navy,margin:"0 0 8px",fontFamily:"'Playfair Display',serif"}}>
          Bienvenue sur votre espace numérique
        </h2>
        <p style={{fontSize:14,color:clr.slate,margin:0}}>
          Choisissez votre profil pour accéder à votre environnement
        </p>
      </div>

      {/* Layout principal */}
      <div style={{
        display:"grid",
        gridTemplateColumns:isDesktop?"1fr 380px":"1fr",
        gap:isDesktop?32:20,
        maxWidth:1060, margin:"0 auto",
        animation:"fadeUp .6s ease .1s both",
      }}>
        {/* Grille profils */}
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
            {PROFILES.map(p=><ProfileCard key={p.key} p={p}/>)}
          </div>
          {/* Mobile : formulaire en accordéon sous les cartes */}
          {isMobile&&mobileFormOpen&&(
            <div style={{marginTop:20,animation:"scaleIn .25s ease"}}>
              {formPanelJSX}
            </div>
          )}
        </div>

        {/* Formulaire (desktop) */}
        {isDesktop&&(
          <div style={{animation:"scaleIn .3s ease"}}>
            {formPanelJSX}
          </div>
        )}
      </div>
    </div>
  );
}


// ─── App Root ─────────────────────────────────────────────────────

// ── Error Boundary global ──────────────────────────────────────────
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }
  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }
  componentDidCatch(error, info) {
    console.error("SVTEEHB Error:", error, info);
  }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{minHeight:"100vh",display:"flex",flexDirection:"column",
          alignItems:"center",justifyContent:"center",background:"#f8fafc",
          fontFamily:"system-ui",padding:20,textAlign:"center"}}>
          <div style={{fontSize:48,marginBottom:16}}>⚠️</div>
          <h2 style={{fontSize:18,fontWeight:800,color:"#1e293b",marginBottom:8}}>
            Une erreur est survenue
          </h2>
          <p style={{fontSize:13,color:"#64748b",marginBottom:24,maxWidth:360}}>
            {String(this.state.error?.message||"Erreur inconnue")}
          </p>
          <button onClick={()=>window.location.reload()}
            style={{padding:"10px 24px",background:"#16a34a",color:"#fff",
              border:"none",borderRadius:10,fontSize:14,fontWeight:700,cursor:"pointer"}}>
            🔄 Recharger l'application
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}


// ─── Filtrage département (scope Proviseur) ────────────────────────
function filterDataByDept(data, deptId) {
  if (!data || !deptId) return data;
  const deptOf = {};
  Object.values(data.users||{}).forEach(u=>{ deptOf[u.id]=u.departement_id; });

  const users = {};
  Object.entries(data.users||{}).forEach(([id,u])=>{ if(u.departement_id===deptId) users[id]=u; });

  const prog = {};
  Object.entries(data.prog||{}).forEach(([k,v])=>{ if(deptOf[k.split("||")[0]]===deptId) prog[k]=v; });

  const absences = {};
  Object.entries(data.absences||{}).forEach(([k,v])=>{ if(deptOf[k.split("||")[0]]===deptId) absences[k]=v; });

  const epreuves = (data.epreuves||[]).filter(e=>deptOf[e.ens_id]===deptId);

  const exceptions = {};
  Object.entries(data.exceptions||{}).forEach(([ensId,v])=>{ if(deptOf[ensId]===deptId) exceptions[ensId]=v; });

  const edtBase = {};
  Object.entries(data.edtBase||{}).forEach(([ensId,v])=>{ if(deptOf[ensId]===deptId) edtBase[ensId]=v; });

  const classesF = (data.classes||[]).filter(c=>c.departement_id===deptId);

  return { ...data, users, prog, absences, epreuves, exceptions, edtBase, classes: classesF, deptFilterActive: true };
}

export default function App() {
  const [screen,setScreen] = useState("splash");
  const [user,setUser]     = useState(null);
  const [page,setPage]     = useState(()=>{
    try { return localStorage.getItem("svt_last_page")||"dashboard"; } catch { return "dashboard"; }
  });
  const [pendingFicheEns, setPendingFicheEns] = useState(null);
  const [pendingClasseSelect, setPendingClasseSelect] = useState(null);
  const [mobileSearchOpen, setMobileSearchOpen] = useState(false);
  const [data,setData]     = useState(null);
  const [viewDeptId,setViewDeptId] = useState(null); // filtre département actif (Proviseur uniquement)
  const [lang,setLang] = useState(()=>{ try { return localStorage.getItem("svt_lang")||"fr"; } catch { return "fr"; } });
  useEffect(()=>{ try { localStorage.setItem("svt_lang", lang); } catch {} },[lang]);
  const t = (fr) => (lang === "en" ? (TRANSLATIONS_EN[fr] || fr) : fr);
  const [online,setOnline] = useState(navigator.onLine);
  const [staticLoaded,setStaticLoaded] = useState(false);
  const [syncing,setSyncing] = useState(false);
  const [toast,setToast]   = useState(null);
  const [realtimeStatus,setRealtimeStatus] = useState("disconnected"); // disconnected | connecting | connected | error

  useEffect(()=>{ try{localStorage.setItem("svt_last_page",page);}catch{} },[page]);
  useEffect(()=>{
    const on=()=>{setOnline(true); showToast("🟢 Connexion rétablie");};
    const off=()=>{setOnline(false); showToast("📡 Mode hors ligne — données locales utilisées", false);};
    window.addEventListener("online",on);window.addEventListener("offline",off);
    return()=>{window.removeEventListener("online",on);window.removeEventListener("offline",off);};
  },[]);

  // ── Synchronisation multi-sessions (Supabase Realtime) ────────────
  // Écoute les changements sur les tables clés et déclenche un refreshData silencieux
  // (anti-rafale : attend 900ms sans nouvel événement avant de rafraîchir, pour éviter
  // de spammer Supabase si plusieurs changements arrivent d'un coup).
  useEffect(()=>{
    if (!user) { setRealtimeStatus("disconnected"); return; }
    let client, channel, debounceTimer;
    try {
      setRealtimeStatus("connecting");
      client = new RealtimeClient(REALTIME_URL, { params: { apikey: SB_KEY } });
      channel = client.channel("svteehb-sync");
      const onChange = () => {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(()=>{ refreshData(true); }, 900);
      };
      REALTIME_TABLES.forEach(table => {
        channel.on("postgres_changes", { event:"*", schema:"public", table }, onChange);
      });
      channel.subscribe((status)=>{
        if (status==="SUBSCRIBED") setRealtimeStatus("connected");
        else if (status==="CHANNEL_ERROR" || status==="TIMED_OUT") setRealtimeStatus("error");
        else if (status==="CLOSED") setRealtimeStatus("disconnected");
      });
    } catch { setRealtimeStatus("error"); }
    return ()=>{
      clearTimeout(debounceTimer);
      try { channel?.unsubscribe(); client?.removeChannel?.(channel); client?.disconnect?.(); } catch {}
      setRealtimeStatus("disconnected");
    };
  },[user?.id]);

  const showToast = useCallback((msg,ok=true)=>{
    setToast({msg,ok});setTimeout(()=>setToast(null),3000);
  },[]);

  const deptIdRef = useRef(null);
  const refreshData = useCallback(async(silent=false)=>{
    setSyncing(true);
    try {
      const d = await loadAllData(deptIdRef.current);
      if (d) {
        setData(prev=>({...prev, ...d}));
        if (!silent) showToast("✓ Données actualisées");
      } else {
        if (!silent) showToast("⚠ Actualisation impossible — vérifiez la connexion", false);
      }
    } catch {
      if (!silent) showToast("⚠ Actualisation impossible", false);
    } finally {
      setSyncing(false);
    }
  },[showToast]);

  const handleLogin = useCallback(async(acc)=>{
    setSyncing(true);setScreen("loading");
    window.__svtSessionToken = acc.token || null;
    await loadStaticData(); // Charger les données statiques en parallèle
    deptIdRef.current = isAdminRole(acc.role) && acc.role !== "proviseur" ? acc.departement_id : null;
    const d = await loadAllData(deptIdRef.current);
    const safeD = d || { users:{}, prog:{}, epreuves:[], classes:[], exceptions:{} };
    const sbUser = safeD.users?.[acc.id];
    // Fallback: si Supabase ne retourne pas les classes, utiliser les données EDT
    const classes = (sbUser?.classes||[]).length > 0
      ? sbUser.classes
      : (acc.classes||ENS_CLASSES_REF[acc.id]||[]);
    setData(safeD);
    setUser({...acc, classes, photo: sbUser?.photo||null});
    setSyncing(false); setScreen("app");
    const savedPage = !acc.mustChangePwd ? (localStorage.getItem("svt_last_page")||"dashboard") : "settings";
    setPage(savedPage);
  },[]);

  const handleLogout = ()=>{setUser(null);setData(null);setPage("dashboard");setScreen("login");};

  // Bouton retour Android — intercepter pour naviguer dans l'app
  useEffect(()=>{
    const handlePop = ()=>{
      if(mobileSearchOpen){ setMobileSearchOpen(false); return; }
      if(screen==="app" && page && page!=="dashboard"){
        setPage("dashboard");
        window.history.pushState({page:"dashboard"},"","");
      }
    };
    window.addEventListener("popstate",handlePop);
    return()=>window.removeEventListener("popstate",handlePop);
  },[screen,page,mobileSearchOpen]);

  useEffect(()=>{ if(screen==="app"&&page) window.history.pushState({page},"",""); },[page,screen]);

  const scopedData = useMemo(()=>filterDataByDept(data, user?.role==="proviseur"?viewDeptId:null), [data, viewDeptId, user?.role]);

  const ctx = {user,setUser,page,setPage,data:scopedData,rawData:data,setData,viewDeptId,setViewDeptId,online,syncing,toast,showToast,staticLoaded,setStaticLoaded,refreshData,pendingFicheEns,setPendingFicheEns,pendingClasseSelect,setPendingClasseSelect,mobileSearchOpen,setMobileSearchOpen,realtimeStatus,lang,setLang,t};

  return(
    <AppCtx.Provider value={ctx}>
      <style>{GLOBAL_CSS}</style>
      {screen==="splash"  && <Splash onDone={()=>setScreen("login")}/>}
      {screen==="login"   && <LoginPage onLogin={handleLogin}/>}
      {screen==="loading" && (
        <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:C.bg,flexDirection:"column",gap:16}}>
          <Spinner size={36} color={C.green}/><p style={{color:C.txtMuted,fontSize:13}}>Chargement des données Supabase…</p>
        </div>
      )}
      {screen==="app" && <AppLayout onLogout={handleLogout}/>}
    </AppCtx.Provider>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <ErrorBoundary>
    <App />
  </ErrorBoundary>
);




