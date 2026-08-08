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
        "submit_absence","submit_note","submit_prog","submit_epreuve","submit_eleves_import",
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
    --c-sidebar:#0f1f14;--c-sidebarBorder:rgba(255,255,255,0.07);--c-sidebarActive:rgba(34,197,94,0.15);--c-sidebarActiveText:#4ade80;--c-sidebarText:rgba(255,255,255,0.55);--c-sidebarHover:rgba(255,255,255,0.05);
    --c-green:#16a34a;--c-greenLight:#22c55e;--c-greenDark:#0c3d24;--c-greenPale:#f0fdf4;--c-greenBorder:#bbf7d0;
    --c-gold:#c8a951;--c-goldPale:#fdf6e3;--c-goldBorder:rgba(200,169,81,.35);
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
        {[{id:"liste",label:"👥 Liste & présences"},{id:"notes",label:"📝 Notes"}].map(t=>(
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
                  <span style={{flex:1, fontSize: isMobile?12.5:13, fontWeight:600, color:C.txt, minWidth:0, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>{e.nom}</span>
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

  return (
    <div style={{display:"flex", flexDirection:"column", gap:12}}>
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
                  <div style={{display:"flex", alignItems:"center", gap:8}}>
                    <div style={{flex:1, maxWidth:200}}><ProgBar value={ens.taux}/></div>
                    <span style={{fontSize:12, fontWeight:800, color:taux2col(ens.taux), minWidth:36}}>{ens.taux}%</span>
                    <span style={{fontSize:11, color:C.txtMuted}}>{ens.totalFait}/{ens.totalRef} leçons</span>
                    <span style={{fontSize:11, color:C.txtMuted}}>· {(ens.classes||[]).length} classe{(ens.classes||[]).length>1?"s":""}</span>
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
    setPreviewHtml(html);
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

function DashboardSurveillance() {
  const {rawData:data,user} = useApp();
  const {isMobile} = useDevice();
  const [loading,setLoading]   = useState(true);
  const [tab,setTab]           = useState("absences");
  const [vieSco,setVieSco]     = useState([]);
  const [vieLoading,setVieLoading] = useState(true);
  const [stats,setStats]       = useState({total:0,parEleve:[],parDept:[],nbAlerte:0});
  const [showForm,setShowForm] = useState(false);
  const [saving,setSaving]     = useState(false);
  const [form,setForm]         = useState({eleve_id:"",classe:"",motif:"",details:"",gravite:"faible"});
  const [selClasse,setSelClasse] = useState("");
  const [elevesSel,setElevesSel] = useState([]);
  const [formErr,setFormErr]   = useState("");

  useEffect(()=>{
    if(!data) return;
    const deptOf={};
    Object.values(data.users||{}).forEach(u=>{deptOf[u.id]=u.departement_id||1;});
    const parEleveMap={};let total=0;
    Object.entries(data.absences||{}).forEach(([k,absents])=>{
      const [,classe,seance]=k.split("||");
      (absents||[]).forEach(eleveId=>{
        total++;
        if(!parEleveMap[eleveId]) parEleveMap[eleveId]={id:eleveId,classe,count:0,dernier:seance};
        parEleveMap[eleveId].count++;
        if(seance>parEleveMap[eleveId].dernier)parEleveMap[eleveId].dernier=seance;
      });
    });
    const parEleve=Object.values(parEleveMap)
      .map(e=>({...e,nom:(ELEVES_DB[e.classe]||[]).find(x=>x.id===e.id)?.nom||e.id}))
      .sort((a,b)=>b.count-a.count).slice(0,25);
    const nbAlerte=Object.values(parEleveMap).filter(e=>e.count>=3).length;
    const absParDept={};
    Object.entries(data.absences||{}).forEach(([k,absents])=>{
      const dId=deptOf[k.split("||")[0]]||1;
      absParDept[dId]=(absParDept[dId]||0)+(absents?absents.length:0);
    });
    const parDept=DEPARTEMENTS_LIST.map(d=>({...d,total:absParDept[d.id]||0})).sort((a,b)=>b.total-a.total);
    setStats({total,parEleve,parDept,nbAlerte});
    setLoading(false);
  },[data]);

  const loadVieSco = async()=>{
    setVieLoading(true);
    const rows = await sb.get("vie_scolaire","?select=*&order=date.desc,created_at.desc&limit=300");
    setVieSco(rows||[]);
    setVieLoading(false);
  };
  useEffect(()=>{ loadVieSco(); },[]);

  useEffect(()=>{
    if(!selClasse){setElevesSel([]);setForm(f=>({...f,eleve_id:"",classe:""}));return;}
    setElevesSel(ELEVES_DB[selClasse]||[]);
    setForm(f=>({...f,classe:selClasse,eleve_id:""}));
  },[selClasse]);

  const typeMap = {retards:"retard",sanctions:"sanction",incidents:"incident"};

  const openForm = ()=>{
    setFormErr("");
    setForm({eleve_id:"",classe:"",motif:"",details:"",gravite:"faible"});
    setSelClasse("");
    setShowForm(!showForm);
  };

  const saveEntry = async()=>{
    if(!form.eleve_id||!form.classe){setFormErr("Sélectionnez une classe et un élève.");return;}
    setFormErr("");setSaving(true);
    const payload={
      type:typeMap[tab]||"retard",
      eleve_id:form.eleve_id, classe:form.classe,
      motif:form.motif||null, details:form.details||null,
      gravite:(tab==="retards")?"faible":form.gravite||"faible",
      enregistre_par:user?.id||"surveillance",
    };
    const ok = await sb.upsert("vie_scolaire",payload);
    if(ok){ await loadVieSco(); setShowForm(false); }
    else { setFormErr("Erreur lors de l'enregistrement. Vérifiez la connexion."); }
    setSaving(false);
  };

  const filteredVie = vieSco.filter(v=>v.type===(typeMap[tab]||"retard"));

  const TABS=[
    {id:"absences",  label:"Absences",  emoji:"📋"},
    {id:"retards",   label:"Retards",   emoji:"⏰"},
    {id:"sanctions", label:"Sanctions", emoji:"⚠️"},
    {id:"incidents", label:"Incidents", emoji:"🚨"},
  ];

  const GraviteBadge=({g})=>{
    const cfgG={faible:{bg:"#fefce8",fg:"#854d0e"},moyen:{bg:"#fff7ed",fg:"#c2410c"},grave:{bg:"#fef2f2",fg:"#b91c1c"}};
    const gc=cfgG[g]||cfgG.faible;
    return React.createElement("span",{style:{fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:10,background:gc.bg,color:gc.fg}},g||"faible");
  };

  const NomEleve=({eleveId,classe})=>{
    const nom=(ELEVES_DB[classe]||[]).find(x=>x.id===eleveId)?.nom||eleveId;
    return React.createElement("span",null,nom);
  };

  return(
    <div style={{padding:"20px 20px 40px",display:"flex",flexDirection:"column",gap:18}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:10}}>
        <div>
          <h2 style={{fontSize:20,fontWeight:800,color:C.txt,margin:0}}>Surveillance générale 🛡️</h2>
          <p style={{color:C.txtMuted,margin:"3px 0 0",fontSize:12}}>{new Date().toLocaleDateString("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric"})} · Vue école entière</p>
        </div>
        {tab!=="absences" && (
          <button onClick={openForm}
            style={{padding:"9px 18px",borderRadius:10,border:"none",background:showForm?C.border:C.green,color:showForm?C.txt:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>
            {showForm?"✕ Annuler":"➕ Enregistrer"}
          </button>
        )}
      </div>

      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <KpiCard label="Absences" value={stats.total} sub="Toutes classes" iconEmoji="📋" bg={C.bluePale} subColor={C.blue} loading={loading} delay={0}/>
        <KpiCard label="Élèves en alerte" value={stats.nbAlerte} sub="≥ 3 absences" iconEmoji="⚠️" bg={C.redPale} subColor={C.red} loading={loading} delay={0.05}/>
        <KpiCard label="Retards" value={vieSco.filter(v=>v.type==="retard").length} sub="Enregistrés" iconEmoji="⏰" bg={C.amberPale} subColor={C.amber} loading={vieLoading} delay={0.1}/>
        <KpiCard label="Sanctions" value={vieSco.filter(v=>v.type==="sanction").length} sub="Enregistrées" iconEmoji="⚠️" bg={C.redPale} subColor={C.red} loading={vieLoading} delay={0.15}/>
      </div>

      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>{setTab(t.id);setShowForm(false);}}
            style={{padding:"8px 16px",borderRadius:20,border:"1.5px solid "+(tab===t.id?C.green:C.border),
              background:tab===t.id?C.greenPale:C.white,color:tab===t.id?C.green:C.txtMuted,
              fontWeight:700,fontSize:12.5,cursor:"pointer",fontFamily:"inherit",transition:"all .15s"}}>
            {t.emoji} {t.label}
          </button>
        ))}
      </div>

      {showForm && tab!=="absences" && (
        <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:18,display:"flex",flexDirection:"column",gap:12}}>
          <h3 style={{margin:0,fontSize:13,fontWeight:700,color:C.txt}}>
            {tab==="retards"?"⏰ Enregistrer un retard":tab==="sanctions"?"⚠️ Enregistrer une sanction":"🚨 Déclarer un incident"}
          </h3>
          {formErr && <div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:8,padding:"8px 12px",fontSize:12.5,color:"#b91c1c"}}>{formErr}</div>}
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:10}}>
            <div>
              <label style={{fontSize:11,fontWeight:600,color:C.txtMuted,display:"block",marginBottom:4}}>Classe *</label>
              <select value={selClasse} onChange={e=>setSelClasse(e.target.value)}
                style={{width:"100%",padding:"9px 12px",border:"1.5px solid "+C.border,borderRadius:8,fontSize:13,fontFamily:"inherit",background:"#f8fafc"}}>
                <option value="">— Sélectionner —</option>
                {CLASSES_REELLES.map(c=><option key={c.code} value={c.code}>{c.code}</option>)}
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
                placeholder="Motif principal..."
                style={{width:"100%",padding:"9px 12px",border:"1.5px solid "+C.border,borderRadius:8,fontSize:13,fontFamily:"inherit",background:"#f8fafc",boxSizing:"border-box"}}/>
            </div>
            {tab!=="retards" && (
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
                placeholder="Détails complémentaires..."
                style={{width:"100%",padding:"9px 12px",border:"1.5px solid "+C.border,borderRadius:8,fontSize:13,fontFamily:"inherit",background:"#f8fafc",resize:"vertical",minHeight:64,boxSizing:"border-box"}}/>
            </div>
          </div>
          <button onClick={saveEntry} disabled={saving}
            style={{alignSelf:"flex-end",padding:"10px 24px",borderRadius:10,border:"none",background:saving?"#94a3b8":C.green,color:"#fff",fontWeight:700,fontSize:13,cursor:saving?"not-allowed":"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:8}}>
            {saving?<><Spinner size={12} color="#fff"/> Enregistrement…</>:"✓ Enregistrer"}
          </button>
        </div>
      )}

      {tab==="absences" ? (
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1.4fr 1fr",gap:14}}>
          <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:16}}>
            <h3 style={{margin:"0 0 4px",fontSize:12.5,fontWeight:700,color:C.txt}}>🔍 Élèves les plus absents</h3>
            <p style={{margin:"0 0 12px",fontSize:10,color:C.txtMuted}}>Cumul sur toute la période — top 25</p>
            {loading?<Sk h={200} br={8}/>:stats.parEleve.length>0?(
              <div style={{display:"flex",flexDirection:"column",gap:6,maxHeight:420,overflowY:"auto"}}>
                {stats.parEleve.map((e,i)=>(
                  <div key={e.id+e.classe} style={{display:"flex",alignItems:"center",gap:10,padding:"7px 4px",borderBottom:i<stats.parEleve.length-1?"1px solid "+C.border:"none"}}>
                    <span style={{fontSize:11,color:C.txtMuted,width:20}}>{i+1}</span>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:12.5,fontWeight:700,color:C.txt,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{e.nom}</div>
                      <div style={{fontSize:10,color:C.txtMuted}}>{e.classe} · dernière le {e.dernier}</div>
                    </div>
                    <span style={{fontSize:12,fontWeight:800,color:e.count>=3?C.red:C.amber,flexShrink:0}}>{e.count}</span>
                  </div>
                ))}
              </div>
            ):<div style={{fontSize:11,color:C.txtLight,textAlign:"center",padding:"30px 0"}}>Aucune absence enregistrée</div>}
          </div>
          <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,padding:16}}>
            <h3 style={{margin:"0 0 12px",fontSize:12.5,fontWeight:700,color:C.txt}}>🏛️ Par département</h3>
            {loading?<Sk h={150} br={8}/>:stats.parDept.some(d=>d.total>0)?(
              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                {stats.parDept.filter(d=>d.total>0).map(d=>(
                  <div key={d.id} style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:14,flexShrink:0}}>{d.emoji}</span>
                    <span style={{fontSize:11,color:C.txt,flex:1}}>{d.nom}</span>
                    <span style={{fontSize:11,fontWeight:800,color:C.red}}>{d.total}</span>
                  </div>
                ))}
              </div>
            ):<div style={{fontSize:11,color:C.txtLight,textAlign:"center",padding:"30px 0"}}>Aucune absence enregistrée</div>}
          </div>
        </div>
      ) : (
        <div style={{background:C.white,borderRadius:12,border:"1px solid "+C.border,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12.5,minWidth:480}}>
              <thead>
                <tr style={{background:"#f8fafc",borderBottom:"1px solid "+C.border}}>
                  <th style={{padding:"10px 12px",textAlign:"left",fontSize:10,fontWeight:700,color:C.txtMuted}}>Date</th>
                  <th style={{padding:"10px 12px",textAlign:"left",fontSize:10,fontWeight:700,color:C.txtMuted}}>Élève</th>
                  <th style={{padding:"10px 12px",textAlign:"left",fontSize:10,fontWeight:700,color:C.txtMuted}}>Classe</th>
                  <th style={{padding:"10px 12px",textAlign:"left",fontSize:10,fontWeight:700,color:C.txtMuted}}>Motif</th>
                  {tab!=="retards" && <th style={{padding:"10px 12px",textAlign:"center",fontSize:10,fontWeight:700,color:C.txtMuted}}>Gravité</th>}
                </tr>
              </thead>
              <tbody>
                {vieLoading?(
                  <tr><td colSpan={5} style={{padding:24,textAlign:"center",color:C.txtLight}}>Chargement…</td></tr>
                ):filteredVie.length===0?(
                  <tr><td colSpan={5} style={{padding:32,textAlign:"center",color:C.txtLight}}>
                    <div style={{fontSize:24,marginBottom:6}}>📭</div>
                    Aucun enregistrement
                  </td></tr>
                ):filteredVie.map((v,i)=>(
                  <tr key={v.id} style={{borderBottom:"1px solid "+C.border,background:i%2===0?C.white:"#fafafa"}}>
                    <td style={{padding:"10px 12px",color:C.txtMuted,whiteSpace:"nowrap"}}>
                      {new Date(v.date).toLocaleDateString("fr-FR",{day:"2-digit",month:"short"})}
                    </td>
                    <td style={{padding:"10px 12px",fontWeight:600,color:C.txt}}><NomEleve eleveId={v.eleve_id} classe={v.classe}/></td>
                    <td style={{padding:"10px 12px",color:C.txtMuted}}>{v.classe}</td>
                    <td style={{padding:"10px 12px",color:C.txt,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{v.motif||"—"}</td>
                    {tab!=="retards" && <td style={{padding:"10px 12px",textAlign:"center"}}><GraviteBadge g={v.gravite}/></td>}
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
  const {data, showToast} = useApp();
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

      <div style={{display:"grid", gridTemplateColumns: isMobile?"1fr":"repeat(auto-fill, minmax(280px, 1fr))", gap:12}}>
        {DEPARTEMENTS_LIST.map(d => {
          const deptMatieres = (matieres||[]).filter(m=>m.departement_id===d.id);
          const isOpen = openDept===d.id;
          return (
            <div key={d.id} style={{background:C.white, borderRadius:12, border:`1px solid ${C.border}`, padding:16}}>
              <div onClick={()=>setOpenDept(isOpen?null:d.id)} style={{display:"flex", alignItems:"center", gap:10, cursor:"pointer"}}>
                <span style={{fontSize:20}}>{d.emoji}</span>
                <div style={{flex:1}}>
                  <div style={{fontSize:13, fontWeight:700, color:C.txt}}>{d.nom}</div>
                  <div style={{fontSize:10.5, color:C.txtMuted}}>{nbEnsParDept[d.id]||0} enseignant{(nbEnsParDept[d.id]||0)>1?"s":""} · {deptMatieres.length} matière{deptMatieres.length>1?"s":""}</div>
                </div>
                <span style={{fontSize:12, color:C.txtMuted}}>{isOpen?"▲":"▼"}</span>
              </div>

              {isOpen && (
                <div style={{marginTop:12, paddingTop:12, borderTop:`1px solid ${C.border}`, display:"flex", flexDirection:"column", gap:6}}>
                  {loading ? <Sk h={16} w="60%"/> : deptMatieres.length===0 ? (
                    <div style={{fontSize:11, color:C.txtLight, fontStyle:"italic"}}>Aucune matière</div>
                  ) : deptMatieres.map(m => (
                    <div key={m.id} style={{display:"flex", alignItems:"center", gap:8, padding:"6px 8px", background:"#f8fafc", borderRadius:7}}>
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

                  <div style={{display:"flex", gap:6, marginTop:4}}>
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
    if(page==="dashboard")         return <W>{user?.role==="proviseur"?<DashboardProviseur/>:user?.role==="surveillant_general"?<DashboardSurveillance/>:(isAdmin||user?.role==="censeur")?<DashboardAdmin/>:<DashboardTeacher/>}</W>
    if(page==="programme")         return <W>{(isAdmin||user?.role==="censeur")?<SuiviProgrammePage/>:<MonProgrammePage/>}</W>
    if(page==="epreuves")          return <W><EpreuvesPage/></W>
    if(page==="edt-teacher")       return <W><MonEdtPage/></W>
    if(page==="edt")               return <W>{(isAdmin||user?.role==="censeur")?<EdtPage/>:<MonEdtPage/>}</W>
    if(page==="enseignants")       return <W>{isAdmin?<EnseignantsPage/>:null}</W>
    if(page==="gestion-annuelle")  return <W>{isAdmin?<GestionAnnuellePage/>:null}</W>
    if(page==="departements")      return <W>{user?.role==="proviseur"?<DepartementsPage/>:null}</W>
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

  const PROFILES = [
    {key:"direction",  label:"Direction",             sub:"Proviseur",             desc:"Gestion globale de l'établissement", emoji:"👨🏾‍💼", role:"proviseur",          needsDept:false},
    {key:"censeur",    label:"Censeur",               sub:"Organisation pédagogique", desc:"Classes, notes, suivi discipline",  emoji:"📚",  role:"censeur",            needsDept:false},
    {key:"sg",         label:"Surveillance Générale", sub:"Vie scolaire & discipline", desc:"Absences, retards, incidents",    emoji:"🛡️", role:"surveillant_general", needsDept:false},
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
      minHeight:"100vh", position:"relative", overflow:"hidden",
      display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
      background:`url(${LOGIN_BG_B64}) center/cover no-repeat`,
    }}>
      <style>{`
        @keyframes fadeUp{from{opacity:0;transform:translateY(24px);}to{opacity:1;transform:none;}}
        @keyframes scaleIn{from{opacity:0;transform:scale(.92);}to{opacity:1;transform:scale(1);}}
        .lp-btn:hover{filter:brightness(1.06);transform:translateY(-1px);}
        .lp-card-hover:hover{border-color:${clr.forest}!important;box-shadow:0 8px 24px rgba(11,77,44,.15)!important;transform:translateY(-2px);}
        .lp-card-hover:focus{outline:2px solid ${clr.forest};outline-offset:3px;}
        .lp-input:focus{outline:none;border-color:${clr.forest}!important;box-shadow:0 0 0 3px rgba(11,77,44,.12);}
      `}</style>

      {/* Overlay vert foncé */}
      <div style={{position:"absolute",inset:0,background:"linear-gradient(160deg,rgba(8,61,34,.88) 0%,rgba(11,77,44,.82) 50%,rgba(15,23,42,.88) 100%)"}}/>

      {/* Motif points or (subtil) */}
      <svg style={{position:"absolute",inset:0,width:"100%",height:"100%",opacity:.08,pointerEvents:"none"}} xmlns="http://www.w3.org/2000/svg">
        <defs><pattern id="dots" x="0" y="0" width="30" height="30" patternUnits="userSpaceOnUse">
          <circle cx="1.5" cy="1.5" r="1.5" fill="#D4AF37"/>
        </pattern></defs>
        <rect width="100%" height="100%" fill="url(#dots)"/>
      </svg>

      {/* Contenu */}
      <div style={{position:"relative",zIndex:2,display:"flex",flexDirection:"column",alignItems:"center",padding:isMobile?"24px 20px":"40px 32px",animation:"fadeUp .7s ease",maxWidth:520,width:"100%"}}>
        {/* Logo */}
        <div style={{width:isMobile?88:108,height:isMobile?88:108,borderRadius:"50%",overflow:"hidden",border:`4px solid ${clr.gold}`,boxShadow:"0 8px 32px rgba(0,0,0,.4)",marginBottom:24,background:"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
          <img src={LOGO_LYCEE_B64} alt="Logo" style={{width:"90%",height:"90%",objectFit:"contain"}}/>
        </div>

        {/* Drapeau + établissement */}
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
          <svg viewBox="0 0 45 30" width="36" height="24" style={{borderRadius:3,boxShadow:"0 2px 6px rgba(0,0,0,.3)",flexShrink:0}}>
            <rect width="15" height="30" fill="#007A5E"/>
            <rect x="15" width="15" height="30" fill="#CE1126"/>
            <rect x="30" width="15" height="30" fill="#FCD116"/>
            <polygon points="22.5,8 23.9,12.6 28.7,12.6 24.9,15.4 26.3,20 22.5,17.2 18.7,20 20.1,15.4 16.3,12.6 21.1,12.6" fill="#FCD116"/>
          </svg>
          <span style={{fontSize:11,fontWeight:700,letterSpacing:".1em",textTransform:"uppercase",color:clr.gold,fontFamily:"'Plus Jakarta Sans',sans-serif"}}>République du Cameroun</span>
        </div>

        <h1 style={{fontSize:isMobile?"clamp(22px,6vw,28px)":"clamp(28px,4vw,36px)",fontWeight:800,color:"#fff",textAlign:"center",margin:"0 0 6px",fontFamily:"'Playfair Display',serif",lineHeight:1.2}}>
          Lycée de Kakatare – Maroua
        </h1>
        <p style={{fontSize:isMobile?13:14,color:"rgba(255,255,255,.75)",textAlign:"center",margin:"0 0 4px",fontFamily:"'Plus Jakarta Sans',sans-serif",fontWeight:500}}>
          Plateforme numérique de gestion scolaire
        </p>
        <div style={{display:"inline-flex",alignItems:"center",gap:6,background:"rgba(212,175,55,.18)",border:`1px solid ${clr.gold}60`,borderRadius:20,padding:"4px 14px",marginBottom:32}}>
          <span style={{fontSize:12,fontWeight:700,color:clr.gold,fontFamily:"'Plus Jakarta Sans',sans-serif",letterSpacing:".05em"}}>Année scolaire 2025–2026</span>
        </div>

        {/* Métriques (chiffres réels animés) */}
        <div style={{display:"flex",gap:32,marginBottom:36,justifyContent:"center"}}>
          {[
            {val:counter.eleves,label:"Élèves inscrits",color:"#4ade80"},
            {val:counter.ens,   label:"Enseignants",    color:clr.gold},
          ].map(({val,label,color})=>(
            <div key={label} style={{textAlign:"center"}}>
              <div style={{fontSize:isMobile?26:32,fontWeight:900,color,fontFamily:"'Playfair Display',serif",lineHeight:1}}>{val.toLocaleString("fr-FR")}</div>
              <div style={{fontSize:11,color:"rgba(255,255,255,.65)",marginTop:2,fontFamily:"'Plus Jakarta Sans',sans-serif"}}>{label}</div>
            </div>
          ))}
        </div>

        {/* CTA */}
        <button
          className="lp-btn"
          onClick={()=>setPortalStep(1)}
          style={{
            background:`linear-gradient(135deg,${clr.gold},#b8860b)`,
            color:clr.navy, fontWeight:800, fontSize:isMobile?14:15,
            padding:isMobile?"14px 28px":"16px 40px",
            borderRadius:14, border:"none", cursor:"pointer",
            boxShadow:"0 8px 24px rgba(212,175,55,.4)",
            fontFamily:"'Plus Jakarta Sans',sans-serif",
            transition:"all .2s ease", letterSpacing:".02em",
          }}>
          Accéder au Portail Numérique →
        </button>

        <p style={{marginTop:16,fontSize:11,color:"rgba(255,255,255,.4)",fontFamily:"'Plus Jakarta Sans',sans-serif",textAlign:"center",display:"flex",alignItems:"center",gap:6}}>
          <span>🔒</span> Connexion sécurisée SSL • SVTEEHB Kakatare
        </p>
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
          setSelProfile(p.key); setErr("");
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



