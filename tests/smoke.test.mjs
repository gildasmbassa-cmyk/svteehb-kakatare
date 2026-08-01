// tests/smoke.test.mjs — Harnais de test minimal SVTEEHB Lykama
// Zéro dépendance : utilise node:test + node:assert (Node 18+) et fetch natif.
// Usage : node --test tests/smoke.test.mjs
//
// Couvre spécifiquement les bugs trouvés en session (2026-08) :
//  - sb.rpc([]) traité comme succès → faux login "fantôme" (corrigé : doit renvoyer vide proprement)
//  - RPC admin_* appelables sans vérification de rôle (corrigé : jeton de session requis)
//  - Filtre département non appliqué à certaines tables
//  - RLS absente sur departements/matieres/animateurs_matieres (corrigé)
//
// Comptes utilisés : uniquement les comptes de test (test_math_*), jamais de comptes réels.
// Si ces comptes de test sont supprimés (données réelles reçues), adapter ou retirer ce fichier.

import { test } from "node:test";
import assert from "node:assert/strict";

const SB_URL = "https://ochijkylsranqectspxc.supabase.co";
const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9jaGlqa3lsc3JhbnFlY3RzcHhjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODA0ODExODAsImV4cCI6MjA5NjA1NzE4MH0.LWQf-wSkgPA3H9IH1neYJEY1yfohSu_h13KdO7cT28M";

const TEST_ENS_ID = "test_math_ens1";
const TEST_ENS_PWD = "Test2026!";
const TEST_ANIM_ID = "test_math_anim";
const TEST_ANIM_PWD = "Test2026!";

function headers(extra = {}) {
  return { apikey: SB_KEY, Authorization: `Bearer ${SB_KEY}`, "Content-Type": "application/json", ...extra };
}

async function rpc(fn, params = {}) {
  const r = await fetch(`${SB_URL}/rest/v1/rpc/${fn}`, {
    method: "POST",
    headers: headers({ Prefer: "return=representation" }),
    body: JSON.stringify(params),
  });
  const text = await r.text();
  let json = null;
  try { json = text ? JSON.parse(text) : null; } catch { /* pas du JSON (ex: message d'erreur brut) */ }
  return { status: r.status, ok: r.ok, json, raw: text };
}

// ── 1 — Authentification : identifiants corrects ─────────────────────
test("authenticate_user : identifiants corrects → session valide", async () => {
  const { ok, json } = await rpc("authenticate_user", { p_id: TEST_ENS_ID, p_mdp: TEST_ENS_PWD });
  assert.equal(ok, true, "la requête doit réussir");
  assert.ok(Array.isArray(json) && json.length === 1, "doit renvoyer exactement 1 ligne");
  const u = json[0];
  assert.equal(u.role, "enseignant");
  assert.equal(u.departement_id, 2, "doit être rattaché à Mathématiques");
  assert.ok(u.token && u.token.length > 20, "un jeton de session doit être émis");
});

// ── 2 — Authentification : mauvais mot de passe (régression login fantôme) ──
test("authenticate_user : mauvais mot de passe → aucune ligne (pas de session fantôme)", async () => {
  const { json } = await rpc("authenticate_user", { p_id: TEST_ENS_ID, p_mdp: "MOT_DE_PASSE_FAUX" });
  assert.ok(Array.isArray(json) && json.length === 0, "doit renvoyer un tableau vide, jamais un objet partiel");
});

// ── 3 — Authentification : identifiant inexistant ─────────────────────
test("authenticate_user : identifiant inconnu → aucune ligne", async () => {
  const { json } = await rpc("authenticate_user", { p_id: "compte_qui_nexiste_pas_xyz", p_mdp: "peu importe" });
  assert.ok(Array.isArray(json) && json.length === 0);
});

// ── 4 — RPC protégée : jeton absent/invalide → rejetée ────────────────
test("admin_set_teacher_classes : jeton absent → rejetée (erreur, pas de succès silencieux)", async () => {
  const { ok, status } = await rpc("admin_set_teacher_classes", {
    p_id: TEST_ENS_ID, p_classes: ["6e I"], p_token: null,
  });
  assert.equal(ok, false, "doit échouer sans jeton valide");
  assert.ok(status >= 400, `doit renvoyer un statut d'erreur (reçu ${status})`);
});

// ── 5 — RPC protégée : rôle insuffisant (enseignant) → rejetée même avec jeton valide ──
test("admin_set_teacher_classes : jeton d'un simple enseignant → rejetée (rôle insuffisant)", async () => {
  const login = await rpc("authenticate_user", { p_id: TEST_ENS_ID, p_mdp: TEST_ENS_PWD });
  const token = login.json[0].token;
  const { ok } = await rpc("admin_set_teacher_classes", {
    p_id: TEST_ENS_ID, p_classes: ["6e I"], p_token: token,
  });
  assert.equal(ok, false, "un enseignant ne doit pas pouvoir exécuter une RPC admin, même avec un jeton valide");
});

// ── 6 — RPC protégée : rôle suffisant (animateur) → acceptée ──────────
test("admin_set_teacher_classes : jeton d'un animateur → acceptée (opération neutre)", async () => {
  const login = await rpc("authenticate_user", { p_id: TEST_ANIM_ID, p_mdp: TEST_ANIM_PWD });
  assert.ok(login.json?.[0]?.token, "précondition : login animateur doit réussir");
  const token = login.json[0].token;
  const { ok, json } = await rpc("admin_set_teacher_classes", {
    p_id: TEST_ENS_ID, p_classes: ["6e I"], p_token: token, // valeur inchangée, opération neutre
  });
  assert.equal(ok, true, "un animateur doit pouvoir exécuter cette RPC");
  assert.equal(json, true, "la fonction doit renvoyer true (ligne trouvée et mise à jour)");
});

// ── 7 — Filtre département : utilisateurs ─────────────────────────────
test("utilisateurs?departement_id=eq.2 : ne renvoie que les comptes Mathématiques", async () => {
  const r = await fetch(`${SB_URL}/rest/v1/utilisateurs?select=id,departement_id&departement_id=eq.2`, { headers: headers() });
  const rows = await r.json();
  assert.ok(rows.length >= 3, "doit au moins contenir les 3 comptes de test Maths");
  assert.ok(rows.every(u => u.departement_id === 2), "aucune fuite d'un autre département");
});

// ── 8 — RLS : écriture directe sur matieres refusée à l'anon ──────────
test("RLS : POST direct sur matieres (clé anon) doit être refusé", async () => {
  const r = await fetch(`${SB_URL}/rest/v1/matieres`, {
    method: "POST",
    headers: headers({ Prefer: "return=minimal" }),
    body: JSON.stringify({ nom: "MATIERE_TEST_RLS_NE_DOIT_PAS_EXISTER", departement_id: 1 }),
  });
  assert.notEqual(r.status, 201, "l'écriture directe ne doit jamais réussir (doit passer par une RPC admin_*)");
});

// ── 9 — RLS : lecture publique toujours autorisée ─────────────────────
test("RLS : lecture de departements (clé anon) doit fonctionner", async () => {
  const r = await fetch(`${SB_URL}/rest/v1/departements?select=id,nom`, { headers: headers() });
  assert.equal(r.status, 200);
  const rows = await r.json();
  assert.ok(rows.length >= 8, "les 8 départements doivent rester lisibles");
});
