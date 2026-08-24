// src/lib/constants.js — Constantes indépendantes, extraites de App.jsx (première tranche du découpage)
// Aucune dépendance sur le reste de l'app — sûr à importer n'importe où.

export const TRANSLATIONS_EN = {
  "Tableau de bord": "Dashboard",
  "Enseignants": "Teachers",
  "Élèves": "Students",
  "Suivi programme": "Curriculum tracking",
  "Épreuves": "Exams",
  "Emploi du temps": "Timetable",
  "Documents": "Documents",
  "Gestion annuelle": "Annual management",
  "Départements": "Departments",
  "Mes classes": "My classes",
  "Cahier de texte": "Class logbook",
  "Mon programme": "My curriculum",
  "Mon emploi du temps": "My timetable",
  "Déconnexion": "Log out",
  "Paramètres": "Settings",
};

export const DEPARTEMENTS_LIST = [
    {id:1,nom:"SVT",emoji:"🌿"},
    {id:2,nom:"Mathématiques",emoji:"📐"},
    {id:3,nom:"Sciences Physiques / PCT",emoji:"🧪"},
    {id:4,nom:"Lettres Françaises",emoji:"📖"},
    {id:5,nom:"Histoire-Géographie",emoji:"🌍"},
    {id:6,nom:"Espagnol",emoji:"🇪🇸"},
    {id:7,nom:"EPS",emoji:"🏃"},
    {id:8,nom:"Informatique",emoji:"💻"},
    {id:9,nom:"ESF",emoji:"🏠"},
    {id:10,nom:"ECM",emoji:"⚖️"},
    {id:11,nom:"Philosophie",emoji:"🤔"},
    {id:12,nom:"Anglais",emoji:"🇬🇧"},
    {id:13,nom:"Allemand",emoji:"🇩🇪"},
    {id:14,nom:"Chinois",emoji:"🇨🇳"},
    {id:15,nom:"Italien",emoji:"🇮🇹"},
    {id:16,nom:"Arabe",emoji:"🇸🇦"},
    {id:17,nom:"Orientation Scolaire",emoji:"🧭"},
  ];

const _OLD_DEPTS = [
  {id:1,nom:"SVT",emoji:"🌿"},
  {id:2,nom:"Mathématiques",emoji:"📐"},
  {id:3,nom:"Sciences Physiques",emoji:"🧪"},
  {id:4,nom:"Lettres",emoji:"📖"},
  {id:5,nom:"Sciences Humaines",emoji:"🌍"},
  {id:6,nom:"Langues Vivantes",emoji:"🗣️"},
  {id:7,nom:"EPS",emoji:"🏃"},
  {id:8,nom:"Informatique",emoji:"💻"},
];
