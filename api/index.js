// api/index.js
// ———————————————————————————————————————————
const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const fetch = require('node-fetch');
const { MongoClient } = require('mongodb');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// ---------- ENV ----------
const MONGO_URL = process.env.MONGO_URL;
const WORD_TEMPLATE_URL = process.env.WORD_TEMPLATE_URL;          // plan hebdo (existant)
const LESSON_TEMPLATE_URL = process.env.LESSON_TEMPLATE_URL;      // ✅ nouveau : modèle docx Plan de leçon
const GEMINI_API_KEY = process.env.GEMINI_API_KEY;

if (!MONGO_URL) console.error("FATAL: MONGO_URL n'est pas défini.");
if (!LESSON_TEMPLATE_URL) console.warn("⚠️ LESSON_TEMPLATE_URL non défini (génération DOCX plan de leçon indisponible).");

let geminiModel = null;
if (GEMINI_API_KEY) {
  try {
    const genAI = new GoogleGenerativeAI(GEMINI_API_KEY);
    geminiModel = genAI.getGenerativeModel({ model: 'gemini-1.5-flash-latest' });
    console.log('✅ Google Gemini prêt.');
  } catch (e) {
    console.error('❌ Erreur init Gemini:', e);
  }
} else {
  console.warn('⚠️ GEMINI_API_KEY non défini (IA désactivée).');
}

// ---------- DATES SEMAINE (UTC) ----------
const specificWeekDateRanges = {
  1:{start:'2025-08-31',end:'2025-09-04'},2:{start:'2025-09-07',end:'2025-09-11'},3:{start:'2025-09-14',end:'2025-09-18'},
  4:{start:'2025-09-21',end:'2025-09-25'},5:{start:'2025-09-28',end:'2025-10-02'},6:{start:'2025-10-05',end:'2025-10-09'},
  7:{start:'2025-10-12',end:'2025-10-16'},8:{start:'2025-10-19',end:'2025-10-23'},9:{start:'2025-10-26',end:'2025-10-30'},
  10:{start:'2025-11-02',end:'2025-11-06'},11:{start:'2025-11-09',end:'2025-11-13'},12:{start:'2025-11-16',end:'2025-11-20'},
  13:{start:'2025-11-23',end:'2025-11-27'},14:{start:'2025-11-30',end:'2025-12-04'},15:{start:'2025-12-07',end:'2025-12-11'},
  16:{start:'2025-12-14',end:'2025-12-18'},17:{start:'2025-12-21',end:'2025-12-25'},18:{start:'2025-12-28',end:'2026-01-01'},
  19:{start:'2026-01-04',end:'2026-01-08'},20:{start:'2026-01-11',end:'2026-01-15'},21:{start:'2026-01-18',end:'2026-01-22'},
  22:{start:'2026-01-25',end:'2026-01-29'},23:{start:'2026-02-01',end:'2026-02-05'},24:{start:'2026-02-08',end:'2026-02-12'},
  25:{start:'2026-02-15',end:'2026-02-19'},26:{start:'2026-02-22',end:'2026-02-26'},27:{start:'2026-03-01',end:'2026-03-05'},
  28:{start:'2026-03-08',end:'2026-03-12'},29:{start:'2026-03-15',end:'2026-03-19'},30:{start:'2026-03-22',end:'2026-03-26'},
  31:{start:'2026-03-29',end:'2026-04-02'},32:{start:'2026-04-05',end:'2026-04-09'},33:{start:'2026-04-12',end:'2026-04-16'},
  34:{start:'2026-04-19',end:'2026-04-23'},35:{start:'2026-04-26',end:'2026-04-30'},36:{start:'2026-05-03',end:'2026-05-07'},
  37:{start:'2026-05-10',end:'2026-05-14'},38:{start:'2026-05-17',end:'2026-05-21'},39:{start:'2026-05-24',end:'2026-05-28'},
  40:{start:'2026-05-31',end:'2026-06-04'},41:{start:'2026-06-07',end:'2026-06-11'},42:{start:'2026-06-14',end:'2026-06-18'},
  43:{start:'2026-06-21',end:'2026-06-25'},44:{start:'2026-06-28',end:'2026-07-02'},45:{start:'2026-07-05',end:'2026-07-09'},
  46:{start:'2026-07-12',end:'2026-07-16'},47:{start:'2026-07-19',end:'2026-07-23'},48:{start:'2026-07-26',end:'2026-07-30'}
};

// ---------- USERS ----------
const validUsers = {
  Mohamed:"Mohamed", Zohra:"Zohra",
  Abas:"Abas", Jaber:"Jaber", Kamel:"Kamel", Majed:"Majed", "Mohamed Ali":"Mohamed Ali", Morched:"Morched", Saeed:"Saeed", Sami:"Sami", Sylvano:"Sylvano", Tonga:"Tonga", Youssef:"Youssef", Zine:"Zine",
  Abeer:"Abeer", Aichetou:"Aichetou", Amal:"Amal", "Amal Arabic":"Amal Arabic", Ange:"Ange", Anouar:"Anouar", Emen:"Emen", Farah:"Farah", "Fatima Islamic":"Fatima Islamic", Ghadah:"Ghadah", "Hana - Ameni - PE":"Hana - Ameni - PE", Nada:"Nada", "Raghd ART":"Raghd ART", Salma:"Salma", Sara:"Sara", Souha:"Souha", Takwa:"Takwa", "Zohra Zidane":"Zohra Zidane"
};

// ---------- Mongo (cache global pour Vercel) ----------
let _mongo = globalThis.__mongoClient;
let _db = globalThis.__mongoDb;
async function getDb() {
  if (_db) return _db;
  if (!_mongo) {
    _mongo = new MongoClient(MONGO_URL, { serverSelectionTimeoutMS: 5000, maxPoolSize: 5 });
    await _mongo.connect();
    globalThis.__mongoClient = _mongo;
  }
  _db = _mongo.db();
  globalThis.__mongoDb = _db;
  return _db;
}

// ---------- Utils ----------
const findKey = (obj, target) =>
  obj ? Object.keys(obj).find(k => k.trim().toLowerCase() === String(target).toLowerCase()) : undefined;

function formatDateFrench(date) {
  if (!date || isNaN(date.getTime())) return 'Date invalide';
  const days = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
  const months = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
  return `${days[date.getUTCDay()]} ${String(date.getUTCDate()).padStart(2,'0')} ${months[date.getUTCMonth()]} ${date.getUTCFullYear()}`;
}
function getDateForDayName(weekStartDate, dayName) {
  const map = { 'Dimanche':0,'Lundi':1,'Mardi':2,'Mercredi':3,'Jeudi':4 };
  if (!weekStartDate || isNaN(weekStartDate.getTime())) return null;
  const off = map[dayName]; if (off === undefined) return null;
  const d = new Date(Date.UTC(weekStartDate.getUTCFullYear(), weekStartDate.getUTCMonth(), weekStartDate.getUTCDate()));
  d.setUTCDate(d.getUTCDate() + off);
  return d;
}

// =====================================
// ============== ROUTES ===============
// =====================================

app.post('/api/login', (req, res) => {
  const { username, password } = req.body || {};
  if (validUsers[username] && validUsers[username] === password) return res.status(200).json({ success: true, username });
  return res.status(401).json({ success: false, message: 'Identifiants invalides.' });
});

// ------- Plans hebdo (inchangé, avec corrections robustesse) -------
app.get('/api/plans/:week', async (req, res) => {
  try {
    const week = parseInt(req.params.week, 10);
    const section = String(req.query.section || '');
    if (!week || !section) return res.status(400).json({ message: 'Semaine ou section manquante.' });
    const db = await getDb();
    const doc = await db.collection('plans').findOne({ week, section });
    return res.status(200).json({ planData: doc?.data || [], classNotes: doc?.classNotes || {} });
  } catch (e) { console.error('/api/plans', e); return res.status(500).json({ message: 'Erreur serveur.' }); }
});

app.post('/api/save-plan', async (req, res) => {
  try {
    const { week, data, section } = req.body || {};
    if (!week || !Array.isArray(data) || !section) return res.status(400).json({ message: 'Données manquantes.' });
    const db = await getDb();
    await db.collection('plans').updateOne(
      { week: parseInt(week, 10), section },
      { $set: { data, section } },
      { upsert: true }
    );
    return res.status(200).json({ ok: true, message: 'Plan enregistré.' });
  } catch (e) { console.error('/api/save-plan', e); return res.status(500).json({ message: 'Erreur serveur.' }); }
});

app.post('/api/save-row', async (req, res) => {
  try {
    const { week, data: rowData, section } = req.body || {};
    if (!week || !rowData || !section) return res.status(400).json({ message: 'Données manquantes.' });

    const db = await getDb();
    const nowIso = new Date().toISOString();

    const filters = [{
      'elem.Enseignant': rowData[findKey(rowData,'Enseignant')],
      'elem.Classe': rowData[findKey(rowData,'Classe')],
      'elem.Jour': rowData[findKey(rowData,'Jour')],
      'elem.Période': rowData[findKey(rowData,'Période')],
      'elem.Matière': rowData[findKey(rowData,'Matière')]
    }];

    const update = {};
    Object.keys(rowData).forEach(k => { update[`data.$[elem].${k}`] = rowData[k]; });
    update['data.$[elem].updatedAt'] = nowIso;

    const result = await db.collection('plans').updateOne(
      { week: parseInt(week,10), section },
      { $set: update },
      { arrayFilters: filters }
    );
    if (!result.matchedCount) return res.status(404).json({ message: 'Ligne non trouvée.' });
    return res.status(200).json({ message: 'Ligne enregistrée.', updatedData: { updatedAt: nowIso } });
  } catch (e) { console.error('/api/save-row', e); return res.status(500).json({ message: 'Erreur serveur.' }); }
});

app.post('/api/save-notes', async (req, res) => {
  try {
    const { week, classe, notes, section } = req.body || {};
    if (!week || !classe || !section) return res.status(400).json({ message: 'Données manquantes.' });
    const db = await getDb();
    await db.collection('plans').updateOne(
      { week: parseInt(week,10), section },
      { $set: { [`classNotes.${classe}`]: String(notes || ''), section } },
      { upsert: true }
    );
    return res.status(200).json({ ok: true, message: 'Notes enregistrées.' });
  } catch (e) { console.error('/api/save-notes', e); return res.status(500).json({ message: 'Erreur serveur.' }); }
});

app.get('/api/all-classes', async (req, res) => {
  try {
    const section = String(req.query.section || '');
    if (!section) return res.status(400).json({ message: 'Section manquante.' });
    const db = await getDb();
    const classes = await db.collection('plans').distinct('data.Classe', { section, 'data.Classe': { $ne: null, $ne: '' } });
    return res.status(200).json((classes || []).filter(Boolean).sort());
  } catch (e) { console.error('/api/all-classes', e); return res.status(500).json({ message: 'Erreur serveur.' }); }
});

// ------- Exports existants (Word hebdo / Excel / Rapport) -------
app.post('/api/generate-word', async (req, res) => {
  try {
    const { week, classe, data, notes, section } = req.body || {};
    if (!week || !classe || !Array.isArray(data) || !section) return res.status(400).json({ message: 'Données invalides.' });
    if (!WORD_TEMPLATE_URL) return res.status(500).json({ message: "WORD_TEMPLATE_URL non configuré." });

    const resp = await fetch(WORD_TEMPLATE_URL);
    if (!resp.ok) return res.status(500).json({ message: 'Modèle Word introuvable.' });
    const templateBuffer = Buffer.from(await resp.arrayBuffer());

    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => '' });

    const weekInfo = specificWeekDateRanges[Number(week)];
    if (!weekInfo?.start || !weekInfo?.end) return res.status(500).json({ message: `Dates serveur manquantes pour S${week}.` });
    const weekStart = new Date(weekInfo.start + 'T00:00:00Z');

    const dayOrder = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi'];
    const kJour = findKey(data[0] || {}, 'Jour');
    const kPer = findKey(data[0] || {}, 'Période');
    const kMat = findKey(data[0] || {}, 'Matière');
    const kLec = findKey(data[0] || {}, 'Leçon');
    const kTrav = findKey(data[0] || {}, 'Travaux de classe');
    const kSup = findKey(data[0] || {}, 'Support');
    const kDev = findKey(data[0] || {}, 'Devoirs');

    const grouped = {};
    data.forEach(r => {
      const d = r[kJour];
      if (dayOrder.includes(d)) { (grouped[d] ||= []).push(r); }
    });

    const joursData = dayOrder.map(day => {
      const rows = (grouped[day] || []).sort((a,b)=>(parseInt(a[kPer],10)||0)-(parseInt(b[kPer],10)||0));
      if (!rows.length) return null;
      const dt = getDateForDayName(weekStart, day);
      return {
        jourDateComplete: dt ? formatDateFrench(dt) : day,
        matieres: rows.map(r => ({
          matiere: r[kMat] ?? '',
          Lecon: r[kLec] ?? '',
          travailDeClasse: r[kTrav] ?? '',
          Support: r[kSup] ?? '',
          devoirs: r[kDev] ?? ''
        }))
      };
    }).filter(Boolean);

    const plage = `du ${formatDateFrench(new Date(weekInfo.start+'T00:00:00Z'))} à ${formatDateFrench(new Date(weekInfo.end+'T00:00:00Z'))}`;

    doc.render({ semaine: Number(week), classe, jours: joursData, notes: String(notes||''), plageSemaine: plage });

    const buf = doc.getZip().generate({ type:'nodebuffer', compression:'DEFLATE' });
    const filename = `Plan_${section}_S${week}_${String(classe).replace(/[^a-z0-9]/gi,'_')}.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    return res.status(200).send(buf);
  } catch (e) { console.error('/api/generate-word', e); if (!res.headersSent) return res.status(500).json({ message:'Erreur interne /generate-word.' }); }
});

app.post('/api/generate-excel-workbook', async (req, res) => {
  try {
    const { week, section } = req.body || {};
    if (!week || !section) return res.status(400).json({ message: 'Données invalides.' });
    const db = await getDb();
    const doc = await db.collection('plans').findOne({ week: parseInt(week,10), section });
    if (!doc?.data?.length) return res.status(404).json({ message: 'Aucune donnée.' });

    const headers = ['Enseignant','Jour','Période','Classe','Matière','Leçon','Travaux de classe','Support','Devoirs'];
    const rows = doc.data.map(item => {
      const row = {}; headers.forEach(h => { const k = findKey(item,h); row[h] = k ? item[k] : ''; }); return row;
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
    ws['!cols'] = [{wch:20},{wch:15},{wch:10},{wch:12},{wch:20},{wch:45},{wch:45},{wch:25},{wch:45}];
    XLSX.utils.book_append_sheet(wb, ws, `Plan S${week}`);

    const buffer = XLSX.write(wb, { bookType:'xlsx', type:'buffer' });
    const filename = `Plan_Complet_${section}_S${week}.xlsx`;
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    return res.status(200).send(buffer);
  } catch (e) { console.error('/api/generate-excel-workbook', e); return res.status(500).json({ message:'Erreur interne Excel.' }); }
});

app.post('/api/full-report-by-class', async (req, res) => {
  try {
    const { classe, section } = req.body || {};
    if (!classe || !section) return res.status(400).json({ message: 'Classe ou section requise.' });
    const db = await getDb();
    const plans = await db.collection('plans').find({ section }).sort({ week: 1 }).toArray();
    if (!plans.length) return res.status(404).json({ message: 'Aucune donnée.' });

    const months = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
    const bySubject = {};
    plans.forEach(p => {
      const w = p.week;
      const info = specificWeekDateRanges[w];
      let month = 'N/A';
      if (info?.start) month = months[new Date(info.start+'T00:00:00Z').getUTCMonth()];
      (p.data || []).forEach(r => {
        const kC = findKey(r,'Classe'), kS = findKey(r,'Matière');
        if (r[kC] === classe && r[kS]) {
          const s = r[kS];
          (bySubject[s] ||= []).push({
            'Mois': month,
            'Semaine': w,
            'Période': r[findKey(r,'Période')] || '',
            'Leçon': r[findKey(r,'Leçon')] || '',
            'Travaux de classe': r[findKey(r,'Travaux de classe')] || '',
            'Support': r[findKey(r,'Support')] || '',
            'Devoirs': r[findKey(r,'Devoirs')] || ''
          });
        }
      });
    });

    const subjects = Object.keys(bySubject);
    if (!subjects.length) return res.status(404).json({ message: `Aucune donnée pour '${classe}'.` });

    const wb = XLSX.utils.book_new();
    const headers = ['Mois','Semaine','Période','Leçon','Travaux de classe','Support','Devoirs'];
    subjects.sort().forEach(s => {
      const ws = XLSX.utils.json_to_sheet(bySubject[s], { header: headers });
      ws['!cols'] = [{wch:12},{wch:10},{wch:10},{wch:40},{wch:40},{wch:25},{wch:40}];
      XLSX.utils.book_append_sheet(wb, ws, s.substring(0,30).replace(/[*?:/\\\[\]]/g,'_'));
    });
    const buffer = XLSX.write(wb, { bookType:'xlsx', type:'buffer' });
    const filename = `Rapport_Complet_${section}_${String(classe).replace(/[^a-z0-9]/gi,'_')}.xlsx`;
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    return res.status(200).send(buffer);
  } catch (e) { console.error('/api/full-report-by-class', e); return res.status(500).json({ message:'Erreur interne du rapport.' }); }
});

// ===============================================================
// =========== IA PLAN DE LEÇON (45 min) -> DOCX ================
// ===============================================================
app.post('/api/generate-ai-lesson-docx', async (req, res) => {
  try {
    if (!geminiModel) return res.status(503).json({ message: 'Service IA non configuré.' });
    if (!LESSON_TEMPLATE_URL) return res.status(503).json({ message: 'Modèle plan de leçon indisponible.' });

    const { row, week, section } = req.body || {};
    if (!row || !week || !section) return res.status(400).json({ message: 'Paramètres manquants (row, week, section).' });

    const enseignant = row[findKey(row,'Enseignant')] || '';
    const classe = row[findKey(row,'Classe')] || '';
    const matiere = row[findKey(row,'Matière')] || '';
    const jour = row[findKey(row,'Jour')] || '';
    const periode = row[findKey(row,'Période')] || '';
    const lecon = row[findKey(row,'Leçon')] || '';
    const titreUnite = row[findKey(row,'Titre de l’unité')] || row[findKey(row,"Titre de l'unité")] || '';

    // Date du jour (à partir de la semaine)
    const info = specificWeekDateRanges[Number(week)];
    if (!info?.start) return res.status(400).json({ message:`Dates semaine S${week} indisponibles.` });
    const start = new Date(info.start + 'T00:00:00Z');
    const dateJour = formatDateFrench(getDateForDayName(start, jour) || start);

    // Prompt (JSON strict)
    const prompt = `
Tu es un enseignant du secondaire. Produis un plan de leçon COMPLET pour une séance de 45 minutes.
Retourne STRICTEMENT du JSON (aucun texte hors JSON) avec les clés EXACTES ci-dessous.

Contexte:
- Matière: ${matiere}
- Classe/Niveau: ${classe}
- Leçon (thème/cible): ${lecon || '(à préciser si vide)'}
- Objectif: atteignable en 45 minutes, avec différenciation.
- Langue: FR.

Format attendu (exemple de structure, remplis avec ton contenu):
{
  "Methodes": "méthodes d'enseignement (ex: découverte guidée, travail en binômes...)",
  "Outils": "outils/supports de travail (ex: manuel, fiche, vidéo, tableau...)",
  "Etapes": [
    { "Minutage": "5",  "Contenu": "Mise en situation / rappel des prérequis", "Ressources": "questionnaire oral, diapos 1-2" },
    { "Minutage": "30", "Contenu": "Développement : activités principales alignées avec la leçon", "Ressources": "manuel p.xx, activité pratique, fiche" },
    { "Minutage": "10", "Contenu": "Consolidation / évaluation rapide", "Ressources": "quiz rapide / sortie ticket" }
  ],
  "Devoirs": "devoirs/quiz/projet à faire à la maison",
  "DiffLents": "mesures pour apprenants lents",
  "DiffTresPerf": "extension pour très performants",
  "DiffTous": "consignes / critères communs à toute la classe"
}
Respecte le total ≈45 min (5/30/10 recommandé, ajuste si pertinent).
Si l'entrée est vide, propose du contenu raisonnable pour ${matiere}.
`.trim();

    const gen = await geminiModel.generateContent(prompt);
    let txt = gen?.response?.text?.() || '';
    const code = txt.match(/```json([\s\S]*?)```/i);
    if (code) txt = code[1].trim();

    let plan;
    try {
      plan = JSON.parse(txt);
    } catch {
      // dernier recours: extraction accolades
      const s = txt.indexOf('{'), e = txt.lastIndexOf('}');
      if (s !== -1 && e !== -1 && e > s) plan = JSON.parse(txt.slice(s, e + 1));
      else return res.status(502).json({ message: 'Réponse IA non JSON.' });
    }

    // Charger modèle DOCX
    const resp = await fetch(LESSON_TEMPLATE_URL);
    if (!resp.ok) return res.status(500).json({ message: 'Modèle DOCX (plan de leçon) introuvable.' });
    const templateBuffer = Buffer.from(await resp.arrayBuffer());

    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => '' });

    // Données pour le template
    const data = {
      Matiere: matiere,
      Classe: classe,
      Semaine: Number(week),
      Jour: jour,
      Date: dateJour,
      Seance: String(periode || ''),
      TitreUnite: titreUnite,
      Lecon: lecon,
      Methodes: String(plan.Methodes || plan.Méthodes || plan["Méthodes d’enseignement"] || ''),
      Outils: String(plan.Outils || plan["Outils de travail"] || ''),
      Etapes: (Array.isArray(plan.Etapes) ? plan.Etapes : []).map(e => ({
        Minutage: String(e.Minutage || ''),
        Contenu: String(e.Contenu || ''),
        Ressources: String(e.Ressources || '')
      })),
      Devoirs: String(plan.Devoirs || ''),
      DiffLents: String(plan.DiffLents || ''),
      DiffTresPerf: String(plan.DiffTresPerf || ''),
      DiffTous: String(plan.DiffTous || ''),
      NomEnseignant: enseignant
    };

    doc.render(data);

    const buf = doc.getZip().generate({ type:'nodebuffer', compression:'DEFLATE' });
    const safeClass = String(classe || 'Classe').replace(/[^a-z0-9]/gi,'_');
    const safeSubj  = String(matiere || 'Matiere').replace(/[^a-z0-9]/gi,'_');
    const filename = `Plan_de_lecon_${section}_S${week}_${safeClass}_${safeSubj}.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    return res.status(200).send(buf);
  } catch (e) {
    console.error('/api/generate-ai-lesson-docx', e);
    return res.status(500).json({ message: 'Erreur lors de la génération du plan de leçon.' });
  }
});

module.exports = app;
