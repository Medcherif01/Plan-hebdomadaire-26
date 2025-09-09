const express = require('express');
const cors = require('cors');
const fileUpload = require('express-fileupload');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { GoogleGenerativeAI } = require("@google/generative-ai");
const fetch = require('node-fetch');
const { MongoClient } = require('mongodb');

const app = express();

// Middleware de base
app.use(cors({
    origin: true,
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(express.json({ 
    limit: '50mb',
    verify: (req, res, buf) => {
        try {
            JSON.parse(buf);
        } catch (e) {
            res.status(400).json({ error: 'Invalid JSON format' });
            return;
        }
    }
}));

app.use(express.urlencoded({ extended: true, limit: '50mb' }));

app.use(fileUpload({
    limits: { fileSize: 50 * 1024 * 1024 },
    abortOnLimit: true,
    responseOnLimit: "File size limit exceeded"
}));

// --- CONFIGURATION ---
const MONGO_URL = process.env.MONGO_URL;
const WORD_TEMPLATE_URL = process.env.WORD_TEMPLATE_URL;
let geminiModel;

// Initialisation Gemini
if (process.env.GEMINI_API_KEY) {
    try {
        const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
        geminiModel = genAI.getGenerativeModel({ model: "gemini-1.5-flash-latest" });
        console.log('✅ SDK Google Gemini initialisé.');
    } catch (e) {
        console.error("Erreur initialisation Gemini:", e);
    }
}

// Dates des semaines
const specificWeekDateRangesNode = {
    1:{start:'2025-08-31',end:'2025-09-04'}, 2:{start:'2025-09-07',end:'2025-09-11'}, 3:{start:'2025-09-14',end:'2025-09-18'}, 4:{start:'2025-09-21',end:'2025-09-25'}, 5:{start:'2025-09-28',end:'2025-10-02'}, 6:{start:'2025-10-05',end:'2025-10-09'}, 7:{start:'2025-10-12',end:'2025-10-16'}, 8:{start:'2025-10-19',end:'2025-10-23'}, 9:{start:'2025-10-26',end:'2025-10-30'}, 10:{start:'2025-11-02',end:'2025-11-06'}, 11:{start:'2025-11-09',end:'2025-11-13'}, 12:{start:'2025-11-16',end:'2025-11-20'}, 13:{start:'2025-11-23',end:'2025-11-27'}, 14:{start:'2025-11-30',end:'2025-12-04'}, 15:{start:'2025-12-07',end:'2025-12-11'}, 16:{start:'2025-12-14',end:'2025-12-18'}, 17:{start:'2025-12-21',end:'2025-12-25'}, 18:{start:'2025-12-28',end:'2026-01-01'}, 19:{start:'2026-01-04',end:'2026-01-08'}, 20:{start:'2026-01-11',end:'2026-01-15'}, 21:{start:'2026-01-18',end:'2026-01-22'}, 22:{start:'2026-01-25',end:'2026-01-29'}, 23:{start:'2026-02-01',end:'2026-02-05'}, 24:{start:'2026-02-08',end:'2026-02-12'}, 25:{start:'2026-02-15',end:'2026-02-19'}, 26:{start:'2026-02-22',end:'2026-02-26'}, 27:{start:'2026-03-01',end:'2026-03-05'}, 28:{start:'2026-03-08',end:'2026-03-12'}, 29:{start:'2026-03-15',end:'2026-03-19'}, 30:{start:'2026-03-22',end:'2026-03-26'}, 31:{start:'2026-03-29',end:'2026-04-02'}, 32:{start:'2026-04-05',end:'2026-04-09'}, 33:{start:'2026-04-12',end:'2026-04-16'}, 34:{start:'2026-04-19',end:'2026-04-23'}, 35:{start:'2026-04-26',end:'2026-04-30'}, 36:{start:'2026-05-03',end:'2026-05-07'}, 37:{start:'2026-05-10',end:'2026-05-14'}, 38:{start:'2026-05-17',end:'2026-05-21'}, 39:{start:'2026-05-24',end:'2026-05-28'}, 40:{start:'2026-05-31',end:'2026-06-04'}, 41:{start:'2026-06-07',end:'2026-06-11'}, 42:{start:'2026-06-14',end:'2026-06-18'}, 43:{start:'2026-06-21',end:'2026-06-25'}, 44:{start:'2026-06-28',end:'2026-07-02'}, 45:{start:'2026-07-05',end:'2026-07-09'}, 46:{start:'2026-07-12',end:'2026-07-16'}, 47:{start:'2026-07-19',end:'2026-07-23'}, 48:{start:'2026-07-26',end:'2026-07-30'}
};

// Utilisateurs valides
const validUsers = {
    "Mohamed": "Mohamed", "Zohra": "Zohra",
    "Abas": "Abas", "Jaber": "Jaber", "Kamel": "Kamel", "Majed": "Majed", "Mohamed Ali": "Mohamed Ali", "Morched": "Morched", "Saeed": "Saeed", "Sami": "Sami", "Sylvano": "Sylvano", "Tonga": "Tonga", "Youssef": "Youssef", "Zine": "Zine",
    "Abeer": "Abeer", "Aichetou": "Aichetou", "Amal": "Amal", "Amal Arabic": "Amal Arabic", "Ange": "Ange", "Anouar": "Anouar", "Emen": "Emen", "Farah": "Farah", "Fatima Islamic": "Fatima Islamic", "Ghadah": "Ghadah", "Hana - Ameni - PE": "Hana - Ameni - PE", "Nada": "Nada", "Raghd ART": "Raghd ART", "Salma": "Salma", "Sara": "Sara", "Souha": "Souha", "Takwa": "Takwa", "Zohra Zidane": "Zohra Zidane"
};

// Connexion MongoDB sécurisée
let cachedDb = null;
let mongoClient = null;

async function connectToDatabase() {
    if (cachedDb && mongoClient) {
        try {
            await cachedDb.admin().ping();
            return cachedDb;
        } catch (error) {
            console.log('Reconnexion MongoDB nécessaire...');
            cachedDb = null;
            mongoClient = null;
        }
    }
    
    if (!MONGO_URL) {
        throw new Error('MONGO_URL non configurée');
    }
    
    try {
        mongoClient = new MongoClient(MONGO_URL, {
            maxPoolSize: 10,
            serverSelectionTimeoutMS: 5000,
            socketTimeoutMS: 45000,
        });
        
        await mongoClient.connect();
        const db = mongoClient.db();
        cachedDb = db;
        console.log('✅ Connexion MongoDB établie');
        return db;
    } catch (error) {
        console.error('❌ Erreur connexion MongoDB:', error);
        throw new Error('Connexion base de données échouée');
    }
}

// Fonctions utilitaires
function formatDateFrenchNode(date) {
    if (!date || isNaN(date.getTime())) return "Date invalide";
    const days = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"];
    const months = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];
    const dayName = days[date.getUTCDay()];
    const dayNum = String(date.getUTCDate()).padStart(2, '0');
    const monthName = months[date.getUTCMonth()];
    const yearNum = date.getUTCFullYear();
    return `${dayName} ${dayNum} ${monthName} ${yearNum}`;
}

function getDateForDayNameNode(weekStartDate, dayName) {
    if (!weekStartDate || isNaN(weekStartDate.getTime())) return null;
    const dayOrder = { "Dimanche": 0, "Lundi": 1, "Mardi": 2, "Mercredi": 3, "Jeudi": 4 };
    const offset = dayOrder[dayName];
    if (offset === undefined) return null;
    const specificDate = new Date(Date.UTC(weekStartDate.getUTCFullYear(), weekStartDate.getUTCMonth(), weekStartDate.getUTCDate()));
    specificDate.setUTCDate(specificDate.getUTCDate() + offset);
    return specificDate;
}

const findKey = (obj, target) => obj ? Object.keys(obj).find(k => k.trim().toLowerCase() === target.toLowerCase()) : undefined;

// Middleware de gestion d'erreurs JSON
function jsonErrorHandler(err, req, res, next) {
    console.error('Erreur serveur:', err);
    
    // S'assurer que la réponse est toujours du JSON
    if (!res.headersSent) {
        res.status(500).json({ 
            success: false,
            message: 'Erreur interne du serveur',
            error: process.env.NODE_ENV === 'development' ? err.message : 'Une erreur est survenue'
        });
    }
}

// --- ROUTES API ---

// Route de test de santé
app.get('/api/health', (req, res) => {
    try {
        res.json({ 
            status: 'OK', 
            timestamp: new Date().toISOString(),
            mongodb: MONGO_URL ? 'configured' : 'not configured'
        });
    } catch (error) {
        res.status(500).json({ 
            status: 'ERROR', 
            message: error.message 
        });
    }
});

// Route de connexion
app.post('/api/login', async (req, res) => {
    try {
        const { username, password } = req.body;
        
        if (!username || !password) {
            return res.status(400).json({ 
                success: false, 
                message: 'Nom d\'utilisateur et mot de passe requis' 
            });
        }
        
        if (validUsers[username] && validUsers[username] === password) {
            res.json({ 
                success: true, 
                username: username 
            });
        } else {
            res.status(401).json({ 
                success: false, 
                message: 'Identifiants invalides' 
            });
        }
    } catch (error) {
        console.error('Erreur login:', error);
        res.status(500).json({ 
            success: false, 
            message: 'Erreur serveur lors de la connexion' 
        });
    }
});

// Récupérer les plans
app.get('/api/plans/:week', async (req, res) => {
    try {
        const { week } = req.params;
        const { section } = req.query;
        
        if (!week || !section) {
            return res.status(400).json({ 
                success: false,
                message: 'Semaine et section requises' 
            });
        }
        
        const db = await connectToDatabase();
        const planDocument = await db.collection('plans').findOne({ 
            week: parseInt(week), 
            section: section 
        });
        
        res.json({
            success: true,
            planData: planDocument?.data || [],
            classNotes: planDocument?.classNotes || {}
        });
        
    } catch (error) {
        console.error('Erreur récupération plans:', error);
        res.status(500).json({ 
            success: false,
            message: 'Erreur lors de la récupération des plans' 
        });
    }
});

// Sauvegarder un plan complet
app.post('/api/save-plan', async (req, res) => {
    try {
        const { week, data, section } = req.body;
        
        if (!week || !data || !section) {
            return res.status(400).json({ 
                success: false,
                message: 'Semaine, données et section requises' 
            });
        }
        
        const db = await connectToDatabase();
        const result = await db.collection('plans').updateOne(
            { week: parseInt(week), section: section },
            { 
                $set: { 
                    data: data, 
                    section: section,
                    updatedAt: new Date()
                } 
            },
            { upsert: true }
        );
        
        console.log(`Plan sauvegardé pour S${week} ${section}`);
        res.json({ 
            success: true,
            message: `Plan enregistré pour la semaine ${week}` 
        });
        
    } catch (error) {
        console.error('Erreur sauvegarde plan:', error);
        res.status(500).json({ 
            success: false,
            message: 'Erreur lors de la sauvegarde du plan' 
        });
    }
});

// Sauvegarder une ligne
app.post('/api/save-row', async (req, res) => {
    try {
        const { week, data: rowData, section } = req.body;
        
        if (!week || !rowData || !section) {
            return res.status(400).json({ 
                success: false,
                message: 'Semaine, données et section requises' 
            });
        }
        
        const db = await connectToDatabase();
        const updateFields = {};
        const now = new Date();
        
        for (const key in rowData) {
            if (key !== '_id' && key !== 'id') {
                updateFields[`data.$[elem].${key}`] = rowData[key];
            }
        }
        updateFields['data.$[elem].updatedAt'] = now;
        
        const arrayFilters = [{
            "elem.Enseignant": rowData[findKey(rowData, 'Enseignant')],
            "elem.Classe": rowData[findKey(rowData, 'Classe')],
            "elem.Jour": rowData[findKey(rowData, 'Jour')],
            "elem.Période": rowData[findKey(rowData, 'Période')],
            "elem.Matière": rowData[findKey(rowData, 'Matière')]
        }];
        
        const result = await db.collection('plans').updateOne(
            { week: parseInt(week), section: section },
            { $set: updateFields },
            { arrayFilters: arrayFilters }
        );
        
        if (result.matchedCount > 0) {
            res.json({ 
                success: true,
                message: 'Ligne enregistrée avec succès', 
                updatedData: { updatedAt: now } 
            });
        } else {
            res.status(404).json({ 
                success: false,
                message: 'Ligne non trouvée pour la mise à jour' 
            });
        }
        
    } catch (error) {
        console.error('Erreur sauvegarde ligne:', error);
        res.status(500).json({ 
            success: false,
            message: 'Erreur lors de la sauvegarde de la ligne' 
        });
    }
});

// Sauvegarder les notes
app.post('/api/save-notes', async (req, res) => {
    try {
        const { week, classe, notes, section } = req.body;
        
        if (!week || !classe || !section) {
            return res.status(400).json({ 
                success: false,
                message: 'Semaine, classe et section requises' 
            });
        }
        
        const db = await connectToDatabase();
        await db.collection('plans').updateOne(
            { week: parseInt(week), section: section },
            { 
                $set: { 
                    [`classNotes.${classe}`]: notes || '', 
                    section: section,
                    updatedAt: new Date()
                } 
            },
            { upsert: true }
        );
        
        res.json({ 
            success: true,
            message: 'Notes enregistrées avec succès' 
        });
        
    } catch (error) {
        console.error('Erreur sauvegarde notes:', error);
        res.status(500).json({ 
            success: false,
            message: 'Erreur lors de la sauvegarde des notes' 
        });
    }
});

// Récupérer toutes les classes
app.get('/api/all-classes', async (req, res) => {
    try {
        const { section } = req.query;
        
        if (!section) {
            return res.status(400).json({ 
                success: false,
                message: 'Section requise' 
            });
        }
        
        const db = await connectToDatabase();
        const classes = await db.collection('plans').distinct('data.Classe', {
            section: section,
            'data.Classe': { $ne: null, $ne: "" }
        });
        
        res.json({
            success: true,
            classes: classes.sort()
        });
        
    } catch (error) {
        console.error('Erreur récupération classes:', error);
        res.status(500).json({ 
            success: false,
            message: 'Erreur lors de la récupération des classes' 
        });
    }
});

// Générer document Word
app.post('/api/generate-word', async (req, res) => {
    try {
        const { week, classe, data, notes, section } = req.body;
        
        if (!week || !classe || !data || !section) {
            return res.status(400).json({ 
                success: false,
                message: 'Données invalides pour la génération Word' 
            });
        }
        
        if (!WORD_TEMPLATE_URL) {
            return res.status(500).json({ 
                success: false,
                message: 'Template Word non configuré' 
            });
        }
        
        const response = await fetch(WORD_TEMPLATE_URL);
        if (!response.ok) {
            return res.status(500).json({ 
                success: false,
                message: 'Template Word introuvable' 
            });
        }
        
        const templateBuffer = Buffer.from(await response.arrayBuffer());
        const zip = new PizZip(templateBuffer);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            nullGetter: () => ""
        });
        
        // Traitement des données pour le template
        const weekNumber = Number(week);
        const groupedByDay = {};
        const dayOrder = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"];
        
        const datesNode = specificWeekDateRangesNode[weekNumber];
        let weekStartDateNode = null;
        if (datesNode?.start) {
            weekStartDateNode = new Date(datesNode.start + 'T00:00:00Z');
        }
        
        if (!weekStartDateNode || isNaN(weekStartDateNode.getTime())) {
            return res.status(500).json({ 
                success: false,
                message: `Dates manquantes pour la semaine ${weekNumber}` 
            });
        }
        
        const sampleRow = data[0] || {};
        const jourKey = findKey(sampleRow, 'Jour');
        const periodeKey = findKey(sampleRow, 'Période');
        const matiereKey = findKey(sampleRow, 'Matière');
        const leconKey = findKey(sampleRow, 'Leçon');
        const travauxKey = findKey(sampleRow, 'Travaux de classe');
        const supportKey = findKey(sampleRow, 'Support');
        const devoirsKey = findKey(sampleRow, 'Devoirs');
        
        data.forEach(item => {
            const day = item[jourKey];
            if (day && dayOrder.includes(day)) {
                if (!groupedByDay[day]) groupedByDay[day] = [];
                groupedByDay[day].push(item);
            }
        });
        
        const joursData = dayOrder.map(dayName => {
            if (!groupedByDay[dayName]) return null;
            
            const dateOfDay = getDateForDayNameNode(weekStartDateNode, dayName);
            const formattedDate = dateOfDay ? formatDateFrenchNode(dateOfDay) : dayName;
            
            const sortedEntries = groupedByDay[dayName].sort((a, b) => 
                (parseInt(a[periodeKey], 10) || 0) - (parseInt(b[periodeKey], 10) || 0)
            );
            
            const matieres = sortedEntries.map(item => ({
                matiere: item[matiereKey] ?? "",
                Lecon: item[leconKey] ?? "",
                travailDeClasse: item[travauxKey] ?? "",
                Support: item[supportKey] ?? "",
                devoirs: item[devoirsKey] ?? ""
            }));
            
            return {
                jourDateComplete: formattedDate,
                matieres: matieres
            };
        }).filter(Boolean);
        
        let plageSemaineText = `Semaine ${weekNumber}`;
        if (datesNode?.start && datesNode?.end) {
            const startD = new Date(datesNode.start + 'T00:00:00Z');
            const endD = new Date(datesNode.end + 'T00:00:00Z');
            if (!isNaN(startD.getTime()) && !isNaN(endD.getTime())) {
                plageSemaineText = `du ${formatDateFrenchNode(startD)} à ${formatDateFrenchNode(endD)}`;
            }
        }
        
        const templateData = {
            semaine: weekNumber,
            classe: classe,
            jours: joursData,
            notes: (notes || ""),
            plageSemaine: plageSemaineText
        };
        
        doc.render(templateData);
        
        const buf = doc.getZip().generate({
            type: 'nodebuffer',
            compression: 'DEFLATE'
        });
        
        const filename = `Plan_${section}_S${week}_${classe.replace(/[^a-z0-9]/gi, '_')}.docx`;
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.send(buf);
        
    } catch (error) {
        console.error('Erreur génération Word:', error);
        if (!res.headersSent) {
            res.status(500).json({ 
                success: false,
                message: 'Erreur lors de la génération Word' 
            });
        }
    }
});

// Générer fichier Excel complet
app.post('/api/generate-excel-workbook', async (req, res) => {
    try {
        const { week, section } = req.body;
        
        if (!week || !section) {
            return res.status(400).json({ 
                success: false,
                message: 'Semaine et section requises' 
            });
        }
        
        const db = await connectToDatabase();
        const planDocument = await db.collection('plans').findOne({
            week: parseInt(week),
            section: section
        });
        
        if (!planDocument?.data?.length) {
            return res.status(404).json({ 
                success: false,
                message: 'Aucune donnée trouvée' 
            });
        }
        
        const headers = ['Enseignant', 'Jour', 'Période', 'Classe', 'Matière', 'Leçon', 'Travaux de classe', 'Support', 'Devoirs'];
        
        const formattedData = planDocument.data.map(item => {
            const row = {};
            headers.forEach(h => {
                const key = findKey(item, h);
                row[h] = key ? item[key] : '';
            });
            return row;
        });
        
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.json_to_sheet(formattedData, { header: headers });
        
        worksheet['!cols'] = [
            { wch: 20 }, { wch: 15 }, { wch: 10 }, { wch: 12 }, { wch: 20 }, 
            { wch: 45 }, { wch: 45 }, { wch: 25 }, { wch: 45 }
        ];
        
        XLSX.utils.book_append_sheet(workbook, worksheet, `Plan S${week}`);
        
        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        const filename = `Plan_Complet_${section}_S${week}.xlsx`;
        
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
        
    } catch (error) {
        console.error('Erreur génération Excel:', error);
        res.status(500).json({ 
            success: false,
            message: 'Erreur lors de la génération Excel' 
        });
    }
});

// Générer rapport complet par classe
app.post('/api/full-report-by-class', async (req, res) => {
    try {
        const { classe: requestedClass, section } = req.body;
        
        if (!requestedClass || !section) {
            return res.status(400).json({ 
                success: false,
                message: 'Classe et section requises' 
            });
        }
        
        const db = await connectToDatabase();
        const allPlans = await db.collection('plans').find({ section: section }).sort({ week: 1 }).toArray();
        
        if (!allPlans || allPlans.length === 0) {
            return res.status(404).json({ 
                success: false,
                message: 'Aucune donnée trouvée' 
            });
        }
        
        const dataBySubject = {};
        const monthsFrench = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];
        
        allPlans.forEach(plan => {
            const weekNumber = plan.week;
            let monthName = 'N/A';
            
            const weekDates = specificWeekDateRangesNode[weekNumber];
            if (weekDates?.start) {
                try {
                    const startDate = new Date(weekDates.start + 'T00:00:00Z');
                    monthName = monthsFrench[startDate.getUTCMonth()];
                } catch (e) {
                    // Ignorer les erreurs de date
                }
            }
            
            (plan.data || []).forEach(item => {
                const itemClassKey = findKey(item, 'Classe');
                const itemSubjectKey = findKey(item, 'Matière');
                
                if (itemClassKey && item[itemClassKey] === requestedClass && itemSubjectKey && item[itemSubjectKey]) {
                    const subject = item[itemSubjectKey];
                    if (!dataBySubject[subject]) dataBySubject[subject] = [];
                    
                    const row = {
                        'Mois': monthName,
                        'Semaine': weekNumber,
                        'Période': item[findKey(item, 'Période')] || '',
                        'Leçon': item[findKey(item, 'Leçon')] || '',
                        'Travaux de classe': item[findKey(item, 'Travaux de classe')] || '',
                        'Support': item[findKey(item, 'Support')] || '',
                        'Devoirs': item[findKey(item, 'Devoirs')] || ''
                    };
                    dataBySubject[subject].push(row);
                }
            });
        });
        
        const subjectsFound = Object.keys(dataBySubject);
        if (subjectsFound.length === 0) {
            return res.status(404).json({ 
                success: false,
                message: `Aucune donnée trouvée pour la classe '${requestedClass}'` 
            });
        }
        
        const workbook = XLSX.utils.book_new();
        const headers = ['Mois', 'Semaine', 'Période', 'Leçon', 'Travaux de classe', 'Support', 'Devoirs'];
        
        subjectsFound.sort().forEach(subject => {
            const safeSheetName = subject.substring(0, 30).replace(/[*?:/\\\[\]]/g, '_');
            const worksheet = XLSX.utils.json_to_sheet(dataBySubject[subject], { header: headers });
            
            worksheet['!cols'] = [
                { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 40 }, 
                { wch: 40 }, { wch: 25 }, { wch: 40 }
            ];
            
            XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName);
        });
        
        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        const filename = `Rapport_Complet_${section}_${requestedClass.replace(/[^a-z0-9]/gi, '_')}.xlsx`;
        
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
        
    } catch (error) {
        console.error('Erreur rapport complet:', error);
        res.status(500).json({ 
            success: false,
            message: 'Erreur lors de la génération du rapport' 
        });
    }
});

// Route IA (placeholder)
app.post('/api/generate-ai-lesson-plan', async (req, res) => {
    try {
        if (!geminiModel) {
            return res.status(503).json({ 
                success: false,
                message: "Service IA non configuré" 
            });
        }
        
        res.status(501).json({ 
            success: false,
            message: "Fonctionnalité IA non implémentée" 
        });
    } catch (error) {
        res.status(500).json({ 
            success: false,
            message: "Erreur service IA" 
        });
    }
});

// Middleware de gestion d'erreurs (doit être en dernier)
app.use(jsonErrorHandler);

// Route 404
app.use((req, res) => {
    res.status(404).json({ 
        success: false,
        message: 'Route non trouvée' 
    });
});

// Export pour Vercel
module.exports = app;

// Pour développement local
if (require.main === module) {
    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => {
        console.log(`🚀 Serveur démarré sur le port ${PORT}`);
    });
}
