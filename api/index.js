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

// Middleware
app.use(cors({
    origin: true,
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With']
}));

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Configuration pour l'upload de fichiers
app.use(fileUpload({
    limits: { 
        fileSize: 50 * 1024 * 1024 // 50MB max
    },
    abortOnLimit: true,
    createParentPath: true,
    parseNested: true
}));

// --- CONFIGURATION ---
const MONGO_URL = process.env.MONGO_URL;
const WORD_TEMPLATE_URL = process.env.WORD_TEMPLATE_URL;
let geminiModel;

// Validation des variables d'environnement
if (!MONGO_URL) {
    console.error('❌ FATAL: MONGO_URL n\'est pas définie dans les variables d\'environnement.');
} else {
    console.log('✅ MONGO_URL configurée');
}

// Initialisation de Gemini AI
if (process.env.GEMINI_API_KEY) {
    try {
        const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
        geminiModel = genAI.getGenerativeModel({ model: "gemini-1.5-flash-latest" });
        console.log('✅ SDK Google Gemini initialisé.');
    } catch (e) { 
        console.error("❌ Erreur initialisation Gemini:", e.message); 
    }
} else {
    console.warn(⚠️ GEMINI_API_KEY non défini - Fonctionnalités IA désactivées.');
}

// Dates des semaines
const specificWeekDateRangesNode = {
    1:{start:'2025-08-31',end:'2025-09-04'}, 2:{start:'2025-09-07',end:'2025-09-11'}, 
    3:{start:'2025-09-14',end:'2025-09-18'}, 4:{start:'2025-09-21',end:'2025-09-25'}, 
    5:{start:'2025-09-28',end:'2025-10-02'}, 6:{start:'2025-10-05',end:'2025-10-09'}, 
    7:{start:'2025-10-12',end:'2025-10-16'}, 8:{start:'2025-10-19',end:'2025-10-23'}, 
    9:{start:'2025-10-26',end:'2025-10-30'}, 10:{start:'2025-11-02',end:'2025-11-06'}, 
    11:{start:'2025-11-09',end:'2025-11-13'}, 12:{start:'2025-11-16',end:'2025-11-20'}, 
    13:{start:'2025-11-23',end:'2025-11-27'}, 14:{start:'2025-11-30',end:'2025-12-04'}, 
    15:{start:'2025-12-07',end:'2025-12-11'}, 16:{start:'2025-12-14',end:'2025-12-18'}, 
    17:{start:'2025-12-21',end:'2025-12-25'}, 18:{start:'2025-12-28',end:'2026-01-01'}, 
    19:{start:'2026-01-04',end:'2026-01-08'}, 20:{start:'2026-01-11',end:'2026-01-15'}, 
    21:{start:'2026-01-18',end:'2026-01-22'}, 22:{start:'2026-01-25',end:'2026-01-29'}, 
    23:{start:'2026-02-01',end:'2026-02-05'}, 24:{start:'2026-02-08',end:'2026-02-12'}, 
    25:{start:'2026-02-15',end:'2026-02-19'}, 26:{start:'2026-02-22',end:'2026-02-26'}, 
    27:{start:'2026-03-01',end:'2026-03-05'}, 28:{start:'2026-03-08',end:'2026-03-12'}, 
    29:{start:'2026-03-15',end:'2026-03-19'}, 30:{start:'2026-03-22',end:'2026-03-26'}, 
    31:{start:'2026-03-29',end:'2026-04-02'}, 32:{start:'2026-04-05',end:'2026-04-09'}, 
    33:{start:'2026-04-12',end:'2026-04-16'}, 34:{start:'2026-04-19',end:'2026-04-23'}, 
    35:{start:'2026-04-26',end:'2026-04-30'}, 36:{start:'2026-05-03',end:'2026-05-07'}, 
    37:{start:'2026-05-10',end:'2026-05-14'}, 38:{start:'2026-05-17',end:'2026-05-21'}, 
    39:{start:'2026-05-24',end:'2026-05-28'}, 40:{start:'2026-05-31',end:'2026-06-04'}, 
    41:{start:'2026-06-07',end:'2026-06-11'}, 42:{start:'2026-06-14',end:'2026-06-18'}, 
    43:{start:'2026-06-21',end:'2026-06-25'}, 44:{start:'2026-06-28',end:'2026-07-02'}, 
    45:{start:'2026-07-05',end:'2026-07-09'}, 46:{start:'2026-07-12',end:'2026-07-16'}, 
    47:{start:'2026-07-19',end:'2026-07-23'}, 48:{start:'2026-07-26',end:'2026-07-30'}
};

// Utilisateurs et Admins
const validUsers = {
    "Mohamed": "Mohamed", "Zohra": "Zohra",
    "Abas": "Abas", "Jaber": "Jaber", "Kamel": "Kamel", "Majed": "Majed", 
    "Mohamed Ali": "Mohamed Ali", "Morched": "Morched", "Saeed": "Saeed", 
    "Sami": "Sami", "Sylvano": "Sylvano", "Tonga": "Tonga", "Youssef": "Youssef", 
    "Zine": "Zine", "Abeer": "Abeer", "Aichetou": "Aichetou", "Amal": "Amal", 
    "Amal Arabic": "Amal Arabic", "Ange": "Ange", "Anouar": "Anouar", "Emen": "Emen", 
    "Farah": "Farah", "Fatima Islamic": "Fatima Islamic", "Ghadah": "Ghadah", 
    "Hana - Ameni - PE": "Hana - Ameni - PE", "Nada": "Nada", "Raghd ART": "Raghd ART", 
    "Salma": "Salma", "Sara": "Sara", "Souha": "Souha", "Takwa": "Takwa", 
    "Zohra Zidane": "Zohra Zidane"
};

// Connexion MongoDB avec gestion d'erreur améliorée
let cachedDb = null;
let mongoClient = null;

async function connectToDatabase() {
    if (cachedDb && mongoClient) {
        try {
            // Test de la connexion existante
            await cachedDb.admin().ping();
            return cachedDb;
        } catch (error) {
            console.log('🔄 Connexion MongoDB expirée, reconnexion...');
            cachedDb = null;
            mongoClient = null;
        }
    }
    
    try {
        console.log('🔗 Tentative de connexion à MongoDB...');
        mongoClient = new MongoClient(MONGO_URL, {
            maxPoolSize: 10,
            serverSelectionTimeoutMS: 10000,
            socketTimeoutMS: 45000,
            connectTimeoutMS: 10000,
            maxIdleTimeMS: 30000
        });
        
        await mongoClient.connect();
        const db = mongoClient.db();
        
        // Test de la connexion
        await db.admin().ping();
        
        cachedDb = db;
        console.log('✅ Connexion MongoDB établie avec succès');
        return db;
    } catch (error) {
        console.error('❌ Erreur connexion MongoDB:', error.message);
        cachedDb = null;
        mongoClient = null;
        throw new Error(`Erreur de connexion à la base de données: ${error.message}`);
    }
}

// Fonctions Utilitaires
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

const findKey = (obj, target) => {
    if (!obj || typeof obj !== 'object' || !target) return undefined;
    return Object.keys(obj).find(k => k && k.trim().toLowerCase() === target.trim().toLowerCase());
};

// Middleware de validation
function validateSection(req, res, next) {
    const section = req.query.section || req.body.section;
    if (!section || !['garcons', 'filles'].includes(section)) {
        return res.status(400).json({ 
            message: 'Section invalide. Doit être "garcons" ou "filles".' 
        });
    }
    next();
}

function validateWeek(req, res, next) {
    const week = parseInt(req.params.week || req.body.week);
    if (!week || week < 1 || week > 48) {
        return res.status(400).json({ 
            message: 'Semaine invalide. Doit être entre 1 et 48.' 
        });
    }
    next();
}

// --- ROUTES API ---

// Route de santé
app.get('/api/health', (req, res) => {
    res.json({ 
        status: 'OK', 
        timestamp: new Date().toISOString(),
        mongodb: cachedDb ? 'Connected' : 'Disconnected',
        gemini: geminiModel ? 'Available' : 'Unavailable'
    });
});

// Route de test MongoDB
app.get('/api/test-db', async (req, res) => {
    try {
        const db = await connectToDatabase();
        const result = await db.admin().ping();
        res.json({ 
            status: 'MongoDB OK', 
            ping: result,
            timestamp: new Date().toISOString()
        });
    } catch (error) {
        res.status(500).json({ 
            status: 'MongoDB Error', 
            error: error.message,
            timestamp: new Date().toISOString()
        });
    }
});

// Login
app.post('/api/login', (req, res) => {
    try {
        const { username, password } = req.body;
        
        if (!username || !password) {
            return res.status(400).json({ 
                success: false, 
                message: 'Nom d\'utilisateur et mot de passe requis' 
            });
        }

        // Validation des identifiants
        if (validUsers[username] && validUsers[username] === password) {
            console.log(`✅ Connexion réussie pour l'utilisateur: ${username}`);
            res.status(200).json({ 
                success: true, 
                username: username,
                timestamp: new Date().toISOString()
            });
        } else {
            console.log(`❌ Tentative de connexion échouée pour: ${username}`);
            res.status(401).json({ 
                success: false, 
                message: 'Identifiants invalides' 
            });
        }
    } catch (error) {
        console.error('❌ Erreur login:', error);
        res.status(500).json({ 
            success: false, 
            message: 'Erreur serveur lors de la connexion' 
        });
    }
});

// Récupérer les plans pour une semaine
app.get('/api/plans/:week', validateWeek, validateSection, async (req, res) => {
    try {
        const { week } = req.params;
        const { section } = req.query;
        
        console.log(`📖 Récupération des plans - Semaine: ${week}, Section: ${section}`);
        
        const db = await connectToDatabase();
        const planDocument = await db.collection('plans').findOne({ 
            week: parseInt(week), 
            section: section 
        });

        const result = {
            planData: planDocument?.data || [],
            classNotes: planDocument?.classNotes || {},
            week: parseInt(week),
            section: section,
            found: !!planDocument,
            timestamp: new Date().toISOString()
        };

        console.log(`📊 Plans trouvés: ${result.planData.length} éléments`);
        
        res.status(200).json(result);
    } catch (error) {
        cone.error('❌ Erreur récupération plans:', error);
        res.status(500).json({ 
            message: 'Erreur serveur lors de la récupération des plans.',
            error: error.message
        });
    }
});

// Sauvegarder un plan complet
app.post('/api/save-plan', validateSection, (req, res) => {
    const { week, data, section } = req.body;
    
    // Validation des données
    if (!week || !data || !section) {
        return res.status(400).json({ 
            message: 'Données manquantes (week, data, section requis).' 
        });
    }

    if (!Array.isArray(data)) {
        return res.status(400).json({ 
            message: 'Les données doivent être un tableau.' 
        });
    }

    if (data.length === 0) {
        return res.status(400).json({ 
            message: 'Le tableau de données ne peut pas être vide.' 
        });
    }

    // Validation de la semaine
    const weekNum = parseInt(week);
    if (isNaN(weekNum) || weekNum < 1 || weekNum > 48) {
        return res.status(400).json({ 
            message: 'Numéro de semaine invalide (doit être entre 1 et 48).' 
        });
    }

    console.log(`💾 Sauvegarde plan - Semaine: ${weekNum}, Section: ${section}, Éléments: ${data.length}`);

    // Traitement asynchrone
    (async () => {
        try {
            const db = await connectToDatabase();
            
            // Ajouter timestamp et validation à chaque élément
            const dataWithTimestamp = data.map((item, index) => {
                if (!item || typeof item !== 'object') {
                    throw new Error(`Élément invalide à l'index ${index}`);
                }
                
                return {
                    ...item,
                    updatedAt: new Date(),
                    createdAt: item.createdAt || new Date(),
                    week: weekNum,
                    section: section
                };
            });

            // Sauvegarde en base
            const result = await db.collection('plans').updateOne(
                { week: weekNum, section: section },
                { 
                    $set: { 
                        data: dataWithTimestamp, 
                        section: section,
                        week: weekNum,
                        lastModified: new Date(),
                        dataCount: dataWithTimestamp.length
                   } 
                },
                { upsert: true }
            );

            console.log(`✅ Plan sauvegardé - Résultat:`, {
                matched: result.matchedCount,
                modified: result.modifiedCount,
                upserted: result.upsertedCount
            });
            
            res.status(200).json({ 
                message: `Plan enregistré avec succès pour la semaine ${weekNum}.`,
                elementsCount: dataWithTimestamp.length,
                upserted: result.upsertedCount > 0,
                modified: result.modifiedCount > 0,
                timestamp: new Date().toISOString()
            });
            
        } catch (error) {
            console.error('❌ Erreur sauvegarde plan:', error);
            if (!res.headersSent) {
                res.status(500).json({ 
                    message: 'Erreur serveur lors de la sauvegarde du plan.',
                    error: error.message
                });
            }
        }
    })();
});

// Sauvegarder une ligne individuelle
app.post('/api/save-row', validateSection, async (req, res) => {
    try {
        const { week, data: rowData, section } = req.body;
        
        if (!week || !rowData || !section) {
            return res.status(400).json({ 
                message: 'Données manquantes (week, data, section requis).' 
            });
        }

        if (typeof rowData !== 'object') {
            return res.status(400).json({ 
                message: 'Les données de la ligne doivent être un objet.' 
            });
        }

        const weekNum = parseInt(week);
        if (isNaN(weekNum) || weekNum < 1 || weekNum > 48) {
            return res.status(400).json({ 
                message: 'Numéro de semaine invalide.' 
            });
        }

        console.log(`💾 Sauvegarde ligne - Semaine: ${weekNum}, Section: ${section}`);

        const db = await connectToDatabase();
        const now = new Date();
        
        // Construire les champs de mise à jour
        const updateFields = {};
        for (const key in rowData) {
            if (key !== '_id' && key !== 'id') {
                updateFields[`data.$[elem].${key}`] = rowData[key];
            }
        }
        updateFields['data.$[elem].updatedAt'] = now;

        // Construire les filtres pour identifier la ligne
        const enseignantKey = findKey(rowData, 'Enseignant');
        const classeKey = findKey(rowData, 'Classe');
        const jourKey = findKey(rowData, 'Jour');
        const periodeKey = findKey(rowData, 'Période');
        const matiereKey = findKey(rowData, 'Matière');

        if (!enseignantKey || !classeKey || !jourKey || !periodeKey || !matiereKey) {
            return res.status(400).json({ 
                message: 'Données de ligne incomplètes (Enseignant, Classe, Jour, Période, Matière requis).' 
            });
        }

        const arrayFilters = [{
            "elem.Enseignant": rowData[enseignantKey],
            "elem.Classe": rowData[classeKey],
            "elem.Jour": rowData[jourKey],
            "elem.Période": rowData[periodeKey],
            "elem.Matière": rowData[matiereKey]
        }];

        const result = await db.collection('plans').updateOne(
            { week: weekNum, section: section },
            { $set: updateFields },
            { arrayFilters: arrayFilters }
        );

        if (result.matchedCount > 0) {
            console.log(`✅ Ligne mise à jour - Modified: ${result.modifiedCount}`);
            res.status(200).json({ 
                message: 'Ligne enregistrée avec succès.', 
                updatedData: { updatedAt: now },
                modified: result.modifiedCount > 0
            });
        } else {
            console.log(`⚠️ Ligne non trouvée pour mise à jour`);
            res.status(404).json({ 
                message: 'Ligne non trouvée pour mise à jour.' 
            });
        }
    } catch (error) {
        console.error('❌ Erreur sauvegarde ligne:', error);
        res.status(500).json({ 
            message: 'Erreur serveur lors de la sauvegarde de la ligne.',
            error: error.message
        });
    }
});

// Sauvegarder les notes de classe
app.post('/api/save-notes', validateSection, async (req, res) => {
    try {
        const { week, classe, notes, section } = req.body;
        
        if (!week || !classe || !section) {
            return res.status(400).json({ 
                message: 'Données manquantes (week, classe, section requis).' 
            });
        }

        const weekNum = parseInt(week);
        if (isNaN(weekNum) || weekNum < 1 || weekNum > 48) {
            return res.status(400).json({ 
                message: 'Numéro de semaine invalide.' 
            });
        }

        console.log(`📝 Sauvegde notes - Semaine: ${weekNum}, Classe: ${classe}, Section: ${section}`);

        const db = await connectToDatabase();
        
        const result = await db.collection('plans').updateOne(
            { week: weekNum, section: section }, 
            { 
                $set: { 
                    [`classNotes.${classe}`]: notes || '', 
                    section: section,
                    week: weekNum,
                    lastModified: new Date()
                } 
            }, 
            { upsert: true }
        );

        console.log(`✅ Notes sauvegardées - Matched: ${result.matchedCount}, Modified: ${result.modifiedCount}`);

        res.status(200).json({ 
            message: 'Notes enregistrées avec succès.',
            upserted: result.upsertedCount > 0,
            modified: result.modifiedCount > 0
        });
    } catch (error) {
        console.error('❌ Erreur sauvegarde notes:', error);
        res.status(500).json({ 
            message: 'Erreur serveur lors de la sauvegarde des notes.',
            error: error.message
        });
    }
});

// Récupérer toutes les classes
app.get('/api/all-classes', validateSection, async (req, res) => {
    try {
        const { section } = req.query;
        
        console.log(`📚 Récupération classes - Section: ${section}`);

        const db = await connectToDatabase();
        
        const classes = await db.collection('plans').distinct('data.Classe', { 
            section: section, 
            'data.Classe': { $ne: null, $ne: "", $exists: true } 
        });

        console.log(`📊 Classes trouvées: ${classes.length}`);

        res.status(200).json(classes.sort());
    } catch (error) {
        console.error('❌ Erreur récupération classes:', error);
        res.status(500).json({ 
            message: 'Erreur serveur lors de la récupération des classes.',
            error: error.message
        });
    }
});

// Générer document Word
app.post('/api/generate-word', validateSection, async (req, res) => {
    try {
        const { week, classe, data, notes, section } = req.body;
        
        if (!week || !classe || !data || !section) {
            return res.status(400).json({ 
                message: 'Données invalides pour la génération Word (week, classe, data, section requis).' 
            });
        }

        if (!Array.isArray(data) || data.length === 0) {
            return res.status(400).json({ 
                message: 'Données invalides - tleau non vide requis.' 
            });
        }

        console.log(`📄 Génération Word - Semaine: ${week}, Classe: ${classe}, Section: ${section}`);
        
        let templateBuffer;
        try {
            if (!WORD_TEMPLATE_URL) {
                throw new Error('WORD_TEMPLATE_URL n\'est pas configuré sur le serveur.');
            }
            
            console.log('📥 Téléchargement du modèle Word...');
            const response = await fetch(WORD_TEMPLATE_URL);
            if (!response.ok) {
                throw new Error(`Modèle Word introuvable (${response.status}): ${response.statusText}`);
            }
            templateBuffer = Buffer.from(await response.arrayBuffer());
            conole.log('✅ Modèle Word téléchargé');
        } catch (e) { 
            console.error('❌ Erreur téléchargement modèle:', e.message);
            return res.status(500).json({ message: e.message }); 
        }

        const zip = new PizZip(templateBuffer);
        const doc = new Docxtemplater(zip, { 
            paragraphoop: true, 
            linebreaks: true, 
            nullGetter: () => "" 
        });

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
                message: `Dates serveur manquantes pour la semaine ${weekNumber}.` 
            });
        }

        // Recherche des clés dans les données
        const sampleRow = data[0] || {};
        const jourKey = findKey(sampleRow, 'Jour');
        const periodeKey = findKey(sampleRow, 'Période');
        const matiereKey = findKey(sampleRow, 'Matière');
        const leconKey = findKey(sampleRow, 'Leçon');
        const travauxKey = findKey(sampleRow, 'Travaux de classe');
        const supportKey = findKey(sampleRow, 'Support');
        const devoirsKey = findKey(sampleRow, 'Devoirs');
        
        if (!jourKey || !periodeKey || !matiereKey) {
            return res.status(400).json({ 
                message: 'Structure de données invalide - colonnes Jour, Période, Matière requises.' 
            });
        }

        // Groupement par jour
        data.forEach(item => { 
            const day = item[jourKey]; 
            if (day && dayOrder.includes(day)) { 
                if (!groupedByDay[day]) groupedByDay[day] = []; 
                groupedByDay[day].push(item); 
            } 
        });
        
        // Préparation des données pour le template
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
        
        console.log('🔧 Génération du document...');
        doc.render(templateData);
        
        const buf = doc.getZip().generate({ 
            type: 'nodebuffer', 
            compression: 'DEFLATE' 
        });
        
        const filename = `Plan_${section}_S${week}_${classe.replace(/[^a-z0-9]/gi, '_')}.docx`;
        
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Length', buf.length);
        
        console.log(`✅ Document Word généré: ${filename} (${buf.length} bytes)`);
        res.send(buf);
        
    } catch (error) {
        console.error('❌ Erreur génération Word:', error);
        if (!res.headersSent) {
            res.status(500).json({ 
                message: 'Erreur interne lors de la génération Word.',
                error: error.message
            });
        }
    }
});

// Générer fichier Excel complet
app.post('/api/generate-excel-workbook', validateSection, async (req, res) => {
    try {
        const { week, section } = req.body;
        
        if (!week || !section) {
            return res.status(400).json({ 
                message: 'Données invalides (week et section requis).' 
            });
        }

        const weekNum = parseInt(week);
        if (isNaN(weekNum) || weekNum < 1 || weekNum > 48) {
            return res.status(400).json({ 
                message: 'Numéro de semaine invalide.' 
            });
        }
        
        console.log(`📊 Génération Excel - Semaine: ${weekNum}, Section: ${section}`);
        
        const db = await connectToDatabase();
        const planDocument = await db.collection('plans').findOne({ 
            week: weekNum, 
            section: section 
        });
        
        if (!planDocument?.data?.length) {
            return res.status(404).json({ 
              message: 'Aucune donnée trouvée pour cette semaine et section.' 
            });
        }
        
        const headers = [ 
            'Enseignant', 'Jour', 'Période', 'Classe', 'Matière', 
            'Leçon', 'Travaux de classe', 'Support', 'Devoirs' 
        ];
        
        const formattedData = planDocument.data.map(item => { 
            const row = {}; 
            headers.forEach(h => { 
                const key = findKey(item, h); 
                row[h] = key ? (item[key] || '') : ''; 
            }); 
            return row; 
        });
        
        console.log(`📋 Formatage de ${formattedData.length} lignes`);
        
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.json_to_sheet(formattedData, { header: headers });
        
        // Configuration des colonnes
        worksheet['!ols'] = [ 
            { wch: 20 }, { wch: 15 }, { wch: 10 }, { wch: 12 }, { wch: 20 }, 
            { wch: 45 }, { wch: 45 }, { wch: 25 }, { wch: 45 } 
        ];
        
        XLSX.utils.book_append_sheet(workbook, worksheet, `Plan S${week}`);
        
        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        const filename = `Plan_Complet_${section}_S${week}.xlsx`;
        
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Length', buffer.length);
        
        console.log(`✅ Excel généré: ${filename} (${buffer.length} bytes)`);
        res.send(buffer);
        
    } catch (error) {
        console.error('❌ Erreur génération Excel:', error);
        res.status(500).json({ 
            message: 'Erreur interne lors de la génération Excel.',
            error: error.message
        });
    }
});

// Générer rapport complet par classe
app.post('/api/full-report-by-class', validateSection, async (req, res) => {
    try {
        const { classe: requestedClass, section } = req.body;
        
        if (!requestedClass || !section) {
            return res.status(400).json({ 
                message: 'Classe et section requises.' 
            });
        }
        
        console.log(`📈 Génération rapport complet - Classe: ${requestedClass}, Section: ${section}`);
        
        const db = await connectToDatabase();
        const allPlans = await db.collection('plans')
            .find({ section: section })
            .sort({ week: 1 })
            .toArray();
        
        if (!allPlans || allPlans.length === 0) {
            return res.status(404).json({ 
                message: 'Aucune donnée trouvée pour cette section.' 
            });
        }
        
        console.log(`📚 Analyse de ${allPlans.length} plans`);
        
        const dataBySubject = {};
        const monthsFrench = [
            "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", 
            "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
        ];
        
        let totalItems = 0;
        
        allPlans.forEach(plan => {
            const weekNumber = plan.week;
            let monthName = 'N/A';
            
            const weekDates = specificWeekDateRangesNode[weekNumber];
            if (weekDates?.start) { 
                try { 
                    const startDate = new Date(weekDates.start + 'T00:00:00Z'); 
                    monthName = monthsFrench[startDate.getUTCMonth()]; 
                } catch (e) {
                    conse.warn(`⚠️ Erreur parsing date pour semaine ${weekNumber}`);
                }
            }
            
            (plan.data || []).forEach(item => {
                const itemClassKey = findKey(item, 'Classe');
                const itemSubjectKey = findKey(item, 'Matière');
                
                if (itemClassKey && item[itemClassKey] === requestedClass && 
                    itemSubjectKey && item[itemSubjectKey]) {
                    
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
                    totalItems++;
                }
            });
        });
        
        const subjectsFound = Object.keys(dataBySubject);
        if (subjectsFound.length === 0) {
            return res.status(404).json({ 
                message: `Aucune donnée trouvée pour la classe '${requestedClass}' dans la section '${section}'.` 
            });
        }
        
        console.log(`📊 Données trouvées: ${subjectsFound.length} matières, ${totalItems} éléments`);
        
        const workbook = XLSX.utils.book_new();
        const headers = ['Mois', 'Semaine', 'Période', 'Leçon', 'Travaux de classe', 'Support', 'Devoirs'];
        
        subjectsFound.sort().forEach(subject => {
            const safeSheetName = subject.substring(0, 30).eplace(/[*?:/\\\[\]]/g, '_');
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
        res.setHeader('Content-Length', buffer.length);
        
        console.log(`✅ Rapport généré: ${filename} (${buffer.length} bytes)`);
        res.send(buffer);
        
    } catch (error) {
        console.error('❌ Erreur génération rapport:', error);
        res.status(500).json({ 
            message: 'Erreur interne lors de la génération du rapport.',
            error: error.message
        });
    }
});

// Route pour l'IA (placeholder)
app.post('/api/generate-ai-lesson-plan', async (req, res) => {
    if (!geminiModel) {
        return res.status(503).json({ 
            message: "Service IA non configuré sur le serveur." 
        });
    }
    
    res.status(501).json({ 
        message: "Fonctionnalité IA non implémentée dans cette version."
    });
});

// Middleware de gestion des erreurs globales
app.use((error, req, res, next) => {
    console.error('❌ Erreur non gérée:', error);
    
    if (!res.headersSent) {
        res.status(500).json({ 
            message: 'Erreur interne du serveur',
            error: process.env.NODE_ENV === 'development' ? error.message : 'Une erreur est survenue'
        });
    }
});

// Route par défaut pour les routes non trouvées
app.use((req, res) => {
    console.log(`❓ Route non trouvée: ${req.method} ${req.path}`);
    res.status(404).json({ 
        message: 'Route non trouvée',
        path: req.path,
        method: req.method
    });
});

// Gestion propre de l'arrêt du serveur
process.on('SIGINT', async () => {
    console.log('🛑 Arrêt du serveur...');
    if (mongoClient) {
        try {
            await mongoClient.close();
            console.log('✅ Connexion MongoDB fermée');
        } catch (error) {
            console.error('❌ Erreur fermeture MongoDB:', error);
        }
    }
    process.exit(0);
});

// Démarrage du serveur (pour développement local)
const PORT = process.env.PORT || 3000;
if (process.env.NODE_ENV !== 'production') {
    app.listen(PORT, () => {
        console.log(`🚀 Serveur démarré sur le port ${PORT}`);
        console.log(`📍 API disponible sur: http://localhost:${PORT}/api/health`);
    });
}

// Exporter l'app pour Vercel
module.exports = app;
