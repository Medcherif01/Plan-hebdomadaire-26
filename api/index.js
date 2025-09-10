const express = require('express');
const cors = require('cors');
const fileUpload = require('express-fileupload');
const { sql } = require('@vercel/postgres'); // Librairie Vercel pour PostgreSQL

const app = express();

// --- Middleware ---
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(fileUpload());

// --- Configuration ---
// Note : Les dépendances pour la génération de fichiers (XLSX, PizZip, etc.)
// et pour l'IA Gemini devront être installées si vous implémentez
// la logique de génération de fichiers ci-dessous.

// ===== MODIFICATION : NOUVELLE LISTE D'UTILISATEURS AVEC MOTS DE PASSE ET RÔLES =====
// Note : Il est fortement recommandé de changer ces mots de passe par défaut !
const users = {
    // Admins
    "Mohamed":    { password: "mohamed123", role: "Admin" },
    "Zohra":      { password: "zohra123",   role: "Admin" },
    // Utilisateurs
    "Abeer":      { password: "abeer123",      role: "User" },
    "Aichetou":   { password: "aichetou123",   role: "User" },
    "Amal":       { password: "amal123",       role: "User" },
    "Amal Najar": { password: "amalnajar123",  role: "User" },
    "Ange":       { password: "ange123",       role: "User" },
    "Anouar":     { password: "anouar123",     role: "User" },
    "Emen":       { password: "emen123",       role: "User" },
    "Farah":      { password: "farah123",      role: "User" },
    "Fatima":     { password: "fatima123",     role: "User" },
    "Ghadah":     { password: "ghadah123",     role: "User" },
    "Hana":       { password: "hana123",       role: "User" },
    "Nada":       { password: "nada123",       role: "User" },
    "Raghd":      { password: "raghd123",      role: "User" },
    "Salma":      { password: "salma123",      role: "User" },
    "Sara":       { password: "sara123",       role: "User" },
    "Souha":      { password: "souha123",      role: "User" },
    "Takwa":      { password: "takwa123",      role: "User" }
};

// --- Routes de l'API ---

// ===== MODIFICATION : LA ROUTE DE LOGIN EST MISE À JOUR POUR GÉRER LES RÔLES =====
app.post('/api/login', (req, res) => {
    const { username, password } = req.body;
    const user = users[username];

    // Vérifie si l'utilisateur existe et si le mot de passe correspond
    if (user && user.password === password) {
        // En cas de succès, renvoie le nom d'utilisateur et son rôle
        res.status(200).json({ 
            success: true, 
            username: username,
            role: user.role // Envoie le rôle au client (front-end)
        });
    } else {
        // En cas d'échec, renvoie un message d'erreur
        res.status(401).json({ success: false, message: 'Identifiants invalides' });
    }
});

// GET /api/plans/:week - Récupère les données pour une semaine donnée
app.get('/api/plans/:week', async (req, res) => {
    const { week } = req.params;
    const weekNumber = parseInt(week, 10);
    if (isNaN(weekNumber)) return res.status(400).json({ message: 'Semaine invalide.' });

    try {
        const { rows } = await sql`SELECT data, class_notes FROM weekly_plans WHERE week = ${weekNumber};`;
        if (rows.length > 0) {
            res.status(200).json({
                planData: rows[0].data || [],
                classNotes: rows[0].class_notes || {}
            });
        } else {
            // Si aucune donnée n'existe pour cette semaine, renvoyer des tableaux/objets vides
            res.status(200).json({ planData: [], classNotes: {} });
        }
    } catch (error) {
        console.error('Erreur SQL /plans/:week:', error);
        res.status(500).json({ message: 'Erreur serveur.' });
    }
});

// POST /api/save-plan - Utilisé par l'admin pour charger un fichier Excel complet
app.post('/api/save-plan', async (req, res) => {
    const { week, data } = req.body;
    const weekNumber = parseInt(week, 10);
    if (isNaN(weekNumber) || !Array.isArray(data)) return res.status(400).json({ message: 'Données invalides.' });

    const jsonData = JSON.stringify(data);
    try {
        // Utilise "ON CONFLICT" pour INSERER une nouvelle semaine ou METTRE À JOUR une semaine existante
        await sql`
            INSERT INTO weekly_plans (week, data, class_notes)
            VALUES (${weekNumber}, ${jsonData}, '{}')
            ON CONFLICT (week)
            DO UPDATE SET data = EXCLUDED.data;
        `;
        res.status(200).json({ message: `Plan S${weekNumber} enregistré.` });
    } catch (error) {
        console.error('Erreur SQL /save-plan:', error);
        res.status(500).json({ message: 'Erreur serveur.' });
    }
});

// POST /api/save-row - Utilisé pour sauvegarder une seule ligne modifiée
app.post('/api/save-row', async (req, res) => {
    const { week, data: rowData } = req.body;
    const weekNumber = parseInt(week, 10);
    if (isNaN(weekNumber) || typeof rowData !== 'object') return res.status(400).json({ message: 'Données invalides.' });

    try {
        const { rows } = await sql`SELECT data FROM weekly_plans WHERE week = ${weekNumber};`;
        if (rows.length === 0) {
             return res.status(404).json({ message: 'Semaine non trouvée. Impossible d\'enregistrer la ligne.' });
        }

        let planData = rows[0].data || [];
        
        // Fonction utilitaire pour trouver une clé sans se soucier de la casse (ex: 'Classe' vs 'classe')
        const findKey = (obj, target) => Object.keys(obj).find(k => k.trim().toLowerCase() === target.toLowerCase());

        // Trouve l'index de la ligne à mettre à jour en se basant sur ses propriétés uniques
        const rowIndex = planData.findIndex(item =>
            item[findKey(item, 'Enseignant')] === rowData[findKey(rowData, 'Enseignant')] &&
            item[findKey(item, 'Classe')] === rowData[findKey(rowData, 'Classe')] &&
            String(item[findKey(item, 'Période')]) === String(rowData[findKey(rowData, 'Période')]) &&
            item[findKey(item, 'Jour')] === rowData[findKey(rowData, 'Jour')] &&
            item[findKey(item, 'Matière')] === rowData[findKey(rowData, 'Matière')]
        );

        if (rowIndex > -1) {
            // Ligne trouvée : on la met à jour et on ajoute/met à jour la date de modification
            const updatedAtKey = findKey(planData[rowIndex], 'updatedAt') || 'updatedAt';
            planData[rowIndex] = { ...rowData, [updatedAtKey]: new Date().toISOString() };
            
            await sql`UPDATE weekly_plans SET data = ${JSON.stringify(planData)} WHERE week = ${weekNumber};`;
            res.status(200).json({ message: 'Ligne enregistrée.', updatedData: { updatedAt: planData[rowIndex][updatedAtKey] } });
        } else {
            // Ligne non trouvée, ce qui est une erreur dans ce contexte
            res.status(404).json({ message: 'Ligne non trouvée pour la mise à jour.' });
        }
    } catch (error) {
        console.error('Erreur SQL /save-row:', error);
        res.status(500).json({ message: 'Erreur serveur.' });
    }
});

// POST /api/save-notes - Sauvegarde les notes pour une classe et une semaine
app.post('/api/save-notes', async (req, res) => {
    const { week, classe, notes } = req.body;
    const weekNumber = parseInt(week, 10);
    if (isNaN(weekNumber) || !classe) return res.status(400).json({ message: 'Données invalides.' });

    try {
        // S'assure que l'entrée pour la semaine existe avant de la modifier
        await sql`
            INSERT INTO weekly_plans (week, data, class_notes)
            VALUES (${weekNumber}, '[]', '{}')
            ON CONFLICT (week) DO NOTHING;
        `;
        // Met à jour le champ JSONB `class_notes` avec la nouvelle note
        await sql`
            UPDATE weekly_plans
            SET class_notes = class_notes || ${JSON.stringify({[classe]: notes})}
            WHERE week = ${weekNumber};
        `;
        res.status(200).json({ message: 'Notes enregistrées.' });
    } catch (error) {
        console.error('Erreur SQL /save-notes:', error);
        res.status(500).json({ message: 'Erreur serveur.' });
    }
});

// GET /api/all-classes - Récupère la liste unique de toutes les classes pour le menu déroulant de l'admin
app.get('/api/all-classes', async (req, res) => {
    try {
        const { rows } = await sql`
            SELECT DISTINCT value->>'Classe' as classe
            FROM weekly_plans, jsonb_array_elements(data)
            WHERE jsonb_typeof(data) = 'array' AND value->>'Classe' IS NOT NULL
            ORDER BY classe;
        `;
        const classes = rows.map(r => r.classe);
        res.status(200).json(classes);
    } catch (error) {
        console.error('Erreur SQL /api/all-classes:', error);
        res.status(500).json({ message: 'Erreur serveur.' });
    }
});

// --- Routes pour la génération de fichiers (NON IMPLÉMENTÉES) ---
// Ces routes sont définies pour éviter les erreurs "404 Not Found" côté client.
// Vous devrez ajouter la logique de génération de fichiers vous-même.

app.post('/api/generate-word', async (req, res) => {
    console.warn("L'API '/api/generate-word' a été appelée, mais n'est pas implémentée.");
    res.status(501).json({ message: "La génération de documents Word n'est pas encore implémentée sur le serveur." });
});

app.post('/api/generate-excel-workbook', async (req, res) => {
    console.warn("L'API '/api/generate-excel-workbook' a été appelée, mais n'est pas implémentée.");
    res.status(501).json({ message: "La génération de classeurs Excel n'est pas encore implémentée sur le serveur." });
});

app.post('/api/full-report-by-class', async (req, res) => {
    console.warn("L'API '/api/full-report-by-class' a été appelée, mais n'est pas implémentée.");
    res.status(501).json({ message: "La génération de rapports complets n'est pas encore implémentée sur le serveur." });
});

app.post('/api/generate-ai-lesson-plan', async (req, res) => {
    console.warn("L'API '/api/generate-ai-lesson-plan' a été appelée, mais n'est pas implémentée.");
    res.status(501).json({ message: "La génération de plans de cours par IA n'est pas encore implémentée sur le serveur." });
});


// Exporter l'application pour Vercel
module.exports = app;
