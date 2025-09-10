// ====== GLOBALS ======
let loggedInUser = null;
let selectedSection = null; // 'garcons' ou 'filles'
let isAdmin = false;
let currentWeek = null;
let currentUserLanguage = 'fr'; // Langue par défaut

// Données séparées par section
const sectionData = {
    garcons: { planData: [], filteredData: [], headers: [], weeklyClassNotes: {} },
    filles:  { planData: [], filteredData: [], headers: [], weeklyClassNotes: {} }
};

// ==========================
// ====== CONFIGURATION ======
// ==========================

// ---- Utilisateurs & Langues ----
const admins = new Set(['Mohamed', 'Zohra']);
const arabicUsers = new Set(['Jaber', 'Majed', 'Saeed', 'Fatima', 'Amal Najar', 'Emen', 'Ghadah', 'Raghd', 'Sara']);
const englishUsers = new Set(['Amal', 'Kamel', 'Abeer', 'Salma']);

// ---- Listes d'enseignants par section (pour la validation à la connexion) ----
const teachers_garcons = new Set(['Abas','Jaber','Kamel','Majed','Mohamed Ali','Morched','Saeed','Sami','Sylvano','Tonga','Youssef','Zine']);
const teachers_filles = new Set(['Abeer','Aichetou','Amal','Amal Arabic', 'Amal Najar', 'Ange','Anouar','Emen','Farah','Fatima Islamic','Ghadah','Hana - Ameni - PE','Nada','Raghd ART','Salma','Sara','Souha','Takwa','Zohra Zidane']);

// ---- Traductions ----
const translations = {
    fr: {
        login_title: "Connexion", section: "Section", boys_section: "Garçons", girls_section: "Filles",
        login_username_label: "Nom d'utilisateur (Enseignant) :", login_password_label: "Mot de passe (idem Nom) :",
        remember_me: "Rester connecté", login_button_text: "Se connecter", main_page_title: "Plans Hebdomadaires",
        logout_button: "Déconnecter", week_label: "Semaine:", select_week: "-- Sélectionnez une semaine --",
        generate_word_button: "Générer Word par Classe", generate_excel_button: "Générer Excel (1 Fichier)",
        save_all_button: "Enregistrer Lignes Affichées", filter_teacher_label: "Enseignant:",
        filter_class_label: "Classe:", filter_material_label: "Matière:", filter_period_label: "Période:",
        filter_day_label: "Jour:", all: "Tous", all_f: "Toutes", day_sun: "Dimanche", day_mon: "Lundi",
        day_tue: "Mardi", day_wed: "Mercredi", day_thu: "Jeudi", notes_for_class: "Notes pour la classe :",
        select_class: "-- Sélectionnez une classe --", select_class_placeholder: "Sélectionnez une classe pour voir ou ajouter des notes...",
        save_notes_button: "Enregistrer Notes", connected_as: "Connecté: {user}", welcome_user: "Bienvenue {user}!",
        headers: { 'Enseignant': 'Enseignant', 'Jour': 'Jour', 'Période': 'Période', 'Classe': 'Classe', 'Matière': 'Matière', 'Leçon': 'Leçon', 'Travaux de classe': 'Travaux de classe', 'Support': 'Support', 'Devoirs': 'Devoirs', 'Actions': 'Actions', 'updatedAt': 'Mis à jour' }
    },
    en: {
        login_title: "Login", section: "Section", boys_section: "Boys", girls_section: "Girls",
        login_username_label: "Username (Teacher):", login_password_label: "Password (same as Name):",
        remember_me: "Remember me", login_button_text: "Login", main_page_title: "Weekly Plans",
        logout_button: "Logout", week_label: "Week:", select_week: "-- Select a week --",
        generate_word_button: "Generate Word by Class", generate_excel_button: "Generate Excel (1 File)",
        save_all_button: "Save Displayed Rows", filter_teacher_label: "Teacher:",
        filter_class_label: "Class:", filter_material_label: "Subject:", filter_period_label: "Period:",
        filter_day_label: "Day:", all: "All", all_f: "All", day_sun: "Sunday", day_mon: "Monday",
        day_tue: "Tuesday", day_wed: "Wednesday", day_thu: "Thursday", notes_for_class: "Notes for class:",
        select_class: "-- Select a class --", select_class_placeholder: "Select a class to view or add notes...",
        save_notes_button: "Save Notes", connected_as: "Connected: {user}", welcome_user: "Welcome {user}!",
        headers: { 'Enseignant': 'Teacher', 'Jour': 'Day', 'Période': 'Period', 'Classe': 'Class', 'Matière': 'Subject', 'Leçon': 'Lesson', 'Travaux de classe': 'Classwork', 'Support': 'Support', 'Devoirs': 'Homework', 'Actions': 'Actions', 'updatedAt': 'Updated At' }
    },
    ar: {
        login_title: "تسجيل الدخول", section: "قسم", boys_section: "بنين", girls_section: "بنات",
        login_username_label: "اسم المستخدم (المعلم):", login_password_label: "كلمة المرور (نفس الاسم):",
        remember_me: "تذكرني", login_button_text: "تسجيل الدخول", main_page_title: "الخطط الأسبوعية",
        logout_button: "تسجيل الخروج", week_label: "الأسبوع:", select_week: "-- اختر أسبوع --",
        generate_word_button: "إنشاء ملف وورد حسب الفصل", generate_excel_button: "إنشاء ملف اكسل (ملف واحد)",
        save_all_button: "حفظ الصفوف المعروضة", filter_teacher_label: "المعلم:",
        filter_class_label: "الفصل:", filter_material_label: "المادة:", filter_period_label: "الحصة:",
        filter_day_label: "اليوم:", all: "الكل", all_f: "الكل", day_sun: "الأحد", day_mon: "الاثنين",
        day_tue: "الثلاثاء", day_wed: "الأربعاء", day_thu: "الخميس", notes_for_class: "ملاحظات للفصل:",
        select_class: "-- اختر فصل --", select_class_placeholder: "اختر فصلًا لعرض أو إضافة ملاحظات...",
        save_notes_button: "حفظ الملاحظات", connected_as: "متصل: {user}", welcome_user: "مرحباً {user}!",
        headers: { 'Enseignant': 'المعلم', 'Jour': 'اليوم', 'Période': 'الحصة', 'Classe': 'الفصل', 'Matière': 'المادة', 'Leçon': 'الدرس', 'Travaux de classe': 'أعمال الفصل', 'Support': 'الدعم', 'Devoirs': 'الواجبات', 'Actions': 'إجراءات', 'updatedAt': 'آخر تحديث' }
    }
};

const t = (key, params = {}) => {
    let text = translations[currentUserLanguage]?.[key] || translations.fr[key] || `[${key}]`;
    for (const p in params) { text = text.replace(`{${p}}`, params[p]); }
    return text;
};

// ---- Autres configurations ----
const weekDates = {
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
const classOrder = ['PEI1', 'PEI2', 'PEI3', 'PEI4', 'PEI5', 'DP1', 'DP2'];
const compareClasses = (a, b) => { const ia = classOrder.indexOf(a), ib = classOrder.indexOf(b); return (ia !== -1 && ib !== -1) ? ia - ib : String(a || '').localeCompare(String(b || '')); };

// =======================
// ====== UTILS & UI ======
// =======================
const setBtnLoading = (btn, loading, iconClass) => { if (!btn) return; btn.disabled = loading; const i = btn.querySelector('i'); if (i) i.className = loading ? 'fas fa-spinner fa-spin' : iconClass; };
const alertMsg = (msg, type = 'success') => { const div = document.getElementById('message-alerte'); div.className = `message-alert-base alert-${type}`; div.textContent = msg; div.style.display = 'block'; setTimeout(() => { div.style.display = 'none'; }, type === 'error' ? 8000 : 5000); };
const formatUpdatedAt = (s) => { if (!s) return ''; const d = new Date(s); if (isNaN(d.getTime())) return ''; return d.toLocaleString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); };
const findH = (k) => sectionData[selectedSection]?.headers.find(h => h.trim().toLowerCase() === String(k).toLowerCase());

function applyTranslations() {
    document.documentElement.lang = currentUserLanguage;
    document.body.dir = currentUserLanguage === 'ar' ? 'rtl' : 'ltr';

    document.querySelectorAll('[data-key]').forEach(el => {
        el.textContent = t(el.dataset.key);
    });

    document.querySelectorAll('[data-placeholder-key]').forEach(el => {
        el.placeholder = t(el.dataset.placeholderKey);
    });

    if (loggedInUser) {
        document.getElementById('loggedInUserInfo').textContent = t('connected_as', { user: loggedInUser });
        if (currentWeek) renderTable();
    }
}

// =============================
// ====== LOGIN / LOGOUT ======
// =============================
async function handleLogin() {
    const username = document.getElementById('username').value.trim();
    const password = document.getElementById('password').value;
    const section = document.getElementById('sectionSelectorLogin').value;
    const errDiv = document.getElementById('login-error');
    errDiv.textContent = '';

    if (!username || !password) { errDiv.textContent = "Entrez nom d'utilisateur et mot de passe."; return; }

    setBtnLoading(document.getElementById('login-button'), true, 'fas fa-sign-in-alt');
    try {
        const response = await fetch('/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username, password }) });
        const result = await response.json();
        if (!response.ok || !result.success) throw new Error(result.message || 'Échec connexion.');

        loggedInUser = result.username;
        selectedSection = section;
        isAdmin = admins.has(loggedInUser);

        const userList = selectedSection === 'garcons' ? teachers_garcons : teachers_filles;
        if (!isAdmin && !userList.has(loggedInUser)) {
             throw new Error(`Utilisateur non autorisé pour la section ${section === 'garcons' ? 'Garçons' : 'Filles'}.`);
        }

        if (arabicUsers.has(loggedInUser)) currentUserLanguage = 'ar';
        else if (englishUsers.has(loggedInUser)) currentUserLanguage = 'en';
        else currentUserLanguage = 'fr';
        
        applyTranslations();
        
        document.getElementById('login-form').style.display = 'none';
        document.getElementById('main-content').style.display = 'block';
        ['garcons', 'filles'].forEach(s => {
            document.getElementById(`section-content-${s}`).style.display = s === selectedSection ? 'block' : 'none';
        });

        document.getElementById('main-title').textContent = t('main_page_title');
        
        const weekSelector = document.getElementById('weekSelector');
        weekSelector.innerHTML = `<option value="">${t('select_week')}</option>`;
        for (let i = 1; i <= 48; i++) {
            const option = document.createElement('option');
            option.value = i;
            option.textContent = `${t('week_label').replace(':', '')} ${i}`;
            weekSelector.appendChild(option);
        }
        
        alertMsg(t('welcome_user', { user: loggedInUser }), 'success');

    } catch (e) {
        errDiv.textContent = e.message;
    } finally {
        setBtnLoading(document.getElementById('login-button'), false, 'fas fa-sign-in-alt');
    }
}

function handleLogout() {
    loggedInUser = null; selectedSection = null; currentWeek = null; isAdmin = false;
    document.getElementById('main-content').style.display = 'none';
    document.getElementById('login-form').style.display = 'block';
    document.getElementById('username').value = '';
    document.getElementById('password').value = '';
    currentUserLanguage = 'fr';
    applyTranslations();
}

// ==================================
// ====== DATA & RENDERING ======
// ==================================
async function loadPlanForWeek() {
    const week = document.getElementById('weekSelector').value;
    currentWeek = week;
    const sData = sectionData[selectedSection];

    if (!week) {
        sData.planData = []; sData.filteredData = []; sData.headers = []; sData.weeklyClassNotes = {};
        renderTable(); updateActionButtonsState(false);
        return;
    }

    try {
        const response = await fetch(`/api/plans/${week}?section=${selectedSection}`);
        const result = await response.json();
        if (!response.ok) throw new Error(result.message || 'Erreur serveur.');

        sData.planData = Array.isArray(result.planData) ? result.planData : [];
        sData.weeklyClassNotes = result.classNotes || {};
        sData.headers = sData.planData.length ? Object.keys(sData.planData[0]).filter(k => k !== '_id' && k !== 'id') : [];
        
        populateFilterOptions();
        sortAndDisplay();
        populateNotesClassSelector();
        updateActionButtonsState(true);

    } catch (e) {
        alertMsg(`Erreur chargement S${week}: ${e.message}`, 'error');
    }
}

function sortAndDisplay() {
    const section = selectedSection;
    const sData = sectionData[section];
    let dataToFilter = sData.planData;

    const ensK = findH('Enseignant');
    if (!isAdmin && ensK) {
        dataToFilter = dataToFilter.filter(item => item[ensK] === loggedInUser);
    }
    
    const ensF = document.getElementById(`filterEnseignant_${section}`).value;
    const clsF = document.getElementById(`filterClasse_${section}`).value;
    const matF = document.getElementById(`filterMatiere_${section}`).value;
    const perF = document.getElementById(`filterPeriode_${section}`).value;
    const jF = document.getElementById(`filterJour_${section}`).value;

    const clsK = findH('Classe'), matK = findH('Matière'), perK = findH('Période'), jK = findH('Jour');

    sData.filteredData = dataToFilter.filter(item => {
        const passEns = isAdmin ? (!ensF || (ensK && item[ensK] === ensF)) : true;
        const passCls = !clsF || (clsK && item[clsK] === clsF);
        const passMat = !matF || (matK && item[matK] === matF);
        const passPer = !perF || (perK && String(item[perK]) === perF);
        const passJour = !jF || (jK && item[jK] === jF);
        return passEns && passCls && passMat && passPer && passJour;
    }).sort((a, b) => {
        const dayOrder = { "Dimanche": 1, "Lundi": 2, "Mardi": 3, "Mercredi": 4, "Jeudi": 5 };
        const classComp = compareClasses(a[clsK], b[clsK]);
        if (classComp !== 0) return classComp;
        const dayComp = (dayOrder[a[jK]] || 99) - (dayOrder[b[jK]] || 99);
        if (dayComp !== 0) return dayComp;
        return (parseInt(a[perK], 10) || 0) - (parseInt(b[perK], 10) || 0);
    });

    renderTable();
}

function renderTable() {
    const section = selectedSection;
    const thead = document.querySelector(`#planTable_${section} thead tr`);
    const tbody = document.querySelector(`#planTable_${section} tbody`);
    thead.innerHTML = '';
    tbody.innerHTML = '';
    
    const sData = sectionData[section];
    const data = sData.filteredData;
    if (!data || data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="11" style="text-align:center;">Aucune donnée à afficher.</td></tr>`;
        return;
    }

    const headersToDisplay = ['Enseignant', 'Jour', 'Période', 'Classe', 'Matière', 'Leçon', 'Travaux de classe', 'Support', 'Devoirs', 'Actions', 'updatedAt'];
    const headerTranslations = t('headers');
    
    headersToDisplay.forEach(h => {
        if (h === 'Enseignant' && !isAdmin) return;
        const th = document.createElement('th');
        th.textContent = headerTranslations[h] || h;
        thead.appendChild(th);
    });
    
    const editableKeys = new Set(['Leçon', 'Travaux de classe', 'Support', 'Devoirs']);

    data.forEach((rowData) => {
        const tr = document.createElement('tr');
        
        headersToDisplay.forEach(header => {
            if (header === 'Enseignant' && !isAdmin) return;

            const td = document.createElement('td');
            if (header === 'Actions') {
                const saveBtn = document.createElement('button');
                saveBtn.innerHTML = '<i class="fas fa-check"></i>';
                saveBtn.className = 'save-row-button';
                saveBtn.onclick = () => saveRow(rowData, tr);
                td.appendChild(saveBtn);
            } else {
                const key = findH(header);
                td.textContent = key ? (rowData[key] ?? '') : '';
                if (editableKeys.has(header)) {
                    td.contentEditable = true;
                    td.classList.add('editable');
                    td.addEventListener('input', (e) => { if (key) rowData[key] = e.target.textContent; tr.classList.add('modified'); });
                }
                if (header === 'updatedAt' && key) {
                    td.textContent = formatUpdatedAt(rowData[key]);
                }
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}


// ======================================
// ====== ACTIONS & INTERACTIONS ======
// ======================================
async function saveRow(rowData, tr) {
    const btn = tr.querySelector('.save-row-button');
    setBtnLoading(btn, true, 'fas fa-check');
    try {
        const response = await fetch('/api/save-row', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ week: currentWeek, data: rowData, section: selectedSection })
        });
        const result = await response.json();
        if (!response.ok) throw new Error(result.message || 'Erreur de sauvegarde.');
        
        tr.classList.remove('modified');
        const updKey = findH('updatedAt');
        if (updKey && result.updatedData?.updatedAt) {
            rowData[updKey] = result.updatedData.updatedAt;
            const headerCells = Array.from(tr.closest('table').querySelector('thead tr').cells);
            const updIndex = headerCells.findIndex(cell => cell.textContent === t('headers.updatedAt'));
            if(updIndex !== -1) tr.cells[updIndex].textContent = formatUpdatedAt(result.updatedData.updatedAt);
        }
        alertMsg('Ligne enregistrée.', 'success');
    } catch (e) {
        alertMsg(e.message, 'error');
    } finally {
        setBtnLoading(btn, false, 'fas fa-check');
    }
}

async function generateWordByClasse() {
    const section = selectedSection;
    const dataToExport = sectionData[section].filteredData;
    if (!dataToExport.length) { alertMsg('Aucune donnée à exporter.', 'error'); return; }

    setBtnLoading(document.getElementById(`generateWordBtn_${section}`), true, 'fas fa-file-word');
    const byClass = {};
    const classKey = findH('Classe');
    dataToExport.forEach(r => { const c = r[classKey]; if(c) (byClass[c] ||= []).push(r); });
    
    let ok = 0, err = 0;
    for (const classe of Object.keys(byClass)) {
        try {
            const response = await fetch('/api/generate-word', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    week: currentWeek, classe: classe, data: byClass[classe],
                    notes: sectionData[section].weeklyClassNotes[classe] || '', section: section
                })
            });
            if (!response.ok) throw new Error(`Erreur pour la classe ${classe}`);
            const blob = await response.blob();
            saveAs(blob, `Plan_${section}_S${currentWeek}_${classe}.docx`);
            ok++;
        } catch { err++; }
    }
    alertMsg(err ? `Génération Word: ${ok} succès / ${err} erreurs.` : 'Documents Word générés.', 'success');
    setBtnLoading(document.getElementById(`generateWordBtn_${section}`), false, 'fas fa-file-word');
}


function populateFilterOptions() {
    const section = selectedSection;
    const data = sectionData[section].planData;
    const getUniq = (k) => [...new Set(data.map(i => i[k]).filter(Boolean))].sort();

    const updateSel = (id, opts) => {
        const sel = document.getElementById(id);
        const firstOptKey = sel.querySelector('option')?.dataset.key || 'all';
        const currentValue = sel.value;
        sel.innerHTML = `<option value="">${t(firstOptKey)}</option>`;
        opts.forEach(o => {
            const opt = document.createElement('option'); opt.value = o; opt.textContent = o; sel.appendChild(opt);
        });
        sel.value = opts.includes(currentValue) ? currentValue : "";
    };
    
    const ensFilterContainer = document.getElementById(`filter-item-enseignant-${section}`);
    if (isAdmin) {
        ensFilterContainer.style.display = '';
        updateSel(`filterEnseignant_${section}`, getUniq(findH('Enseignant')));
    } else {
        ensFilterContainer.style.display = 'none';
    }

    updateSel(`filterClasse_${section}`, getUniq(findH('Classe')).sort(compareClasses));
    updateSel(`filterMatiere_${section}`, getUniq(findH('Matière')));
    updateSel(`filterPeriode_${section}`, getUniq(findH('Période')));
}

function populateNotesClassSelector() {
    const section = selectedSection;
    const sel = document.getElementById(`notesClassSelector_${section}`);
    sel.innerHTML = `<option value="">${t('select_class')}</option>`;
    const classes = [...new Set(sectionData[section].planData.map(r => r[findH('Classe')]).filter(Boolean))].sort(compareClasses);
    classes.forEach(c => { const o = document.createElement('option'); o.value = c; o.textContent = c; sel.appendChild(o); });
    document.getElementById(`notesInput_${section}`).value = '';
    document.getElementById(`notesInput_${section}`).disabled = true;
    document.getElementById(`saveNotesBtn_${section}`).disabled = true;
}

function updateActionButtonsState(enabled) {
    if (!selectedSection) return;
    [`generateWordBtn_${selectedSection}`, `generateExcelBtn_${selectedSection}`, `saveAllDisplayedBtn_${selectedSection}`].forEach(id => {
        const btn = document.getElementById(id);
        if (btn) btn.disabled = !enabled;
    });
}

// =======================
// ====== EVENTS ======
// =======================
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('login-button').addEventListener('click', handleLogin);
    document.getElementById('logout-button').addEventListener('click', handleLogout);
    document.getElementById('togglePassword').addEventListener('click', () => {
         const p = document.getElementById('password'); 
         p.type = p.type === 'password' ? 'text' : 'password';
         document.getElementById('togglePassword').className = p.type === 'password' ? 'fas fa-eye password-toggle-icon' : 'fas fa-eye-slash password-toggle-icon';
    });

    ['garcons', 'filles'].forEach(section => {
        document.getElementById(`generateWordBtn_${section}`).addEventListener('click', generateWordByClasse);
        // ... attacher les autres événements (generateExcel, saveAll, etc.) ici ...
        
        // Événements pour les filtres
        ['Enseignant', 'Classe', 'Matiere', 'Periode', 'Jour'].forEach(filter => {
            document.getElementById(`filter${filter}_${section}`).addEventListener('change', sortAndDisplay);
        });
    });
    
    applyTranslations(); // Appliquer la langue par défaut au chargement
});
