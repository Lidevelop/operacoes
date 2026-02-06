// Validação obrigatória do campo VTR nas ocorrências
function validateIncidentsVTR() {
    let valid = true;
    let firstInvalid = null;
    if (incidentsRows.length > 0) {
        incidentsRows.forEach((row, idx) => {
            const select = document.getElementById(`incident-vtr-${row.id}`);
            if (!row.vtr || row.vtr === '') {
                valid = false;
                if (select) {
                    select.classList.add('field-error');
                    if (!firstInvalid) firstInvalid = select;
                }
            } else {
                if (select) select.classList.remove('field-error');
            }
        });
        if (firstInvalid) {
            setTimeout(() => { firstInvalid.focus(); }, 100);
        }
    }
    return valid;
}
// Initialize jsPDF
const { jsPDF } = window.jspdf;

// Firebase configuration (substitua pelos dados do seu projeto)
const firebaseConfig = {
    apiKey: "AIzaSyAJxtfzWip7RLrfItlCV7_Bq3g35Zn20SI",
    authDomain: "operacoes-c6b81.firebaseapp.com",
    projectId: "operacoes-c6b81",
    storageBucket: "operacoes-c6b81.firebasestorage.app",
    messagingSenderId: "522409594781",
    appId: "1:522409594781:web:f35354ca63741bbf87a974",
    measurementId: "G-6718L9HGTG"
};

let auth = null;
let db = null;
let currentOperationId = null;
let operationsCache = [];
let isLoadingOperations = false;
let forcedOperationIdFilter = null;
let idleTimeoutId = null;
const IDLE_LIMIT_MS = 30 * 60 * 1000;
const SESSION_COUNTDOWN_MS = 60 * 60 * 1000; // 1 hora
const SESSION_CONFIRM_AT_MS = 5 * 60 * 1000; // 5 minutos
const SESSION_DEADLINE_KEY = 'pmaracaju_session_deadline';
const OPERATIONS_PAGE_SIZE = 5;
let operationsCurrentPage = 1;
const USERS_PAGE_SIZE = 8;
let usersCurrentPage = 1;

// DOM Elements
const loginView = document.getElementById('loginView');
const appView = document.getElementById('appView');
const loginForm = document.getElementById('loginForm');
const loginEmail = document.getElementById('loginEmail');
const loginPassword = document.getElementById('loginPassword');
const loginError = document.getElementById('loginError');
const logoutBtn = document.getElementById('logoutBtn');
const adminMenuContainer = document.getElementById('adminMenuContainer');
const adminMenuBtn = document.getElementById('adminMenuBtn');
const adminMenu = document.getElementById('adminMenu');
const profileMenuBtn = document.getElementById('profileMenuBtn');
const profileMenu = document.getElementById('profileMenu');
const currentOperationBtn = document.getElementById('currentOperationBtn');
const newOperationBtn = document.getElementById('newOperationBtn');
const viewManagerBtn = document.getElementById('viewManagerBtn');
const viewUsersBtn = document.getElementById('viewUsersBtn');
const viewLogsBtn = document.getElementById('viewLogsBtn');
const operationFormView = document.getElementById('operationFormView');
const managerView = document.getElementById('managerView');
const usersView = document.getElementById('usersView');
const logsView = document.getElementById('logsView');
const operationsList = document.getElementById('operationsList');
const operationsPagination = document.getElementById('operationsPagination');
const filterText = document.getElementById('filterText');
const filterDateFrom = document.getElementById('filterDateFrom');
const filterDateTo = document.getElementById('filterDateTo');
const refreshOperations = document.getElementById('refreshOperations');
const formInputs = document.querySelectorAll('#operationFormView input, #operationFormView textarea, #operationFormView select');
const saveIndicator = document.getElementById('saveIndicator');
const saveStatus = document.getElementById('saveStatus');
const lastSaved = document.getElementById('lastSaved');
const sessionTimer = document.getElementById('sessionTimer');
const sessionTimerValue = document.getElementById('sessionTimerValue');
const sessionConfirmModal = document.getElementById('sessionConfirmModal');
const sessionConfirmContinue = document.getElementById('sessionConfirmContinue');
const sessionConfirmLogout = document.getElementById('sessionConfirmLogout');
const sessionTimerModal = document.getElementById('sessionTimerModal');
const sessionTimerValueModal = document.getElementById('sessionTimerValueModal');
const addAreaBtn = document.getElementById('addAreaBtn');
const addVehicleBtn = document.getElementById('addVehicleBtn');
const addServiceChangeBtn = document.getElementById('addServiceChangeBtn');
const areaTableBody = document.getElementById('areaTableBody');
const areaCardsContainer = document.getElementById('areaCardsContainer');
const vehiclesTableBody = document.getElementById('vehiclesTableBody');
const vehiclesCardsContainer = document.getElementById('vehiclesCardsContainer');
const saveDbBtn = document.getElementById('saveDbBtn');
const clearBtn = document.getElementById('clearBtn');
const confirmationModal = document.getElementById('confirmationModal');
const confirmClear = document.getElementById('confirmClear');
const cancelClear = document.getElementById('cancelClear');
const resetPasswordModal = document.getElementById('resetPasswordModal');
const resetPasswordMessage = document.getElementById('resetPasswordMessage');
const resetPasswordError = document.getElementById('resetPasswordError');
const confirmResetPassword = document.getElementById('confirmResetPassword');
const cancelResetPassword = document.getElementById('cancelResetPassword');
const editLockBanner = document.getElementById('editLockBanner');
const usersTableBody = document.getElementById('usersTableBody');
const usersSearchInput = document.getElementById('usersSearchInput');
const usersPagination = document.getElementById('usersPagination');
const summaryField = document.getElementById('summary');
const insertSummaryTemplateBtn = document.getElementById('insertSummaryTemplateBtn');
const summaryTemplateWarning = document.getElementById('summaryTemplateWarning');
const summaryTemplateInfo = document.getElementById('summaryTemplateInfo');
const logsTableBody = document.getElementById('logsTableBody');
const logsDateFrom = document.getElementById('logsDateFrom');
const logsDateTo = document.getElementById('logsDateTo');
const logsLimit = document.getElementById('logsLimit');
const logsClearBtn = document.getElementById('logsClearBtn');
const logsFilterAll = document.getElementById('logsFilterAll');
const logsFilterAccess = document.getElementById('logsFilterAccess');
const logsFilterReport = document.getElementById('logsFilterReport');
const logsFilterUpdate = document.getElementById('logsFilterUpdate');
const logsFilterDelete = document.getElementById('logsFilterDelete');
const openUserCreateBtn = document.getElementById('openUserCreateBtn');
const profileSummaryClass = document.getElementById('profileSummaryClass');
const profileSummaryWarName = document.getElementById('profileSummaryWarName');
const profileSummaryRegistration = document.getElementById('profileSummaryRegistration');
const profileSummaryType = document.getElementById('profileSummaryType');
const profileModal = document.getElementById('profileModal');
const profileForm = document.getElementById('profileForm');
const profileClass = document.getElementById('profileClass');
const profileWarName = document.getElementById('profileWarName');
const profileRegistration = document.getElementById('profileRegistration');
const profileType = document.getElementById('profileType');
const saveProfileBtn = document.getElementById('saveProfileBtn');
const profileError = document.getElementById('profileError');
const userEditModal = document.getElementById('userEditModal');
const userEditForm = document.getElementById('userEditForm');
const userEditClass = document.getElementById('userEditClass');
const userEditWarName = document.getElementById('userEditWarName');
const userEditRegistration = document.getElementById('userEditRegistration');
const userEditType = document.getElementById('userEditType');
const saveUserEditBtn = document.getElementById('saveUserEditBtn');
const cancelUserEditBtn = document.getElementById('cancelUserEditBtn');
const userEditError = document.getElementById('userEditError');
const userCreateModal = document.getElementById('userCreateModal');
const userCreateForm = document.getElementById('userCreateForm');
const userCreateEmail = document.getElementById('userCreateEmail');
const userCreatePassword = document.getElementById('userCreatePassword');
const userCreateClass = document.getElementById('userCreateClass');
const userCreateWarName = document.getElementById('userCreateWarName');
const userCreateRegistration = document.getElementById('userCreateRegistration');
const userCreateType = document.getElementById('userCreateType');
const saveUserCreateBtn = document.getElementById('saveUserCreateBtn');
const cancelUserCreateBtn = document.getElementById('cancelUserCreateBtn');
const userCreateError = document.getElementById('userCreateError');
const changePasswordBtn = document.getElementById('changePasswordBtn');
const changePasswordModal = document.getElementById('changePasswordModal');
const changePasswordForm = document.getElementById('changePasswordForm');
const currentPassword = document.getElementById('currentPassword');
const newPassword = document.getElementById('newPassword');
const confirmPassword = document.getElementById('confirmPassword');
const savePasswordBtn = document.getElementById('savePasswordBtn');
const cancelPasswordBtn = document.getElementById('cancelPasswordBtn');
const changePasswordError = document.getElementById('changePasswordError');
const eventTypeSelect = document.getElementById('eventType');
const addEventTypeBtn = document.getElementById('addEventTypeBtn');
const operationNumberInput = document.getElementById('operationNumber');
const eventTypeModal = document.getElementById('eventTypeModal');
const eventTypeList = document.getElementById('eventTypeList');
const eventTypeInput = document.getElementById('eventTypeInput');
const eventTypeAddConfirm = document.getElementById('eventTypeAddConfirm');
const closeEventTypeModal = document.getElementById('closeEventTypeModal');
const eventTypeModalError = document.getElementById('eventTypeModalError');
const incidentsNoChange = document.getElementById('incidentsNoChange');
const serviceChangesNoChange = document.getElementById('serviceChangesNoChange');
const incidentsNoChangeRow = document.getElementById('incidentsNoChangeRow');
const serviceChangesNoChangeRow = document.getElementById('serviceChangesNoChangeRow');
const incidentsContainer = document.getElementById('incidentsContainer');
const addIncidentBtn = document.getElementById('addIncidentBtn');

// State variables
let saveTimeout = null;
let isSaving = false;
let sessionIntervalId = null;
let sessionPrompted = false;
let areaRows = [];
let vehicleRows = [];
let serviceChangeRows = [];
let incidentsRows = [];
let currentUserProfile = null;
let isAdmin = false;
let isRestricted = false;
let usersCache = [];
let isLoadingUsers = false;
let resetPasswordTarget = null;
let currentOperationMeta = null;
let editingUserId = null;
let logsCache = [];
let isLoadingLogs = false;
let logsCurrentLimit = 100;
let logsActionFilter = 'all';
let eventTypesCache = [];
let eventTypesCacheTimestamp = 0;
const EVENT_TYPES_CACHE_TTL = 5 * 60 * 1000; // 5 minutos

const CLASSE_OPTIONS = [
    'S INSP 2ª CL',
    'S INSP 3ª CL',
    'SUP',
    'GUARDA AUXILIAR',
    'GM'
];

const EVENT_TYPES_COLLECTION = 'settings';
const EVENT_TYPES_DOC_ID = 'eventTypes';
const SUMMARY_TEMPLATE_TEXT = `Durante o período programado, as equipes da Polícia Municipal estiveram presentes no local designado, realizando patrulhamento preventivo e acompanhamento das atividades relacionadas ao evento/operação. As ações foram executadas conforme o planejamento estabelecido, com foco na preservação da ordem pública, na segurança dos participantes e na proteção do patrimônio.

Foram realizadas rondas periódicas, orientações ao público quando necessário e monitoramento contínuo das áreas relatadas. Não houve intercorrências relevantes. As ocorrências registradas nas seções anteriores do relatório, quando existentes, foram devidamente atendidas e solucionadas pela equipe no local.

Ao término das atividades, o local foi deixado em condições normais, sem registros adicionais, sendo a operação finalizada conforme previsto.`;

function buildClasseOptions(selectedValue) {
    const options = ['<option value="">Selecione</option>'];
    CLASSE_OPTIONS.forEach(option => {
        const isSelected = option === selectedValue ? ' selected' : '';
        options.push(`<option value="${option}"${isSelected}>${option}</option>`);
    });
    return options.join('');
}

function getEventTypesDoc() {
    if (!db) return null;
    return db.collection(EVENT_TYPES_COLLECTION).doc(EVENT_TYPES_DOC_ID);
}

function normalizeEventType(value) {
    return (value || '').trim().toUpperCase();
}

function renderEventTypeOptions(items, selectedValue = '') {
    if (!eventTypeSelect) return;
    const normalizedSelected = normalizeEventType(selectedValue);
    const uniqueItems = Array.from(new Set((items || []).map(normalizeEventType).filter(Boolean)));
    if (normalizedSelected && !uniqueItems.includes(normalizedSelected)) {
        uniqueItems.push(normalizedSelected);
    }
    uniqueItems.sort((a, b) => a.localeCompare(b, 'pt-BR'));

    const options = ['<option value="">Selecione</option>'];
    uniqueItems.forEach(option => {
        const isSelected = option === normalizedSelected ? ' selected' : '';
        options.push(`<option value="${option}"${isSelected}>${option}</option>`);
    });
    eventTypeSelect.innerHTML = options.join('');
}

function renderEventTypeModalList() {
    if (!eventTypeList) return;
    eventTypeList.innerHTML = '';
    if (!eventTypesCache.length) {
        const empty = document.createElement('p');
        empty.className = 'modal-helper';
        empty.textContent = 'Nenhum tipo cadastrado.';
        eventTypeList.appendChild(empty);
        return;
    }
    eventTypesCache.forEach(type => {
        const row = document.createElement('div');
        row.className = 'event-type-item';
        row.innerHTML = `
            <span class="event-type-name">${type}</span>
            <button class="btn-small btn-danger" type="button" data-type="${type}">
                <i class="fas fa-trash"></i>
                Excluir
            </button>
        `;
        eventTypeList.appendChild(row);
    });

    eventTypeList.querySelectorAll('button[data-type]').forEach(button => {
        button.addEventListener('click', () => {
            const type = button.getAttribute('data-type');
            if (!type) return;
            deleteEventType(type);
        });
    });
}

function openEventTypeModal() {
    if (!eventTypeModal) return;
    if (eventTypeModalError) eventTypeModalError.textContent = '';
    if (eventTypeInput) eventTypeInput.value = '';
    renderEventTypeModalList();
    eventTypeModal.style.display = 'flex';
}

function closeEventTypeModalView() {
    if (!eventTypeModal) return;
    eventTypeModal.style.display = 'none';
}

async function ensureEventTypesDocument() {
    const docRef = getEventTypesDoc();
    if (!docRef) return [];
    try {
        const snap = await docRef.get();
        let items = [];
        if (!snap.exists) {
            items = ['JOGOS DE FUTEBOL'];
            await docRef.set({
                items,
                updatedAt: firebase.firestore.FieldValue.serverTimestamp()
            });
        } else {
            items = Array.isArray(snap.data()?.items) ? snap.data().items : [];
            if (items.length === 0) {
                items = ['JOGOS DE FUTEBOL'];
                await docRef.set({
                    items,
                    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
                }, { merge: true });
            }
        }
        return items;
    } catch (error) {
        console.error('Erro ao carregar tipos de evento:', error);
        return [];
    }
}

async function loadEventTypes() {
    if (!db) return;
    const selectedValue = eventTypeSelect?.value || '';
    // Se o cache for recente, usa o cache
    if (eventTypesCache.length > 0 && (Date.now() - eventTypesCacheTimestamp < EVENT_TYPES_CACHE_TTL)) {
        renderEventTypeOptions(eventTypesCache, selectedValue);
        renderEventTypeModalList();
        return;
    }
    const items = await ensureEventTypesDocument();
    eventTypesCache = Array.from(new Set(items.map(normalizeEventType).filter(Boolean)));
    eventTypesCacheTimestamp = Date.now();
    renderEventTypeOptions(eventTypesCache, selectedValue);
    renderEventTypeModalList();
}

async function addEventTypeFromModal() {
    if (!db || !eventTypeInput) return;
    if (eventTypeModalError) eventTypeModalError.textContent = '';
    const normalized = normalizeEventType(eventTypeInput.value);
    if (!normalized) return;

    if (eventTypesCache.includes(normalized)) {
        if (eventTypeModalError) {
            eventTypeModalError.textContent = 'Este tipo já existe.';
        }
        return;
    }

    const docRef = getEventTypesDoc();
    if (!docRef) return;
    try {
        await docRef.set({
            items: firebase.firestore.FieldValue.arrayUnion(normalized),
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        }, { merge: true });
        // Atualiza cache local imediatamente
        eventTypesCache.push(normalized);
        eventTypesCache = Array.from(new Set(eventTypesCache));
        eventTypesCacheTimestamp = Date.now();
        renderEventTypeOptions(eventTypesCache, normalized);
        renderEventTypeModalList();
        eventTypeSelect.value = normalized;
        eventTypeInput.value = '';
        triggerSaveDebounced();
    } catch (error) {
        console.error('Erro ao adicionar tipo de evento:', error);
        if (eventTypeModalError) {
            eventTypeModalError.textContent = 'Não foi possível adicionar o tipo de evento.';
        }
    }
}

async function deleteEventType(type) {
    if (!db) return;
    const normalized = normalizeEventType(type);
    if (!normalized) return;
    const confirmed = window.confirm(`Excluir o tipo "${normalized}"?`);
    if (!confirmed) return;

    const docRef = getEventTypesDoc();
    if (!docRef) return;

    try {
        await docRef.set({
            items: firebase.firestore.FieldValue.arrayRemove(normalized),
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        }, { merge: true });
        // Remove do cache local imediatamente
        eventTypesCache = eventTypesCache.filter(item => item !== normalized);
        eventTypesCacheTimestamp = Date.now();
        if (eventTypeSelect && eventTypeSelect.value === normalized) {
            eventTypeSelect.value = '';
        }
        renderEventTypeOptions(eventTypesCache, eventTypeSelect?.value || '');
        renderEventTypeModalList();
        await registerEventTypeLog({ action: 'delete-event-type', eventType: normalized });
        triggerSaveDebounced();
    } catch (error) {
        console.error('Erro ao excluir tipo de evento:', error);
        if (eventTypeModalError) {
            eventTypeModalError.textContent = 'Não foi possível excluir o tipo de evento.';
        }
    }
// Debounce utilitário para evitar múltiplos saves em sequência
function debounce(func, wait) {
    let timeout;
    return function(...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
    };
}

// Substitua triggerSave por uma versão com debounce
const triggerSaveDebounced = debounce(triggerSave, 2000); // 2 segundos

// Exemplo: onde houver triggerSave(), troque por triggerSaveDebounced() nos pontos de autosave
}

function setOperationNumber(sequence, year) {
    if (!operationNumberInput) return;
    const formatted = `${String(sequence).padStart(4, '0')}/${year}`;
    operationNumberInput.value = formatted;
    operationNumberInput.dataset.sequence = String(sequence);
    operationNumberInput.dataset.year = String(year);
}

function getOperationNumberMeta() {
    if (!operationNumberInput) {
        return { number: '', sequence: null, year: null };
    }
    const number = (operationNumberInput.value || '').trim();
    let sequence = parseInt(operationNumberInput.dataset.sequence, 10);
    let year = parseInt(operationNumberInput.dataset.year, 10);
    if (Number.isNaN(sequence) || Number.isNaN(year)) {
        const match = number.match(/^(\d{4})\/(\d{4})$/);
        if (match) {
            sequence = parseInt(match[1], 10);
            year = parseInt(match[2], 10);
        }
    }
    return {
        number,
        sequence: Number.isNaN(sequence) ? null : sequence,
        year: Number.isNaN(year) ? null : year
    };
}

function getOperationYearForNew() {
    const dateStr = document.getElementById('operationDate')?.value;
    if (dateStr) {
        const date = new Date(`${dateStr}T00:00:00`);
        if (!Number.isNaN(date.getTime())) return date.getFullYear();
    }
    return new Date().getFullYear();
}

async function assignNextOperationNumber() {
    if (!db || !operationNumberInput) return;
    if (operationNumberInput.value && operationNumberInput.value.trim() !== '') return;

    const year = getOperationYearForNew();
    let maxSequence = 0;
    try {
        const snapshot = await db.collection('operations').where('operationYear', '==', year).get();
        snapshot.forEach(doc => {
            const data = doc.data() || {};
            const sequence = parseInt(data.operationSequence, 10) || 0;
            if (sequence > maxSequence) {
                maxSequence = sequence;
            }
        });
    } catch (error) {
        console.error('Erro ao gerar número da operação:', error);
    }
    const nextSequence = maxSequence + 1;
    setOperationNumber(nextSequence, year);
}

// Initialize with default rows
function bootApp() {
    initFirebase();

    // Set today's date as default
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('operationDate').value = today;

    if (eventTypeSelect) {
        eventTypeSelect.value = '';
    }
    assignNextOperationNumber();

    // Load saved data
    loadFormData();

    // Add initial rows if none loaded
    if (areaRows.length === 0) addAreaRow();
    if (vehicleRows.length === 0) addVehicleRow();
    if (serviceChangeRows.length === 0) updateServiceChangesContainer();
    if (incidentsRows.length === 0) updateIncidentsContainer();

    // Update time display
    updateLastSavedTime();

    // Initialize event listeners
    initializeEventListeners();
    initializeAuthListeners();

    // Auto-expand textareas after load/reload
    setTimeout(expandAllTextareas, 0);

    // Start idle watcher
    initializeIdleWatcher();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bootApp);
} else {
    bootApp();
}

function initFirebase() {
    if (!window.firebase) {
        console.error('Firebase SDK não carregado.');
        return;
    }

    if (!firebase.apps.length) {
        firebase.initializeApp(firebaseConfig);
    }

    auth = firebase.auth();
    db = firebase.firestore();
    auth.onAuthStateChanged((user) => {
        if (user) {
            const deadline = getSessionDeadline();
            if (deadline && deadline <= Date.now()) {
                clearSessionDeadline();
                auth.signOut();
                return;
            }
            showAppView();
            loadOperationsList();
            initializeUserProfileFlow(user);
            loadEventTypes();
            assignNextOperationNumber();
            if (deadline) {
                sessionPrompted = false;
                startSessionCountdown();
            } else {
                resetSessionCountdown();
            }
        } else {
            showLoginView();
            resetUserProfileState();
            clearSessionDeadline();
        }
    });
}

function initializeEventListeners() {
    // Save on input with debounce
    formInputs.forEach(input => {
        input.addEventListener('input', () => {
            if (input.classList.contains('field-error') && input.value.trim() !== '') {
                input.classList.remove('field-error');
            }
            triggerSaveDebounced();
        });
    });
    
    // Auto-expand textareas
    const textareas = document.querySelectorAll('#operationFormView textarea');
    textareas.forEach(textarea => {
        // Initialize height on page load
        autoExpandTextarea(textarea);
        
        // Expand on input
        textarea.addEventListener('input', function() {
            autoExpandTextarea(this);
            triggerSaveDebounced();
        });
    });
    
    // Add buttons
    addAreaBtn.addEventListener('click', addAreaRow);
    addVehicleBtn.addEventListener('click', addVehicleRow);
    addServiceChangeBtn.addEventListener('click', addServiceChangeRow);
    if (addIncidentBtn) {
        addIncidentBtn.addEventListener('click', addIncidentRow);
    }
    if (addEventTypeBtn) {
        addEventTypeBtn.addEventListener('click', openEventTypeModal);
    }
    if (eventTypeAddConfirm) {
        eventTypeAddConfirm.addEventListener('click', addEventTypeFromModal);
    }
    if (eventTypeInput) {
        eventTypeInput.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                addEventTypeFromModal();
            }
        });
    }
    if (closeEventTypeModal) {
        closeEventTypeModal.addEventListener('click', closeEventTypeModalView);
    }
    if (eventTypeModal) {
        eventTypeModal.addEventListener('click', (event) => {
            if (event.target === eventTypeModal) {
                closeEventTypeModalView();
            }
        });
    }
    
    // Save to Firebase explicitly
    saveDbBtn.addEventListener('click', async () => {
        if (!auth || !auth.currentUser) return;
        if (!validateVehicleDistribution()) {
            saveStatus.textContent = 'Distribua os agentes por viatura antes de salvar.';
            return;
        }
        if (!validateRequiredFields() || !validateIncidentsVTR()) {
            saveStatus.textContent = 'Preencha os campos obrigatórios antes de salvar. Todas as ocorrências devem ter um Posto/VTR selecionado.';
            return;
        }
        if (!validateOptionalFlags()) {
            saveStatus.textContent = 'Preencha Ocorrências e Alterações de Serviço ou marque "Sem alteração."';
            return;
        }
        if (!isEditAllowed()) {
            saveStatus.textContent = 'Edição bloqueada para esta operação.';
            return;
        }
        saveStatus.textContent = 'Salvando no Firebase...';
        await assignNextOperationNumber();
        const formData = getFormData();
        const savedId = await saveOperationToFirebase(formData);
        if (savedId) {
            forcedOperationIdFilter = savedId;
            clearForm();
            setActiveAppView('manager');
            loadOperationsList(true);
        }
    });
    
    // Clear form button
    clearBtn.addEventListener('click', () => {
        confirmationModal.style.display = 'flex';
    });
    
    // Confirmation modal buttons
    confirmClear.addEventListener('click', clearForm);
    cancelClear.addEventListener('click', () => {
        confirmationModal.style.display = 'none';
    });
    
    // Close modal when clicking outside
    confirmationModal.addEventListener('click', (e) => {
        if (e.target === confirmationModal) {
            confirmationModal.style.display = 'none';
        }
    });

    if (confirmResetPassword) {
        confirmResetPassword.addEventListener('click', async () => {
            await sendResetPasswordEmail();
        });
    }
    if (cancelResetPassword) {
        cancelResetPassword.addEventListener('click', () => {
            closeResetPasswordModal();
        });
    }
    if (resetPasswordModal) {
        resetPasswordModal.addEventListener('click', (event) => {
            if (event.target === resetPasswordModal) {
                closeResetPasswordModal();
            }
        });
    }

    if (sessionConfirmContinue) {
        sessionConfirmContinue.addEventListener('click', () => {
            resetSessionCountdown();
        });
    }
    if (sessionConfirmLogout) {
        sessionConfirmLogout.addEventListener('click', () => {
            handleSessionExpired();
        });
    }
    // Não fecha o modal ao clicar fora
    
    // Update total agents count when agents are added/removed
    document.addEventListener('input', () => {
        updateTotalAgentsCount();
    });

    // Revalidate edit window when date changes
    const operationDateInput = document.getElementById('operationDate');
    operationDateInput.addEventListener('change', () => {
        enforceEditWindow();
        assignNextOperationNumber();
    });

    // App navigation
    currentOperationBtn.addEventListener('click', () => setActiveAppView('form'));
    newOperationBtn.addEventListener('click', () => {
        confirmNewOperation();
        setActiveAppView('form');
    });
    viewManagerBtn.addEventListener('click', () => setActiveAppView('manager'));
    if (viewUsersBtn) {
        viewUsersBtn.addEventListener('click', () => {
            if (isAdmin) {
                setActiveAppView('users');
            }
            closeAdminMenu();
        });
    }
    if (viewLogsBtn) {
        viewLogsBtn.addEventListener('click', () => {
            if (isAdmin) {
                setActiveAppView('logs');
                loadLogsList(true);
            }
            closeAdminMenu();
        });
    }

    [logsDateFrom, logsDateTo].forEach(input => {
        if (!input) return;
        input.addEventListener('change', () => {
            applyLogsFilter();
        });
    });

    if (logsLimit) {
        logsLimit.addEventListener('change', () => {
            const value = parseInt(logsLimit.value, 10);
            logsCurrentLimit = Number.isNaN(value) ? 100 : value;
            loadLogsList(true);
        });
    }

    if (logsClearBtn) {
        logsClearBtn.addEventListener('click', () => {
            if (logsDateFrom) logsDateFrom.value = '';
            if (logsDateTo) logsDateTo.value = '';
            if (logsLimit) logsLimit.value = '100';
            logsCurrentLimit = 100;
            loadLogsList(true);
        });
    }

    if (logsFilterAll) {
        logsFilterAll.addEventListener('click', () => {
            logsActionFilter = 'all';
            applyLogsFilter();
        });
    }
    if (logsFilterAccess) {
        logsFilterAccess.addEventListener('click', () => {
            logsActionFilter = 'access';
            applyLogsFilter();
        });
    }
    if (logsFilterReport) {
        logsFilterReport.addEventListener('click', () => {
            logsActionFilter = 'report';
            applyLogsFilter();
        });
    }
    if (logsFilterUpdate) {
        logsFilterUpdate.addEventListener('click', () => {
            logsActionFilter = 'update';
            applyLogsFilter();
        });
    }
    if (logsFilterDelete) {
        logsFilterDelete.addEventListener('click', () => {
            logsActionFilter = 'delete';
            applyLogsFilter();
        });
    }

    // Manager filters
    [filterText, filterDateFrom, filterDateTo].forEach(input => {
        input.addEventListener('input', () => {
            forcedOperationIdFilter = null;
            operationsCurrentPage = 1;
            applyOperationsFilter();
        });
        input.addEventListener('change', () => {
            forcedOperationIdFilter = null;
            operationsCurrentPage = 1;
            applyOperationsFilter();
        });
    });
    refreshOperations.addEventListener('click', () => {
        filterText.value = '';
        filterDateFrom.value = '';
        filterDateTo.value = '';
        forcedOperationIdFilter = null;
        operationsCurrentPage = 1;
        loadOperationsList(true);
    });

    if (usersSearchInput) {
        usersSearchInput.addEventListener('input', () => {
            usersCurrentPage = 1;
            applyUsersFilter();
        });
        usersSearchInput.addEventListener('change', () => {
            usersCurrentPage = 1;
            applyUsersFilter();
        });
    }

    if (profileMenuBtn && profileMenu) {
        profileMenuBtn.addEventListener('click', (event) => {
            event.stopPropagation();
            const isHidden = profileMenu.classList.contains('hidden');
            profileMenu.classList.toggle('hidden', !isHidden);
            profileMenuBtn.setAttribute('aria-expanded', String(isHidden));
        });
    }

    if (adminMenuBtn && adminMenu) {
        adminMenuBtn.addEventListener('click', (event) => {
            event.stopPropagation();
            const isHidden = adminMenu.classList.contains('hidden');
            adminMenu.classList.toggle('hidden', !isHidden);
            adminMenuBtn.setAttribute('aria-expanded', String(isHidden));
        });
        adminMenu.addEventListener('click', (event) => {
            event.stopPropagation();
        });
    }

    document.addEventListener('click', () => {
        closeAdminMenu();
        closeProfileMenu();
    });

    if (saveProfileBtn) {
        saveProfileBtn.addEventListener('click', (event) => {
            event.preventDefault();
            saveCurrentUserProfile();
        });
    }

    if (saveUserEditBtn) {
        saveUserEditBtn.addEventListener('click', (event) => {
            event.preventDefault();
            saveEditedUserProfile();
        });
    }

    if (cancelUserEditBtn) {
        cancelUserEditBtn.addEventListener('click', (event) => {
            event.preventDefault();
            closeUserEditModal();
        });
    }

    if (openUserCreateBtn) {
        openUserCreateBtn.addEventListener('click', (event) => {
            event.preventDefault();
            if (isAdmin) {
                openUserCreateModal();
            }
        });
    }

    if (changePasswordBtn) {
        changePasswordBtn.addEventListener('click', (event) => {
            event.preventDefault();
            openChangePasswordModal();
            closeProfileMenu();
        });
    }

    if (savePasswordBtn) {
        savePasswordBtn.addEventListener('click', (event) => {
            event.preventDefault();
            updateUserPassword();
        });
    }

    if (cancelPasswordBtn) {
        cancelPasswordBtn.addEventListener('click', (event) => {
            event.preventDefault();
            closeChangePasswordModal();
        });
    }

    if (saveUserCreateBtn) {
        saveUserCreateBtn.addEventListener('click', (event) => {
            event.preventDefault();
            createNewUser();
        });
    }

    if (cancelUserCreateBtn) {
        cancelUserCreateBtn.addEventListener('click', (event) => {
            event.preventDefault();
            closeUserCreateModal();
        });
    }

    if (summaryField) {
        summaryField.addEventListener('input', () => {
            updateSummaryTemplateState();
        });
        summaryField.addEventListener('change', () => {
            updateSummaryTemplateState();
        });
    }
    if (insertSummaryTemplateBtn) {
        insertSummaryTemplateBtn.addEventListener('click', () => {
            insertSummaryTemplate();
        });
    }
}

function validateRequiredFields() {
    const requiredIds = [
        'operationNumber',
        'operationName',
        'operationComplexity',
        'eventType',
        'operationDate',
        'startTime',
        'endTime',
        'location',
        'neighborhood',
        'coordinatorName',
        'registration',
        'position',
        'summary'
    ];

    let firstInvalid = null;

    requiredIds.forEach(id => {
        const field = document.getElementById(id);
        if (!field) return;
        const isEmpty = !field.value || field.value.trim() === '';
        if (isEmpty) {
            field.classList.add('field-error');
            if (!firstInvalid) firstInvalid = field;
        } else {
            field.classList.remove('field-error');
        }
    });

    if (firstInvalid) {
        firstInvalid.focus();
        return false;
    }

    return true;
}

function validateVehicleDistribution() {
    const requiredVehicles = vehicleRows
        .map(row => {
            const prefix = (row.prefix || '').trim();
            const post = (row.post || '').trim();
            return [prefix, post].filter(Boolean).join(' - ');
        })
        .filter(label => label);

    if (!requiredVehicles.length) return true;

    const areaMap = new Map();
    areaRows.forEach(row => {
        const post = (row.post || '').trim();
        const agents = parseInt(row.agents, 10) || 0;
        if (!post) return;
        areaMap.set(post, (areaMap.get(post) || 0) + agents);
    });

    const missing = requiredVehicles.filter(label => {
        const totalAgents = areaMap.get(label) || 0;
        return totalAgents <= 0;
    });

    if (missing.length > 0) {
        alert(`Distribua os agentes para todas as viaturas cadastradas: ${missing.join(', ')}`);
        return false;
    }

    return true;
}

function initializeAuthListeners() {
    loginForm.addEventListener('submit', async (event) => {
        event.preventDefault();
        loginError.textContent = '';

        if (!auth) {
            loginError.textContent = 'Firebase não configurado.';
            return;
        }

        try {
            await auth.signInWithEmailAndPassword(loginEmail.value.trim(), loginPassword.value);
            loginForm.reset();
        } catch (error) {
            loginError.textContent = 'Falha no login. Verifique e-mail e senha.';
            console.error(error);
        }
    });

    logoutBtn.addEventListener('click', async () => {
        if (!auth) return;
        await auth.signOut();
        clearSessionDeadline();
        clearSessionCountdown();
        currentOperationId = null;
        operationsCache = [];
        resetUserProfileState();
    });
}

function showLoginView() {
    loginView.classList.remove('hidden');
    appView.classList.add('hidden');
    clearIdleTimer();
    clearSessionCountdown();
    updateSessionTimerUI(SESSION_COUNTDOWN_MS);
    closeSessionConfirmModal();
}

function showAppView() {
    loginView.classList.add('hidden');
    appView.classList.remove('hidden');
    setActiveAppView('form');
    resetIdleTimer();
}

function confirmNewOperation() {
    const confirmed = window.confirm('Deseja iniciar um novo evento? Isso limpará o formulário atual.');
    if (!confirmed) return;
    clearForm();
    currentOperationId = null;
    currentOperationBtn.classList.remove('active');
    newOperationBtn.classList.add('active');
}

function initializeIdleWatcher() {
    const activityEvents = ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll'];
    activityEvents.forEach(eventName => {
        document.addEventListener(eventName, resetIdleTimer, { passive: true });
    });
    resetIdleTimer();
}

function clearIdleTimer() {
    if (idleTimeoutId) {
        clearTimeout(idleTimeoutId);
        idleTimeoutId = null;
    }
}

function resetIdleTimer() {
    clearIdleTimer();
    if (!auth || !auth.currentUser) return;

    idleTimeoutId = setTimeout(() => {
        if (auth && auth.currentUser) {
            auth.signOut();
        }
    }, IDLE_LIMIT_MS);
}

function getSessionDeadline() {
    const raw = localStorage.getItem(SESSION_DEADLINE_KEY);
    const parsed = parseInt(raw, 10);
    return Number.isNaN(parsed) ? null : parsed;
}

function setSessionDeadline(timestamp) {
    localStorage.setItem(SESSION_DEADLINE_KEY, String(timestamp));
}

function clearSessionDeadline() {
    localStorage.removeItem(SESSION_DEADLINE_KEY);
}

function formatSessionCountdown(ms) {
    const safeMs = Math.max(0, ms);
    const totalSeconds = Math.ceil(safeMs / 1000);
    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;
    return `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

function updateSessionTimerUI(remainingMs) {
    if (sessionTimerValue) {
        sessionTimerValue.textContent = formatSessionCountdown(remainingMs);
    }
    if (sessionTimerValueModal) {
        sessionTimerValueModal.textContent = formatSessionCountdown(remainingMs);
    }
    if (sessionTimer) {
        sessionTimer.classList.toggle('is-warning', remainingMs <= SESSION_CONFIRM_AT_MS);
    }
    if (sessionTimerModal) {
        sessionTimerModal.classList.toggle('is-warning', remainingMs <= SESSION_CONFIRM_AT_MS);
    }
}

function openSessionConfirmModal() {
    if (!sessionConfirmModal) return;
    sessionConfirmModal.style.display = 'flex';
}

function closeSessionConfirmModal() {
    if (!sessionConfirmModal) return;
    sessionConfirmModal.style.display = 'none';
}

function clearSessionCountdown() {
    if (sessionIntervalId) {
        clearInterval(sessionIntervalId);
        sessionIntervalId = null;
    }
}

function handleSessionExpired() {
    clearSessionCountdown();
    clearSessionDeadline();
    sessionPrompted = false;
    updateSessionTimerUI(0);
    closeSessionConfirmModal();
    if (auth && auth.currentUser) {
        auth.signOut();
    }
}

function checkSessionCountdown() {
    const deadline = getSessionDeadline();
    if (!deadline) return;
    const remaining = deadline - Date.now();

    updateSessionTimerUI(remaining);

    if (remaining <= 0) {
        handleSessionExpired();
        return;
    }

    if (remaining <= SESSION_CONFIRM_AT_MS && !sessionPrompted) {
        sessionPrompted = true;
        openSessionConfirmModal();
    }
}

function startSessionCountdown() {
    clearSessionCountdown();
    checkSessionCountdown();
    sessionIntervalId = setInterval(checkSessionCountdown, 1000);
}

function resetSessionCountdown() {
    if (!auth || !auth.currentUser) return;
    const deadline = Date.now() + SESSION_COUNTDOWN_MS;
    setSessionDeadline(deadline);
    sessionPrompted = false;
    closeSessionConfirmModal();
    updateSessionTimerUI(deadline - Date.now());
    startSessionCountdown();
}

function closeProfileMenu() {
    if (!profileMenu || !profileMenuBtn) return;
    profileMenu.classList.add('hidden');
    profileMenuBtn.setAttribute('aria-expanded', 'false');
}

function closeAdminMenu() {
    if (!adminMenu || !adminMenuBtn) return;
    adminMenu.classList.add('hidden');
    adminMenuBtn.setAttribute('aria-expanded', 'false');
}

function resetUserProfileState() {
    currentUserProfile = null;
    isAdmin = false;
    usersCache = [];
    editingUserId = null;
    logsCache = [];
    closeAdminMenu();
    closeProfileMenu();
    closeUserCreateModal();
    updateProfileSummary(null);
    if (adminMenuContainer) {
        adminMenuContainer.classList.add('hidden');
    }
    if (openUserCreateBtn) {
        openUserCreateBtn.classList.add('hidden');
    }
    if (viewUsersBtn) {
        viewUsersBtn.classList.add('hidden');
        viewUsersBtn.classList.remove('active');
    }
    if (viewLogsBtn) {
        viewLogsBtn.classList.add('hidden');
        viewLogsBtn.classList.remove('active');
    }
    if (usersView) {
        usersView.classList.add('hidden');
    }
    if (logsView) {
        logsView.classList.add('hidden');
    }
    if (profileModal) {
        profileModal.style.display = 'none';
    }
    if (userEditModal) {
        userEditModal.style.display = 'none';
    }
    if (userCreateModal) {
        userCreateModal.style.display = 'none';
    }
    if (changePasswordModal) {
        changePasswordModal.style.display = 'none';
    }
}

async function initializeUserProfileFlow(user) {
    if (!db || !user) return;
    await ensureUserProfile(user);
    await registerAccessLog(user);
}

function getUsersCollection() {
    return db.collection('users');
}

function getAccessLogsCollection() {
    return db.collection('accessLogs');
}

function isProfileComplete(profile) {
    return Boolean(profile && profile.classe && profile.nomeGuerra && profile.matricula);
}

function updateAdminVisibility() {
    if (adminMenuContainer) {
        adminMenuContainer.classList.toggle('hidden', !isAdmin);
    }
    if (openUserCreateBtn) {
        openUserCreateBtn.classList.toggle('hidden', !isAdmin);
    }
    if (viewUsersBtn) {
        viewUsersBtn.classList.toggle('hidden', !isAdmin);
    }
    if (viewLogsBtn) {
        viewLogsBtn.classList.toggle('hidden', !isAdmin);
    }
    if (!isAdmin && usersView && !usersView.classList.contains('hidden')) {
        setActiveAppView('form');
    }
    if (!isAdmin && logsView && !logsView.classList.contains('hidden')) {
        setActiveAppView('form');
    }
}

function updateProfileSummary(profile) {
    if (!profileSummaryClass) return;
    const data = profile || {};
    profileSummaryClass.textContent = data.classe || '-';
    profileSummaryWarName.textContent = data.nomeGuerra || '-';
    profileSummaryRegistration.textContent = data.matricula || '-';
    const tipoValue = data.tipo || 'usuario';
    profileSummaryType.textContent = tipoValue === 'usuario_restrito' ? 'Usuário restrito' : tipoValue;
}

function openProfileModal() {
    if (!profileModal) return;
    profileError.textContent = '';
    profileClass.value = currentUserProfile?.classe || '';
    profileWarName.value = currentUserProfile?.nomeGuerra || '';
    profileRegistration.value = currentUserProfile?.matricula || '';
    profileType.value = currentUserProfile?.tipo || 'usuario';
    profileType.disabled = true;
    profileModal.style.display = 'flex';
}

function closeProfileModal() {
    if (!profileModal) return;
    profileModal.style.display = 'none';
}

function openUserEditModal(user) {
    if (!userEditModal || !user) return;
    editingUserId = user.id;
    userEditError.textContent = '';
    userEditClass.value = user.classe || '';
    userEditWarName.value = user.nomeGuerra || '';
    userEditRegistration.value = user.matricula || '';
    userEditType.value = user.tipo || 'usuario';
    userEditModal.style.display = 'flex';
}

function closeUserEditModal() {
    if (!userEditModal) return;
    editingUserId = null;
    userEditModal.style.display = 'none';
}

function openUserCreateModal() {
    if (!userCreateModal || !isAdmin) return;
    if (userCreateError) userCreateError.textContent = '';
    if (userCreateForm) userCreateForm.reset();
    if (userCreateType) userCreateType.value = 'usuario';
    userCreateModal.style.display = 'flex';
}

function closeUserCreateModal() {
    if (!userCreateModal) return;
    userCreateModal.style.display = 'none';
}

function openChangePasswordModal() {
    if (!changePasswordModal) return;
    if (changePasswordError) changePasswordError.textContent = '';
    if (changePasswordForm) changePasswordForm.reset();
    changePasswordModal.style.display = 'flex';
    currentPassword?.focus();
}

function closeChangePasswordModal() {
    if (!changePasswordModal) return;
    changePasswordModal.style.display = 'none';
}

async function updateUserPassword() {
    if (!auth || !auth.currentUser) return;
    if (!currentPassword || !newPassword || !confirmPassword) return;

    const currentValue = currentPassword.value.trim();
    const newValue = newPassword.value.trim();
    const confirmValue = confirmPassword.value.trim();

    if (!currentValue || !newValue || !confirmValue) {
        if (changePasswordError) changePasswordError.textContent = 'Preencha todos os campos.';
        return;
    }

    if (newValue.length < 6) {
        if (changePasswordError) changePasswordError.textContent = 'A nova senha deve ter pelo menos 6 caracteres.';
        return;
    }

    if (newValue !== confirmValue) {
        if (changePasswordError) changePasswordError.textContent = 'A confirmação de senha não confere.';
        return;
    }

    try {
        const user = auth.currentUser;
        const credential = firebase.auth.EmailAuthProvider.credential(user.email, currentValue);
        await user.reauthenticateWithCredential(credential);
        await user.updatePassword(newValue);
        if (changePasswordError) changePasswordError.textContent = 'Senha alterada com sucesso.';
        setTimeout(closeChangePasswordModal, 800);
    } catch (error) {
        console.error('Erro ao alterar senha:', error);
        if (changePasswordError) changePasswordError.textContent = 'Falha ao alterar senha. Verifique a senha atual.';
    }
}

function validateProfileFields({ classe, nomeGuerra, matricula }) {
    return Boolean(classe && nomeGuerra && matricula);
}

function validateUserCreateFields({ email, password, classe, nomeGuerra, matricula }) {
    return Boolean(email && password && classe && nomeGuerra && matricula);
}

async function ensureUserProfile(user) {
    try {
        const docRef = getUsersCollection().doc(user.uid);
        const docSnap = await docRef.get();
        let profileData = docSnap.exists ? docSnap.data() : null;

        await docRef.set(
            {
                uid: user.uid,
                email: user.email || null,
                updatedAt: firebase.firestore.FieldValue.serverTimestamp()
            },
            { merge: true }
        );

        if (!profileData) {
            profileData = {
                uid: user.uid,
                email: user.email || null,
                classe: '',
                nomeGuerra: '',
                matricula: '',
                tipo: 'usuario',
                createdAt: firebase.firestore.FieldValue.serverTimestamp(),
                updatedAt: firebase.firestore.FieldValue.serverTimestamp()
            };
            await docRef.set(profileData, { merge: true });
        } else if (!profileData.tipo) {
            profileData.tipo = 'usuario';
            await docRef.set({ tipo: 'usuario', updatedAt: firebase.firestore.FieldValue.serverTimestamp() }, { merge: true });
        }

        const refreshedSnap = await docRef.get();
        if (refreshedSnap.exists) {
            profileData = refreshedSnap.data();
        }

        currentUserProfile = profileData;
        isAdmin = currentUserProfile.tipo === 'administrador';
        isRestricted = currentUserProfile.tipo === 'usuario_restrito';
        updateAdminVisibility();
        updateProfileSummary(currentUserProfile);
    } catch (error) {
        console.error('Erro ao carregar perfil do usuário:', error);
        if (!currentUserProfile) {
            currentUserProfile = {
                uid: user.uid,
                email: user.email || null,
                classe: '',
                nomeGuerra: '',
                matricula: '',
                tipo: 'usuario'
            };
        }
    }
}

async function saveCurrentUserProfile() {
    if (!auth || !auth.currentUser || !db) return;

    const classe = profileClass.value.trim();
    const nomeGuerra = profileWarName.value.trim();
    const matricula = profileRegistration.value.trim();
    const tipoValue = currentUserProfile?.tipo || 'usuario';

    if (!validateProfileFields({ classe, nomeGuerra, matricula })) {
        profileError.textContent = 'Preencha todos os campos obrigatórios.';
        return;
    }

    try {
        const payload = {
            classe,
            nomeGuerra,
            matricula,
            tipo: tipoValue,
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };
        await getUsersCollection().doc(auth.currentUser.uid).set(payload, { merge: true });

        currentUserProfile = {
            ...(currentUserProfile || {}),
            ...payload,
            uid: auth.currentUser.uid,
            email: auth.currentUser.email || currentUserProfile?.email || null
        };
        isAdmin = currentUserProfile.tipo === 'administrador';
        updateAdminVisibility();
        updateProfileSummary(currentUserProfile);
        closeProfileModal();
    } catch (error) {
        console.error('Erro ao salvar perfil do usuário:', error);
        profileError.textContent = 'Falha ao salvar o perfil. Tente novamente.';
    }
}

async function createNewUser() {
    if (!auth || !db || !isAdmin) return;
    if (!userCreateEmail || !userCreatePassword || !userCreateClass || !userCreateWarName || !userCreateRegistration || !userCreateType) return;

    const email = userCreateEmail.value.trim();
    const password = userCreatePassword.value.trim();
    const classe = userCreateClass.value.trim();
    const nomeGuerra = userCreateWarName.value.trim();
    const matricula = userCreateRegistration.value.trim();
    const tipo = userCreateType.value || 'usuario';

    if (!validateUserCreateFields({ email, password, classe, nomeGuerra, matricula })) {
        if (userCreateError) userCreateError.textContent = 'Preencha todos os campos obrigatórios.';
        return;
    }

    try {
        const existingSecondary = firebase.apps.find(app => app.name === 'adminCreate');
        const secondaryApp = existingSecondary || firebase.initializeApp(firebaseConfig, 'adminCreate');
        const secondaryAuth = secondaryApp.auth();

        const userCredential = await secondaryAuth.createUserWithEmailAndPassword(email, password);
        const createdUser = userCredential.user;

        const payload = {
            classe,
            nomeGuerra,
            matricula,
            tipo,
            email,
            createdAt: firebase.firestore.FieldValue.serverTimestamp(),
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };

        await getUsersCollection().doc(createdUser.uid).set(payload, { merge: true });
        await secondaryAuth.signOut();

        closeUserCreateModal();
        if (userCreateForm) userCreateForm.reset();
        loadUsersList(true);
    } catch (error) {
        console.error('Erro ao criar usuário:', error);
        if (userCreateError) {
            userCreateError.textContent = 'Falha ao criar usuário. Verifique os dados e tente novamente.';
        }
    }
}

async function loadUsersList(force = false) {
    if (!db || !auth || !auth.currentUser || !isAdmin || isLoadingUsers) return;
    if (usersCache.length > 0 && !force) {
        applyUsersFilter();
        return;
    }

    isLoadingUsers = true;
    if (usersTableBody) {
        usersTableBody.innerHTML = '<tr><td colspan="6">Carregando usuários...</td></tr>';
    }

    try {
        const snapshot = await getUsersCollection().orderBy('nomeGuerra').get();
        usersCache = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        usersCurrentPage = 1;
        applyUsersFilter();
    } catch (error) {
        console.error('Erro ao carregar usuários:', error);
        if (usersTableBody) {
            usersTableBody.innerHTML = '<tr><td colspan="6">Falha ao carregar usuários.</td></tr>';
        }
    } finally {
        isLoadingUsers = false;
    }
}

function applyUsersFilter() {
    const query = (usersSearchInput?.value || '').trim().toLowerCase();
    const filtered = usersCache.filter(user => {
        const searchText = [
            user.classe || '',
            user.nomeGuerra || '',
            user.email || '',
            user.matricula || '',
            user.tipo || ''
        ].join(' ').toLowerCase();
        if (query && !searchText.includes(query)) return false;
        return true;
    });

    renderUsersTable(filtered);
}

function renderUsersTable(users) {
    if (!usersTableBody) return;
    if (!users.length) {
        usersTableBody.innerHTML = '<tr><td colspan="6">Nenhum usuário encontrado.</td></tr>';
        if (usersPagination) {
            usersPagination.innerHTML = '';
        }
        return;
    }

    const totalPages = Math.ceil(users.length / USERS_PAGE_SIZE);
    if (usersCurrentPage > totalPages) {
        usersCurrentPage = totalPages;
    }
    const startIndex = (usersCurrentPage - 1) * USERS_PAGE_SIZE;
    const pageItems = users.slice(startIndex, startIndex + USERS_PAGE_SIZE);

    usersTableBody.innerHTML = '';
    pageItems.forEach(user => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${user.classe || '-'}</td>
            <td>${user.nomeGuerra || '-'}</td>
            <td>${user.email || '-'}</td>
            <td>${user.matricula || '-'}</td>
            <td>${user.tipo || 'usuario'}</td>
            <td>
                <button class="btn-small" data-user-id="${user.id}">
                    <i class="fas fa-pen"></i> Editar
                </button>
                <button class="btn-small" data-user-reset-id="${user.id}">
                    <i class="fas fa-envelope"></i> Redefinir senha
                </button>
                ${isAdmin ? `
                <button class="btn-small btn-danger" data-user-delete-id="${user.id}">
                    <i class="fas fa-trash"></i> Excluir
                </button>
                ` : ''}
            </td>
        `;
        usersTableBody.appendChild(tr);
    });

    usersTableBody.querySelectorAll('button[data-user-id]').forEach(button => {
        button.addEventListener('click', () => {
            if (!isAdmin) return;
            const userId = button.getAttribute('data-user-id');
            const user = usersCache.find(item => item.id === userId);
            if (user) {
                openUserEditModal(user);
            }
        });
    });

    usersTableBody.querySelectorAll('button[data-user-reset-id]').forEach(button => {
        button.addEventListener('click', () => {
            if (!isAdmin) return;
            const userId = button.getAttribute('data-user-reset-id');
            const user = usersCache.find(item => item.id === userId);
            if (user) {
                openResetPasswordModal(user);
            }
        });
    });

    usersTableBody.querySelectorAll('button[data-user-delete-id]').forEach(button => {
        button.addEventListener('click', async () => {
            if (!isAdmin) return;
            const userId = button.getAttribute('data-user-delete-id');
            const user = usersCache.find(item => item.id === userId);
            if (!user) return;

            const justification = prompt('Informe a justificativa para excluir este usuário:');
            if (!justification || !justification.trim()) return;

            try {
                const targetName = user.nomeGuerra || user.email || user.id;
                const fullJustification = `${targetName}: ${justification.trim()}`;
                await getUsersCollection().doc(userId).delete();
                usersCache = usersCache.filter(item => item.id !== userId);
                await registerAdminActionLog({
                    action: 'delete-user',
                    justification: fullJustification,
                    operationId: null,
                    operationName: null,
                    targetUserId: userId,
                    targetUserEmail: user.email || null
                });
                applyUsersFilter();
            } catch (error) {
                console.error('Erro ao excluir usuário:', error);
                alert('Falha ao excluir o usuário.');
            }
        });
    });

    renderUsersPagination(totalPages);
}

function renderUsersPagination(totalPages) {
    if (!usersPagination) return;
    usersPagination.innerHTML = '';
    if (totalPages <= 1) return;

    const prevBtn = document.createElement('button');
    prevBtn.className = 'pagination-btn';
    prevBtn.textContent = 'Anterior';
    prevBtn.disabled = usersCurrentPage === 1;
    prevBtn.addEventListener('click', () => {
        usersCurrentPage -= 1;
        applyUsersFilter();
    });
    usersPagination.appendChild(prevBtn);

    for (let i = 1; i <= totalPages; i += 1) {
        const pageBtn = document.createElement('button');
        pageBtn.className = 'pagination-btn';
        if (i === usersCurrentPage) pageBtn.classList.add('active');
        pageBtn.textContent = String(i);
        pageBtn.addEventListener('click', () => {
            usersCurrentPage = i;
            applyUsersFilter();
        });
        usersPagination.appendChild(pageBtn);
    }

    const nextBtn = document.createElement('button');
    nextBtn.className = 'pagination-btn';
    nextBtn.textContent = 'Próximo';
    nextBtn.disabled = usersCurrentPage === totalPages;
    nextBtn.addEventListener('click', () => {
        usersCurrentPage += 1;
        applyUsersFilter();
    });
    usersPagination.appendChild(nextBtn);
}

function updateSummaryTemplateState() {
    if (!summaryField || !insertSummaryTemplateBtn) return;
    const isBlank = summaryField.value.trim() === '';
    insertSummaryTemplateBtn.disabled = !isBlank;
    if (summaryTemplateWarning) {
        summaryTemplateWarning.classList.toggle('hidden', isBlank);
    }
    if (isBlank && summaryTemplateInfo) {
        summaryTemplateInfo.classList.add('hidden');
    }
}

function insertSummaryTemplate() {
    if (!summaryField) return;
    if (summaryField.value.trim() !== '') return;
    summaryField.value = SUMMARY_TEMPLATE_TEXT;
    if (summaryTemplateInfo) {
        summaryTemplateInfo.classList.remove('hidden');
    }
    updateSummaryTemplateState();
    autoExpandTextarea(summaryField);
    triggerSaveDebounced();
}

function openResetPasswordModal(user) {
    if (!resetPasswordModal) return;
    resetPasswordTarget = user || null;
    if (resetPasswordError) resetPasswordError.textContent = '';
    const name = user?.nomeGuerra || user?.email || 'este usuário';
    const email = user?.email || 'sem e-mail cadastrado';
    if (resetPasswordMessage) {
        resetPasswordMessage.textContent = `Deseja enviar um e-mail de redefinição de senha para ${name} (${email})? Oriente o usuário a verificar também a caixa de spam.`;
    }
    resetPasswordModal.style.display = 'flex';
}

function closeResetPasswordModal() {
    if (!resetPasswordModal) return;
    resetPasswordModal.style.display = 'none';
    resetPasswordTarget = null;
    if (resetPasswordError) resetPasswordError.textContent = '';
}

async function sendResetPasswordEmail() {
    if (!auth || !isAdmin) return;
    if (!resetPasswordTarget) return;
    const email = (resetPasswordTarget.email || '').trim();
    if (!email) {
        if (resetPasswordError) resetPasswordError.textContent = 'Este usuário não possui e-mail cadastrado.';
        return;
    }

    try {
        await auth.sendPasswordResetEmail(email);
        closeResetPasswordModal();
        alert('E-mail de redefinição enviado com sucesso.');
    } catch (error) {
        console.error('Erro ao enviar e-mail de redefinição:', error);
        if (resetPasswordError) {
            resetPasswordError.textContent = 'Falha ao enviar o e-mail. Tente novamente.';
        }
    }
}

async function saveEditedUserProfile() {
    if (!isAdmin || !editingUserId) return;

    const classe = userEditClass.value.trim();
    const nomeGuerra = userEditWarName.value.trim();
    const matricula = userEditRegistration.value.trim();
    const tipoValue = userEditType.value;

    if (!validateProfileFields({ classe, nomeGuerra, matricula })) {
        userEditError.textContent = 'Preencha todos os campos obrigatórios.';
        return;
    }

    try {
        const payload = {
            classe,
            nomeGuerra,
            matricula,
            tipo: tipoValue,
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };
        await getUsersCollection().doc(editingUserId).update(payload);

        if (auth?.currentUser?.uid === editingUserId) {
            currentUserProfile = {
                ...(currentUserProfile || {}),
                ...payload
            };
            isAdmin = currentUserProfile.tipo === 'administrador';
            updateAdminVisibility();
        }

        closeUserEditModal();
        loadUsersList(true);
    } catch (error) {
        console.error('Erro ao atualizar usuário:', error);
        userEditError.textContent = 'Falha ao salvar. Tente novamente.';
    }
}

function setActiveAppView(view) {
    const isForm = view === 'form';
    const isManager = view === 'manager';
    const isUsers = view === 'users';
    const isLogs = view === 'logs';

    operationFormView.classList.toggle('hidden', !isForm);
    managerView.classList.toggle('hidden', !isManager);
    if (usersView) {
        usersView.classList.toggle('hidden', !isUsers);
    }
    if (logsView) {
        logsView.classList.toggle('hidden', !isLogs);
    }

    currentOperationBtn.classList.toggle('active', isForm);
    viewManagerBtn.classList.toggle('active', isManager);
    if (viewUsersBtn) {
        viewUsersBtn.classList.toggle('active', isUsers);
    }
    if (viewLogsBtn) {
        viewLogsBtn.classList.toggle('active', isLogs);
    }

    if (isForm) {
        newOperationBtn.classList.remove('active');
        requestAnimationFrame(() => {
            requestAnimationFrame(expandAllTextareas);
        });
        enforceEditWindow();
    }

    if (isManager) {
        loadOperationsList();
    }

    if (isUsers && isAdmin) {
        loadUsersList(true);
    }
}

// Function to auto-expand textarea
function autoExpandTextarea(textarea) {
    // Reset height to auto to get the correct scrollHeight
    textarea.style.height = 'auto';
    // Set the height to the scrollHeight
    textarea.style.height = Math.max(textarea.scrollHeight, 80) + 'px';
}

function expandAllTextareas() {
    document.querySelectorAll('#operationFormView textarea').forEach(textarea => {
        autoExpandTextarea(textarea);
    });
}

function isEditAllowed() {
    if (!auth || !auth.currentUser) return false;
    if (!currentOperationId) return true;
    const ownerId = currentOperationMeta?.ownerId || null;
    if (!ownerId) return false;
    return auth.currentUser.uid === ownerId;
}

function enforceEditWindow() {
    const allowed = isEditAllowed();

    editLockBanner.classList.toggle('hidden', allowed);

    document.querySelectorAll('#operationFormView input, #operationFormView textarea, #operationFormView select').forEach(input => {
        if (input.id === 'operationDate') {
            input.disabled = !allowed;
        } else {
            input.disabled = !allowed;
        }
    });

    addAreaBtn.disabled = !allowed;
    addVehicleBtn.disabled = !allowed;
    addServiceChangeBtn.disabled = !allowed;
    if (addIncidentBtn) {
        addIncidentBtn.disabled = !allowed;
    }
    clearBtn.disabled = !allowed;
    saveDbBtn.disabled = !allowed;
    if (addEventTypeBtn) {
        addEventTypeBtn.disabled = !allowed;
    }
}

function triggerSave() {
    // Show saving indicator
    if (!isSaving) {
        saveStatus.textContent = 'Salvando...';
        saveIndicator.classList.add('saving');
        isSaving = true;
    }
    
    // Clear previous timeout
    if (saveTimeout) {
        clearTimeout(saveTimeout);
    }
    
    // Set new timeout for saving
    saveTimeout = setTimeout(saveFormData, 500);
}

// Update total agents count based on actual agents added
function updateTotalAgentsCount() {
    const total = areaRows.reduce((sum, row) => sum + (parseInt(row.agents, 10) || 0), 0);
    document.getElementById('totalAgents').textContent = total;
}

function getPostOptionsList(currentRowId) {
    const options = new Set();
    vehicleRows.forEach(row => {
        const prefix = (row.prefix || '').trim();
        const post = (row.post || '').trim();
        const label = [prefix, post].filter(Boolean).join(' - ');
        if (label) {
            options.add(label);
        }
    });

    const usedPosts = new Set();
    areaRows.forEach(row => {
        if (currentRowId && row.id === currentRowId) return;
        const value = (row.post || row.vehicle || '').trim();
        if (value) {
            usedPosts.add(value);
        }
    });

    const filtered = Array.from(options).filter(option => !usedPosts.has(option));
    const result = new Set(filtered);
    result.add('Agentes a Pé');
    return Array.from(result).sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

function buildPostOptions(selectedValue, currentRowId) {
    const options = ['<option value="">Selecione</option>'];
    getPostOptionsList(currentRowId).forEach(option => {
        const isSelected = option === selectedValue ? ' selected' : '';
        options.push(`<option value="${option}"${isSelected}>${option}</option>`);
    });
    return options.join('');
}

// AREA DISTRIBUTION FUNCTIONS
function addAreaRow() {
    const rowId = Date.now();
    const newRow = {
        id: rowId,
        area: '',
        agents: '',
        post: ''
    };
    
    areaRows.push(newRow);
    updateAreaTable();
    updateAreaCards();
    updateTotalAgentsCount();
    triggerSaveDebounced();
}

function removeAreaRow(rowId) {
    if (areaRows.length <= 1) {
        alert('Deve haver pelo menos uma área cadastrada.');
        return;
    }
    
    areaRows = areaRows.filter(row => row.id !== rowId);
    updateAreaTable();
    updateAreaCards();
    updateTotalAgentsCount();
    triggerSaveDebounced();
}

function updateAreaTable() {
    areaTableBody.innerHTML = '';
    
    areaRows.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <select class="area-input" data-id="${row.id}" data-field="post">
                    ${buildPostOptions(row.post || row.vehicle || '', row.id)}
                </select>
            </td>
            <td><input type="number" class="area-input" data-id="${row.id}" data-field="agents" value="${row.agents}" min="0" placeholder="0"></td>
            <td><input type="text" class="area-input" data-id="${row.id}" data-field="area" value="${row.area}" placeholder="Ex: Centro Comercial"></td>
            <td><button class="btn-small btn-danger remove-area-row" data-id="${row.id}"><i class="fas fa-trash-alt"></i> Remover</button></td>
        `;
        areaTableBody.appendChild(tr);
    });
    
    document.querySelectorAll('.area-input').forEach(input => {
        const handler = (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            const field = e.target.getAttribute('data-field');
            const value = e.target.value;

            const rowIndex = areaRows.findIndex(row => row.id === rowId);
            if (rowIndex !== -1) {
                areaRows[rowIndex][field] = value;
                if (field === 'agents') {
                    updateTotalAgentsCount();
                }
                triggerSaveDebounced();
            }
        };
        if (input.tagName === 'SELECT') {
            input.addEventListener('change', handler);
        } else {
            input.addEventListener('input', handler);
        }
    });
    
    document.querySelectorAll('.remove-area-row').forEach(button => {
        button.addEventListener('click', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            removeAreaRow(rowId);
        });
    });
}

function updateAreaCards() {
    areaCardsContainer.innerHTML = '';
    
    areaRows.forEach(row => {
        const card = document.createElement('div');
        card.className = 'area-card';
        card.innerHTML = `
            <div class="card-row">
                <span class="card-label">Posto:</span>
                <span>
                    <select class="card-input" data-id="${row.id}" data-field="post">
                        ${buildPostOptions(row.post || row.vehicle || '', row.id)}
                    </select>
                </span>
            </div>
            <div class="card-row">
                <span class="card-label">Qtd. Agentes:</span>
                <span><input type="number" class="card-input" data-id="${row.id}" data-field="agents" value="${row.agents}" min="0" placeholder="0"></span>
            </div>
            <div class="card-row">
                <span class="card-label">Área/Local:</span>
                <span><input type="text" class="card-input" data-id="${row.id}" data-field="area" value="${row.area}" placeholder="Ex: Centro Comercial"></span>
            </div>
            <div class="card-actions">
                <button class="btn-small btn-danger remove-area-row" data-id="${row.id}"><i class="fas fa-trash-alt"></i> Remover</button>
            </div>
        `;
        areaCardsContainer.appendChild(card);
    });
    
    document.querySelectorAll('.cards-container .card-input').forEach(input => {
        const handler = (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            const field = e.target.getAttribute('data-field');
            const value = e.target.value;

            const rowIndex = areaRows.findIndex(row => row.id === rowId);
            if (rowIndex !== -1) {
                areaRows[rowIndex][field] = value;
                if (field === 'agents') {
                    updateTotalAgentsCount();
                }
                triggerSaveDebounced();
            }
        };
        if (input.tagName === 'SELECT') {
            input.addEventListener('change', handler);
        } else {
            input.addEventListener('input', handler);
        }
    });
    
    document.querySelectorAll('.cards-container .remove-area-row').forEach(button => {
        button.addEventListener('click', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            removeAreaRow(rowId);
        });
    });
}

// VEHICLES FUNCTIONS
function addVehicleRow() {
    const rowId = Date.now();
    const newRow = {
        id: rowId,
        prefix: '',
        post: '',
        type: ''
    };
    
    vehicleRows.push(newRow);
    updateVehiclesTable();
    updateVehiclesCards();
    triggerSaveDebounced();
}

function removeVehicleRow(rowId) {
    if (vehicleRows.length <= 1) {
        alert('Deve haver pelo menos uma viatura cadastrada.');
        return;
    }
    
    vehicleRows = vehicleRows.filter(row => row.id !== rowId);
    updateVehiclesTable();
    updateVehiclesCards();
    triggerSaveDebounced();
}

function updateVehiclesTable() {
    vehiclesTableBody.innerHTML = '';
    
    vehicleRows.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" class="vehicle-input" data-id="${row.id}" data-field="prefix" value="${row.prefix}" placeholder="Ex: VTR-001"></td>
            <td><input type="text" class="vehicle-input" data-id="${row.id}" data-field="post" value="${row.post}" placeholder="Ex: RONDAC"></td>
            <td>
                <select class="vehicle-input" data-id="${row.id}" data-field="type" value="${row.type}">
                    <option value="">Selecione</option>
                    <option value="VTR Convencional">VTR Convencional</option>
                    <option value="VTR Pickup">VTR Pickup</option>
                    <option value="Motocicleta">Motocicleta</option>
                    <option value="Bicicleta">Bicicleta</option>
                    <option value="Ônibus">Ônibus</option>
                    <option value="Outra">Outra</option>
                </select>
            </td>
            <td><button class="btn-small btn-danger remove-vehicle-row" data-id="${row.id}"><i class="fas fa-trash-alt"></i> Remover</button></td>
        `;
        vehiclesTableBody.appendChild(tr);
        
        // Set select value
        const select = tr.querySelector('select');
        if (select) select.value = row.type;
    });
    
    document.querySelectorAll('.vehicle-input').forEach(input => {
        if (input.tagName === 'INPUT') {
            input.addEventListener('input', (e) => {
                const rowId = parseInt(e.target.getAttribute('data-id'));
                const field = e.target.getAttribute('data-field');
                const value = e.target.value;
                
                const rowIndex = vehicleRows.findIndex(row => row.id === rowId);
                if (rowIndex !== -1) {
                    vehicleRows[rowIndex][field] = value;
                    triggerSaveDebounced();
                }
            });
        } else if (input.tagName === 'SELECT') {
            input.addEventListener('change', (e) => {
                const rowId = parseInt(e.target.getAttribute('data-id'));
                const field = e.target.getAttribute('data-field');
                const value = e.target.value;
                
                const rowIndex = vehicleRows.findIndex(row => row.id === rowId);
                if (rowIndex !== -1) {
                    vehicleRows[rowIndex][field] = value;
                    triggerSaveDebounced();
                }
            });
        }
    });
    
    document.querySelectorAll('.remove-vehicle-row').forEach(button => {
        button.addEventListener('click', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            removeVehicleRow(rowId);
        });
    });

    updateAreaTable();
    updateAreaCards();
}

function updateVehiclesCards() {
    vehiclesCardsContainer.innerHTML = '';
    
    vehicleRows.forEach(row => {
        const card = document.createElement('div');
        card.className = 'vehicle-card';
        card.innerHTML = `
            <div class="card-row">
                <span class="card-label">Prefixo:</span>
                <span><input type="text" class="card-input" data-id="${row.id}" data-field="prefix" value="${row.prefix}" placeholder="Ex: VTR-001"></span>
            </div>
            <div class="card-row">
                <span class="card-label">Posto:</span>
                <span><input type="text" class="card-input" data-id="${row.id}" data-field="post" value="${row.post}" placeholder="Ex: RONDAC"></span>
            </div>
            <div class="card-row">
                <span class="card-label">Tipo:</span>
                <span>
                    <select class="card-input" data-id="${row.id}" data-field="type" style="width: 100%;">
                        <option value="">Selecione</option>
                        <option value="VTR Blindada">VTR Blindada</option>
                        <option value="VTR Convencional">VTR Convencional</option>
                        <option value="Motocicleta">Motocicleta</option>
                        <option value="Viatura de Comando">Viatura de Comando</option>
                        <option value="Outra">Outra</option>
                    </select>
                </span>
            </div>
            <div class="card-actions">
                <button class="btn-small btn-danger remove-vehicle-row" data-id="${row.id}"><i class="fas fa-trash-alt"></i> Remover</button>
            </div>
        `;
        vehiclesCardsContainer.appendChild(card);
        
        // Set select value
        const select = card.querySelector('select');
        if (select) select.value = row.type;
    });
    
    document.querySelectorAll('#vehiclesCardsContainer .card-input').forEach(input => {
        if (input.tagName === 'INPUT') {
            input.addEventListener('input', (e) => {
                const rowId = parseInt(e.target.getAttribute('data-id'));
                const field = e.target.getAttribute('data-field');
                const value = e.target.value;
                
                const rowIndex = vehicleRows.findIndex(row => row.id === rowId);
                if (rowIndex !== -1) {
                    vehicleRows[rowIndex][field] = value;
                    triggerSaveDebounced();
                }
            });
        } else if (input.tagName === 'SELECT') {
            input.addEventListener('change', (e) => {
                const rowId = parseInt(e.target.getAttribute('data-id'));
                const field = e.target.getAttribute('data-field');
                const value = e.target.value;
                
                const rowIndex = vehicleRows.findIndex(row => row.id === rowId);
                if (rowIndex !== -1) {
                    vehicleRows[rowIndex][field] = value;
                    triggerSaveDebounced();
                }
            });
        }
    });
    
    document.querySelectorAll('#vehiclesCardsContainer .remove-vehicle-row').forEach(button => {
        button.addEventListener('click', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            removeVehicleRow(rowId);
        });
    });

    updateAreaTable();
    updateAreaCards();
}

// SERVICE CHANGES FUNCTIONS
function addServiceChangeRow() {
    const rowId = Date.now();
    const itemCount = serviceChangeRows.length + 1;
    
    const newRow = {
        id: rowId,
        description: '',
        index: itemCount
    };
    
    serviceChangeRows.push(newRow);
    updateServiceChangesContainer();
    triggerSaveDebounced();
}

function removeServiceChangeRow(rowId) {
    serviceChangeRows = serviceChangeRows.filter(row => row.id !== rowId);
    // Reindex items
    serviceChangeRows.forEach((row, index) => {
        row.index = index + 1;
    });
    updateServiceChangesContainer();
    triggerSaveDebounced();
}

function updateServiceChangesContainer() {
    const container = document.getElementById('serviceChangesContainer');
    container.innerHTML = '';
    
    serviceChangeRows.forEach(row => {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'service-change-item';
        itemDiv.innerHTML = `
            <div class="service-change-header">
                <h4>Alteração ${row.index}</h4>
                <button class="btn-small btn-danger remove-service-change-row" data-id="${row.id}">
                    <i class="fas fa-trash-alt"></i> Remover
                </button>
            </div>
            <textarea class="service-change-textarea" 
                      data-id="${row.id}" 
                      placeholder="Descreva a alteração de serviço (substituições, mudanças de escala, etc.)">${row.description}</textarea>
        `;
        container.appendChild(itemDiv);
    });
    
    // Add event listeners to textareas
    document.querySelectorAll('.service-change-textarea').forEach(textarea => {
        // Initialize height
        autoExpandTextarea(textarea);
        
        textarea.addEventListener('input', (e) => {
            // Auto-expand on input
            autoExpandTextarea(e.target);
            
            const rowId = parseInt(e.target.getAttribute('data-id'));
            const value = e.target.value;
            
            const rowIndex = serviceChangeRows.findIndex(row => row.id === rowId);
            if (rowIndex !== -1) {
                serviceChangeRows[rowIndex].description = value;
                triggerSaveDebounced();
            }
        });
    });
    
    // Add event listeners to remove buttons
    document.querySelectorAll('.remove-service-change-row').forEach(button => {
        button.addEventListener('click', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            removeServiceChangeRow(rowId);
        });
    });

    if (serviceChangesNoChangeRow) {
        serviceChangesNoChangeRow.classList.toggle('hidden', serviceChangeRows.length > 0);
        if (serviceChangeRows.length > 0 && serviceChangesNoChange) {
            serviceChangesNoChange.checked = false;
        }
    }
}

// INCIDENTS FUNCTIONS
function addIncidentRow() {
    const rowId = Date.now();
    const newRow = {
        id: rowId,
        description: '',
        vtr: '' // Novo campo para associar VTR
    };
    incidentsRows.push(newRow);
    updateIncidentsContainer();
    triggerSaveDebounced();
}

function removeIncidentRow(rowId) {
    incidentsRows = incidentsRows.filter(row => row.id !== rowId);
    updateIncidentsContainer();
    triggerSave();
}

function updateIncidentsContainer() {
    if (!incidentsContainer) return;
    incidentsContainer.innerHTML = '';

    // Opções de VTRs cadastradas
    const vtrOptions = vehicleRows.map(row => {
        const label = row.prefix ? `${row.prefix} - ${row.post || ''}` : (row.post || '');
        return `<option value="${row.prefix || row.post || ''}">${label}</option>`;
    });

    incidentsRows.forEach((row, index) => {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'service-change-item';
        itemDiv.innerHTML = `
            <div class="service-change-header">
                <h4>Ocorrência ${index + 1}</h4>
                <button class="btn-small btn-danger remove-incident-row" data-id="${row.id}">
                    <i class="fas fa-trash-alt"></i> Remover
                </button>
            </div>
            <div class="incident-vtr-row">
                <label for="incident-vtr-${row.id}">Posto/VTR:</label>
                <select class="incident-vtr-select" data-id="${row.id}" id="incident-vtr-${row.id}">
                    <option value="">Selecione</option>
                    ${vtrOptions.join('')}
                </select>
            </div>
            <textarea class="service-change-textarea incident-textarea" data-id="${row.id}" placeholder="Descreva a ocorrência registrada">${row.description}</textarea>
        `;
        incidentsContainer.appendChild(itemDiv);
        // Seleciona o valor salvo
        const select = itemDiv.querySelector('.incident-vtr-select');
        if (select) select.value = row.vtr || '';
    });

    document.querySelectorAll('.incident-vtr-select').forEach(select => {
        select.addEventListener('change', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            const value = e.target.value;
            const rowIndex = incidentsRows.findIndex(row => row.id === rowId);
            if (rowIndex !== -1) {
                incidentsRows[rowIndex].vtr = value;
                triggerSave();
            }
        });
    });

    document.querySelectorAll('.incident-textarea').forEach(textarea => {
        autoExpandTextarea(textarea);
        textarea.addEventListener('input', (e) => {
            autoExpandTextarea(e.target);
            const rowId = parseInt(e.target.getAttribute('data-id'));
            const value = e.target.value;
            const rowIndex = incidentsRows.findIndex(row => row.id === rowId);
            if (rowIndex !== -1) {
                incidentsRows[rowIndex].description = value;
                triggerSave();
            }
        });
    });

    document.querySelectorAll('.remove-incident-row').forEach(button => {
        button.addEventListener('click', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            removeIncidentRow(rowId);
        });
    });

    if (incidentsNoChangeRow) {
        incidentsNoChangeRow.classList.toggle('hidden', incidentsRows.length > 0);
        if (incidentsRows.length > 0 && incidentsNoChange) {
            incidentsNoChange.checked = false;
        }
    }
}

function getFormData() {
    const operationNumberMeta = getOperationNumberMeta();
    const incidentsText = incidentsRows
        .map(row => (row.description || '').trim())
        .filter(Boolean)
        .join('\n\n');
    return {
        operationNumber: operationNumberMeta.number,
        operationSequence: operationNumberMeta.sequence,
        operationYear: operationNumberMeta.year,
        operationName: document.getElementById('operationName').value,
        operationComplexity: document.getElementById('operationComplexity').value,
        eventType: document.getElementById('eventType').value,
        operationDate: document.getElementById('operationDate').value,
        startTime: document.getElementById('startTime').value,
        endTime: document.getElementById('endTime').value,
        location: document.getElementById('location').value,
        neighborhood: document.getElementById('neighborhood').value,
        coordinatorName: document.getElementById('coordinatorName').value,
        registration: document.getElementById('registration').value,
        position: document.getElementById('position').value,
        totalAgents: document.getElementById('totalAgents').textContent,
        incidents: incidentsText,
        incidentsRows: incidentsRows,
        summary: document.getElementById('summary').value,
        areaRows: areaRows,
        vehicleRows: vehicleRows,
        serviceChangeRows: serviceChangeRows,
        lastSaved: new Date().toISOString()
    };
}

function applyFormData(formData) {
    if (operationNumberInput) {
        operationNumberInput.value = formData.operationNumber || '';
        if (formData.operationSequence) {
            operationNumberInput.dataset.sequence = String(formData.operationSequence);
        }
        if (formData.operationYear) {
            operationNumberInput.dataset.year = String(formData.operationYear);
        }
        if ((!formData.operationSequence || !formData.operationYear) && formData.operationNumber) {
            const match = String(formData.operationNumber).match(/^(\d{4})\/(\d{4})$/);
            if (match) {
                operationNumberInput.dataset.sequence = match[1];
                operationNumberInput.dataset.year = match[2];
            }
        }
    }
    document.getElementById('operationName').value = formData.operationName || '';
    document.getElementById('operationComplexity').value = formData.operationComplexity || '';
    if (eventTypeSelect) {
        eventTypeSelect.value = formData.eventType || '';
    }
    document.getElementById('operationDate').value = formData.operationDate || '';
    document.getElementById('startTime').value = formData.startTime || '';
    document.getElementById('endTime').value = formData.endTime || '';
    document.getElementById('location').value = formData.location || '';
    document.getElementById('neighborhood').value = formData.neighborhood || '';
    document.getElementById('coordinatorName').value = formData.coordinatorName || '';
    document.getElementById('registration').value = formData.registration || '';
    document.getElementById('position').value = formData.position || '';
    document.getElementById('totalAgents').textContent = formData.totalAgents || '0';
    document.getElementById('summary').value = formData.summary || '';
    currentOperationMeta = {
        ownerId: formData.ownerId || null,
        createdBy: formData.createdBy || null,
        updatedBy: formData.updatedBy || null,
        createdAt: formData.createdAt || null,
        updatedAt: formData.updatedAt || null
    };

    if (formData.areaRows && formData.areaRows.length > 0) {
        areaRows = formData.areaRows.map(row => ({
            ...row,
            post: row.post || row.vehicle || ''
        }));
        updateAreaTable();
        updateAreaCards();
        updateTotalAgentsCount();
    }

    if (formData.vehicleRows && formData.vehicleRows.length > 0) {
        vehicleRows = formData.vehicleRows;
        updateVehiclesTable();
        updateVehiclesCards();
    }

    if (formData.serviceChangeRows && formData.serviceChangeRows.length > 0) {
        serviceChangeRows = formData.serviceChangeRows.map((row, index) => ({
            ...row,
            index: index + 1
        }));
        updateServiceChangesContainer();
    } else {
        serviceChangeRows = [];
        updateServiceChangesContainer();
    }

    if (formData.incidentsRows && formData.incidentsRows.length > 0) {
        incidentsRows = formData.incidentsRows.map(row => ({
            ...row,
            description: row.description || ''
        }));
        updateIncidentsContainer();
    } else if (formData.incidents && String(formData.incidents).trim()) {
        incidentsRows = [{ id: Date.now(), description: String(formData.incidents) }];
        updateIncidentsContainer();
    } else {
        incidentsRows = [];
        updateIncidentsContainer();
    }

    updateTotalAgentsCount();

    requestAnimationFrame(() => {
        requestAnimationFrame(expandAllTextareas);
    });

    assignNextOperationNumber();
    enforceEditWindow();
    updateSummaryTemplateState();
}

// Save form data to localStorage
function saveFormData() {
    const formData = getFormData();
    
    // Save to localStorage
    localStorage.setItem('policiaMunicipalOperacaoForm', JSON.stringify(formData));
    
    // Update UI
    updateLastSavedTime();
    saveIndicator.classList.remove('saving');
    isSaving = false;
    saveStatus.textContent = 'Salvo automaticamente';
}

// Load form data from localStorage
function loadFormData() {
    const savedData = localStorage.getItem('policiaMunicipalOperacaoForm');
    
    if (savedData) {
        try {
            const formData = JSON.parse(savedData);
            
            applyFormData(formData);
            
            // Update last saved time
            if (formData.lastSaved) {
                const lastSavedDate = new Date(formData.lastSaved);
                lastSaved.textContent = `Último salvamento: ${formatTime(lastSavedDate)}`;
            }
            
        } catch (e) {
            console.error('Erro ao carregar dados salvos:', e);
        }
    }
}

async function saveOperationToFirebase(formData) {
    if (!db || !auth || !auth.currentUser) return;

    try {
        const operationsRef = db.collection('operations');
        let docRef = null;
        const user = auth.currentUser;
        if (currentOperationId && currentOperationMeta?.ownerId !== user.uid) {
            saveStatus.textContent = 'Você não tem permissão para editar esta operação.';
            return null;
        }
        let previousData = null;
        const userInfo = {
            uid: user.uid,
            nomeGuerra: currentUserProfile?.nomeGuerra || user.displayName || null,
            classe: currentUserProfile?.classe || null,
            matricula: currentUserProfile?.matricula || null
        };
        const payload = {
            ...formData,
            ownerId: auth.currentUser.uid,
            updatedBy: userInfo,
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };

        if (currentOperationId) {
            docRef = operationsRef.doc(currentOperationId);
            const previousSnap = await docRef.get();
            previousData = previousSnap.exists ? previousSnap.data() : null;
        } else {
            docRef = operationsRef.doc();
            currentOperationId = docRef.id;
            payload.createdAt = firebase.firestore.FieldValue.serverTimestamp();
            payload.createdBy = userInfo;
        }

        await docRef.set(payload, { merge: true });
        if (previousData) {
            const changedFields = getChangedOperationFields(previousData, payload);
            if (changedFields.length > 0) {
                await registerOperationEditLog({
                    operationId: docRef.id,
                    operationNumber: payload.operationNumber || previousData.operationNumber || null,
                    operationName: payload.operationName || previousData.operationName || null,
                    changedFields
                });
            }
        }
        saveStatus.textContent = 'Salvo!';
        return docRef.id;
    } catch (error) {
        console.error('Erro ao salvar:', error);
        saveStatus.textContent = 'Falha ao salvar';
        return null;
    }
}

async function loadOperationsList(force = false) {
    if (!db || !auth || !auth.currentUser || isLoadingOperations) return;
    if (operationsCache.length > 0 && !force) {
        applyOperationsFilter();
        return;
    }

    isLoadingOperations = true;
    operationsList.innerHTML = '<div class="operation-empty">Carregando operações...</div>';

    try {
        const snapshot = await db
            .collection('operations')
            .orderBy('updatedAt', 'desc')
            .get();

        operationsCache = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        operationsCurrentPage = 1;
        applyOperationsFilter();
    } catch (error) {
        console.error('Erro ao carregar operações:', error);
        operationsList.innerHTML = '<div class="operation-empty">Falha ao carregar operações.</div>';
        operationsPagination.innerHTML = '';
    } finally {
        isLoadingOperations = false;
    }
}

function applyOperationsFilter() {
    const textValue = (filterText.value || '').toLowerCase();
    const dateFrom = filterDateFrom.value ? new Date(filterDateFrom.value) : null;
    const dateTo = filterDateTo.value ? new Date(filterDateTo.value) : null;

    const filtered = operationsCache.filter(operation => {
        if (forcedOperationIdFilter && operation.id !== forcedOperationIdFilter) return false;
        if (isRestricted && auth?.currentUser && operation.ownerId !== auth.currentUser.uid) return false;
        const searchText = [
            operation.operationNumber || '',
            operation.eventType || '',
            operation.operationComplexity || '',
            operation.operationName || '',
            formatDate(operation.operationDate),
            operation.location || '',
            operation.neighborhood || '',
            operation.totalAgents || '',
            operation.coordinatorName || '',
            operation.registration || '',
            operation.createdBy?.nomeGuerra || '',
            formatDateTime(operation.createdAt),
            operation.updatedBy?.nomeGuerra || '',
            formatDateTime(operation.updatedAt)
        ].join(' ').toLowerCase();
        if (textValue && !searchText.includes(textValue)) return false;

        if (dateFrom || dateTo) {
            const operationDate = operation.operationDate ? new Date(operation.operationDate) : null;
            if (!operationDate) return false;
            if (dateFrom && operationDate < dateFrom) return false;
            if (dateTo && operationDate > dateTo) return false;
        }

        return true;
    });

    renderOperationsList(filtered);
}

function renderOperationsList(operations) {
    if (!operations.length) {
        operationsList.innerHTML = '<div class="operation-empty">Nenhuma operação encontrada.</div>';
        operationsPagination.innerHTML = '';
        return;
    }

    const totalPages = Math.ceil(operations.length / OPERATIONS_PAGE_SIZE);
    if (operationsCurrentPage > totalPages) {
        operationsCurrentPage = totalPages;
    }
    const startIndex = (operationsCurrentPage - 1) * OPERATIONS_PAGE_SIZE;
    const pageItems = operations.slice(startIndex, startIndex + OPERATIONS_PAGE_SIZE);

    operationsList.innerHTML = '';
    pageItems.forEach(operation => {
        const showUpdateInfo = shouldShowUpdateInfo(operation);
        const card = document.createElement('div');
        card.className = 'operation-card';
        card.innerHTML = `
            <div class="operation-header">
                <div class="operation-title">${operation.operationName || 'Operação sem nome'}</div>
                <div class="operation-highlight">
                        <span><strong>Nº:</strong> ${operation.operationNumber || 'Não informado'}</span>
                        <span><strong>Tipo:</strong> ${operation.eventType || 'Não informado'}</span>
                        <span><strong>Complexidade:</strong> ${operation.operationComplexity || 'Não informado'}</span>
                    <span><strong>Data:</strong> ${formatDate(operation.operationDate)}</span>
                    <span><strong>Local:</strong> ${(operation.location || 'Não informado')}${operation.neighborhood ? ` - ${operation.neighborhood}` : ''}</span>
                    <span><strong>Efetivo:</strong> ${operation.totalAgents || '0'}</span>
                </div>
            </div>
            <div class="operation-meta">
                <div><strong>Coordenador:</strong> ${operation.coordinatorName || 'Não informado'}</div>
                <div><strong>Registrado por:</strong> ${(operation.createdBy?.nomeGuerra || 'Não informado')}</div>
                <div><strong>Criado em:</strong> ${formatDateTime(operation.createdAt)}</div>
                ${showUpdateInfo ? `
                <div><strong>Editado por:</strong> ${(operation.updatedBy?.nomeGuerra || 'Não informado')}</div>
                <div><strong>Atualizado em:</strong> ${formatDateTime(operation.updatedAt)}</div>
                ` : ''}
            </div>
            <div class="operation-actions">
                <button class="btn-small" data-action="edit" data-id="${operation.id}">
                    <i class="fas fa-pen"></i> Editar
                </button>
                <button class="btn-small" data-action="pdf" data-id="${operation.id}">
                    <i class="fas fa-file-pdf"></i> Gerar PDF
                </button>
                ${isAdmin ? `
                <button class="btn-small btn-danger" data-action="delete" data-id="${operation.id}">
                    <i class="fas fa-trash"></i> Excluir
                </button>
                ` : ''}
            </div>
        `;

        operationsList.appendChild(card);
    });

    operationsList.querySelectorAll('button[data-action="edit"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = button.getAttribute('data-id');
            loadOperationById(id);
        });
    });

    operationsList.querySelectorAll('button[data-action="pdf"]').forEach(button => {
        button.addEventListener('click', async () => {
            const id = button.getAttribute('data-id');
            const operation = operationsCache.find(item => item.id === id);
            if (operation) {
                await exportToPDF(operation);
            }
        });
    });

    operationsList.querySelectorAll('button[data-action="delete"]').forEach(button => {
        button.addEventListener('click', async () => {
            if (!isAdmin) return;
            const id = button.getAttribute('data-id');
            const operation = operationsCache.find(item => item.id === id);
            if (!operation) return;

            const justification = prompt('Informe a justificativa para excluir este registro:');
            if (!justification || !justification.trim()) return;

            try {
                await db.collection('operations').doc(id).delete();
                operationsCache = operationsCache.filter(item => item.id !== id);
                if (forcedOperationIdFilter === id) {
                    forcedOperationIdFilter = null;
                }
                await registerAdminActionLog({
                    action: 'delete-operation',
                    justification: justification.trim(),
                    operationId: id,
                    operationNumber: operation.operationNumber || null,
                    operationName: operation.operationName || null
                });
                applyOperationsFilter();
            } catch (error) {
                console.error('Erro ao excluir operação:', error);
                alert('Falha ao excluir a operação.');
            }
        });
    });

    renderOperationsPagination(totalPages);
}

function renderOperationsPagination(totalPages) {
    operationsPagination.innerHTML = '';
    if (totalPages <= 1) return;

    const prevBtn = document.createElement('button');
    prevBtn.className = 'pagination-btn';
    prevBtn.textContent = 'Anterior';
    prevBtn.disabled = operationsCurrentPage === 1;
    prevBtn.addEventListener('click', () => {
        operationsCurrentPage -= 1;
        applyOperationsFilter();
    });
    operationsPagination.appendChild(prevBtn);

    for (let i = 1; i <= totalPages; i += 1) {
        const pageBtn = document.createElement('button');
        pageBtn.className = 'pagination-btn';
        if (i === operationsCurrentPage) pageBtn.classList.add('active');
        pageBtn.textContent = String(i);
        pageBtn.addEventListener('click', () => {
            operationsCurrentPage = i;
            applyOperationsFilter();
        });
        operationsPagination.appendChild(pageBtn);
    }

    const nextBtn = document.createElement('button');
    nextBtn.className = 'pagination-btn';
    nextBtn.textContent = 'Próximo';
    nextBtn.disabled = operationsCurrentPage === totalPages;
    nextBtn.addEventListener('click', () => {
        operationsCurrentPage += 1;
        applyOperationsFilter();
    });
    operationsPagination.appendChild(nextBtn);
}

async function loadOperationById(id) {
    if (!db || !auth || !auth.currentUser) return;

    try {
        const docRef = db.collection('operations').doc(id);
        const docSnap = await docRef.get();
        if (docSnap.exists) {
            const data = docSnap.data();
            if (data.ownerId && auth.currentUser && data.ownerId !== auth.currentUser.uid) {
                saveStatus.textContent = 'Você não tem permissão para editar esta operação.';
                alert('Você não tem permissão para editar esta operação.');
                return;
            }
            currentOperationId = docSnap.id;
            applyFormData(data);
            setActiveAppView('form');
            saveStatus.textContent = 'Operação carregada';
        }
    } catch (error) {
        console.error('Erro ao abrir operação:', error);
    }
}

// Update last saved time display
function updateLastSavedTime() {
    const now = new Date();
    lastSaved.textContent = `Último salvamento: ${formatTime(now)}`;
}

// Format time as HH:MM:SS
function formatTime(date) {
    const hours = date.getHours().toString().padStart(2, '0');
    const minutes = date.getMinutes().toString().padStart(2, '0');
    const seconds = date.getSeconds().toString().padStart(2, '0');
    return `${hours}:${minutes}:${seconds}`;
}

// Clear form data
function clearForm() {
        currentOperationMeta = null;
    // Clear all form inputs
    formInputs.forEach(input => {
        if (input.type === 'number') {
            input.value = '0';
        } else {
            input.value = '';
        }
    });
    if (incidentsNoChange) incidentsNoChange.checked = false;
    if (serviceChangesNoChange) serviceChangesNoChange.checked = false;
    if (operationNumberInput) {
        operationNumberInput.dataset.sequence = '';
        operationNumberInput.dataset.year = '';
    }
    
    // Clear date/time fields explicitly
    document.getElementById('operationDate').value = '';
    document.getElementById('startTime').value = '';
    document.getElementById('endTime').value = '';
    
    // Reset dynamic rows
    areaRows = [{
        id: Date.now(),
        area: '',
        agents: '',
        post: ''
    }];
    
    vehicleRows = [{
        id: Date.now() + 1,
        prefix: '',
        post: '',
        type: ''
    }];
    
    serviceChangeRows = [];
    incidentsRows = [];
    
    // Update tables and cards
    updateAreaTable();
    updateAreaCards();
    updateTotalAgentsCount();
    updateVehiclesTable();
    updateVehiclesCards();
    updateServiceChangesContainer();
    updateIncidentsContainer();
    
    // Update counts
    updateTotalAgentsCount();
    
    // Clear localStorage
    localStorage.removeItem('policiaMunicipalOperacaoForm');
    currentOperationId = null;
    
    // Close modal
    confirmationModal.style.display = 'none';
    
    // Update save indicator
    updateLastSavedTime();
    saveStatus.textContent = 'Formulário limpo';
    setTimeout(() => {
        saveStatus.textContent = 'Salvo automaticamente';
    }, 2000);

    updateSummaryTemplateState();
}

// Export to PDF with improved layout
function loadImage(src) {
    return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => resolve(null);
        img.src = src;
    });
}

// Nova versão: recebe dados da operação como parâmetro
async function exportToPDF(operationData) {
    // Create PDF document in A4 format
    const doc = new jsPDF('p', 'mm', 'a4');
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 25;
    const bottomMargin = 35; // Espaço reservado para rodapé
    const contentWidth = pageWidth - 2 * margin;
    let yPos = margin;
    
    // Helper function to add new page if needed
    // Store current font state
    let currentFontSize = 11;
    let currentTextColor = { r: 0, g: 0, b: 0 };
    let currentFontStyle = 'normal';
    
    // Wrapper functions to track font state
    function setFontSize(size) {
        doc.setFontSize(size);
        currentFontSize = size;
    }
    
    function setTextColor(r, g, b) {
        doc.setTextColor(r, g, b);
        currentTextColor = { r, g, b };
    }
    
    function setFontStyle(style) {
        doc.setFont('helvetica', style);
        currentFontStyle = style;
    }
    
    // Function to render text with justification
    function renderJustifiedText(text, x, y, maxWidth) {
        const words = text.trim().split(/\s+/);
        
        if (words.length === 1) {
            // Single word, just render it left-aligned
            doc.text(words[0], x, y);
            return;
        }
        
        // Calculate the width of the text with normal spacing
        const normalSpacing = words.map(word => doc.getTextWidth(word)).reduce((a, b) => a + b, 0);
        const normalSpaceWidth = doc.getTextWidth(' ');
        const normalTotalWidth = normalSpacing + (normalSpaceWidth * (words.length - 1));
        
        // If the line doesn't reach 85% of maxWidth, render left-aligned (it's a short line)
        if (normalTotalWidth < maxWidth * 0.85) {
            doc.text(text, x, y);
            return;
        }
        
        // Get the width of a space character
        const spaceWidth = doc.getTextWidth(' ');
        const totalWordsWidth = words.reduce((sum, word) => sum + doc.getTextWidth(word), 0);
        const totalSpacesNeeded = maxWidth - totalWordsWidth;
        const spaceBetweenWords = totalSpacesNeeded / (words.length - 1);
        
        // Render each word with calculated spacing
        let currentX = x;
        words.forEach((word, index) => {
            doc.text(word, currentX, y);
            
            if (index < words.length - 1) {
                currentX += doc.getTextWidth(word) + spaceBetweenWords;
            }
        });
    }
    
    function checkPageBreak(heightNeeded) {
        if (yPos + heightNeeded > pageHeight - bottomMargin) {
            // Save current font state
            const savedFontSize = currentFontSize;
            const savedTextColor = { ...currentTextColor };
            const savedFontStyle = currentFontStyle;
            
            doc.addPage();
            yPos = margin;
            // Add header to new page
            addPageHeader();
            
            // Restore previous font state
            setFontSize(savedFontSize);
            setTextColor(savedTextColor.r, savedTextColor.g, savedTextColor.b);
            setFontStyle(savedFontStyle);
            
            return true;
        }
        return false;
    }
    
    // Function to add page header
    function addPageHeader() {
        doc.setFontSize(10);
        doc.setTextColor(100, 100, 100);
        doc.text(`Polícia Municipal de Aracaju - Relatório de Operação - Página ${doc.getNumberOfPages()}`, margin, 15);
        doc.line(margin, 18, pageWidth - margin, 18);
        yPos = margin;
    }
    
    // Clean operation name for filename
    const operationName = operationData?.operationName || 'Não informado';
    const cleanOperationName = operationName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
    
    const logoImage = await loadImage('logo.jpg');

    // Add main header
    doc.setFillColor(30, 58, 95);
    doc.rect(0, 0, pageWidth, 55, 'F');

    let logoBottomY = 6;
    if (logoImage) {
        const logoMaxWidth = 22;
        const logoMaxHeight = 22;
        const ratio = logoImage.naturalWidth / logoImage.naturalHeight;
        let logoWidth = logoMaxWidth;
        let logoHeight = logoWidth / ratio;
        if (logoHeight > logoMaxHeight) {
            logoHeight = logoMaxHeight;
            logoWidth = logoHeight * ratio;
        }
        const logoX = (pageWidth - logoWidth) / 2;
        const logoY = 6;
        logoBottomY = logoY + logoHeight;
        doc.addImage(logoImage, 'JPEG', logoX, logoY, logoWidth, logoHeight);
    }
    
    // Title in white
    setTextColor(255, 255, 255);
    const titleY = Math.max(logoBottomY + 10, 30);
    const subtitleY = titleY + 10;
    setFontSize(22);
    setFontStyle('bold');
    doc.text('POLÍCIA MUNICIPAL DE ARACAJU', pageWidth / 2, titleY, { align: 'center' });
    
    setFontSize(18);
    doc.text('RELATÓRIO DE OPERAÇÃO', pageWidth / 2, subtitleY, { align: 'center' });
    
    // Add decorative line
    doc.setDrawColor(201, 162, 39);
    doc.setLineWidth(2);
    doc.line(margin, 60, pageWidth - margin, 60);
    
    yPos = 70;
    
    // Function to add section with smaller spacing
    function addSection(title, content, isTable = false) {
        // Check page break for title only
        checkPageBreak(8);
        
        // Section title
        setTextColor(30, 58, 95);
        setFontSize(14);
        setFontStyle('bold');
        doc.text(title, margin, yPos);
        
        yPos += 8;
        
        // Section content
        setTextColor(0, 0, 0);
        setFontSize(11);
        setFontStyle('normal');
        
        if (!isTable && content) {
            // Split text into lines with proper width
            const lines = doc.splitTextToSize(content, contentWidth - 2);
            
            // Render each line
            lines.forEach((line, index) => {
                // Check page break for each line to prevent overflow
                checkPageBreak(6);
                
                doc.text(line, margin + 1, yPos);
                yPos += 7.5;
            });
            
            yPos += 8;
        }
        

        return yPos;
    }
    
    // Function to add section with justified text (for textareas)
    function addSectionJustified(title, content, isTable = false) {
        // Check page break for title only
        checkPageBreak(8);
        
        // Section title
        setTextColor(30, 58, 95);
        setFontSize(14);
        setFontStyle('bold');
        doc.text(title, margin, yPos);
        
        yPos += 8;
        
        // Section content
        setTextColor(0, 0, 0);
        setFontSize(11);
        setFontStyle('normal');
        
        if (!isTable && content) {
            // Split text into lines with proper width
            const lines = doc.splitTextToSize(content, contentWidth - 2);
            
            // Render each line with justification
            lines.forEach((line, index) => {
                // Check page break for each line to prevent overflow
                checkPageBreak(6);
                
                // Render justified text for all lines except the last
                if (index < lines.length - 1) {
                    renderJustifiedText(line, margin + 1, yPos, contentWidth - 2);
                } else {
                    // Last line: render left-aligned
                    doc.text(line, margin + 1, yPos);
                }
                yPos += 7.5;
            });
            
            yPos += 8;
        }
        

        return yPos;
    }
    
    // Function to create a professional table
    function createTable(headers, data, colWidths, showTotal = false, totalLabel = '') {
        // Sempre desenha o cabeçalho antes de cada página
        function drawTableHeader() {
            doc.setFillColor(30, 58, 95);
            doc.rect(margin, yPos, contentWidth, 7, 'F');
            doc.setTextColor(255, 255, 255);
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(10);
            let xPos = margin;
            for (let i = 0; i < headers.length; i++) {
                doc.text(headers[i], xPos + 2, yPos + 4.5);
                xPos += colWidths[i];
            }
            yPos += 7;
        }

        drawTableHeader();
        doc.setTextColor(0, 0, 0);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(10);

        data.forEach((row, rowIndex) => {
            // Se não couber mais uma linha, quebra a página e redesenha o cabeçalho
            if (yPos + 7 > pageHeight - bottomMargin) {
                doc.addPage();
                yPos = margin;
                addPageHeader();
                drawTableHeader();
                doc.setTextColor(0, 0, 0); // Corrige cor do texto após quebra de página
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(10);
            }
            // Alternate row colors
            if (rowIndex % 2 === 0) {
                doc.setFillColor(245, 245, 245);
                doc.rect(margin, yPos, contentWidth, 7, 'F');
            }
            // Add row content
            let xPos = margin;
            for (let i = 0; i < row.length; i++) {
                const cellContent = row[i] || '-';
                const lines = doc.splitTextToSize(cellContent, colWidths[i] - 4);
                doc.text(lines[0] || '-', xPos + 2, yPos + 4.5);
                xPos += colWidths[i];
            }
            yPos += 7;
        });

        yPos += 8;

        // Add total if requested
        if (showTotal && totalLabel) {
            if (yPos + 7 > pageHeight - bottomMargin) {
                doc.addPage();
                yPos = margin;
                addPageHeader();
            }
            doc.setFillColor(240, 240, 240);
            doc.rect(margin, yPos, contentWidth, 7, 'F');
            doc.setTextColor(30, 58, 95);
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(10);
            doc.text(totalLabel, margin + 5, yPos + 4.5);
            yPos += 10;
        }

        yPos += 12;
    }
    
    // 1. IDENTIFICAÇÃO DA OPERAÇÃO
    const operationNumber = operationData?.operationNumber || 'Não informado';
    const operationDate = operationData?.operationDate || 'Não informado';
    const startTime = operationData?.startTime || 'Não informado';
    const endTime = operationData?.endTime || 'Não informado';
    const eventType = operationData?.eventType || 'Não informado';
    const location = operationData?.location || 'Não informado';
    const neighborhood = operationData?.neighborhood || 'Não informado';
    
    const identificationContent = `Operação: ${operationName}
Nº da Operação: ${operationNumber}
Tipo Evento: ${eventType}
Data: ${formatDate(operationDate)}
Período: ${startTime} às ${endTime}
Local: ${location} - ${neighborhood}`;
    
    addSection('1. IDENTIFICAÇÃO DA OPERAÇÃO', identificationContent);
     // Adicione espaço extra aqui:
    yPos += 2;
    
    // 2. COORDENAÇÃO / RESPONSÁVEL
    const coordinatorName = operationData?.coordinatorName || 'Não informado';
    const registration = operationData?.registration || 'Não informado';
    const position = operationData?.position || 'Não informado';
    
    const coordinationContent = `Coordenador: ${coordinatorName}
Matrícula: ${registration}
Cargo/Função: ${position}`;
    
    addSection('2. COORDENAÇÃO / RESPONSÁVEL', coordinationContent);
 // Adicione espaço extra aqui:
    yPos += 2;
    
    // 3. VIATURAS (VTRs) UTILIZADAS - Table
    checkPageBreak(15);
    addSection('3. VIATURAS (VTRs) UTILIZADAS', '');

     // Adicione espaço extra aqui:
    yPos += 2;
    
    const vehicleRowsData = operationData?.vehicleRows || [];
    if (vehicleRowsData.length > 0) {
        const vehicleHeaders = ['PREFIXO', 'POSTO', 'TIPO'];
        const vehicleData = vehicleRowsData.map(row => [
            row.prefix || '-',
            row.post || '-',
            row.type || '-'
        ]);
        const vehicleColWidths = [40, 50, 60];
        createTable(vehicleHeaders, vehicleData, vehicleColWidths, true, `TOTAL DE VIATURAS: ${vehicleRowsData.length}`);
    } else {
        setTextColor(100, 100, 100);
        doc.text('Nenhuma viatura cadastrada.', margin, yPos);
        yPos += 8;
    }
    
    // 4. DISTRIBUIÇÃO DO EFETIVO POR ÁREA - Table
    checkPageBreak(25);
    addSection('4. DISTRIBUIÇÃO DO EFETIVO POR ÁREA', '');

     // Adicione espaço extra aqui:
    yPos += 2;
    
    const areaRowsData = operationData?.areaRows || [];
    if (areaRowsData.length > 0) {
        const areaHeaders = ['POSTO', 'QTD. AGENTES', 'ÁREA/LOCAL'];
        const areaData = areaRowsData.map(row => [
            row.post || row.vehicle || '-',
            row.agents || '0',
            row.area || '-'
        ]);
        const areaColWidths = [45, 35, 70];
        const totalAgentsDistributed = areaRowsData.reduce((sum, row) => sum + (parseInt(row.agents) || 0), 0);
        createTable(areaHeaders, areaData, areaColWidths, true, `TOTAL DE AGENTES DISTRIBUÍDOS: ${totalAgentsDistributed}`);
    } else {
        setTextColor(100, 100, 100);
        doc.text('Nenhuma área cadastrada.', margin, yPos);
        yPos += 8;
    }

    // 5. REGISTRO DE OCORRÊNCIAS
    const incidentsRowsData = operationData?.incidentsRows || [];
    if (incidentsRowsData.length > 0) {
        checkPageBreak(8);
        setTextColor(30, 58, 95);
        setFontSize(14);
        setFontStyle('bold');
        doc.text('5. REGISTRO DE OCORRÊNCIAS', margin, yPos);
        yPos += 8;

        incidentsRowsData.forEach((row, index) => {
            checkPageBreak(10);
            // Título em negrito
            setTextColor(30, 58, 95);
            setFontSize(12);
            setFontStyle('bold');
            doc.text(`Ocorrência ${index + 1}`, margin + 1, yPos);
            // VTR associada
            setFontStyle('normal');
            setTextColor(0, 0, 0);
            if (row.vtr) {
                doc.text(`Posto/VTR: ${row.vtr}`, margin + 50, yPos);
            }
            yPos += 7.5;
            // Descrição
            setFontSize(11);
            setFontStyle('normal');
            if (row.description && row.description.trim() !== '') {
                const lines = doc.splitTextToSize(row.description, contentWidth - 2);
                lines.forEach((line, lineIndex) => {
                    checkPageBreak(6);
                    // Justifica todas as linhas exceto a última
                    if (lineIndex < lines.length - 1) {
                        renderJustifiedText(line, margin + 3, yPos, contentWidth - 2);
                    } else {
                        doc.text(line, margin + 3, yPos);
                    }
                    yPos += 7.5;
                });
            }
            yPos += 2; // Espaço de 2px entre ocorrências
        });
        yPos += 8;
    } else {
        addSectionJustified('5. REGISTRO DE OCORRÊNCIAS', 'Nenhuma ocorrência registrada durante a operação.');
    }
    
     // Adicione espaço extra aqui:
    yPos += 2;

    // 6. ALTERAÇÕES DE SERVIÇO
    checkPageBreak(8);
    addSection('6. ALTERAÇÕES DE SERVIÇO', '');

    // Adicione espaço extra aqui:
    yPos += 2;
    
    // Verificar se há alterações de serviço com conteúdo
    const serviceChangeRowsData = operationData?.serviceChangeRows?.map((row, index) => ({ ...row, index: index + 1 })) || [];
    const hasServiceChanges = serviceChangeRowsData.some(row => row.description && row.description.trim() !== '');
    
    if (hasServiceChanges) {
        setTextColor(0, 0, 0);
        setFontSize(11);
        setFontStyle('normal');
        serviceChangeRowsData.forEach((row) => {
            const description = row.description && row.description.trim() !== '' ? row.description : null;
            if (description) {
                checkPageBreak(8);
                setTextColor(30, 58, 95);
                setFontStyle('bold');
                doc.text(`Alteração ${row.index}:`, margin, yPos);
                yPos += 8;
                setTextColor(0, 0, 0);
                setFontStyle('normal');
                const lines = doc.splitTextToSize(description, contentWidth - 2);
                lines.forEach((line, lineIndex) => {
                    checkPageBreak(6);
                    if (lineIndex < lines.length - 1) {
                        renderJustifiedText(line, margin + 1, yPos, contentWidth - 2);
                    } else {
                        doc.text(line, margin + 1, yPos);
                    }
                    yPos += 7.5;
                });
                yPos += 7.5;
            }
        });
    } else {
        setTextColor(100, 100, 100);
        doc.text('Nenhuma alteração de serviço registrada.', margin, yPos);
        yPos += 8;
    }
     // Adicione espaço extra aqui:
    yPos += 2;

    // 7. RESUMO DA OPERAÇÃO
    const summary = operationData?.summary || 'Nenhum resumo fornecido.';
    addSectionJustified('7. RESUMO DA OPERAÇÃO', summary);

     // Adicione espaço extra aqui:
    yPos += 2;
    
    // User info area (registrador + emissão)
    checkPageBreak(60);
    yPos += 6;

    doc.setDrawColor(150, 150, 150);
    doc.line(margin, yPos, pageWidth - margin, yPos);
    yPos += 12;

    setTextColor(30, 58, 95);
    setFontSize(12);
    setFontStyle('bold');
    doc.text('USUÁRIO REGISTRADOR', pageWidth / 2, yPos, { align: 'center' });

    yPos += 10;

    const registradorProfile = operationData?.createdBy || {};
    const createdAtText = formatDateTime(operationData?.createdAt);
    const updatedAtText = formatDateTime(operationData?.updatedAt);

    setTextColor(0, 0, 0);
    setFontSize(11);
    setFontStyle('normal');

    const infoLeftX = margin + 5;
    const infoRightX = pageWidth / 2 + 5;

    doc.text(`Classe: ${registradorProfile.classe || 'Não informado'}`, infoLeftX, yPos);
    doc.text(`Nome: ${registradorProfile.nomeGuerra || 'Não informado'}`, infoRightX, yPos);
    yPos += 7;
    doc.text(`Matrícula: ${registradorProfile.matricula || 'Não informado'}`, infoLeftX, yPos);
    doc.text(`Criado em: ${createdAtText}`, infoRightX, yPos);
    yPos += 7;
    doc.text(`Atualizado em: ${updatedAtText}`, infoLeftX, yPos);

    yPos += 10;
    doc.setDrawColor(150, 150, 150);
    doc.line(margin, yPos, pageWidth - margin, yPos);
    yPos += 12;

    setTextColor(30, 58, 95);
    setFontSize(12);
    setFontStyle('bold');
    doc.text('RESPONSÁVEL PELA EMISSÃO', pageWidth / 2, yPos, { align: 'center' });

    yPos += 10;

    const emissionProfile = currentUserProfile || {};
    const emissionDateTime = new Date();

    setTextColor(0, 0, 0);
    setFontSize(11);
    setFontStyle('normal');

    doc.text(`Classe: ${emissionProfile.classe || 'Não informado'}`, infoLeftX, yPos);
    doc.text(`Nome: ${emissionProfile.nomeGuerra || 'Não informado'}`, infoRightX, yPos);
    yPos += 7;
    doc.text(`Matrícula: ${emissionProfile.matricula || 'Não informado'}`, infoLeftX, yPos);
    doc.text(`Emissão: ${emissionDateTime.toLocaleString('pt-BR')}`, infoRightX, yPos);
    
    // Add footer on all pages
    const totalPages = doc.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        
        // Bottom border
        doc.setDrawColor(201, 162, 39);
        doc.setLineWidth(0.5);
        doc.line(margin, pageHeight - 20, pageWidth - margin, pageHeight - 20);
        
        // Page number
        setFontSize(9);
        setTextColor(100, 100, 100);
        doc.text(`Página ${i} de ${totalPages}`, pageWidth / 2, pageHeight - 15, { align: 'center' });
        
        // Footer text
        setFontSize(8);
        doc.text('Polícia Municipal de Aracaju - Gestão de Operações - © 2026', 
               pageWidth / 2, pageHeight - 10, { align: 'center' });
    }
    
    const dateStr = new Date().toLocaleDateString('pt-BR');
    const fileName = `Relatorio_Operacao_${cleanOperationName}_${dateStr.replace(/\//g, '-')}.pdf`;

    // Save the PDF with cleaned filename
    doc.save(fileName);

    await registerReportLog({
        operationId: currentOperationId || null,
        operationNumber: operationData?.operationNumber || null,
        operationName: operationData?.operationName || null
    });
}

// Helper function to format date
function formatDate(dateString) {
    if (!dateString) return 'Não informado';
    const date = new Date(dateString);
    return date.toLocaleDateString('pt-BR');
}

function formatDateTime(value) {
    if (!value) return 'Não informado';
    let date = value;
    if (typeof value?.toDate === 'function') {
        date = value.toDate();
    } else if (typeof value === 'string' || typeof value === 'number') {
        date = new Date(value);
    }
    if (date instanceof Date && !Number.isNaN(date.getTime())) {
        return date.toLocaleString('pt-BR');
    }
    return 'Não informado';
}

function shouldShowUpdateInfo(operation) {
    if (!operation || !operation.updatedAt || !operation.createdAt) return false;
    const createdAt = normalizeTimestamp(operation.createdAt);
    const updatedAt = normalizeTimestamp(operation.updatedAt);
    if (!createdAt || !updatedAt) return false;
    return updatedAt.getTime() > createdAt.getTime();
}

function normalizeTimestamp(value) {
    if (!value) return null;
    if (typeof value?.toDate === 'function') {
        return value.toDate();
    }
    if (typeof value === 'string' || typeof value === 'number') {
        const date = new Date(value);
        return Number.isNaN(date.getTime()) ? null : date;
    }
    if (value instanceof Date) return value;
    return null;
}

async function registerAccessLog(user) {
    if (!db || !user) return;
    try {
        const payload = {
            uid: user.uid,
            email: user.email || null,
            classe: currentUserProfile?.classe || null,
            nomeGuerra: currentUserProfile?.nomeGuerra || user.displayName || null,
            matricula: currentUserProfile?.matricula || null,
            action: 'login',
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
        };
        await getAccessLogsCollection().add(payload);
    } catch (error) {
        console.error('Erro ao registrar acesso:', error);
    }
}

async function registerAdminActionLog({ action, justification, operationId, operationNumber, operationName, targetUserId, targetUserEmail }) {
    if (!db || !auth || !auth.currentUser) return;
    try {
        const payload = {
            uid: auth.currentUser.uid,
            email: auth.currentUser.email || null,
            classe: currentUserProfile?.classe || null,
            nomeGuerra: currentUserProfile?.nomeGuerra || auth.currentUser.displayName || null,
            matricula: currentUserProfile?.matricula || null,
            action,
            justification,
            operationId,
            operationNumber: operationNumber || null,
            operationName: operationName || null,
            targetUserId: targetUserId || null,
            targetUserEmail: targetUserEmail || null,
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
        };
        await getAccessLogsCollection().add(payload);
    } catch (error) {
        console.error('Erro ao registrar ação administrativa:', error);
    }
}

async function registerEventTypeLog({ action, eventType }) {
    if (!db || !auth || !auth.currentUser) return;
    try {
        const payload = {
            uid: auth.currentUser.uid,
            email: auth.currentUser.email || null,
            classe: currentUserProfile?.classe || null,
            nomeGuerra: currentUserProfile?.nomeGuerra || auth.currentUser.displayName || null,
            matricula: currentUserProfile?.matricula || null,
            action,
            justification: `Tipo de evento: ${eventType}`,
            operationId: null,
            operationName: null,
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
        };
        await getAccessLogsCollection().add(payload);
    } catch (error) {
        console.error('Erro ao registrar log de tipo de evento:', error);
    }
}

async function registerOperationEditLog({ operationId, operationNumber, operationName, changedFields }) {
    if (!db || !auth || !auth.currentUser) return;
    try {
        const payload = {
            uid: auth.currentUser.uid,
            email: auth.currentUser.email || null,
            classe: currentUserProfile?.classe || null,
            nomeGuerra: currentUserProfile?.nomeGuerra || auth.currentUser.displayName || null,
            matricula: currentUserProfile?.matricula || null,
            action: 'update-operation',
            justification: changedFields.join(', '),
            operationId,
            operationNumber: operationNumber || null,
            operationName: operationName || null,
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
        };
        await getAccessLogsCollection().add(payload);
    } catch (error) {
        console.error('Erro ao registrar edição da operação:', error);
    }
}

async function registerReportLog({ operationId, operationNumber, operationName }) {
    if (!db || !auth || !auth.currentUser) return;
    try {
        const payload = {
            uid: auth.currentUser.uid,
            email: auth.currentUser.email || null,
            classe: currentUserProfile?.classe || null,
            nomeGuerra: currentUserProfile?.nomeGuerra || auth.currentUser.displayName || null,
            matricula: currentUserProfile?.matricula || null,
            action: 'report-issued',
            justification: `operação "${operationName || 'Não informado'}" emitida.`,
            operationId: operationId || null,
            operationNumber: operationNumber || null,
            operationName: operationName || null,
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
        };
        await getAccessLogsCollection().add(payload);
    } catch (error) {
        console.error('Erro ao registrar emissão de relatório:', error);
    }
}

function getChangedOperationFields(previousData, nextData) {
    const fields = [
        { key: 'operationNumber', label: 'Nº da operação' },
        { key: 'operationName', label: 'Nome da operação' },
        { key: 'operationComplexity', label: 'Complexidade' },
        { key: 'eventType', label: 'Tipo de evento' },
        { key: 'operationDate', label: 'Data' },
        { key: 'startTime', label: 'Horário de início' },
        { key: 'endTime', label: 'Horário de término' },
        { key: 'location', label: 'Local' },
        { key: 'neighborhood', label: 'Bairro' },
        { key: 'coordinatorName', label: 'Coordenador' },
        { key: 'registration', label: 'Matrícula' },
        { key: 'position', label: 'Cargo/Função' },
        { key: 'totalAgents', label: 'Efetivo' },
        { key: 'incidents', label: 'Ocorrências' },
        { key: 'summary', label: 'Resumo' },
        { key: 'areaRows', label: 'Áreas de atuação' },
        { key: 'vehicleRows', label: 'Viaturas' },
        { key: 'serviceChangeRows', label: 'Alterações de Serviço' }
    ];

    const normalizeText = (value) => (value || '').toString().trim();

    const normalizeArrayRows = (key, value) => {
        if (!Array.isArray(value)) return value;
        const mapped = value.map(row => {
            if (!row || typeof row !== 'object') return row;
            if (key === 'areaRows') {
                return {
                    area: normalizeText(row.area),
                    agents: normalizeText(row.agents),
                    post: normalizeText(row.post || row.vehicle)
                };
            }
            if (key === 'vehicleRows') {
                return {
                    prefix: normalizeText(row.prefix),
                    post: normalizeText(row.post),
                    type: normalizeText(row.type)
                };
            }
            if (key === 'serviceChangeRows') {
                return { description: normalizeText(row.description) };
            }
            if (key === 'incidentsRows') {
                return { description: normalizeText(row.description) };
            }
            const { id, index, ...rest } = row;
            return rest;
        });

        const filtered = mapped.filter(item => {
            if (!item || typeof item !== 'object') return Boolean(item);
            return Object.values(item).some(value => normalizeText(value) !== '');
        });

        return filtered
            .map(item => JSON.stringify(item))
            .sort()
            .join('|');
    };

    return fields
        .filter(field => {
            const prevValue = previousData?.[field.key] ?? null;
            const nextValue = nextData?.[field.key] ?? null;
            if (Array.isArray(prevValue) || Array.isArray(nextValue)) {
                return normalizeArrayRows(field.key, prevValue) !== normalizeArrayRows(field.key, nextValue);
            }
            if (typeof prevValue === 'object' || typeof nextValue === 'object') {
                return JSON.stringify(prevValue || {}) !== JSON.stringify(nextValue || {});
            }
            return normalizeText(prevValue) !== normalizeText(nextValue);
        })
        .map(field => field.label);
}

async function loadLogsList(force = false) {
    if (!db || !auth || !auth.currentUser || !isAdmin || isLoadingLogs) return;
    if (logsCache.length > 0 && !force) {
        applyLogsFilter();
        return;
    }

    isLoadingLogs = true;
    if (logsTableBody) {
        logsTableBody.innerHTML = '<tr><td colspan="8">Carregando logs...</td></tr>';
    }

    try {
        const snapshot = await getAccessLogsCollection()
            .orderBy('createdAt', 'desc')
            .limit(logsCurrentLimit || 100)
            .get();
        logsCache = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        applyLogsFilter();
    } catch (error) {
        console.error('Erro ao carregar logs:', error);
        if (logsTableBody) {
            logsTableBody.innerHTML = '<tr><td colspan="8">Falha ao carregar logs.</td></tr>';
        }
    } finally {
        isLoadingLogs = false;
    }
}

function renderLogsTable(logs) {
    if (!logsTableBody) return;
    if (!logs.length) {
        logsTableBody.innerHTML = '<tr><td colspan="8">Nenhum log encontrado.</td></tr>';
        return;
    }

    logsTableBody.innerHTML = '';
    logs.forEach(log => {
        const actionLabel = getLogActionLabel(log.action);
        const operationInfo = log.operationNumber || log.operationId || '-';
        const justification = log.justification || '-';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${formatDateTime(log.createdAt)}</td>
            <td>${actionLabel}</td>
            <td>${log.classe || '-'}</td>
            <td>${log.nomeGuerra || '-'}</td>
            <td>${log.matricula || '-'}</td>
            <td>${log.email || '-'}</td>
            <td>${operationInfo}</td>
            <td>${justification}</td>
        `;
        logsTableBody.appendChild(tr);
    });
}

function getLogActionLabel(action) {
    if (!action || action === 'login') return 'Acesso';
    if (action === 'update-operation') return 'Alteração de operação';
    if (action === 'delete-operation') return 'Exclusão de operação';
    if (action === 'report-issued') return 'Emissão de relatório';
    if (action === 'delete-user') return 'Exclusão de usuário';
    return action;
}

function applyLogsFilter() {
    if (!logsCache.length) {
        renderLogsTable([]);
        return;
    }

    const fromValue = logsDateFrom?.value ? new Date(logsDateFrom.value) : null;
    const toValue = logsDateTo?.value ? new Date(logsDateTo.value) : null;

    const filtered = logsCache.filter(log => {
        const createdAt = normalizeTimestamp(log.createdAt);
        if (!createdAt) return false;
        if (fromValue && createdAt < fromValue) return false;
        if (toValue) {
            const endOfDay = new Date(toValue);
            endOfDay.setHours(23, 59, 59, 999);
            if (createdAt > endOfDay) return false;
        }
        if (logsActionFilter === 'access') {
            return !log.action || log.action === 'login';
        }
        if (logsActionFilter === 'report') {
            return log.action === 'report-issued';
        }
        if (logsActionFilter === 'update') {
            return log.action && log.action.startsWith('update');
        }
        if (logsActionFilter === 'delete') {
            return log.action && log.action.startsWith('delete');
        }
        return true;
    });

    renderLogsTable(filtered);
}

function hasServiceChangesContent() {
    return serviceChangeRows.some(row => (row.description || '').trim() !== '');
}

function validateOptionalFlags() {
    if (incidentsRows.length === 0 && incidentsNoChange && !incidentsNoChange.checked) {
        incidentsNoChange.focus();
        return false;
    }

    if (serviceChangeRows.length === 0 && serviceChangesNoChange && !serviceChangesNoChange.checked) {
        serviceChangesNoChange.focus();
        return false;
    }

    return true;
}