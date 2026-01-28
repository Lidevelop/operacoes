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
const OPERATIONS_PAGE_SIZE = 5;
let operationsCurrentPage = 1;

// DOM Elements
const loginView = document.getElementById('loginView');
const appView = document.getElementById('appView');
const loginForm = document.getElementById('loginForm');
const loginEmail = document.getElementById('loginEmail');
const loginPassword = document.getElementById('loginPassword');
const loginError = document.getElementById('loginError');
const logoutBtn = document.getElementById('logoutBtn');
const currentOperationBtn = document.getElementById('currentOperationBtn');
const newOperationBtn = document.getElementById('newOperationBtn');
const viewManagerBtn = document.getElementById('viewManagerBtn');
const operationFormView = document.getElementById('operationFormView');
const managerView = document.getElementById('managerView');
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
const addAreaBtn = document.getElementById('addAreaBtn');
const addAgentBtn = document.getElementById('addAgentBtn');
const addVehicleBtn = document.getElementById('addVehicleBtn');
const addServiceChangeBtn = document.getElementById('addServiceChangeBtn');
const areaTableBody = document.getElementById('areaTableBody');
const areaCardsContainer = document.getElementById('areaCardsContainer');
const agentsTableBody = document.getElementById('agentsTableBody');
const agentsCardsContainer = document.getElementById('agentsCardsContainer');
const vehiclesTableBody = document.getElementById('vehiclesTableBody');
const vehiclesCardsContainer = document.getElementById('vehiclesCardsContainer');
const saveDbBtn = document.getElementById('saveDbBtn');
const clearBtn = document.getElementById('clearBtn');
const confirmationModal = document.getElementById('confirmationModal');
const confirmClear = document.getElementById('confirmClear');
const cancelClear = document.getElementById('cancelClear');
const editLockBanner = document.getElementById('editLockBanner');

// State variables
let saveTimeout = null;
let isSaving = false;
let areaRows = [];
let agentRows = [];
let vehicleRows = [];
let serviceChangeRows = [];

// Initialize with default rows
document.addEventListener('DOMContentLoaded', () => {
    initFirebase();

    // Set today's date as default
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('operationDate').value = today;
    
    // Set current time as default for start time
    const now = new Date();
    const hours = now.getHours().toString().padStart(2, '0');
    const minutes = now.getMinutes().toString().padStart(2, '0');
    document.getElementById('startTime').value = `${hours}:${minutes}`;
    
    // Load saved data
    loadFormData();
    
    // Add initial rows if none loaded
    if (areaRows.length === 0) addAreaRow();
    if (agentRows.length === 0) addAgentRow();
    if (vehicleRows.length === 0) addVehicleRow();
    if (serviceChangeRows.length === 0) addServiceChangeRow();
    
    // Update time display
    updateLastSavedTime();
    
    // Initialize event listeners
    initializeEventListeners();
    initializeAuthListeners();

    // Auto-expand textareas after load/reload
    setTimeout(expandAllTextareas, 0);

    // Start idle watcher
    initializeIdleWatcher();
});

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
            showAppView();
            loadOperationsList();
        } else {
            showLoginView();
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
            triggerSave();
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
            triggerSave();
        });
    });
    
    // Add buttons
    addAreaBtn.addEventListener('click', addAreaRow);
    addAgentBtn.addEventListener('click', addAgentRow);
    addVehicleBtn.addEventListener('click', addVehicleRow);
    addServiceChangeBtn.addEventListener('click', addServiceChangeRow);
    
    // Save to Firebase explicitly
    saveDbBtn.addEventListener('click', async () => {
        if (!auth || !auth.currentUser) return;
        if (!validateRequiredFields()) {
            saveStatus.textContent = 'Preencha os campos obrigatórios antes de salvar.';
            return;
        }
        if (!isEditAllowed()) {
            saveStatus.textContent = 'Edição bloqueada para esta operação.';
            return;
        }
        saveStatus.textContent = 'Salvando no Firebase...';
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
    
    // Update total agents count when agents are added/removed
    document.addEventListener('input', () => {
        updateTotalAgentsCount();
    });

    // Revalidate edit window when date changes
    const operationDateInput = document.getElementById('operationDate');
    operationDateInput.addEventListener('change', enforceEditWindow);

    // App navigation
    currentOperationBtn.addEventListener('click', () => setActiveAppView('form'));
    newOperationBtn.addEventListener('click', () => {
        confirmNewOperation();
        setActiveAppView('form');
    });
    viewManagerBtn.addEventListener('click', () => setActiveAppView('manager'));

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
}

function validateRequiredFields() {
    const requiredIds = [
        'operationName',
        'operationDate',
        'startTime',
        'endTime',
        'location',
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
        currentOperationId = null;
        operationsCache = [];
    });
}

function showLoginView() {
    loginView.classList.remove('hidden');
    appView.classList.add('hidden');
    clearIdleTimer();
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

function setActiveAppView(view) {
    const isForm = view === 'form';
    operationFormView.classList.toggle('hidden', !isForm);
    managerView.classList.toggle('hidden', isForm);
    currentOperationBtn.classList.toggle('active', isForm);
    if (isForm) {
        newOperationBtn.classList.remove('active');
    }
    viewManagerBtn.classList.toggle('active', !isForm);

    if (!isForm) {
        loadOperationsList();
    } else {
        requestAnimationFrame(() => {
            requestAnimationFrame(expandAllTextareas);
        });
        enforceEditWindow();
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
    const dateValue = document.getElementById('operationDate').value;
    if (!dateValue) return true;

    const operationDate = new Date(dateValue + 'T00:00:00');
    if (Number.isNaN(operationDate.getTime())) return true;

    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const cutoff = new Date(operationDate.getTime() + 3 * 24 * 60 * 60 * 1000);
    return today.getTime() <= cutoff.getTime();
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
    addAgentBtn.disabled = !allowed;
    addVehicleBtn.disabled = !allowed;
    addServiceChangeBtn.disabled = !allowed;
    clearBtn.disabled = !allowed;
    saveDbBtn.disabled = !allowed;
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
    const total = agentRows.length;
    document.getElementById('totalAgents').textContent = total;
}

// AREA DISTRIBUTION FUNCTIONS
function addAreaRow() {
    const rowId = Date.now();
    const newRow = {
        id: rowId,
        area: '',
        agents: '',
        vehicle: ''
    };
    
    areaRows.push(newRow);
    updateAreaTable();
    updateAreaCards();
    triggerSave();
}

function removeAreaRow(rowId) {
    if (areaRows.length <= 1) {
        alert('Deve haver pelo menos uma área cadastrada.');
        return;
    }
    
    areaRows = areaRows.filter(row => row.id !== rowId);
    updateAreaTable();
    updateAreaCards();
    triggerSave();
}

function updateAreaTable() {
    areaTableBody.innerHTML = '';
    
    areaRows.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" class="area-input" data-id="${row.id}" data-field="area" value="${row.area}" placeholder="Ex: Centro Comercial"></td>
            <td><input type="number" class="area-input" data-id="${row.id}" data-field="agents" value="${row.agents}" min="0" placeholder="0"></td>
            <td><input type="text" class="area-input" data-id="${row.id}" data-field="vehicle" value="${row.vehicle}" placeholder="Ex: VTR-001"></td>
            <td><button class="btn-small btn-danger remove-area-row" data-id="${row.id}"><i class="fas fa-trash-alt"></i> Remover</button></td>
        `;
        areaTableBody.appendChild(tr);
    });
    
    document.querySelectorAll('.area-input').forEach(input => {
        input.addEventListener('input', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            const field = e.target.getAttribute('data-field');
            const value = e.target.value;
            
            const rowIndex = areaRows.findIndex(row => row.id === rowId);
            if (rowIndex !== -1) {
                areaRows[rowIndex][field] = value;
                triggerSave();
            }
        });
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
                <span class="card-label">Área/Local:</span>
                <span><input type="text" class="card-input" data-id="${row.id}" data-field="area" value="${row.area}" placeholder="Ex: Centro Comercial"></span>
            </div>
            <div class="card-row">
                <span class="card-label">Qtd. Agentes:</span>
                <span><input type="number" class="card-input" data-id="${row.id}" data-field="agents" value="${row.agents}" min="0" placeholder="0"></span>
            </div>
            <div class="card-row">
                <span class="card-label">Viatura:</span>
                <span><input type="text" class="card-input" data-id="${row.id}" data-field="vehicle" value="${row.vehicle}" placeholder="Ex: VTR-001"></span>
            </div>
            <div class="card-actions">
                <button class="btn-small btn-danger remove-area-row" data-id="${row.id}"><i class="fas fa-trash-alt"></i> Remover</button>
            </div>
        `;
        areaCardsContainer.appendChild(card);
    });
    
    document.querySelectorAll('.cards-container .card-input').forEach(input => {
        input.addEventListener('input', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            const field = e.target.getAttribute('data-field');
            const value = e.target.value;
            
            const rowIndex = areaRows.findIndex(row => row.id === rowId);
            if (rowIndex !== -1) {
                areaRows[rowIndex][field] = value;
                triggerSave();
            }
        });
    });
    
    document.querySelectorAll('.cards-container .remove-area-row').forEach(button => {
        button.addEventListener('click', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            removeAreaRow(rowId);
        });
    });
}

// AGENTS FUNCTIONS
function addAgentRow() {
    const rowId = Date.now();
    const newRow = {
        id: rowId,
        classe: '',
        warName: ''
    };
    
    agentRows.push(newRow);
    updateAgentsTable();
    updateAgentsCards();
    updateTotalAgentsCount();
    triggerSave();
}

function removeAgentRow(rowId) {
    if (agentRows.length <= 1) {
        alert('Deve haver pelo menos um agente cadastrado.');
        return;
    }
    
    agentRows = agentRows.filter(row => row.id !== rowId);
    updateAgentsTable();
    updateAgentsCards();
    updateTotalAgentsCount();
    triggerSave();
}

function updateAgentsTable() {
    agentsTableBody.innerHTML = '';
    
    agentRows.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" class="agent-input" data-id="${row.id}" data-field="classe" value="${row.classe}" placeholder="Ex: SUPERVISOR"></td>
            <td><input type="text" class="agent-input" data-id="${row.id}" data-field="warName" value="${row.warName}" placeholder="Nome de guerra"></td>
            <td><button class="btn-small btn-danger remove-agent-row" data-id="${row.id}"><i class="fas fa-trash-alt"></i> Remover</button></td>
        `;
        agentsTableBody.appendChild(tr);
    });
    
    document.querySelectorAll('.agent-input').forEach(input => {
        input.addEventListener('input', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            const field = e.target.getAttribute('data-field');
            const value = e.target.value;
            
            const rowIndex = agentRows.findIndex(row => row.id === rowId);
            if (rowIndex !== -1) {
                agentRows[rowIndex][field] = value;
                triggerSave();
            }
        });
    });
    
    document.querySelectorAll('.remove-agent-row').forEach(button => {
        button.addEventListener('click', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            removeAgentRow(rowId);
        });
    });
}

function updateAgentsCards() {
    agentsCardsContainer.innerHTML = '';
    
    agentRows.forEach(row => {
        const card = document.createElement('div');
        card.className = 'agent-card';
        card.innerHTML = `
            <div class="card-row">
                <span class="card-label">Classe:</span>
                <span><input type="text" class="card-input" data-id="${row.id}" data-field="classe" value="${row.classe}" placeholder="Ex: PM3"></span>
            </div>
            <div class="card-row">
                <span class="card-label">Nome de Guerra:</span>
                <span><input type="text" class="card-input" data-id="${row.id}" data-field="warName" value="${row.warName}" placeholder="Nome de guerra"></span>
            </div>
            <div class="card-actions">
                <button class="btn-small btn-danger remove-agent-row" data-id="${row.id}"><i class="fas fa-trash-alt"></i> Remover</button>
            </div>
        `;
        agentsCardsContainer.appendChild(card);
    });
    
    document.querySelectorAll('#agentsCardsContainer .card-input').forEach(input => {
        input.addEventListener('input', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            const field = e.target.getAttribute('data-field');
            const value = e.target.value;
            
            const rowIndex = agentRows.findIndex(row => row.id === rowId);
            if (rowIndex !== -1) {
                agentRows[rowIndex][field] = value;
                triggerSave();
            }
        });
    });
    
    document.querySelectorAll('#agentsCardsContainer .remove-agent-row').forEach(button => {
        button.addEventListener('click', (e) => {
            const rowId = parseInt(e.target.getAttribute('data-id'));
            removeAgentRow(rowId);
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
    triggerSave();
}

function removeVehicleRow(rowId) {
    if (vehicleRows.length <= 1) {
        alert('Deve haver pelo menos uma viatura cadastrada.');
        return;
    }
    
    vehicleRows = vehicleRows.filter(row => row.id !== rowId);
    updateVehiclesTable();
    updateVehiclesCards();
    triggerSave();
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
                    triggerSave();
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
                    triggerSave();
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
                <span><input type="text" class="card-input" data-id="${row.id}" data-field="post" value="${row.post}" placeholder="Ex: 1ª Cia"></span>
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
                    triggerSave();
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
                    triggerSave();
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
    triggerSave();
}

function removeServiceChangeRow(rowId) {
    if (serviceChangeRows.length <= 1) {
        return;
    }
    
    serviceChangeRows = serviceChangeRows.filter(row => row.id !== rowId);
    // Reindex items
    serviceChangeRows.forEach((row, index) => {
        row.index = index + 1;
    });
    updateServiceChangesContainer();
    triggerSave();
}

function updateServiceChangesContainer() {
    const container = document.getElementById('serviceChangesContainer');
    container.innerHTML = '';
    const isOnlyOne = serviceChangeRows.length === 1;
    
    serviceChangeRows.forEach(row => {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'service-change-item';
        itemDiv.innerHTML = `
            <div class="service-change-header">
                <h4>Alteração ${row.index}</h4>
                <button class="btn-small btn-danger remove-service-change-row" data-id="${row.id}" ${isOnlyOne ? 'disabled' : ''}>
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
                triggerSave();
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
}

function getFormData() {
    return {
        operationName: document.getElementById('operationName').value,
        operationDate: document.getElementById('operationDate').value,
        startTime: document.getElementById('startTime').value,
        endTime: document.getElementById('endTime').value,
        location: document.getElementById('location').value,
        coordinatorName: document.getElementById('coordinatorName').value,
        registration: document.getElementById('registration').value,
        position: document.getElementById('position').value,
        totalAgents: document.getElementById('totalAgents').textContent,
        incidents: document.getElementById('incidents').value,
        summary: document.getElementById('summary').value,
        areaRows: areaRows,
        agentRows: agentRows,
        vehicleRows: vehicleRows,
        serviceChangeRows: serviceChangeRows,
        lastSaved: new Date().toISOString()
    };
}

function applyFormData(formData) {
    document.getElementById('operationName').value = formData.operationName || '';
    document.getElementById('operationDate').value = formData.operationDate || '';
    document.getElementById('startTime').value = formData.startTime || '';
    document.getElementById('endTime').value = formData.endTime || '';
    document.getElementById('location').value = formData.location || '';
    document.getElementById('coordinatorName').value = formData.coordinatorName || '';
    document.getElementById('registration').value = formData.registration || '';
    document.getElementById('position').value = formData.position || '';
    document.getElementById('totalAgents').textContent = formData.totalAgents || '0';
    document.getElementById('incidents').value = formData.incidents || '';
    document.getElementById('summary').value = formData.summary || '';

    if (formData.areaRows && formData.areaRows.length > 0) {
        areaRows = formData.areaRows;
        updateAreaTable();
        updateAreaCards();
    }

    if (formData.agentRows && formData.agentRows.length > 0) {
        agentRows = formData.agentRows;
        updateAgentsTable();
        updateAgentsCards();
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
    }

    requestAnimationFrame(() => {
        requestAnimationFrame(expandAllTextareas);
    });

    enforceEditWindow();
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
        const userInfo = {
            uid: user.uid,
            email: user.email || null,
            displayName: user.displayName || null
        };
        const payload = {
            ...formData,
            ownerId: auth.currentUser.uid,
            updatedBy: userInfo,
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };

        if (currentOperationId) {
            docRef = operationsRef.doc(currentOperationId);
        } else {
            docRef = operationsRef.doc();
            currentOperationId = docRef.id;
            payload.createdAt = firebase.firestore.FieldValue.serverTimestamp();
            payload.createdBy = userInfo;
        }

        await docRef.set(payload, { merge: true });
        saveStatus.textContent = 'Salvo no Firebase';
        return docRef.id;
    } catch (error) {
        console.error('Erro ao salvar no Firebase:', error);
        saveStatus.textContent = 'Falha ao salvar no Firebase';
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
        const searchText = `${operation.operationName || ''} ${operation.coordinatorName || ''} ${operation.location || ''} ${operation.registration || ''}`.toLowerCase();
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
        const card = document.createElement('div');
        card.className = 'operation-card';
        card.innerHTML = `
            <div class="operation-title">${operation.operationName || 'Operação sem nome'}</div>
            <div class="operation-meta">
                <div><strong>Data:</strong> ${formatDate(operation.operationDate)}</div>
                <div><strong>Local:</strong> ${operation.location || 'Não informado'}</div>
                <div><strong>Coordenador:</strong> ${operation.coordinatorName || 'Não informado'}</div>
                <div><strong>Matrícula:</strong> ${operation.registration || 'Não informado'}</div>
            </div>
            <div class="operation-actions">
                <button class="btn-small" data-action="edit" data-id="${operation.id}">
                    <i class="fas fa-pen"></i> Editar
                </button>
                <button class="btn-small" data-action="pdf" data-id="${operation.id}">
                    <i class="fas fa-file-pdf"></i> Gerar PDF
                </button>
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
        button.addEventListener('click', () => {
            const id = button.getAttribute('data-id');
            const operation = operationsCache.find(item => item.id === id);
            if (operation) {
                applyFormData(operation);
                exportToPDF();
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
    // Clear all form inputs
    formInputs.forEach(input => {
        if (input.type === 'number') {
            input.value = '0';
        } else if (input.type !== 'date' && input.type !== 'time') {
            input.value = '';
        }
    });
    
    // Set today's date as default
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('operationDate').value = today;
    
    // Set current time as default for start time
    const now = new Date();
    const hours = now.getHours().toString().padStart(2, '0');
    const minutes = now.getMinutes().toString().padStart(2, '0');
    document.getElementById('startTime').value = `${hours}:${minutes}`;
    
    // Reset dynamic rows
    areaRows = [{
        id: Date.now(),
        area: '',
        agents: '',
        vehicle: ''
    }];
    
    agentRows = [{
        id: Date.now() + 1,
        classe: '',
        warName: ''
    }];
    
    vehicleRows = [{
        id: Date.now() + 2,
        prefix: '',
        post: '',
        type: ''
    }];
    
    serviceChangeRows = [{
        id: Date.now() + 3,
        description: '',
        index: 1
    }];
    
    // Update tables and cards
    updateAreaTable();
    updateAreaCards();
    updateAgentsTable();
    updateAgentsCards();
    updateVehiclesTable();
    updateVehiclesCards();
    updateServiceChangesContainer();
    
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
}

// Export to PDF with improved layout
function exportToPDF() {
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
    const operationName = document.getElementById('operationName').value || 'Não informado';
    const cleanOperationName = operationName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
    
    // Add main header
    doc.setFillColor(30, 58, 95);
    doc.rect(0, 0, pageWidth, 40, 'F');
    
    // Title in white
    setTextColor(255, 255, 255);
    setFontSize(22);
    setFontStyle('bold');
    doc.text('POLÍCIA MUNICIPAL DE ARACAJU', pageWidth / 2, 20, { align: 'center' });
    
    setFontSize(18);
    doc.text('RELATÓRIO DE OPERAÇÃO', pageWidth / 2, 30, { align: 'center' });
    
    // Add generation date
    const now = new Date();
    const dateStr = now.toLocaleDateString('pt-BR');
    const timeStr = now.toLocaleTimeString('pt-BR');
    setFontSize(10);
    doc.text(`Gerado em: ${dateStr} às ${timeStr}`, pageWidth - margin, 37, { align: 'right' });
    
    // Add decorative line
    doc.setDrawColor(201, 162, 39);
    doc.setLineWidth(2);
    doc.line(margin, 45, pageWidth - margin, 45);
    
    yPos = 55;
    
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
        checkPageBreak(15 + (data.length * 7)); // REDUZIDO: altura das linhas
        
        // Draw table header
        doc.setFillColor(30, 58, 95);
        doc.rect(margin, yPos, contentWidth, 7, 'F'); // REDUZIDO: altura do cabeçalho
        
        doc.setTextColor(255, 255, 255);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(10);
        
        let xPos = margin;
        for (let i = 0; i < headers.length; i++) {
            doc.text(headers[i], xPos + 2, yPos + 4.5); // REDUZIDO: posicionamento vertical
            xPos += colWidths[i];
        }
        
        yPos += 7; // REDUZIDO
        
        // Draw table rows
        doc.setTextColor(0, 0, 0);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(10);
        
        data.forEach((row, rowIndex) => {
            checkPageBreak(10); // REDUZIDO
            
            // Alternate row colors
            if (rowIndex % 2 === 0) {
                doc.setFillColor(245, 245, 245);
                doc.rect(margin, yPos, contentWidth, 7, 'F'); // REDUZIDO: altura da linha
            }
            
            // Add row content
            xPos = margin;
            for (let i = 0; i < row.length; i++) {
                const cellContent = row[i] || '-';
                const lines = doc.splitTextToSize(cellContent, colWidths[i] - 4);
                doc.text(lines[0] || '-', xPos + 2, yPos + 4.5); // REDUZIDO: posicionamento vertical
                xPos += colWidths[i];
            }
            
            yPos += 7; // REDUZIDO: altura da linha
        });
        
        yPos += 8; // REDUZIDO: espaço após tabela
        
        // Add total if requested
        if (showTotal && totalLabel) {
            doc.setFillColor(240, 240, 240);
            doc.rect(margin, yPos, contentWidth, 7, 'F'); // REDUZIDO
            
            doc.setTextColor(30, 58, 95);
            setFontStyle('bold');
            doc.text(totalLabel, margin + 5, yPos + 4.5); // REDUZIDO
            yPos += 10; // REDUZIDO
        }
        
        yPos += 12; // REDUZIDO
    }
    
    // 1. IDENTIFICAÇÃO DA OPERAÇÃO
    const operationDate = document.getElementById('operationDate').value || 'Não informado';
    const startTime = document.getElementById('startTime').value || 'Não informado';
    const endTime = document.getElementById('endTime').value || 'Não informado';
    const location = document.getElementById('location').value || 'Não informado';
    
    const identificationContent = `Operação: ${operationName}
Data: ${formatDate(operationDate)}
Período: ${startTime} às ${endTime}
Local: ${location}`;
    
    addSection('1. IDENTIFICAÇÃO DA OPERAÇÃO', identificationContent);
     // Adicione espaço extra aqui:
    yPos += 2;
    
    // 2. COORDENAÇÃO / RESPONSÁVEL
    const coordinatorName = document.getElementById('coordinatorName').value || 'Não informado';
    const registration = document.getElementById('registration').value || 'Não informado';
    const position = document.getElementById('position').value || 'Não informado';
    
    const coordinationContent = `Coordenador: ${coordinatorName}
Matrícula: ${registration}
Cargo/Função: ${position}`;
    
    addSection('2. COORDENAÇÃO / RESPONSÁVEL', coordinationContent);
 // Adicione espaço extra aqui:
    yPos += 2;
    
    // 3. EFETIVO EMPREGADO - Table
    checkPageBreak(15);
    addSection('3. EFETIVO EMPREGADO', '');

     // Adicione espaço extra aqui:
    yPos += 2;
    
    if (agentRows.length > 0) {
        const agentHeaders = ['CLASSE', 'NOME DE GUERRA'];
        const agentData = agentRows.map(row => [
            row.classe || '-',
            row.warName || '-'
        ]);
        
        const agentColWidths = [50, 100];
        createTable(agentHeaders, agentData, agentColWidths, true, `TOTAL DE AGENTES: ${agentRows.length}`);
    } else {
        setTextColor(100, 100, 100);
        doc.text('Nenhum agente cadastrado.', margin, yPos);
        yPos += 8;
    }
    
    // 4. VIATURAS (VTRs) UTILIZADAS - Table
    checkPageBreak(15);
    addSection('4. VIATURAS (VTRs) UTILIZADAS', '');

     // Adicione espaço extra aqui:
    yPos += 2;
    
    if (vehicleRows.length > 0) {
        const vehicleHeaders = ['PREFIXO', 'POSTO', 'TIPO'];
        const vehicleData = vehicleRows.map(row => [
            row.prefix || '-',
            row.post || '-',
            row.type || '-'
        ]);
        
        const vehicleColWidths = [40, 50, 60];
        createTable(vehicleHeaders, vehicleData, vehicleColWidths, true, `TOTAL DE VIATURAS: ${vehicleRows.length}`);
    } else {
        setTextColor(100, 100, 100);
        doc.text('Nenhuma viatura cadastrada.', margin, yPos);
        yPos += 8;
    }
    
    // 5. DISTRIBUIÇÃO DO EFETIVO POR ÁREA - Table
    checkPageBreak(25);
    addSection('5. DISTRIBUIÇÃO DO EFETIVO POR ÁREA', '');

     // Adicione espaço extra aqui:
    yPos += 2;
    
    if (areaRows.length > 0) {
        const areaHeaders = ['ÁREA/LOCAL', 'QTD. AGENTES', 'VIATURA'];
        const areaData = areaRows.map(row => [
            row.area || '-',
            row.agents || '0',
            row.vehicle || '-'
        ]);
        
        const areaColWidths = [70, 35, 45];
        
        // Calculate total agents distributed
        const totalAgentsDistributed = areaRows.reduce((sum, row) => sum + (parseInt(row.agents) || 0), 0);
        createTable(areaHeaders, areaData, areaColWidths, true, `TOTAL DE AGENTES DISTRIBUÍDOS: ${totalAgentsDistributed}`);
    } else {
        setTextColor(100, 100, 100);
        doc.text('Nenhuma área cadastrada.', margin, yPos);
        yPos += 8;
    }

    // 6. REGISTRO DE OCORRÊNCIAS
    const incidents = document.getElementById('incidents').value || 'Nenhuma ocorrência registrada durante a operação.';
     addSectionJustified('6. REGISTRO DE OCORRÊNCIAS', incidents);
    
     // Adicione espaço extra aqui:
    yPos += 2;

    // 7. ALTERAÇÕES DE SERVIÇO
    checkPageBreak(8);
    addSection('7. ALTERAÇÕES DE SERVIÇO', '');

    // Adicione espaço extra aqui:
    yPos += 2;
    
    // Verificar se há alterações de serviço com conteúdo
    const hasServiceChanges = serviceChangeRows.some(row => row.description && row.description.trim() !== '');
    
    if (hasServiceChanges) {
        setTextColor(0, 0, 0);
        setFontSize(11);
        setFontStyle('normal');
        
        serviceChangeRows.forEach((row, index) => {
            // Verificar se esta alteração tem conteúdo
            const description = row.description && row.description.trim() !== '' ? row.description : null;
            
            if (description) {
                // Check page break for title only
                checkPageBreak(8);
                
                // Add alteration title
                setTextColor(30, 58, 95);
                setFontStyle('bold');
                doc.text(`Alteração ${row.index}:`, margin, yPos);
                yPos += 8;
                
                // Add alteration description com texto justificado
                setTextColor(0, 0, 0);
                setFontStyle('normal');
                const lines = doc.splitTextToSize(description, contentWidth - 2);
                lines.forEach((line, lineIndex) => {
                    // Check page break for each line
                    checkPageBreak(6);
                    
                    // Render justified text for all lines except the last
                    if (lineIndex < lines.length - 1) {
                        renderJustifiedText(line, margin + 1, yPos, contentWidth - 2);
                    } else {
                        // Last line: render left-aligned
                        doc.text(line, margin + 1, yPos);
                    }
                    yPos += 7.5;
                });
                
                yPos += 7.5; // Espaço entre alterações
            }
        });
    } else {
        setTextColor(100, 100, 100);
        doc.text('Nenhuma alteração de serviço registrada.', margin, yPos);
        yPos += 8;
    }
     // Adicione espaço extra aqui:
    yPos += 2;

    // 8. RESUMO DA OPERAÇÃO
    const summary = document.getElementById('summary').value || 'Nenhum resumo fornecido.';
    addSectionJustified('8. RESUMO DA OPERAÇÃO', summary);

     // Adicione espaço extra aqui:
    yPos += 2;
    
    // Add signature area com menos espaçamento
    checkPageBreak(35);
    yPos += 6;
    
    doc.setDrawColor(150, 150, 150);
    doc.line(margin, yPos, pageWidth - margin, yPos);
    yPos += 15;
    
    setTextColor(30, 58, 95);
    setFontSize(12);
    setFontStyle('bold');
    doc.text('RESPONSÁVEL PELA OPERAÇÃO', pageWidth / 2, yPos, { align: 'center' });
    
    yPos += 12;
    
    // Signature lines
    const lineLength = 80;
    const leftLineX = pageWidth / 2 - lineLength - 10;
    const rightLineX = pageWidth / 2 + 10;
    
    doc.setDrawColor(200, 200, 200);
    doc.line(leftLineX, yPos, leftLineX + lineLength, yPos);
    doc.line(rightLineX, yPos, rightLineX + lineLength, yPos);
    
    yPos += 8;
    setTextColor(100, 100, 100);
    setFontSize(9);
    setFontStyle('normal');
    doc.text('Assinatura', leftLineX + (lineLength / 2), yPos, { align: 'center' });
    doc.text('Carimbo/Matrícula', rightLineX + (lineLength / 2), yPos, { align: 'center' });
    
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
        doc.text('Polícia Municipal de Aracaju - Formulário de Gestão de Operações - © 2026', 
               pageWidth / 2, pageHeight - 10, { align: 'center' });
    }
    
    const fileName = `Relatorio_Operacao_${cleanOperationName}_${dateStr.replace(/\//g, '-')}.pdf`;

    // Save the PDF with cleaned filename
    doc.save(fileName);
}

// Helper function to format date
function formatDate(dateString) {
    if (!dateString) return 'Não informado';
    const date = new Date(dateString);
    return date.toLocaleDateString('pt-BR');
}