// ==================== CONFIGURATION ====================
const CONFIG = {
    PROXY_URL: 'https://script.google.com/macros/s/AKfycbwC9lJdYj4askujDNO2GfK-Rqq02VBcr90NXhifgvpawboEK1YCyUfbi2GA2hFL2UghkA/exec',
    AUTH_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbwC9lJdYj4askujDNO2GfK-Rqq02VBcr90NXhifgvpawboEK1YCyUfbi2GA2hFL2UghkA/exec',
    LOGO_URL: 'https://github.com/mohamedsillahkanu/gdp-dashboard-2/raw/6c7463b0d5c3be150aafae695a4bcbbd8aeb1499/ICF-SL.jpg'
};

// ==================== STORAGE HELPERS ====================
// In-memory fallback storage for when localStorage is blocked
const memoryStorage = {};

// Check if localStorage is available
let storageAvailable = true;
let storageWarningShown = false;
try {
    localStorage.setItem('__test__', '1');
    localStorage.removeItem('__test__');
} catch (e) {
    storageAvailable = false;
    console.warn('localStorage not available:', e.message);
}

// Safe localStorage wrapper with in-memory fallback
const safeStorage = {
    getItem: function(key) {
        if (!storageAvailable) {
            return memoryStorage[key] || null;
        }
        try {
            return localStorage.getItem(key);
        } catch (e) {
            console.warn('Storage read error:', e);
            return memoryStorage[key] || null;
        }
    },
    setItem: function(key, value) {
        // Always store in memory as backup
        memoryStorage[key] = value;
        
        if (!storageAvailable) {
            if (!storageWarningShown) {
                storageWarningShown = true;
                setTimeout(() => {
                    if (typeof notify === 'function') {
                        notify('⚠️ Storage blocked by browser. Data stored in memory only - will be lost on refresh. Host file on a web server for persistence.', 'warning');
                    }
                }, 2000);
            }
            return;
        }
        try {
            localStorage.setItem(key, value);
        } catch (e) {
            console.warn('Storage write error:', e);
        }
    },
    removeItem: function(key) {
        delete memoryStorage[key];
        if (!storageAvailable) return;
        try {
            localStorage.removeItem(key);
        } catch (e) {
            console.warn('Storage remove error:', e);
        }
    }
};

// ==================== PROXY CONFIGURATION ====================
// Multiple proxies for DHIS2 connection fallback
const PROXIES = [
    { 
        name: 'Direct',
        url: '',
        type: 'direct'  // Try direct connection first (works if DHIS2 has CORS enabled)
    },
    { 
        name: 'GAS Proxy',
        url: 'https://script.google.com/macros/s/AKfycbwC9lJdYj4askujDNO2GfK-Rqq02VBcr90NXhifgvpawboEK1YCyUfbi2GA2hFL2UghkA/exec',
        type: 'gas'
    }
  
];

function buildProxyUrl(proxy, targetUrl, auth, method) {
    if (proxy.type === 'direct') {
        return targetUrl; // Direct connection
    } else if (proxy.type === 'gas') {
        let proxyUrl = proxy.url + '?url=' + encodeURIComponent(targetUrl);
        if (auth) {
            proxyUrl += '&auth=' + encodeURIComponent(auth);
        }
        if (method && method !== 'GET') {
            proxyUrl += '&method=' + method;
        }
        return proxyUrl;
    } else {
        return proxy.url + encodeURIComponent(targetUrl);
    }
}

// ==================== GLOBAL STATE ====================
const state = {
    user: null,
    fields: [],
    selectedFieldId: null,
    fieldCounter: 0,
    isSharedMode: false,
    settings: {
        title: 'My Data Collection Form',
        originalTitle: '', // Track original title for rename detection
        formId: '', // Unique form identifier
        logo: CONFIG.LOGO_URL,
        aggregateColumn: '', // Legacy single column (for backward compatibility)
        aggregateColumns: [], // Multiple columns for grouping in aggregate view
        gpsField: '' // Which GPS field to use for mapping
    },
    sheets: {
        scriptUrl: CONFIG.AUTH_SCRIPT_URL,
        sheetId: '',
        connected: true
    },
    dhis2: {
        url: '',
        username: '',
        password: '',
        syncMode: 'aggregate', // 'aggregate' or 'tracker'
        orgUnitLevel: 5,
        periodType: 'Monthly',
        periodColumn: '',
        orgUnitColumn: '',
        programId: '',
        connected: false,
        datasetId: null,
        dataElements: {},
        orgUnits: [],
        orgUnitMap: {}
    },
    collectedData: [],
    filters: {},
    filterOrder: [], // Custom order for filter fields
    dateFilter: { start: '', end: '' },
    currentDataView: 'case', // 'case' or 'aggregate'
    chartInstances: {}
};

const fieldDefs = {
    period: { label: 'Period', icon: 'calendar-days', defaultLabel: 'Reporting Period', valueType: 'TEXT', isDhis2: true, category: 'dhis2' },
    text: { label: 'Text', icon: 'type', defaultLabel: 'Text Field', valueType: 'TEXT', category: 'text' },
    number: { label: 'Number', icon: 'hash', defaultLabel: 'Number Field', valueType: 'NUMBER', category: 'numeric' },
    calculation: { label: 'Calculation', icon: 'calculator', defaultLabel: 'Calculated Field', valueType: 'NUMBER', category: 'numeric' },
    date: { label: 'Date', icon: 'calendar', defaultLabel: 'Date Field', valueType: 'DATE', category: 'date' },
    time: { label: 'Time', icon: 'clock', defaultLabel: 'Time Field', valueType: 'TIME', category: 'time' },
    email: { label: 'Email', icon: 'mail', defaultLabel: 'Email Address', valueType: 'TEXT', category: 'text' },
    phone: { label: 'Phone', icon: 'phone', defaultLabel: 'Phone Number', valueType: 'TEXT', category: 'text' },
    textarea: { label: 'Long Text', icon: 'align-left', defaultLabel: 'Long Text', valueType: 'LONG_TEXT', category: 'text' },
    select: { label: 'Dropdown', icon: 'chevron-down-square', defaultLabel: 'Dropdown', valueType: 'TEXT', category: 'categorical', options: ['Option 1', 'Option 2', 'Option 3'] },
    radio: { label: 'Radio', icon: 'circle-dot', defaultLabel: 'Radio Choice', valueType: 'TEXT', category: 'categorical', options: ['Option 1', 'Option 2', 'Option 3'] },
    checkbox: { label: 'Checkbox', icon: 'check-square', defaultLabel: 'Checkbox', valueType: 'TRUE_ONLY', category: 'categorical', options: ['Option 1', 'Option 2'] },
    yesno: { label: 'Yes/No', icon: 'toggle-left', defaultLabel: 'Yes/No Question', valueType: 'BOOLEAN', category: 'categorical' },
    gps: { label: 'GPS', icon: 'map-pin', defaultLabel: 'GPS Location', valueType: 'TEXT', category: 'text' },
    qrcode: { label: 'QR Code', icon: 'qr-code', defaultLabel: 'QR/Barcode Scanner', valueType: 'TEXT', category: 'text' },
    cascade: { label: 'Cascade', icon: 'git-branch', defaultLabel: 'Cascading Dropdown', valueType: 'TEXT', category: 'categorical', cascadeData: [], cascadeColumns: [] },
    rating: { label: 'Rating', icon: 'star', defaultLabel: 'Rating', valueType: 'INTEGER', category: 'categorical', max: 5 },
    section: { label: 'Section', icon: 'folder', defaultLabel: 'New Section', valueType: null, category: 'structure' }
};

// ==================== PERIODS ====================
function generatePeriods() {
    const periods = [];
    for (let year = 2020; year <= 2026; year++) {
        for (let month = 1; month <= 12; month++) {
            periods.push(`${year}${String(month).padStart(2, '0')}`);
        }
    }
    return periods;
}
const PERIODS = generatePeriods();

// ==================== INITIALIZATION ====================
function init() {
    // Show warning if running locally without a server
    if (window.location.protocol === 'file:') {
        const warning = document.createElement('div');
        warning.style.cssText = 'background:#ffc107;color:#000;padding:10px 15px;text-align:center;font-size:13px;font-family:Oswald,sans-serif;position:fixed;top:0;left:0;right:0;z-index:10000;';
        warning.innerHTML = '<strong>⚠️ Local File Mode:</strong> Some features may not work. For full functionality, host this file on a web server (e.g., GitHub Pages, localhost). <button onclick="this.parentElement.remove()" style="margin-left:10px;padding:2px 8px;cursor:pointer;">✕</button>';
        document.body.insertBefore(warning, document.body.firstChild);
    }
    
    // Check storage availability
    if (!storageAvailable) {
        console.warn('localStorage is not available - data will not persist');
    }
    
    const urlParams = new URLSearchParams(window.location.search);
    const compressedData = urlParams.get('d');
    
    if (compressedData) {
        try {
            const decoded = atob(decodeURIComponent(compressedData));
            const charData = decoded.split('').map(c => c.charCodeAt(0));
            const binData = new Uint8Array(charData);
            const decompressed = pako.inflate(binData, { to: 'string' });
            const data = JSON.parse(decompressed);
            
            // Cache the form data for offline use
            const cacheKey = 'shared_form_' + window.location.search.substring(0, 100);
            try {
                safeStorage.setItem(cacheKey, decompressed);
                console.log('Shared form cached for offline use');
            } catch (e) {
                console.log('Could not cache form data');
            }
            
            state.isSharedMode = true;
            renderSharedForm(data);
            return;
        } catch (err) { 
            console.error('Decode error:', err);
            
            // Try to load from cache if offline
            if (!navigator.onLine) {
                const cacheKey = 'shared_form_' + window.location.search.substring(0, 100);
                const cached = safeStorage.getItem(cacheKey);
                if (cached) {
                    try {
                        const data = JSON.parse(cached);
                        console.log('Loaded form from offline cache');
                        state.isSharedMode = true;
                        renderSharedForm(data);
                        return;
                    } catch (e) {
                        console.error('Cache parse error:', e);
                    }
                }
            }
        }
    }
    
    loadFromStorage();
    loadConfigs();
    setupEventListeners();
    checkAuth();
}

function setupEventListeners() {
    document.querySelectorAll('.auth-tab').forEach(tab => tab.addEventListener('click', () => switchAuthTab(tab.dataset.tab)));
    document.getElementById('loginForm').addEventListener('submit', handleLogin);
    document.getElementById('signupForm').addEventListener('submit', handleSignup);
    document.getElementById('forgotForm').addEventListener('submit', handleForgotPassword);
    document.querySelectorAll('.field-type').forEach(el => el.addEventListener('click', () => addField(el.dataset.type)));
    document.getElementById('formTitle').addEventListener('input', (e) => {
        state.settings.title = e.target.value;
        document.getElementById('previewTitle').textContent = e.target.value;
        saveToStorage();
    });
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
        overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.classList.remove('show'); });
    });
}

// ==================== GOOGLE SHEETS ====================
// Automatic sheet creation based on form name
// Row 1 = Labels, Row 2 = Variable Names, Row 3+ = Data

async function submitToGoogleSheets(data) {
    const scriptUrl = CONFIG.AUTH_SCRIPT_URL;
    if (!scriptUrl) {
        console.log('No script URL configured');
        saveOffline(data);
        return { success: false, offline: true };
    }
    
    // Check if online
    if (!navigator.onLine) {
        console.log('Browser reports offline');
        saveOffline(data);
        return { success: false, offline: true };
    }
    
    console.log('Submitting to Google Sheets...', { formTitle: state.settings.title });
    
    // Build field definitions (labels and names)
    const fieldDefs = state.fields
        .filter(f => f.type !== 'section')
        .map(f => ({ label: f.label, name: f.name }));
    
    // Add system fields
    fieldDefs.unshift(
        { label: 'Record ID', name: '_id' },
        { label: 'Timestamp', name: '_timestamp' },
        { label: 'Synced', name: '_synced' }
    );
    
    try {
        // Determine if title was changed
        const oldFormTitle = state.settings.originalTitle && 
                             state.settings.originalTitle !== state.settings.title 
                             ? state.settings.originalTitle : null;
        
        const payload = {
            action: 'submit',
            formTitle: state.settings.title,
            formId: state.settings.formId || null,
            oldFormTitle: oldFormTitle,
            syncSchema: true, // Sync columns with form definition
            fields: fieldDefs,
            data: data
        };
        
        // Update original title after successful submit
        state.settings.originalTitle = state.settings.title;
        
        console.log('Payload:', payload);
        
        const response = await fetch(scriptUrl, {
            method: 'POST', mode: 'cors', redirect: 'follow',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify(payload)
        });
        
        console.log('Response status:', response.status);
        
        const text = await response.text();
        console.log('Response text:', text);
        
        let result;
        try {
            result = JSON.parse(text);
        } catch(e) {
            console.error('Failed to parse response:', text);
            saveOffline(data);
            return { success: false, offline: true, error: 'Invalid response from server' };
        }
        
        console.log('Parsed result:', result);
        
        if (result.success) {
            // Mark as synced
            data._synced = true;
            data._syncedAt = new Date().toISOString();
            console.log('✓ Submitted successfully');
            return { success: true, offline: false };
        } else {
            console.error('Sheets submission failed:', result.error);
            saveOffline(data);
            return { success: false, offline: true, error: result.error };
        }
    } catch (err) {
        console.error('Sheets submission error:', err);
        saveOffline(data);
        return { success: false, offline: true, error: err.message };
    }
}

async function loadFromGoogleSheets() {
    console.log('=== loadFromGoogleSheets called ===');
    const scriptUrl = CONFIG.AUTH_SCRIPT_URL;
    if (!scriptUrl) {
        console.log('No script URL configured');
        return [];
    }
    
    const formTitle = state.settings.title;
    console.log('Form title:', formTitle);
    
    if (!formTitle) {
        console.log('No form title set');
        return [];
    }
    
    const url = `${scriptUrl}?action=getData&formTitle=${encodeURIComponent(formTitle)}&limit=1000`;
    console.log('Fetching from:', url);
    
    try {
        const response = await fetch(url);
        console.log('Response status:', response.status);
        
        if (!response.ok) {
            console.error('HTTP error:', response.status, response.statusText);
            return [];
        }
        
        const text = await response.text();
        console.log('Response text (first 500 chars):', text.substring(0, 500));
        
        if (!text || text.trim() === '') {
            console.log('Empty response from server');
            return [];
        }
        
        let result;
        try {
            result = JSON.parse(text);
        } catch (parseErr) {
            console.error('JSON parse error:', parseErr);
            console.error('Response was:', text.substring(0, 200));
            return [];
        }
        
        if (result.success && result.data) {
            console.log('✓ Loaded', result.data.length, 'records from Google Sheets');
            return result.data;
        } else {
            console.log('No data or error:', result.error || result.message || 'Unknown');
            return [];
        }
    } catch (err) {
        console.error('Load error:', err.name, err.message);
        return [];
    }
}

function saveOffline(data) {
    data._offline = true;
    
    // Include field definitions for syncing later
    const fieldDefs = state.fields
        .filter(f => f.type !== 'section')
        .map(f => ({ label: f.label, name: f.name }));
    
    fieldDefs.unshift(
        { label: 'Record ID', name: '_id' },
        { label: 'Timestamp', name: '_timestamp' },
        { label: 'Synced', name: '_synced' }
    );
    
    const offlineData = JSON.parse(safeStorage.getItem('icfOfflineData') || '[]');
    offlineData.push({ 
        formTitle: state.settings.title, 
        fields: fieldDefs,
        data: data 
    });
    safeStorage.setItem('icfOfflineData', JSON.stringify(offlineData));
}

async function syncOfflineData() {
    if (!navigator.onLine) {
        notify('No internet connection', 'error');
        return;
    }
    
    const offlineData = JSON.parse(safeStorage.getItem('icfOfflineData') || '[]');
    if (offlineData.length === 0) {
        notify('No offline data to sync', 'info');
        return;
    }
    
    notify(`Syncing ${offlineData.length} offline records...`, 'info');
    
    let synced = 0;
    let failed = 0;
    const remaining = [];
    
    for (const item of offlineData) {
        try {
            const response = await fetch(CONFIG.AUTH_SCRIPT_URL, {
                method: 'POST', mode: 'cors', redirect: 'follow',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({
                    action: 'submit',
                    formTitle: item.formTitle,
                    fields: item.fields,
                    data: item.data
                })
            });
            
            const result = await response.json();
            
            if (result.success) {
                synced++;
                // Update local data
                const localRecord = state.collectedData.find(d => d._id === item.data._id);
                if (localRecord) {
                    localRecord._synced = true;
                    localRecord._syncedAt = new Date().toISOString();
                    delete localRecord._offline;
                }
            } else {
                failed++;
                remaining.push(item);
            }
        } catch (err) {
            failed++;
            remaining.push(item);
        }
    }
    
    // Save remaining offline data
    safeStorage.setItem('icfOfflineData', JSON.stringify(remaining));
    saveToStorage();
    
    if (synced > 0) {
        notify(`Synced ${synced} records to Google Sheets!`, 'success');
    }
    if (failed > 0) {
        notify(`${failed} records failed to sync`, 'error');
    }
    
    loadViewerData();
}

// Auto-sync when coming online
window.addEventListener('online', () => {
    const offlineData = JSON.parse(safeStorage.getItem('icfOfflineData') || '[]');
    if (offlineData.length > 0) {
        notify('Back online! Syncing offline data...', 'info');
        setTimeout(syncOfflineData, 1000);
    }
});

function getOfflineCount() {
    const offlineData = JSON.parse(safeStorage.getItem('icfOfflineData') || '[]');
    return offlineData.length;
}

function updateConnectionStatus() {
    const statusEl = document.getElementById('connectionStatus');
    if (!statusEl) return;
    
    const offlineCount = getOfflineCount();
    const isOnline = navigator.onLine;
    
    statusEl.className = 'connection-status ' + (isOnline ? 'online' : 'offline');
    
    if (isOnline) {
        if (offlineCount > 0) {
            statusEl.innerHTML = `<span class="inline-icon" style="color:#ffc107;">${getIcon('wifi', 14)}</span> Online (${offlineCount} pending)`;
            statusEl.style.background = '#ffc107';
            statusEl.style.color = '#000';
        } else {
            statusEl.innerHTML = `<span class="inline-icon">${getIcon('wifi', 14)}</span> Online`;
            statusEl.style.background = '#28a745';
            statusEl.style.color = '#fff';
        }
    } else {
        statusEl.innerHTML = `<span class="inline-icon">${getIcon('wifi-off', 14)}</span> Offline`;
        statusEl.style.background = '#dc3545';
        statusEl.style.color = '#fff';
    }
}

// Update connection status on network changes
window.addEventListener('online', updateConnectionStatus);
window.addEventListener('offline', updateConnectionStatus);


// ==================== DHIS2 ====================
async function dhis2Request(endpoint, method = 'GET', payload = null) {
    const { url, username, password } = state.dhis2;
    if (!url || !username || !password) {
        throw new Error('DHIS2 not configured');
    }

    const serverUrl = url.replace(/\/$/, '');
    const apiUrl = `${serverUrl}/api/${endpoint}`;
    const auth = btoa(`${username}:${password}`);
    
    let lastError = null;
    
    // Try each proxy in order
    for (let i = 0; i < PROXIES.length; i++) {
        const proxy = PROXIES[i];
        const proxiedUrl = buildProxyUrl(proxy, apiUrl, auth, method);
        
        let options;
        
        if (proxy.type === 'gas') {
            // GAS proxy: NO custom headers to avoid CORS preflight
            // Auth is already in the query string via buildProxyUrl
            options = {
                method: method === 'GET' ? 'GET' : 'POST',
                redirect: 'follow'
            };
            if (payload && method !== 'GET') {
                options.headers = { 'Content-Type': 'text/plain' };
                options.body = JSON.stringify(payload);
            }
        } else {
            // Direct/CORS: send full headers
            options = {
                method: method,
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                }
            };
            if (proxy.type === 'direct' || proxy.type === 'cors') {
                options.headers['Authorization'] = 'Basic ' + auth;
            }
            if (payload && method !== 'GET') {
                options.body = JSON.stringify(payload);
            }
        }
        
        try {
            const response = await fetch(proxiedUrl, options);
            const text = await response.text();
            
            let data;
            try { data = JSON.parse(text); } catch (e) { data = text; }
            
            // Check for proxy errors - try next proxy
            if (data && data.error) {
                if (typeof data.error === 'object' && (data.error.status === 502 || data.error.status === 500)) {
                    lastError = `${proxy.name}: ${data.error.message || 'Proxy error'}`;
                    continue;
                }
                if (typeof data.error === 'string' && (data.error.includes('Exception') || data.error.includes('fetch'))) {
                    lastError = `${proxy.name}: ${data.error}`;
                    continue;
                }
            }
            
            // Check for CORS error (direct connection blocked)
            if (proxy.type === 'direct' && typeof data === 'string' && data === '') {
                lastError = 'Direct: CORS blocked';
                continue;
            }
            
            // Check for DHIS2 error responses
            if (data && (data.httpStatus === 'Unauthorized' || data.httpStatusCode === 401)) {
                return { success: false, status: 401, error: 'Unauthorized - check credentials', data: data };
            }
            if (data && (data.httpStatus === 'Not Found' || data.httpStatusCode === 404)) {
                return { success: false, status: 404, error: 'Not found', data: data };
            }
            if (data && data.httpStatusCode && data.httpStatusCode >= 400) {
                return { success: false, status: data.httpStatusCode, error: data.message || 'Request failed', data: data };
            }


                    
            // Success check - include 201 Created and other success patterns
            const isSuccess = data && (
                data.httpStatus === 'Created' ||
                data.httpStatus === 'OK' ||
                data.httpStatusCode === 200 ||
                data.httpStatusCode === 201 ||
                data.response?.uid ||
                data.status === 'OK' ||
                data.status === 'SUCCESS' ||
                (typeof data === 'object' && !data.error && !data.httpStatus) ||
                Array.isArray(data) ||
                (data.organisationUnits !== undefined) ||
                (data.dataSets !== undefined) ||
                (data.dataElements !== undefined) ||
                (data.programs !== undefined) ||
                (data.systemName !== undefined)
            );
            
            if (isSuccess) {
                return { success: true, status: data.httpStatusCode || 200, data: data };
            }
            
            // If we got data but it's not clearly success, still return it (let caller decide)
            if (data && typeof data === 'object') {
                return { success: true, status: 200, data: data };
            }
            
            // Not successful, try next proxy
            lastError = `${proxy.name}: Invalid response`;
            
        } catch (error) {
            // Network error - try next proxy
            lastError = `${proxy.name}: ${error.message}`;
            continue;
        }
    }
    
    // All proxies failed
    return { success: false, status: 0, error: 'All proxies failed: ' + lastError };
}

async function testDhis2Connection() {
    const url = document.getElementById('dhis2Url').value.trim();
    const username = document.getElementById('dhis2Username').value.trim();
    const password = document.getElementById('dhis2Password').value;
    
    if (!url || !username || !password) {
        notify('Please fill server URL, username and password', 'error');
        return;
    }
    
    state.dhis2.url = url;
    state.dhis2.username = username;
    state.dhis2.password = password;
    
    updateDhis2Status('syncing', 'Testing connection...');
    addLog('info', 'Connecting to DHIS2...');
    
    try {
        const result = await dhis2Request('system/info.json');
        
        if (result.success && result.data?.systemName) {
            state.dhis2.connected = true;
            updateDhis2Status('connected', `Connected to ${result.data.systemName}`);
            addLog('success', `✓ Connected: ${result.data.systemName} v${result.data.version || ''}`);
            
            if (state.dhis2.syncMode === 'aggregate') {
                await fetchOrgUnits();
            }
            
            saveDhis2Config();
            notify('Connected to DHIS2!');
        } else {
            state.dhis2.connected = false;
            updateDhis2Status('disconnected', 'Connection failed');
            addLog('error', `✗ Failed: ${result.error || 'Could not connect'}`);
            notify('Connection failed. Check URL and credentials.', 'error');
        }
    } catch (error) {
        state.dhis2.connected = false;
        updateDhis2Status('disconnected', 'Connection error');
        addLog('error', `✗ Error: ${error.message}`);
        notify('Connection error: ' + error.message, 'error');
    }
}

async function fetchOrgUnits() {
    const level = state.dhis2.orgUnitLevel;
    addLog('info', `Fetching org units at level ${level}...`);
    
    const result = await dhis2Request(`organisationUnits.json?fields=id,name,displayName,code&filter=level:eq:${level}&paging=false`);
    
    if (result.success && result.data?.organisationUnits) {
        state.dhis2.orgUnits = result.data.organisationUnits;
        state.dhis2.orgUnitMap = {};
        
        result.data.organisationUnits.forEach(ou => {
            const names = [
                ou.displayName?.toLowerCase().trim(),
                ou.name?.toLowerCase().trim(),
                ou.code?.toLowerCase().trim()
            ].filter(Boolean);
            
            names.forEach(name => {
                state.dhis2.orgUnitMap[name] = ou.id;
            });
        });
        
        addLog('success', `✓ Loaded ${state.dhis2.orgUnits.length} org units at level ${level}`);
    } else {
        addLog('warning', `⚠ Could not load org units: ${result.error || 'Unknown error'}`);
    }
}

function selectSyncMode(mode) {
    state.dhis2.syncMode = mode;
    document.querySelectorAll('.sync-mode-option').forEach(el => {
        el.classList.remove('selected');
        if (el.dataset.mode === mode) el.classList.add('selected');
    });
    
    document.getElementById('aggregateConfig').style.display = mode === 'aggregate' ? 'block' : 'none';
    document.getElementById('trackerConfig').style.display = mode === 'tracker' ? 'block' : 'none';
}

function saveDhis2Config() {
    state.dhis2.url = document.getElementById('dhis2Url').value.trim();
    state.dhis2.username = document.getElementById('dhis2Username').value.trim();
    state.dhis2.password = document.getElementById('dhis2Password').value;
    state.dhis2.orgUnitLevel = parseInt(document.getElementById('dhis2OrgLevel').value);
    state.dhis2.periodType = document.getElementById('dhis2PeriodType').value;
    state.dhis2.programId = document.getElementById('dhis2ProgramId').value.trim();
    state.dhis2.periodColumn = document.getElementById('dhis2PeriodColumn').value;
    state.dhis2.orgUnitColumn = document.getElementById('dhis2OrgUnitColumn').value;
    
    safeStorage.setItem('icfDhis2Config', JSON.stringify({
        url: state.dhis2.url,
        username: state.dhis2.username,
        password: state.dhis2.password,
        syncMode: state.dhis2.syncMode,
        orgUnitLevel: state.dhis2.orgUnitLevel,
        periodType: state.dhis2.periodType,
        periodColumn: state.dhis2.periodColumn,
        orgUnitColumn: state.dhis2.orgUnitColumn,
        programId: state.dhis2.programId
    }));
    
    notify('DHIS2 config saved!');
}

function updateDhis2Status(type, message) {
    const el = document.getElementById('dhis2Status');
    if (!el) return; // Safe for shared mode where element doesn't exist
    el.className = `status-badge ${type}`;
    const icon = type === 'connected' ? '✓' : type === 'syncing' ? '⋯' : '✗';
    el.innerHTML = `${icon} <span>${message}</span>`;
}

function addLog(type, message) {
    const log = document.getElementById('syncLog');
    if (!log) return;
    const time = new Date().toLocaleTimeString();
    log.innerHTML += `<div class="log-entry ${type}">[${time}] ${message}</div>`;
    log.scrollTop = log.scrollHeight;
}

function clearLog() {
    const log = document.getElementById('syncLog');
    if (log) log.innerHTML = '';
}

async function setupDhis2() {
    if (!state.dhis2.connected) {
        notify('Connect to DHIS2 first', 'error');
        return;
    }
    
    const periodColumn = state.dhis2.periodColumn;
    const orgUnitColumn = state.dhis2.orgUnitColumn;
    
    clearLog();
    addLog('info', `Setting up DHIS2 (${state.dhis2.syncMode} mode)...`);
    addLog('info', `Period column: ${periodColumn || 'Not set'}`);
    addLog('info', `Org Unit column: ${orgUnitColumn || 'Not set'} (will be used for assignment, not as data element)`);
    updateDhis2Status('syncing', 'Setting up...');
    
    // Skip fields: sections, period column, org unit column
    const skipFieldNames = [periodColumn, orgUnitColumn].filter(Boolean);
    
    const dataFields = state.fields.filter(f => {
        if (f.type === 'section') return false;
        if (skipFieldNames.includes(f.name)) return false;
        return true;
    });
    
    if (dataFields.length === 0) {
        notify('Add data fields first', 'error');
        return;
    }
    
    addLog('info', `Data fields to create: ${dataFields.map(f => f.label).join(', ')}`);
    
    // Ensure org units are loaded
    if (state.dhis2.orgUnits.length === 0) {
        await fetchOrgUnits();
    }
    
    try {
        // Create Data Elements
        addLog('info', 'Creating Data Elements...');
        const createdElements = [];
        const categoricalTypes = ['select', 'radio', 'yesno', 'checkbox'];
        
        // Field types that cannot be used in aggregate mode (text-based)
        const textBasedTypes = ['phone', 'gps', 'email', 'text', 'textarea'];
        
        if (state.dhis2.syncMode === 'tracker') {
            // ========== TRACKER MODE ==========
            // Create ONE data element per field, all as TEXT
            for (const field of dataFields) {
                const code = field.name.toUpperCase().replace(/[^A-Z0-9_]/g, '_');
                const cleanLabel = field.label.replace(/[<>]/g, ''); // Remove special chars
                
                // First check if exists by code
                let found = false;
                const checkResult = await dhis2Request(`dataElements.json?filter=code:eq:${code}&fields=id,name`);
                
                if (checkResult.success && checkResult.data?.dataElements?.length > 0) {
                    const existing = checkResult.data.dataElements[0];
                    state.dhis2.dataElements[field.name] = existing.id;
                    createdElements.push({ id: existing.id, name: field.label });
                    addLog('info', `  ↳ Already exists: ${field.label}`);
                    found = true;
                }
                
                if (!found) {
                    // Try to find by name
                    const nameCheck = await dhis2Request(`dataElements.json?filter=name:eq:${encodeURIComponent(cleanLabel)}&fields=id,name`);
                    if (nameCheck.success && nameCheck.data?.dataElements?.length > 0) {
                        const existing = nameCheck.data.dataElements[0];
                        state.dhis2.dataElements[field.name] = existing.id;
                        createdElements.push({ id: existing.id, name: field.label });
                        addLog('info', `  ↳ Already exists: ${field.label}`);
                        found = true;
                    }
                }
                
                if (!found) {
                    // Try to create
                    let valueType = 'TEXT';
                    if (field.type === 'number') valueType = 'NUMBER';
                    else if (field.type === 'date') valueType = 'DATE';
                    
                    const payload = {
                        name: cleanLabel,
                        shortName: cleanLabel.substring(0, 50),
                        code: code,
                        domainType: 'TRACKER',
                        valueType: valueType,
                        aggregationType: 'NONE'
                    };
                    
                    const createResult = await dhis2Request('dataElements', 'POST', payload);
                    
                    const newId = createResult.data?.response?.uid || 
                                 createResult.data?.uid ||
                                 (createResult.data?.status === 'OK' && createResult.data?.response?.uid);
                    
                    if (createResult.success && newId) {
                        state.dhis2.dataElements[field.name] = newId;
                        createdElements.push({ id: newId, name: field.label });
                        addLog('success', `  ✓ Created: ${field.label}`);
                    } else {
                        // Creation failed - search more broadly by name containing
                        const likeCheck = await dhis2Request(`dataElements.json?filter=name:ilike:${encodeURIComponent(cleanLabel)}&fields=id,name`);
                        if (likeCheck.success && likeCheck.data?.dataElements?.length > 0) {
                            const existing = likeCheck.data.dataElements[0];
                            state.dhis2.dataElements[field.name] = existing.id;
                            createdElements.push({ id: existing.id, name: field.label });
                            addLog('info', `  ↳ Already exists: ${field.label}`);
                        } else {
                            // Still try to use it - maybe it exists with different domainType
                            addLog('info', `  ⊘ Skipped: ${field.label} (may exist with different type)`);
                        }
                    }
                }
                await sleep(200);
            }
            
        } else {
            // ========== AGGREGATE MODE ==========
            // Only NUMBER and CATEGORICAL fields - skip text-based fields
            for (const field of dataFields) {
                // Skip text-based fields completely
                if (textBasedTypes.includes(field.type)) {
                    addLog('info', `  ⊘ Skipped: ${field.label} (${field.type} cannot be aggregated)`);
                    continue;
                }
                
                if (categoricalTypes.includes(field.type)) {
                    // Split categorical into separate INTEGER columns
                    const options = field.type === 'yesno' ? ['Yes', 'No'] : (field.options || []);
                    
                    for (const opt of options) {
                        const colName = `${field.name}_${opt}`;
                        const code = colName.toUpperCase().replace(/[^A-Z0-9_]/g, '_');
                        const label = `${field.label} (${opt})`;
                        const cleanLabel = label.replace(/[<>]/g, '');
                        
                        let found = false;
                        
                        // First check if exists by code
                        const checkResult = await dhis2Request(`dataElements.json?filter=code:eq:${code}&fields=id,name`);
                        
                        if (checkResult.success && checkResult.data?.dataElements?.length > 0) {
                            const existing = checkResult.data.dataElements[0];
                            state.dhis2.dataElements[colName] = existing.id;
                            createdElements.push({ id: existing.id, name: label });
                            addLog('info', `  ↳ Already exists: ${label}`);
                            found = true;
                        }
                        
                        if (!found) {
                            // Check by name
                            const nameCheck = await dhis2Request(`dataElements.json?filter=name:eq:${encodeURIComponent(cleanLabel)}&fields=id,name`);
                            if (nameCheck.success && nameCheck.data?.dataElements?.length > 0) {
                                const existing = nameCheck.data.dataElements[0];
                                state.dhis2.dataElements[colName] = existing.id;
                                createdElements.push({ id: existing.id, name: label });
                                addLog('info', `  ↳ Already exists: ${label}`);
                                found = true;
                            }
                        }
                        
                        if (!found) {
                            // Try to create
                            const payload = {
                                name: cleanLabel,
                                shortName: cleanLabel.substring(0, 50),
                                code: code,
                                domainType: 'AGGREGATE',
                                valueType: 'INTEGER',
                                aggregationType: 'SUM'
                            };
                            
                            const createResult = await dhis2Request('dataElements', 'POST', payload);
                            
                            const newId = createResult.data?.response?.uid || 
                                         createResult.data?.uid ||
                                         (createResult.data?.status === 'OK' && createResult.data?.response?.uid);
                            
                            if (createResult.success && newId) {
                                state.dhis2.dataElements[colName] = newId;
                                createdElements.push({ id: newId, name: label });
                                addLog('success', `  ✓ Created: ${label}`);
                            } else {
                                // Creation failed - search more broadly
                                const likeCheck = await dhis2Request(`dataElements.json?filter=name:ilike:${encodeURIComponent(cleanLabel)}&fields=id,name`);
                                if (likeCheck.success && likeCheck.data?.dataElements?.length > 0) {
                                    const existing = likeCheck.data.dataElements[0];
                                    state.dhis2.dataElements[colName] = existing.id;
                                    createdElements.push({ id: existing.id, name: label });
                                    addLog('info', `  ↳ Already exists: ${label}`);
                                } else {
                                    addLog('info', `  ⊘ Skipped: ${label} (may exist with different type)`);
                                }
                            }
                        }
                        await sleep(200);
                    }
                } else if (field.type === 'number') {
                    // Numeric field - SUM aggregation
                    const code = field.name.toUpperCase().replace(/[^A-Z0-9_]/g, '_');
                    const cleanLabel = field.label.replace(/[<>]/g, '');
                    
                    let found = false;
                    
                    // First check if exists by code
                    const checkResult = await dhis2Request(`dataElements.json?filter=code:eq:${code}&fields=id,name`);
                    
                    if (checkResult.success && checkResult.data?.dataElements?.length > 0) {
                        const existing = checkResult.data.dataElements[0];
                        state.dhis2.dataElements[field.name] = existing.id;
                        createdElements.push({ id: existing.id, name: field.label });
                        addLog('info', `  ↳ Already exists: ${field.label}`);
                        found = true;
                    }
                    
                    if (!found) {
                        // Check by name
                        const nameCheck = await dhis2Request(`dataElements.json?filter=name:eq:${encodeURIComponent(cleanLabel)}&fields=id,name`);
                        if (nameCheck.success && nameCheck.data?.dataElements?.length > 0) {
                            const existing = nameCheck.data.dataElements[0];
                            state.dhis2.dataElements[field.name] = existing.id;
                            createdElements.push({ id: existing.id, name: field.label });
                            addLog('info', `  ↳ Already exists: ${field.label}`);
                            found = true;
                        }
                    }
                    
                    if (!found) {
                        // Try to create
                        const payload = {
                            name: cleanLabel,
                            shortName: cleanLabel.substring(0, 50),
                            code: code,
                            domainType: 'AGGREGATE',
                            valueType: 'NUMBER',
                            aggregationType: 'SUM'
                        };
                        
                        const createResult = await dhis2Request('dataElements', 'POST', payload);
                        
                        const newId = createResult.data?.response?.uid || 
                                     createResult.data?.uid ||
                                     (createResult.data?.status === 'OK' && createResult.data?.response?.uid);
                        
                        if (createResult.success && newId) {
                            state.dhis2.dataElements[field.name] = newId;
                            createdElements.push({ id: newId, name: field.label });
                            addLog('success', `  ✓ Created: ${field.label}`);
                        } else {
                            // Creation failed - search more broadly
                            const likeCheck = await dhis2Request(`dataElements.json?filter=name:ilike:${encodeURIComponent(cleanLabel)}&fields=id,name`);
                            if (likeCheck.success && likeCheck.data?.dataElements?.length > 0) {
                                const existing = likeCheck.data.dataElements[0];
                                state.dhis2.dataElements[field.name] = existing.id;
                                createdElements.push({ id: existing.id, name: field.label });
                                addLog('info', `  ↳ Already exists: ${field.label}`);
                            } else {
                                addLog('info', `  ⊘ Skipped: ${field.label} (may exist with different type)`);
                            }
                        }
                    }
                    await sleep(200);
                } else {
                    // Skip other field types (date, time, rating, etc.)
                    addLog('info', `  ⊘ Skipped: ${field.label} (${field.type} not supported in aggregate)`);
                }
            }
        }
        
        addLog('success', `✓ ${createdElements.length} data elements ready`);
        
        if (state.dhis2.syncMode === 'aggregate') {
            await setupAggregateDataset(createdElements);
        } else {
            await setupTrackerProgram(createdElements);
        }
        
        updateDhis2Status('connected', 'Setup complete!');
        addLog('success', 'DHIS2 SETUP COMPLETE!');
        saveToStorage();
        saveDhis2Config(); // Also save DHIS2 config with programId
        notify('DHIS2 setup complete!');
        
    } catch (error) {
        updateDhis2Status('disconnected', 'Setup failed');
        addLog('error', '✗ Error: ' + error.message);
        notify('Setup failed', 'error');
    }
}

async function setupTrackerProgram(createdElements) {
    // Ensure org units are loaded
    if (state.dhis2.orgUnits.length === 0) {
        await fetchOrgUnits();
    }
    
    const programName = state.settings.title;
    const programCode = programName.toUpperCase().replace(/[^A-Z0-9]/g, '_').substring(0, 30);
    
    // Check if user provided an existing Program ID
    const userProvidedProgramId = document.getElementById('dhis2ProgramId').value.trim();
    
    let programId = null;
    let programStageId = null;
    
    if (userProvidedProgramId) {
        // User provided a Program ID - verify it exists
        addLog('info', `Checking existing Program: ${userProvidedProgramId}...`);
        const existingProgResult = await dhis2Request(`programs/${userProvidedProgramId}.json?fields=id,name,programStages[id]`);
        
        if (existingProgResult.success && existingProgResult.data?.id) {
            programId = existingProgResult.data.id;
            programStageId = existingProgResult.data.programStages?.[0]?.id;
            state.dhis2.programId = programId;
            addLog('info', `  ↳ Using existing Program: ${existingProgResult.data.name}`);
        } else {
            addLog('info', `  Program not found: ${userProvidedProgramId}, will create new...`);
        }
    }
    
    if (!programId) {
        addLog('info', `Setting up Event Program: ${programName}...`);
        
        // Check if program exists by code
        const progCheckResult = await dhis2Request(`programs.json?filter=code:eq:${programCode}&fields=id,name,programStages[id]`);
        
        if (progCheckResult.success && progCheckResult.data?.programs?.length > 0) {
            const existingProg = progCheckResult.data.programs[0];
            programId = existingProg.id;
            programStageId = existingProg.programStages?.[0]?.id;
            state.dhis2.programId = programId;
            addLog('info', `  ↳ Program already exists: ${existingProg.name} (will add data elements to it)`);
        } else {
            // Check by name if not found by code
            const nameCheck = await dhis2Request(`programs.json?filter=name:eq:${encodeURIComponent(programName)}&fields=id,name,programStages[id]`);
            if (nameCheck.success && nameCheck.data?.programs?.length > 0) {
                const existingProg = nameCheck.data.programs[0];
                programId = existingProg.id;
                programStageId = existingProg.programStages?.[0]?.id;
                state.dhis2.programId = programId;
                addLog('info', `  ↳ Program already exists: ${existingProg.name} (will add data elements to it)`);
            } else {
                // Create new Event Program (WITHOUT_REGISTRATION = Event Program)
                const programPayload = {
                    name: programName,
                    shortName: programName.substring(0, 50),
                    code: programCode,
                    programType: 'WITHOUT_REGISTRATION', // Event Program for case-based data
                    organisationUnits: state.dhis2.orgUnits.map(ou => ({ id: ou.id }))
                };
                
                const createProgResult = await dhis2Request('programs', 'POST', programPayload);
                
                if (createProgResult.success && createProgResult.data?.response?.uid) {
                    programId = createProgResult.data.response.uid;
                    state.dhis2.programId = programId;
                    addLog('success', `  ✓ Program created: ${programName}`);
                    
                    // Create Program Stage
                    await sleep(300);
                    const stagePayload = {
                        name: programName + ' - Data Entry',
                        program: { id: programId },
                        sortOrder: 1
                    };
                    
                    const createStageResult = await dhis2Request('programStages', 'POST', stagePayload);
                    
                    if (createStageResult.success && createStageResult.data?.response?.uid) {
                        programStageId = createStageResult.data.response.uid;
                        addLog('success', `  ✓ Program Stage created`);
                    } else {
                        addLog('warning', '  ⚠ Could not create Program Stage');
                    }
                } else {
                    // Creation failed - try to find again by name or ilike
                    const retryCheck = await dhis2Request(`programs.json?filter=name:eq:${encodeURIComponent(programName)}&fields=id,name,programStages[id]`);
                    if (retryCheck.success && retryCheck.data?.programs?.length > 0) {
                        const existingProg = retryCheck.data.programs[0];
                        programId = existingProg.id;
                        programStageId = existingProg.programStages?.[0]?.id;
                        state.dhis2.programId = programId;
                        addLog('info', `  ↳ Program already exists: ${existingProg.name} (will add data elements to it)`);
                    } else {
                        // Try broader search
                        const likeCheck = await dhis2Request(`programs.json?filter=name:ilike:${encodeURIComponent(programName)}&fields=id,name,programStages[id]`);
                        if (likeCheck.success && likeCheck.data?.programs?.length > 0) {
                            const existingProg = likeCheck.data.programs[0];
                            programId = existingProg.id;
                            programStageId = existingProg.programStages?.[0]?.id;
                            state.dhis2.programId = programId;
                            addLog('info', `  ↳ Program already exists: ${existingProg.name} (will add data elements to it)`);
                        } else {
                            // Last resort - get all programs and look for partial match
                            const allProgsCheck = await dhis2Request(`programs.json?fields=id,name,programStages[id]&pageSize=100`);
                            if (allProgsCheck.success && allProgsCheck.data?.programs?.length > 0) {
                                const searchName = programName.toLowerCase();
                                const matchedProg = allProgsCheck.data.programs.find(p => 
                                    p.name.toLowerCase().includes(searchName) || 
                                    searchName.includes(p.name.toLowerCase())
                                );
                                if (matchedProg) {
                                    programId = matchedProg.id;
                                    programStageId = matchedProg.programStages?.[0]?.id;
                                    state.dhis2.programId = programId;
                                    addLog('info', `  ↳ Program found: ${matchedProg.name} (will add data elements to it)`);
                                } else {
                                    addLog('info', `  ⊘ Program could not be created (may need manual creation in DHIS2)`);
                                    return;
                                }
                            } else {
                                addLog('info', `  ⊘ Program could not be created (may need manual creation in DHIS2)`);
                                return;
                            }
                        }
                    }
                }
            }
        }
    }
    
    // If program exists but has no stage, create one
    if (programId && !programStageId) {
        addLog('info', '  Creating Program Stage for existing program...');
        const stagePayload = {
            name: programName + ' - Data Entry',
            program: { id: programId },
            sortOrder: 1
        };
        
        const createStageResult = await dhis2Request('programStages', 'POST', stagePayload);
        
        if (createStageResult.success && createStageResult.data?.response?.uid) {
            programStageId = createStageResult.data.response.uid;
            addLog('success', `  ✓ Program Stage created`);
        } else {
            // Stage creation failed - try to find existing stage
            const stageCheck = await dhis2Request(`programStages.json?filter=program.id:eq:${programId}&fields=id,name`);
            if (stageCheck.success && stageCheck.data?.programStages?.length > 0) {
                programStageId = stageCheck.data.programStages[0].id;
                addLog('info', `  ↳ Program Stage already exists`);
            } else {
                addLog('warning', '  ⚠ Could not create or find Program Stage');
            }
        }
    }
    
    // Assign Data Elements to Program Stage (merge with existing)
    if (programStageId && createdElements.length > 0) {
        addLog('info', `Adding ${createdElements.length} data elements to Program Stage...`);
        
        // Get current program stage to update it
        const stageResult = await dhis2Request(`programStages/${programStageId}.json?fields=:owner`);
        
        if (stageResult.success && stageResult.data) {
            const stageData = stageResult.data;
            
            // Get existing data element IDs
            const existingElementIds = new Set(
                (stageData.programStageDataElements || []).map(psde => psde.dataElement?.id).filter(Boolean)
            );
            
            // Use form field order from createdElements, then add any existing that are not in new list
            const newElementIds = createdElements.map(el => el.id);
            const extraExistingIds = [...existingElementIds].filter(id => !newElementIds.includes(id));
            
            // Form order first, then any extras that existed before
            const orderedElementIds = [...newElementIds, ...extraExistingIds];
            
            // Build programStageDataElements array with correct sortOrder
            const programStageDataElements = orderedElementIds.map((id, idx) => ({
                dataElement: { id: id },
                compulsory: false,
                sortOrder: idx + 1
            }));
            
            stageData.programStageDataElements = programStageDataElements;
            
            const updateResult = await dhis2Request(`programStages/${programStageId}`, 'PUT', stageData);
            
            if (updateResult.success) {
                const newCount = newElementIds.filter(id => !existingElementIds.has(id)).length;
                addLog('success', `  ✓ Program Stage updated: ${newCount} new elements added, ${orderedElementIds.length} total`);
            } else {
                addLog('warning', '  ⚠ Could not update Program Stage');
            }
        } else {
            addLog('warning', '  ⚠ Could not load Program Stage data');
        }
    } else if (!programStageId) {
        addLog('warning', '  ⚠ No Program Stage available - data elements not assigned');
    } else if (createdElements.length === 0) {
        addLog('info', '  ⊘ No data elements to add to Program Stage');
    }
    
    // Update Program ID in the UI
    if (document.getElementById('dhis2ProgramId')) {
        document.getElementById('dhis2ProgramId').value = programId;
    }
    addLog('success', `Event Program ready: ${programName}`);
}

async function setupAggregateDataset(createdElements) {
    if (state.dhis2.orgUnits.length === 0) {
        await fetchOrgUnits();
    }
    
    if (createdElements.length === 0) {
        addLog('warning', '⚠ No data elements to add to dataset');
        return;
    }
    
    const datasetName = state.settings.title;
    const datasetCode = datasetName.toUpperCase().replace(/[^A-Z0-9]/g, '_').substring(0, 30);
    
    addLog('info', `Setting up Dataset: ${datasetName}...`);
    
    // Check if dataset exists by code
    let datasetId = null;
    const dsCheckResult = await dhis2Request(`dataSets.json?filter=code:eq:${datasetCode}&fields=id,name,dataSetElements[dataElement[id]]`);
    
    if (dsCheckResult.success && dsCheckResult.data?.dataSets?.length > 0) {
        datasetId = dsCheckResult.data.dataSets[0].id;
        addLog('info', `  ↳ Dataset already exists: ${datasetName}`);
    } else {
        // Check by name if not found by code
        const nameCheck = await dhis2Request(`dataSets.json?filter=name:eq:${encodeURIComponent(datasetName)}&fields=id,name,dataSetElements[dataElement[id]]`);
        if (nameCheck.success && nameCheck.data?.dataSets?.length > 0) {
            datasetId = nameCheck.data.dataSets[0].id;
            addLog('info', `  ↳ Dataset already exists: ${datasetName}`);
        }
    }
    
    if (datasetId) {
        // Dataset exists - update with merged elements
        state.dhis2.datasetId = datasetId;
        
        addLog('info', '  Merging data elements into dataset...');
        const fullDs = await dhis2Request(`dataSets/${datasetId}.json?fields=:owner`);
        
        if (fullDs.success && fullDs.data) {
            // Get existing data element IDs
            const existingElementIds = new Set(
                (fullDs.data.dataSetElements || []).map(dse => dse.dataElement?.id).filter(Boolean)
            );
            
            // Add new elements (merge, don't replace)
            const newElementIds = createdElements.map(de => de.id);
            const allElementIds = [...new Set([...existingElementIds, ...newElementIds])];
            
            const merged = {
                ...fullDs.data,
                dataSetElements: allElementIds.map(id => ({ 
                    dataSet: { id: datasetId },
                    dataElement: { id: id } 
                })),
                organisationUnits: state.dhis2.orgUnits.map(ou => ({ id: ou.id }))
            };
            
            const updateResult = await dhis2Request(`dataSets/${datasetId}`, 'PUT', merged);
            if (updateResult.success) {
                const newCount = newElementIds.filter(id => !existingElementIds.has(id)).length;
                addLog('success', `  ✓ Dataset updated: ${newCount} new elements added, ${allElementIds.length} total elements`);
            } else {
                const errorMsg = updateResult.data?.message || updateResult.error || JSON.stringify(updateResult.data);
                addLog('warning', `  ⚠ Could not update dataset: ${errorMsg}`);
            }
        }
    } else {
        // Create new dataset
        const createPayload = {
            name: datasetName,
            shortName: datasetName.substring(0, 50),
            code: datasetCode,
            periodType: state.dhis2.periodType,
            dataSetElements: createdElements.map(de => ({ dataElement: { id: de.id } })),
            organisationUnits: state.dhis2.orgUnits.map(ou => ({ id: ou.id }))
        };
        
        addLog('info', `  Creating dataset with ${createdElements.length} elements...`);
        
        const createResult = await dhis2Request('dataSets', 'POST', createPayload);
        
        const newId = createResult.data?.response?.uid || createResult.data?.uid;
        
        if (createResult.success && newId) {
            state.dhis2.datasetId = newId;
            addLog('success', `  ✓ Dataset created: ${datasetName}`);
        } else {
            // Creation failed - try to find by name again (maybe created by someone else)
            const retryCheck = await dhis2Request(`dataSets.json?filter=name:eq:${encodeURIComponent(datasetName)}&fields=id,name`);
            if (retryCheck.success && retryCheck.data?.dataSets?.length > 0) {
                state.dhis2.datasetId = retryCheck.data.dataSets[0].id;
                addLog('info', `  ↳ Dataset already exists: ${datasetName}`);
                // Now update it with elements
                const fullDs = await dhis2Request(`dataSets/${state.dhis2.datasetId}.json?fields=:owner`);
                if (fullDs.success && fullDs.data) {
                    const existingElementIds = new Set(
                        (fullDs.data.dataSetElements || []).map(dse => dse.dataElement?.id).filter(Boolean)
                    );
                    const allElementIds = [...new Set([...existingElementIds, ...createdElements.map(de => de.id)])];
                    
                    const merged = {
                        ...fullDs.data,
                        dataSetElements: allElementIds.map(id => ({ 
                            dataSet: { id: state.dhis2.datasetId },
                            dataElement: { id: id } 
                        })),
                        organisationUnits: state.dhis2.orgUnits.map(ou => ({ id: ou.id }))
                    };
                    
                    const updateResult = await dhis2Request(`dataSets/${state.dhis2.datasetId}`, 'PUT', merged);
                    if (updateResult.success) {
                        addLog('success', `  ✓ Dataset updated with ${createdElements.length} elements`);
                    }
                }
            } else {
                // Try broader search
                const likeCheck = await dhis2Request(`dataSets.json?filter=name:ilike:${encodeURIComponent(datasetName)}&fields=id,name`);
                if (likeCheck.success && likeCheck.data?.dataSets?.length > 0) {
                    state.dhis2.datasetId = likeCheck.data.dataSets[0].id;
                    addLog('info', `  ↳ Dataset already exists: ${likeCheck.data.dataSets[0].name}`);
                } else {
                    addLog('info', `  ⊘ Dataset setup skipped (may need manual creation)`);
                }
            }
        }
    }
}

async function syncToDhis2() {
    if (!state.dhis2.url || !state.dhis2.username) {
        notify('DHIS2 not configured', 'error');
        return;
    }
    
    // Ensure connected flag is set
    state.dhis2.connected = true;
    
    clearLog();
    
    if (state.dhis2.syncMode === 'aggregate') {
        await syncAggregateData();
    } else {
        await syncTrackerData();
    }
}

// Separate sync function for Case-Based (Tracker)
window.syncCaseBased = async function() {
    if (!state.dhis2.url || !state.dhis2.username) {
        notify('DHIS2 not configured', 'error');
        return;
    }
    state.dhis2.connected = true;
    clearLog();
    await syncTrackerData();
};

// Separate sync function for Aggregate
window.syncAggregate = async function() {
    if (!state.dhis2.url || !state.dhis2.username) {
        notify('DHIS2 not configured', 'error');
        return;
    }
    state.dhis2.connected = true;
    clearLog();
    await syncAggregateData();
};

async function syncAggregateData() {
    addLog('info', 'Syncing AGGREGATE data to DHIS2...');
    updateDhis2Status('syncing', 'Syncing...');
    
    // Check if DHIS2 columns are configured
    if (!state.dhis2.periodColumn || !state.dhis2.orgUnitColumn) {
        addLog('error', 'Period Column and Org Unit Column must be set in DHIS2 settings');
        notify('Set Period Column and Org Unit Column in DHIS2 settings first', 'error');
        updateDhis2Status('disconnected', 'Config needed');
        return;
    }
    
    // Get aggregate data
    const aggregateData = calculateAggregateData();
    
    if (aggregateData.length === 0) {
        notify('No data to sync', 'info');
        addLog('info', 'No aggregate data to sync');
        return;
    }
    
    const periodColumn = state.dhis2.periodColumn;
    const orgUnitColumn = state.dhis2.orgUnitColumn;
    const aggregateColumn = state.settings.aggregateColumn || orgUnitColumn;
    
    let success = 0, failed = 0;
    
    for (const record of aggregateData) {
        const period = record._period;
        const groupValue = record._group;
        const orgUnitId = state.dhis2.orgUnitMap[groupValue?.toLowerCase().trim()];
        
        if (!orgUnitId) {
            addLog('error', `  ✗ No org unit match: ${groupValue}`);
            failed++;
            continue;
        }
        
    const dataValues = [];
        // Skip text-based fields that can't be aggregated
        const skipTypes = ['phone', 'gps', 'email', 'text', 'textarea', 'date', 'time'];
        const dataFields = state.fields.filter(f => 
            f.type !== 'section' && 
            f.name !== periodColumn && 
            f.name !== aggregateColumn &&
            !skipTypes.includes(f.type)
        );
        
        const categoricalTypes = ['select', 'radio', 'yesno', 'checkbox'];
        
        dataFields.forEach(field => {
            const def = fieldDefs[field.type];
            
            if (categoricalTypes.includes(field.type)) {
                // For categorical fields, send each option as separate value
                const options = field.type === 'yesno' ? ['Yes', 'No'] : (field.options || []);
                options.forEach(opt => {
                    const colName = `${field.name}_${opt}`;
                    const deId = state.dhis2.dataElements[colName];
                    const value = record[colName];
                    
                    if (deId && value !== undefined && value !== null) {
                        dataValues.push({
                            dataElement: deId,
                            value: String(value),
                            period: period,
                            orgUnit: orgUnitId
                        });
                    }
                });
            } else if (field.type === 'number') {
                // Only numeric fields
                const deId = state.dhis2.dataElements[field.name];
                const value = record[field.name];
                
                if (deId && value !== undefined && value !== null && value !== '') {
                    dataValues.push({
                        dataElement: deId,
                        value: String(value),
                        period: period,
                        orgUnit: orgUnitId
                    });
                }
            }
        });
        
        if (dataValues.length === 0) continue;
        
        const result = await dhis2Request('dataValueSets', 'POST', { dataValues });
        
        if (result.success) {
            success++;
            addLog('success', `  ✓ Synced: ${orgUnitName} / ${period}`);
        } else {
            failed++;
            addLog('error', `  ✗ Failed: ${orgUnitName} / ${period}`);
        }
        
        await sleep(300);
    }
    
    updateDhis2Status('connected', 'Sync complete');
    addLog('info', `Sync complete: ${success} success, ${failed} failed`);
    notify(`Synced ${success} records to DHIS2!`);
}

async function syncTrackerData() {
    addLog('info', 'Syncing TRACKER/EVENT data to DHIS2...');
    updateDhis2Status('syncing', 'Syncing...');
    
    let programId = state.dhis2.programId;
    
    // If no programId, try to find the program by name
    if (!programId) {
        addLog('info', 'No Program ID saved, searching for existing program...');
        const programName = state.settings.title;
        const programCode = programName.toUpperCase().replace(/[^A-Z0-9]/g, '_').substring(0, 30);
        
        // Check by code first
        const codeCheck = await dhis2Request(`programs.json?filter=code:eq:${programCode}&fields=id,name,programStages[id]`);
        if (codeCheck.success && codeCheck.data?.programs?.length > 0) {
            const prog = codeCheck.data.programs[0];
            programId = prog.id;
            state.dhis2.programId = programId;
            addLog('info', `  ↳ Found program: ${prog.name}`);
            if (document.getElementById('dhis2ProgramId')) {
                document.getElementById('dhis2ProgramId').value = programId;
            }
            saveToStorage();
            saveDhis2Config();
        } else {
            // Check by name
            const nameCheck = await dhis2Request(`programs.json?filter=name:eq:${encodeURIComponent(programName)}&fields=id,name,programStages[id]`);
            if (nameCheck.success && nameCheck.data?.programs?.length > 0) {
                const prog = nameCheck.data.programs[0];
                programId = prog.id;
                state.dhis2.programId = programId;
                addLog('info', `  ↳ Found program: ${prog.name}`);
                if (document.getElementById('dhis2ProgramId')) {
                    document.getElementById('dhis2ProgramId').value = programId;
                }
                saveToStorage();
                saveDhis2Config();
            } else {
                // Try ilike search
                const likeCheck = await dhis2Request(`programs.json?filter=name:ilike:${encodeURIComponent(programName)}&fields=id,name,programStages[id]`);
                if (likeCheck.success && likeCheck.data?.programs?.length > 0) {
                    const prog = likeCheck.data.programs[0];
                    programId = prog.id;
                    state.dhis2.programId = programId;
                    addLog('info', `  ↳ Found program: ${prog.name}`);
                    if (document.getElementById('dhis2ProgramId')) {
                        document.getElementById('dhis2ProgramId').value = programId;
                    }
                    saveToStorage();
                    saveDhis2Config();
                }
            }
        }
    }
    
    if (!programId) {
        notify('Run Setup DHIS2 first to create the Program', 'error');
        addLog('info', '⊘ No Program found - run Setup first to create it');
        return;
    }
    
    // Ensure org units are loaded
    if (state.dhis2.orgUnits.length === 0) {
        await fetchOrgUnits();
    }
    
    const pendingData = state.collectedData.filter(d => !d._synced);
    
    if (pendingData.length === 0) {
        notify('No pending data to sync', 'info');
        addLog('info', 'No pending data to sync');
        return;
    }
    
    const periodColumn = state.dhis2.periodColumn;
    const orgUnitColumn = state.dhis2.orgUnitColumn;
    
    addLog('info', `Syncing ${pendingData.length} events...`);
    
    let success = 0, failed = 0;
    
    for (const record of pendingData) {
        // Get org unit
        const orgUnitName = orgUnitColumn ? record[orgUnitColumn] : '';
        let orgUnitId = state.dhis2.orgUnitMap[orgUnitName?.toLowerCase().trim()];
        
        // Try partial match if exact match fails
        if (!orgUnitId && orgUnitName) {
            const searchKey = orgUnitName.toLowerCase().trim();
            for (const [name, id] of Object.entries(state.dhis2.orgUnitMap)) {
                if (name.includes(searchKey) || searchKey.includes(name)) {
                    orgUnitId = id;
                    break;
                }
            }
        }
        
        if (!orgUnitId) {
            // Use first org unit as fallback
            orgUnitId = state.dhis2.orgUnits[0]?.id;
            if (!orgUnitId) {
                addLog('error', `  ✗ No org unit: ${orgUnitName}`);
                failed++;
                continue;
            }
            addLog('warning', `  ⚠ Using default org unit for: ${orgUnitName}`);
        }
        
        // Get event date from period or timestamp
        let eventDate = record._timestamp?.split('T')[0] || new Date().toISOString().split('T')[0];
        if (periodColumn && record[periodColumn]) {
            // Convert YYYYMM to YYYY-MM-01
            const period = record[periodColumn];
            if (period.length === 6) {
                eventDate = `${period.substring(0, 4)}-${period.substring(4, 6)}-01`;
            }
        }
        
        // Build data values (excluding period and orgunit columns)
        const dataValues = [];
        const dataFields = state.fields.filter(f => 
            f.type !== 'section' && f.name !== periodColumn && f.name !== orgUnitColumn
        );
        
        dataFields.forEach(field => {
            const deId = state.dhis2.dataElements[field.id];
            const value = record[field.name];
            
            if (deId && value !== undefined && value !== null && value !== '') {
                dataValues.push({
                    dataElement: deId,
                    value: String(value)
                });
            }
        });
        
        if (dataValues.length === 0) {
            addLog('warning', '  ⚠ No data values in record');
            failed++;
            continue;
        }
        
        // Create event payload
        const event = {
            program: programId,
            orgUnit: orgUnitId,
            eventDate: eventDate,
            status: 'COMPLETED',
            dataValues: dataValues
        };
        
        const result = await dhis2Request('events', 'POST', event);
        
        if (result.success && (result.data?.response?.imported > 0 || result.data?.response?.importSummaries)) {
            record._synced = true;
            record._syncedAt = new Date().toISOString();
            delete record._syncError;
            success++;
            addLog('success', `  ✓ Event synced: ${orgUnitName || 'record'}`);
        } else {
            record._syncError = result.error || JSON.stringify(result.data?.message || result.data?.response?.importSummaries?.[0]?.description) || 'Unknown error';
            failed++;
            addLog('error', `  ✗ Event failed: ${record._syncError}`);
        }
        
        await sleep(300);
    }
    
    updateDhis2Status('connected', 'Sync complete');
    addLog('info', `Sync complete: ${success} success, ${failed} failed`);
    saveToStorage();
    saveFormToList();
    
    if (success > 0) notify(`Synced ${success} events to DHIS2!`);
    if (failed > 0) notify(`${failed} events failed`, 'error');
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ==================== DATA AGGREGATION ====================
function calculateAggregateData() {
    const periodColumn = state.dhis2.periodColumn;
    const orgUnitColumn = state.dhis2.orgUnitColumn;
    
    // Support both legacy single column and new multi-column
    let aggregateColumns = state.settings.aggregateColumns || [];
    if (aggregateColumns.length === 0 && state.settings.aggregateColumn) {
        aggregateColumns = [state.settings.aggregateColumn];
    }
    if (aggregateColumns.length === 0 && orgUnitColumn) {
        aggregateColumns = [orgUnitColumn];
    }
    
    // If no aggregation column set, use default grouping by month
    const useDefaultGrouping = aggregateColumns.length === 0 && !periodColumn;
    
    const grouped = {};
    const data = getFilteredData();
    
    if (data.length === 0) return [];
    
    // Get field names to skip (period and aggregate columns)
    const skipFields = [periodColumn, ...aggregateColumns].filter(Boolean);
    // Field types that cannot be aggregated
    const skipTypes = ['phone', 'gps', 'email', 'text', 'textarea', 'date', 'time'];
    
    data.forEach(record => {
        let period, groupValue, key;
        
        if (useDefaultGrouping) {
            // Use timestamp month as period, "All" as group
            const ts = record._timestamp ? new Date(record._timestamp) : new Date();
            period = `${ts.getFullYear()}${String(ts.getMonth() + 1).padStart(2, '0')}`;
            groupValue = 'All';
            key = `${groupValue}|||${period}`;
        } else {
            // Use period column or timestamp month
            if (periodColumn && record[periodColumn]) {
                period = record[periodColumn];
            } else {
                const ts = record._timestamp ? new Date(record._timestamp) : new Date();
                period = `${ts.getFullYear()}${String(ts.getMonth() + 1).padStart(2, '0')}`;
            }
            
            // Build composite group value from multiple columns
            if (aggregateColumns.length > 0) {
                const groupParts = aggregateColumns.map(col => record[col] || 'Unknown');
                groupValue = groupParts.join(' | ');
            } else {
                groupValue = 'All';
            }
            key = `${groupValue}|||${period}`;
        }
        
        if (!grouped[key]) {
            grouped[key] = {
                _group: groupValue,
                _period: period,
                _count: 0
            };
            
            // Add individual group column values for clarity
            aggregateColumns.forEach(col => {
                grouped[key]['_grp_' + col] = record[col] || 'Unknown';
            });
        }
        
        grouped[key]._count++;
        
        // Aggregate other fields (excluding period, aggregate columns, section, and text types)
        state.fields.forEach(field => {
            if (field.type === 'section') return;
            
            // Skip period and aggregate columns - they are used for grouping, not aggregation
            if (skipFields.includes(field.name)) return;
            
            // Skip non-aggregatable field types
            if (skipTypes.includes(field.type)) return;
            
            const value = record[field.name];
            const def = fieldDefs[field.type];
            
            if (def?.category === 'numeric' || field.type === 'number') {
                // Sum numbers
                const numVal = parseFloat(value) || 0;
                grouped[key][field.name] = (grouped[key][field.name] || 0) + numVal;
            } else if ((def?.category === 'categorical' || field.type === 'yesno') && field.type !== 'text') {
                // Split categorical fields (select, radio, yesno, checkbox) into separate columns
                // Example: Sex with Male/Female becomes Sex_Male, Sex_Female
                const options = field.type === 'yesno' ? ['Yes', 'No'] : (field.options || []);
                
                // Initialize all option columns to 0 if not exists
                options.forEach(opt => {
                    const colName = `${field.name}_${opt}`;
                    if (grouped[key][colName] === undefined) {
                        grouped[key][colName] = 0;
                    }
                });
                
                // Increment the matching option column
                if (value && options.includes(value)) {
                    const colName = `${field.name}_${value}`;
                    grouped[key][colName] = (grouped[key][colName] || 0) + 1;
                }
            }
            // Text fields and other types are skipped (not aggregated)
        });
    });
    
    return Object.values(grouped);
}

// ==================== FILTERING ====================
function getFilterableFields() {
    const categoricalTypes = ['select', 'radio', 'yesno', 'checkbox', 'rating'];
    return state.fields.filter(f => categoricalTypes.includes(f.type));
}

function getFilteredData() {
    let data = [...state.collectedData];
    
    // Date filter
    if (state.dateFilter.start || state.dateFilter.end) {
        data = data.filter(row => {
            const ts = row._timestamp;
            if (!ts) return true;
            const rowDate = new Date(ts);
            
            if (state.dateFilter.start && rowDate < new Date(state.dateFilter.start)) return false;
            if (state.dateFilter.end) {
                const endDate = new Date(state.dateFilter.end);
                endDate.setHours(23, 59, 59, 999);
                if (rowDate > endDate) return false;
            }
            return true;
        });
    }
    
    // Field filters
    Object.keys(state.filters).forEach(fieldName => {
        const filterValue = state.filters[fieldName];
        if (filterValue) {
            data = data.filter(row => {
                const rowValue = row[fieldName];
                return rowValue && rowValue.toString().trim() === filterValue;
            });
        }
    });
    
    return data;
}

window.updateFilter = function(fieldName, value) {
    if (value) {
        state.filters[fieldName] = value;
    } else {
        delete state.filters[fieldName];
    }
    applyFilters();
};

window.updateDateFilter = function(type, value) {
    state.dateFilter[type] = value;
    applyFilters();
};

window.clearAllFilters = function() {
    state.filters = {};
    state.dateFilter = { start: '', end: '' };
    document.querySelectorAll('.filter-select, .filter-input').forEach(el => el.value = '');
    applyFilters();
    notify('Filters cleared');
};

window.moveFilter = function(fieldName, direction) {
    const filterableFields = getFilterableFields();
    const fieldNames = filterableFields.map(f => f.name);
    
    // Initialize filterOrder if empty
    if (state.filterOrder.length === 0) {
        state.filterOrder = [...fieldNames];
    }
    
    // Add any new fields not in filterOrder
    fieldNames.forEach(name => {
        if (!state.filterOrder.includes(name)) {
            state.filterOrder.push(name);
        }
    });
    
    // Remove any fields no longer in filterableFields
    state.filterOrder = state.filterOrder.filter(name => fieldNames.includes(name));
    
    const idx = state.filterOrder.indexOf(fieldName);
    if (idx === -1) return;
    
    if (direction === 'up' && idx > 0) {
        [state.filterOrder[idx], state.filterOrder[idx - 1]] = [state.filterOrder[idx - 1], state.filterOrder[idx]];
    } else if (direction === 'down' && idx < state.filterOrder.length - 1) {
        [state.filterOrder[idx], state.filterOrder[idx + 1]] = [state.filterOrder[idx + 1], state.filterOrder[idx]];
    }
    
    saveToStorage();
    renderDataContent();
    renderDashboard();
};

function getOrderedFilterFields() {
    const filterableFields = getFilterableFields();
    
    // Initialize filterOrder if empty
    if (state.filterOrder.length === 0) {
        return filterableFields;
    }
    
    // Sort fields by filterOrder
    const ordered = [];
    state.filterOrder.forEach(name => {
        const field = filterableFields.find(f => f.name === name);
        if (field) ordered.push(field);
    });
    
    // Add any new fields not in filterOrder
    filterableFields.forEach(field => {
        if (!ordered.find(f => f.name === field.name)) {
            ordered.push(field);
        }
    });
    
    return ordered;
}

function applyFilters() {
    renderDataContent();
    renderDashboard();
}

// ==================== UI HELPERS ====================
function openDhis2Config() {
    document.getElementById('dhis2Url').value = state.dhis2.url || '';
    document.getElementById('dhis2Username').value = state.dhis2.username || '';
    document.getElementById('dhis2Password').value = state.dhis2.password || '';
    document.getElementById('dhis2OrgLevel').value = state.dhis2.orgUnitLevel || 5;
    document.getElementById('dhis2PeriodType').value = state.dhis2.periodType || 'Monthly';
    document.getElementById('dhis2ProgramId').value = state.dhis2.programId || '';
    selectSyncMode(state.dhis2.syncMode || 'aggregate');
    
    // Populate Period Column and Org Unit Column dropdowns
    const periodSelect = document.getElementById('dhis2PeriodColumn');
    const orgUnitSelect = document.getElementById('dhis2OrgUnitColumn');
    
    // Clear existing options except first
    periodSelect.innerHTML = '<option value="">-- Select field --</option>';
    orgUnitSelect.innerHTML = '<option value="">-- Select field --</option>';
    
    // Add form fields as options (exclude section)
    state.fields.filter(f => f.type !== 'section').forEach(f => {
        const periodOpt = document.createElement('option');
        periodOpt.value = f.name;
        periodOpt.textContent = f.label;
        periodSelect.appendChild(periodOpt);
        
        const orgOpt = document.createElement('option');
        orgOpt.value = f.name;
        orgOpt.textContent = f.label;
        orgUnitSelect.appendChild(orgOpt);
    });
    
    // Set selected values
    if (state.dhis2.periodColumn) periodSelect.value = state.dhis2.periodColumn;
    if (state.dhis2.orgUnitColumn) orgUnitSelect.value = state.dhis2.orgUnitColumn;
    
    document.getElementById('dhis2Modal').classList.add('show');
}

function closeModal(id) {
    document.getElementById(id).classList.remove('show');
}

// Compress cascade data: [{col1: val1, col2: val2}, ...] -> "val1|val2||val3|val4||..."
function compressCascadeData(data, columns) {
    if (!data || !Array.isArray(data) || data.length === 0) {
        console.log('compressCascadeData: No data or empty array');
        return null;
    }
    if (!columns || !Array.isArray(columns) || columns.length === 0) {
        console.log('compressCascadeData: No columns or empty array');
        return null;
    }
    
    try {
        console.log(`Compressing ${data.length} rows with ${columns.length} columns:`, columns);
        
        // Each row's values joined by |, rows joined by ||
        const result = data.map((row, rowIdx) => {
            if (!row || typeof row !== 'object') {
                console.log(`Row ${rowIdx} is invalid:`, row);
                return columns.map(() => '').join('|');
            }
            return columns.map((col, colIdx) => {
                try {
                    const val = row[col];
                    if (val === null || val === undefined) return '';
                    return String(val).replace(/\|/g, '¦');
                } catch (e) {
                    console.error(`Error at row ${rowIdx}, col ${colIdx} (${col}):`, e);
                    return '';
                }
            }).join('|');
        }).join('||');
        
        console.log(`Compressed to ${result.length} characters`);
        return result;
    } catch (err) {
        console.error('Compress cascade error:', err);
        return null;
    }
}

// Decompress cascade data: "val1|val2||val3|val4||..." -> [{col1: val1, col2: val2}, ...]
function decompressCascadeData(compressed, columns) {
    if (!compressed || typeof compressed !== 'string') return [];
    if (!columns || !Array.isArray(columns) || columns.length === 0) return [];
    
    try {
        const rows = compressed.split('||');
        return rows.map(row => {
            const values = row.split('|');
            const obj = {};
            columns.forEach((col, i) => {
                obj[col] = (values[i] || '').replace(/¦/g, '|');
            });
            return obj;
        }).filter(row => Object.values(row).some(v => v)); // Filter empty rows
    } catch (err) {
        console.error('Decompress cascade error:', err);
        return [];
    }
}

function notify(message, type = 'success') {
    const el = document.getElementById('notification');
    el.textContent = message;
    el.className = 'notification show' + (type === 'error' ? ' error' : type === 'info' ? ' info' : type === 'warning' ? ' warning' : '');
    setTimeout(() => el.classList.remove('show'), 4000);
}

function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ==================== AUTH ====================
function checkAuth() {
    const saved = safeStorage.getItem('icfCollectUser');
    if (saved) { state.user = JSON.parse(saved); showBuilder(); }
}

function switchAuthTab(tab) {
    document.querySelectorAll('.auth-tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.auth-form').forEach(f => f.classList.remove('active'));
    document.querySelector(`.auth-tab[data-tab="${tab}"]`)?.classList.add('active');
    document.getElementById(tab + 'Form')?.classList.add('active');
}

function showForgotPassword() {
    document.querySelectorAll('.auth-form').forEach(f => f.classList.remove('active'));
    document.querySelectorAll('.auth-tab').forEach(t => t.classList.remove('active'));
    document.getElementById('forgotForm').classList.add('active');
    document.getElementById('authError').style.display = 'none';
    document.getElementById('authSuccess').style.display = 'none';
}

function showAuthLoading(show) {
    document.getElementById('authLoading').style.display = show ? 'block' : 'none';
    document.querySelectorAll('.auth-btn').forEach(btn => btn.disabled = show);
}

async function handleLogin(e) {
    e.preventDefault();
    const email = document.getElementById('loginEmail').value.trim();
    const password = document.getElementById('loginPassword').value;
    
    document.getElementById('authError').style.display = 'none';
    showAuthLoading(true);
    
    try {
        const response = await fetch(CONFIG.AUTH_SCRIPT_URL, {
            method: 'POST', mode: 'cors', redirect: 'follow',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action: 'login',
                email: email,
                password: password
            })
        });
        const result = await response.json();
        
        showAuthLoading(false);
        
        if (result.success && result.user) {
            state.user = result.user;
            safeStorage.setItem('icfCollectUser', JSON.stringify(result.user));
            showBuilder();
        } else {
            document.getElementById('authError').style.display = 'block';
            document.getElementById('authError').textContent = result.error || 'Invalid credentials';
        }
    } catch (error) {
        showAuthLoading(false);
        // Fallback to local storage if offline
        const users = JSON.parse(safeStorage.getItem('icfCollectUsers') || '[]');
        const user = users.find(u => u.email === email && u.password === password);
        if (user) {
            state.user = user;
            safeStorage.setItem('icfCollectUser', JSON.stringify(user));
            showBuilder();
        } else {
            document.getElementById('authError').style.display = 'block';
            document.getElementById('authError').textContent = 'Connection error. Please try again.';
        }
    }
}

async function handleSignup(e) {
    e.preventDefault();
    const name = document.getElementById('signupName').value.trim();
    const email = document.getElementById('signupEmail').value.trim();
    const password = document.getElementById('signupPassword').value;
    
    document.getElementById('authError').style.display = 'none';
    document.getElementById('authSuccess').style.display = 'none';
    showAuthLoading(true);
    
    try {
        const response = await fetch(CONFIG.AUTH_SCRIPT_URL, {
            method: 'POST', mode: 'cors', redirect: 'follow',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action: 'signup',
                name: name,
                email: email,
                password: password
            })
        });
        const result = await response.json();
        
        showAuthLoading(false);
        
        if (result.success) {
            // Also save locally for offline
            const users = JSON.parse(safeStorage.getItem('icfCollectUsers') || '[]');
            if (!users.find(u => u.email === email)) {
                users.push({ id: result.user?.id || Date.now().toString(), name, email, password });
                safeStorage.setItem('icfCollectUsers', JSON.stringify(users));
            }
            
            document.getElementById('authSuccess').style.display = 'block';
            document.getElementById('authSuccess').textContent = 'Account created! Please login.';
            setTimeout(() => switchAuthTab('login'), 1500);
        } else {
            document.getElementById('authError').style.display = 'block';
            document.getElementById('authError').textContent = result.error || 'Registration failed';
        }
    } catch (error) {
        showAuthLoading(false);
        document.getElementById('authError').style.display = 'block';
        document.getElementById('authError').textContent = 'Connection error. Please try again.';
    }
}

async function handleForgotPassword(e) {
    e.preventDefault();
    const email = document.getElementById('forgotEmail').value.trim();
    
    document.getElementById('authError').style.display = 'none';
    document.getElementById('authSuccess').style.display = 'none';
    showAuthLoading(true);
    
    try {
        const response = await fetch(CONFIG.AUTH_SCRIPT_URL + '?action=forgotPassword&email=' + encodeURIComponent(email), {
            mode: 'cors',
            redirect: 'follow'
        });
        const result = await response.json();
        
        showAuthLoading(false);
        
        if (result.success) {
            document.getElementById('authSuccess').style.display = 'block';
            document.getElementById('authSuccess').textContent = 'Password sent to your email!';
            setTimeout(() => switchAuthTab('login'), 2000);
        } else {
            document.getElementById('authError').style.display = 'block';
            document.getElementById('authError').textContent = result.message || 'Email not found';
        }
    } catch (error) {
        showAuthLoading(false);
        document.getElementById('authError').style.display = 'block';
        document.getElementById('authError').textContent = 'Connection error. Please try again.';
    }
}

function showBuilder() {
    document.getElementById('authContainer').style.display = 'none';
    document.getElementById('mainContainer').classList.add('show');
    document.getElementById('headerUser').innerHTML = '<span data-icon="user" data-size="14"></span> ' + (state.user?.name || 'User');
    renderFields();
    initIcons();
}

function logout() {
    state.user = null;
    safeStorage.removeItem('icfCollectUser');
    document.getElementById('mainContainer').classList.remove('show');
    document.getElementById('authContainer').style.display = 'flex';
}

// ==================== FORM BUILDER ====================
function addField(type) {
    const def = fieldDefs[type];
    if (!def) return;
    
    // Check for duplicate period field
    if (type === 'period' && state.fields.find(f => f.type === 'period')) {
        notify('Period field already exists', 'error');
        return;
    }
    
    state.fieldCounter++;
    const field = {
        id: 'field_' + state.fieldCounter,
        type: type,
        label: def.defaultLabel,
        name: type + '_' + state.fieldCounter,
        required: true,
        options: def.options ? [...def.options] : [],
        max: def.max || 5
    };
    
    // Special naming for period field
    if (type === 'period') {
        field.name = 'period';
        field.label = 'Reporting Period';
    }
    
    state.fields.push(field);
    renderFields();
    selectField(field.id);
    saveToStorage();
}

function removeField(id) {
    event?.stopPropagation();
    if (!confirm('Delete field?')) return;
    state.fields = state.fields.filter(f => f.id !== id);
    if (state.selectedFieldId === id) { state.selectedFieldId = null; renderProperties(); }
    renderFields();
    saveToStorage();
}

function moveField(id, direction) {
    event?.stopPropagation();
    const index = state.fields.findIndex(f => f.id === id);
    if (index < 0) return;
    const newIndex = direction === 'up' ? index - 1 : index + 1;
    if (newIndex < 0 || newIndex >= state.fields.length) return;
    [state.fields[index], state.fields[newIndex]] = [state.fields[newIndex], state.fields[index]];
    renderFields();
    saveToStorage();
}

function selectField(id) {
    state.selectedFieldId = id;
    renderFields();
    renderProperties();
}

function updateField(prop, value) {
    const field = state.fields.find(f => f.id === state.selectedFieldId);
    if (field) { field[prop] = value; renderFields(); saveToStorage(); }
}

// Download sample cascade template
window.downloadCascadeTemplate = function() {
    // Create sample data with different hierarchy examples
    const sampleData = [
        // Headers
        ['Level_1', 'Level_2', 'Level_3', 'Level_4', 'UID'],
        // Example 1: Location hierarchy
        ['Eastern Region', 'Kailahun District', 'Luawa Chiefdom', 'Kailahun Gov Hospital', 'UID001'],
        ['Eastern Region', 'Kailahun District', 'Luawa Chiefdom', 'Luawa CHC', 'UID002'],
        ['Eastern Region', 'Kailahun District', 'Jawei Chiefdom', 'Daru CHC', 'UID003'],
        ['Eastern Region', 'Kenema District', 'Nongowa Chiefdom', 'Kenema Gov Hospital', 'UID004'],
        ['Southern Region', 'Bo District', 'Badjia Chiefdom', 'Ngelehun CHC', 'UID005'],
        ['Southern Region', 'Bo District', 'Bagbwe Chiefdom', 'Barlie MCHP', 'UID006'],
        ['Southern Region', 'Bonthe District', 'Jong Chiefdom', 'Mattru Hospital', 'UID007'],
        ['Western Region', 'Western Urban', 'East 1', 'Connaught Hospital', 'UID008'],
        ['Western Region', 'Western Rural', 'Koya Chiefdom', 'Waterloo CHC', 'UID009'],
    ];
    
    // Create workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(sampleData);
    
    // Set column widths
    ws['!cols'] = [
        { wch: 18 }, // Level_1
        { wch: 18 }, // Level_2
        { wch: 18 }, // Level_3
        { wch: 22 }, // Level_4
        { wch: 10 }, // UID
    ];
    
    XLSX.utils.book_append_sheet(wb, ws, 'CascadeData');
    
    // Download
    XLSX.writeFile(wb, 'cascade_template.xlsx');
    notify('Template downloaded! Edit columns and data as needed.', 'success');
};

// Handle cascade Excel file upload - creates separate linked fields
window.handleCascadeUpload = async function(input) {
    const file = input.files[0];
    if (!file) return;
    
    const currentField = state.fields.find(f => f.id === state.selectedFieldId);
    if (!currentField || currentField.type !== 'cascade') return;
    
    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        // Convert to JSON with headers
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        if (jsonData.length < 2) {
            notify('Excel file must have headers and at least one data row', 'error');
            return;
        }
        
        // First row is headers (column names)
        const columns = jsonData[0].map(h => String(h || '').trim()).filter(h => h);
        
        // Rest is data rows - filter out completely empty rows
        const rows = jsonData.slice(1).filter(row => {
            if (!row || !Array.isArray(row)) return false;
            return row.some(cell => cell !== null && cell !== undefined && cell !== '');
        });
        
        // Store as array of objects
        const cascadeData = rows.map(row => {
            const obj = {};
            columns.forEach((col, i) => {
                // Safely access row[i] - array might be shorter than columns
                const cellValue = Array.isArray(row) && i < row.length ? row[i] : null;
                obj[col] = cellValue !== null && cellValue !== undefined ? String(cellValue).trim() : '';
            });
            return obj;
        });
        
        // Generate unique group ID
        const cascadeGroupId = 'cascade_' + Date.now();
        
        // Find current field index
        const currentIndex = state.fields.findIndex(f => f.id === currentField.id);
        
        // Remove the cascade placeholder field
        state.fields.splice(currentIndex, 1);
        
        // Create separate select fields for each column
        const newFields = columns.map((col, idx) => ({
            id: cascadeGroupId + '_' + idx,
            type: 'select',
            label: col,
            name: col.toLowerCase().replace(/\s+/g, '_'),
            required: false,
            options: [], // Will be populated dynamically
            cascadeGroup: cascadeGroupId,
            cascadeLevel: idx,
            cascadeColumn: col,
            cascadeData: idx === 0 ? cascadeData : null, // Only first field stores data
            cascadeColumns: idx === 0 ? columns : null
        }));
        
        // Insert new fields at the same position
        state.fields.splice(currentIndex, 0, ...newFields);
        
        state.selectedFieldId = newFields[0].id;
        renderFields();
        renderProperties();
        saveToStorage();
        
        notify(`Created ${columns.length} linked cascade fields with ${cascadeData.length} data rows`, 'success');
    } catch (err) {
        console.error('Excel parse error:', err);
        notify('Error reading Excel file: ' + err.message, 'error');
    }
    
    // Reset input
    input.value = '';
};

function renderFields() {
    const dropZone = document.getElementById('dropZone');
    const dataFieldCount = state.fields.filter(f => f.type !== 'section').length;
    document.getElementById('fieldCount').innerHTML = `<span data-icon="bar-chart-3" data-size="12"></span> ${dataFieldCount} fields`;
    initIcons();
    
    if (state.fields.length === 0) {
        dropZone.classList.remove('has-fields');
        dropZone.innerHTML = '<p style="font-size:48px;margin-bottom:12px;"><span class="inline-icon">' + getIcon('inbox', 48) + '</span></p><p style="font-weight:600;">Click field types to add them</p><p style="font-size:11px;color:#868e96;margin-top:10px;">Add Period + Org Unit fields for DHIS2 sync</p>';
        return;
    }
    
    dropZone.classList.add('has-fields');
    
    dropZone.innerHTML = state.fields.map(f => {
        if (f.type === 'section') {
            return `<div class="section-divider ${f.id === state.selectedFieldId ? 'selected' : ''}" data-id="${f.id}">
                <span><span class="inline-icon">${getIcon('folder', 14)}</span> ${f.label.toUpperCase()}</span>
                <div style="display:flex;gap:5px;">
                    <button class="field-action-btn" onclick="moveField('${f.id}','up')"><span class="inline-icon">${getIcon('chevron-up', 12)}</span></button>
                    <button class="field-action-btn" onclick="moveField('${f.id}','down')"><span class="inline-icon">${getIcon('chevron-down', 12)}</span></button>
                    <button class="field-action-btn delete" onclick="removeField('${f.id}')"><span class="inline-icon">${getIcon('trash-2', 12)}</span></button>
                </div>
            </div>`;
        }
        
        const def = fieldDefs[f.type];
        const isDhis2Field = f.type === 'period';
        const isNumeric = f.type === 'number';
        const isCascade = f.cascadeGroup;
        const isCalc = f.type === 'calculation';
        
        return `<div class="form-field ${f.id === state.selectedFieldId ? 'selected' : ''} ${isDhis2Field ? 'dhis2-field' : ''} ${isCascade ? 'cascade-field' : ''}" data-id="${f.id}">
            <div class="form-field-header">
                <span class="form-field-label"><span class="inline-icon">${getIcon(def?.icon || 'file-text', 14)}</span> ${f.label}</span>
                <div class="form-field-actions">
                    <button class="field-action-btn" onclick="moveField('${f.id}','up')"><span class="inline-icon">${getIcon('chevron-up', 12)}</span></button>
                    <button class="field-action-btn" onclick="moveField('${f.id}','down')"><span class="inline-icon">${getIcon('chevron-down', 12)}</span></button>
                    <button class="field-action-btn delete" onclick="removeField('${f.id}')"><span class="inline-icon">${getIcon('trash-2', 12)}</span></button>
                </div>
            </div>
            <div style="font-size:10px;color:#868e96;">
                Code: <code>${f.name}</code>
                ${f.required ? '<span class="field-badge required">Required</span>' : ''}
                ${isDhis2Field ? '<span class="field-badge dhis2">DHIS2</span>' : ''}
                ${isNumeric ? '<span class="field-badge aggregate">SUM</span>' : ''}
                ${isCascade ? `<span class="field-badge cascade">Cascade L${(f.cascadeLevel || 0) + 1}</span>` : ''}
                ${isCalc ? '<span class="field-badge calc">CALC</span>' : ''}
            </div>
        </div>`;
    }).join('');
    
    dropZone.querySelectorAll('.form-field, .section-divider').forEach(el => {
        el.addEventListener('click', (e) => {
            if (!e.target.closest('.field-action-btn')) selectField(el.dataset.id);
        });
    });
}

function renderProperties() {
    const container = document.getElementById('propertiesContent');
    if (!state.selectedFieldId) {
        container.innerHTML = '<div class="no-selection"><p style="font-size:32px;margin-bottom:12px;"><span data-icon="chevron-up" data-size="32"></span></p><p>Select a field to edit</p></div>';
        return;
    }
    
    const field = state.fields.find(f => f.id === state.selectedFieldId);
    if (!field) return;
    
    // Initialize logic arrays if not exist
    if (!field.showLogic) field.showLogic = [];
    if (!field.validation) field.validation = [];
    
    let html = `
        <div class="prop-section">
            <div class="prop-section-title"><span data-icon="edit-3" data-size="12"></span> Field Settings</div>
            <div class="property-group">
                <label class="property-label">Label</label>
                <input type="text" class="property-input" id="propLabel" value="${escapeHtml(field.label)}">
            </div>
            <div class="property-group">
                <label class="property-label">Code / Name</label>
                <input type="text" class="property-input" id="propName" value="${escapeHtml(field.name)}">
            </div>
            ${field.type !== 'section' ? `
                <div class="property-group">
                    <label class="property-checkbox">
                        <input type="checkbox" id="propRequired" ${field.required ? 'checked' : ''}>
                        <span>Required Field</span>
                    </label>
                </div>
                <div class="property-group">
                    <label class="property-checkbox">
                        <input type="checkbox" id="propCheckDuplicate" ${field.checkDuplicate ? 'checked' : ''}>
                        <span>Check for Duplicates</span>
                    </label>
                    <p style="font-size:9px;color:#868e96;margin-top:4px;">Block submission if this value already exists in Google Sheets</p>
                </div>
            ` : ''}
        </div>
    `;
    
    if (['select', 'radio', 'checkbox'].includes(field.type) && !field.cascadeGroup) {
        html += `
            <div class="prop-section">
                <div class="prop-section-title"><span data-icon="list" data-size="12"></span> Options (one per line)</div>
                <textarea class="property-textarea" id="propOptions">${(field.options || []).join('\n')}</textarea>
            </div>
        `;
    }
    
    // Calculation field properties
    if (field.type === 'calculation') {
        const numberFields = state.fields.filter(f => f.type === 'number' && f.id !== field.id);
        html += `
            <div class="prop-section">
                <div class="prop-section-title"><span data-icon="calculator" data-size="12"></span> Formula Builder</div>
                
                <div class="property-group">
                    <label class="property-label">Formula</label>
                    <input type="text" class="property-input" id="propFormula" value="${escapeHtml(field.formula || '')}" placeholder="Click fields and operators below">
                </div>
                
                <!-- Operator Buttons -->
                <div style="margin:10px 0;">
                    <div style="font-size:10px;font-weight:700;color:#004080;margin-bottom:6px;">Operators:</div>
                    <div style="display:flex;flex-wrap:wrap;gap:4px;">
                        <button type="button" onclick="insertFormulaText(' + ')" class="calc-op-btn">+</button>
                        <button type="button" onclick="insertFormulaText(' - ')" class="calc-op-btn">−</button>
                        <button type="button" onclick="insertFormulaText(' * ')" class="calc-op-btn">×</button>
                        <button type="button" onclick="insertFormulaText(' / ')" class="calc-op-btn">÷</button>
                        <button type="button" onclick="insertFormulaText('(')" class="calc-op-btn">(</button>
                        <button type="button" onclick="insertFormulaText(')')" class="calc-op-btn">)</button>
                        <button type="button" onclick="insertFormulaText(' % ')" class="calc-op-btn">%</button>
                    </div>
                </div>
                
                <!-- Number Fields -->
                ${numberFields.length > 0 ? `
                    <div style="background:#d4edda;padding:10px;border-radius:6px;margin-top:10px;border:1px solid #c3e6cb;">
                        <div style="font-size:10px;font-weight:700;color:#155724;margin-bottom:8px;"><span data-icon="hash" data-size="10"></span> Click to Insert Number Fields:</div>
                        <div style="display:flex;flex-direction:column;gap:6px;">
                            ${numberFields.map(f => `
                                <div onclick="insertFieldRef('${f.name}')" style="background:white;padding:8px 10px;border-radius:6px;cursor:pointer;border:2px solid #28a745;display:flex;justify-content:space-between;align-items:center;" onmouseover="this.style.background='#e8f5e9'" onmouseout="this.style.background='white'">
                                    <span style="font-weight:600;font-size:11px;color:#155724;">${escapeHtml(f.label)}</span>
                                    <code style="background:#28a745;color:white;padding:2px 8px;border-radius:4px;font-size:10px;">{${f.name}}</code>
                                </div>
                            `).join('')}
                        </div>
                    </div>
                ` : `
                    <div style="background:#fff3cd;padding:12px;border-radius:6px;font-size:11px;color:#856404;margin-top:10px;border:1px solid #ffeeba;">
                        <span data-icon="alert-triangle" data-size="12"></span> <strong>No number fields available.</strong><br>
                        Add Number fields to your form first, then use them in calculations.
                    </div>
                `}
                
                <!-- Quick Numbers -->
                <div style="margin-top:10px;">
                    <div style="font-size:10px;font-weight:700;color:#004080;margin-bottom:6px;">Quick Numbers:</div>
                    <div style="display:flex;flex-wrap:wrap;gap:4px;">
                        ${[0, 1, 2, 5, 10, 100, 0.5, 0.1].map(n => `<button type="button" onclick="insertFormulaText('${n}')" class="calc-num-btn">${n}</button>`).join('')}
                    </div>
                </div>
                
                <!-- Formula Preview -->
                <div style="background:#f8f9fa;padding:10px;border-radius:6px;margin-top:10px;border:1px solid #dee2e6;">
                    <div style="font-size:10px;font-weight:700;color:#666;margin-bottom:4px;">Formula Preview:</div>
                    <div id="formulaPreview" style="font-family:monospace;font-size:12px;color:#004080;word-break:break-all;">${escapeHtml(field.formula || '(empty)')}</div>
                </div>
                
                <!-- Clear Button -->
                <button type="button" onclick="clearFormula()" style="margin-top:10px;padding:6px 12px;background:#dc3545;color:white;border:none;border-radius:4px;font-size:11px;cursor:pointer;width:100%;">
                    <span data-icon="trash-2" data-size="10"></span> Clear Formula
                </button>
            </div>
        `;
    }
    
    // Cascade field properties (show group info)
    if (field.type === 'cascade') {
        const cascadeColumns = field.cascadeColumns || [];
        const cascadeDataCount = (field.cascadeData || []).length;
        
        html += `
            <div class="prop-section">
                <div class="prop-section-title"><span data-icon="git-branch" data-size="12"></span> Cascade Data</div>
                <p style="font-size:10px;color:#666;margin-bottom:10px;">Upload Excel file with hierarchical data. Each column = dropdown level.</p>
                
                <div style="background:#f8f9fa;padding:10px;border-radius:6px;margin-bottom:12px;border:1px dashed #dee2e6;">
                    <div style="font-size:10px;font-weight:700;color:#004080;margin-bottom:6px;"><span data-icon="info" data-size="10"></span> Excel Format:</div>
                    <p style="font-size:9px;color:#666;margin-bottom:8px;">Row 1 = Column headers (dropdown labels)<br>Row 2+ = Data rows (each row = one complete path)</p>
                    <button class="logic-add-btn" style="background:#6c757d;" onclick="downloadCascadeTemplate()">
                        <span data-icon="download" data-size="10"></span> Download Sample Template
                    </button>
                </div>
                
                <div style="margin-bottom:12px;">
                    <input type="file" id="cascadeFileInput" accept=".xlsx,.xls" style="display:none;" onchange="handleCascadeUpload(this)">
                    <button class="logic-add-btn" style="background:#17a2b8;" onclick="document.getElementById('cascadeFileInput').click()">
                        <span data-icon="upload" data-size="10"></span> Upload Excel File
                    </button>
                </div>
                
                ${cascadeColumns.length > 0 ? `
                    <div style="background:#e8f4fc;padding:10px;border-radius:6px;margin-bottom:10px;">
                        <div style="font-size:10px;font-weight:700;color:#004080;margin-bottom:6px;">Loaded Columns:</div>
                        <div style="display:flex;flex-wrap:wrap;gap:6px;">
                            ${cascadeColumns.map((col, i) => `<span style="background:#004080;color:white;padding:3px 8px;border-radius:12px;font-size:10px;">${i+1}. ${col}</span>`).join('')}
                        </div>
                        <div style="font-size:9px;color:#666;margin-top:6px;">${cascadeDataCount} rows loaded</div>
                    </div>
                    
                    <div class="property-group">
                        <label class="property-label">Store Value From Column</label>
                        <select class="property-input" id="propCascadeValueColumn" onchange="updateField('cascadeValueColumn', this.value)">
                            ${cascadeColumns.map(col => `<option value="${col}" ${field.cascadeValueColumn === col ? 'selected' : ''}>${col}</option>`).join('')}
                        </select>
                        <p style="font-size:9px;color:#868e96;margin-top:4px;">Which column value to save (usually last column or UID)</p>
                    </div>
                ` : `
                    <div style="background:#fff3cd;padding:10px;border-radius:6px;font-size:11px;color:#856404;">
                        <span data-icon="alert-triangle" data-size="12"></span> No data loaded. Upload an Excel file with cascading data.
                    </div>
                `}
            </div>
        `;
    }
    
    // Get other fields that can be used for conditions (exclude current field and sections)
    const otherFields = state.fields.filter(f => f.id !== field.id && f.type !== 'section');
    const categoricalTypes = ['select', 'radio', 'yesno', 'checkbox'];
    const numericTypes = ['number'];
    
    // Show/Hide Logic Section
    html += `
        <div class="prop-section">
            <div class="prop-section-title"><span data-icon="eye" data-size="12"></span> Show/Hide Logic</div>
            <p style="font-size:10px;color:#666;margin-bottom:10px;">Show or hide this field based on other field values</p>
            
            <div id="logicRulesContainer">
                ${renderLogicRules(field.showLogic, otherFields)}
            </div>
            
            <div style="display:flex;gap:8px;margin-top:10px;">
                <button class="logic-add-btn" onclick="addLogicRule('show')"><span data-icon="eye" data-size="10"></span> Add SHOW Rule</button>
                <button class="logic-add-btn" style="background:#dc3545;" onclick="addLogicRule('hide')"><span data-icon="eye-off" data-size="10"></span> Add HIDE Rule</button>
            </div>
        </div>
    `;
    
    // Validation Section - different for each field type
    const validatableTypes = ['number', 'text', 'textarea', 'email', 'phone', 'date', 'time'];
    
    if (validatableTypes.includes(field.type)) {
        let validationHtml = '';
        
        if (field.type === 'number') {
            validationHtml = `
                <p style="font-size:10px;color:#666;margin-bottom:10px;">Set numeric constraints</p>
                <div id="validationRulesContainer">
                    ${renderNumericValidation(field.validation)}
                </div>
                <button class="logic-add-btn" style="background:#ffc107;color:#000;" onclick="addValidationRule('number')"><span data-icon="plus" data-size="10"></span> Add Rule</button>
            `;
        } else if (field.type === 'text' || field.type === 'textarea') {
            validationHtml = `
                <p style="font-size:10px;color:#666;margin-bottom:10px;">Set character limits</p>
                <div class="validation-rule">
                    <div class="logic-rule-row">
                        <span style="font-size:10px;font-weight:700;">Min characters:</span>
                        <input type="number" class="logic-input" min="0" value="${field.minLength || ''}" onchange="updateFieldValidation('minLength', this.value)">
                    </div>
                    <div class="logic-rule-row" style="margin-top:8px;">
                        <span style="font-size:10px;font-weight:700;">Max characters:</span>
                        <input type="number" class="logic-input" min="0" value="${field.maxLength || ''}" onchange="updateFieldValidation('maxLength', this.value)">
                    </div>
                    <div class="logic-rule-row" style="margin-top:8px;">
                        <span style="font-size:10px;font-weight:700;">Exact length:</span>
                        <input type="number" class="logic-input" min="0" value="${field.exactLength || ''}" onchange="updateFieldValidation('exactLength', this.value)" placeholder="e.g. 10">
                    </div>
                </div>
            `;
        } else if (field.type === 'email') {
            validationHtml = `
                <p style="font-size:10px;color:#666;margin-bottom:10px;">Email validation is automatic (must contain @)</p>
                <div class="validation-rule" style="background:#d4edda;border-color:#28a745;">
                    <div style="font-size:11px;color:#155724;">✓ Email format validation enabled (requires @)</div>
                </div>
            `;
        } else if (field.type === 'phone') {
            validationHtml = `
                <p style="font-size:10px;color:#666;margin-bottom:10px;">Set phone number format</p>
                <div class="validation-rule">
                    <div class="logic-rule-row">
                        <span style="font-size:10px;font-weight:700;">Exact digits:</span>
                        <input type="number" class="logic-input" min="1" max="20" value="${field.exactDigits || ''}" onchange="updateFieldValidation('exactDigits', this.value)" placeholder="e.g. 8">
                        <span style="font-size:9px;color:#666;">digits only</span>
                    </div>
                    <div class="logic-rule-row" style="margin-top:8px;">
                        <span style="font-size:10px;font-weight:700;">Min digits:</span>
                        <input type="number" class="logic-input" min="1" value="${field.minDigits || ''}" onchange="updateFieldValidation('minDigits', this.value)">
                    </div>
                    <div class="logic-rule-row" style="margin-top:8px;">
                        <span style="font-size:10px;font-weight:700;">Max digits:</span>
                        <input type="number" class="logic-input" min="1" value="${field.maxDigits || ''}" onchange="updateFieldValidation('maxDigits', this.value)">
                    </div>
                </div>
            `;
        } else if (field.type === 'date') {
            validationHtml = `
                <p style="font-size:10px;color:#666;margin-bottom:10px;">Set date restrictions</p>
                <div class="validation-rule">
                    <div class="logic-rule-row">
                        <label class="property-checkbox" style="padding:0;">
                            <input type="checkbox" ${field.noFutureDates ? 'checked' : ''} onchange="updateFieldValidation('noFutureDates', this.checked)">
                            <span style="font-size:10px;">No future dates (today or before)</span>
                        </label>
                    </div>
                    <div class="logic-rule-row" style="margin-top:8px;">
                        <label class="property-checkbox" style="padding:0;">
                            <input type="checkbox" ${field.noPastDates ? 'checked' : ''} onchange="updateFieldValidation('noPastDates', this.checked)">
                            <span style="font-size:10px;">No past dates (today or after)</span>
                        </label>
                    </div>
                    <div class="logic-rule-row" style="margin-top:8px;">
                        <span style="font-size:10px;font-weight:700;">Not before:</span>
                        <input type="date" class="logic-input" style="width:auto;" value="${field.minDate || ''}" onchange="updateFieldValidation('minDate', this.value)">
                    </div>
                    <div class="logic-rule-row" style="margin-top:8px;">
                        <span style="font-size:10px;font-weight:700;">Not after:</span>
                        <input type="date" class="logic-input" style="width:auto;" value="${field.maxDate || ''}" onchange="updateFieldValidation('maxDate', this.value)">
                    </div>
                </div>
            `;
        } else if (field.type === 'time') {
            validationHtml = `
                <p style="font-size:10px;color:#666;margin-bottom:10px;">Time field settings</p>
                <div class="validation-rule" style="background:#e8f4fc;border-color:#17a2b8;">
                    <div class="logic-rule-row">
                        <label class="property-checkbox" style="padding:0;">
                            <input type="checkbox" ${field.autoTime !== false ? 'checked' : ''} onchange="updateFieldValidation('autoTime', this.checked)">
                            <span style="font-size:10px;">Auto-fill with current time</span>
                        </label>
                    </div>
                </div>
            `;
        }
        
        html += `
            <div class="prop-section">
                <div class="prop-section-title"><span data-icon="check-circle" data-size="12"></span> Validation Rules</div>
                ${validationHtml}
            </div>
        `;
    }
    
    container.innerHTML = html;
    initIcons();
    
    document.getElementById('propLabel')?.addEventListener('change', (e) => updateField('label', e.target.value));
    document.getElementById('propName')?.addEventListener('change', (e) => updateField('name', e.target.value));
    document.getElementById('propRequired')?.addEventListener('change', (e) => updateField('required', e.target.checked));
    document.getElementById('propCheckDuplicate')?.addEventListener('change', (e) => updateField('checkDuplicate', e.target.checked));
    document.getElementById('propOptions')?.addEventListener('change', (e) => updateField('options', e.target.value.split('\n').filter(o => o.trim())));
    document.getElementById('propFormula')?.addEventListener('input', (e) => {
        updateField('formula', e.target.value);
        updateFormulaPreview(e.target.value);
    });
}

// Insert field reference into formula
window.insertFieldRef = function(fieldName) {
    const formulaInput = document.getElementById('propFormula');
    if (formulaInput) {
        const cursorPos = formulaInput.selectionStart || formulaInput.value.length;
        const before = formulaInput.value.substring(0, cursorPos);
        const after = formulaInput.value.substring(cursorPos);
        formulaInput.value = before + '{' + fieldName + '}' + after;
        formulaInput.focus();
        updateField('formula', formulaInput.value);
        updateFormulaPreview(formulaInput.value);
    }
};

// Insert text (operators, numbers) into formula
window.insertFormulaText = function(text) {
    const formulaInput = document.getElementById('propFormula');
    if (formulaInput) {
        const cursorPos = formulaInput.selectionStart || formulaInput.value.length;
        const before = formulaInput.value.substring(0, cursorPos);
        const after = formulaInput.value.substring(cursorPos);
        formulaInput.value = before + text + after;
        formulaInput.focus();
        updateField('formula', formulaInput.value);
        updateFormulaPreview(formulaInput.value);
    }
};

// Clear the formula
window.clearFormula = function() {
    const formulaInput = document.getElementById('propFormula');
    if (formulaInput) {
        formulaInput.value = '';
        updateField('formula', '');
        updateFormulaPreview('');
    }
};

// Update formula preview
function updateFormulaPreview(formula) {
    const preview = document.getElementById('formulaPreview');
    if (preview) {
        preview.textContent = formula || '(empty)';
    }
}

function renderLogicRules(rules, otherFields) {
    if (!rules || rules.length === 0) {
        return '<p style="font-size:10px;color:#868e96;text-align:center;padding:10px;">No logic rules. Field always visible.</p>';
    }
    
    const categoricalTypes = ['select', 'radio', 'yesno', 'checkbox'];
    
    return rules.map((rule, ruleIdx) => {
        const conditions = rule.conditions || [];
        
        return `
            <div class="logic-rule ${rule.action}">
                <div class="logic-rule-header">
                    <span class="logic-rule-type ${rule.action}">${rule.action.toUpperCase()} this field when:</span>
                    <button class="logic-rule-delete" onclick="deleteLogicRule(${ruleIdx})" title="Delete rule">✕</button>
                </div>
                
                ${conditions.map((cond, condIdx) => {
                    const sourceField = otherFields.find(f => f.name === cond.field);
                    const isCategorical = sourceField && categoricalTypes.includes(sourceField.type);
                    const isNumeric = sourceField && sourceField.type === 'number';
                    const options = sourceField?.type === 'yesno' ? ['Yes', 'No'] : (sourceField?.options || []);
                    
                    return `
                        ${condIdx > 0 ? `<div style="text-align:center;margin:5px 0;"><span class="logic-connector">${rule.connector || 'AND'}</span></div>` : ''}
                        <div class="logic-rule-row">
                            <select class="logic-select" style="flex:1;" onchange="updateLogicCondition(${ruleIdx}, ${condIdx}, 'field', this.value)">
                                <option value="">-- Select Field --</option>
                                ${otherFields.map(f => `<option value="${f.name}" ${cond.field === f.name ? 'selected' : ''}>${f.label}</option>`).join('')}
                            </select>
                            
                            ${isCategorical ? `
                                <select class="logic-select" onchange="updateLogicCondition(${ruleIdx}, ${condIdx}, 'operator', this.value)">
                                    <option value="=" ${cond.operator === '=' ? 'selected' : ''}>equals</option>
                                    <option value="!=" ${cond.operator === '!=' ? 'selected' : ''}>not equals</option>
                                </select>
                                <select class="logic-select" style="flex:1;" onchange="updateLogicCondition(${ruleIdx}, ${condIdx}, 'value', this.value)">
                                    <option value="">-- Select Value --</option>
                                    ${options.map(o => `<option value="${o}" ${cond.value === o ? 'selected' : ''}>${o}</option>`).join('')}
                                </select>
                            ` : isNumeric ? `
                                <select class="logic-select" onchange="updateLogicCondition(${ruleIdx}, ${condIdx}, 'operator', this.value)">
                                    <option value="=" ${cond.operator === '=' ? 'selected' : ''}>=</option>
                                    <option value="!=" ${cond.operator === '!=' ? 'selected' : ''}>≠</option>
                                    <option value=">" ${cond.operator === '>' ? 'selected' : ''}>&gt;</option>
                                    <option value=">=" ${cond.operator === '>=' ? 'selected' : ''}>≥</option>
                                    <option value="<" ${cond.operator === '<' ? 'selected' : ''}>&lt;</option>
                                    <option value="<=" ${cond.operator === '<=' ? 'selected' : ''}>≤</option>
                                    <option value="between" ${cond.operator === 'between' ? 'selected' : ''}>between</option>
                                </select>
                                ${cond.operator === 'between' ? `
                                    <input type="number" class="logic-input" placeholder="min" value="${cond.min || ''}" onchange="updateLogicCondition(${ruleIdx}, ${condIdx}, 'min', this.value)">
                                    <span style="font-size:10px;">and</span>
                                    <input type="number" class="logic-input" placeholder="max" value="${cond.max || ''}" onchange="updateLogicCondition(${ruleIdx}, ${condIdx}, 'max', this.value)">
                                ` : `
                                    <input type="number" class="logic-input" placeholder="value" value="${cond.value || ''}" onchange="updateLogicCondition(${ruleIdx}, ${condIdx}, 'value', this.value)">
                                `}
                            ` : `
                                <span style="font-size:10px;color:#868e96;">Select a field first</span>
                            `}
                            
                            <button class="logic-rule-delete" onclick="deleteLogicCondition(${ruleIdx}, ${condIdx})" title="Remove condition">✕</button>
                        </div>
                    `;
                }).join('')}
                
                <div style="display:flex;gap:8px;margin-top:8px;">
                    <button class="logic-add-condition" onclick="addLogicCondition(${ruleIdx}, 'AND')">+ AND</button>
                    <button class="logic-add-condition" onclick="addLogicCondition(${ruleIdx}, 'OR')">+ OR</button>
                </div>
            </div>
        `;
    }).join('');
}

function renderNumericValidation(rules) {
    if (!rules || rules.length === 0) {
        return '<p style="font-size:10px;color:#868e96;text-align:center;padding:10px;">No validation rules.</p>';
    }
    
    return rules.map((rule, idx) => `
        <div class="validation-rule">
            <div class="logic-rule-row">
                <span style="font-size:10px;font-weight:700;">Value must be</span>
                <select class="logic-select" onchange="updateValidationRule(${idx}, 'operator', this.value)">
                    <option value=">" ${rule.operator === '>' ? 'selected' : ''}>&gt;</option>
                    <option value=">=" ${rule.operator === '>=' ? 'selected' : ''}>≥</option>
                    <option value="<" ${rule.operator === '<' ? 'selected' : ''}>&lt;</option>
                    <option value="<=" ${rule.operator === '<=' ? 'selected' : ''}>≤</option>
                    <option value="=" ${rule.operator === '=' ? 'selected' : ''}>=</option>
                    <option value="!=" ${rule.operator === '!=' ? 'selected' : ''}>≠</option>
                    <option value="between" ${rule.operator === 'between' ? 'selected' : ''}>between</option>
                </select>
                ${rule.operator === 'between' ? `
                    <input type="number" class="logic-input" placeholder="min" value="${rule.min || ''}" onchange="updateValidationRule(${idx}, 'min', this.value)">
                    <span style="font-size:10px;">and</span>
                    <input type="number" class="logic-input" placeholder="max" value="${rule.max || ''}" onchange="updateValidationRule(${idx}, 'max', this.value)">
                ` : `
                    <input type="number" class="logic-input" placeholder="value" value="${rule.value || ''}" onchange="updateValidationRule(${idx}, 'value', this.value)">
                `}
                <button class="logic-rule-delete" onclick="deleteValidationRule(${idx})">✕</button>
            </div>
            <div class="validation-msg">
                <input type="text" class="property-input" style="font-size:10px;padding:5px;" placeholder="Error message (optional)" value="${rule.message || ''}" onchange="updateValidationRule(${idx}, 'message', this.value)">
            </div>
        </div>
    `).join('');
}

// Update field validation property directly
window.updateFieldValidation = function(prop, value) {
    const field = state.fields.find(f => f.id === state.selectedFieldId);
    if (!field) return;
    
    // Convert to number for numeric props, boolean for boolean props
    if (['minLength', 'maxLength', 'exactLength', 'exactDigits', 'minDigits', 'maxDigits'].includes(prop)) {
        field[prop] = value === '' ? null : parseInt(value);
    } else if (['noFutureDates', 'noPastDates', 'autoTime'].includes(prop)) {
        field[prop] = value;
    } else {
        field[prop] = value || null;
    }
    
    saveToStorage();
};

// Logic Management Functions
window.addLogicRule = function(action) {
    const field = state.fields.find(f => f.id === state.selectedFieldId);
    if (!field) return;
    if (!field.showLogic) field.showLogic = [];
    
    field.showLogic.push({
        action: action, // 'show' or 'hide'
        connector: 'AND',
        conditions: [{ field: '', operator: '=', value: '' }]
    });
    
    saveToStorage();
    renderProperties();
};

window.deleteLogicRule = function(ruleIdx) {
    const field = state.fields.find(f => f.id === state.selectedFieldId);
    if (!field || !field.showLogic) return;
    
    field.showLogic.splice(ruleIdx, 1);
    saveToStorage();
    renderProperties();
};

window.addLogicCondition = function(ruleIdx, connector) {
    const field = state.fields.find(f => f.id === state.selectedFieldId);
    if (!field || !field.showLogic || !field.showLogic[ruleIdx]) return;
    
    field.showLogic[ruleIdx].connector = connector;
    field.showLogic[ruleIdx].conditions.push({ field: '', operator: '=', value: '' });
    
    saveToStorage();
    renderProperties();
};

window.deleteLogicCondition = function(ruleIdx, condIdx) {
    const field = state.fields.find(f => f.id === state.selectedFieldId);
    if (!field || !field.showLogic || !field.showLogic[ruleIdx]) return;
    
    if (field.showLogic[ruleIdx].conditions.length <= 1) {
        // If only one condition, delete the whole rule
        field.showLogic.splice(ruleIdx, 1);
    } else {
        field.showLogic[ruleIdx].conditions.splice(condIdx, 1);
    }
    
    saveToStorage();
    renderProperties();
};

window.updateLogicCondition = function(ruleIdx, condIdx, prop, value) {
    const field = state.fields.find(f => f.id === state.selectedFieldId);
    if (!field || !field.showLogic || !field.showLogic[ruleIdx]) return;
    
    const condition = field.showLogic[ruleIdx].conditions[condIdx];
    if (!condition) return;
    
    condition[prop] = value;
    
    // If field changed, reset other values
    if (prop === 'field') {
        condition.operator = '=';
        condition.value = '';
        condition.min = '';
        condition.max = '';
    }
    
    // If operator changed to between, reset value
    if (prop === 'operator' && value === 'between') {
        condition.value = '';
    }
    
    saveToStorage();
    renderProperties();
};

// Validation Management Functions
window.addValidationRule = function(type) {
    const field = state.fields.find(f => f.id === state.selectedFieldId);
    if (!field) return;
    if (!field.validation) field.validation = [];
    
    field.validation.push({
        operator: '>=',
        value: '',
        min: '',
        max: '',
        message: ''
    });
    
    saveToStorage();
    renderProperties();
};

window.deleteValidationRule = function(idx) {
    const field = state.fields.find(f => f.id === state.selectedFieldId);
    if (!field || !field.validation) return;
    
    field.validation.splice(idx, 1);
    saveToStorage();
    renderProperties();
};

window.updateValidationRule = function(idx, prop, value) {
    const field = state.fields.find(f => f.id === state.selectedFieldId);
    if (!field || !field.validation || !field.validation[idx]) return;
    
    field.validation[idx][prop] = value;
    
    if (prop === 'operator' && value === 'between') {
        field.validation[idx].value = '';
    }
    
    saveToStorage();
    renderProperties();
};

// Evaluate Logic for a field given current form data
function evaluateFieldLogic(field, formData) {
    if (!field.showLogic || field.showLogic.length === 0) {
        return { visible: true, valid: true, errorMessage: '' };
    }
    
    let shouldShow = null; // null means no show rules
    let shouldHide = false;
    
    for (const rule of field.showLogic) {
        const result = evaluateRule(rule, formData);
        
        if (rule.action === 'show') {
            if (shouldShow === null) shouldShow = false;
            if (result) shouldShow = true;
        } else if (rule.action === 'hide') {
            if (result) shouldHide = true;
        }
    }
    
    // Hide takes precedence
    if (shouldHide) return { visible: false, valid: true, errorMessage: '' };
    
    // If there are show rules, field is only visible if at least one matches
    if (shouldShow !== null) return { visible: shouldShow, valid: true, errorMessage: '' };
    
    return { visible: true, valid: true, errorMessage: '' };
}

function evaluateRule(rule, formData) {
    const conditions = rule.conditions || [];
    if (conditions.length === 0) return false;
    
    const results = conditions.map(cond => evaluateCondition(cond, formData));
    
    if (rule.connector === 'OR') {
        return results.some(r => r);
    } else {
        return results.every(r => r);
    }
}

function evaluateCondition(condition, formData) {
    const fieldValue = formData[condition.field];
    const operator = condition.operator;
    
    if (fieldValue === undefined || fieldValue === null || fieldValue === '') {
        return false;
    }
    
    if (operator === 'between') {
        const num = parseFloat(fieldValue);
        const min = parseFloat(condition.min);
        const max = parseFloat(condition.max);
        return !isNaN(num) && !isNaN(min) && !isNaN(max) && num >= min && num <= max;
    }
    
    const compareValue = condition.value;
    
    // For numeric comparisons
    if (['>', '>=', '<', '<='].includes(operator)) {
        const num = parseFloat(fieldValue);
        const comp = parseFloat(compareValue);
        if (isNaN(num) || isNaN(comp)) return false;
        
        switch (operator) {
            case '>': return num > comp;
            case '>=': return num >= comp;
            case '<': return num < comp;
            case '<=': return num <= comp;
        }
    }
    
    // For equality
    if (operator === '=') {
        return String(fieldValue) === String(compareValue);
    }
    if (operator === '!=') {
        return String(fieldValue) !== String(compareValue);
    }
    
    return false;
}

function validateField(field, value) {
    if (value === '' || value === null || value === undefined) {
        return { valid: true, message: '' }; // Don't validate empty (use required for that)
    }
    
    // Numeric validation
    if (field.type === 'number') {
        const num = parseFloat(value);
        if (isNaN(num)) {
            return { valid: false, message: 'Please enter a valid number' };
        }
        
        if (field.validation && field.validation.length > 0) {
            for (const rule of field.validation) {
                let valid = true;
                
                if (rule.operator === 'between') {
                    const min = parseFloat(rule.min);
                    const max = parseFloat(rule.max);
                    valid = !isNaN(min) && !isNaN(max) && num >= min && num <= max;
                } else {
                    const comp = parseFloat(rule.value);
                    if (!isNaN(comp)) {
                        switch (rule.operator) {
                            case '>': valid = num > comp; break;
                            case '>=': valid = num >= comp; break;
                            case '<': valid = num < comp; break;
                            case '<=': valid = num <= comp; break;
                            case '=': valid = num === comp; break;
                            case '!=': valid = num !== comp; break;
                        }
                    }
                }
                
                if (!valid) {
                    const defaultMsg = rule.operator === 'between' 
                        ? `Value must be between ${rule.min} and ${rule.max}`
                        : `Value must be ${rule.operator} ${rule.value}`;
                    return { valid: false, message: rule.message || defaultMsg };
                }
            }
        }
    }
    
    // Text/Textarea validation
    if (field.type === 'text' || field.type === 'textarea') {
        const len = value.length;
        
        if (field.exactLength && len !== parseInt(field.exactLength)) {
            return { valid: false, message: `Must be exactly ${field.exactLength} characters` };
        }
        if (field.minLength && len < parseInt(field.minLength)) {
            return { valid: false, message: `Minimum ${field.minLength} characters required` };
        }
        if (field.maxLength && len > parseInt(field.maxLength)) {
            return { valid: false, message: `Maximum ${field.maxLength} characters allowed` };
        }
    }
    
    // Email validation
    if (field.type === 'email') {
        if (!value.includes('@') || !value.includes('.')) {
            return { valid: false, message: 'Please enter a valid email address' };
        }
        // Basic email pattern
        const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailPattern.test(value)) {
            return { valid: false, message: 'Please enter a valid email address' };
        }
    }
    
    // Phone validation
    if (field.type === 'phone') {
        // Extract only digits
        const digits = value.replace(/\D/g, '');
        
        if (field.exactDigits && digits.length !== parseInt(field.exactDigits)) {
            return { valid: false, message: `Phone number must be exactly ${field.exactDigits} digits` };
        }
        if (field.minDigits && digits.length < parseInt(field.minDigits)) {
            return { valid: false, message: `Phone number must have at least ${field.minDigits} digits` };
        }
        if (field.maxDigits && digits.length > parseInt(field.maxDigits)) {
            return { valid: false, message: `Phone number must have at most ${field.maxDigits} digits` };
        }
    }
    
    // Date validation
    if (field.type === 'date') {
        const dateValue = new Date(value);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        
        if (field.noFutureDates && dateValue > today) {
            return { valid: false, message: 'Future dates are not allowed' };
        }
        if (field.noPastDates && dateValue < today) {
            return { valid: false, message: 'Past dates are not allowed' };
        }
        if (field.minDate) {
            const minDate = new Date(field.minDate);
            if (dateValue < minDate) {
                return { valid: false, message: `Date must be on or after ${field.minDate}` };
            }
        }
        if (field.maxDate) {
            const maxDate = new Date(field.maxDate);
            if (dateValue > maxDate) {
                return { valid: false, message: `Date must be on or before ${field.maxDate}` };
            }
        }
    }
    
    return { valid: true, message: '' };
}

// Apply logic to all fields in the form
function applyFormLogic() {
    const form = document.getElementById('viewerForm');
    if (!form) return;
    
    // Get current form data
    const formData = {};
    new FormData(form).forEach((v, k) => {
        // Handle multiple values (checkboxes)
        if (formData[k]) {
            formData[k] = Array.isArray(formData[k]) ? [...formData[k], v] : [formData[k], v];
        } else {
            formData[k] = v;
        }
    });
    
    // Also get values from inputs that might not be in FormData yet
    form.querySelectorAll('input, select, textarea').forEach(el => {
        const name = el.name;
        if (!name) return;
        
        if (el.type === 'radio') {
            if (el.checked) formData[name] = el.value;
        } else if (el.type === 'checkbox') {
            if (el.checked && !formData[name]) {
                formData[name] = el.value;
            }
        } else {
            if (!formData[name]) formData[name] = el.value;
        }
    });
    
    // Build section visibility map - track which sections are hidden
    const hiddenSections = new Set();
    
    // First pass: evaluate section visibility
    state.fields.forEach(field => {
        if (field.type !== 'section') return;
        
        const logicResult = evaluateFieldLogic(field, formData);
        if (!logicResult.visible) {
            hiddenSections.add(field.id);
        }
    });
    
    // Build a map of which section each field belongs to
    let currentSectionId = null;
    const fieldSectionMap = {};
    state.fields.forEach(field => {
        if (field.type === 'section') {
            currentSectionId = field.id;
        } else {
            fieldSectionMap[field.id] = currentSectionId;
        }
    });
    
    // Apply logic to each field
    state.fields.forEach(field => {
        // For sections - find and hide the page that has this section ID
        if (field.type === 'section') {
            const logicResult = evaluateFieldLogic(field, formData);
            const page = document.querySelector(`.form-page[data-section-id="${field.id}"]`);
            if (page) {
                if (logicResult.visible) {
                    page.classList.remove('section-hidden');
                } else {
                    page.classList.add('section-hidden');
                    // Also clear all fields in this hidden page
                    page.querySelectorAll('input, select, textarea').forEach(input => {
                        if (input.type === 'radio' || input.type === 'checkbox') {
                            input.checked = false;
                        } else if (input.type !== 'time') {
                            input.value = '';
                        }
                    });
                }
            }
            return;
        }
        
        const fieldEl = document.querySelector(`[data-field-name="${field.name}"]`);
        if (!fieldEl) return;
        
        // Check if field's section is hidden
        const fieldSectionId = fieldSectionMap[field.id];
        const sectionHidden = fieldSectionId && hiddenSections.has(fieldSectionId);
        
        // Evaluate field's own logic
        const logicResult = evaluateFieldLogic(field, formData);
        
        // Field is visible only if both its section is visible AND its own logic says visible
        const shouldBeVisible = logicResult.visible && !sectionHidden;
        
        // Show/Hide field
        if (shouldBeVisible) {
            fieldEl.classList.remove('field-hidden');
        } else {
            fieldEl.classList.add('field-hidden');
            // Clear hidden field values
            const inputs = fieldEl.querySelectorAll('input, select, textarea');
            inputs.forEach(input => {
                if (input.type === 'radio' || input.type === 'checkbox') {
                    input.checked = false;
                } else if (input.type !== 'time') { // Don't clear auto-time
                    input.value = '';
                }
            });
        }
        
        // Apply validation styling for validatable fields
        const validatableTypes = ['number', 'text', 'textarea', 'email', 'phone', 'date'];
        if (validatableTypes.includes(field.type)) {
            const input = fieldEl.querySelector('input, textarea');
            const errorDiv = fieldEl.querySelector('.field-error-msg');
            
            if (input && input.value) {
                const validResult = validateField(field, input.value);
                if (!validResult.valid) {
                    fieldEl.classList.add('field-error');
                    if (errorDiv) errorDiv.textContent = validResult.message;
                } else {
                    fieldEl.classList.remove('field-error');
                    if (errorDiv) errorDiv.textContent = '';
                }
            } else {
                fieldEl.classList.remove('field-error');
                if (errorDiv) errorDiv.textContent = '';
            }
        }
    });
    
    // Also handle sections (hide section if all its fields are hidden)
    let currentSection = null;
    let allFieldsHidden = true;
    
    state.fields.forEach((field, idx) => {
        if (field.type === 'section') {
            // Check if previous section should be hidden
            if (currentSection && allFieldsHidden) {
                const sectionEl = document.querySelector(`[data-field-name="${currentSection.name}"]`);
                if (sectionEl) sectionEl.classList.add('field-hidden');
            }
            currentSection = field;
            allFieldsHidden = true;
        } else if (currentSection) {
            const fieldEl = document.querySelector(`[data-field-name="${field.name}"]`);
            if (fieldEl && !fieldEl.classList.contains('field-hidden')) {
                allFieldsHidden = false;
            }
        }
    });
    
    // Check if currently displayed page is now empty and should auto-skip
    checkCurrentPageVisibility();
}

// Check if current page has visible fields, if not, navigate to next valid page
function checkCurrentPageVisibility() {
    const pages = document.querySelectorAll('.form-page');
    const visiblePage = Array.from(pages).find(p => p.style.display !== 'none');
    
    if (!visiblePage) return;
    
    // Check if this page has any visible fields
    const visibleFields = visiblePage.querySelectorAll('.viewer-field:not(.field-hidden)');
    
    if (visibleFields.length === 0) {
        console.log('Current page has no visible fields, finding next valid page...');
        
        // Find the index of the current page
        const currentIndex = parseInt(visiblePage.dataset.page || '0');
        
        // Try to find next page with visible fields
        let nextIndex = currentIndex + 1;
        while (nextIndex < pages.length) {
            const nextPage = pages[nextIndex];
            if (!nextPage.classList.contains('section-hidden')) {
                const nextFields = nextPage.querySelectorAll('.viewer-field:not(.field-hidden)');
                if (nextFields.length > 0) {
                    // Found a valid page, navigate to it
                    console.log('Auto-navigating to page', nextIndex);
                    pages.forEach(p => p.style.display = 'none');
                    nextPage.style.display = 'block';
                    return;
                }
            }
            nextIndex++;
        }
        
        // If no next page found, try previous pages
        let prevIndex = currentIndex - 1;
        while (prevIndex >= 0) {
            const prevPage = pages[prevIndex];
            if (!prevPage.classList.contains('section-hidden')) {
                const prevFields = prevPage.querySelectorAll('.viewer-field:not(.field-hidden)');
                if (prevFields.length > 0) {
                    console.log('Auto-navigating back to page', prevIndex);
                    pages.forEach(p => p.style.display = 'none');
                    prevPage.style.display = 'block';
                    return;
                }
            }
            prevIndex--;
        }
    }
}

// Validate all visible fields before submit
function validateAllFields() {
    const form = document.getElementById('viewerForm');
    if (!form) return { valid: true };
    
    const formData = {};
    new FormData(form).forEach((v, k) => formData[k] = v);
    
    for (const field of state.fields) {
        if (field.type === 'section') continue;
        
        const fieldEl = document.querySelector(`[data-field-name="${field.name}"]`);
        if (!fieldEl || fieldEl.classList.contains('field-hidden')) continue;
        
        // Check required
        if (field.required) {
            const value = formData[field.name];
            if (!value || value === '') {
                return { valid: false, message: `${field.label} is required` };
            }
        }
        
        // Check validation rules for all validatable field types
        const validatableTypes = ['number', 'text', 'textarea', 'email', 'phone', 'date'];
        if (validatableTypes.includes(field.type)) {
            const value = formData[field.name];
            if (value) {
                const result = validateField(field, value);
                if (!result.valid) {
                    return { valid: false, message: `${field.label}: ${result.message}` };
                }
            }
        }
    }
    
    return { valid: true };
}

// ==================== STORAGE ====================
function saveToStorage() {
    safeStorage.setItem('icfCollectForm', JSON.stringify({
        fields: state.fields,
        settings: state.settings,
        fieldCounter: state.fieldCounter,
        collectedData: state.collectedData,
        filterOrder: state.filterOrder,
        dhis2: {
            datasetId: state.dhis2.datasetId,
            programId: state.dhis2.programId,
            dataElements: state.dhis2.dataElements
        }
    }));
}

function loadFromStorage() {
    const saved = safeStorage.getItem('icfCollectForm');
    if (saved) {
        try {
            const data = JSON.parse(saved);
            state.fields = data.fields || [];
            state.settings = { ...state.settings, ...data.settings };
            state.fieldCounter = data.fieldCounter || 0;
            state.collectedData = data.collectedData || [];
            state.filterOrder = data.filterOrder || [];
            if (data.dhis2) {
                state.dhis2.datasetId = data.dhis2.datasetId;
                state.dhis2.programId = data.dhis2.programId;
                state.dhis2.dataElements = data.dhis2.dataElements || {};
            }
            
            // Ensure formId and originalTitle are set
            if (!state.settings.formId) {
                state.settings.formId = 'form_' + Date.now();
            }
            state.settings.originalTitle = state.settings.title;
            
            document.getElementById('formTitle').value = state.settings.title;
            document.getElementById('previewTitle').textContent = state.settings.title;
        } catch (e) {}
    }
}

function loadConfigs() {
    // Load Sheets config
    const sheetsConfig = safeStorage.getItem('icfSheetsConfig');
    if (sheetsConfig) {
        try {
            const config = JSON.parse(sheetsConfig);
            state.sheets = { ...state.sheets, ...config };
        } catch (e) {}
    }
    
    // Load DHIS2 config
    const dhis2Config = safeStorage.getItem('icfDhis2Config');
    if (dhis2Config) {
        try {
            const config = JSON.parse(dhis2Config);
            if (!state.dhis2) {
                state.dhis2 = { dataElements: {}, datasetId: null, programId: '', url: '', username: '', password: '', syncMode: 'aggregate', orgUnitLevel: 5, periodType: 'Monthly', periodColumn: '', orgUnitColumn: '', connected: false, orgUnits: [], orgUnitMap: {} };
            }
            state.dhis2 = { ...state.dhis2, ...config };
        } catch (e) {}
    }
}

// ==================== FORM OPERATIONS ====================
function newForm() {
    if (state.fields.length > 0 && !confirm('Create new form?')) return;
    state.fields = [];
    state.selectedFieldId = null;
    state.fieldCounter = 0;
    
    // Ensure dhis2 object exists
    if (!state.dhis2) {
        state.dhis2 = { dataElements: {}, datasetId: null, programId: '', url: '', username: '', password: '', syncMode: 'aggregate', orgUnitLevel: 5, periodType: 'Monthly', periodColumn: '', orgUnitColumn: '', connected: false, orgUnits: [], orgUnitMap: {} };
    }
    state.dhis2.dataElements = {};
    state.dhis2.datasetId = null;
    state.collectedData = [];
    
    const title = prompt('Form Name:', 'New Form');
    if (title) {
        state.settings.title = title;
        state.settings.originalTitle = title; // Track original title
        state.settings.formId = 'form_' + Date.now(); // Generate unique ID
        document.getElementById('formTitle').value = title;
        document.getElementById('previewTitle').textContent = title;
    }
    saveToStorage();
    renderFields();
    renderProperties();
    notify('New form created!');
}

function saveCurrentForm() {
    if (state.fields.length === 0) { notify('Add fields first!', 'error'); return; }
    saveToStorage();
    saveFormToList();
    notify('Form saved!');
}

function saveFormToList() {
    const forms = JSON.parse(safeStorage.getItem('icfCollectForms') || '[]');
    
    // Use existing formId from state, or find by title, or generate new
    let formId = state.settings.formId;
    if (!formId) {
        const existingForm = forms.find(f => f.title === state.settings.title || f.id === state.settings.formId);
        formId = existingForm ? existingForm.id : 'form_' + Date.now();
        state.settings.formId = formId;
    }
    
    const existingIndex = forms.findIndex(f => f.id === formId);
    const formEntry = {
        id: formId,
        title: state.settings.title,
        fieldCount: state.fields.filter(f => f.type !== 'section').length,
        updatedAt: new Date().toISOString(),
        fields: state.fields,
        settings: state.settings,
        collectedData: state.collectedData,
        dhis2: {
            datasetId: state.dhis2?.datasetId,
            programId: state.dhis2?.programId,
            dataElements: state.dhis2?.dataElements
        }
    };
    if (existingIndex >= 0) forms[existingIndex] = formEntry;
    else forms.push(formEntry);
    safeStorage.setItem('icfCollectForms', JSON.stringify(forms));
    
    // Update original title after save
    state.settings.originalTitle = state.settings.title;
    
    // Also save to cloud if connected
    saveFormToCloud(formEntry);
}

// Save form to Google Sheets cloud storage
async function saveFormToCloud(formEntry) {
    if (!state.sheets.scriptUrl || !state.user) return;
    
    try {
        const response = await fetch(state.sheets.scriptUrl, {
            method: 'POST', mode: 'cors', redirect: 'follow',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action: 'saveForm',
                email: state.user.email,
                formId: formEntry.id,
                formData: JSON.stringify(formEntry)
            })
        });
        const result = await response.json();
        if (result.success) {
            console.log('Form saved to cloud');
        }
    } catch (err) {
        console.error('Cloud save error:', err);
    }
}

// Load all forms from Google Sheets cloud storage
async function loadFormsFromCloud() {
    if (!state.sheets.scriptUrl || !state.user) return [];
    
    try {
        const response = await fetch(state.sheets.scriptUrl, {
            method: 'POST', mode: 'cors', redirect: 'follow',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action: 'loadForms',
                email: state.user.email
            })
        });
        const result = await response.json();
        if (result.success && result.forms) {
            return result.forms;
        }
    } catch (err) {
        console.error('Cloud load error:', err);
    }
    return [];
}

// Sync forms between local storage and cloud
async function syncForms() {
    if (!state.user) return;
    
    const localForms = JSON.parse(safeStorage.getItem('icfCollectForms') || '[]');
    const cloudForms = await loadFormsFromCloud();
    
    // Build cloud forms map for quick lookup
    const cloudFormsMap = {};
    cloudForms.forEach(f => { cloudFormsMap[f.id] = f; });
    
    // Merge: use most recently updated version
    const mergedForms = {};
    const formsToUpload = [];
    
    localForms.forEach(f => {
        mergedForms[f.id] = f;
        // Check if local is newer than cloud
        const cloudVersion = cloudFormsMap[f.id];
        if (!cloudVersion || new Date(f.updatedAt) > new Date(cloudVersion.updatedAt)) {
            formsToUpload.push(f);
        }
    });
    
    cloudForms.forEach(f => {
        if (!mergedForms[f.id] || new Date(f.updatedAt) > new Date(mergedForms[f.id].updatedAt)) {
            mergedForms[f.id] = f;
        }
    });
    
    const finalForms = Object.values(mergedForms);
    safeStorage.setItem('icfCollectForms', JSON.stringify(finalForms));
    
    // Only push forms that are newer locally (in parallel for speed)
    if (formsToUpload.length > 0) {
        console.log(`Uploading ${formsToUpload.length} changed forms...`);
        await Promise.all(formsToUpload.map(form => saveFormToCloud(form)));
    }
    
    return finalForms;
}

// Delete form from cloud
async function deleteFormFromCloud(formId, formTitle) {
    if (!state.sheets.scriptUrl || !state.user) return;
    
    try {
        await fetch(state.sheets.scriptUrl, {
            method: 'POST', mode: 'cors', redirect: 'follow',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action: 'deleteForm',
                email: state.user.email,
                formId: formId,
                formTitle: formTitle || '', // Also delete the data sheet
                deleteSheet: 'true'
            })
        });
        console.log('Form and sheet deleted from cloud');
    } catch (err) {
        console.error('Cloud delete error:', err);
    }
}

// Save large cascade data to Google Sheets cloud storage
async function saveCascadeDataToCloud(cascadeId, compressedData, columns) {
    // Use CONFIG.AUTH_SCRIPT_URL or fallback to state.sheets.scriptUrl
    const scriptUrl = CONFIG.AUTH_SCRIPT_URL || state.sheets.scriptUrl;
    if (!scriptUrl) {
        throw new Error('No script URL configured');
    }
    
    console.log('Saving cascade to:', scriptUrl);
    console.log('Cascade ID:', cascadeId, 'Data length:', compressedData.length);
    
    // Check if data is too large
    if (compressedData.length > 5000000) {
        throw new Error('Cascade data too large (>5MB). Please reduce the number of rows.');
    }
    
    try {
        // Build URL with params, send data in POST body
        const url = new URL(scriptUrl);
        url.searchParams.set('action', 'saveCascadeData');
        url.searchParams.set('cascadeId', cascadeId);
        url.searchParams.set('columns', JSON.stringify(columns));
        
        console.log('Sending POST to:', url.toString());
        
        const response = await fetch(url.toString(), {
            method: 'POST', mode: 'cors', redirect: 'follow',
            headers: { 'Content-Type': 'text/plain' },
            body: compressedData
        });
        
        console.log('Response status:', response.status);
        
        if (!response.ok) {
            const text = await response.text();
            console.error('Response text:', text);
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        const result = await response.json();
        console.log('Response result:', result);
        
        if (!result.success) {
            throw new Error(result.error || 'Failed to save cascade data');
        }
        console.log('Cascade data saved to cloud:', cascadeId);
        return true;
    } catch (err) {
        console.error('Cascade cloud save error:', err);
        if (err.name === 'TypeError' && err.message === 'Failed to fetch') {
            throw new Error('Network error - check internet connection and that Apps Script is deployed');
        }
        throw err;
    }
}

// Load cascade data from Google Sheets cloud storage
async function loadCascadeDataFromCloud(cascadeId, columns) {
    // First check localStorage cache
    const cacheKey = 'cascade_cache_' + cascadeId;
    const cached = safeStorage.getItem(cacheKey);
    if (cached) {
        try {
            const cachedData = JSON.parse(cached);
            console.log('Cascade data loaded from cache:', cascadeId, cachedData.data.length, 'rows');
            return cachedData;
        } catch (e) {
            safeStorage.removeItem(cacheKey);
        }
    }
    
    const scriptUrl = CONFIG.AUTH_SCRIPT_URL;
    if (!scriptUrl) {
        console.log('No script URL, cannot load cascade data');
        return null;
    }
    
    try {
        const response = await fetch(scriptUrl + '?' + new URLSearchParams({
            action: 'getCascadeData',
            cascadeId: cascadeId
        }), {
            mode: 'cors',
            redirect: 'follow'
        });
        const result = await response.json();
        if (result.success && result.data) {
            // Decompress the data
            const cols = result.columns || columns || [];
            const decompressed = decompressCascadeData(result.data, cols);
            console.log('Cascade data loaded from cloud:', cascadeId, decompressed.length, 'rows');
            
            // Cache in localStorage for offline use
            const cacheData = { data: decompressed, columns: cols };
            try {
                safeStorage.setItem(cacheKey, JSON.stringify(cacheData));
                console.log('Cascade data cached locally');
            } catch (e) {
                console.log('Could not cache cascade data (localStorage full?)');
            }
            
            return cacheData;
        }
    } catch (err) {
        console.error('Cascade load error:', err);
    }
    return null;
}

function previewForm() {
    if (state.fields.length === 0) { notify('Add fields!', 'error'); return; }
    state.isSharedMode = false;
    document.querySelector('.header').style.display = 'none';
    document.getElementById('mainContainer').classList.remove('show');
    document.querySelector('.footer').style.display = 'none';
    document.getElementById('viewerContainer').classList.add('show');
    renderFormViewer();
}

function closeViewer() {
    document.getElementById('viewerContainer').classList.remove('show');
    document.querySelector('.header').style.display = '';
    document.getElementById('mainContainer').classList.add('show');
    document.querySelector('.footer').style.display = '';
}

async function shareForm() {
    if (state.fields.length === 0) { notify('Add fields!', 'error'); return; }
    
    // Ensure settings exist
    if (!state.settings) {
        state.settings = { title: 'Form', logo: CONFIG.LOGO_URL, aggregateColumn: '' };
    }
    
    try {
        console.log('Starting shareForm...');
        console.log('Fields count:', state.fields.length);
        
        // Debug: check cascade fields
        state.fields.forEach((f, idx) => {
            if (f.cascadeData && f.cascadeData.length > 0) {
                console.log(`Field ${idx} (${f.name}): cascadeData=${f.cascadeData.length} rows, cascadeColumns=`, f.cascadeColumns);
            }
        });
        
        const formData = {
            s: { 
                t: state.settings.title || 'Form', 
                l: state.settings.logo || CONFIG.LOGO_URL, 
                ac: state.settings.aggregateColumn || '' 
            },
            f: state.fields.map((f, idx) => {
                try {
                    console.log(`Mapping field ${idx}: ${f.name}, type: ${f.type}`);
                    
                    // For cascade data, save to cloud and use reference ID
                    let compressedCascade = null;
                    let cascadeRef = null;
                    if (f.cascadeData && Array.isArray(f.cascadeData) && f.cascadeData.length > 0 && 
                        f.cascadeColumns && Array.isArray(f.cascadeColumns) && f.cascadeColumns.length > 0) {
                        // Generate a unique reference ID for this cascade data
                        cascadeRef = f.cascadeDataRef || ('csc_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6));
                        f.cascadeDataRef = cascadeRef; // Store for future use
                        console.log(`Cascade field ${idx}: ${f.cascadeData.length} rows will be saved with ref: ${cascadeRef}`);
                    }
                    
                    // Safely get options
                    let safeOptions = [];
                    if (f.options) {
                        if (Array.isArray(f.options)) {
                            safeOptions = f.options;
                        } else if (typeof f.options === 'string') {
                            safeOptions = f.options.split('\n').filter(o => o.trim());
                        }
                    }
                    
                    return { 
                        i: f.id || '', 
                        y: f.type || '', 
                        l: f.label || '', 
                        n: f.name || '', 
                        r: f.required || false, 
                        o: safeOptions, 
                        mx: f.max || 5,
                        // Validation fields - ensure arrays
                        sl: Array.isArray(f.showLogic) ? f.showLogic : [],
                        v: Array.isArray(f.validation) ? f.validation : [],
                        mnl: f.minLength, mxl: f.maxLength, el: f.exactLength,
                        ed: f.exactDigits, mnd: f.minDigits, mxd: f.maxDigits,
                        nfd: f.noFutureDates, npd: f.noPastDates, mid: f.minDate, mxd2: f.maxDate,
                        at: f.autoTime,
                        cd: f.checkDuplicate || false,
                        // Cascade linked fields
                        cg: f.cascadeGroup || null,
                        cl: typeof f.cascadeLevel === 'number' ? f.cascadeLevel : null,
                        ccol: f.cascadeColumn || null,
                        cc: Array.isArray(f.cascadeColumns) && f.cascadeColumns.length > 0 ? f.cascadeColumns : null,
                        cref: cascadeRef, // Reference ID to load cascade data from cloud
                        cvc: f.cascadeValueColumn || null,
                        // Calculation field
                        frm: f.formula || null
                    };
                } catch (fieldErr) {
                    console.error(`Error mapping field ${idx} (${f?.name || 'unknown'}):`, fieldErr);
                    console.error('Field data:', JSON.stringify(f, null, 2));
                    throw fieldErr;
                }
            }),
            d: (state.dhis2?.datasetId || state.dhis2?.programId) ? { 
                id: state.dhis2.datasetId,
                pid: state.dhis2.programId
                // Note: dataElements not included - too large for URL, will be fetched from DHIS2
            } : null,
        // Include DHIS2 credentials for sync
        h: state.dhis2?.url ? {
            u: state.dhis2.url,
            n: state.dhis2.username,
            p: btoa(state.dhis2.password || ''), // Encode password
            m: state.dhis2.syncMode,
            ol: state.dhis2.orgUnitLevel,
            pt: state.dhis2.periodType,
            pc: state.dhis2.periodColumn,
            oc: state.dhis2.orgUnitColumn
        } : null
    };
    
        // Save cascade data to Google Sheets
        const cascadeFieldsToSave = state.fields.filter(f => 
            f.cascadeData && f.cascadeData.length > 0 && 
            f.cascadeColumns && f.cascadeColumns.length > 0 &&
            f.cascadeDataRef
        );
        
        if (cascadeFieldsToSave.length > 0) {
            console.log(`Saving ${cascadeFieldsToSave.length} cascade datasets to cloud...`);
            notify('Saving cascade data to cloud...', 'info');
            
            for (const field of cascadeFieldsToSave) {
                try {
                    const compressed = compressCascadeData(field.cascadeData, field.cascadeColumns);
                    await saveCascadeDataToCloud(field.cascadeDataRef, compressed, field.cascadeColumns);
                    console.log(`Saved cascade data: ${field.cascadeDataRef}`);
                } catch (err) {
                    console.error('Failed to save cascade data:', err);
                    notify('Error saving cascade data: ' + err.message, 'error');
                    return;
                }
            }
        }
    
        // Clean up undefined values to reduce URL size
        console.log('Cleaning up form data...');
        formData.f = formData.f.map((f, idx) => {
            try {
                const cleaned = {};
                Object.keys(f).forEach(k => {
                    if (f[k] !== null && f[k] !== undefined && f[k] !== '') {
                        cleaned[k] = f[k];
                    }
                });
                return cleaned;
            } catch (cleanErr) {
                console.error(`Error cleaning field ${idx}:`, cleanErr);
                throw cleanErr;
            }
        });
        
        console.log('Stringifying form data...');
        
        // Debug: log sizes
        console.log('Settings size:', JSON.stringify(formData.s).length);
        console.log('DHIS2 d size:', JSON.stringify(formData.d).length);
        console.log('DHIS2 h size:', JSON.stringify(formData.h).length);
        console.log('Fields count:', formData.f.length);
        
        // Debug: log field sizes to see what's large
        formData.f.forEach((field, idx) => {
            const fieldStr = JSON.stringify(field);
            console.log(`Field ${idx} (${field.n}): ${fieldStr.length} chars`);
            if (fieldStr.length > 500) {
                console.log('Large field keys:', Object.keys(field).filter(k => JSON.stringify(field[k]).length > 100));
            }
        });
        
        const jsonStr = JSON.stringify(formData);
        console.log('Total JSON string length:', jsonStr.length);
        
        console.log('Compressing with pako...');
        const compressed = pako.deflate(jsonStr);
        console.log('Compressed length:', compressed.length);
        
        // Convert Uint8Array to base64 safely (handle large arrays)
        console.log('Converting to base64...');
        let binary = '';
        const chunkSize = 8192;
        for (let i = 0; i < compressed.length; i += chunkSize) {
            const chunk = compressed.subarray(i, i + chunkSize);
            binary += String.fromCharCode.apply(null, Array.from(chunk));
        }
        const encoded = btoa(binary);
        console.log('Base64 length:', encoded.length);
        
        const shareUrl = window.location.origin + window.location.pathname + '?d=' + encodeURIComponent(encoded);
        
        // Check URL length and show info
        const urlLen = shareUrl.length;
        console.log('Share URL length:', urlLen, 'characters');
        
        if (urlLen > 100000) {
            notify(`URL too large (${Math.round(urlLen/1000)}KB). Reduce cascade data rows.`, 'error');
            return;
        } else if (urlLen > 50000) {
            notify(`Large URL (${Math.round(urlLen/1000)}KB). May not work in all browsers.`, 'warning');
        } else if (urlLen > 8000) {
            notify(`URL: ${Math.round(urlLen/1000)}KB - Works in modern browsers`, 'info');
        }
        
        const shareUrlEl = document.getElementById('shareUrl');
        if (shareUrlEl) shareUrlEl.textContent = shareUrl;
        
        saveFormToList();
        const modal = document.getElementById('shareModal');
        if (modal) modal.classList.add('show');
    } catch (err) { 
        console.error('Share error:', err);
        notify('Error: ' + err.message, 'error'); 
    }
}

function copyShareUrl() {
    navigator.clipboard.writeText(document.getElementById('shareUrl').textContent)
        .then(() => notify('Copied to clipboard!'))
        .catch(() => notify('Copy failed', 'error'));
}

// ==================== HOME ====================
async function showHome() {
    document.querySelector('.header').style.display = 'none';
    document.getElementById('mainContainer').classList.remove('show');
    document.querySelector('.footer').style.display = 'none';
    document.getElementById('homeContainer').classList.add('show');
    
    // Show local forms IMMEDIATELY (no waiting)
    renderHomeContent(false);
    
    // Sync with cloud in BACKGROUND (don't block UI)
    if (state.user) {
        syncFormsInBackground();
    }
}

async function syncFormsInBackground() {
    try {
        console.log('Background sync started...');
        const cloudForms = await loadFormsFromCloud();
        
        if (cloudForms.length > 0) {
            const localForms = JSON.parse(safeStorage.getItem('icfCollectForms') || '[]');
            const mergedForms = {};
            
            localForms.forEach(f => { mergedForms[f.id] = f; });
            cloudForms.forEach(f => {
                if (!mergedForms[f.id] || new Date(f.updatedAt) > new Date(mergedForms[f.id].updatedAt)) {
                    mergedForms[f.id] = f;
                }
            });
            
            const finalForms = Object.values(mergedForms);
            safeStorage.setItem('icfCollectForms', JSON.stringify(finalForms));
            
            // Re-render only if we're still on home page
            if (document.getElementById('homeContainer').classList.contains('show')) {
                renderHomeContent(true);
            }
        }
        console.log('Background sync complete');
    } catch (err) {
        console.error('Background sync error:', err);
    }
}

function renderHomeContent(synced) {
    const forms = JSON.parse(safeStorage.getItem('icfCollectForms') || '[]');
    const isLoggedIn = !!state.user;
    
    let formsHtml = forms.length === 0 ? `
        <div style="text-align:center;padding:60px;color:#868e96;">
            <p style="font-size:64px;margin-bottom:12px;"><span class="inline-icon">${getIcon('inbox', 64)}</span></p>
            <h3 style="color:#004080;">No Forms Yet</h3>
            <p style="font-size:12px;margin-top:10px;">Create your first data collection form</p>
            <button onclick="closeHome()" style="padding:15px 30px;background:#28a745;color:#fff;border:none;border-radius:8px;cursor:pointer;margin-top:20px;font-weight:bold;"><span class="inline-icon">${getIcon('plus', 14)}</span> Create Form</button>
        </div>
    ` : forms.map((f, i) => `
        <div style="background:#fff;border-radius:12px;margin-bottom:15px;box-shadow:0 2px 10px rgba(0,0,0,0.1);overflow:hidden;">
            <div style="background:linear-gradient(135deg,#004080,#002855);color:#fff;padding:12px 15px;">
                <h3 style="margin:0;font-size:14px;"><span class="inline-icon">${getIcon('clipboard-list', 14)}</span> ${escapeHtml(f.title)}</h3>
            </div>
            <div style="padding:12px 15px;">
                <div style="font-size:11px;color:#666;margin-bottom:10px;">
                    <span class="inline-icon">${getIcon('bar-chart-3', 12)}</span> ${f.fieldCount} fields | <span class="inline-icon">${getIcon('edit-3', 12)}</span> ${(f.collectedData || []).length} records
                    ${f.dhis2?.datasetId ? '<span style="color:#17a2b8;margin-left:8px;"><span class="inline-icon">' + getIcon('link', 12) + '</span> DHIS2</span>' : ''}
                </div>
                <div style="display:flex;gap:6px;">
                    <button onclick="editForm(${i})" style="flex:1;padding:8px;background:#004080;color:#fff;border:none;border-radius:4px;cursor:pointer;font-weight:bold;font-size:11px;"><span class="inline-icon">${getIcon('pencil', 12)}</span> Edit</button>
                    <button onclick="deleteForm(${i})" style="padding:8px 12px;background:#dc3545;color:#fff;border:none;border-radius:4px;cursor:pointer;font-weight:bold;font-size:11px;"><span class="inline-icon">${getIcon('trash-2', 12)}</span></button>
                </div>
            </div>
        </div>
    `).join('');
    
    const syncStatus = isLoggedIn ? (synced ? 
        '<span style="color:#28a745;"><span class="inline-icon">' + getIcon('check-circle', 12) + '</span> Synced</span>' : 
        '<span style="color:#ffc107;"><span class="inline-icon">' + getIcon('loader', 12) + '</span> Syncing...</span>') : '';
    
    document.getElementById('homeContainer').innerHTML = `
        <div style="max-width:600px;margin:0 auto;padding:20px;">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;">
                <h2 style="color:#004080;"><span class="inline-icon">${getIcon('home', 20)}</span> My Forms</h2>
                <div style="display:flex;gap:8px;align-items:center;">
                    ${syncStatus}
                    ${isLoggedIn ? `
                        <button onclick="syncFormsManually()" style="padding:10px 15px;background:#17a2b8;color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:bold;" title="Sync with Cloud">
                            <span class="inline-icon">${getIcon('refresh-cw', 14)}</span>
                        </button>
                    ` : ''}
                    <button onclick="closeHome()" style="padding:10px 20px;background:#28a745;color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:bold;"><span class="inline-icon">${getIcon('plus', 14)}</span> New</button>
                </div>
            </div>
            
            ${isLoggedIn ? `
                <div style="background:#d4edda;border:1px solid #c3e6cb;border-radius:8px;padding:10px 15px;margin-bottom:15px;display:flex;align-items:center;gap:10px;">
                    <span class="inline-icon" style="color:#155724;">${getIcon('cloud', 16)}</span>
                    <span style="font-size:12px;color:#155724;">Signed in as <strong>${escapeHtml(state.user.email)}</strong> - Forms sync across devices</span>
                </div>
            ` : `
                <div style="background:#fff3cd;border:1px solid #ffeeba;border-radius:8px;padding:10px 15px;margin-bottom:15px;display:flex;align-items:center;gap:10px;">
                    <span class="inline-icon" style="color:#856404;">${getIcon('alert-triangle', 16)}</span>
                    <span style="font-size:12px;color:#856404;">Forms are stored locally. Sign in to sync across devices.</span>
                </div>
            `}
            
            ${formsHtml}
        </div>
    `;
    
    initIcons();
}

// Manual sync button handler
async function syncFormsManually() {
    notify('Syncing forms...', 'info');
    try {
        await syncForms();
        notify('Forms synced!', 'success');
        renderHomeContent(true);
    } catch (err) {
        notify('Sync failed: ' + err.message, 'error');
    }
}

function closeHome() {
    document.getElementById('homeContainer').classList.remove('show');
    document.querySelector('.header').style.display = 'flex';
    document.getElementById('mainContainer').classList.add('show');
    document.querySelector('.footer').style.display = 'block';
}

function editForm(index) {
    const forms = JSON.parse(safeStorage.getItem('icfCollectForms') || '[]');
    const form = forms[index];
    if (!form) return;
    
    state.fields = form.fields || [];
    state.settings = { ...state.settings, ...form.settings };
    state.fieldCounter = Math.max(...state.fields.map(f => parseInt(f.id.replace('field_', '')) || 0), 0) + 1;
    state.collectedData = form.collectedData || [];
    
    if (form.dhis2) {
        state.dhis2.datasetId = form.dhis2.datasetId;
        state.dhis2.dataElements = form.dhis2.dataElements || {};
    }
    
    document.getElementById('formTitle').value = state.settings.title;
    document.getElementById('previewTitle').textContent = state.settings.title;
    saveToStorage();
    closeHome();
    renderFields();
    renderProperties();
    notify('Form loaded!');
}

function deleteForm(index) {
    const forms = JSON.parse(safeStorage.getItem('icfCollectForms') || '[]');
    const form = forms[index];
    if (!form || !confirm(`Delete "${form.title}"?\n\nThis will also delete the data sheet in Google Sheets.`)) return;
    
    // Delete from cloud (form definition AND data sheet)
    deleteFormFromCloud(form.id, form.title);
    
    forms.splice(index, 1);
    safeStorage.setItem('icfCollectForms', JSON.stringify(forms));
    showHome();
    notify('Deleted');
}

// ==================== FORM VIEWER ====================
async function renderSharedForm(data) {
    document.getElementById('authContainer').style.display = 'none';
    document.getElementById('mainContainer').classList.remove('show');
    document.querySelector('.header').style.display = 'none';
    document.querySelector('.footer').style.display = 'none';
    
    // Show loading indicator
    document.getElementById('viewerContainer').classList.add('show');
    document.getElementById('viewerContainer').innerHTML = `
        <div style="display:flex;align-items:center;justify-content:center;height:100vh;flex-direction:column;">
            <div style="font-size:48px;margin-bottom:20px;">⏳</div>
            <div style="color:#004080;font-weight:600;">Loading form...</div>
        </div>
    `;
    
    state.settings = { title: data.s?.t || 'Form', logo: data.s?.l || state.settings.logo, aggregateColumn: data.s?.ac || '' };
    
    // First pass: create fields with placeholder for cascade data
    state.fields = (data.f || []).map(f => ({
        id: f.i, type: f.y, label: f.l, name: f.n, required: f.r || false, options: f.o || [], max: f.mx || 5,
        showLogic: f.sl || [],
        validation: f.v || [],
        minLength: f.mnl, maxLength: f.mxl, exactLength: f.el,
        exactDigits: f.ed, minDigits: f.mnd, maxDigits: f.mxd,
        noFutureDates: f.nfd, noPastDates: f.npd, minDate: f.mid, maxDate: f.mxd2,
        autoTime: f.at !== false,
        checkDuplicate: f.cd || false,
        cascadeGroup: f.cg || null,
        cascadeLevel: typeof f.cl === 'number' ? f.cl : null,
        cascadeColumn: f.ccol || null,
        cascadeColumns: f.cc || [],
        cascadeData: [], // Will be loaded
        cascadeDataRef: f.cref || null, // Reference ID for cloud data
        cascadeValueColumn: f.cvc || null,
        formula: f.frm || ''
    }));
    
    // Load cascade data from cloud for fields with reference IDs
    const cascadeFields = state.fields.filter(f => f.cascadeDataRef && f.cascadeColumns && f.cascadeColumns.length > 0);
    if (cascadeFields.length > 0) {
        console.log(`Loading cascade data for ${cascadeFields.length} fields...`);
        for (const field of cascadeFields) {
            try {
                const cloudData = await loadCascadeDataFromCloud(field.cascadeDataRef, field.cascadeColumns);
                if (cloudData && cloudData.data) {
                    field.cascadeData = cloudData.data;
                    console.log(`Loaded cascade data for ${field.name}: ${field.cascadeData.length} rows`);
                }
            } catch (err) {
                console.error('Failed to load cascade data for', field.name, err);
            }
        }
    }
    
    // Load DHIS2 dataset/program info
    if (data.d) {
        state.dhis2.datasetId = data.d.id || '';
        state.dhis2.programId = data.d.pid || '';
        state.dhis2.dataElements = data.d.de || {};
    }
    
    // Load DHIS2 credentials from shared link
    if (data.h) {
        state.dhis2.url = data.h.u || '';
        state.dhis2.username = data.h.n || '';
        state.dhis2.password = data.h.p ? atob(data.h.p) : ''; // Decode password
        state.dhis2.syncMode = data.h.m || 'aggregate';
        state.dhis2.orgUnitLevel = data.h.ol || 5;
        state.dhis2.periodType = data.h.pt || 'Monthly';
        state.dhis2.periodColumn = data.h.pc || '';
        state.dhis2.orgUnitColumn = data.h.oc || '';
        state.dhis2.connected = true; // Mark as configured
        
        // Fetch org units in background
        if (state.dhis2.url) {
            fetchOrgUnits().catch(err => console.log('Could not fetch org units:', err));
        }
    }
    
    renderFormViewer();
}

function renderFormViewer() {
    const s = state.settings;
    let pages = [], currentPage = { title: 'Page 1', fields: [], sectionId: null };
    
    console.log('renderFormViewer - Total fields:', state.fields.length);
    
    state.fields.forEach((field, idx) => {
        console.log(`  Field ${idx}: type=${field.type}, label=${field.label}`);
        if (field.type === 'section') {
            if (currentPage.fields.length > 0) pages.push(currentPage);
            currentPage = { title: field.label, fields: [], sectionId: field.id };
        } else {
            currentPage.fields.push(field);
        }
    });
    if (currentPage.fields.length > 0) pages.push(currentPage);
    if (pages.length === 0) pages = [{ title: s.title, fields: state.fields.filter(f => f.type !== 'section'), sectionId: null }];
    
    console.log('renderFormViewer - Pages created:', pages.length);
    pages.forEach((p, i) => console.log(`  Page ${i}: "${p.title}" with ${p.fields.length} fields`));
    
    let pagesHtml = pages.map((page, pageIndex) => {
        console.log(`Rendering page ${pageIndex}: "${page.title}" with ${page.fields.length} fields`);
        
        let fieldsHtml = page.fields.map((field, fieldIdx) => {
            try {
                console.log(`  Field ${fieldIdx}: ${field.name} (${field.type}), options:`, field.options?.length || 0);
            const req = field.required ? '<span style="color:#dc3545;">*</span>' : '';
            const reqAttr = field.required ? 'required' : '';
            let input = '';
            
            // Build validation attributes
            let textAttrs = '';
            if (field.minLength) textAttrs += ` minlength="${field.minLength}"`;
            if (field.maxLength) textAttrs += ` maxlength="${field.maxLength}"`;
            if (field.exactLength) textAttrs += ` minlength="${field.exactLength}" maxlength="${field.exactLength}"`;
            
            let phoneAttrs = '';
            if (field.exactDigits) {
                phoneAttrs = ` pattern="[0-9]{${field.exactDigits}}" title="Must be exactly ${field.exactDigits} digits"`;
            } else if (field.minDigits || field.maxDigits) {
                const min = field.minDigits || 1;
                const max = field.maxDigits || 20;
                phoneAttrs = ` pattern="[0-9]{${min},${max}}" title="Must be ${min}-${max} digits"`;
            }
            
            let dateAttrs = '';
            const today = new Date().toISOString().split('T')[0];
            if (field.noFutureDates) dateAttrs += ` max="${today}"`;
            if (field.noPastDates) dateAttrs += ` min="${today}"`;
            if (field.minDate) dateAttrs += ` min="${field.minDate}"`;
            if (field.maxDate) dateAttrs += ` max="${field.maxDate}"`;
            
            const currentTime = new Date().toTimeString().slice(0, 5);
            const autoTimeValue = (field.autoTime !== false) ? ` value="${currentTime}"` : '';
            
            switch (field.type) {
                case 'period':
                    input = `<select name="${field.name}" class="viewer-input" ${reqAttr}>
                        <option value="">-- Select Period --</option>
                        ${PERIODS.map(p => `<option value="${p}">${p.substring(0,4)}-${p.substring(4)}</option>`).join('')}
                    </select>`;
                    break;
                case 'text':
                    input = `<input type="text" name="${field.name}" class="viewer-input" ${reqAttr}${textAttrs}>`;
                    break;
                case 'email':
                    input = `<input type="email" name="${field.name}" class="viewer-input" ${reqAttr} pattern="[^\\s@]+@[^\\s@]+\\.[^\\s@]+" title="Please enter a valid email (e.g., name@example.com)">`;
                    break;
                case 'phone':
                    input = `<input type="tel" name="${field.name}" class="viewer-input" ${reqAttr}${phoneAttrs}>`;
                    break;
                case 'date':
                    input = `<input type="date" name="${field.name}" class="viewer-input" ${reqAttr}${dateAttrs}>`;
                    break;
                case 'time':
                    input = `<input type="time" name="${field.name}" class="viewer-input" ${reqAttr}${autoTimeValue} readonly style="background:#f0f0f0;">`;
                    break;
                case 'number':
                    input = `<input type="number" name="${field.name}" class="viewer-input" ${reqAttr}>`;
                    break;
                case 'calculation':
                    input = `<input type="text" name="${field.name}" class="viewer-input calculation-field" 
                        data-formula="${escapeHtml(field.formula || '')}" 
                        readonly 
                        style="background:#e9ecef;font-weight:600;color:#004080;">`;
                    break;
                case 'textarea':
                    input = `<textarea name="${field.name}" class="viewer-input" rows="3" ${reqAttr}${textAttrs}></textarea>`;
                    break;
                case 'select':
                    if (field.cascadeGroup) {
                        // Cascade-linked select field
                        const isFirst = (field.cascadeLevel || 0) === 0;
                        const cascadeLevel = field.cascadeLevel || 0;
                        input = `<select name="${field.name}" class="viewer-input cascade-field" 
                            data-cascade-group="${field.cascadeGroup}" 
                            data-cascade-level="${cascadeLevel}"
                            data-cascade-column="${field.cascadeColumn || ''}"
                            onchange="handleCascadeChange(this)"
                            ${isFirst ? '' : 'disabled'} ${reqAttr}>
                            <option value="">-- Select ${escapeHtml(field.label)} --</option>
                        </select>`;
                    } else {
                        input = `<select name="${field.name}" class="viewer-input" ${reqAttr}><option value="">-- Select --</option>${(field.options||[]).map(o => `<option value="${escapeHtml(o)}">${escapeHtml(o)}</option>`).join('')}</select>`;
                    }
                    break;
                case 'radio': case 'yesno':
                    const opts = field.type === 'yesno' ? ['Yes','No'] : (field.options||[]);
                    input = `<div class="viewer-radio-group">${opts.map(o => `<label class="viewer-radio-option"><input type="radio" name="${field.name}" value="${escapeHtml(o)}" ${reqAttr}><span>${escapeHtml(o)}</span></label>`).join('')}</div>`;
                    break;
                case 'checkbox':
                    // Don't add HTML required - custom validation checks if ANY checkbox is checked
                    input = `<div class="viewer-radio-group">${(field.options||[]).map((o, idx) => `<label class="viewer-radio-option"><input type="checkbox" name="${field.name}" value="${escapeHtml(o)}"><span>${escapeHtml(o)}</span></label>`).join('')}</div>`;
                    break;
                case 'gps':
                    input = `<div><button type="button" onclick="captureGPS(this)" style="padding:10px 20px;background:#004080;color:white;border:none;border-radius:6px;cursor:pointer;"><span data-icon="map-pin" data-size="14"></span> Get Location</button><input type="hidden" name="${field.name}"><div class="gps-status" style="margin-top:6px;font-size:11px;"></div>
        `;
                    break;
                case 'qrcode':
                    input = `<div class="qr-scanner-container">
                        <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
                            <button type="button" onclick="startQRScanner(this)" style="padding:10px 20px;background:#6f42c1;color:white;border:none;border-radius:6px;cursor:pointer;"><span data-icon="scan" data-size="14"></span> Scan QR/Barcode</button>
                            <input type="hidden" name="${field.name}" ${reqAttr}>
                            <span class="qr-value" style="font-size:12px;color:#28a745;font-weight:600;"></span>
                        </div>
                        <div class="qr-preview" style="margin-top:10px;display:none;">
                            <video class="qr-video" style="width:100%;max-width:300px;border-radius:8px;"></video>
                            <button type="button" onclick="stopQRScanner(this)" style="margin-top:8px;padding:6px 12px;background:#dc3545;color:white;border:none;border-radius:4px;cursor:pointer;">Stop Scanner</button>
                        </div>
                        <div class="qr-status" style="margin-top:6px;font-size:11px;"></div>
                    </div>`;
                    break;
                case 'cascade':
                    const cascadeColumns = field.cascadeColumns || [];
                    const cascadeId = field.id;
                    if (cascadeColumns.length > 0 && field.cascadeData && field.cascadeData.length > 0) {
                        input = `<div class="cascade-container" data-cascade-id="${cascadeId}" data-field-name="${field.name}" data-value-column="${field.cascadeValueColumn || cascadeColumns[cascadeColumns.length - 1]}">
                            ${cascadeColumns.map((col, idx) => `
                                <div class="cascade-level" data-level="${idx}">
                                    <label class="cascade-label">${escapeHtml(col)}</label>
                                    <select class="viewer-input cascade-select" data-column="${col}" data-level="${idx}" onchange="handleCascadeSelect(this)" ${idx === 0 ? '' : 'disabled'}>
                                        <option value="">-- Select ${escapeHtml(col)} --</option>
                                    </select>
                                </div>
                            `).join('')}
                            <input type="hidden" name="${field.name}" ${reqAttr}>
                        </div>`;
                    } else {
                        input = `<div style="background:#fff3cd;padding:10px;border-radius:6px;font-size:11px;color:#856404;">
                            <span data-icon="alert-triangle" data-size="12"></span> Cascade data not configured. Upload Excel in form builder.
                        </div>`;
                    }
                    break;
                case 'rating':
                    input = `<div>${Array(field.max || 5).fill().map((_, i) => `<span onclick="setRating(this,${i+1})" class="rating-star" style="font-size:24px;cursor:pointer;opacity:0.3;color:#ffc107;">★</span>`).join('')}<input type="hidden" name="${field.name}"></div>`;
                    break;
                default:
                    input = `<input type="text" name="${field.name}" class="viewer-input" ${reqAttr}>`;
            }
            
            return `<div class="viewer-field" data-field-name="${field.name}" data-field-id="${field.id}"><label class="viewer-field-label">${escapeHtml(field.label)} ${req}</label>${input}<div class="field-error-msg"></div></div>`;
            } catch (err) {
                console.error('Error rendering field:', field, err);
                return `<div class="viewer-field" style="color:red;">Error rendering field: ${field.label || field.name}</div>`;
            }
        }).join('');
        
        const isFirst = pageIndex === 0, isLast = pageIndex === pages.length - 1;
        
        return `
            <div class="form-page" data-page="${pageIndex}" data-section-id="${page.sectionId || ''}" style="${pageIndex === 0 ? '' : 'display:none;'}">
                ${pages.length > 1 ? `<div class="page-header"><span class="page-indicator">Page ${pageIndex + 1}/${pages.length}</span><h3 class="page-title">${escapeHtml(page.title)}</h3></div>` : ''}
                ${fieldsHtml}
                <div class="page-navigation">
                    ${!isFirst ? `<button type="button" class="nav-btn back-btn" onclick="goToPrevPage(${pageIndex})"><span class="inline-icon">${getIcon('arrow-left', 14)}</span> Back</button>` : '<div></div>'}
                    ${!isLast ? `<button type="button" class="nav-btn next-btn" onclick="goToNextPage(${pageIndex})">Next <span class="inline-icon">${getIcon('arrow-right', 14)}</span></button>` : ''}
                    ${isLast ? `<button type="submit" class="nav-btn submit-btn"><span data-icon="check" data-size="14"></span> Submit</button>` : ''}
                </div>
            </div>
        `;
    }).join('');
    
    const backBtn = state.isSharedMode ? '' : '<button class="viewer-back-btn" onclick="closeViewer()"><span class="inline-icon">' + getIcon('arrow-left', 14) + '</span> Back</button>';
    
    document.getElementById('viewerContainer').innerHTML = `
        <div class="viewer-nav">
            ${backBtn}
            <button class="viewer-nav-btn active" style="background:#004080;color:#fff;" onclick="showTab('form',this)"><span data-icon="edit-3" data-size="14"></span> Form</button>
            <button class="viewer-nav-btn" style="background:#28a745;color:#fff;" onclick="showTab('data',this)"><span data-icon="table" data-size="14"></span> Data</button>
            <button class="viewer-nav-btn" style="background:#17a2b8;color:#fff;" onclick="showTab('dashboard',this)"><span data-icon="bar-chart-2" data-size="14"></span> Dashboard</button>
            ${state.dhis2.url ? `<div class="connection-status" style="background:#17a2b8;"><span class="inline-icon">${getIcon('link', 14)}</span> DHIS2</div>` : ''}
            <div id="connectionStatus" class="connection-status ${navigator.onLine ? 'online' : 'offline'}">
                ${navigator.onLine ? '<span class="inline-icon" style="color:#28a745;">' + getIcon('wifi', 14) + '</span> Online' : '<span class="inline-icon" style="color:#dc3545;">' + getIcon('wifi-off', 14) + '</span> Offline'}
            </div>
        </div>
        
        <div id="tabForm" class="viewer-tab active">
            <div class="viewer-form">
                <div class="viewer-form-box">
                    <div class="viewer-header">
                        <img src="${s.logo}" alt="Logo">
                        <h1>${escapeHtml(s.title)}</h1>
                        <p>ICF-SL Data Collection System</p>
                    </div>
                    <div class="viewer-body">
                        <form id="viewerForm">${pagesHtml}</form>
                        
                        <!-- Draft Actions -->
                        <div class="draft-actions-bar">
                            <button type="button" class="draft-save-btn" onclick="saveDraft()">
                                <span data-icon="file-edit" data-size="14"></span> Save as Draft
                            </button>
                            <button type="button" class="draft-toggle-btn" onclick="toggleDraftsPanel()">
                                <span data-icon="folder" data-size="14"></span> <span id="draftsCount">Drafts (0)</span>
                            </button>
                        </div>
                        
                        <!-- Drafts Panel -->
                        <div id="draftsPanelWrapper" class="drafts-panel-wrapper" style="display:none;">
                            <div class="drafts-panel-header">
                                <span><span data-icon="file-edit" data-size="14"></span> Saved Drafts</span>
                                <button onclick="toggleDraftsPanel()" style="background:none;border:none;cursor:pointer;color:#666;">✕</button>
                            </div>
                            <div id="draftsPanel"></div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <div id="tabData" class="viewer-tab">
            <div style="max-width:1100px;margin:0 auto;" id="dataContent">
                <p style="text-align:center;padding:40px;"><span data-icon="loader" data-size="14"></span> Loading...</p>
            </div>
        </div>
        
        <div id="tabDashboard" class="viewer-tab">
            <div class="dashboard-container" id="dashboardContent">
                <p style="text-align:center;padding:40px;"><span data-icon="loader" data-size="14"></span> Loading...</p>
            </div>
        </div>
    `;
    
    window.totalPages = pages.length;
    setupFormSubmit();
    loadViewerData();
    initIcons();
    
    // Initialize drafts panel
    setTimeout(() => {
        renderDraftsPanel();
        updateDraftsCount();
    }, 100);
    
    // Apply initial logic and initialize cascades after form renders
    setTimeout(() => {
        applyFormLogic();
        initCascades();
    }, 100);
}

function setupFormSubmit() {
    const form = document.getElementById('viewerForm');
    if (!form) return;
    
    // Add change listener to apply logic and run calculations when form values change
    form.addEventListener('input', () => { applyFormLogic(); runCalculations(); });
    form.addEventListener('change', () => { applyFormLogic(); runCalculations(); });
    
    form.onsubmit = async function(e) {
        e.preventDefault();
        
        // Validate all visible fields
        const validationResult = validateAllFields();
        if (!validationResult.valid) {
            notify(validationResult.message, 'error');
            return;
        }
        
        const data = { _id: Date.now(), _timestamp: new Date().toISOString(), _synced: false };
        new FormData(form).forEach((v, k) => data[k] = v);
        
        // Show submitting state
        const submitBtn = form.querySelector('.submit-btn');
        const originalText = submitBtn ? submitBtn.innerHTML : '';
        if (submitBtn) {
            submitBtn.innerHTML = '<span class="spin" style="display:inline-block;">' + getIcon('loader', 14) + '</span> Checking...';
            submitBtn.disabled = true;
        }
        
        // Check for duplicates
        const duplicateFields = state.fields.filter(f => f.checkDuplicate && data[f.name]);
        if (duplicateFields.length > 0) {
            const duplicateResult = await checkDuplicates(data, duplicateFields);
            if (duplicateResult.hasDuplicate) {
                if (submitBtn) {
                    submitBtn.innerHTML = originalText;
                    submitBtn.disabled = false;
                }
                notify(`Duplicate found: ${duplicateResult.field} = "${duplicateResult.value}" already exists`, 'error');
                return;
            }
        }
        
        if (submitBtn) {
            submitBtn.innerHTML = '<span class="spin" style="display:inline-block;">' + getIcon('loader', 14) + '</span> Submitting...';
        }
        
        // Submit to Google Sheets first (if online)
        const result = await submitToGoogleSheets(data);
        
        // Update sync status based on result
        if (result.success && !result.offline) {
            data._synced = true;
        }
        
        // Save locally
        state.collectedData.push(data);
        saveToStorage();
        saveFormToList();
        
        // Remove from drafts if it was a draft being submitted
        const draftId = form.dataset.draftId ? parseInt(form.dataset.draftId) : null;
        if (draftId) {
            removeDraft(draftId);
            delete form.dataset.draftId;
            updateDraftsCount();
        }
        
        // Reset form
        form.reset();
        goToPage(0);
        
        // Re-apply logic after reset
        setTimeout(() => applyFormLogic(), 100);
        
        // Restore button
        if (submitBtn) {
            submitBtn.innerHTML = originalText;
            submitBtn.disabled = false;
        }
        
        // Show appropriate notification
        if (result.success && !result.offline) {
            notify('Submitted to Google Sheets!', 'success');
        } else if (result.offline) {
            notify('Saved offline (will sync when online)', 'info');
        } else {
            notify('Saved locally', 'info');
        }
        
        loadViewerData();
    };
}

// Check for duplicate entries in Google Sheets
async function checkDuplicates(data, duplicateFields) {
    try {
        // First check local data
        for (const field of duplicateFields) {
            const value = data[field.name];
            const localDuplicate = state.collectedData.find(d => d[field.name] === value);
            if (localDuplicate) {
                return { hasDuplicate: true, field: field.label, value: value };
            }
        }
        
        // Then check Google Sheets
        const existingData = await loadFromGoogleSheets();
        for (const field of duplicateFields) {
            const value = data[field.name];
            const sheetDuplicate = existingData.find(d => d[field.name] === value);
            if (sheetDuplicate) {
                return { hasDuplicate: true, field: field.label, value: value };
            }
        }
        
        return { hasDuplicate: false };
    } catch (err) {
        console.error('Duplicate check error:', err);
        return { hasDuplicate: false }; // Allow submission if check fails
    }
}

// ==================== DRAFT FUNCTIONALITY ====================
function getDraftStorageKey() {
    return `icf_drafts_${state.settings.title.replace(/\s+/g, '_')}`;
}

function loadDrafts() {
    try {
        return JSON.parse(safeStorage.getItem(getDraftStorageKey()) || '[]');
    } catch (err) {
        return [];
    }
}

function saveDrafts(drafts) {
    safeStorage.setItem(getDraftStorageKey(), JSON.stringify(drafts));
}

window.saveDraft = function() {
    const form = document.getElementById('viewerForm');
    if (!form) return;
    
    // Check if any field has data first
    const formData = new FormData(form);
    const tempData = {};
    formData.forEach((v, k) => tempData[k] = v);
    
    const hasData = state.fields.some(f => f.type !== 'section' && tempData[f.name]);
    if (!hasData) {
        notify('No data to save as draft', 'info');
        return;
    }
    
    // Prompt for draft name
    const draftName = prompt('Enter a name for this draft:', 'Draft ' + (loadDrafts().length + 1));
    if (!draftName) return; // User cancelled
    
    const data = { 
        _draftId: Date.now(), 
        _draftName: draftName.trim(),
        _savedAt: new Date().toISOString()
    };
    formData.forEach((v, k) => data[k] = v);
    
    const drafts = loadDrafts();
    drafts.push(data);
    saveDrafts(drafts);
    
    notify(`Draft "${draftName}" saved!`, 'success');
    renderDraftsPanel();
    updateDraftsCount();
};

window.loadDraft = function(draftId) {
    const drafts = loadDrafts();
    const draft = drafts.find(d => d._draftId === draftId);
    if (!draft) return;
    
    const form = document.getElementById('viewerForm');
    if (!form) return;
    
    // Populate form with draft data
    state.fields.forEach(field => {
        if (field.type === 'section') return;
        const value = draft[field.name];
        if (value === undefined) return;
        
        const input = form.querySelector(`[name="${field.name}"]`);
        if (!input) return;
        
        if (input.type === 'radio') {
            // Find the radio with matching value
            const radio = form.querySelector(`input[name="${field.name}"][value="${value}"]`);
            if (radio) radio.checked = true;
        } else if (input.type === 'checkbox') {
            const values = Array.isArray(value) ? value : [value];
            form.querySelectorAll(`input[name="${field.name}"]`).forEach(cb => {
                cb.checked = values.includes(cb.value);
            });
        } else {
            input.value = value;
        }
    });
    
    // Store draft ID for removal after submit
    form.dataset.draftId = draftId;
    
    notify('Draft loaded!', 'info');
    goToPage(0);
    setTimeout(() => applyFormLogic(), 100);
};

window.deleteDraft = function(draftId) {
    if (!confirm('Delete this draft?')) return;
    removeDraft(draftId);
    notify('Draft deleted');
    renderDraftsPanel();
    updateDraftsCount();
};

function removeDraft(draftId) {
    if (!draftId) return;
    const drafts = loadDrafts();
    const filtered = drafts.filter(d => d._draftId !== draftId);
    saveDrafts(filtered);
}

function renderDraftsPanel() {
    const container = document.getElementById('draftsPanel');
    if (!container) return;
    
    const drafts = loadDrafts();
    
    if (drafts.length === 0) {
        container.innerHTML = '<p style="text-align:center;color:#868e96;font-size:11px;padding:10px;">No saved drafts</p>';
        return;
    }
    
    let html = '<div class="drafts-list">';
    drafts.forEach(draft => {
        const savedDate = new Date(draft._savedAt).toLocaleString();
        // Use draft name if available, otherwise show preview
        const draftName = draft._draftName || 'Unnamed Draft';
        // Get first few non-empty field values for preview subtitle
        const preview = state.fields
            .filter(f => f.type !== 'section' && draft[f.name])
            .slice(0, 2)
            .map(f => String(draft[f.name]).substring(0, 15))
            .join(', ') || '';
        
        html += `
            <div class="draft-item">
                <div class="draft-info">
                    <div class="draft-preview" style="font-weight:700;">${escapeHtml(draftName)}</div>
                    <div class="draft-date">${savedDate}${preview ? ' • ' + escapeHtml(preview) : ''}</div>
                </div>
                <div class="draft-actions">
                    <button class="draft-btn load" onclick="loadDraft(${draft._draftId})" title="Load Draft">
                        <span data-icon="file-edit" data-size="12"></span>
                    </button>
                    <button class="draft-btn delete" onclick="deleteDraft(${draft._draftId})" title="Delete Draft">
                        <span data-icon="trash-2" data-size="12"></span>
                    </button>
                </div>
            </div>
        `;
    });
    html += '</div>';
    
    container.innerHTML = html;
    initIcons();
}

window.toggleDraftsPanel = function() {
    const wrapper = document.getElementById('draftsPanelWrapper');
    if (wrapper) {
        wrapper.style.display = wrapper.style.display === 'none' ? 'block' : 'none';
        if (wrapper.style.display === 'block') {
            renderDraftsPanel();
        }
    }
};

function updateDraftsCount() {
    const countEl = document.getElementById('draftsCount');
    if (countEl) {
        const drafts = loadDrafts();
        countEl.textContent = `Drafts (${drafts.length})`;
    }
}

async function loadViewerData() {
    console.log('=== loadViewerData called ===');
    console.log('Form title:', state.settings.title);
    console.log('Fields count:', state.fields.length);
    console.log('Existing collected data:', state.collectedData.length);
    
    // Show loading state
    const dataContainer = document.getElementById('dataContent');
    const dashContainer = document.getElementById('dashboardContent');
    
    if (dataContainer) {
        dataContainer.innerHTML = `<p style="text-align:center;padding:40px;"><span class="spin" style="display:inline-block;">${getIcon('loader', 20)}</span> Loading data from Google Sheets...</p>`;
    }
    if (dashContainer) {
        dashContainer.innerHTML = `<p style="text-align:center;padding:40px;"><span class="spin" style="display:inline-block;">${getIcon('loader', 20)}</span> Loading dashboard...</p>`;
    }
    
    // Try to load from Google Sheets
    try {
        console.log('Calling loadFromGoogleSheets...');
        const sheetsData = await loadFromGoogleSheets();
        console.log('Sheets data received:', sheetsData?.length || 0, 'records');
        
        if (sheetsData && sheetsData.length > 0) {
            state.collectedData = sheetsData.map(d => {
                const sanitized = { ...d };
                sanitized._id = d._id ? (typeof d._id === 'string' ? parseInt(d._id) || d._id : d._id) : Date.now();
                sanitized._synced = true;
                return sanitized;
            });
            console.log('✓ Data loaded from Google Sheets:', state.collectedData.length, 'records');
        } else {
            console.log('No data in Sheets, using local data:', state.collectedData.length, 'records');
        }
    } catch (err) {
        console.error('Could not load from Sheets:', err.message);
        console.log('Fallback to local data:', state.collectedData.length, 'records');
    }
    
    console.log('Rendering data content...');
    renderDataContent();
    console.log('Rendering dashboard...');
    renderDashboard();
    updateConnectionStatus();
    console.log('=== loadViewerData complete ===');
}

function renderDataContent() {
    console.log('=== renderDataContent called ===');
    const container = document.getElementById('dataContent');
    if (!container) {
        console.log('Data content container not found!');
        return;
    }
    
    try {
    console.log('Collected data count:', state.collectedData.length);
    
    const orderedFilterFields = getOrderedFilterFields();
    const filteredData = getFilteredData();
    const aggregateData = calculateAggregateData();
    console.log('Filtered data:', filteredData.length, 'Aggregate:', aggregateData.length);
    
    const activeFilterCount = Object.keys(state.filters).filter(k => state.filters[k]).length + 
                             (state.dateFilter.start || state.dateFilter.end ? 1 : 0);
    
    // Build filter panel
    let filtersHtml = `
        <div class="filter-group">
            <label class="filter-label"><span class="inline-icon">${getIcon('calendar', 12)}</span> From</label>
            <input type="date" class="filter-input" value="${state.dateFilter.start}" onchange="updateDateFilter('start', this.value)">
        </div>
        <div class="filter-group">
            <label class="filter-label"><span class="inline-icon">${getIcon('calendar', 12)}</span> To</label>
            <input type="date" class="filter-input" value="${state.dateFilter.end}" onchange="updateDateFilter('end', this.value)">
        </div>
    `;
    
    orderedFilterFields.forEach((field, idx) => {
        const uniqueValues = [...new Set(state.collectedData.map(d => d[field.name]).filter(Boolean))];
        filtersHtml += `
            <div class="filter-group with-arrows">
                <button class="filter-arrow-btn left" onclick="moveFilter('${field.name}','up')" title="Move Left">◀</button>
                <label class="filter-label">${escapeHtml(field.label)}</label>
                <select class="filter-select" onchange="updateFilter('${field.name}', this.value)">
                    <option value="">All</option>
                    ${uniqueValues.map(v => `<option value="${escapeHtml(v)}" ${state.filters[field.name] === v ? 'selected' : ''}>${escapeHtml(v)}</option>`).join('')}
                </select>
                <button class="filter-arrow-btn right" onclick="moveFilter('${field.name}','down')" title="Move Right">▶</button>
            </div>
        `;
    });
    
    // Build aggregate column options (categorical fields that can be used for grouping)
    const aggregatableFields = state.fields.filter(f => 
        ['select', 'radio', 'yesno', 'text'].includes(f.type) && f.type !== 'section'
    );
    const selectedColumns = state.settings.aggregateColumns || [];
    const aggregateColumnHtml = `
        <div class="config-section" style="margin-bottom:15px;padding:12px;">
            <div style="display:flex;align-items:flex-start;gap:15px;flex-wrap:wrap;">
                <div style="display:flex;flex-direction:column;gap:8px;">
                    <div style="display:flex;align-items:center;gap:8px;">
                        <span class="inline-icon">${getIcon('layers', 14)}</span>
                        <strong style="font-size:11px;">Aggregate By (select one or more):</strong>
                    </div>
                    <div style="display:flex;flex-wrap:wrap;gap:8px;max-width:500px;">
                        ${aggregatableFields.map(f => `
                            <label style="display:flex;align-items:center;gap:4px;padding:4px 8px;background:${selectedColumns.includes(f.name) ? '#004080' : '#f1f3f5'};color:${selectedColumns.includes(f.name) ? '#fff' : '#333'};border-radius:4px;cursor:pointer;font-size:11px;">
                                <input type="checkbox" value="${f.name}" ${selectedColumns.includes(f.name) ? 'checked' : ''} onchange="toggleAggregateColumn('${f.name}')" style="margin:0;">
                                ${escapeHtml(f.label)}
                            </label>
                        `).join('')}
                    </div>
                </div>
                <div style="font-size:10px;color:#28a745;">
                    ${selectedColumns.length > 0 ? `<span class="inline-icon">${getIcon('check-circle', 12)}</span> Grouping by: ${selectedColumns.map(c => {
                        const field = aggregatableFields.find(f => f.name === c);
                        return field ? field.label : c;
                    }).join(' + ')}` : '<span style="color:#868e96;">Select columns to group data</span>'}
                </div>
            </div>
        </div>
    `;
    
    container.innerHTML = `
        <div class="filter-panel">
            <div class="filter-header">
                <div class="filter-title"><span class="inline-icon">${getIcon('filter', 14)}</span> Filters ${activeFilterCount > 0 ? `<span class="filter-count">${activeFilterCount} active</span>` : ''}</div>
                <button class="filter-btn clear" onclick="clearAllFilters()"><span class="inline-icon">${getIcon('trash-2', 12)}</span> Clear</button>
            </div>
            <div class="filter-controls">${filtersHtml}</div>
        </div>
        
        ${aggregateColumnHtml}
        
        <div class="data-view-tabs">
            <div class="data-view-tab ${state.currentDataView === 'case' ? 'active' : ''}" onclick="switchDataView('case')"><span class="inline-icon">${getIcon('list', 14)}</span> Case-Based (${filteredData.length})</div>
            <div class="data-view-tab ${state.currentDataView === 'aggregate' ? 'active aggregate' : ''}" onclick="switchDataView('aggregate')"><span class="inline-icon">${getIcon('bar-chart-3', 14)}</span> Aggregate (${aggregateData.length})</div>
        </div>
        
        <div id="dataTableContainer">${state.currentDataView === 'case' ? renderCaseTable(filteredData) : renderAggregateTable(aggregateData)}</div>
        
        <div style="margin-top:20px;display:flex;gap:10px;flex-wrap:wrap;">
            <button class="modal-btn primary" onclick="refreshData()"><span class="inline-icon">${getIcon('refresh-cw', 14)}</span> Refresh</button>
            <button class="modal-btn success" onclick="downloadCSV()"><span class="inline-icon">${getIcon('download', 14)}</span> Download CSV</button>
            ${getOfflineCount() > 0 ? `<button class="modal-btn" style="background:#ffc107;color:#000;" onclick="syncOfflineData()"><span class="inline-icon">${getIcon('upload', 14)}</span> Sync Offline (${getOfflineCount()})</button>` : ''}
            ${state.dhis2.url ? `<button class="modal-btn" style="background:#6f42c1;color:#fff;" onclick="syncCaseBased()"><span class="inline-icon">${getIcon('list', 14)}</span> Sync Case-Based</button>` : ''}
            ${state.dhis2.url ? `<button class="modal-btn" style="background:#17a2b8;color:#fff;" onclick="syncAggregate()"><span class="inline-icon">${getIcon('bar-chart-3', 14)}</span> Sync Aggregate</button>` : ''}
        </div>
    `;
    initIcons();
    
    } catch (err) {
        console.error('Error in renderDataContent:', err);
        container.innerHTML = `<div style="text-align:center;padding:40px;color:#dc3545;"><p>Error loading data</p><p style="font-size:12px;">${escapeHtml(err.message)}</p></div>`;
    }
}

window.switchDataView = function(view) {
    state.currentDataView = view;
    renderDataContent();
};

window.setAggregateColumn = function(columnName) {
    // Legacy function - now uses array
    state.settings.aggregateColumns = columnName ? [columnName] : [];
    state.settings.aggregateColumn = columnName; // Keep for backward compatibility
    saveToStorage();
    renderDataContent();
    renderDashboard();
};

window.toggleAggregateColumn = function(columnName) {
    // Initialize array if needed
    if (!state.settings.aggregateColumns) {
        state.settings.aggregateColumns = [];
    }
    
    const index = state.settings.aggregateColumns.indexOf(columnName);
    if (index >= 0) {
        // Remove column
        state.settings.aggregateColumns.splice(index, 1);
    } else {
        // Add column
        state.settings.aggregateColumns.push(columnName);
    }
    
    // Update legacy single column (use first one if any)
    state.settings.aggregateColumn = state.settings.aggregateColumns[0] || '';
    
    saveToStorage();
    renderDataContent();
    renderDashboard();
};

function renderCaseTable(data) {
    if (data.length === 0) {
        return '<p style="text-align:center;padding:40px;color:#868e96;"><span data-icon="inbox" data-size="20"></span> No records found</p>';
    }
    
    const fields = state.fields.filter(f => f.type !== 'section');
    
    let html = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>#</th><th><span data-icon="calendar" data-size="12"></span> Timestamp</th>';
    fields.slice(0, 8).forEach(f => html += `<th>${escapeHtml(f.label)}</th>`);
    html += '<th>Status</th></tr></thead><tbody>';
    
    data.slice().reverse().forEach((row, i) => {
        const syncStatus = row._synced ? 
            '<span class="sync-badge synced"><span data-icon="check" data-size="10"></span> Synced</span>' : 
            (row._syncError ? '<span class="sync-badge failed"><span data-icon="x" data-size="10"></span> Failed</span>' : '<span class="sync-badge pending"><span data-icon="clock" data-size="10"></span> Pending</span>');
            
        html += `<tr><td>${data.length - i}</td><td style="font-size:10px;">${new Date(row._timestamp).toLocaleString()}</td>`;
        fields.slice(0, 8).forEach(f => html += `<td>${escapeHtml(String(row[f.name] || '-').substring(0, 25))}</td>`);
        html += `<td>${syncStatus}</td></tr>`;
    });
    
    html += '</tbody></table></div>';
    return html;
}

function renderAggregateTable(data) {
    if (data.length === 0) {
        return '<p style="text-align:center;padding:40px;color:#868e96;"><span data-icon="inbox" data-size="20"></span> No aggregate data available</p>';
    }
    
    const periodColumn = state.dhis2.periodColumn;
    const aggregateColumn = state.settings.aggregateColumn || state.dhis2.orgUnitColumn;
    
    // Fields to skip (period, aggregate column, and non-aggregatable types)
    const skipFields = [periodColumn, aggregateColumn].filter(Boolean);
    const skipTypes = ['phone', 'gps', 'email', 'text', 'textarea', 'date', 'time'];
    
    // Exclude period, aggregate column, and text-based fields
    const dataFields = state.fields.filter(f => 
        f.type !== 'section' && 
        !skipFields.includes(f.name) &&
        !skipTypes.includes(f.type)
    );
    
    // Build column definitions - expand categorical fields into multiple columns
    const columns = [];
    const categoricalTypes = ['select', 'radio', 'yesno', 'checkbox'];
    
    dataFields.forEach(f => {
        const def = fieldDefs[f.type];
        
        if (categoricalTypes.includes(f.type)) {
            // Split categorical field into option columns (e.g., Sex_Male, Sex_Female)
            const options = f.type === 'yesno' ? ['Yes', 'No'] : (f.options || []);
            options.forEach(opt => {
                columns.push({
                    key: `${f.name}_${opt}`,
                    label: `${f.label} (${opt})`,
                    type: 'categorical'
                });
            });
        } else if (def?.category === 'numeric' || f.type === 'number') {
            columns.push({
                key: f.name,
                label: `${f.label} (SUM)`,
                type: 'numeric'
            });
        }
        // Skip text fields and other non-aggregatable types
    });
    
    // Get the aggregate column label (aggregateColumn already defined above)
    const aggregateField = state.fields.find(f => f.name === aggregateColumn);
    const aggregateLabel = aggregateField?.label || 'Group';
    
    // Build table
    let html = '<div style="overflow-x:auto;"><table class="aggregate-table"><thead><tr>';
    html += '<th style="position:sticky;left:0;background:#004080;z-index:2;"><span class="inline-icon">' + getIcon('layers', 14) + '</span> ' + escapeHtml(aggregateLabel) + '</th>';
    html += '<th><span class="inline-icon">' + getIcon('calendar', 14) + '</span> Period</th>';
    html += '<th><span class="inline-icon">' + getIcon('bar-chart-3', 14) + '</span> N</th>';
    
    columns.forEach(col => {
        const bgColor = col.type === 'categorical' ? '#0066cc' : '#004080';
        html += `<th style="background:${bgColor};white-space:nowrap;font-size:11px;">${escapeHtml(col.label)}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    data.forEach(row => {
        html += `<tr>`;
        html += `<td style="position:sticky;left:0;background:#fff;z-index:1;font-weight:600;">${escapeHtml(row._group)}</td>`;
        html += `<td>${escapeHtml(row._period)}</td>`;
        html += `<td class="aggregate-value" style="font-weight:600;">${row._count}</td>`;
        
        columns.forEach(col => {
            const value = row[col.key];
            const displayVal = value !== undefined && value !== null ? value : 0;
            html += `<td class="aggregate-value">${displayVal}</td>`;
        });
        html += '</tr>';
    });
    
    html += '</tbody></table></div>';
    return html;
}

window.refreshData = async function() {
    notify('Refreshing...');
    await loadViewerData();
    notify('Refreshed!');
};

window.downloadCSV = function() {
    const isAggregate = state.currentDataView === 'aggregate';
    const data = isAggregate ? calculateAggregateData() : getFilteredData();
    if (data.length === 0) { notify('No data', 'error'); return; }
    
    // Build a map of variable names to labels
    const nameToLabel = {};
    state.fields.forEach(f => {
        if (f.name && f.label) {
            nameToLabel[f.name] = f.label;
        }
    });
    // Add system field labels
    nameToLabel['_timestamp'] = 'Timestamp';
    nameToLabel['_id'] = 'Record ID';
    nameToLabel['_group'] = 'Group';
    nameToLabel['_period'] = 'Period';
    nameToLabel['_count'] = 'Count';
    
    // For aggregate data, include _group, _period, _count; for case data, include _timestamp
    const headers = Object.keys(data[0]).filter(k => {
        if (isAggregate) {
            return k === '_group' || k === '_period' || k === '_count' || !k.startsWith('_');
        }
        return !k.startsWith('_') || k === '_timestamp';
    });
    
    // Create header row with labels
    const headerLabels = headers.map(h => nameToLabel[h] || h);
    let csv = headerLabels.map(h => `"${String(h).replace(/"/g, '""')}"`).join(',') + '\n';
    
    data.forEach(row => {
        csv += headers.map(h => `"${String(row[h] || '').replace(/"/g, '""')}"`).join(',') + '\n';
    });
    
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${state.settings.title}_${state.currentDataView}_data.csv`;
    a.click();
    URL.revokeObjectURL(url);
    notify('Downloaded!');
};

function renderDashboard() {
    console.log('=== renderDashboard called ===');
    const container = document.getElementById('dashboardContent');
    if (!container) {
        console.log('Dashboard container not found!');
        return;
    }
    
    console.log('Collected data count:', state.collectedData.length);
    console.log('Fields count:', state.fields.length);
    
    try {
    const orderedFilterFields = getOrderedFilterFields();
    const filteredData = getFilteredData();
    const aggregateData = calculateAggregateData();
    console.log('Filtered data:', filteredData.length, 'Aggregate data:', aggregateData.length);
    
    const activeFilterCount = Object.keys(state.filters).filter(k => state.filters[k]).length + 
                             (state.dateFilter.start || state.dateFilter.end ? 1 : 0);
    
    // Destroy old charts
    Object.values(state.chartInstances).forEach(chart => { if (chart) chart.destroy(); });
    state.chartInstances = {};
    
    // Stats
    const synced = state.collectedData.filter(d => d._synced).length;
    const pending = state.collectedData.filter(d => !d._synced && !d._syncError).length;
    const periodColumn = state.dhis2.periodColumn;
    const uniquePeriods = periodColumn ? [...new Set(filteredData.map(d => d[periodColumn]).filter(Boolean))].length : 0;
    
    // Use aggregateColumn from settings, or fall back to orgUnitColumn for DHIS2
    const aggregateColumn = state.settings.aggregateColumn || state.dhis2.orgUnitColumn;
    const aggregateField = state.fields.find(f => f.name === aggregateColumn);
    const uniqueGroups = aggregateColumn ? [...new Set(filteredData.map(d => d[aggregateColumn]).filter(Boolean))].length : 0;
    
    // Build filter panel HTML (same as Data tab)
    let filtersHtml = `
        <div class="filter-group">
            <label class="filter-label"><span class="inline-icon">${getIcon('calendar', 12)}</span> From</label>
            <input type="date" class="filter-input" value="${state.dateFilter.start}" onchange="updateDateFilter('start', this.value)">
        </div>
        <div class="filter-group">
            <label class="filter-label"><span class="inline-icon">${getIcon('calendar', 12)}</span> To</label>
            <input type="date" class="filter-input" value="${state.dateFilter.end}" onchange="updateDateFilter('end', this.value)">
        </div>
    `;
    
    orderedFilterFields.forEach((field, idx) => {
        const uniqueValues = [...new Set(state.collectedData.map(d => d[field.name]).filter(Boolean))];
        filtersHtml += `
            <div class="filter-group with-arrows">
                <button class="filter-arrow-btn left" onclick="moveFilter('${field.name}','up')" title="Move Left">◀</button>
                <label class="filter-label">${escapeHtml(field.label)}</label>
                <select class="filter-select" onchange="updateFilter('${field.name}', this.value)">
                    <option value="">All</option>
                    ${uniqueValues.map(v => `<option value="${escapeHtml(v)}" ${state.filters[field.name] === v ? 'selected' : ''}>${escapeHtml(v)}</option>`).join('')}
                </select>
                <button class="filter-arrow-btn right" onclick="moveFilter('${field.name}','down')" title="Move Right">▶</button>
            </div>
        `;
    });
    
    // Build charts - EXCLUDE aggregate column from charts
    const categoricalTypes = ['select', 'radio', 'yesno', 'checkbox'];
    const categoricalFields = state.fields.filter(f => 
        categoricalTypes.includes(f.type) && 
        f.name !== aggregateColumn && 
        f.name !== periodColumn
    );
    const numericFields = state.fields.filter(f => f.type === 'number');
    
    let chartsHtml = '';
    
    categoricalFields.forEach((field, idx) => {
        const valueCounts = {};
        filteredData.forEach(row => {
            const value = row[field.name];
            if (value) valueCounts[value] = (valueCounts[value] || 0) + 1;
        });
        
        if (Object.keys(valueCounts).length > 0) {
            chartsHtml += `
                <div class="chart-container">
                    <h4><span class="inline-icon">${getIcon('bar-chart-3', 16)}</span> ${escapeHtml(field.label)}</h4>
                    <div class="chart-wrapper">
                        <div class="chart-box"><canvas id="barChart_${idx}"></canvas></div>
                        <div class="chart-box"><canvas id="pieChart_${idx}"></canvas></div>
                    </div>
                </div>
            `;
        }
    });
    
    if (numericFields.length > 0 && aggregateData.length > 0) {
        chartsHtml += `
            <div class="chart-container">
                <h4><span data-icon="hash" data-size="14"></span> Numeric Summary (Aggregate)</h4>
                <div style="overflow-x:auto;">
                    <table class="data-table">
                        <thead><tr><th>Field</th><th>Total Sum</th><th>Average per Location</th><th>Min</th><th>Max</th></tr></thead>
                        <tbody>
                            ${numericFields.map(field => {
                                const values = aggregateData.map(d => d[field.name] || 0);
                                const sum = values.reduce((a, b) => a + b, 0);
                                const avg = values.length > 0 ? (sum / values.length).toFixed(2) : 0;
                                const min = values.length > 0 ? Math.min(...values) : 0;
                                const max = values.length > 0 ? Math.max(...values) : 0;
                                return `<tr><td><strong>${escapeHtml(field.label)}</strong></td><td>${sum}</td><td>${avg}</td><td>${min}</td><td>${max}</td></tr>`;
                            }).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        `;
    }
    
    container.innerHTML = `
        <div class="filter-panel">
            <div class="filter-header">
                <div class="filter-title"><span class="inline-icon">${getIcon('filter', 14)}</span> Filters ${activeFilterCount > 0 ? `<span class="filter-count">${activeFilterCount} active</span>` : ''}</div>
                <button class="filter-btn clear" onclick="clearAllFilters()"><span class="inline-icon">${getIcon('trash-2', 12)}</span> Clear</button>
            </div>
            <div class="filter-controls">${filtersHtml}</div>
        </div>
        
        <div class="dashboard-header">
            <img src="${state.settings.logo}" alt="Logo">
            <h2><span class="inline-icon">${getIcon('bar-chart-2', 20)}</span> ${escapeHtml(state.settings.title)} - Dashboard</h2>
            <p>ICF-SL Data Analytics</p>
        </div>
        
        <div class="dashboard-stats">
            <div class="stat-card"><div class="stat-value">${filteredData.length}</div><div class="stat-label">Total Records</div></div>
            <div class="stat-card success"><div class="stat-value">${synced}</div><div class="stat-label">Synced</div></div>
            <div class="stat-card warning"><div class="stat-value">${pending}</div><div class="stat-label">Pending</div></div>
            ${periodColumn ? `<div class="stat-card info"><div class="stat-value">${uniquePeriods}</div><div class="stat-label">Periods</div></div>` : ''}
            ${aggregateColumn ? `<div class="stat-card purple"><div class="stat-value">${uniqueGroups}</div><div class="stat-label">${escapeHtml(aggregateField?.label || 'Groups')}</div></div>` : ''}
            <div class="stat-card"><div class="stat-value">${aggregateData.length}</div><div class="stat-label">Aggregate Rows</div></div>
        </div>
        
        ${chartsHtml || '<div class="chart-container"><p style="text-align:center;color:#868e96;padding:30px;"><span class="inline-icon">' + getIcon('bar-chart-3', 18) + '</span> Add categorical fields for charts</p></div>'}
        
        ${renderGpsMapSection(filteredData, 'dashboard')}
    `;
    
    // Render charts
    setTimeout(() => {
        const colors = ['#004080', '#28a745', '#dc3545', '#ffc107', '#17a2b8', '#6f42c1', '#fd7e14', '#20c997', '#e83e8c', '#6c757d'];
        
        categoricalFields.forEach((field, idx) => {
            const valueCounts = {};
            filteredData.forEach(row => {
                const value = row[field.name];
                if (value) valueCounts[value] = (valueCounts[value] || 0) + 1;
            });
            
            if (Object.keys(valueCounts).length === 0) return;
            
            const labels = Object.keys(valueCounts);
            const values = Object.values(valueCounts);
            const total = values.reduce((a, b) => a + b, 0);
            
            const barCtx = document.getElementById(`barChart_${idx}`);
            if (barCtx && typeof Chart !== 'undefined') {
                state.chartInstances[`bar_${idx}`] = new Chart(barCtx, {
                    type: 'bar',
                    data: { labels, datasets: [{ label: 'Count', data: values, backgroundColor: colors.slice(0, labels.length) }] },
                    options: {
                        responsive: true,
                        plugins: { legend: { display: false }, title: { display: true, text: 'Bar Chart', font: { family: 'Oswald' } } },
                        scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } }
                    }
                });
            } else if (barCtx && typeof Chart === 'undefined') {
                barCtx.parentElement.innerHTML = '<p style="text-align:center;color:#868e96;padding:20px;">Charts unavailable (Chart.js not loaded)</p>';
            }
            
            const pieCtx = document.getElementById(`pieChart_${idx}`);
            if (pieCtx && typeof Chart !== 'undefined') {
                state.chartInstances[`pie_${idx}`] = new Chart(pieCtx, {
                    type: 'pie',
                    data: { 
                        labels: labels.map((l, i) => `${l} (${((values[i] / total) * 100).toFixed(1)}%)`), 
                        datasets: [{ data: values, backgroundColor: colors.slice(0, labels.length) }] 
                    },
                    options: {
                        responsive: true,
                        plugins: { legend: { position: 'bottom', labels: { font: { family: 'Oswald', size: 10 } } }, title: { display: true, text: 'Pie Chart', font: { family: 'Oswald' } } }
                    }
                });
            }
        });
        
        // Initialize GPS map after charts
        initGpsMap(filteredData, 'dashboard');
    }, 100);
    
    } catch (err) {
        console.error('Error in renderDashboard:', err);
        container.innerHTML = `<div style="text-align:center;padding:40px;color:#dc3545;"><p>Error loading dashboard</p><p style="font-size:12px;">${escapeHtml(err.message)}</p></div>`;
    }
}

// ==================== GPS MAP FUNCTIONS ====================
const gpsMapInstances = {}; // Store multiple map instances by suffix
const gpsMarkersLayers = {};

function renderGpsMapSection(filteredData, suffix = '') {
    // Find GPS fields in the form
    const gpsFields = state.fields.filter(f => f.type === 'gps');
    
    // Show map only if has GPS fields
    if (gpsFields.length === 0) {
        return ''; // No GPS fields, don't show map section
    }
    
    // Count records with GPS data
    const selectedGpsField = state.settings.gpsField || gpsFields[0]?.name;
    let recordsWithGps = 0;
    if (selectedGpsField) {
        recordsWithGps = filteredData.filter(d => {
            const gpsValue = d[selectedGpsField];
            return gpsValue && parseGpsCoordinates(gpsValue);
        }).length;
    }
    
    const gpsFieldOptions = gpsFields.length > 0 ? gpsFields.map(f => 
        `<option value="${f.name}" ${selectedGpsField === f.name ? 'selected' : ''}>${escapeHtml(f.label)}</option>`
    ).join('') : '<option value="">No GPS field in form</option>';
    
    const mapId = suffix ? `gpsMap_${suffix}` : 'gpsMap';
    
    return `
        <div class="gps-map-container">
            <h4><span class="inline-icon">${getIcon('map-pin', 16)}</span> GPS Points Map (${recordsWithGps} locations)</h4>
            
            <div style="display:flex;gap:15px;margin-bottom:15px;flex-wrap:wrap;align-items:center;">
                <div style="display:flex;align-items:center;gap:8px;">
                    <label style="font-size:11px;font-weight:bold;">GPS Field:</label>
                    <select class="filter-select" onchange="setGpsField(this.value)" style="min-width:150px;">
                        ${gpsFieldOptions}
                    </select>
                </div>
            </div>
            
            <div id="${mapId}"></div>
            
            <div class="map-controls">
                <button class="modal-btn" onclick="zoomToAllPoints()" style="font-size:11px;padding:8px 12px;">
                    <span class="inline-icon">${getIcon('maximize-2', 12)}</span> Fit All Points
                </button>
                <button class="modal-btn success" onclick="downloadGeoJSON()" style="font-size:11px;padding:8px 12px;">
                    <span class="inline-icon">${getIcon('download', 12)}</span> Download GeoJSON
                </button>
                <button class="modal-btn" onclick="downloadKML()" style="font-size:11px;padding:8px 12px;background:#17a2b8;color:white;">
                    <span class="inline-icon">${getIcon('download', 12)}</span> Download KML
                </button>
            </div>
        </div>
    `;
}

function parseGpsCoordinates(gpsString) {
    if (!gpsString) return null;
    
    // Handle different GPS formats
    // Format 1: "lat,lng" or "lat, lng"
    // Format 2: "lat,lng,accuracy" 
    // Format 3: JSON object {lat, lng}
    
    try {
        if (typeof gpsString === 'object') {
            if (gpsString.lat && gpsString.lng) {
                return { lat: parseFloat(gpsString.lat), lng: parseFloat(gpsString.lng) };
            }
        }
        
        const str = String(gpsString).trim();
        
        // Try JSON parse
        if (str.startsWith('{')) {
            const obj = JSON.parse(str);
            if (obj.lat && obj.lng) {
                return { lat: parseFloat(obj.lat), lng: parseFloat(obj.lng) };
            }
        }
        
        // Try comma-separated
        const parts = str.split(',').map(p => parseFloat(p.trim()));
        if (parts.length >= 2 && !isNaN(parts[0]) && !isNaN(parts[1])) {
            return { lat: parts[0], lng: parts[1] };
        }
    } catch (e) {
        console.log('Failed to parse GPS:', gpsString);
    }
    
    return null;
}

function initGpsMap(filteredData, suffix = '') {
    const mapId = suffix ? `gpsMap_${suffix}` : 'gpsMap';
    const mapContainer = document.getElementById(mapId);
    if (!mapContainer) return;
    
    // Check if Leaflet is available
    if (typeof L === 'undefined') {
        mapContainer.innerHTML = '<p style="text-align:center;padding:40px;color:#868e96;">Map unavailable (Leaflet not loaded)</p>';
        return;
    }
    
    // Destroy existing map for this suffix
    if (gpsMapInstances[suffix]) {
        try {
            gpsMapInstances[suffix].remove();
        } catch (e) {}
        delete gpsMapInstances[suffix];
        delete gpsMarkersLayers[suffix];
    }
    
    // Get GPS field to use
    const gpsFields = state.fields.filter(f => f.type === 'gps');
    const selectedGpsField = state.settings.gpsField || gpsFields[0]?.name;
    
    // Collect GPS points (if GPS field exists)
    const points = [];
    if (selectedGpsField) {
        filteredData.forEach((record, idx) => {
            const coords = parseGpsCoordinates(record[selectedGpsField]);
            if (coords) {
                points.push({
                    ...coords,
                    record: record,
                    index: idx
                });
            }
        });
    }
    
    // Show message if no GPS data
    if (points.length === 0) {
        if (!selectedGpsField) {
            mapContainer.innerHTML = '<p style="text-align:center;padding:40px;color:#868e96;"><span class="inline-icon">' + getIcon('map', 16) + '</span> Add a GPS field to your form to see the map.</p>';
        } else {
            mapContainer.innerHTML = '<p style="text-align:center;padding:40px;color:#868e96;"><span class="inline-icon">' + getIcon('map-pin', 16) + '</span> No GPS data collected yet. Submit data with GPS coordinates to see points on map.</p>';
        }
        return;
    }
    
    // Initialize map - default to Sierra Leone center if no points
    const defaultCenter = points.length > 0 ? [points[0].lat, points[0].lng] : [8.4606, -11.7799];
    const map = L.map(mapId).setView(defaultCenter, 8);
    gpsMapInstances[suffix] = map;
    
    // Add tile layer (OpenStreetMap)
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '© OpenStreetMap contributors',
        maxZoom: 19
    }).addTo(map);
    
    // Add markers layer
    gpsMarkersLayers[suffix] = L.layerGroup().addTo(map);
    
    // Get color field for differentiation
    const colorField = state.settings.aggregateColumns?.[0] || state.settings.aggregateColumn;
    const colorMap = {};
    const colors = ['#dc3545', '#28a745', '#007bff', '#ffc107', '#17a2b8', '#6f42c1', '#fd7e14', '#20c997', '#e83e8c', '#6c757d'];
    let colorIdx = 0;
    
    // Add markers
    points.forEach(point => {
        // Get color based on category
        let markerColor = '#004080';
        if (colorField && point.record[colorField]) {
            const val = point.record[colorField];
            if (!colorMap[val]) {
                colorMap[val] = colors[colorIdx % colors.length];
                colorIdx++;
            }
            markerColor = colorMap[val];
        }
        
        // Create custom marker icon
        const markerIcon = L.divIcon({
            className: 'custom-marker',
            html: `<div style="background:${markerColor};width:12px;height:12px;border-radius:50%;border:2px solid white;box-shadow:0 0 4px rgba(0,0,0,0.4);"></div>`,
            iconSize: [12, 12],
            iconAnchor: [6, 6]
        });
        
        const marker = L.marker([point.lat, point.lng], { icon: markerIcon });
        
        // Build popup content
        let popupContent = '<div style="max-width:250px;max-height:200px;overflow-y:auto;">';
        popupContent += `<strong style="color:#004080;">Record #${point.index + 1}</strong><br>`;
        popupContent += `<small>${point.record._timestamp || 'No timestamp'}</small><hr style="margin:5px 0;">`;
        
        // Add first few fields
        state.fields.filter(f => f.type !== 'section' && f.type !== 'gps').slice(0, 5).forEach(field => {
            const val = point.record[field.name];
            if (val) {
                popupContent += `<strong>${escapeHtml(field.label)}:</strong> ${escapeHtml(String(val))}<br>`;
            }
        });
        
        popupContent += '</div>';
        marker.bindPopup(popupContent);
        
        gpsMarkersLayers[suffix].addLayer(marker);
    });
    
    // Add legend if color field is used
    if (colorField && Object.keys(colorMap).length > 0) {
        const legend = L.control({ position: 'bottomright' });
        legend.onAdd = function() {
            const div = L.DomUtil.create('div', 'map-legend');
            div.innerHTML = '<strong style="font-size:10px;display:block;margin-bottom:5px;">' + escapeHtml(state.fields.find(f => f.name === colorField)?.label || colorField) + '</strong>';
            Object.entries(colorMap).forEach(([val, color]) => {
                div.innerHTML += `<div class="map-legend-item"><div class="map-legend-color" style="background:${color};"></div><span>${escapeHtml(val)}</span></div>`;
            });
            return div;
        };
        legend.addTo(map);
    }
    
    // Fit to all points
    if (points.length > 0) {
        const bounds = L.latLngBounds(points.map(p => [p.lat, p.lng]));
        map.fitBounds(bounds, { padding: [20, 20] });
    }
}

window.setGpsField = function(fieldName) {
    state.settings.gpsField = fieldName;
    saveToStorage();
    renderDashboard();
};

// ==================== END MAP FUNCTIONS ====================


window.zoomToAllPoints = function() {
    // Find the first available map
    const suffixes = Object.keys(gpsMapInstances);
    if (suffixes.length === 0) return;
    
    suffixes.forEach(suffix => {
        const map = gpsMapInstances[suffix];
        const markersLayer = gpsMarkersLayers[suffix];
        if (!map || !markersLayer) return;
        
        const markers = [];
        markersLayer.eachLayer(m => markers.push(m.getLatLng()));
        
        if (markers.length > 0) {
            const bounds = L.latLngBounds(markers);
            map.fitBounds(bounds, { padding: [20, 20] });
        }
    });
};

window.downloadGeoJSON = function() {
    const gpsFields = state.fields.filter(f => f.type === 'gps');
    const selectedGpsField = state.settings.gpsField || gpsFields[0]?.name;
    const filteredData = getFilteredData();
    
    const features = [];
    filteredData.forEach((record, idx) => {
        const coords = parseGpsCoordinates(record[selectedGpsField]);
        if (coords) {
            const properties = {};
            state.fields.filter(f => f.type !== 'section').forEach(field => {
                properties[field.label || field.name] = record[field.name] || '';
            });
            properties['_timestamp'] = record._timestamp || '';
            properties['_id'] = record._id || idx;
            
            features.push({
                type: 'Feature',
                geometry: {
                    type: 'Point',
                    coordinates: [coords.lng, coords.lat]
                },
                properties: properties
            });
        }
    });
    
    const geojson = {
        type: 'FeatureCollection',
        features: features
    };
    
    const blob = new Blob([JSON.stringify(geojson, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${state.settings.title}_gps_points.geojson`;
    a.click();
    URL.revokeObjectURL(url);
    notify('GeoJSON downloaded!');
};

window.downloadKML = function() {
    const gpsFields = state.fields.filter(f => f.type === 'gps');
    const selectedGpsField = state.settings.gpsField || gpsFields[0]?.name;
    const filteredData = getFilteredData();
    
    let kml = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
<name>${escapeHtml(state.settings.title)} GPS Points</name>
<description>Exported from ICF Collect</description>
<Style id="pointStyle">
<IconStyle>
    <color>ff0000ff</color>
    <scale>1.0</scale>
    <Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-circle.png</href></Icon>
</IconStyle>
</Style>
`;
    
    filteredData.forEach((record, idx) => {
        const coords = parseGpsCoordinates(record[selectedGpsField]);
        if (coords) {
            const name = record._id || `Point ${idx + 1}`;
            let description = '';
            state.fields.filter(f => f.type !== 'section' && f.type !== 'gps').slice(0, 10).forEach(field => {
                if (record[field.name]) {
                    description += `${field.label}: ${record[field.name]}\n`;
                }
            });
            
            kml += `    <Placemark>
<name>${escapeHtml(String(name))}</name>
<description><![CDATA[${description}]]></description>
<styleUrl>#pointStyle</styleUrl>
<Point><coordinates>${coords.lng},${coords.lat},0</coordinates></Point>
</Placemark>
`;
        }
    });
    
    kml += `</Document>
</kml>`;
    
    const blob = new Blob([kml], { type: 'application/vnd.google-earth.kml+xml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${state.settings.title}_gps_points.kml`;
    a.click();
    URL.revokeObjectURL(url);
    notify('KML downloaded!');
};
// ==================== END GPS MAP FUNCTIONS ====================

function showTab(tab, btn) {
    console.log('=== showTab called ===', tab);
    document.querySelectorAll('.viewer-tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.viewer-nav-btn').forEach(b => b.classList.remove('active'));
    const tabEl = document.getElementById('tab' + tab.charAt(0).toUpperCase() + tab.slice(1));
    console.log('Tab element ID:', 'tab' + tab.charAt(0).toUpperCase() + tab.slice(1));
    console.log('Tab element found:', !!tabEl);
    if (tabEl) tabEl.classList.add('active');
    if (btn) btn.classList.add('active');
    
    if (tab === 'data') {
        console.log('Rendering data tab...');
        renderDataContent();
    }
    if (tab === 'dashboard') {
        console.log('Rendering dashboard tab...');
        renderDashboard();
    }
}

function goToPage(pageIndex, direction = 1) {
    console.log('=== goToPage called, pageIndex:', pageIndex, 'direction:', direction, '===');
    const pages = document.querySelectorAll('.form-page');
    const totalPages = pages.length;
    
    if (totalPages === 0) {
        console.log('No pages found');
        return;
    }
    
    // Find a visible page with visible fields starting from pageIndex
    let targetPage = pageIndex;
    let attempts = 0;
    
    while (attempts < totalPages) {
        if (targetPage < 0 || targetPage >= totalPages) {
            targetPage = direction > 0 ? totalPages - 1 : 0;
            break;
        }
        
        const page = document.querySelector(`.form-page[data-page="${targetPage}"]`);
        if (page) {
            // Skip hidden sections
            if (page.classList.contains('section-hidden')) {
                console.log('Page', targetPage, 'is section-hidden, skipping');
                targetPage += direction;
                attempts++;
                continue;
            }
            
            // Check if page has visible fields
            const visibleFields = page.querySelectorAll('.viewer-field:not(.field-hidden)');
            if (visibleFields.length === 0) {
                console.log('Page', targetPage, 'has no visible fields, skipping');
                targetPage += direction;
                attempts++;
                continue;
            }
            
            // Found a good page
            console.log('Page', targetPage, 'is valid with', visibleFields.length, 'visible fields');
            break;
        }
        
        targetPage += direction;
        attempts++;
    }
    
    // Clamp to valid range
    if (targetPage < 0) targetPage = 0;
    if (targetPage >= totalPages) targetPage = totalPages - 1;
    
    // Hide all pages and show target
    pages.forEach(p => p.style.display = 'none');
    const targetPageEl = document.querySelector(`.form-page[data-page="${targetPage}"]`);
    if (targetPageEl) {
        targetPageEl.style.display = 'block';
        console.log('Now showing page', targetPage);
    } else {
        console.log('Target page element not found');
    }
}

// Validate required fields on current page
function validateCurrentPage(currentPageIndex) {
    console.log('=== validateCurrentPage called for page', currentPageIndex, '===');
    const currentPageEl = document.querySelector(`.form-page[data-page="${currentPageIndex}"]`);
    if (!currentPageEl) {
        console.log('Page element not found');
        return true;
    }
    
    // Skip validation if page is hidden
    if (currentPageEl.classList.contains('section-hidden')) {
        console.log('Page is hidden, skipping validation');
        return true;
    }
    
    let isValid = true;
    let firstInvalid = null;
    
    // Find all viewer-field elements on current page
    const allFieldEls = currentPageEl.querySelectorAll('.viewer-field');
    console.log('Found', allFieldEls.length, 'total fields on page');
    
    allFieldEls.forEach(fieldEl => {
        // Skip if field is hidden
        if (fieldEl.classList.contains('field-hidden')) {
            return;
        }
        
        // Check if element is actually visible (not display:none or hidden parent)
        if (fieldEl.offsetParent === null && !fieldEl.closest('.viewer-body')) {
            return;
        }
        
        const fieldName = fieldEl.dataset.fieldName;
        if (!fieldName) {
            console.log('Field element has no data-field-name');
            return;
        }
        
        const fieldDef = state.fields.find(f => f.name === fieldName);
        if (!fieldDef) {
            console.log('Field definition not found for:', fieldName);
            return;
        }
        
        if (!fieldDef.required) {
            return;
        }
        
        console.log('Validating required field:', fieldName, 'type:', fieldDef.type);
        
        // Remove previous error styling
        const existingError = fieldEl.querySelector('.field-error');
        if (existingError) existingError.remove();
        fieldEl.querySelectorAll('input, select, textarea').forEach(el => {
            el.style.borderColor = '';
        });
        
        let fieldValue = '';
        let inputEl = null;
        let validationFailed = false;
        
        // Handle different field types
        if (fieldDef.type === 'checkbox') {
            const checkboxes = fieldEl.querySelectorAll('input[type="checkbox"]');
            const anyChecked = Array.from(checkboxes).some(cb => cb.checked);
            if (!anyChecked) {
                validationFailed = true;
                inputEl = checkboxes[0];
                if (inputEl) showFieldError(inputEl, 'Please select at least one option');
                console.log('  - FAILED: No checkbox checked');
            }
        } else if (fieldDef.type === 'radio' || fieldDef.type === 'yesno') {
            const radios = fieldEl.querySelectorAll('input[type="radio"]');
            const anySelected = Array.from(radios).some(r => r.checked);
            if (!anySelected) {
                validationFailed = true;
                inputEl = radios[0];
                if (inputEl) showFieldError(inputEl, 'Please select an option');
                console.log('  - FAILED: No radio selected');
            }
        } else if (fieldDef.type === 'cascade') {
            // For cascade, check the hidden input
            const hiddenInput = fieldEl.querySelector('input[type="hidden"]');
            fieldValue = hiddenInput ? hiddenInput.value.trim() : '';
            inputEl = fieldEl.querySelector('select');
            if (!fieldValue) {
                validationFailed = true;
                if (inputEl) showFieldError(inputEl, 'Please complete all selections');
                console.log('  - FAILED: Cascade not completed');
            }
        } else if (fieldDef.type === 'gps' || fieldDef.type === 'qrcode' || fieldDef.type === 'rating') {
            // For these types, check the hidden input
            const hiddenInput = fieldEl.querySelector('input[type="hidden"]');
            fieldValue = hiddenInput ? hiddenInput.value.trim() : '';
            inputEl = hiddenInput;
            if (!fieldValue) {
                validationFailed = true;
                const btn = fieldEl.querySelector('button');
                if (btn) showFieldError(btn, 'This field is required');
                console.log('  - FAILED: Hidden input empty');
            }
        } else {
            // Text, select, textarea, number, date, time, email, phone, etc.
            inputEl = fieldEl.querySelector('input, select, textarea');
            if (inputEl) {
                fieldValue = inputEl.value ? inputEl.value.trim() : '';
            }
            if (!fieldValue) {
                validationFailed = true;
                if (inputEl) showFieldError(inputEl, 'This field is required');
                console.log('  - FAILED: Empty value');
            } else {
                console.log('  - PASSED: Value =', fieldValue.substring(0, 20));
            }
        }
        
        if (validationFailed) {
            isValid = false;
            if (!firstInvalid && inputEl) firstInvalid = inputEl;
        }
    });
    
    // Scroll to first invalid field
    if (firstInvalid) {
        firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
        try { firstInvalid.focus(); } catch(e) {}
    }
    
    console.log('Validation result:', isValid ? 'VALID' : 'INVALID');
    return isValid;
}

function showFieldError(field, message) {
    field.style.borderColor = '#dc3545';
    const errorDiv = document.createElement('div');
    errorDiv.className = 'field-error';
    errorDiv.style.cssText = 'color:#dc3545;font-size:12px;margin-top:4px;';
    errorDiv.textContent = message;
    field.parentElement.appendChild(errorDiv);
}

// Navigate to next visible page
window.goToNextPage = function(currentPage) {
    console.log('=== goToNextPage called from page', currentPage, '===');
    
    // Validate current page before moving to next
    if (!validateCurrentPage(currentPage)) {
        notify('Please fill in all required fields', 'error');
        return;
    }
    
    // Find next visible page with visible fields
    const pages = document.querySelectorAll('.form-page');
    let nextPage = currentPage + 1;
    
    while (nextPage < pages.length) {
        const pageEl = pages[nextPage];
        // Skip hidden sections
        if (pageEl.classList.contains('section-hidden')) {
            console.log('Skipping hidden page', nextPage);
            nextPage++;
            continue;
        }
        // Check if page has any visible fields
        const visibleFields = pageEl.querySelectorAll('.viewer-field:not(.field-hidden)');
        if (visibleFields.length === 0) {
            console.log('Skipping page', nextPage, 'with no visible fields');
            nextPage++;
            continue;
        }
        break;
    }
    
    if (nextPage >= pages.length) {
        console.log('No more visible pages');
        return;
    }
    
    console.log('Moving to page', nextPage);
    goToPage(nextPage, 1);
};

// Navigate to previous visible page  
window.goToPrevPage = function(currentPage) {
    console.log('=== goToPrevPage called from page', currentPage, '===');
    
    // Find previous visible page with visible fields
    const pages = document.querySelectorAll('.form-page');
    let prevPage = currentPage - 1;
    
    while (prevPage >= 0) {
        const pageEl = pages[prevPage];
        // Skip hidden sections
        if (pageEl.classList.contains('section-hidden')) {
            console.log('Skipping hidden page', prevPage);
            prevPage--;
            continue;
        }
        // Check if page has any visible fields
        const visibleFields = pageEl.querySelectorAll('.viewer-field:not(.field-hidden)');
        if (visibleFields.length === 0) {
            console.log('Skipping page', prevPage, 'with no visible fields');
            prevPage--;
            continue;
        }
        break;
    }
    
    if (prevPage < 0) {
        console.log('No previous visible pages');
        return;
    }
    
    console.log('Moving to page', prevPage);
    goToPage(prevPage, -1);
};

// Initialize cascade dropdowns - handles both cascade-container and linked fields
function initCascades() {
    // Method 1: cascade-container (single cascade field with multiple dropdowns inside)
    document.querySelectorAll('.cascade-container').forEach(container => {
        const cascadeId = container.dataset.cascadeId;
        const field = state.fields.find(f => f.id === cascadeId);
        if (!field || !field.cascadeData || field.cascadeData.length === 0) return;
        
        const columns = field.cascadeColumns || [];
        if (columns.length === 0) return;
        
        const firstSelect = container.querySelector('select[data-level="0"]');
        if (firstSelect) {
            const uniqueValues = [...new Set(field.cascadeData.map(row => row[columns[0]]).filter(Boolean))].sort();
            uniqueValues.forEach(val => {
                const option = document.createElement('option');
                option.value = val;
                option.textContent = val;
                firstSelect.appendChild(option);
            });
        }
    });
    
    // Method 2: Linked cascade fields (separate select fields with cascadeGroup)
    const cascadeGroups = {};
    state.fields.forEach(f => {
        if (f.cascadeGroup) {
            if (!cascadeGroups[f.cascadeGroup]) {
                cascadeGroups[f.cascadeGroup] = [];
            }
            cascadeGroups[f.cascadeGroup].push(f);
        }
    });
    
    Object.keys(cascadeGroups).forEach(groupId => {
        const fields = cascadeGroups[groupId].sort((a, b) => (a.cascadeLevel || 0) - (b.cascadeLevel || 0));
        const firstField = fields[0];
        
        if (!firstField || !firstField.cascadeData || firstField.cascadeData.length === 0) return;
        
        const firstSelect = document.querySelector(`select[data-cascade-group="${groupId}"][data-cascade-level="0"]`);
        if (firstSelect && firstSelect.options.length <= 1) {
            const column = firstField.cascadeColumn;
            if (!column) return;
            const uniqueValues = [...new Set(firstField.cascadeData.map(row => row[column]).filter(Boolean))].sort();
            
            uniqueValues.forEach(val => {
                const option = document.createElement('option');
                option.value = val;
                option.textContent = val;
                firstSelect.appendChild(option);
            });
        }
    });
}

// Handle cascade dropdown selection (cascade-container method)
window.handleCascadeSelect = function(select) {
    const container = select.closest('.cascade-container');
    const cascadeId = container.dataset.cascadeId;
    const field = state.fields.find(f => f.id === cascadeId);
    if (!field || !field.cascadeData) return;
    
    const columns = field.cascadeColumns || [];
    const level = parseInt(select.dataset.level);
    const selectedValue = select.value;
    
    for (let i = level + 1; i < columns.length; i++) {
        const nextSelect = container.querySelector(`select[data-level="${i}"]`);
        if (nextSelect) {
            nextSelect.innerHTML = `<option value="">-- Select ${columns[i]} --</option>`;
            nextSelect.disabled = true;
            nextSelect.value = '';
        }
    }
    
    const hiddenInput = container.querySelector('input[type="hidden"]');
    if (hiddenInput) hiddenInput.value = '';
    
    if (!selectedValue) return;
    
    const filters = {};
    for (let i = 0; i <= level; i++) {
        const sel = container.querySelector(`select[data-level="${i}"]`);
        if (sel && sel.value) {
            filters[columns[i]] = sel.value;
        }
    }
    
    const filteredData = field.cascadeData.filter(row => {
        return Object.keys(filters).every(col => row[col] === filters[col]);
    });
    
    if (level + 1 < columns.length) {
        const nextSelect = container.querySelector(`select[data-level="${level + 1}"]`);
        if (nextSelect) {
            const nextColumn = columns[level + 1];
            const uniqueValues = [...new Set(filteredData.map(row => row[nextColumn]).filter(Boolean))].sort();
            
            uniqueValues.forEach(val => {
                const option = document.createElement('option');
                option.value = val;
                option.textContent = val;
                nextSelect.appendChild(option);
            });
            
            nextSelect.disabled = false;
        }
    }
    
    const valueColumn = container.dataset.valueColumn;
    if (level === columns.length - 1 && filteredData.length > 0) {
        const value = filteredData[0][valueColumn] || filteredData[0][columns[columns.length - 1]];
        if (hiddenInput) hiddenInput.value = value;
    }
};

// Handle linked cascade field change (separate select fields method)
window.handleCascadeChange = function(select) {
    const groupId = select.dataset.cascadeGroup;
    const level = parseInt(select.dataset.cascadeLevel) || 0;
    const selectedValue = select.value;
    
    if (!groupId) return;
    
    // Find all fields in this cascade group
    const groupFields = state.fields.filter(f => f.cascadeGroup === groupId).sort((a, b) => (a.cascadeLevel || 0) - (b.cascadeLevel || 0));
    const firstField = groupFields[0];
    
    if (!firstField || !firstField.cascadeData || firstField.cascadeData.length === 0) return;
    
    const cascadeData = firstField.cascadeData;
    
    // Clear and disable all subsequent dropdowns
    for (let i = level + 1; i < groupFields.length; i++) {
        const nextSelect = document.querySelector(`select[data-cascade-group="${groupId}"][data-cascade-level="${i}"]`);
        if (nextSelect) {
            const col = groupFields[i]?.cascadeColumn || '';
            nextSelect.innerHTML = `<option value="">-- Select ${col} --</option>`;
            nextSelect.disabled = true;
            nextSelect.value = '';
        }
    }
    
    if (!selectedValue) return;
    
    // Build filter based on all selections up to current level
    const filters = {};
    for (let i = 0; i <= level; i++) {
        const sel = document.querySelector(`select[data-cascade-group="${groupId}"][data-cascade-level="${i}"]`);
        const col = groupFields[i]?.cascadeColumn;
        if (sel && sel.value && col) {
            filters[col] = sel.value;
        }
    }
    
    // Filter data based on current selections
    const filteredData = cascadeData.filter(row => {
        return Object.keys(filters).every(col => row[col] === filters[col]);
    });
    
    // If there's a next level, populate it
    if (level + 1 < groupFields.length) {
        const nextField = groupFields[level + 1];
        const nextSelect = document.querySelector(`select[data-cascade-group="${groupId}"][data-cascade-level="${level + 1}"]`);
        if (nextSelect && nextField) {
            const nextColumn = nextField.cascadeColumn;
            if (nextColumn) {
                const uniqueValues = [...new Set(filteredData.map(row => row[nextColumn]).filter(Boolean))].sort();
                
                uniqueValues.forEach(val => {
                    const option = document.createElement('option');
                    option.value = val;
                    option.textContent = val;
                    nextSelect.appendChild(option);
                });
            }
            nextSelect.disabled = false;
        }
    }
};

// Run all calculations
function runCalculations() {
    const form = document.getElementById('viewerForm');
    if (!form) return;
    
    const calcFields = state.fields.filter(f => f.type === 'calculation' && f.formula);
    
    calcFields.forEach(field => {
        const input = form.querySelector(`input[name="${field.name}"]`);
        if (!input) return;
        
        let formula = field.formula;
        
        // Replace field references with values
        const fieldRefs = formula.match(/\{([^}]+)\}/g) || [];
        fieldRefs.forEach(ref => {
            const fieldName = ref.replace(/[{}]/g, '');
            const fieldInput = form.querySelector(`[name="${fieldName}"]`);
            const value = fieldInput ? (parseFloat(fieldInput.value) || 0) : 0;
            formula = formula.replace(ref, value);
        });
        
        // Evaluate the formula safely
        try {
            // Only allow numbers, operators, parentheses, decimal, and spaces
            const safeFormula = formula.replace(/[^0-9+\-*/().%\s]/g, '');
            if (safeFormula.trim()) {
                const result = Function('"use strict"; return (' + safeFormula + ')')();
                input.value = isNaN(result) || !isFinite(result) ? '' : Math.round(result * 100) / 100;
            }
        } catch (err) {
            input.value = 'Error';
        }
    });
}

window.captureGPS = function(btn) {
    const container = btn.parentElement;
    const status = container.querySelector('.gps-status');
    const input = container.querySelector('input[type="hidden"]');
    status.innerHTML = '<span class="inline-icon">' + getIcon('loader', 12) + '</span> Getting location...';
    
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(
            pos => {
                const coords = `[${pos.coords.latitude.toFixed(6)},${pos.coords.longitude.toFixed(6)}]`;
                input.value = coords;
                status.innerHTML = '<span class="inline-icon" style="color:#28a745;">' + getIcon('check-circle', 12) + '</span> ' + coords;
            },
            err => { status.innerHTML = '<span class="inline-icon" style="color:#dc3545;">' + getIcon('x-circle', 12) + '</span> ' + err.message; }
        );
    } else {
        status.innerHTML = '<span class="inline-icon" style="color:#dc3545;">' + getIcon('x-circle', 12) + '</span> Not supported';
    }
};

window.setRating = function(star, val) {
    const container = star.parentElement;
    container.querySelectorAll('span').forEach((s, i) => s.style.opacity = i < val ? '1' : '0.3');
    container.querySelector('input').value = val;
};

// QR/Barcode Scanner
window.activeQRStream = null;

window.startQRScanner = async function(btn) {
    const container = btn.closest('.qr-scanner-container');
    const preview = container.querySelector('.qr-preview');
    const video = container.querySelector('.qr-video');
    const status = container.querySelector('.qr-status');
    const input = container.querySelector('input[type="hidden"]');
    const valueDisplay = container.querySelector('.qr-value');
    
    status.innerHTML = '<span class="inline-icon">' + getIcon('loader', 12) + '</span> Starting camera...';
    preview.style.display = 'block';
    
    try {
        const stream = await navigator.mediaDevices.getUserMedia({ 
            video: { facingMode: 'environment' } 
        });
        video.srcObject = stream;
        video.play();
        window.activeQRStream = stream;
        
        status.innerHTML = '<span class="inline-icon" style="color:#17a2b8;">' + getIcon('scan', 12) + '</span> Point camera at QR code or barcode...';
        
        // Check for BarcodeDetector API
        if ('BarcodeDetector' in window) {
            const detector = new BarcodeDetector({ formats: ['qr_code', 'ean_13', 'ean_8', 'code_128', 'code_39', 'code_93', 'upc_a', 'upc_e', 'itf', 'codabar'] });
            
            const scanLoop = async () => {
                if (!window.activeQRStream) return;
                
                try {
                    const barcodes = await detector.detect(video);
                    if (barcodes.length > 0) {
                        const value = barcodes[0].rawValue;
                        input.value = value;
                        if (valueDisplay) valueDisplay.textContent = value;
                        status.innerHTML = '<span class="inline-icon" style="color:#28a745;">' + getIcon('check-circle', 12) + '</span> Scanned!';
                        stopQRScanner(btn);
                        return;
                    }
                } catch (err) {
                    console.log('Scan error:', err);
                }
                
                if (window.activeQRStream) {
                    requestAnimationFrame(scanLoop);
                }
            };
            
            scanLoop();
        } else {
            status.innerHTML = '<span class="inline-icon" style="color:#dc3545;">' + getIcon('x-circle', 12) + '</span> Scanner not supported on this device';
        }
    } catch (err) {
        status.innerHTML = '<span class="inline-icon" style="color:#dc3545;">' + getIcon('x-circle', 12) + '</span> Camera error: ' + err.message;
        preview.style.display = 'none';
    }
};

window.stopQRScanner = function(btn) {
    const container = btn.closest('.qr-scanner-container');
    const preview = container.querySelector('.qr-preview');
    const video = container.querySelector('.qr-video');
    
    if (window.activeQRStream) {
        window.activeQRStream.getTracks().forEach(track => track.stop());
        window.activeQRStream = null;
    }
    
    video.srcObject = null;
    preview.style.display = 'none';
};

// Offline detection
function updateOnlineStatus() {
    const indicator = document.getElementById('offlineIndicator');
    if (indicator) {
        indicator.style.display = navigator.onLine ? 'none' : 'block';
    }
}

window.addEventListener('online', () => {
    updateOnlineStatus();
    notify('Back online!', 'success');
    // Try to sync offline data
    syncOfflineData();
});

window.addEventListener('offline', () => {
    updateOnlineStatus();
    notify('You are offline - data will be saved locally', 'warning');
});

// Initialize
init();
initIcons();
updateOnlineStatus();
