/* global Chart, gapi, google */

// Google Sheets configuration
// You can either define CONFIG here or load it from config.js
const DEFAULT_CONFIG = {
    // Choose your authentication method: 'csv', 'api_key', or 'oauth'
    AUTH_METHOD: 'oauth', // Using OAuth for secure access
    
    // For CSV method (easiest - just publish your sheet to web)
    CSV_URL: 'https://docs.google.com/spreadsheets/d/e/YOUR_SHEET_ID/pub?output=csv',
    
    // For API Key method (free, requires Google Cloud setup)
    SPREADSHEET_ID: '1XhNpvY1SYsvszBugJeD-gQKar9OwU4lLLiF5cTkAk6I',
    SHEET_NAME: 'PaymentOpps', // The actual tab name in your spreadsheet
    SHEET_GID: '575449998', // The gid from the URL
    RANGE: 'A1:Z1000', // Fetching all columns, first 1000 rows
    API_KEY: 'YOUR_GOOGLE_SHEETS_API_KEY',
    
    // For OAuth method (free, most secure)
    CLIENT_ID: '630950450890-lofitb7ofs2q6ae3olqv1jis88uhfjeg.apps.googleusercontent.com',
    DISCOVERY_DOCS: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
    SCOPES: 'https://www.googleapis.com/auth/spreadsheets.readonly'
};
const CONFIG = Object.assign({}, DEFAULT_CONFIG, (window.CONFIG || {}));

// Chart instances
let volumeChart, paymentMethodsChart, successRateChart, revenueByRegionChart;
let ownerPerformanceChart, winRateChart, sourceChart, ageChart;
let processorChart, segmentChart, ownerSourceChart, activitiesChart, funnelChart;

// Google API client state
let gapiInited = false;
let gisInited = false;
let tokenClient;
let isLoadingData = false; // Prevent multiple simultaneous data loads 
let dataLoaded = false; // Global date filter - initialized to show all data
let currentDateFilter = {
    startDate: null, // null means no start date restriction (include all past data)
    endDate: null, // null means no end date (include all future data)
    filterType: 'all-time'
};

// Activities date filter - separate from global
let activitiesDateFilter = {
    startDate: null, // No start date restriction - include all activities
    endDate: null,
    filterType: 'all-time'
};

// Global view mode for all charts
let globalViewMode = 'closedWon'; // 'closedWon' or 'totalOpportunities'

// Owner Source Chart view mode (separate from global)
let ownerSourceViewMode = 'closedWon'; // 'closedWon' or 'totalOpportunities'

// User filter - empty array means show all users
let selectedUsers = []; // Empty array = all users, populated array = selected users only

// Define allowed opportunity owners
const ALLOWED_OWNERS = [
    'Jaxon Eady',
    'Carson Taylor', 
    'Matt Hoffman',
    'Nicholas Linares',
    'Jacob Lucas'
];

function nameTokens(s) {
    return String(s || '').toLowerCase().replace(/[^a-z]/g, ' ').split(/\s+/).filter(Boolean);
}

// Toggle Activities section to reduce memory usage by lazy rendering and destroying the chart when hidden
function initActivitiesToggle() {
    const toggleBtn = document.getElementById('toggleActivitiesBtn');
    const section = document.getElementById('activitiesSection');
    if (!toggleBtn || !section) return;

    // Ensure hidden by default (HTML already has hidden class)
    let activitiesRendered = false;

    const destroyActivities = () => {
        try {
            if (activitiesChart) {
                activitiesChart.destroy();
                activitiesChart = null;
            }
            const metricsDiv = document.getElementById('activities-metrics');
            if (metricsDiv) metricsDiv.innerHTML = '';
            const tbody = document.getElementById('activitiesBreakdownTableBody');
            if (tbody) tbody.innerHTML = '';
        } catch (e) {
            console.warn('Failed to destroy activities resources:', e);
        }
    };

    const renderIfNeeded = () => {
        // If there is a global cached activities dataset, use it; otherwise render with empty data
        try {
            if (!activitiesRendered) {
                const dataset = window.activitiesData || [];
                renderActivitiesChart(Array.isArray(dataset) ? dataset : []);
                activitiesRendered = true;
            }
        } catch (e) {
            console.warn('Activities render failed:', e);
        }
    };

    toggleBtn.addEventListener('click', () => {
        const isHidden = section.classList.contains('hidden');
        if (isHidden) {
            section.classList.remove('hidden');
            toggleBtn.textContent = 'Hide Activities';
            renderIfNeeded();
        } else {
            section.classList.add('hidden');
            toggleBtn.textContent = 'Show Activities';
            activitiesRendered = false;
            destroyActivities();
        }
    });
}

function canonicalizeOwner(owner) {
    const t = nameTokens(owner);
    if (t.length === 0) return String(owner || '').trim();
    const combos = [
        { firsts: ['matt', 'matthew'], last: 'hoffman', canonical: 'Matt Hoffman' },
        { firsts: ['carson'], last: 'taylor', canonical: 'Carson Taylor' },
        { firsts: ['jacob', 'jake'], last: 'lucas', canonical: 'Jacob Lucas' },
        { firsts: ['jaxon', 'jaxson', 'jax'], last: 'eady', canonical: 'Jaxon Eady' },
        { firsts: ['nicholas', 'nick', 'nich'], last: 'linares', canonical: 'Nicholas Linares' }
    ];
    for (const c of combos) {
        const firstMatch = t.some(x => c.firsts.some(f => tokenPrefixMatch(x, f)));
        const lastMatch = t.some(x => tokenPrefixMatch(x, c.last));
        if (firstMatch && lastMatch) return c.canonical;
    }
    return String(owner || '').trim();
}

function tokenPrefixMatch(a, b) {
    return a.startsWith(b) || b.startsWith(a);
}

function isAllowedOwnerName(owner) {
    const t = nameTokens(owner);
    const combos = [
        { firsts: ['matt', 'matthew'], last: 'hoffman' },
        { firsts: ['carson'], last: 'taylor' },
        { firsts: ['jacob', 'jake'], last: 'lucas' },
        { firsts: ['jaxon', 'jaxson', 'jax'], last: 'eady' },
        { firsts: ['nicholas', 'nick', 'nich'], last: 'linares' }
    ];
    return combos.some(c => c.firsts.some(f => t.some(x => tokenPrefixMatch(x, f))) && t.some(x => tokenPrefixMatch(x, c.last)));
}

function ownerMatchesSelected(owner) {
    if (!selectedUsers || selectedUsers.length === 0) return true;
    const t = nameTokens(owner);
    return selectedUsers.some(sel => {
        const ts = nameTokens(sel);
        if (ts.length === 0) return false;
        const first = ts[0];
        const last = ts[ts.length - 1];
        return t.some(x => tokenPrefixMatch(x, first)) && t.some(x => tokenPrefixMatch(x, last));
    });
}

// Modern color palette based on user's brand colors
const GRADIENT_COLORS = [
    { start: 'rgba(43, 75, 255, 0.9)', end: 'rgba(11, 230, 199, 0.9)', border: 'rgba(43, 75, 255, 1)' }, // Canopy Blue to Aqua
    { start: 'rgba(11, 230, 199, 0.9)', end: 'rgba(43, 75, 255, 0.9)', border: 'rgba(11, 230, 199, 1)' }, // Aqua to Canopy Blue
    { start: 'rgba(43, 75, 255, 0.85)', end: 'rgba(43, 75, 255, 0.6)', border: 'rgba(43, 75, 255, 1)' }, // Canopy Blue gradient
    { start: 'rgba(11, 230, 199, 0.85)', end: 'rgba(11, 230, 199, 0.6)', border: 'rgba(11, 230, 199, 1)' }, // Aqua gradient
    { start: 'rgba(43, 75, 255, 0.75)', end: 'rgba(0, 0, 0, 0.85)', border: 'rgba(43, 75, 255, 1)' } // Canopy Blue to Black
];

// High-contrast categorical palette for per-owner bars (better legibility across charts)
const OWNER_COLOR_PALETTE = [
    { fill: 'rgba(31, 119, 180, 0.85)', border: '#1f77b4' }, // Blue
    { fill: 'rgba(255, 127, 14, 0.85)', border: '#ff7f0e' }, // Orange
    { fill: 'rgba(44, 160, 44, 0.85)', border: '#2ca02c' }, // Green
    { fill: 'rgba(214, 39, 40, 0.85)', border: '#d62728' }, // Red
    { fill: 'rgba(148, 103, 189, 0.85)', border: '#9467bd' }  // Purple
];

// Helper function to create gradient for chart context
function createGradient(ctx, chartArea, colorIndex) {
    if (!chartArea) return GRADIENT_COLORS[colorIndex].start;
    
    const gradient = ctx.createLinearGradient(0, chartArea.bottom, 0, chartArea.top);
    gradient.addColorStop(0, GRADIENT_COLORS[colorIndex].start);
    gradient.addColorStop(1, GRADIENT_COLORS[colorIndex].end);
    return gradient;
}

// Register datalabels plugin if available
if (typeof window !== 'undefined' && window.Chart && window.ChartDataLabels) {
    Chart.register(window.ChartDataLabels);
}

// Global datalabels formatter: round to 2 decimals for numeric values
if (typeof window !== 'undefined' && window.Chart && Chart.defaults && Chart.defaults.plugins) {
    Chart.defaults.plugins.datalabels = Object.assign({}, Chart.defaults.plugins.datalabels || {}, {
        display: false,
        formatter: function(value) {
            const n = Number(value);
            return Number.isFinite(n) ? n.toFixed(2) : String(value);
        }
    });
}

function formatCurrencyShort(n) {
    const num = Number(n) || 0;
    const abs = Math.abs(num);
    if (abs >= 1e9) return '$' + (num / 1e9).toFixed(1).replace(/\.0$/, '') + 'B';
    if (abs >= 1e6) return '$' + (num / 1e6).toFixed(1).replace(/\.0$/, '') + 'M';
    if (abs >= 1e3) return '$' + (num / 1e3).toFixed(1).replace(/\.0$/, '') + 'K';
    return '$' + Math.round(num).toLocaleString();
}

// Initialize the dashboard
document.addEventListener('DOMContentLoaded', async () => {
    // Initialize UI pieces that don't depend on data, regardless of auth flow
    initDateFilterControls();
    initActivitiesToggle();
    try {
        console.log('Dashboard initializing...');
        console.log('Auth method:', CONFIG.AUTH_METHOD);
        
        // Initialize global toggle buttons
        updateGlobalToggleButtons();
        
        // Load Google Sheets data based on auth method
        if (CONFIG.AUTH_METHOD === 'oauth') {
            console.log('Initializing OAuth...');
            await initOAuth();
            return; // OAuth will render after authentication
        } else if (CONFIG.AUTH_METHOD === 'csv') {
            const data = await fetchCSVData();
            renderCharts(data);
            initUserFilter(); // Initialize user filter after data loads
        } else {
            const data = await fetchGoogleSheetsData();
            renderCharts(data);
            initUserFilter(); // Initialize user filter after data loads
        }
    } catch (error) {
        console.error('Error initializing dashboard:', error);
        showError('Failed to load dashboard data: ' + error.message);
    }
});

// Initialize user filter dropdown
function initUserFilter() {
    console.log('Initializing user filter...');
    const userFilterSelect = document.getElementById('userFilterSelect');

    if (!userFilterSelect) {
        console.warn('User filter select element not found');
        return;
    }

    console.log('User filter element found, ALLOWED_OWNERS:', ALLOWED_OWNERS);

    // Clear existing options
    userFilterSelect.innerHTML = '';

    // Add "All Users" option first (default)
    const allUsersOption = document.createElement('option');
    allUsersOption.value = 'all';
    allUsersOption.textContent = 'All Users';
    allUsersOption.selected = true; // Default to all users selected
    userFilterSelect.appendChild(allUsersOption);

    // Add individual user options
    ALLOWED_OWNERS.forEach(owner => {
        const option = document.createElement('option');
        option.value = owner;
        option.textContent = owner;
        userFilterSelect.appendChild(option);
        console.log('Added user option:', owner);
    });

    // Set initial state to all users
    selectedUsers = [];
    console.log('🔍 DEBUG: Initial selectedUsers set to:', selectedUsers);

    // Add change event listener
    userFilterSelect.addEventListener('change', handleUserFilterChange);

    console.log('User filter initialized successfully');
}

function handleUserFilterChange(event) {
    const userFilterSelect = event.target;
    const values = Array.from(userFilterSelect.selectedOptions).map(o => o.value);

    console.log('🔍 DEBUG: User filter change triggered, values:', values);
    console.log('🔍 DEBUG: selectedUsers before change:', selectedUsers);

    // If 'All Users' is selected along with others, deselect 'all'
    if (values.includes('all') && values.length > 1) {
        const allOption = Array.from(userFilterSelect.options).find(o => o.value === 'all');
        if (allOption) allOption.selected = false;
    }

    const finalValues = Array.from(userFilterSelect.selectedOptions).map(o => o.value).filter(v => v !== 'all');
    if (finalValues.length === 0 || values.includes('all')) {
        // No specific users selected or only 'All Users' selected
        selectedUsers = [];
        console.log('User filter changed: All users');
    } else {
        selectedUsers = finalValues;
        console.log('User filter changed: Selected users:', selectedUsers);
    }

    console.log('🔍 DEBUG: selectedUsers after change:', selectedUsers);

    // Re-render charts with new user filter
    if (window.currentData) {
        console.log('🔍 DEBUG: Re-rendering charts with user filter');
        renderCharts(window.currentData);
    }
}

function initDateFilterControls() {
    const filterSelect = document.getElementById('dateFilterSelect');
    const customDateRange = document.getElementById('customDateRange');
    const customDateRangeEnd = document.getElementById('customDateRangeEnd');
    const applyFilterBtn = document.getElementById('applyFilterBtn');
    const startDateInput = document.getElementById('startDateInput');
    const endDateInput = document.getElementById('endDateInput');

    // Note: initUserFilter() is now called after data loads in the main initialization

    // Global view mode toggle buttons
    const globalClosedWonToggle = document.getElementById('globalClosedWonToggle');
    const globalTotalOppsToggle = document.getElementById('globalTotalOppsToggle');
    const globalPipelineToggle = document.getElementById('globalPipelineToggle');

    if (globalClosedWonToggle && globalTotalOppsToggle && globalPipelineToggle) {
        globalClosedWonToggle.addEventListener('click', () => {
            globalViewMode = 'closedWon';
            updateGlobalToggleButtons();
            if (window.currentData) {
                renderCharts(window.currentData);
            }
        });

        globalTotalOppsToggle.addEventListener('click', () => {
            globalViewMode = 'totalOpportunities';
            updateGlobalToggleButtons();
            if (window.currentData) {
                renderCharts(window.currentData);
            }
        });

        globalPipelineToggle.addEventListener('click', () => {
            globalViewMode = 'pipeline';
            updateGlobalToggleButtons();
            if (window.currentData) {
                renderCharts(window.currentData);
            }
        });
    }

    if (filterSelect) {
        // Set default value to 'this-quarter' and apply on first init
        if (!filterSelect.dataset.defaultApplied) {
            filterSelect.value = 'this-quarter';
            applyDateFilter('this-quarter');
            filterSelect.dataset.defaultApplied = 'true';
        }
        const updateApplyBtnVisibility = () => {
            const isCustom = filterSelect.value === 'custom';
            if (applyFilterBtn) {
                if (isCustom) {
                    applyFilterBtn.classList.remove('hidden');
                    applyFilterBtn.disabled = false;
                    applyFilterBtn.classList.remove('opacity-50', 'cursor-not-allowed');
                } else {
                    applyFilterBtn.classList.add('hidden');
                    applyFilterBtn.disabled = true;
                    applyFilterBtn.classList.add('opacity-50', 'cursor-not-allowed');
                }
            }
        };

        filterSelect.addEventListener('change', (e) => {
            const value = e.target.value;
            if (value === 'custom') {
                customDateRange.classList.remove('hidden');
                customDateRangeEnd.classList.remove('hidden');
            } else {
                customDateRange.classList.add('hidden');
                customDateRangeEnd.classList.add('hidden');
                applyDateFilter(value);
            }
            updateApplyBtnVisibility();
        });

        // Initialize visibility on load
        updateApplyBtnVisibility();

        // Auto-apply custom range when dates change (no Apply button)
        const maybeApplyCustom = () => {
            if (filterSelect.value === 'custom') {
                const s = startDateInput && startDateInput.value ? new Date(startDateInput.value) : null;
                const e = endDateInput && endDateInput.value ? new Date(endDateInput.value) : null;
                if (s) applyCustomDateFilter(s, e);
            }
        };
        if (startDateInput) startDateInput.addEventListener('change', maybeApplyCustom);
        if (endDateInput) endDateInput.addEventListener('change', maybeApplyCustom);
    }

    if (applyFilterBtn) {
        applyFilterBtn.addEventListener('click', () => {
            const filterType = filterSelect.value;
            if (filterType === 'custom') {
                const startDate = startDateInput.value ? new Date(startDateInput.value) : null;
                const endDate = endDateInput.value ? new Date(endDateInput.value) : null;
                applyCustomDateFilter(startDate, endDate);
            }
        });
    }

    // Owner Source Chart toggle buttons
    const closedWonToggle = document.getElementById('closedWonToggle');
    const totalOppsToggle = document.getElementById('totalOppsToggle');

    // Activities date filter controls
    const activitiesDateFilterSelect = document.getElementById('activitiesDateFilterSelect');
    const activitiesCustomDateRange = document.getElementById('activitiesCustomDateRange');
    const activitiesCustomDateRangeEnd = document.getElementById('activitiesCustomDateRangeEnd');
    const activitiesApplyFilterBtn = document.getElementById('activitiesApplyFilterBtn');
    const activitiesStartDateInput = document.getElementById('activitiesStartDateInput');
    const activitiesEndDateInput = document.getElementById('activitiesEndDateInput');

    // Set default value to 'this-month' explicitly for faster initial load
    if (activitiesDateFilterSelect) {
        activitiesDateFilterSelect.value = 'this-month';
        console.log('🔍 DEBUG: Set activities date filter select to this-month');
        // Apply the default filter
        applyActivitiesDateFilter('this-month');
    }

    if (activitiesDateFilterSelect) {
        activitiesDateFilterSelect.addEventListener('change', (e) => {
            const value = e.target.value;
            
            if (value === 'custom') {
                activitiesCustomDateRange.classList.remove('hidden');
                activitiesCustomDateRangeEnd.classList.remove('hidden');
            } else {
                activitiesCustomDateRange.classList.add('hidden');
                activitiesCustomDateRangeEnd.classList.add('hidden');
                applyActivitiesDateFilter(value);
            }
        });
    }

    if (activitiesApplyFilterBtn) {
        activitiesApplyFilterBtn.addEventListener('click', () => {
            const filterType = activitiesDateFilterSelect.value;
            if (filterType === 'custom') {
                const startDate = activitiesStartDateInput.value ? new Date(activitiesStartDateInput.value) : null;
                const endDate = activitiesEndDateInput.value ? new Date(activitiesEndDateInput.value) : null;
                applyCustomActivitiesDateFilter(startDate, endDate);
            }
        });
    }

    if (closedWonToggle && totalOppsToggle) {
        closedWonToggle.addEventListener('click', () => {
            ownerSourceViewMode = 'closedWon';
            updateToggleButtons();
            if (window.currentData) {
                updateOwnerSourceChart(window.currentData);
            }
        });

        totalOppsToggle.addEventListener('click', () => {
            ownerSourceViewMode = 'totalOpportunities';
            updateToggleButtons();
            if (window.currentData) {
                updateOwnerSourceChart(window.currentData);
            }
        });
    }
}

function updateToggleButtons() {
    const closedWonToggle = document.getElementById('closedWonToggle');
    const totalOppsToggle = document.getElementById('totalOppsToggle');

    if (ownerSourceViewMode === 'closedWon') {
        closedWonToggle.className = 'px-3 py-1 text-sm rounded-md bg-blue-500 text-white transition-colors';
        totalOppsToggle.className = 'px-3 py-1 text-sm rounded-md text-gray-700 hover:bg-gray-200 transition-colors';
    } else {
        closedWonToggle.className = 'px-3 py-1 text-sm rounded-md text-gray-700 hover:bg-gray-200 transition-colors';
        totalOppsToggle.className = 'px-3 py-1 text-sm rounded-md bg-blue-500 text-white transition-colors';
    }
}

function updateGlobalToggleButtons() {
    const globalClosedWonToggle = document.getElementById('globalClosedWonToggle');
    const globalTotalOppsToggle = document.getElementById('globalTotalOppsToggle');
    const globalPipelineToggle = document.getElementById('globalPipelineToggle');

    if (!globalClosedWonToggle || !globalTotalOppsToggle || !globalPipelineToggle) return;

    // Helper to set active state
    const setActive = (el, isActive) => {
        if (isActive) {
            el.classList.add('active');
        } else {
            el.classList.remove('active');
        }
    };

    setActive(globalClosedWonToggle, globalViewMode === 'closedWon');
    setActive(globalTotalOppsToggle, globalViewMode === 'totalOpportunities');
    setActive(globalPipelineToggle, globalViewMode === 'pipeline');
}

function updateOwnerSourceChart(data) {
    // Destroy existing chart
    if (ownerSourceChart) {
        ownerSourceChart.destroy();
        ownerSourceChart = null;
    }

    // Filter data based on current date filter and owner filter
    const startDate = currentDateFilter.startDate;
    const endDate = currentDateFilter.endDate;

    // Filter data based on view mode
    const filteredData = data.filter(row => {
        const dealDate = parseSheetDate(row[7]); // Column H
        // Include rows without close date when showing all-time
        const inDate = (!startDate && !endDate)
            ? true
            : (dealDate ? ((!startDate || dealDate >= startDate) && (!endDate || dealDate <= endDate)) : false);
        
        const owner = String(row[0]).trim(); // Column A
        const isAllowedOwner = isAllowedOwnerName(owner);
        const userFilterMatch = ownerMatchesSelected(owner);
        
        // Apply stage filter based on GLOBAL view mode
        let stageFilter = true;
        const isWon = String(row[4]).toLowerCase().trim() === 'closed won';
        const isLost = String(row[4]).toLowerCase().trim() === 'closed lost';
        if (globalViewMode === 'closedWon') stageFilter = isWon;
        else if (globalViewMode === 'pipeline') stageFilter = !isWon && !isLost;
        
        return inDate && isAllowedOwner && userFilterMatch && stageFilter;
    });

    // Owner Source Analysis Chart - Shows each owner's opportunities by source
    const ownerSourceData = {};
    const ownerTotals = {}; // per-owner totals for percent calcs
    const sourceTotals = {}; // overall per-source totals
    filteredData.forEach(row => {
        const owner = canonicalizeOwner(row[0] || 'Unknown'); // Column A
        const subSource = row[12] || 'Implementation'; // Column M
        const gpv = parseSheetNumber(row[9]); // Column J

        if (!ownerSourceData[owner]) {
            ownerSourceData[owner] = {};
        }
        if (!ownerSourceData[owner][subSource]) {
            ownerSourceData[owner][subSource] = { gpv: 0, count: 0 };
        }
        ownerSourceData[owner][subSource].gpv += gpv;
        ownerSourceData[owner][subSource].count += 1;

        if (!ownerTotals[owner]) ownerTotals[owner] = { gpv: 0, count: 0 };
        ownerTotals[owner].gpv += gpv;
        ownerTotals[owner].count += 1;

        if (!sourceTotals[subSource]) sourceTotals[subSource] = { gpv: 0, count: 0 };
        sourceTotals[subSource].gpv += gpv;
        sourceTotals[subSource].count += 1;
    });

    // Prepare data for grouped bar chart by OWNER with sources as series (GPV only)
    const owners = ALLOWED_OWNERS.filter(o => ownerMatchesSelected(o)); // Respect user filter
    const sources = [...new Set(filteredData.map(row => row[12] || 'Implementation'))].sort();

    // Define distinct colors for sources - map by source name for consistency
    const sourceColorMap = {
        'Implementation': 'rgba(59, 130, 246, 0.8)',  // Blue
        'Referral': 'rgba(16, 185, 129, 0.8)',       // Emerald
        'Website': 'rgba(245, 158, 11, 0.8)',        // Amber
        'Email': 'rgba(239, 68, 68, 0.8)',           // Red
        'Social Media': 'rgba(168, 85, 247, 0.8)',   // Violet
        'PPC': 'rgba(236, 72, 153, 0.8)',            // Pink
        'Content Marketing': 'rgba(34, 197, 94, 0.8)', // Green
        'SEO': 'rgba(251, 191, 36, 0.8)',            // Yellow
        'Partnership': 'rgba(139, 69, 19, 0.8)',     // Brown
        'Other': 'rgba(107, 114, 128, 0.8)'          // Gray
    };
    
    // Fallback colors for unknown sources
    const fallbackColors = [
        'rgba(59, 130, 246, 0.8)',  // Blue
        'rgba(16, 185, 129, 0.8)',  // Emerald
        'rgba(245, 158, 11, 0.8)',  // Amber
        'rgba(239, 68, 68, 0.8)',   // Red
        'rgba(168, 85, 247, 0.8)',  // Violet
        'rgba(236, 72, 153, 0.8)',  // Pink
        'rgba(34, 197, 94, 0.8)',   // Green
        'rgba(251, 191, 36, 0.8)',  // Yellow
        'rgba(139, 69, 19, 0.8)',   // Brown
        'rgba(107, 114, 128, 0.8)'  // Gray
    ];
    
    const sourceDatasets = sources.map((source, index) => ({
        label: source,
        data: owners.map(owner => ownerSourceData[owner]?.[source]?.gpv || 0),
        backgroundColor: sourceColorMap[source] || fallbackColors[index % fallbackColors.length],
        borderColor: (sourceColorMap[source] || fallbackColors[index % fallbackColors.length]).replace('0.8', '1'),
        borderWidth: 2,
        borderRadius: 6
    }));

    const ownerSourceCanvas = document.getElementById('ownerSourceChart');
    if (ownerSourceCanvas) {
        const ownerSourceCtx = ownerSourceCanvas.getContext('2d');
        ownerSourceChart = new Chart(ownerSourceCtx, {
            type: 'bar',
            data: {
                labels: owners,
                datasets: sourceDatasets
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                resizeDelay: 200,
                plugins: {
                    legend: { position: 'top' },
                    title: { 
                        display: true, 
                        text: `Owner Opportunities by Source (GPV) — ${globalViewMode === 'closedWon' ? 'Closed Won' : globalViewMode === 'pipeline' ? 'Pipeline' : 'Total Opportunities'}`
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const value = context.raw;
                                const owner = context.label;
                                const sourceName = context.dataset.label;
                                const count = ownerSourceData[owner]?.[sourceName]?.count || 0;
                                const ownerTotal = ownerTotals[owner]?.gpv || 0;
                                const pctOfOwner = ownerTotal > 0 ? ((value / ownerTotal) * 100).toFixed(1) : '0.0';
                                return `${sourceName}: $${value.toLocaleString()} (${pctOfOwner}% of ${owner}, ${count} opps)`;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        stacked: false,
                        title: {
                            display: true,
                            text: 'Opportunity Owner'
                        }
                    },
                    y: {
                        stacked: false,
                        title: {
                            display: true,
                            text: 'GPV'
                        },
                        ticks: {
                            callback: function(value) {
                                return formatCurrencyShort(value);
                            }
                        }
                    }
                }
            }
        });
    }
}

function applyDateFilter(filterType) {
    const now = new Date();
    let startDate, endDate;
    
    switch (filterType) {
        case 'all-time':
            startDate = null;
            endDate = null;
            break;
        case 'from-2025-04-01':
            startDate = new Date('2025-04-01');
            endDate = null;
            break;
        case 'this-month':
            startDate = new Date(now.getFullYear(), now.getMonth(), 1);
            endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
            break;
        case 'last-month':
            startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
            endDate = new Date(now.getFullYear(), now.getMonth(), 0);
            break;
        case 'this-quarter':
            const quarterStartMonth = Math.floor(now.getMonth() / 3) * 3;
            startDate = new Date(now.getFullYear(), quarterStartMonth, 1);
            endDate = new Date(now.getFullYear(), quarterStartMonth + 3, 0);
            break;
        case 'last-quarter':
            const lastQuarterStartMonth = Math.floor((now.getMonth() - 3) / 3) * 3;
            startDate = new Date(now.getFullYear(), lastQuarterStartMonth, 1);
            endDate = new Date(now.getFullYear(), lastQuarterStartMonth + 3, 0);
            break;
        case 'this-year':
            startDate = new Date(now.getFullYear(), 0, 1);
            endDate = new Date(now.getFullYear(), 11, 31);
            break;
        default:
            startDate = new Date('2025-04-01');
            endDate = null;
    }
    
    currentDateFilter = { startDate, endDate, filterType };
    console.log('Applied date filter:', currentDateFilter);
    
    // Re-render with new filter
    if (window.currentData) {
        renderCharts(window.currentData);
        if (window.activitiesData) {
            renderActivitiesChart(window.activitiesData);
        }
    }
}

function applyCustomDateFilter(startDate, endDate) {
    if (!startDate) {
        alert('Please select a start date');
        return;
    }
    
    currentDateFilter = { 
        startDate, 
        endDate: endDate || null, 
        filterType: 'custom' 
    };
    console.log('Applied custom date filter:', currentDateFilter);
    
    // Re-render with new filter
    if (window.currentData) {
        renderCharts(window.currentData);
        if (window.activitiesData) {
            renderActivitiesChart(window.activitiesData);
        }
    }
}

function applyActivitiesDateFilter(filterType) {
    const now = new Date();
    let startDate, endDate;
    
    switch (filterType) {
        case 'from-2025-01-01':
            startDate = new Date('2025-01-01');
            endDate = null;
            break;
        case 'all-time':
            startDate = null;
            endDate = null;
            break;
        case 'this-year':
            startDate = new Date(now.getFullYear(), 0, 1);
            endDate = new Date(now.getFullYear(), 11, 31);
            break;
        case 'this-quarter':
            const quarterStartMonth = Math.floor(now.getMonth() / 3) * 3;
            startDate = new Date(now.getFullYear(), quarterStartMonth, 1);
            endDate = new Date(now.getFullYear(), quarterStartMonth + 3, 0);
            break;
        case 'last-quarter':
            const lastQuarterStartMonth = Math.floor((now.getMonth() - 3) / 3) * 3;
            startDate = new Date(now.getFullYear(), lastQuarterStartMonth, 1);
            endDate = new Date(now.getFullYear(), lastQuarterStartMonth + 3, 0);
            break;
        case 'this-month':
            startDate = new Date(now.getFullYear(), now.getMonth(), 1);
            endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
            break;
        case 'last-month':
            startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
            endDate = new Date(now.getFullYear(), now.getMonth(), 0);
            break;
        default:
            startDate = new Date(now.getFullYear(), 0, 1);
            endDate = null;
    }
    
    activitiesDateFilter = { startDate, endDate, filterType };
    console.log('Applied activities date filter:', activitiesDateFilter);
    
    // Re-render activities chart with new filter
    if (window.activitiesData) {
        renderActivitiesChart(window.activitiesData);
    }
}

function applyCustomActivitiesDateFilter(startDate, endDate) {
    if (!startDate) {
        alert('Please select a start date');
        return;
    }
    
    activitiesDateFilter = { 
        startDate, 
        endDate: endDate || null, 
        filterType: 'custom' 
    };
    console.log('Applied custom activities date filter:', activitiesDateFilter);
    
    // Re-render activities chart with new filter
    if (window.activitiesData) {
        renderActivitiesChart(window.activitiesData);
    }
}

// Show configuration warning
function showConfigWarning() {
    const warning = document.createElement('div');
    warning.className = 'bg-yellow-100 border-l-4 border-yellow-500 text-yellow-700 p-4 mb-4';
    warning.innerHTML = `
        <p class="font-bold">Configuration Required</p>
        <p class="text-sm">You're viewing mock data. To connect to your Google Sheet:</p>
        <p class="text-sm font-semibold mt-2">🚀 Quick Start (2 minutes):</p>
        <ol class="text-sm list-decimal ml-4 mt-1">
            <li>Open your Google Sheet → File → Share → Publish to web</li>
            <li>Choose "Comma-separated values (.csv)" and click Publish</li>
            <li>Copy the URL and paste it into <code>CONFIG.CSV_URL</code> in app.js</li>
        </ol>
        <p class="text-sm mt-2">📖 See <code>SIMPLE_SETUP.md</code> for detailed instructions</p>
    `;
    document.querySelector('.container').insertBefore(warning, document.querySelector('header').nextSibling);
}

// Show error message
function showError(message) {
    const error = document.createElement('div');
    error.className = 'bg-red-100 border-l-4 border-red-500 text-red-700 p-4 mb-4';
    error.innerHTML = `
        <p class="font-bold">Error</p>
        <p class="text-sm">${message}</p>
        <p class="text-sm mt-2">Check the browser console for more details.</p>
    `;
    document.querySelector('.container').insertBefore(error, document.querySelector('header').nextSibling);
}

// Helpers to parse Google Sheets cell values
function parseSheetNumber(value) {
    if (value === null || value === undefined || value === '') return 0;
    if (typeof value === 'number') return value;
    let s = String(value).trim();
    let multiplier = 1;
    const suffixMatch = s.match(/([mk])$/i);
    if (suffixMatch) {
        const suf = suffixMatch[1].toLowerCase();
        if (suf === 'm') multiplier = 1e6;
        if (suf === 'k') multiplier = 1e3;
        s = s.slice(0, -1);
    }
    s = s.replace(/[$,]/g, '');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n * multiplier;
}

function parseSheetDate(value) {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'number') {
        // Google Sheets serial number -> JS Date (days since 1899-12-30)
        const ms = (value - 25569) * 86400000;
        return new Date(ms);
    }
    const d = new Date(String(value));
    return isNaN(d.getTime()) ? null : d;
}

// Normalize sheet duration to minutes. Supports "HH:MM:SS" or "MM:SS" strings and day-fraction numbers.
function normalizeDurationToMinutes(raw) {
    if (raw == null || raw === '') return 0;
    if (typeof raw === 'string') {
        const s = raw.trim();
        if (/^\d{1,2}:\d{2}(:\d{1,2})?$/.test(s)) {
            const parts = s.split(':').map(Number);
            let h = 0, m = 0, sec = 0;
            if (parts.length === 3) { [h, m, sec] = parts; }
            else if (parts.length === 2) { [m, sec] = parts; }
            return (h * 60) + m + (sec / 60);
        }
        const n = parseSheetNumber(s);
        // Values < 1 are likely day-fractions → minutes
        return n < 1 ? n * 24 * 60 : n;
    }
    if (typeof raw === 'number') {
        // Values < 1 → day-fraction; convert to minutes
        return raw < 1 ? raw * 24 * 60 : raw;
    }
    return 0;
}

// Initialize OAuth 2.0
async function initOAuth() {
    console.log('Loading Google API...');
    try {
        // Load the Google API client library
        await loadGoogleAPI();
        console.log('Google API loaded successfully');
        
        // Create a sign-in button
        const signInButton = document.createElement('button');
        signInButton.textContent = '🔐 Sign in with Google to view your Payment Opps data';
        signInButton.className = 'bg-blue-500 hover:bg-blue-700 text-white font-bold py-3 px-6 rounded-lg mb-6 shadow-lg transition duration-200';
        signInButton.onclick = handleAuthClick;
        
        const container = document.querySelector('.container');
        const header = document.querySelector('header');
        
        // Add info message
        const infoDiv = document.createElement('div');
        infoDiv.className = 'bg-blue-50 border-l-4 border-blue-500 text-blue-700 p-4 mb-4';
        infoDiv.innerHTML = `
            <p class="font-bold">🔒 Authentication Required</p>
            <p class="text-sm mt-1">Click the button below to sign in with your Google account and access your PaymentOpps sheet.</p>
            <p class="text-xs mt-2 text-gray-600">Spreadsheet ID: ${CONFIG.SPREADSHEET_ID}</p>
        `;
        
        container.insertBefore(infoDiv, header.nextSibling);
        container.insertBefore(signInButton, header.nextSibling.nextSibling);
    } catch (error) {
        console.error('Failed to initialize OAuth:', error);
        showError('Failed to initialize Google authentication: ' + error.message);
    }
}

// Load Google API
function loadGoogleAPI() {
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://apis.google.com/js/api.js';
        script.onload = () => {
            gapi.load('client', async () => {
                try {
                    // Initialize without API key for OAuth flow
                    await gapi.client.init({
                        discoveryDocs: CONFIG.DISCOVERY_DOCS,
                    });
                    gapiInited = true;
                    console.log('Google API client initialized successfully');
                    resolve();
                } catch (error) {
                    console.error('Error initializing Google API client:', error);
                    reject(error);
                }
            });
        };
        script.onerror = (error) => {
            console.error('Error loading Google API script:', error);
            reject(error);
        };
        document.head.appendChild(script);
    });
}

// Handle authentication
async function handleAuthClick() {
    console.log('Auth button clicked');
    
    // Prevent multiple clicks
    if (isLoadingData || dataLoaded) {
        console.log('Already authenticated or loading, ignoring click');
        return;
    }
    
    try {
        if (!tokenClient) {
            console.log('Creating token client...');
            tokenClient = google.accounts.oauth2.initTokenClient({
                client_id: CONFIG.CLIENT_ID,
                scope: CONFIG.SCOPES,
                prompt: '', // Don't show consent screen if already granted
                callback: async (response) => {
                    console.log('OAuth callback received:', response);
                    
                    // Prevent multiple loads - if data already loaded, ignore
                    if (dataLoaded) {
                        console.log('Data already loaded, ignoring callback');
                        return;
                    }
                    
                    // Prevent multiple simultaneous loads
                    if (isLoadingData) {
                        console.log('Already loading data, skipping duplicate callback');
                        return;
                    }
                    
                    if (response.error) {
                        console.error('Auth error:', response.error);
                        showError('Authentication failed: ' + response.error);
                        isLoadingData = false;
                        // Re-enable button on error
                        const signInButton = document.querySelector('button');
                        if (signInButton) {
                            signInButton.disabled = false;
                            signInButton.textContent = '🔐 Sign in with Google';
                            signInButton.className = 'bg-blue-500 hover:bg-blue-600 text-white font-bold py-3 px-6 rounded-lg mb-6 shadow-lg transition duration-200';
                        }
                        return;
                    }
                    
                    isLoadingData = true;
                    console.log('Authentication successful, fetching data...');
                    
                    // Set a timeout to prevent infinite loading
                    const loadTimeout = setTimeout(() => {
                        if (isLoadingData && !dataLoaded) {
                            console.error('Data loading timeout');
                            showError('Loading timeout - please refresh and try again');
                            isLoadingData = false;
                            const signInButton = document.querySelector('button');
                            if (signInButton) {
                                signInButton.disabled = false;
                                signInButton.textContent = '🔐 Sign in with Google';
                                signInButton.className = 'bg-blue-500 hover:bg-blue-600 text-white font-bold py-3 px-6 rounded-lg mb-6 shadow-lg transition duration-200';
                            }
                        }
                    }, 30000); // 30 second timeout
                    
                    try {
                        const data = await fetchGoogleSheetsDataOAuth();
                        clearTimeout(loadTimeout);
                        console.log('Data fetched:', data);
                        
                        // Remove sign-in UI
                        const signInButton = document.querySelector('button');
                        const infoDiv = document.querySelector('.bg-blue-50');
                        if (signInButton) signInButton.remove();
                        if (infoDiv) infoDiv.remove();
                        
                        // Store data globally for filtering
                        window.currentData = data;
                        
                        renderCharts(data);
                        
                        // Fetch and render activities data
                        try {
                            const activitiesData = await fetchActivitiesDataOAuth();
                            console.log('Activities data fetched:', activitiesData);
                            window.activitiesData = activitiesData; // do not render now; render lazily on toggle
                            const section = document.getElementById('activitiesSection');
                            if (section && !section.classList.contains('hidden')) {
                                renderActivitiesChart(activitiesData);
                            }
                        } catch (activitiesError) {
                            console.error('Error fetching activities data:', activitiesError);
                            // Show error but don't break the main dashboard
                            const activitiesSection = document.querySelector('#activities-metrics');
                            if (activitiesSection) {
                                activitiesSection.innerHTML = `
                                    <div class="bg-red-50 border-l-4 border-red-500 text-red-700 p-4 col-span-3">
                                        <p class="font-bold">Activities Data Error</p>
                                        <p class="text-sm">Unable to load activities data: ${activitiesError.message}</p>
                                    </div>
                                `;
                            }
                        }
                        
                        // Initialize user filter now that data is loaded
                        initUserFilter();
                        
                        // Initialize date filter controls now that DOM is ready
                        initDateFilterControls();
                        
                        dataLoaded = true; // Mark as loaded
                        console.log('Dashboard loaded successfully');
                    } catch (error) {
                        clearTimeout(loadTimeout);
                        console.error('Error fetching data:', error);
                        showError('Failed to fetch data: ' + error.message);
                        // Re-enable button on error
                        const signInButton = document.querySelector('button');
                        if (signInButton) {
                            signInButton.disabled = false;
                            signInButton.textContent = '🔐 Sign in with Google';
                            signInButton.className = 'bg-blue-500 hover:bg-blue-600 text-white font-bold py-3 px-6 rounded-lg mb-6 shadow-lg transition duration-200';
                        }
                    } finally {
                        isLoadingData = false;
                    }
                },
            });
        }
        
        // Disable the button immediately
        const signInButton = document.querySelector('button');
        if (signInButton) {
            signInButton.disabled = true;
            signInButton.textContent = '⏳ Authenticating...';
            signInButton.className = 'bg-gray-400 text-white font-bold py-3 px-6 rounded-lg mb-6 shadow-lg cursor-not-allowed';
        }
        
        console.log('Requesting access token...');
        tokenClient.requestAccessToken();
    } catch (error) {
        console.error('Error in handleAuthClick:', error);
        showError('Authentication error: ' + error.message);
        isLoadingData = false;
        // Re-enable button on error
        const signInButton = document.querySelector('button');
        if (signInButton) {
            signInButton.disabled = false;
            signInButton.textContent = '🔐 Sign in with Google';
            signInButton.className = 'bg-blue-500 hover:bg-blue-600 text-white font-bold py-3 px-6 rounded-lg mb-6 shadow-lg transition duration-200';
        }
    }
}

// Fetch data using OAuth
async function fetchGoogleSheetsDataOAuth() {
    console.log('Fetching from spreadsheet:', CONFIG.SPREADSHEET_ID);
    console.log('Sheet name:', CONFIG.SHEET_NAME);
    console.log('Range:', CONFIG.RANGE);
    
    try {
        // Load the Sheets API if not already loaded
        if (!gapi.client.sheets) {
            console.log('Loading Sheets API...');
            await gapi.client.load('sheets', 'v4');
            console.log('Sheets API loaded');
        }
        
        // Try different range formats
        let response;
        let sheetRange;
        
        // First try: Use gid (sheet ID) - most reliable
        try {
            sheetRange = `PaymentOpps!${CONFIG.RANGE}`;
            console.log('Trying with PaymentOpps sheet:', sheetRange);
            response = await gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: CONFIG.SPREADSHEET_ID,
                range: sheetRange,
                valueRenderOption: 'UNFORMATTED_VALUE',
                dateTimeRenderOption: 'SERIAL_NUMBER',
            });
        } catch (err1) {
            console.log('Sheet name failed, trying gid approach...');
            // Second try: Use gid directly
            try {
                sheetRange = `${CONFIG.SHEET_GID}!${CONFIG.RANGE}`;
                console.log('Trying with gid:', sheetRange);
                response = await gapi.client.sheets.spreadsheets.values.get({
                    spreadsheetId: CONFIG.SPREADSHEET_ID,
                    range: sheetRange,
                    valueRenderOption: 'UNFORMATTED_VALUE',
                    dateTimeRenderOption: 'SERIAL_NUMBER',
                });
            } catch (err2) {
                console.log('Gid failed, trying just range...');
                // Third try: Just the range (first sheet)
                sheetRange = CONFIG.RANGE;
                console.log('Trying just range:', sheetRange);
                response = await gapi.client.sheets.spreadsheets.values.get({
                    spreadsheetId: CONFIG.SPREADSHEET_ID,
                    range: sheetRange,
                    valueRenderOption: 'UNFORMATTED_VALUE',
                    dateTimeRenderOption: 'SERIAL_NUMBER',
                });
            }
        }
        
        console.log('Raw response:', response);
        console.log('Values:', response.result.values);
        
        const rows = response.result.values || [];
        if (rows.length <= 1) {
            throw new Error('No data rows found in the spreadsheet');
        }
        
        // Drop header row; keep only data rows
        const dataRows = rows.slice(1);
        
        // Debug: Log first few rows to see raw data structure
        console.log('First 3 data rows (raw from API):', dataRows.slice(0, 3));
        console.log('Column H (index 7) values:', dataRows.slice(0, 5).map(r => r[7]));
        console.log('Column J (index 9) values:', dataRows.slice(0, 5).map(r => r[9]));
        
        return dataRows;
    } catch (error) {
        console.error('Error in fetchGoogleSheetsDataOAuth:', error);
        throw new Error(`Failed to fetch spreadsheet data: ${error.message}`);
    }
}

// Fetch Activities tab data using OAuth
async function fetchActivitiesDataOAuth() {
    console.log('Fetching Activities tab data...');

    try {
        // Load the Sheets API if not already loaded
        if (!gapi.client.sheets) {
            console.log('Loading Sheets API...');
            await gapi.client.load('sheets', 'v4');
            console.log('Sheets API loaded');
        }

        // Try different range formats for Activities tab
        let response;

        // First try: List all sheets to find the correct one
        console.log('🔍 DEBUG: Fetching sheet metadata to find Activities tab...');
        try {
            const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get({
                spreadsheetId: CONFIG.SPREADSHEET_ID
            });
            console.log('🔍 DEBUG: Available sheets:', spreadsheetResponse.result.sheets.map(s => ({
                name: s.properties.title,
                gid: s.properties.sheetId
            })));
        } catch (metaError) {
            console.log('🔍 DEBUG: Could not fetch sheet metadata:', metaError);
        }

        // Try different approaches to find Activities tab
        const attempts = [
            { name: 'Activities sheet name', range: 'Activities!A:Z' }, // Fetch all rows, all columns
            { name: 'activities lowercase', range: 'activities!A:Z' },
            { name: 'Activity sheet name', range: 'Activity!A:Z' },
            { name: 'GID from URL', range: '790013075!A:Z' },
            { name: 'First sheet fallback', range: 'A:Z' }
        ];

        for (const attempt of attempts) {
            try {
                console.log(`🔍 DEBUG: Trying ${attempt.name}:`, attempt.range);
                response = await gapi.client.sheets.spreadsheets.values.get({
                    spreadsheetId: CONFIG.SPREADSHEET_ID,
                    range: attempt.range,
                    valueRenderOption: 'UNFORMATTED_VALUE',
                    dateTimeRenderOption: 'SERIAL_NUMBER',
                });
                console.log(`✅ DEBUG: Success with ${attempt.name}!`);
                break;
            } catch (err) {
                console.log(`❌ DEBUG: ${attempt.name} failed:`, err.message);
                if (attempt === attempts[attempts.length - 1]) {
                    throw new Error(`All attempts to access Activities tab failed. Last error: ${err.message}`);
                }
            }
        }

        console.log('Activities response:', response);
        console.log('Activities values:', response.result.values);

        const values = response.result.values || [];
        if (values.length <= 1) {
            throw new Error('No data rows found in Activities tab');
        }

        // Capture headers and rows
        const headers = values[0] || [];
        const dataRows = values.slice(1);
        window.activitiesHeaders = headers;

        // Debug: Log first few rows
        console.log('Activities headers:', headers);
        console.log('Activities first 3 data rows:', dataRows.slice(0, 3));

        return dataRows;
    } catch (error) {
        console.error('Error in fetchActivitiesDataOAuth:', error);
        throw new Error(`Failed to fetch Activities data: ${error.message}`);
    }
}

// Fetch data from published Google Sheets CSV (simplest method)
async function fetchCSVData() {
    const response = await fetch(CONFIG.CSV_URL);
    if (!response.ok) {
        throw new Error(`Failed to fetch CSV: ${response.statusText}`);
    }
    
    const csvText = await response.text();
    return parseCSV(csvText);
}

// Parse CSV text into structured data
function parseCSV(csvText) {
    const lines = csvText.trim().split('\n');
    if (lines.length < 2) return [];
    
    // Parse headers
    const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, ''));
    
    // Parse data rows
    return lines.slice(1).map(line => {
        const values = line.split(',').map(v => v.trim().replace(/"/g, ''));
        const entry = {};
        headers.forEach((header, index) => {
            entry[header.toLowerCase().replace(/\s+/g, '_')] = values[index] || '';
        });
        return entry;
    });
}

// Fetch data from Google Sheets using API Key
async function fetchGoogleSheetsData() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SPREADSHEET_ID}/values/${CONFIG.SHEET_NAME}!${CONFIG.RANGE}?key=${CONFIG.API_KEY}`;
    
    const response = await fetch(url);
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(`Google Sheets API error: ${response.statusText}. ${errorData.error?.message || ''}`);
    }
    
    const data = await response.json();
    return processSheetData(data.values);
}

// Process raw sheet data into structured format
function processSheetData(rows) {
    if (!rows || rows.length < 2) return [];
    
    const headers = rows[0];
    return rows.slice(1).map(row => {
        const entry = {};
        headers.forEach((header, index) => {
            entry[header.toLowerCase().replace(/\s+/g, '_')] = row[index] || '';
        });
        return entry;
    });
}

// Show metrics summary
function showMetricsSummary(metrics) {
    const container = document.querySelector('.container');
    const header = document.querySelector('header');
    
    // Remove existing summary if any
    const existing = document.getElementById('metrics-summary');
    if (existing) existing.remove();
    
    const summaryDiv = document.createElement('div');
    summaryDiv.id = 'metrics-summary';
    summaryDiv.className = 'grid grid-cols-1 md:grid-cols-3 gap-4 mb-8';
    
    // Format the current filter range for display
    const formatDate = (date) => {
        if (!date) return '';
        // Use UTC so 2025-04-01 isn't displayed as Mar 2025 in some timezones
        const y = date.getUTCFullYear();
        const m = date.getUTCMonth();
        return new Date(Date.UTC(y, m, 1)).toLocaleDateString('en-US', { month: 'short', year: 'numeric', timeZone: 'UTC' });
    };
    const startDisplay = formatDate(currentDateFilter.startDate);
    const endDisplay = currentDateFilter.endDate ? formatDate(currentDateFilter.endDate) : 'present';
    const rangeText = `${startDisplay} - ${endDisplay}`;
    const viewLabel = metrics.viewLabel || (globalViewMode === 'closedWon' ? 'Closed Won' : globalViewMode === 'pipeline' ? 'Pipeline' : 'All Opportunities');
    
    summaryDiv.innerHTML = `
        <div class="bg-white p-6 rounded-lg shadow">
            <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Avg Deal Size (${viewLabel} • ${rangeText})</h3>
            <p class="text-4xl font-bold mt-3 mb-2" style="background: linear-gradient(135deg, #2B4BFF 0%, #0BE6C7 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">$${metrics.avgDealSize.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</p>
            <p class="text-sm text-gray-600 font-medium">${metrics.dealsInView} deals in view</p>
        </div>
        <div class="bg-white p-6 rounded-lg shadow">
            <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Total GPV (${viewLabel} • ${rangeText})</h3>
            <p class="text-4xl font-bold mt-3 mb-2" style="background: linear-gradient(135deg, #0BE6C7 0%, #2B4BFF 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">$${metrics.totalGPV.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</p>
            <p class="text-sm text-gray-600 font-medium">${metrics.dealsInView} total opportunities in view</p>
        </div>
        <div class="bg-white p-6 rounded-lg shadow">
            <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Team Win Rate (Rolling 6 months)</h3>
            <p class="text-4xl font-bold mt-3 mb-2" style="background: linear-gradient(135deg, #2B4BFF 0%, #000000 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">${metrics.winRate.toFixed(1)}%</p>
            <p class="text-sm text-gray-600 font-medium">Avg deal age: ${Math.round(metrics.avgDealAge)} days</p>
        </div>
    `;
    
    // Insert after header
    const nextElement = header.nextSibling;
    container.insertBefore(summaryDiv, nextElement);
}

// Render Key Metrics Cards
function renderKeyMetrics(data) {
    // Apply current filters to data first (Owner, Date)
    const filteredData = data.filter(row => {
        // Owner Filter
        const owner = String(row[0] || '').trim();
        if (!isAllowedOwnerName(owner)) return false;
        if (!ownerMatchesSelected(owner)) return false;

        // Date Filter
        const dealDate = parseSheetDate(row[7]);
        const startDate = currentDateFilter.startDate;
        const endDate = currentDateFilter.endDate;
        if (startDate && (!dealDate || dealDate < startDate)) return false;
        if (endDate && (!dealDate || dealDate > endDate)) return false;

        return true;
    });

    let totalGPV = 0;
    let activePipeline = 0;
    let dealsWon = 0;

    filteredData.forEach(row => {
        const stage = String(row[4] || '').toLowerCase().trim();
        const gpv = parseSheetNumber(row[9]);
        
        if (stage === 'closed won') {
            totalGPV += gpv;
            dealsWon++;
        } else if (stage !== 'closed lost' && stage !== 'lost' && stage !== '') {
            activePipeline += gpv;
        }
    });

    const avgDealSize = dealsWon > 0 ? totalGPV / dealsWon : 0;

    const elTotalGPV = document.getElementById('metricTotalGPV');
    const elPipeline = document.getElementById('metricPipeline');
    const elDealsWon = document.getElementById('metricDealsWon');
    const elAvgDealSize = document.getElementById('metricAvgDealSize');

    if (elTotalGPV) elTotalGPV.textContent = formatCurrencyShort(totalGPV);
    if (elPipeline) elPipeline.textContent = formatCurrencyShort(activePipeline);
    if (elDealsWon) elDealsWon.textContent = dealsWon;
    if (elAvgDealSize) elAvgDealSize.textContent = formatCurrencyShort(avgDealSize);
}

// Render Top Performers Leaderboard
function renderLeaderboard(data) {
    // Apply current filters (Owner, Date)
    const filteredData = data.filter(row => {
        const owner = String(row[0] || '').trim();
        if (!isAllowedOwnerName(owner)) return false;
        if (!ownerMatchesSelected(owner)) return false;

        const dealDate = parseSheetDate(row[7]);
        const startDate = currentDateFilter.startDate;
        const endDate = currentDateFilter.endDate;
        if (startDate && (!dealDate || dealDate < startDate)) return false;
        if (endDate && (!dealDate || dealDate > endDate)) return false;

        return true;
    });

    const tbody = document.getElementById('leaderboardBody');
    if (!tbody) return;

    const stats = {};
    filteredData.forEach(row => {
        const owner = canonicalizeOwner(row[0] || 'Unknown');
        const stage = String(row[4] || '').toLowerCase().trim();
        const gpv = parseSheetNumber(row[9]);

        if (!stats[owner]) stats[owner] = { gpv: 0, deals: 0, won: 0, total: 0 };
        
        stats[owner].total++;
        if (stage === 'closed won') {
            stats[owner].gpv += gpv;
            stats[owner].deals++; // Won deals
            stats[owner].won++;
        }
    });

    const sortedOwners = Object.keys(stats)
        .filter(owner => stats[owner].gpv > 0 || stats[owner].won > 0)
        .sort((a, b) => stats[b].gpv - stats[a].gpv);

    tbody.innerHTML = sortedOwners.map((owner, index) => {
        const s = stats[owner];
        const winRate = s.total > 0 ? ((s.won / s.total) * 100).toFixed(1) : 0;
        const rank = index + 1;
        const rankColor = rank === 1 ? 'text-yellow-500' : rank === 2 ? 'text-gray-400' : rank === 3 ? 'text-yellow-700' : 'text-gray-500';
        
        return `
            <tr class="hover:bg-gray-50 transition-colors">
                <td class="px-6 py-4 whitespace-nowrap">
                    <div class="flex items-center">
                        <span class="font-bold mr-3 ${rankColor}">#${rank}</span>
                        <div class="text-sm font-medium text-gray-900">${owner}</div>
                    </div>
                </td>
                <td class="px-6 py-4 whitespace-nowrap text-sm font-semibold text-gray-900">$${s.gpv.toLocaleString()}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${s.won}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                    <div class="flex items-center">
                        <span class="mr-2 w-10 text-right">${winRate}%</span>
                        <div class="w-24 bg-gray-200 rounded-full h-2">
                            <div class="bg-blue-500 h-2 rounded-full" style="width: ${winRate}%"></div>
                        </div>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

// Render Funnel Chart
function renderFunnelChart(data) {
    // Apply filters
    const filteredData = data.filter(row => {
        const owner = String(row[0] || '').trim();
        if (!isAllowedOwnerName(owner)) return false;
        if (!ownerMatchesSelected(owner)) return false;

        const dealDate = parseSheetDate(row[7]);
        const startDate = currentDateFilter.startDate;
        const endDate = currentDateFilter.endDate;
        if (startDate && (!dealDate || dealDate < startDate)) return false;
        if (endDate && (!dealDate || dealDate > endDate)) return false;
        return true;
    });

    const funnelGroups = {
        'Discovery/Qual': 0,
        'Proposal/Quote': 0,
        'Negotiation/Review': 0,
        'Closed Won': 0
    };

    filteredData.forEach(row => {
        const s = String(row[4] || '').toLowerCase().trim();
        if (s.includes('won')) funnelGroups['Closed Won']++;
        else if (s.includes('negotiat') || s.includes('review') || s.includes('contract')) funnelGroups['Negotiation/Review']++;
        else if (s.includes('prop') || s.includes('quote') || s.includes('present')) funnelGroups['Proposal/Quote']++;
        else if (s.includes('disc') || s.includes('qual') || s.includes('lead') || s.includes('meet') || s.includes('sched')) funnelGroups['Discovery/Qual']++;
    });
    
    const labels = ['Discovery/Qual', 'Proposal/Quote', 'Negotiation/Review', 'Closed Won'];
    const values = labels.map(l => funnelGroups[l]);

    const ctx = document.getElementById('funnelChart');
    if (!ctx) return;

    if (funnelChart) funnelChart.destroy();

    funnelChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Deals',
                data: values,
                backgroundColor: [
                    'rgba(43, 75, 255, 0.5)',
                    'rgba(43, 75, 255, 0.7)',
                    'rgba(43, 75, 255, 0.9)',
                    'rgba(11, 230, 199, 0.9)'
                ],
                borderRadius: 4,
                barPercentage: 0.6
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: (ctx) => `${ctx.raw} deals`
                    }
                },
                datalabels: {
                    anchor: 'end',
                    align: 'end',
                    color: '#333',
                    font: { weight: 'bold' },
                    formatter: (value) => value > 0 ? value : ''
                }
            },
            scales: {
                x: { display: false, grid: { display: false } },
                y: { 
                    grid: { display: false },
                    ticks: { font: { weight: '600', size: 12 } }
                }
            }
        }
    });
}

// Calculate metrics from Payment Opps data
function calculateMetrics(data) {
    console.log('Calculating metrics from data:', data);
    
    // Debug: Show unique owners in data
    const uniqueOwners = [...new Set(data.map(row => String(row[0] || '').trim()))].filter(o => o);
    console.log('🔍 DEBUG: Unique owners in raw data:', uniqueOwners);
    
    // Use current date filter (no stage filter here)
    const startDate = currentDateFilter.startDate;
    const endDate = currentDateFilter.endDate;
    
    // Deals in range and by allowed owners (and user filter)
    console.log('🔍 DEBUG calculateMetrics: startDate:', startDate, 'endDate:', endDate, 'view:', globalViewMode);
    console.log('🔍 DEBUG calculateMetrics: selectedUsers:', selectedUsers);
    console.log('🔍 DEBUG calculateMetrics: Total data rows:', data.length);
    
    const inRangeOwnerDeals = data.filter(row => {
        const dealDate = parseSheetDate(row[7]); // Column H
        const inDate = (!startDate && !endDate)
            ? true
            : (dealDate ? ((!startDate || dealDate >= startDate) && (!endDate || dealDate <= endDate)) : false);
        const owner = String(row[0] || '').trim(); // Column A
        const isAllowedOwner = isAllowedOwnerName(owner);
        
        // Apply user filter (empty array means all users)
        const userFilterMatch = ownerMatchesSelected(owner);
        
        return inDate && isAllowedOwner && userFilterMatch;
    });
    
    // Debug: Show unique owners in filtered deals
    const filteredOwners = [...new Set(inRangeOwnerDeals.map(row => String(row[0] || '').trim()))].filter(o => o);
    console.log('🔍 DEBUG calculateMetrics: Filtered deals:', inRangeOwnerDeals.length);
    console.log('🔍 DEBUG calculateMetrics: Unique owners in filtered deals:', filteredOwners);
    
    console.log('Deals in range for metrics:', inRangeOwnerDeals.length, 'filter:', currentDateFilter);
    
    // Apply stage filter according to global view for the summary metrics
    const dealsInView = inRangeOwnerDeals.filter(row => {
        const stage = String(row[4] || '').toLowerCase().trim(); // Column E
        const isWon = stage === 'closed won';
        const isLost = stage === 'closed lost';
        if (globalViewMode === 'closedWon') return isWon;
        if (globalViewMode === 'pipeline') return !isWon && !isLost;
        return true; // totalOpportunities
    });

    // Average deal size and totals based on deals in view
    const gpvValues = dealsInView
        .map(row => parseSheetNumber(row[9])) // Column J
        .filter(val => val > 0);

    const avgDealSize = gpvValues.length > 0
        ? gpvValues.reduce((a, b) => a + b, 0) / gpvValues.length
        : 0;

    const totalGPV = gpvValues.reduce((a, b) => a + b, 0);
    const dealCount = dealsInView.length;
    
    console.log('Metrics:', { avgDealSize, totalGPV, dealCount });
    
    // Win rate = Closed Won / (Closed Won + Closed Lost) - Rolling 6 months
    const now = new Date();
    const sixMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 6, now.getDate());
    
    // Filter deals for the last 6 months from the full dataset (respect user filter & allowed owners)
    // This avoids inflating the win rate when the global view is set to Closed Won or Pipeline
    const sixMonthDeals = data.filter(row => {
        const dealDate = parseSheetDate(row[7]); // Column H
        const owner = String(row[0] || '').trim();
        const isAllowedOwner = isAllowedOwnerName(owner);
        const userFilterMatch = ownerMatchesSelected(owner);
        return isAllowedOwner && userFilterMatch && dealDate && dealDate >= sixMonthsAgo && dealDate <= now;
    });
    
    const sixMonthClosedWon = sixMonthDeals.filter(row => {
        const s = String(row[4] || '').toLowerCase().trim();
        return s === 'closed won' || s.startsWith('closed won');
    });
    const sixMonthClosedLost = sixMonthDeals.filter(row => {
        const s = String(row[4] || '').toLowerCase().trim();
        return s === 'closed lost' || s.startsWith('closed lost');
    });
    
    const sixMonthAttempts = sixMonthClosedWon.length + sixMonthClosedLost.length;
    const winRate = sixMonthAttempts > 0 ? (sixMonthClosedWon.length / sixMonthAttempts) * 100 : 0;
    
    // Average deal age from deals in view (Column G index 6)
    const avgDealAge = dealsInView.length > 0
        ? dealsInView.reduce((sum, row) => sum + (parseInt(row[6]) || 0), 0) / dealsInView.length
        : 0;
    
    const totalOpportunities = inRangeOwnerDeals.length;
    console.log('Additional metrics:', { totalOpportunities, winRate, avgDealAge });
    
    return {
        avgDealSize,
        totalGPV,
        dealCount,
        thisMonthDeals: inRangeOwnerDeals,
        allDeals: data,
        totalOpportunities,
        winRate,
        avgDealAge,
        dealsInView: dealCount,
        viewLabel: (globalViewMode === 'closedWon' ? 'Closed Won' : globalViewMode === 'pipeline' ? 'Pipeline' : 'All Opportunities')
    };
}

// Calculate metrics from Activities data
function calculateActivitiesMetrics(activitiesData) {
    console.log('🔍 DEBUG: Starting activities metrics calculation');
    console.log('🔍 DEBUG: Raw activities data length:', activitiesData.length);

    if (!activitiesData || activitiesData.length === 0) {
        console.log('🔍 DEBUG: No activities data available');
        return {
            totalCalls: 0,
            totalEmails: 0,
            demoSets: 0,
            avgCallDurationForDemos: 0,
            callsPerDemo: 0,
            emailsPerDemo: 0,
            callsPerAnswered: 0,
            filteredActivities: 0
        };
    }

    // Column mapping based on headers:
    // 0: Date, 1: Assigned, 2: Assigned Role, 3: Contact, 4: Lead, 5: Subject,
    // 6: Call Type, 7: Activity Type, 8: Call Duration (minutes), 9: Call Result,
    // 10: Activity Type WFR, 11: Disposition, 12: Set By, 13: Activity Created By,
    // 14: Set Source, 15: Task, 16: Status, 17: Created Date, 18: Step Number,
    // 19: Attributed Sequence Name, 20: Queue Name CP, 21: Connect

    const COL_DATE = 0;
    const COL_ASSIGNED = 1;
    const COL_ASSIGNED_ROLE = 2;
    const COL_ACTIVITY_TYPE = 7;
    const COL_ACTIVITY_TYPE_WFR = 10;
    const COL_CALL_DURATION = 8;
    const COL_CALL_RESULT = 9;
    const COL_DISPOSITION = 11;
    const COL_CONNECT = 21;

    // Use activities date filter
    const startDate = activitiesDateFilter.startDate;
    const endDate = activitiesDateFilter.endDate;
    console.log('🔍 DEBUG: Activities date filter:', activitiesDateFilter);
    console.log('🔍 DEBUG: startDate:', startDate, 'endDate:', endDate, 'filterType:', activitiesDateFilter.filterType);

    // Step 1: Filter for specific Assigned Role (column C)
    let roleFilteredActivities = activitiesData.filter(row => {
        const assignedRole = String(row[COL_ASSIGNED_ROLE] || '').toLowerCase().trim();
        return assignedRole === 'account manager - payments';
    });
    // Exclude specific users
    const EXCLUDED_USERS = ['forrest bernhardt', 'dylan favila'];
    roleFilteredActivities = roleFilteredActivities.filter(row => {
        const assignedName = String(row[COL_ASSIGNED] || '').toLowerCase().trim();
        return !EXCLUDED_USERS.includes(assignedName);
    });
    console.log('🔍 DEBUG: Activities after user filter:', roleFilteredActivities.length);

    // Step 2: Filter by date range
    let dateFilteredActivities = roleFilteredActivities.filter(row => {
        const dateCell = row[COL_DATE];
        const activityDate = parseSheetDate(dateCell);

        // Allow rows with empty dates to pass through (include them)
        if (!activityDate) {
            return true;
        }

        const afterStart = !startDate || activityDate >= startDate;
        const beforeEnd = !endDate || activityDate <= endDate;

        return afterStart && beforeEnd;
    });
    console.log('🔍 DEBUG: Activities in date range:', dateFilteredActivities.length);

    // Step 3: Process activities and count metrics
    let totalCalls = 0;
    let totalEmails = 0;
    let demoSets = 0;
    let callDurationsForDemos = [];
    let connectedCalls = 0;
    // Per-user aggregation map
    const perUserMap = {};

    dateFilteredActivities.forEach((row, index) => {
        // Get values from relevant columns
        const activityTypeRaw = String(row[COL_ACTIVITY_TYPE] || '').toLowerCase().trim();
        const activityTypeWfr = String(row[COL_ACTIVITY_TYPE_WFR] || '').toLowerCase().trim();
        const activityType = activityTypeWfr || activityTypeRaw;
        const callResult = String(row[COL_CALL_RESULT] || '').toLowerCase().trim();
        const callDuration = normalizeDurationToMinutes(row[COL_CALL_DURATION]);
        const disposition = String(row[COL_DISPOSITION] || '').toLowerCase().trim();
        const connect = String(row[COL_CONNECT] || '').toLowerCase().trim();
        const assignedUser = String(row[COL_ASSIGNED] || 'Unknown').trim() || 'Unknown';

        // Debug first few rows
        if (index < 5) {
            console.log(`🔍 DEBUG: Row ${index}: ActivityType="${activityType}", CallResult="${callResult}", Duration=${callDuration}`);
        }

        // Count calls and emails using normalized activity type
        const isCall = activityType.includes('call') || activityType.includes('phone');
        const isEmail = activityType.includes('email') || activityType.includes('e-mail');

        if (isCall) {
            totalCalls++;
            // Check if call was connected
            if (connect === 'connected' || connect === 'yes' || disposition === 'connected') {
                connectedCalls++;
            }
            // per-user
            if (!perUserMap[assignedUser]) perUserMap[assignedUser] = { calls: 0, emails: 0, demos: 0, connected: 0, demoDurSum: 0, demoDurCnt: 0 };
            perUserMap[assignedUser].calls += 1;
            if (connect === 'connected' || connect === 'yes' || disposition === 'connected') {
                perUserMap[assignedUser].connected += 1;
            }
        } else if (isEmail) {
            totalEmails++;
            if (!perUserMap[assignedUser]) perUserMap[assignedUser] = { calls: 0, emails: 0, demos: 0, connected: 0, demoDurSum: 0, demoDurCnt: 0 };
            perUserMap[assignedUser].emails += 1;
        }

        // Count demo sets
        if (callResult === 'demo set' || callResult.includes('demo')) {
            demoSets++;
            if (callDuration > 0) {
                callDurationsForDemos.push(callDuration);
            }
            if (!perUserMap[assignedUser]) perUserMap[assignedUser] = { calls: 0, emails: 0, demos: 0, connected: 0, demoDurSum: 0, demoDurCnt: 0 };
            perUserMap[assignedUser].demos += 1;
            if (callDuration > 0) {
                perUserMap[assignedUser].demoDurSum += callDuration;
                perUserMap[assignedUser].demoDurCnt += 1;
            }
        }
    });

    // Calculate derived metrics
    const avgCallDurationForDemos = callDurationsForDemos.length > 0
        ? callDurationsForDemos.reduce((a, b) => a + b, 0) / callDurationsForDemos.length
        : 0;

    const callsPerDemo = demoSets > 0 ? totalCalls / demoSets : 0;
    const emailsPerDemo = demoSets > 0 ? totalEmails / demoSets : 0;
    const callsPerAnswered = totalCalls > 0 ? (connectedCalls / totalCalls) * 100 : 0;

    console.log('🔍 DEBUG: Final metrics:', {
        totalCalls,
        totalEmails,
        demoSets,
        avgCallDurationForDemos,
        callsPerDemo,
        emailsPerDemo,
        callsPerAnswered
    });

    // Build per-user metrics array
    const perUser = Object.entries(perUserMap).map(([user, s]) => ({
        user,
        totalCalls: s.calls,
        totalEmails: s.emails,
        demoSets: s.demos,
        avgCallDurationForDemos: s.demoDurCnt > 0 ? s.demoDurSum / s.demoDurCnt : 0,
        callsPerDemo: s.demos > 0 ? s.calls / s.demos : 0,
        emailsPerDemo: s.demos > 0 ? s.emails / s.demos : 0,
        callsPerAnswered: s.calls > 0 ? (s.connected / s.calls) * 100 : 0,
        totalActivities: s.calls + s.emails
    })).sort((a, b) => b.totalActivities - a.totalActivities);

    return {
        totalCalls,
        totalEmails,
        demoSets,
        avgCallDurationForDemos,
        callsPerDemo,
        emailsPerDemo,
        callsPerAnswered,
        filteredActivities: dateFilteredActivities.length,
        perUser
    };
}

// Render all charts
function renderCharts(data) {
    // Update Key Metrics and New Charts
    renderKeyMetrics(data);
    renderLeaderboard(data);
    renderFunnelChart(data);

    console.log('Rendering charts with data:', data);
    
    // Destroy existing charts to prevent duplicates
    if (volumeChart) {
        volumeChart.destroy();
        volumeChart = null;
    }
    if (paymentMethodsChart) {
        paymentMethodsChart.destroy();
        paymentMethodsChart = null;
    }
    if (successRateChart) {
        successRateChart.destroy();
        successRateChart = null;
    }
    if (revenueByRegionChart) {
        revenueByRegionChart.destroy();
        revenueByRegionChart = null;
    }
    // Destroy new charts
    if (ownerPerformanceChart) {
        ownerPerformanceChart.destroy();
        ownerPerformanceChart = null;
    }
    if (winRateChart) {
        winRateChart.destroy();
        winRateChart = null;
    }
    if (sourceChart) {
        sourceChart.destroy();
        sourceChart = null;
    }
    if (ageChart) {
        ageChart.destroy();
        ageChart = null;
    }
    if (processorChart) {
        processorChart.destroy();
        processorChart = null;
    }
    if (segmentChart) {
        segmentChart.destroy();
        segmentChart = null;
    }
    
    // Calculate metrics (this already filters for Closed Won deals)
    const metrics = calculateMetrics(data);
    
    // Show metrics summary at top
    showMetricsSummary(metrics);
    
    // Filter data based on global view mode for charts
    const startDate = currentDateFilter.startDate;
    const endDate = currentDateFilter.endDate;
    const chartData = data.filter(row => {
        const dealDate = parseSheetDate(row[7]); // Column H
        const inDate = (!startDate && !endDate)
            ? true
            : (dealDate ? ((!startDate || dealDate >= startDate) && (!endDate || dealDate <= endDate)) : false);
        const owner = String(row[0]).trim(); // Column A
        const isAllowedOwner = isAllowedOwnerName(owner);
        const isClosedWon = String(row[4]).toLowerCase().trim() === 'closed won'; // Column E
        
        // Apply user filter (empty array means all users)
        const userFilterMatch = ownerMatchesSelected(owner);
        
        // Apply stage filter based on global view mode
        let stageFilter = true;
        const isClosedLost = String(row[4]).toLowerCase().trim() === 'closed lost';
        if (globalViewMode === 'closedWon') stageFilter = isClosedWon;
        else if (globalViewMode === 'pipeline') stageFilter = !isClosedWon && !isClosedLost;
        
        return inDate && isAllowedOwner && userFilterMatch && stageFilter;
    });
    
    console.log('Charts will use data filtered by global view mode:', globalViewMode, 'deals:', chartData.length);
    
    // Group deals by month for trend analysis - WITH INDIVIDUAL BREAKDOWN
    const dealsByMonth = {};
    const ownerMonthlyGPV = {};
    const ownerMonthlyDeals = {};
    const ownerMonthlyAvgDealSize = {};
    chartData.forEach(row => {
        const closeDate = row[7]; // Column H
        const owner = canonicalizeOwner(row[0] || 'Unknown'); // Column A
        const gpv = parseSheetNumber(row[9]); // Column J
        
        if (closeDate) {
            const date = parseSheetDate(closeDate);
            const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
            
            // Total deals by month
            if (!dealsByMonth[monthKey]) {
                dealsByMonth[monthKey] = [];
            }
            dealsByMonth[monthKey].push(gpv);
            
            // Owner-specific GPV by month
            if (!ownerMonthlyGPV[monthKey]) {
                ownerMonthlyGPV[monthKey] = {};
            }
            if (!ownerMonthlyGPV[monthKey][owner]) {
                ownerMonthlyGPV[monthKey][owner] = 0;
            }
            ownerMonthlyGPV[monthKey][owner] += gpv;
            
            // Owner-specific deal count by month
            if (!ownerMonthlyDeals[monthKey]) {
                ownerMonthlyDeals[monthKey] = {};
            }
            if (!ownerMonthlyDeals[monthKey][owner]) {
                ownerMonthlyDeals[monthKey][owner] = 0;
            }
            ownerMonthlyDeals[monthKey][owner] += 1;
        }
    });
    const monthLabels = Object.keys(dealsByMonth).sort((a, b) => {
        const [yearA, monthA] = a.split('-').map(Number);
        const [yearB, monthB] = b.split('-').map(Number);
        if (yearA !== yearB) return yearA - yearB;
        return monthA - monthB;
    });
    const monthlyTotals = monthLabels.map(month => 
        dealsByMonth[month].reduce((a, b) => a + b, 0)
    );
    const monthlyAvgDealSize = monthLabels.map(month => {
        const arr = dealsByMonth[month];
        return arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0;
    });
    
    // Monthly GPV Trend - STACKED BY OWNER
    const volumeCanvas = document.getElementById('volumeChart');
    if (volumeCanvas) {
        const volumeCtx = volumeCanvas.getContext('2d');
        
        // Prepare datasets for each owner
        const ownerDatasets = ALLOWED_OWNERS.map((owner, index) => {
            const ownerData = monthLabels.map(month => ownerMonthlyGPV[month]?.[owner] || 0);
            return {
                label: owner,
                data: ownerData,
                backgroundColor: OWNER_COLOR_PALETTE[index % OWNER_COLOR_PALETTE.length].fill,
                borderColor: OWNER_COLOR_PALETTE[index % OWNER_COLOR_PALETTE.length].border,
                borderWidth: 2,
                borderRadius: 8,
                stack: 'gpv'
            };
        });
        
        volumeChart = new Chart(volumeCtx, {
        type: 'bar',
        data: {
            labels: monthLabels.map(m => {
                const [year, month] = m.split('-');
                return new Date(year, month - 1).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
            }),
            datasets: ownerDatasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 200,
            animation: {
                duration: 800,
                easing: 'easeInOutQuart'
            },
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: { 
                    display: true,
                    position: 'top',
                    labels: {
                        font: {
                            size: 12,
                            weight: '500'
                        },
                        padding: 15,
                        usePointStyle: true,
                        pointStyle: 'circle'
                    }
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    backgroundColor: 'rgba(255, 255, 255, 0.95)',
                    titleColor: '#000000',
                    bodyColor: '#000000',
                    borderColor: 'rgba(43, 75, 255, 0.5)',
                    borderWidth: 1,
                    padding: 12,
                    displayColors: true,
                    callbacks: {
                        label: function(context) {
                            const owner = context.dataset.label;
                            const gpv = context.parsed.y;
                            const totalGPV = monthlyTotals[context.dataIndex];
                            const percentage = totalGPV > 0 ? ((gpv / totalGPV) * 100).toFixed(1) : 0;
                            return `${owner}: $${gpv.toLocaleString()} (${percentage}%)`;
                        },
                        footer: function(context) {
                            const totalGPV = monthlyTotals[context[0].dataIndex];
                            return `Total: $${totalGPV.toLocaleString()}`;
                        }
                    }
                },
                datalabels: {
                    display: function(ctx) {
                        const chart = ctx.chart;
                        for (let i = chart.data.datasets.length - 1; i >= 0; i--) {
                            const meta = chart.getDatasetMeta(i);
                            if (!meta.hidden) return ctx.datasetIndex === i;
                        }
                        return false;
                    },
                    anchor: 'end',
                    align: 'top',
                    color: '#111111',
                    font: { weight: '600', size: 11 },
                    clip: false,
                    formatter: function(value, ctx) {
                        return formatCurrencyShort(monthlyTotals[ctx.dataIndex]);
                    }
                },
                title: { 
                    display: true, 
                    text: 'Monthly GPV Trend by Owner',
                    font: {
                        size: 16,
                        weight: '600'
                    },
                    padding: 20
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: {
                        display: false
                    },
                    title: {
                        display: true,
                        text: 'Month',
                        font: {
                            size: 13,
                            weight: '500'
                        }
                    },
                    ticks: {
                        font: {
                            size: 11
                        }
                    }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    grid: { 
                        display: true,
                        color: 'rgba(0, 0, 0, 0.05)'
                    },
                    ticks: {
                        callback: function(value) {
                            return formatCurrencyShort(value);
                        },
                        font: {
                            size: 11
                        }
                    },
                    title: {
                        display: true,
                        text: 'GPV',
                        font: {
                            size: 13,
                            weight: '500'
                        }
                    }
                }
            }
        }
    });
    }

    // Deal Size Distribution Chart - STACKED BY OWNER
    const dealSizes = chartData.map(row => ({
        size: parseSheetNumber(row[9]),
        owner: canonicalizeOwner(row[0] || 'Unknown')
    })).filter(item => item.size > 0);
    
    const ranges = [
        { label: '$100K-$300K', min: 100000, max: 300000 },
        { label: '$300K-$500K', min: 300000, max: 500000 },
        { label: '$500K-$700K', min: 500000, max: 700000 },
        { label: '$700K-$1M', min: 700000, max: 1000000 },
        { label: '$1M-$3M', min: 1000000, max: 3000000 },
        { label: '$3M-$5M', min: 3000000, max: 5000000 },
        { label: '$5M-$8M', min: 5000000, max: 8000000 },
        { label: '$8M-$10M', min: 8000000, max: 10000000 },
        { label: '$10M+', min: 10000000, max: Infinity }
    ];
    
    // Create stacked datasets for each owner
    const dealSizeDatasets = ALLOWED_OWNERS.map((owner, index) => {
        const ownerData = ranges.map(range => {
            return dealSizes.filter(item => 
                item.owner === owner && 
                item.size >= range.min && 
                item.size < range.max
            ).length;
        });
        
        return {
            label: owner,
            data: ownerData,
            backgroundColor: OWNER_COLOR_PALETTE[index % OWNER_COLOR_PALETTE.length].fill,
            borderColor: OWNER_COLOR_PALETTE[index % OWNER_COLOR_PALETTE.length].border,
            borderWidth: 2,
            borderRadius: 8,
            stack: 'deals'
        };
    });
    
    // Calculate total deals per range for tooltips
    const rangeCounts = ranges.map(range => 
        dealSizes.filter(item => item.size >= range.min && item.size < range.max).length
    );
    
    const paymentMethodsCanvas = document.getElementById('paymentMethodsChart');
    if (paymentMethodsCanvas) {
        const paymentMethodsCtx = paymentMethodsCanvas.getContext('2d');
        paymentMethodsChart = new Chart(paymentMethodsCtx, {
        type: 'bar',
        data: {
            labels: ranges.map(r => r.label),
            datasets: dealSizeDatasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 200,
            animation: {
                duration: 800,
                easing: 'easeInOutQuart'
            },
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: { 
                    display: true,
                    position: 'top',
                    labels: {
                        font: {
                            size: 12,
                            weight: '500'
                        },
                        padding: 15,
                        usePointStyle: true,
                        pointStyle: 'circle'
                    }
                },
                title: { 
                    display: true, 
                    text: 'Deal Size Distribution by Owner',
                    font: {
                        size: 16,
                        weight: '600'
                    },
                    padding: 20
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    backgroundColor: 'rgba(255, 255, 255, 0.95)',
                    titleColor: '#000000',
                    bodyColor: '#000000',
                    borderColor: 'rgba(43, 75, 255, 0.5)',
                    borderWidth: 1,
                    padding: 12,
                    displayColors: true,
                    callbacks: {
                        label: function(context) {
                            const owner = context.dataset.label;
                            const dealCount = context.parsed.y;
                            const totalDeals = rangeCounts[context.dataIndex];
                            const percentage = totalDeals > 0 ? ((dealCount / totalDeals) * 100).toFixed(1) : 0;
                            return `${owner}: ${dealCount} deals (${percentage}%)`;
                        },
                        footer: function(context) {
                            const totalDeals = rangeCounts[context[0].dataIndex];
                            const rangeLabel = ranges[context[0].dataIndex].label;
                            return `${rangeLabel}: ${totalDeals} total deals`;
                        }
                    }
                },
                datalabels: {
                    display: function(ctx) {
                        const chart = ctx.chart;
                        for (let i = chart.data.datasets.length - 1; i >= 0; i--) {
                            const meta = chart.getDatasetMeta(i);
                            if (!meta.hidden) return ctx.datasetIndex === i;
                        }
                        return false;
                    },
                    anchor: 'end',
                    align: 'top',
                    color: '#111111',
                    font: { weight: '600', size: 11 },
                    clip: false,
                    formatter: function(value, ctx) {
                        return rangeCounts[ctx.dataIndex];
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: {
                        display: false
                    },
                    title: {
                        display: true,
                        text: 'Deal Size Range',
                        font: {
                            size: 13,
                            weight: '500'
                        }
                    },
                    ticks: {
                        font: {
                            size: 11
                        }
                    }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    grid: {
                        display: true,
                        color: 'rgba(0, 0, 0, 0.05)'
                    },
                    ticks: {
                        stepSize: 1,
                        font: {
                            size: 11
                        }
                    },
                    title: {
                        display: true,
                        text: 'Number of Deals',
                        font: {
                            size: 13,
                            weight: '500'
                        }
                    }
                }
            }
        }
    });
    }

    // Monthly Deal Count Chart - STACKED BY OWNER
    const monthlyDealCounts = monthLabels.map(month => dealsByMonth[month].length);
    
    // Prepare datasets for each owner (deal count)
    const ownerDealCountDatasets = ALLOWED_OWNERS.map((owner, index) => {
        const ownerData = monthLabels.map(month => ownerMonthlyDeals[month]?.[owner] || 0);
        return {
            label: owner,
            data: ownerData,
            backgroundColor: OWNER_COLOR_PALETTE[index % OWNER_COLOR_PALETTE.length].fill,
            borderColor: OWNER_COLOR_PALETTE[index % OWNER_COLOR_PALETTE.length].border,
            borderWidth: 2,
            borderRadius: 8,
            stack: 'deals'
        };
    });
    
    const successRateCanvas = document.getElementById('successRateChart');
    if (successRateCanvas) {
        const successRateCtx = successRateCanvas.getContext('2d');
        successRateChart = new Chart(successRateCtx, {
        type: 'bar',
        data: {
            labels: monthLabels.map(m => {
                const [year, month] = m.split('-');
                return new Date(year, month - 1).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
            }),
            datasets: ownerDealCountDatasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 200,
            animation: {
                duration: 800,
                easing: 'easeInOutQuart'
            },
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: { 
                    display: true,
                    position: 'top',
                    labels: {
                        font: {
                            size: 12,
                            weight: '500'
                        },
                        padding: 15,
                        usePointStyle: true,
                        pointStyle: 'circle'
                    }
                },
                title: { 
                    display: true, 
                    text: 'Monthly Deal Count by Owner',
                    font: {
                        size: 16,
                        weight: '600'
                    },
                    padding: 20
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    backgroundColor: 'rgba(255, 255, 255, 0.95)',
                    titleColor: '#000000',
                    bodyColor: '#000000',
                    borderColor: 'rgba(43, 75, 255, 0.5)',
                    borderWidth: 1,
                    padding: 12,
                    displayColors: true,
                    callbacks: {
                        label: function(context) {
                            const owner = context.dataset.label;
                            const dealCount = context.parsed.y;
                            const totalDeals = monthlyDealCounts[context.dataIndex];
                            const percentage = totalDeals > 0 ? ((dealCount / totalDeals) * 100).toFixed(1) : 0;
                            return `${owner}: ${dealCount} deals (${percentage}%)`;
                        },
                        footer: function(context) {
                            const totalDeals = monthlyDealCounts[context[0].dataIndex];
                            return `Total: ${totalDeals} deals`;
                        }
                    }
                },
                datalabels: {
                    display: function(ctx) {
                        const chart = ctx.chart;
                        for (let i = chart.data.datasets.length - 1; i >= 0; i--) {
                            const meta = chart.getDatasetMeta(i);
                            if (!meta.hidden) return ctx.datasetIndex === i;
                        }
                        return false;
                    },
                    anchor: 'end',
                    align: 'top',
                    color: '#111111',
                    font: { weight: '600', size: 11 },
                    clip: false,
                    formatter: function(value, ctx) {
                        return monthlyDealCounts[ctx.dataIndex];
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: {
                        display: false
                    },
                    title: {
                        display: true,
                        text: 'Month',
                        font: {
                            size: 13,
                            weight: '500'
                        }
                    },
                    ticks: {
                        font: {
                            size: 11
                        }
                    }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    grid: {
                        display: true,
                        color: 'rgba(0, 0, 0, 0.05)'
                    },
                    ticks: {
                        stepSize: 1,
                        font: {
                            size: 11
                        }
                    },
                    title: {
                        display: true,
                        text: 'Number of Deals',
                        font: {
                            size: 13,
                            weight: '500'
                        }
                    }
                }
            }
        }
    });
    }

    // Average Deal Size by Month Chart - INDIVIDUAL BARS BY OWNER
    // Calculate average deal size per owner per month
    monthLabels.forEach(month => {
        ownerMonthlyAvgDealSize[month] = {};
        ALLOWED_OWNERS.forEach(owner => {
            const ownerGPV = ownerMonthlyGPV[month]?.[owner] || 0;
            const ownerDeals = ownerMonthlyDeals[month]?.[owner] || 0;
            ownerMonthlyAvgDealSize[month][owner] = ownerDeals > 0 ? ownerGPV / ownerDeals : null;
        });
    });
    
    // Prepare datasets for each owner (average deal size)
    const ownerAvgDealSizeDatasets = ALLOWED_OWNERS.map((owner, index) => {
        const ownerData = monthLabels.map(month => ownerMonthlyAvgDealSize[month]?.[owner]);
        return {
            label: owner,
            data: ownerData,
            backgroundColor: OWNER_COLOR_PALETTE[index % OWNER_COLOR_PALETTE.length].fill,
            borderColor: OWNER_COLOR_PALETTE[index % OWNER_COLOR_PALETTE.length].border,
            borderWidth: 1.5,
            borderRadius: 6,
            barPercentage: 0.9,
            categoryPercentage: 0.85
        };
    });
    
    // Add team average line
    ownerAvgDealSizeDatasets.push({
        label: 'Team Average',
        data: monthlyAvgDealSize,
        type: 'line',
        borderColor: 'rgba(0, 0, 0, 0.9)',
        backgroundColor: 'rgba(0, 0, 0, 0.08)',
        borderWidth: 3,
        pointBackgroundColor: 'rgba(0, 0, 0, 0.9)',
        pointBorderColor: '#ffffff',
        pointBorderWidth: 2,
        pointRadius: 5,
        pointHoverRadius: 7,
        fill: false,
        tension: 0.3,
        borderDash: [6, 3],
        order: 10,
        datalabels: {
            display: true,
            align: 'top',
            anchor: 'end',
            color: '#111111',
            font: { weight: '600', size: 11 },
            clip: false,
            formatter: function(value) { return formatCurrencyShort(value); }
        }
    });
    
    const revenueByRegionCanvas = document.getElementById('revenueByRegionChart');
    if (revenueByRegionCanvas) {
        const revenueByRegionCtx = revenueByRegionCanvas.getContext('2d');
        revenueByRegionChart = new Chart(revenueByRegionCtx, {
        type: 'bar',
        data: {
            labels: monthLabels.map(m => {
                const [year, month] = m.split('-');
                return new Date(year, month - 1).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
            }),
            datasets: ownerAvgDealSizeDatasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 200,
            animation: {
                duration: 800,
                easing: 'easeInOutQuart'
            },
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: { 
                    display: true,
                    position: 'top',
                    labels: {
                        font: {
                            size: 12,
                            weight: '500'
                        },
                        padding: 15,
                        usePointStyle: true,
                        pointStyle: 'circle'
                    }
                },
                title: { 
                    display: true, 
                    text: 'Average Deal Size by Month & Owner',
                    font: {
                        size: 16,
                        weight: '600'
                    },
                    padding: 20
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    backgroundColor: 'rgba(255, 255, 255, 0.95)',
                    titleColor: '#000000',
                    bodyColor: '#000000',
                    borderColor: 'rgba(43, 75, 255, 0.5)',
                    borderWidth: 1,
                    padding: 12,
                    displayColors: true,
                    callbacks: {
                        label: function(context) {
                            const datasetLabel = context.dataset.label;
                            const avgSize = context.parsed.y;
                            
                            if (datasetLabel === 'Team Average') {
                                return `Team Average: $${avgSize.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 0})}`;
                            } else {
                                const owner = datasetLabel;
                                if (avgSize !== null && avgSize !== undefined) {
                                    const month = monthLabels[context.dataIndex];
                                    const dealCount = ownerMonthlyDeals[month]?.[owner] || 0;
                                    return `${owner}: $${avgSize.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 0})} (${dealCount} deals)`;
                                }
                                return `${owner}: No data`;
                            }
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { 
                        display: true,
                        color: 'rgba(0, 0, 0, 0.05)'
                    },
                    ticks: {
                        callback: function(value) {
                            return formatCurrencyShort(value);
                        },
                        font: {
                            size: 11
                        }
                    },
                    title: {
                        display: true,
                        text: 'Average Deal Size',
                        font: {
                            size: 13,
                            weight: '500'
                        }
                    }
                },
                x: { 
                    grid: { display: false },
                    title: {
                        display: true,
                        text: 'Month',
                        font: {
                            size: 13,
                            weight: '500'
                        }
                    },
                    ticks: {
                        font: {
                            size: 11
                        }
                    }
                }
            }
        }
    });
    }

    // Performance by Opportunity Owner Chart
    const ownerPerformance = {};
    chartData.forEach(row => {
        const owner = canonicalizeOwner(row[0] || 'Unknown'); // Column A
        const gpv = parseSheetNumber(row[9]); // Column J
        if (!ownerPerformance[owner]) {
            ownerPerformance[owner] = { gpv: 0, deals: 0 };
        }
        ownerPerformance[owner].gpv += gpv;
        ownerPerformance[owner].deals += 1;
    });

    // Sort owners by total GPV descending for left-to-right highest ordering
    const sortedOwners = Object.keys(ownerPerformance)
        .sort((a, b) => ownerPerformance[b].gpv - ownerPerformance[a].gpv);
    const ownerLabels = sortedOwners;
    const ownerGPV = sortedOwners.map(owner => ownerPerformance[owner].gpv);
    const ownerDeals = sortedOwners.map(owner => ownerPerformance[owner].deals);

    const ownerPerformanceCanvas = document.getElementById('ownerPerformanceChart');
    if (ownerPerformanceCanvas) {
        const ownerPerformanceCtx = ownerPerformanceCanvas.getContext('2d');
        ownerPerformanceChart = new Chart(ownerPerformanceCtx, {
        type: 'bar',
        data: {
            labels: ownerLabels,
            datasets: [{
                label: 'Total GPV',
                data: ownerGPV,
                backgroundColor: 'rgba(34, 197, 94, 0.8)',
                borderColor: 'rgba(34, 197, 94, 1)',
                borderWidth: 1,
                yAxisID: 'y'
            }, {
                label: 'Deal Count',
                data: ownerDeals,
                backgroundColor: 'rgba(59, 130, 246, 0.8)',
                borderColor: 'rgba(59, 130, 246, 1)',
                borderWidth: 1,
                yAxisID: 'y1'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 200,
            plugins: {
                legend: { display: true },
                title: { display: true, text: 'Performance by Opportunity Owner' }
            },
            scales: {
                y: {
                    type: 'linear',
                    display: true,
                    position: 'left',
                    ticks: {
                        callback: function(value) {
                            return '$' + value.toLocaleString();
                        }
                    }
                },
                y1: {
                    type: 'linear',
                    display: true,
                    position: 'right',
                    grid: { drawOnChartArea: false }
                }
            }
        }
    });
    }

    // Win Rate by Opportunity Owner Chart (Rolling 6 Months)
    const winRates = {};
    // Calculate date range for last 6 months
    const now = new Date();
    const wrStartDate = new Date(now.getFullYear(), now.getMonth() - 6, now.getDate());
    const wrEndDate = now; // Up to current date
    data.forEach(row => {
        const owner = String(row[0] || '').trim(); // Column A
        const ownerLC = owner.toLowerCase();
        if (!ALLOWED_OWNERS.some(o => o.toLowerCase() === ownerLC)) return; // user filter
        const dealDate = parseSheetDate(row[7]); // Column H
        if (!dealDate) return;
        const afterStart = dealDate >= wrStartDate;
        const beforeEnd = dealDate <= wrEndDate; // Include up to current date
        if (!afterStart || !beforeEnd) return; // date filter for last 6 months

        const stage = String(row[4]).toLowerCase().trim(); // Column E
        if (!winRates[owner]) {
            winRates[owner] = { total: 0, won: 0 };
        }
        winRates[owner].total += 1;
        if (stage === 'closed won') {
            winRates[owner].won += 1;
        }
    });

    const winRateLabels = Object.keys(winRates).filter(owner => winRates[owner].total >= 3); // Only show owners with 3+ deals
    const winRateData = winRateLabels.map(owner => {
        const rate = (winRates[owner].won / winRates[owner].total) * 100;
        return rate;
    });

    const winRateCanvas = document.getElementById('winRateChart');
    if (winRateCanvas) {
        const winRateCtx = winRateCanvas.getContext('2d');
        winRateChart = new Chart(winRateCtx, {
        type: 'bar',
        data: {
            labels: winRateLabels,
            datasets: [{
                label: 'Win Rate (%)',
                data: winRateData,
                backgroundColor: 'rgba(251, 191, 36, 0.8)',
                borderColor: 'rgba(251, 191, 36, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 200,
            plugins: {
                legend: { display: false },
                title: { display: true, text: 'Win Rate by Opportunity Owner (Rolling 6 Months)' },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const owner = context.label;
                            const stats = winRates[owner];
                            return `${context.raw.toFixed(1)}% (${stats.won}/${stats.total} deals)`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 100,
                    ticks: { callback: value => value + '%' }
                }
            }
        }
    });
    }

    // Opportunity Sub-Source Analysis Chart - WITH INDIVIDUAL BREAKDOWN
    const sourceStats = {};
    const sourceOwnerStats = {};
    data.forEach(row => {
        const subSource = row[12] || 'Implementation'; // Column M - Opportunity Sub-Source, default to Implementation if blank
        const stage = String(row[4]).toLowerCase().trim(); // Column E
        const owner = canonicalizeOwner(row[0] || 'Unknown'); // Column A
        
        if (!sourceStats[subSource]) {
            sourceStats[subSource] = { total: 0, won: 0 };
            sourceOwnerStats[subSource] = {};
        }
        sourceStats[subSource].total += 1;
        if (stage === 'closed won') {
            sourceStats[subSource].won += 1;
        }
        
        // Track per-owner stats for this source
        if (!sourceOwnerStats[subSource][owner]) {
            sourceOwnerStats[subSource][owner] = { total: 0, won: 0 };
        }
        sourceOwnerStats[subSource][owner].total += 1;
        if (stage === 'closed won') {
            sourceOwnerStats[subSource][owner].won += 1;
        }
    });

    const sourceLabels = Object.keys(sourceStats).sort();
    const sourceData = sourceLabels.map(source => sourceStats[source].total);
    const sourceTotal = sourceData.reduce((a, b) => a + b, 0);

    const sourceCanvas = document.getElementById('sourceChart');
    if (sourceCanvas) {
        const sourceCtx = sourceCanvas.getContext('2d');
        sourceChart = new Chart(sourceCtx, {
        type: 'doughnut',
        data: {
            labels: sourceLabels,
            datasets: [{
                data: sourceData,
                backgroundColor: [
                    'rgba(59, 130, 246, 0.8)',
                    'rgba(16, 185, 129, 0.8)',
                    'rgba(245, 158, 11, 0.8)',
                    'rgba(239, 68, 68, 0.8)',
                    'rgba(168, 85, 247, 0.8)',
                    'rgba(236, 72, 153, 0.8)',
                    'rgba(34, 197, 94, 0.8)',
                    'rgba(251, 191, 36, 0.8)'
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 200,
            plugins: {
                legend: { position: 'right' },
                title: { display: true, text: 'Opportunity Sub-Source Distribution' },
                datalabels: {
                    display: function(ctx) {
                        const total = sourceTotal;
                        const v = ctx.dataset.data[ctx.dataIndex] || 0;
                        return total > 0 && v / total >= 0.06;
                    },
                    color: '#111111',
                    font: { weight: '600', size: 11 },
                    formatter: function(value) {
                        if (!sourceTotal) return '';
                        const pct = (value / sourceTotal) * 100;
                        return Math.round(pct) + '%';
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const sourceIndex = context.dataIndex;
                            const sourceName = sourceLabels[sourceIndex];
                            const totalOpps = context.raw;
                            const sourceStatsData = sourceStats[sourceName];
                            const ownerBreakdown = sourceOwnerStats[sourceName] || {};
                            
                            let tooltipText = `${sourceName}: ${totalOpps} opportunities`;
                            if (sourceStatsData) {
                                const winRate = sourceStatsData.total > 0 ? 
                                    ((sourceStatsData.won / sourceStatsData.total) * 100).toFixed(1) : 0;
                                tooltipText += ` (${sourceStatsData.won} won, ${winRate}% win rate)`;
                            }
                            
                            // Add individual owner contributions
                            const sortedOwners = Object.entries(ownerBreakdown)
                                .sort(([,a], [,b]) => b.total - a.total)
                                .filter(([,ownerStats]) => ownerStats.total > 0);
                            
                            if (sortedOwners.length > 0) {
                                tooltipText += '\n\nBy Owner:';
                                sortedOwners.forEach(([owner, ownerStats]) => {
                                    const ownerWinRate = ownerStats.total > 0 ? 
                                        ((ownerStats.won / ownerStats.total) * 100).toFixed(1) : 0;
                                    tooltipText += `\n${owner}: ${ownerStats.total} opps (${ownerStats.won} won, ${ownerWinRate}%)`;
                                });
                            }
                            
                            return tooltipText;
                        }
                    }
                }
            },
            cutout: '60%'
        }
    });
    }

    // Deal Age by Segment Chart - INDIVIDUAL OWNER BARS
    const segmentAgeStats = {};
    const segmentOwnerStats = {};
    chartData.forEach(row => {
        const segment = row[18] || 'Unknown'; // Column S
        const age = parseInt(row[6]) || 0; // Column G - deal age in days
        const owner = canonicalizeOwner(row[0] || 'Unknown'); // Column A
        
        if (!segmentAgeStats[segment]) {
            segmentAgeStats[segment] = { totalAge: 0, count: 0 };
            segmentOwnerStats[segment] = {};
        }
        segmentAgeStats[segment].totalAge += age;
        segmentAgeStats[segment].count += 1;
        
        // Track per-owner stats for this segment
        if (!segmentOwnerStats[segment][owner]) {
            segmentOwnerStats[segment][owner] = { totalAge: 0, count: 0 };
        }
        segmentOwnerStats[segment][owner].totalAge += age;
        segmentOwnerStats[segment][owner].count += 1;
    });

    const ageSegmentLabels = Object.keys(segmentAgeStats).sort();
    const avgAges = ageSegmentLabels.map(segment => {
        const stats = segmentAgeStats[segment];
        return stats.count > 0 ? Math.round(stats.totalAge / stats.count) : 0;
    });

    // Create datasets for each owner showing their average age per segment
    const ageDatasets = ALLOWED_OWNERS.map((owner, index) => {
        const ownerData = ageSegmentLabels.map(segment => {
            const ownerStats = segmentOwnerStats[segment]?.[owner];
            return ownerStats && ownerStats.count > 0 ? Math.round(ownerStats.totalAge / ownerStats.count) : null;
        });
        
        return {
            label: owner,
            data: ownerData,
            backgroundColor: OWNER_COLOR_PALETTE[index % OWNER_COLOR_PALETTE.length].fill,
            borderColor: OWNER_COLOR_PALETTE[index % OWNER_COLOR_PALETTE.length].border,
            borderWidth: 1.5,
            borderRadius: 6,
            barThickness: 20
        };
    });

    // Add overall average line (put at the end so it renders on top)
    ageDatasets.push({
        label: 'Team Average',
        data: avgAges,
        type: 'line',
        borderColor: 'rgba(0, 0, 0, 0.9)',
        backgroundColor: 'rgba(0, 0, 0, 0.08)',
        borderWidth: 3,
        pointBackgroundColor: 'rgba(0, 0, 0, 0.9)',
        pointBorderColor: '#ffffff',
        pointBorderWidth: 2,
        pointRadius: 5,
        pointHoverRadius: 7,
        fill: false,
        tension: 0.3,
        borderDash: [6, 3],
        order: 10,
        datalabels: {
            display: true,
            align: 'top',
            anchor: 'end',
            color: '#111111',
            font: { weight: '600', size: 11 },
            clip: false,
            formatter: function(value) { return value + 'd'; }
        }
    });

    const ageCanvas = document.getElementById('ageChart');
    if (ageCanvas) {
        const ageCtx = ageCanvas.getContext('2d');
        ageChart = new Chart(ageCtx, {
        type: 'bar',
        data: {
            labels: ageSegmentLabels,
            datasets: ageDatasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 200,
            animation: {
                duration: 800,
                easing: 'easeInOutQuart'
            },
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: { 
                    display: true,
                    position: 'top',
                    labels: {
                        font: {
                            size: 12,
                            weight: '500'
                        },
                        padding: 15,
                        usePointStyle: true,
                        pointStyle: 'circle'
                    }
                },
                title: { 
                    display: true, 
                    text: 'Average Deal Age by Segment & Owner',
                    font: {
                        size: 16,
                        weight: '600'
                    },
                    padding: 20
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    backgroundColor: 'rgba(255, 255, 255, 0.95)',
                    titleColor: '#000000',
                    bodyColor: '#000000',
                    borderColor: 'rgba(43, 75, 255, 0.5)',
                    borderWidth: 1,
                    padding: 12,
                    displayColors: true,
                    callbacks: {
                        label: function(context) {
                            const segment = context.label;
                            const datasetLabel = context.dataset.label;
                            
                            if (datasetLabel === 'Team Average') {
                                const avgAge = context.parsed.y;
                                return `Team Average: ${avgAge} days`;
                            } else {
                                const owner = datasetLabel;
                                const avgAge = context.parsed.y;
                                if (avgAge !== null && avgAge !== undefined) {
                                    const ownerStats = segmentOwnerStats[segment]?.[owner];
                                    return `${owner}: ${avgAge} days (${ownerStats?.count || 0} deals)`;
                                }
                                return `${owner}: No data`;
                            }
                        }
                    }
                }
            },
            scales: {
                y: { 
                    beginAtZero: true,
                    grid: {
                        display: true,
                        color: 'rgba(0, 0, 0, 0.05)'
                    },
                    title: { 
                        display: true, 
                        text: 'Average Days',
                        font: {
                            size: 13,
                            weight: '500'
                        }
                    },
                    ticks: { 
                        stepSize: 10,
                        font: {
                            size: 11
                        }
                    }
                },
                x: {
                    grid: {
                        display: false
                    },
                    title: { 
                        display: true, 
                        text: 'Customer Segment',
                        font: {
                            size: 13,
                            weight: '500'
                        }
                    },
                    ticks: {
                        font: {
                            size: 11
                        }
                    }
                }
            }
        }
    });
    }

    // Current Payment Processors Chart
    const processorStats = {};
    data.forEach(row => {
        const processor = row[16] || 'Unknown'; // Column Q
        if (!processorStats[processor]) {
            processorStats[processor] = 0;
        }
        processorStats[processor] += 1;
    });

    // Sort by count descending and take top 5
    const sortedProcessors = Object.entries(processorStats)
        .sort(([,a], [,b]) => b - a)
        .slice(0, 5);

    const processorLabels = sortedProcessors.map(([processor]) => processor);
    const processorData = sortedProcessors.map(([,count]) => count);

    // Create top 5 text for subtitle
    const top5Text = sortedProcessors.length >= 5
        ? `Top 5: ${processorLabels.join(', ')}`
        : `Top ${sortedProcessors.length}: ${processorLabels.join(', ')}`;

    const processorCanvas = document.getElementById('processorChart');
    if (processorCanvas) {
        const processorCtx = processorCanvas.getContext('2d');
        processorChart = new Chart(processorCtx, {
        type: 'doughnut',
        data: {
            labels: processorLabels,
            datasets: [{
                data: processorData,
                backgroundColor: [
                    'rgba(99, 102, 241, 0.8)',
                    'rgba(168, 85, 247, 0.8)',
                    'rgba(236, 72, 153, 0.8)',
                    'rgba(251, 146, 60, 0.8)',
                    'rgba(34, 197, 94, 0.8)'
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 200,
            plugins: {
                legend: { position: 'right' },
                title: {
                    display: true,
                    text: 'Current Payment Processors',
                    font: {
                        size: 16,
                        weight: '600'
                    },
                    padding: 20
                },
                subtitle: {
                    display: true,
                    text: top5Text,
                    font: {
                        size: 12,
                        weight: '400'
                    },
                    padding: {
                        top: 0,
                        bottom: 10
                    }
                },
                datalabels: {
                    display: function(ctx) {
                        const total = processorData.reduce((a, b) => a + b, 0);
                        const v = ctx.dataset.data[ctx.dataIndex] || 0;
                        return total > 0 && v / total >= 0.08;
                    },
                    color: '#111111',
                    font: { weight: '600', size: 11 },
                    formatter: function(value) {
                        const total = processorData.reduce((a, b) => a + b, 0);
                        if (!total) return '';
                        const pct = (value / total) * 100;
                        return Math.round(pct) + '%';
                    }
                }
            },
            cutout: '60%'
        }
    });
    }

    // GPV Segment Distribution Chart
    const segmentStats = {};
    const segmentGpvTotals = {};
    data.forEach(row => {
        const segment = row[18] || 'Unknown'; // Column S
        if (!segmentStats[segment]) {
            segmentStats[segment] = 0;
            segmentGpvTotals[segment] = 0;
        }
        segmentStats[segment] += 1;
        segmentGpvTotals[segment] += parseSheetNumber(row[9] || 0); // Column J - GPV
    });

    const segmentLabels = Object.keys(segmentStats).sort();
    const segmentData = segmentLabels.map(seg => segmentStats[seg]);
    const segmentTotal = segmentData.reduce((a, b) => a + b, 0);

    const segmentCanvas = document.getElementById('segmentChart');
    if (segmentCanvas) {
        const segmentCtx = segmentCanvas.getContext('2d');
        segmentChart = new Chart(segmentCtx, {
            type: 'pie',
            data: {
                labels: segmentLabels,
                datasets: [{
                    data: segmentData,
                    backgroundColor: [
                        'rgba(59, 130, 246, 0.8)',
                        'rgba(16, 185, 129, 0.8)',
                        'rgba(251, 146, 60, 0.8)',
                        'rgba(239, 68, 68, 0.8)',
                        'rgba(168, 85, 247, 0.8)',
                        'rgba(236, 72, 153, 0.8)'
                    ]
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                resizeDelay: 200,
                plugins: {
                    legend: { position: 'right' },
                    title: { display: true, text: 'GPV Segment Distribution' },
                    datalabels: {
                        display: function(ctx) {
                            const total = segmentTotal;
                            const v = ctx.dataset.data[ctx.dataIndex] || 0;
                            return total > 0 && v / total >= 0.06;
                        },
                        color: '#111111',
                        font: { weight: '600', size: 11 },
                        formatter: function(value) {
                            if (!segmentTotal) return '';
                            const pct = (value / segmentTotal) * 100;
                            return Math.round(pct) + '%';
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const label = context.label || '';
                                const count = context.raw || 0;
                                const gpv = segmentGpvTotals[label] || 0;
                                return `${label}: ${count} opps, $${gpv.toLocaleString()}`;
                            }
                        }
                    }
                }
            }
        });
    }

    // Initialize the owner source chart
    updateOwnerSourceChart(data);
}

// Render Activities chart for Account Managers - Payments
function renderActivitiesChart(activitiesData) {
    console.log('Rendering activities chart with data:', activitiesData);

    // Destroy existing chart
    if (activitiesChart) {
        activitiesChart.destroy();
        activitiesChart = null;
    }

    // Calculate activities metrics
    const metrics = calculateActivitiesMetrics(activitiesData);

    // Helper: format minutes as seconds if < 60s, else m/s
    const formatDurationMinutes = (mins) => {
        const n = Number(mins);
        if (!isFinite(n) || n <= 0) return '0s';
        const totalSeconds = Math.round(n * 60);
        if (totalSeconds < 60) return `${totalSeconds}s`;
        const m = Math.floor(totalSeconds / 60);
        const s = totalSeconds % 60;
        return s === 0 ? `${m}m` : `${m}m ${s}s`;
    };

    // Create a comprehensive chart showing multiple metrics
    const activitiesCanvas = document.getElementById('activitiesChart');
    if (activitiesCanvas) {
        const activitiesCtx = activitiesCanvas.getContext('2d');

        // Prepare data for mixed chart
        const labels = ['Calls', 'Emails', 'Demo Sets', 'Avg Call Duration', 'Calls/Demo', 'Emails/Demo', 'Answered Calls %'];
        const dataValues = [
            metrics.totalCalls,
            metrics.totalEmails,
            metrics.demoSets,
            metrics.avgCallDurationForDemos,
            metrics.callsPerDemo,
            metrics.emailsPerDemo,
            metrics.callsPerAnswered
        ];

        // Use different chart types for different data types
        const colors = [
            'rgba(59, 130, 246, 0.8)',  // Blue - Calls
            'rgba(16, 185, 129, 0.8)',  // Emerald - Emails
            'rgba(245, 158, 11, 0.8)',  // Amber - Demo Sets
            'rgba(239, 68, 68, 0.8)',   // Red - Avg Duration
            'rgba(168, 85, 247, 0.8)',  // Violet - Calls/Demo
            'rgba(236, 72, 153, 0.8)',  // Pink - Emails/Demo
            'rgba(34, 197, 94, 0.8)'    // Green - Answered %
        ];

        activitiesChart = new Chart(activitiesCtx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: 'Activities Metrics',
                    data: dataValues,
                    backgroundColor: colors,
                    borderColor: colors.map(color => color.replace('0.8', '1')),
                    borderWidth: 1,
                    borderRadius: 6,
                    barThickness: 40
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                resizeDelay: 200,
                plugins: {
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'Payments AE Activity',
                        font: {
                            size: 16,
                            weight: '600'
                        },
                        padding: 20
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const value = context.raw;
                                const label = context.label;

                                if (label === 'Avg Call Duration') {
                                    return `${label}: ${formatDurationMinutes(value)}`;
                                } else if (label === 'Calls/Demo') {
                                    return `${label}: ${value.toFixed(1)} calls per demo`;
                                } else if (label === 'Emails/Demo') {
                                    return `${label}: ${value.toFixed(1)} emails per demo`;
                                } else if (label === 'Answered Calls %') {
                                    return `${label}: ${value.toFixed(1)}%`;
                                } else {
                                    return `${label}: ${value.toLocaleString()}`;
                                }
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: {
                            display: true,
                            color: 'rgba(0, 0, 0, 0.05)'
                        },
                        ticks: {
                            font: {
                                size: 11
                            }
                        },
                        title: {
                            display: true,
                            text: 'Value',
                            font: {
                                size: 13,
                                weight: '500'
                            }
                        }
                    },
                    x: {
                        grid: {
                            display: false
                        },
                        ticks: {
                            font: {
                                size: 11
                            },
                            maxRotation: 45,
                            minRotation: 45
                        }
                    }
                }
            }
        });
    }

    // Update activities metrics display
    const activitiesMetricsDiv = document.getElementById('activities-metrics');
    if (activitiesMetricsDiv) {
        const formatDate = (date) => {
            if (!date) return '';
            const y = date.getUTCFullYear();
            const m = date.getUTCMonth();
            return new Date(Date.UTC(y, m, 1)).toLocaleDateString('en-US', { month: 'short', year: 'numeric', timeZone: 'UTC' });
        };
        const startDisplay = activitiesDateFilter.startDate ? formatDate(activitiesDateFilter.startDate) : 'All Time';
        const endDisplay = activitiesDateFilter.endDate ? formatDate(activitiesDateFilter.endDate) : 'present';

        activitiesMetricsDiv.innerHTML = `
            <div class="bg-white p-4 rounded-lg shadow border">
                <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Total Activities (${startDisplay} - ${endDisplay})</h3>
                <p class="text-2xl font-bold mt-2 mb-1" style="background: linear-gradient(135deg, #2B4BFF 0%, #0BE6C7 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">${(metrics.totalCalls + metrics.totalEmails).toLocaleString()}</p>
                <p class="text-sm text-gray-600">${metrics.totalCalls} calls, ${metrics.totalEmails} emails</p>
            </div>
            <div class="bg-white p-4 rounded-lg shadow border">
                <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Demo Sets</h3>
                <p class="text-2xl font-bold mt-2 mb-1" style="background: linear-gradient(135deg, #0BE6C7 0%, #2B4BFF 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">${metrics.demoSets}</p>
                <p class="text-sm text-gray-600">Successful demo bookings</p>
            </div>
            <div class="bg-white p-4 rounded-lg shadow border">
                <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Avg Call Duration</h3>
                <p class="text-2xl font-bold mt-2 mb-1" style="background: linear-gradient(135deg, #2B4BFF 0%, #000000 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">${formatDurationMinutes(metrics.avgCallDurationForDemos)}</p>
                <p class="text-sm text-gray-600">For demo-set calls</p>
            </div>
        `;

        // Append Top Demo Set Performers card
        const topPerformers = (metrics.perUser || [])
            .slice()
            .sort((a, b) => (b.demoSets || 0) - (a.demoSets || 0))
            .filter(p => (p.demoSets || 0) > 0)
            .slice(0, 3);
        const performersHtml = topPerformers.length
            ? topPerformers.map(p => `<li class="flex items-center justify-between"><span>${canonicalizeOwner(p.user)}</span><span class="font-semibold">${p.demoSets}</span></li>`).join('')
            : '<li class="text-gray-500">No demo sets in range</li>';
        activitiesMetricsDiv.insertAdjacentHTML('beforeend', `
            <div class="bg-white p-4 rounded-lg shadow border">
                <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Top Demo Set Performers</h3>
                <ul class="text-sm text-gray-700 mt-2 space-y-1">${performersHtml}</ul>
            </div>
        `);
    }

    // Populate AE breakdown table
    const breakdownBody = document.getElementById('activitiesBreakdownTableBody');
    if (breakdownBody) {
        breakdownBody.innerHTML = '';
        const rows = (metrics && Array.isArray(metrics.perUser)) ? metrics.perUser : [];
        if (rows.length === 0) {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td colspan="8" class="px-6 py-4 text-center text-sm text-gray-500">No activity data for the selected range</td>`;
            breakdownBody.appendChild(tr);
        } else {
            rows.forEach((r, index) => {
                const tr = document.createElement('tr');
                tr.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
                const answered = Number.isFinite(r.callsPerAnswered) ? r.callsPerAnswered.toFixed(1) : '0.0';
                const callsPerDemo = Number.isFinite(r.callsPerDemo) ? r.callsPerDemo.toFixed(1) : '0.0';
                const emailsPerDemo = Number.isFinite(r.emailsPerDemo) ? r.emailsPerDemo.toFixed(1) : '0.0';
                const avgDemo = formatDurationMinutes(r.avgCallDurationForDemos);
                tr.innerHTML = `
                    <td class="px-6 py-3 text-sm text-gray-900">${canonicalizeOwner(r.user)}</td>
                    <td class="px-6 py-3 text-sm text-gray-900">${(r.totalCalls || 0).toLocaleString()}</td>
                    <td class="px-6 py-3 text-sm text-gray-900">${(r.totalEmails || 0).toLocaleString()}</td>
                    <td class="px-6 py-3 text-sm text-gray-900">${(r.demoSets || 0).toLocaleString()}</td>
                    <td class="px-6 py-3 text-sm text-gray-900">${answered}%</td>
                    <td class="px-6 py-3 text-sm text-gray-900">${callsPerDemo}</td>
                    <td class="px-6 py-3 text-sm text-gray-900">${emailsPerDemo}</td>
                    <td class="px-6 py-3 text-sm text-gray-900">${avgDemo}</td>
                `;
                breakdownBody.appendChild(tr);
            });
        }
    }
}

// Render projected deals to close this month
// Global variables for projected deals table sorting
let projectedDealsSortColumn = null; // null, 'name', 'owner', 'date', 'gpv', 'stage', 'days'
let projectedDealsSortDirection = 'asc'; // 'asc' or 'desc'

function renderProjectedDeals(data) {
    try {
        console.log('Rendering projected deals for current month');

        // Get current month date range
        const now = new Date();
        const currentMonthStart = new Date(now.getFullYear(), now.getMonth(), 1);
        const currentMonthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);

        // Update date display
        const dateElement = document.getElementById('projected-deals-date');
        if (dateElement) {
            const monthName = currentMonthStart.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            dateElement.textContent = `Showing opportunities expected to close in ${monthName}`;
        }

        // Filter opportunities that are not closed (not "Closed Won" or "Closed Lost") and close this month
        let projectedDeals = data.filter(row => {
            const closeDate = parseSheetDate(row[7]); // Column H - Close Date
            const stage = String(row[4] || '').toLowerCase().trim(); // Column E - Stage
            const owner = String(row[0] || '').trim(); // Column A - Owner
            const ownerLC = owner.toLowerCase();
            // Check if deal closes this month
            const closesThisMonth = closeDate >= currentMonthStart && closeDate <= currentMonthEnd;

            // Check if deal is not closed (not "closed won" or "closed lost")
            const isOpen = stage !== 'closed won' && stage !== 'closed lost';

            // Check if owner is allowed and matches user filter
            const isAllowedOwner = ALLOWED_OWNERS.some(o => o.toLowerCase() === ownerLC);
            const userFilterMatch = selectedUsers.length === 0 || selectedUsers.includes(owner);

            return closesThisMonth && isOpen && isAllowedOwner && userFilterMatch;
        });

        // Sort by GPV descending initially (if no sort column selected)
        if (!projectedDealsSortColumn) {
            projectedDeals.sort((a, b) => parseSheetNumber(b[9]) - parseSheetNumber(a[9]));
        } else {
            projectedDeals = sortProjectedDeals(projectedDeals);
        }

        // Limit to top 15
        projectedDeals = projectedDeals.slice(0, 15);

        // Calculate metrics
        const totalProjectedGPV = projectedDeals.reduce((sum, row) => sum + parseSheetNumber(row[9]), 0);
        const totalDeals = projectedDeals.length;
        const avgDealSize = totalDeals > 0 ? totalProjectedGPV / totalDeals : 0;

        // Update metrics summary
        const metricsDiv = document.getElementById('projected-deals-metrics');
        if (metricsDiv) {
            metricsDiv.innerHTML = `
                <div class="bg-white p-4 rounded-lg shadow border">
                    <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Projected GPV</h3>
                    <p class="text-2xl font-bold mt-2 mb-1" style="background: linear-gradient(135deg, #2B4BFF 0%, #0BE6C7 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">$${totalProjectedGPV.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 0})}</p>
                    <p class="text-sm text-gray-600">${totalDeals} deals projected</p>
                </div>
                <div class="bg-white p-4 rounded-lg shadow border">
                    <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Avg Deal Size</h3>
                    <p class="text-2xl font-bold mt-2 mb-1" style="background: linear-gradient(135deg, #0BE6C7 0%, #2B4BFF 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">$${avgDealSize.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 0})}</p>
                    <p class="text-sm text-gray-600">Average projected deal</p>
                </div>
                <div class="bg-white p-4 rounded-lg shadow border">
                    <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Opportunities</h3>
                    <p class="text-2xl font-bold mt-2 mb-1" style="background: linear-gradient(135deg, #2B4BFF 0%, #000000 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">${totalDeals}</p>
                    <p class="text-sm text-gray-600">Open opportunities</p>
                </div>
            `;
        }

        // Update table
        const tableBody = document.getElementById('projectedDealsTableBody');
        if (!tableBody) {
            console.warn('Projected deals table body not found');
            return;
        }

        tableBody.innerHTML = '';

        // Setup sortable headers
        setupProjectedDealsSorting();

        // Format currency
        const formatCurrency = (value) => {
            const num = parseSheetNumber(value);
            return new Intl.NumberFormat('en-US', {
                style: 'currency',
                currency: 'USD',
                minimumFractionDigits: 0,
                maximumFractionDigits: 0
            }).format(num);
        };

        // Add rows with data
        projectedDeals.forEach((row, index) => {
            const tr = document.createElement('tr');
            tr.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';

            const oppName = row[3] || 'N/A'; // Column D - Opportunity Name
            const owner = row[0] || 'N/A'; // Column A - Owner
            const closeDate = parseSheetDate(row[7]); // Column H - Close Date
            const closeDateStr = closeDate ? closeDate.toLocaleDateString() : 'N/A';
            const gpv = formatCurrency(row[9]); // Column J - GPV
            const stage = row[4] || 'N/A'; // Column E - Stage

            // Calculate days until close
            const daysUntilClose = closeDate ? Math.ceil((closeDate - now) / (1000 * 60 * 60 * 24)) : 'N/A';

            tr.innerHTML = `
                <td class="px-6 py-4 text-sm text-gray-900">${oppName}</td>
                <td class="px-6 py-4 text-sm text-gray-900">${owner}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${closeDateStr}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm font-semibold text-green-600">${gpv}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${stage}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm ${daysUntilClose === 'N/A' ? 'text-gray-500' : daysUntilClose <= 7 ? 'text-red-600 font-semibold' : daysUntilClose <= 14 ? 'text-orange-600' : 'text-green-600'}">${daysUntilClose}</td>
            `;

            tableBody.appendChild(tr);
        });

        if (projectedDeals.length === 0) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td colspan="6" class="px-6 py-4 text-center text-sm text-gray-500">
                    No open opportunities projected to close this month
                </td>
            `;
            tableBody.appendChild(tr);
        }
    } catch (error) {
        console.error('Error rendering projected deals:', error);
        // Don't throw - just log the error so other parts of the dashboard still load
    }
}

// Sort projected deals based on current sort settings
function sortProjectedDeals(deals) {
    return deals.sort((a, b) => {
        let valueA, valueB;

        switch (projectedDealsSortColumn) {
            case 'name':
                valueA = String(a[3] || '').toLowerCase(); // Column D - Opportunity Name
                valueB = String(b[3] || '').toLowerCase();
                break;
            case 'owner':
                valueA = String(a[0] || '').toLowerCase(); // Column A - Owner
                valueB = String(b[0] || '').toLowerCase();
                break;
            case 'date':
                valueA = parseSheetDate(a[7]); // Column H - Close Date
                valueB = parseSheetDate(b[7]);
                break;
            case 'gpv':
                valueA = parseSheetNumber(a[9]); // Column J - GPV
                valueB = parseSheetNumber(b[9]);
                break;
            case 'stage':
                valueA = String(a[4] || '').toLowerCase(); // Column E - Stage
                valueB = String(b[4] || '').toLowerCase();
                break;
            case 'days':
                const now = new Date();
                const dateA = parseSheetDate(a[7]);
                const dateB = parseSheetDate(b[7]);
                valueA = dateA ? Math.ceil((dateA - now) / (1000 * 60 * 60 * 24)) : Infinity;
                valueB = dateB ? Math.ceil((dateB - now) / (1000 * 60 * 60 * 24)) : Infinity;
                break;
            default:
                return 0;
        }

        // Handle null/undefined values
        if (valueA == null && valueB == null) return 0;
        if (valueA == null) return projectedDealsSortDirection === 'asc' ? -1 : 1;
        if (valueB == null) return projectedDealsSortDirection === 'asc' ? 1 : -1;

        // Compare values
        let result;
        if (valueA < valueB) result = -1;
        else if (valueA > valueB) result = 1;
        else result = 0;

        return projectedDealsSortDirection === 'asc' ? result : -result;
    });
}

// Setup sortable headers for projected deals table
function setupProjectedDealsSorting() {
    const table = document.getElementById('projectedDealsTable');
    if (!table) return;

    const headers = table.querySelectorAll('thead th');
    headers.forEach((header, index) => {
        // Skip if already has event listener
        if (header.dataset.sortable === 'true') return;

        header.dataset.sortable = 'true';
        header.style.cursor = 'pointer';
        header.style.userSelect = 'none';

        // Add sort indicator span
        const sortIndicator = document.createElement('span');
        sortIndicator.className = 'ml-2 text-xs opacity-50';
        sortIndicator.textContent = '↕️';
        header.appendChild(sortIndicator);

        header.addEventListener('click', () => {
            const columnMap = ['name', 'owner', 'date', 'gpv', 'stage', 'days'];
            const column = columnMap[index];

            if (!column) return;

            // Toggle sort direction if same column, otherwise set to ascending
            if (projectedDealsSortColumn === column) {
                projectedDealsSortDirection = projectedDealsSortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                projectedDealsSortColumn = column;
                projectedDealsSortDirection = 'asc';
            }

            // Update visual indicators
            updateProjectedDealsSortIndicators();

            // Re-render table with new sort
            if (window.currentData) {
                renderProjectedDeals(window.currentData);
            }
        });
    });

    // Initial visual indicators
    updateProjectedDealsSortIndicators();
}

// Update sort indicators for projected deals table
function updateProjectedDealsSortIndicators() {
    const table = document.getElementById('projectedDealsTable');
    if (!table) return;

    const headers = table.querySelectorAll('thead th');
    const columnMap = ['name', 'owner', 'date', 'gpv', 'stage', 'days'];

    headers.forEach((header, index) => {
        const sortIndicator = header.querySelector('span');
        if (!sortIndicator) return;

        const column = columnMap[index];
        if (projectedDealsSortColumn === column) {
            sortIndicator.textContent = projectedDealsSortDirection === 'asc' ? '↑' : '↓';
            sortIndicator.className = 'ml-2 text-xs opacity-100 font-bold';
            header.classList.add('bg-gray-100');
        } else {
            sortIndicator.textContent = '↕️';
            sortIndicator.className = 'ml-2 text-xs opacity-50';
            header.classList.remove('bg-gray-100');
        }
    });
}

// For development/testing without Google Sheets
async function fetchMockData() {
    // Simulate API delay
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    // Generate mock data
    const mockData = [];
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    
    for (let i = 0; i < 12; i++) {
        const baseAmount = 10000 + Math.random() * 5000;
        const successRate = 85 + Math.random() * 10;
        
        mockData.push({
            date: `${months[i % 12]} 2023`,
            volume: Math.round(baseAmount * (1 + i * 0.1)),
            success_rate: successRate.toFixed(2),
            avg_transaction: (50 + Math.random() * 20).toFixed(2)
        });
    }
    
    return mockData;
}
 
