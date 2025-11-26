/* global Chart, gapi, google */

const CONFIG = window.CONFIG || {
  AUTH_METHOD: 'oauth',
  SPREADSHEET_ID: '1XhNpvY1SYsvszBugJeD-gQKar9OwU4lLLiF5cTkAk6I',
  CLIENT_ID: '630950450890-lofitb7ofs2q6ae3olqv1jis88uhfjeg.apps.googleusercontent.com',
  DISCOVERY_DOCS: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
  SCOPES: 'https://www.googleapis.com/auth/spreadsheets.readonly'
};

let tokenClient; let gapiInited = false; let dataLoaded = false; let isLoading = false;
let aeData = []; let oppSplitData = [];

// Charts
let attainmentChart = null;
let winRateChart = null;
let callsChart = null;
let selfDemosChart = null;
let selfOppsChart = null;
let pipelineChart = null;
let pipelineAgeChart = null;
let avgDaysToCloseChart = null;
let adsChart = null;
let arpuChart = null;
let discountChart = null;

// Date Filter State
let aeDateFilter = { startDate: null, endDate: null, filterType: 'this-quarter' };

// --- Helpers ---
function parseSheetNumber(v){ if(v==null||v==='') return 0; if(typeof v==='number') return v; const n=parseFloat(String(v).replace(/[$,]/g,'')); return isNaN(n)?0:n; }
function parseSheetDate(v){ if(v==null||v==='') return null; if(typeof v==='number'){ return new Date((v-25569)*86400000); } const d=new Date(String(v)); return isNaN(d.getTime())?null:d; }
function to2(n){ const x=Number(n); return Number.isFinite(x)? Math.round(x*100)/100 : 0; }
function toCurrency(n){ return new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',maximumFractionDigits:0}).format(Number(n)||0); }
function formatDate(d){ if(!d) return ''; return d.toLocaleDateString(); }

// --- Date Range Logic ---
function computeQuickRange(type){
  const now=new Date();
  const som=new Date(now.getFullYear(), now.getMonth(), 1);
  const eom=new Date(now.getFullYear(), now.getMonth()+1, 0);
  const q=Math.floor(now.getMonth()/3);
  const soq=new Date(now.getFullYear(), q*3, 1);
  const eoq=new Date(now.getFullYear(), q*3+3, 0);
  const soy=new Date(now.getFullYear(),0,1); const eoy=new Date(now.getFullYear(),11,31);
  switch(type){
    case 'this-month': return {start:som,end:eom};
    case 'last-month': return {start:new Date(now.getFullYear(), now.getMonth()-1, 1), end:new Date(now.getFullYear(), now.getMonth(), 0)};
    case 'this-quarter': return {start:soq,end:eoq};
    case 'last-quarter': return {start:new Date(now.getFullYear(), q*3-3,1), end:new Date(now.getFullYear(), q*3,0)};
    case 'this-year': return {start:soy,end:eoy};
    case 'all-time': return {start:new Date(2020,0,1), end:null};
    default: return {start:null,end:null};
  }
}

// --- Column Mapping ---
// Based on User Request:
// 0: Rep, 1: Date
// 5: Quota, 6: Attained
// 13: Opps Won
// 15: Win Rate
// 17: ADS
// 19: ARPU
// 23: Avg Discount
// 25: Self Demos
// 27: Self Opps Created
// 36: Open Pipeline To Close This Month, 37: Open Pipeline Total, 38: Pipeline Coverage
// 40: Avg Age of Open Opps
// 41: Avg Days To Close Win Opps
// 46: Total Calls, 48: Calls/Day
// 72: QTD ARR, 73: QTD Quota, 76: YTD ARR, 77: YTD Quota

const COL_MAP = {
  REP: 0,
  DATE: 1,
  QUOTA: 5,
  ATTAINED: 6,
  OPPS_WON: 13,
  WIN_RATE: 15,
  ADS: 17,
  ARPU: 19,
  AVG_DISCOUNT: 23,
  SELF_DEMOS: 25,
  SELF_OPPS: 27,
  PIPELINE_TOTAL: 37,
  PIPELINE_COVERAGE: 38,
  AVG_AGE_OPEN: 40,
  AVG_DAYS_CLOSE_WIN: 41,
  TOTAL_CALLS: 46,
  QTD_ARR: 72,
  QTD_QUOTA: 73,
  YTD_ARR: 76,
  YTD_QUOTA: 77
};

// OppSplitData
const SPLIT_MAP = {
  USER: 0,
  OPP_NAME: 1,
  OWNER: 2,
  ACCOUNT: 3,
  STAGE: 4,
  PERCENT: 5,
  AMOUNT: 6,
  ARR: 7,
  TYPE: 8,
  CLOSE_DATE: 9
};

// --- Data Fetching ---
async function loadGoogleAPI(){
  if(gapiInited) return;
  await new Promise((res,rej)=>{
    const s=document.createElement('script');
    s.src='https://apis.google.com/js/api.js';
    s.onload=()=>{ gapi.load('client', async()=>{ try{ await gapi.client.init({ discoveryDocs: CONFIG.DISCOVERY_DOCS }); gapiInited=true; res(); }catch(e){ rej(e); } }); };
    s.onerror=rej;
    document.head.appendChild(s);
  });
}

async function fetchAeScorecard(){
  if(!gapi.client.sheets) await gapi.client.load('sheets','v4');
  const range='AE Scorecard!A:CK'; 
  const resp=await gapi.client.sheets.spreadsheets.values.get({ spreadsheetId: CONFIG.SPREADSHEET_ID, range, valueRenderOption:'UNFORMATTED_VALUE', dateTimeRenderOption:'SERIAL_NUMBER' });
  const values=resp.result.values||[];
  if(values.length<=1) return [];
  return values.slice(1); 
}

async function fetchOppSplitData(){
    if(!gapi.client.sheets) await gapi.client.load('sheets','v4');
    const range='OppSplitData!A:J';
    const resp=await gapi.client.sheets.spreadsheets.values.get({ spreadsheetId: CONFIG.SPREADSHEET_ID, range, valueRenderOption:'UNFORMATTED_VALUE', dateTimeRenderOption:'SERIAL_NUMBER' });
    const values=resp.result.values||[];
    if(values.length<=1) return [];
    return values.slice(1);
}

// --- Processing & Rendering ---

function calculateMetrics(rows){
  const start=aeDateFilter.startDate, end=aeDateFilter.endDate;
  
  // Filter rows
  const filtered = rows.filter(r => {
    const d = parseSheetDate(r[COL_MAP.DATE]);
    if(!d) return false;
    const a = !start || d >= start;
    const b = !end || d <= end;
    return a && b;
  });

  const byRep = {};
  
  filtered.forEach(r => {
    const rep = String(r[COL_MAP.REP]||'').trim();
    if(!rep) return;
    
    if(!byRep[rep]) {
      byRep[rep] = { 
        rows: [], 
        totalCalls: 0, 
        oppsWon: 0, 
        selfDemos: 0,
        selfOpps: 0,
        latestDate: 0,
        latestRow: null
      };
    }
    
    // Summing Activity Metrics
    byRep[rep].rows.push(r);
    byRep[rep].totalCalls += parseSheetNumber(r[COL_MAP.TOTAL_CALLS]);
    byRep[rep].oppsWon += parseSheetNumber(r[COL_MAP.OPPS_WON]);
    byRep[rep].selfDemos += parseSheetNumber(r[COL_MAP.SELF_DEMOS]);
    byRep[rep].selfOpps += parseSheetNumber(r[COL_MAP.SELF_OPPS]);
    
    const dVal = r[COL_MAP.DATE];
    if(dVal > byRep[rep].latestDate){
      byRep[rep].latestDate = dVal;
      byRep[rep].latestRow = r;
    }
  });

  const repList = Object.keys(byRep).map(rep => {
    const data = byRep[rep];
    const lr = data.latestRow || {}; 
    
    const isQuarterly = aeDateFilter.filterType === 'this-quarter' || aeDateFilter.filterType === 'last-quarter';
    const isYearly = aeDateFilter.filterType === 'this-year';
    
    let attained = 0;
    let quota = 0;
    
    if(isQuarterly) {
        attained = parseSheetNumber(lr[COL_MAP.QTD_ARR]);
        quota = parseSheetNumber(lr[COL_MAP.QTD_QUOTA]);
    } else if(isYearly) {
        attained = parseSheetNumber(lr[COL_MAP.YTD_ARR]);
        quota = parseSheetNumber(lr[COL_MAP.YTD_QUOTA]);
    } else {
        attained = parseSheetNumber(lr[COL_MAP.ATTAINED]);
        quota = parseSheetNumber(lr[COL_MAP.QUOTA]);
    }
    
    // Snapshot Metrics from Latest Row
    const winRate = parseSheetNumber(lr[COL_MAP.WIN_RATE]);
    const pipeline = parseSheetNumber(lr[COL_MAP.PIPELINE_TOTAL]);
    const pipeCoverage = parseSheetNumber(lr[COL_MAP.PIPELINE_COVERAGE]);
    const ads = parseSheetNumber(lr[COL_MAP.ADS]);
    const arpu = parseSheetNumber(lr[COL_MAP.ARPU]);
    const avgDiscount = parseSheetNumber(lr[COL_MAP.AVG_DISCOUNT]);
    const avgAgeOpen = parseSheetNumber(lr[COL_MAP.AVG_AGE_OPEN]);
    const avgDaysCloseWin = parseSheetNumber(lr[COL_MAP.AVG_DAYS_CLOSE_WIN]);

    return {
      rep,
      totalCalls: data.totalCalls,
      oppsWon: data.oppsWon,
      selfDemos: data.selfDemos,
      selfOpps: data.selfOpps,
      attained,
      quota,
      winRate, 
      pipeline,
      pipeCoverage,
      ads,
      arpu,
      avgDiscount,
      avgAgeOpen,
      avgDaysCloseWin
    };
  });
  
  return repList.sort((a,b) => b.attained - a.attained);
}

function renderMetrics(repList){
  const totalQuota = repList.reduce((s,r)=>s+r.quota,0);
  const totalAttained = repList.reduce((s,r)=>s+r.attained,0);
  const totalPipeline = repList.reduce((s,r)=>s+r.pipeline,0);
  const avgWinRate = repList.length ? repList.reduce((s,r)=>s+r.winRate,0)/repList.length : 0;

  const cards = document.getElementById('ae-metrics-cards');
  if(cards){
    cards.innerHTML = `
      <div class="bg-white p-4 rounded-lg shadow border">
        <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Total Attained</h3>
        <p class="text-2xl font-bold mt-2 mb-1" style="color:#10B981;">${toCurrency(totalAttained)}</p>
        <p class="text-sm text-gray-600">Quota: ${toCurrency(totalQuota)}</p>
      </div>
      <div class="bg-white p-4 rounded-lg shadow border">
        <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Avg Win Rate</h3>
        <p class="text-2xl font-bold mt-2 mb-1" style="color:#3B82F6;">${to2(avgWinRate*100)}%</p>
        <p class="text-sm text-gray-600">Across ${repList.length} Reps</p>
      </div>
      <div class="bg-white p-4 rounded-lg shadow border">
        <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Total Pipeline</h3>
        <p class="text-2xl font-bold mt-2 mb-1" style="color:#F59E0B;">${toCurrency(totalPipeline)}</p>
        <p class="text-sm text-gray-600">Open Pipeline</p>
      </div>
      <div class="bg-white p-4 rounded-lg shadow border">
        <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Total Calls</h3>
        <p class="text-2xl font-bold mt-2 mb-1" style="color:#6366F1;">${repList.reduce((s,r)=>s+r.totalCalls,0).toLocaleString()}</p>
        <p class="text-sm text-gray-600">Activities in period</p>
      </div>
    `;
  }
}

function createBarChart(id, title, labels, data, color, formatFn){
    const canvas = document.getElementById(id);
    if(!canvas) return null;
    const ctx = canvas.getContext('2d');
    return new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{ label: title, data: data, backgroundColor: color, borderRadius: 4 }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                title: { display: true, text: title, font: { size: 14, weight: '600' } },
                legend: { display: false },
                datalabels: {
                    color: '#111', anchor: 'end', align: 'top', offset: -2,
                    font: { size: 10 },
                    formatter: formatFn || ((v)=>v)
                }
            },
            scales: { y: { beginAtZero: true, grid: { display: true, color: '#f3f4f6' } }, x: { grid: { display: false } } }
        }
    });
}

function renderCharts(repList){
    const labels = repList.map(r=>r.rep);
    
    // Cleanup
    [attainmentChart, winRateChart, callsChart, selfDemosChart, selfOppsChart, pipelineChart, pipelineAgeChart, avgDaysToCloseChart, adsChart, arpuChart, discountChart].forEach(c => { if(c) c.destroy(); });
    
    // 1. Attainment
    const ctxAtt = document.getElementById('attainmentChart');
    if(ctxAtt){
        attainmentChart = new Chart(ctxAtt.getContext('2d'), {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    { label: 'Attained', data: repList.map(r=>r.attained), backgroundColor: 'rgba(16,185,129,0.8)', order: 2 },
                    { label: 'Quota', data: repList.map(r=>r.quota), borderColor: 'rgba(107,114,128,0.8)', type: 'line', fill: false, order: 1, borderDash: [5,5] }
                ]
            },
            options: { responsive:true, maintainAspectRatio:false, plugins: { title: {display:true, text:'Attainment vs Quota'} } }
        });
    }

    // 2. Win Rate
    winRateChart = createBarChart('winRateChart', 'Win Rate %', labels, repList.map(r=>to2(r.winRate*100)), 'rgba(59,130,246,0.8)', (v)=>v+'%');
    
    // 3. Activity
    callsChart = createBarChart('callsChart', 'Total Calls', labels, repList.map(r=>r.totalCalls), 'rgba(99,102,241,0.8)');
    selfDemosChart = createBarChart('selfDemosChart', 'Self Demos', labels, repList.map(r=>r.selfDemos), 'rgba(139,92,246,0.8)');
    selfOppsChart = createBarChart('selfOppsChart', 'Self Opps Created', labels, repList.map(r=>r.selfOpps), 'rgba(217,70,239,0.8)');

    // 4. Pipeline Health
    pipelineChart = createBarChart('pipelineChart', 'Pipeline Coverage (x)', labels, repList.map(r=>r.pipeCoverage), 'rgba(245,158,11,0.8)', (v)=>v+'x');
    pipelineAgeChart = createBarChart('pipelineAgeChart', 'Avg Age Open Opps (Days)', labels, repList.map(r=>Math.round(r.avgAgeOpen)), 'rgba(234,179,8,0.8)');
    avgDaysToCloseChart = createBarChart('avgDaysToCloseChart', 'Avg Days to Close Win', labels, repList.map(r=>Math.round(r.avgDaysCloseWin)), 'rgba(22,163,74,0.8)');

    // 5. Deal Quality
    adsChart = createBarChart('adsChart', 'Avg Deal Size', labels, repList.map(r=>r.ads), 'rgba(14,165,233,0.8)', (v)=>'$'+Math.round(v/1000)+'k');
    arpuChart = createBarChart('arpuChart', 'ARPU', labels, repList.map(r=>r.arpu), 'rgba(6,182,212,0.8)', (v)=>'$'+Math.round(v));
    discountChart = createBarChart('discountChart', 'Avg Discount %', labels, repList.map(r=>to2(r.avgDiscount*100)), 'rgba(244,63,94,0.8)', (v)=>v+'%');
}

function renderTable(repList){
    const head = document.getElementById('aeScorecardHeaderRow');
    if(head){
        head.innerHTML = `
            <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Rep</th>
            <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Quota</th>
            <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Attained</th>
            <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">% Attained</th>
            <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Win Rate</th>
            <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Opps Won</th>
            <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Calls</th>
            <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Self Demos</th>
            <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Pipeline</th>
            <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Coverage</th>
        `;
    }
    
    const body = document.getElementById('aeScorecardBody');
    if(body){
        body.innerHTML = '';
        repList.forEach((r, idx) => {
            const tr = document.createElement('tr');
            tr.className = idx%2===0?'bg-white':'bg-gray-50';
            const pct = r.quota > 0 ? (r.attained/r.quota)*100 : 0;
            tr.innerHTML = `
                <td class="px-6 py-3 text-sm text-gray-900 font-medium">${r.rep}</td>
                <td class="px-6 py-3 text-sm text-gray-900">${toCurrency(r.quota)}</td>
                <td class="px-6 py-3 text-sm text-gray-900">${toCurrency(r.attained)}</td>
                <td class="px-6 py-3 text-sm text-gray-900">${to2(pct)}%</td>
                <td class="px-6 py-3 text-sm text-gray-900">${to2(r.winRate*100)}%</td>
                <td class="px-6 py-3 text-sm text-gray-900">${r.oppsWon}</td>
                <td class="px-6 py-3 text-sm text-gray-900">${r.totalCalls}</td>
                <td class="px-6 py-3 text-sm text-gray-900">${r.selfDemos}</td>
                <td class="px-6 py-3 text-sm text-gray-900">${toCurrency(r.pipeline)}</td>
                <td class="px-6 py-3 text-sm text-gray-900">${to2(r.pipeCoverage)}x</td>
            `;
            body.appendChild(tr);
        });
    }
}

function renderOppTable(rows){
    const start=aeDateFilter.startDate, end=aeDateFilter.endDate;
    // Filter by Close Date
    const filtered = rows.filter(r => {
        const d = parseSheetDate(r[SPLIT_MAP.CLOSE_DATE]);
        if(!d) return false;
        const a = !start || d >= start;
        const b = !end || d <= end;
        return a && b;
    }).sort((a,b) => parseSheetDate(b[SPLIT_MAP.CLOSE_DATE]) - parseSheetDate(a[SPLIT_MAP.CLOSE_DATE]));
    
    const body = document.getElementById('aeOppSplitsBody');
    if(body){
        body.innerHTML = '';
        if(filtered.length === 0){
            body.innerHTML = '<tr><td colspan="7" class="px-6 py-4 text-center text-gray-500">No splits found for this period</td></tr>';
            return;
        }
        // Limit to top 50
        filtered.slice(0,50).forEach((r, idx) => {
            const tr = document.createElement('tr');
            tr.className = idx%2===0?'bg-white':'bg-gray-50';
            const d = parseSheetDate(r[SPLIT_MAP.CLOSE_DATE]);
            const amt = parseSheetNumber(r[SPLIT_MAP.AMOUNT]);
            const arr = parseSheetNumber(r[SPLIT_MAP.ARR]);
            
            tr.innerHTML = `
                <td class="px-6 py-3 text-sm text-gray-900">${formatDate(d)}</td>
                <td class="px-6 py-3 text-sm text-gray-900">
                    <div class="font-medium">${r[SPLIT_MAP.OWNER]||''}</div>
                    <div class="text-xs text-gray-500">${r[SPLIT_MAP.USER]||''}</div>
                </td>
                <td class="px-6 py-3 text-sm text-gray-900">
                    <div class="font-medium">${r[SPLIT_MAP.ACCOUNT]||''}</div>
                    <div class="text-xs text-gray-500">${r[SPLIT_MAP.OPP_NAME]||''}</div>
                </td>
                <td class="px-6 py-3 text-sm text-gray-900">${r[SPLIT_MAP.STAGE]||''}</td>
                <td class="px-6 py-3 text-sm text-gray-900">${r[SPLIT_MAP.TYPE]||''}</td>
                <td class="px-6 py-3 text-sm text-gray-900">${toCurrency(amt)}</td>
                <td class="px-6 py-3 text-sm text-gray-900">${toCurrency(arr)}</td>
            `;
            body.appendChild(tr);
        });
    }
}

function updateDashboard(){
    const metrics = calculateMetrics(aeData);
    renderMetrics(metrics);
    renderCharts(metrics);
    renderTable(metrics);
    renderOppTable(oppSplitData);
}

// --- Auth & Init ---
async function handleAuth(){
  if(isLoading||dataLoaded) return;
  
  // Check if Google scripts are loaded
  if(typeof google === 'undefined' || !google.accounts || !google.accounts.oauth2){
    console.warn('Google GSI script not loaded yet');
    alert('Google Sign-In script is still loading. Please wait a moment and try again.');
    return;
  }

  try{
    await loadGoogleAPI();
    if(!tokenClient){
        tokenClient = google.accounts.oauth2.initTokenClient({
            client_id: CONFIG.CLIENT_ID,
            scope: CONFIG.SCOPES,
            prompt: '',
            callback: async(resp)=>{
                if(resp.error) return;
                isLoading=true;
                const btn=document.getElementById('aeAuthBtn');
                if(btn) { btn.textContent='Loading...'; btn.disabled=true; }
                
                try{
                    const [sc, os] = await Promise.all([fetchAeScorecard(), fetchOppSplitData()]);
                    aeData = sc;
                    oppSplitData = os;
                    dataLoaded = true;
                    updateDashboard();
                    if(btn) btn.textContent='Data Loaded';
                } catch(e){
                    console.error(e);
                    alert('Failed to load data: '+e.message);
                    if(btn) { btn.textContent='Retry Load'; btn.disabled=false; }
                } finally {
                    isLoading=false;
                }
            }
        });
    }
    tokenClient.requestAccessToken();
  } catch(e){
      console.error(e);
      alert('Auth Init Error: '+e.message);
  }
}

function initDateFilters(){
    const select=document.getElementById('aeDateFilterSelect');
    const sIn=document.getElementById('aeStartDateInput');
    const eIn=document.getElementById('aeEndDateInput');
    const btn=document.getElementById('aeApplyFilterBtn');
    const cS=document.getElementById('aeCustomDateRange');
    const cE=document.getElementById('aeCustomDateRangeEnd');
    
    const sync=()=>{ 
        const isC=select.value==='custom'; 
        [cS,cE].forEach(el=>isC?el.classList.remove('hidden'):el.classList.add('hidden')); 
        isC?btn.classList.remove('hidden'):btn.classList.add('hidden'); 
    };
    
    select.addEventListener('change',()=>{ 
        if(select.value==='custom'){ sync(); } 
        else { 
            const r=computeQuickRange(select.value); 
            aeDateFilter={startDate:r.start,endDate:r.end,filterType:select.value}; 
            if(dataLoaded) updateDashboard(); 
            sync(); 
        }
    });
    
    if(btn){ 
        btn.addEventListener('click',()=>{ 
            const s=sIn&&sIn.value?new Date(sIn.value):null; 
            const e=eIn&&eIn.value?new Date(eIn.value):null; 
            aeDateFilter={startDate:s,endDate:e,filterType:'custom'}; 
            if(dataLoaded) updateDashboard(); 
        }); 
    }
    
    const def=computeQuickRange('this-quarter'); 
    aeDateFilter={startDate:def.start,endDate:def.end,filterType:'this-quarter'}; 
    sync();
}

document.addEventListener('DOMContentLoaded', ()=>{
    initDateFilters();
    const btn=document.getElementById('aeAuthBtn');
    if(btn) btn.addEventListener('click', handleAuth);
});
