/* global Chart, gapi, google */

// Minimal config for ADM page (falls back to window.CONFIG if present)
const CONFIG = window.CONFIG || {
  AUTH_METHOD: 'oauth',
  SPREADSHEET_ID: '1XhNpvY1SYsvszBugJeD-gQKar9OwU4lLLiF5cTkAk6I',
  CLIENT_ID: '630950450890-lofitb7ofs2q6ae3olqv1jis88uhfjeg.apps.googleusercontent.com',
  DISCOVERY_DOCS: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
  SCOPES: 'https://www.googleapis.com/auth/spreadsheets.readonly'
};

let tokenClient; let gapiInited = false; let dataLoaded = false; let isLoading = false;
let admActivitiesChart = null; let admActivitiesData = [];
// OppData charts
let oppInboundOutboundCountChart = null, oppAmountBySourceChart = null, oppByADMChart = null, oppStageChart = null, oppCreatedByMonthChart = null;
let oppClosedLostReasonBySourceChart = null, oppAvgWonBySourceChart = null, oppAvgWonByADMChart = null;
let admOppData = [];
// Sequence charts
let seqDemoSetsChart = null, seqAnswersChart = null;

// Date filter state (ADM-only)
let admActivitiesDateFilter = { startDate: null, endDate: null, filterType: 'this-quarter' };

// Allowed years restriction for ADM page
const ALLOWED_YEARS = new Set([2024, 2025]);
function inAllowedYears(d){ return d instanceof Date && !isNaN(d) && ALLOWED_YEARS.has(d.getFullYear()); }

// Inbound/Outbound config and state
const ADM_INBOUND_USERS = ['Tati Kallianiotis', 'Ashlyn Sallade', 'Eric Wintch'];
function nameTokens(s){ return String(s||'').toLowerCase().replace(/[^a-z]/g,' ').split(/\s+/).filter(Boolean); }
function tokenPrefixMatch(a,b){ return a && b && (a.startsWith(b) || b.startsWith(a)); }
const inboundPairs = ADM_INBOUND_USERS.map(n=>{ const t=nameTokens(n); return { first: t[0]||'', last: t[t.length-1]||'' }; });
function isInboundAssigned(name){ const t=nameTokens(name); if(t.length===0) return false; return inboundPairs.some(p=> t.some(x=>tokenPrefixMatch(x,p.first)) && t.some(x=>tokenPrefixMatch(x,p.last)) ); }
let admTeamFilter = 'all'; // 'all' | 'inbound' | 'outbound'

// Register datalabels if available
if (typeof window !== 'undefined' && window.Chart && window.ChartDataLabels) {
  Chart.register(window.ChartDataLabels);
}

function parseSheetNumber(v){ if(v==null||v==='') return 0; if(typeof v==='number') return v; const n=parseFloat(String(v).replace(/[$,]/g,'')); return isNaN(n)?0:n; }
function parseSheetDate(v){ if(v==null||v==='') return null; if(typeof v==='number'){ return new Date((v-25569)*86400000); } const d=new Date(String(v)); return isNaN(d.getTime())?null:d; }
function formatDurationMinutes(mins){ const n=Number(mins); if(!isFinite(n)||n<=0) return '0s'; const totalSeconds = Math.round(n*60); if(totalSeconds < 60) return `${totalSeconds}s`; const m=Math.floor(totalSeconds/60), s=totalSeconds%60; return s===0?`${m}m`:`${m}m ${s}s`; }
function to2(n){ const x=Number(n); return Number.isFinite(x)? Math.round(x*100)/100 : 0; }
function toCurrency(n){ return new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',maximumFractionDigits:0}).format(Number(n)||0); }
// Normalize Column I duration to minutes: Sheets returns day-fractions for durations with UNFORMATTED_VALUE.
function normalizeDurationToMinutes(raw){
  if(raw==null || raw==='') return 0;
  if(typeof raw==='string'){
    // Try HH:MM:SS or MM:SS
    if(/:\d{1,2}(?::\d{1,2})?$/.test(raw)){
      const parts = raw.split(':').map(Number);
      let h=0,m=0,s=0;
      if(parts.length===3){ [h,m,s]=parts; }
      else if(parts.length===2){ [m,s]=parts; }
      return h*60 + m + (s/60);
    }
    const n = parseSheetNumber(raw);
    // Heuristic: values < 1 are likely day-fractions → minutes
    return n < 1 ? n*24*60 : n;
  }
  if(typeof raw==='number'){
    // Values < 1 → day-fraction; convert to minutes
    return raw < 1 ? raw*24*60 : raw;
  }
  return 0;
}

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
    case 'from-2025-01-01': return {start:new Date(2025,0,1), end:null};
    default: return {start:null,end:null};
  }
}

// Attempt to get a sequence value from a row.
// If a dedicated Sequence column exists, plug its index here. Fallback: combine Disposition + Activity Type.
function extractSequenceFromRow(r){
  // Activities column T = index 19
  const SEQ_COL = 19;
  const COL_TYPE=7, COL_DISP=11;
  const raw = r && r.length>SEQ_COL ? r[SEQ_COL] : '';
  const val = String(raw||'').trim();
  if(val && val.toLowerCase()!=='n/a' && val.toLowerCase()!=='none') return val;
  // Fallback: Disposition • Activity Type
  const disp = String(r[COL_DISP]||'').trim();
  const type = String(r[COL_TYPE]||'').trim();
  const seq = [disp||'', type||''].filter(Boolean).join(' • ');
  return seq || 'Unspecified';
}

function destroySequenceCharts(){ if(seqDemoSetsChart){ seqDemoSetsChart.destroy(); seqDemoSetsChart=null; } if(seqAnswersChart){ seqAnswersChart.destroy(); seqAnswersChart=null; } }

function renderActivitiesSequenceInsights(rows){
  try{
    destroySequenceCharts();
    const COL_DATE=17, COL_ASSIGNED=1, COL_ROLE=2, COL_TYPE=7, COL_CALL_RESULT=9, COL_DISP=11, COL_CONNECT=21;
    if(!Array.isArray(rows) || rows.length===0) return;
    // Filter: ADM role
    const roleRows = rows.filter(r => String(r[COL_ROLE]||'').toLowerCase().trim()==='adm');
    // Filter: team selection using assigned name
    const teamRows = roleRows.filter(r => {
      const assigned = r[COL_ASSIGNED]; const inbound = isInboundAssigned(assigned);
      if (admTeamFilter === 'inbound') return inbound; if (admTeamFilter === 'outbound') return !inbound; return true;
    });
    // Filter: date + allowed years
    const start=admActivitiesDateFilter.startDate, end=admActivitiesDateFilter.endDate;
    const dateRows = teamRows.filter(r=>{ const d=parseSheetDate(r[COL_DATE]); if(!d) return false; if(!inAllowedYears(d)) return false; const a=!start||d>=start; const b=!end||d<=end; return a&&b; });
    // Continue rendering even if empty; we'll show 'No Data' placeholder

    // Aggregate by sequence
    const bySeq = {};
    dateRows.forEach(r=>{
      const seq = extractSequenceFromRow(r);
      const typeRaw = String(r[COL_TYPE]||'').toLowerCase();
      const callRes = String(r[COL_CALL_RESULT]||'').toLowerCase();
      const disp = String(r[COL_DISP]||'').toLowerCase();
      const conn = String(r[COL_CONNECT]||'').toLowerCase();
      const isCall = typeRaw.includes('call')||typeRaw.includes('phone');
      if(!bySeq[seq]) bySeq[seq] = { demos:0, calls:0, answered:0 };
      if(callRes==='demo set' || callRes.includes('demo')) bySeq[seq].demos += 1;
      if(isCall){ bySeq[seq].calls += 1; if(conn==='connected'||conn==='yes'||disp==='connected') bySeq[seq].answered += 1; }
    });

    // Top sequences by demo sets
    const entries = Object.entries(bySeq);
    const topDemos = entries.sort((a,b)=>b[1].demos - a[1].demos).slice(0,10);
    const demoLabels = topDemos.length? topDemos.map(e=>e[0]) : ['No Data'];
    const demoValues = topDemos.length? topDemos.map(([,v])=>v.demos) : [0];
    const demoTitle = topDemos.length ? 'Top Sequences by Demo Sets' : 'No data for selection';
    const c1 = document.getElementById('seqDemoSetsChart');
    if(c1){
      seqDemoSetsChart = new Chart(c1.getContext('2d'),{
        type:'bar', data:{ labels: demoLabels, datasets:[{ label:'Demo Sets', data: demoValues, backgroundColor:'rgba(16,185,129,0.8)', borderRadius:6 }] },
        options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, title:{display:true,text: demoTitle}, datalabels:{ color:'#111', font:{size:10,weight:'600'}, formatter:(v)=>v||'' } }, scales:{ x:{ beginAtZero:true } } }
      });
    }

    // Top sequences by answered rate (min 20 calls to avoid noise)
    const minCalls = 10;
    const qualifiedRaw = entries.map(([k,v])=>({ key:k, rate: v.calls>0? v.answered/v.calls : 0, calls:v.calls })).filter(x=>x.calls>=minCalls).sort((a,b)=>b.rate-a.rate).slice(0,10);
    const ansLabels = qualifiedRaw.length? qualifiedRaw.map(x=>x.key) : ['No Data'];
    const ansValues = qualifiedRaw.length? qualifiedRaw.map(x=> to2(x.rate*100)) : [0];
    const ansTitle = qualifiedRaw.length ? 'Top Sequences by Answered % (min 10 calls)' : 'No data for selection';
    const c2 = document.getElementById('seqAnswersChart');
    if(c2){
      seqAnswersChart = new Chart(c2.getContext('2d'),{
        type:'bar', data:{ labels: ansLabels, datasets:[{ label:'Answered %', data: ansValues, backgroundColor:'rgba(59,130,246,0.8)', borderRadius:6 }] },
        options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, title:{display:true,text: ansTitle}, datalabels:{ color:'#111', font:{size:10,weight:'600'}, formatter:(v)=> `${to2(v)}%` } }, scales:{ x:{ beginAtZero:true, ticks:{ callback:(v)=>`${v}%` } } } }
      });
    }
  }catch(e){ console.error('Error rendering Activities Sequence Insights', e); }
}

function initADMDateFilterControls(){
  const select=document.getElementById('admActivitiesDateFilterSelect');
  const sIn=document.getElementById('admActivitiesStartDateInput');
  const eIn=document.getElementById('admActivitiesEndDateInput');
  const btn=document.getElementById('admActivitiesApplyFilterBtn');
  const cS=document.getElementById('admActivitiesCustomDateRange');
  const cE=document.getElementById('admActivitiesCustomDateRangeEnd');
  if(!select) return;
  const sync=()=>{ const isC=select.value==='custom'; [cS,cE].forEach(el=>isC?el.classList.remove('hidden'):el.classList.add('hidden')); isC?btn.classList.remove('hidden'):btn.classList.add('hidden'); };
  select.addEventListener('change',()=>{ if(select.value==='custom'){ sync(); } else { const r=computeQuickRange(select.value); admActivitiesDateFilter={startDate:r.start,endDate:r.end,filterType:select.value}; if(admActivitiesData.length) renderADMActivities(admActivitiesData); if(admOppData.length) renderOppCharts(admOppData); sync(); }});
  if(btn){ btn.addEventListener('click',()=>{ const s=sIn&&sIn.value?new Date(sIn.value):null; const e=eIn&&eIn.value?new Date(eIn.value):null; admActivitiesDateFilter={startDate:s,endDate:e,filterType:'custom'}; if(admActivitiesData.length) renderADMActivities(admActivitiesData); if(admOppData.length) renderOppCharts(admOppData); }); }
  const def=computeQuickRange('this-quarter'); admActivitiesDateFilter={startDate:def.start,endDate:def.end,filterType:'this-quarter'}; sync();
}

function calculateADMActivitiesMetrics(rows){
  if(!rows||rows.length===0){ return { totalCalls:0,totalEmails:0,demoSets:0,avgCallDurationForDemos:0,callsPerDemo:0,emailsPerDemo:0,callsPerAnswered:0,filteredActivities:0, perUser:[] }; }
  const COL_DATE=17, COL_ASSIGNED=1, COL_ROLE=2, COL_TYPE=7, COL_CALL_DUR=8, COL_CALL_RESULT=9, COL_DISP=11, COL_CONNECT=21;
  // Filter: role === 'adm'
  const roleRows = rows.filter(r => String(r[COL_ROLE]||'').toLowerCase().trim()==='adm');
  // Filter by team selection
  const teamRows = roleRows.filter(r => {
    const assigned = r[COL_ASSIGNED];
    const inbound = isInboundAssigned(assigned);
    if (admTeamFilter === 'inbound') return inbound;
    if (admTeamFilter === 'outbound') return !inbound;
    return true;
  });
  // Date filter
  const start=admActivitiesDateFilter.startDate, end=admActivitiesDateFilter.endDate;
  const dateRows = teamRows.filter(r=>{ const d=parseSheetDate(r[COL_DATE]); if(!d) return false; if(!inAllowedYears(d)) return false; const a=!start||d>=start; const b=!end||d<=end; return a&&b; });
  let totalCalls=0,totalEmails=0,demoSets=0, callDurForDemos=[], connected=0;
  const perUser={};
  dateRows.forEach(r=>{
    const typeRaw=String(r[COL_TYPE]||'').toLowerCase();
    const callRes=String(r[COL_CALL_RESULT]||'').toLowerCase();
    const disp=String(r[COL_DISP]||'').toLowerCase();
    const conn=String(r[COL_CONNECT]||'').toLowerCase();
    const user=String(r[COL_ASSIGNED]||'Unknown').trim()||'Unknown';
    const isCall=typeRaw.includes('call')||typeRaw.includes('phone');
    const isEmail=typeRaw.includes('email')||typeRaw.includes('e-mail');
    if(!perUser[user]) perUser[user]={calls:0,emails:0,demos:0,connected:0,demoDurSum:0,demoDurCnt:0};
    if(isCall){ totalCalls++; perUser[user].calls++; if(conn==='connected'||conn==='yes'||disp==='connected'){ connected++; perUser[user].connected++; } }
    if(isEmail){ totalEmails++; perUser[user].emails++; }
    if(callRes==='demo set'||callRes.includes('demo')){
      demoSets++; perUser[user].demos++;
      if(isCall){
        const durMin = normalizeDurationToMinutes(r[COL_CALL_DUR]);
        if(durMin>0){ callDurForDemos.push(durMin); perUser[user].demoDurSum+=durMin; perUser[user].demoDurCnt++; }
      }
    }
  });
  const avgDemo = to2(callDurForDemos.length? callDurForDemos.reduce((a,b)=>a+b,0)/callDurForDemos.length:0);
  const callsPerDemo = to2(demoSets>0? totalCalls/demoSets:0);
  const emailsPerDemo = to2(demoSets>0? totalEmails/demoSets:0);
  const callsPerAnswered = to2(totalCalls>0? (connected/totalCalls)*100:0);
  const perUserArr = Object.entries(perUser).map(([u,s])=>({
    user:u,
    totalCalls:s.calls, totalEmails:s.emails, demoSets:s.demos,
    avgCallDurationForDemos: to2(s.demoDurCnt>0? s.demoDurSum/s.demoDurCnt:0),
    callsPerDemo: to2(s.demos>0? s.calls/s.demos:0),
    emailsPerDemo: to2(s.demos>0? s.emails/s.demos:0),
    callsPerAnswered: to2(s.calls>0? (s.connected/s.calls)*100:0),
    totalActivities: s.calls+s.emails
  })).sort((a,b)=>b.totalActivities-a.totalActivities);
  return { totalCalls,totalEmails,demoSets,avgCallDurationForDemos:avgDemo,callsPerDemo,emailsPerDemo,callsPerAnswered, filteredActivities:dateRows.length, perUser:perUserArr };
}

function renderADMActivities(rows){
  // Destroy old chart
  if(admActivitiesChart){ admActivitiesChart.destroy(); admActivitiesChart=null; }
  const metrics = calculateADMActivitiesMetrics(rows);
  const labels=['Calls','Emails','Demo Sets','Avg Call Duration','Calls/Demo','Emails/Demo','Answered Calls %'];
  const values=[metrics.totalCalls,metrics.totalEmails,metrics.demoSets,metrics.avgCallDurationForDemos,metrics.callsPerDemo,metrics.emailsPerDemo,metrics.callsPerAnswered];
  const colors=['rgba(59,130,246,0.8)','rgba(16,185,129,0.8)','rgba(245,158,11,0.8)','rgba(239,68,68,0.8)','rgba(168,85,247,0.8)','rgba(236,72,153,0.8)','rgba(34,197,94,0.8)'];
  const canvas=document.getElementById('admActivitiesChart');
  if(canvas){
    const ctx=canvas.getContext('2d');
    admActivitiesChart=new Chart(ctx,{
      type:'bar',
      data:{ labels, datasets:[{ label:'ADM Metrics', data:values.map(to2), backgroundColor:colors, borderColor:colors.map(c=>c.replace('0.8','1')), borderWidth:1, borderRadius:6, barThickness:40 }] },
      options:{ responsive:true, maintainAspectRatio:false, resizeDelay:200,
        plugins:{ legend:{display:false}, title:{display:true, text:'ADM Activities', font:{size:16, weight:'600'}, padding:20}, datalabels:{ display:false }, tooltip:{ callbacks:{ label:(ctx)=>{ const v=ctx.raw, l=ctx.label; if(l==='Avg Call Duration') return `${l}: ${formatDurationMinutes(v)}`; if(l==='Calls/Demo'||l==='Emails/Demo') return `${l}: ${to2(v)}`; if(l==='Answered Calls %') return `${l}: ${to2(v)}%`; return `${l}: ${Number(v).toLocaleString()}`; } } } },
        scales:{ y:{ beginAtZero:true, grid:{display:true,color:'rgba(0,0,0,0.05)'}, ticks:{font:{size:11}} }, x:{ grid:{display:false}, ticks:{font:{size:11}, maxRotation:45, minRotation:45} } }
      }
    });
  }
  // Metrics cards
  const mDiv=document.getElementById('adm-activities-metrics');
  if(mDiv){
    mDiv.innerHTML=`
      <div class="bg-white p-4 rounded-lg shadow border">
        <h3 class="text-sm font-semibold text-gray-600 uppercase tracking-wide">Total Activities</h3>
        <p class="text-2xl font-bold mt-2 mb-1" style="background: linear-gradient(135deg, #2B4BFF 0%, #0BE6C7 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">${(metrics.totalCalls+metrics.totalEmails).toLocaleString()}</p>
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
      </div>`;
  }
  // Table
  const tbody=document.getElementById('admActivitiesBreakdownTableBody');
  if(tbody){
    tbody.innerHTML='';
    const rows=Array.isArray(metrics.perUser)?metrics.perUser:[];
    if(rows.length===0){ const tr=document.createElement('tr'); tr.innerHTML='<td colspan="8" class="px-6 py-4 text-center text-sm text-gray-500">No activity data for the selected range</td>'; tbody.appendChild(tr); }
    else {
      rows.forEach((r,idx)=>{
        const tr=document.createElement('tr'); tr.className=idx%2===0?'bg-white':'bg-gray-50';
        const answered=Number.isFinite(r.callsPerAnswered)? to2(r.callsPerAnswered).toFixed(2):'0.00';
        const cpd=Number.isFinite(r.callsPerDemo)? to2(r.callsPerDemo).toFixed(2):'0.00';
        const epd=Number.isFinite(r.emailsPerDemo)? to2(r.emailsPerDemo).toFixed(2):'0.00';
        const avgDemo=formatDurationMinutes(r.avgCallDurationForDemos);
        tr.innerHTML=`
          <td class="px-6 py-3 text-sm text-gray-900">${r.user}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${(r.totalCalls||0).toLocaleString()}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${(r.totalEmails||0).toLocaleString()}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${(r.demoSets||0).toLocaleString()}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${answered}%</td>
          <td class="px-6 py-3 text-sm text-gray-900">${cpd}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${epd}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${avgDemo}</td>`;
        tbody.appendChild(tr);
      });
    }
  }
  // Render Activities Sequence Insights after activities are computed
  renderActivitiesSequenceInsights(admActivitiesData || []);
}

async function loadGoogleAPI(){
  if(gapiInited) return; await new Promise((res,rej)=>{ const s=document.createElement('script'); s.src='https://apis.google.com/js/api.js'; s.onload=()=>{ gapi.load('client', async()=>{ try{ await gapi.client.init({ discoveryDocs: CONFIG.DISCOVERY_DOCS }); gapiInited=true; res(); }catch(e){ rej(e); } }); }; s.onerror=rej; document.head.appendChild(s); });
}

async function fetchActivitiesDataOAuth(){
  if(!gapi.client.sheets) await gapi.client.load('sheets','v4');
  const range='Activities!A:Z';
  const resp=await gapi.client.sheets.spreadsheets.values.get({ spreadsheetId: CONFIG.SPREADSHEET_ID, range, valueRenderOption:'FORMATTED_VALUE' });
  const values=resp.result.values||[]; if(values.length<=1) return []; return values.slice(1); // drop header
}

function setLoading(active){
  const overlay = document.getElementById('loadingOverlay');
  if(overlay) overlay.style.display = active ? 'flex' : 'none';
  const btn = document.getElementById('admAuthBtn');
  if(btn){
    if(active){
      btn.disabled = true;
      btn.textContent = 'Loading...';
      btn.classList.add('opacity-70');
    } else {
      btn.disabled = false;
      btn.textContent = dataLoaded ? 'Data Loaded' : 'Sign in to Load Data';
      btn.classList.remove('opacity-70');
    }
  }
}

async function handleAdmAuth(){
  if(isLoading||dataLoaded) return;
  try{
    await loadGoogleAPI();
    if(!tokenClient){ tokenClient=google.accounts.oauth2.initTokenClient({ client_id: CONFIG.CLIENT_ID, scope: CONFIG.SCOPES, prompt:'', callback: async(resp)=>{
      if(resp.error) return;
      isLoading=true;
      setLoading(true);
      try{
        console.log('[ADM] OAuth callback OK, fetching Activities...');
        const rows=await fetchActivitiesDataOAuth();
        console.log('[ADM] Activities rows:', rows.length);
        admActivitiesData=rows;
        renderADMActivities(admActivitiesData);
        // Fetch OppData in parallel after Activities
        try {
          console.log('[ADM] Fetching OppData...');
          const oppRows = await fetchOppDataOAuth();
          console.log('[ADM] OppData rows:', oppRows.length);
          admOppData = oppRows;
          renderOppCharts(admOppData);
          renderThisWeekSection();
        } catch (e) {
          console.warn('Failed to fetch OppData', e);
        }
        dataLoaded=true;
      }catch(e){ console.error('ADM fetch error', e); alert('Failed to load ADM activities: '+e.message); }
      finally{ isLoading=false; setLoading(false); }
    }}); }
    tokenClient.requestAccessToken();
  }catch(e){ console.error('Auth init error', e); alert('Authentication error: '+e.message); setLoading(false); }
}

document.addEventListener('DOMContentLoaded',()=>{
  initADMDateFilterControls();
  const btn=document.getElementById('admAuthBtn'); if(btn) btn.addEventListener('click', handleAdmAuth);
  initADMTeamToggles();
});

// Initialize Inbound/Outbound/All toggle and wire events
function initADMTeamToggles(){
  const inboundBtn = document.getElementById('admInboundToggle');
  const outboundBtn = document.getElementById('admOutboundToggle');
  const allBtn = document.getElementById('admAllToggle');
  if(!inboundBtn || !outboundBtn || !allBtn) return;

  const setActive = () => {
    const activeClasses = ['bg-blue-500','text-white'];
    const inactiveClasses = ['text-gray-700'];
    const reset = (el, active) => {
      if(!el) return;
      el.classList.remove('bg-blue-500','text-white','hover:bg-gray-200');
      if(active){ activeClasses.forEach(c=>el.classList.add(c)); }
      else { inactiveClasses.forEach(c=>el.classList.add(c)); el.classList.add('hover:bg-gray-200'); }
    };
    reset(inboundBtn, admTeamFilter==='inbound');
    reset(outboundBtn, admTeamFilter==='outbound');
    reset(allBtn, admTeamFilter==='all');
  };

  inboundBtn.addEventListener('click', ()=>{ admTeamFilter='inbound'; console.log('[ADM] Team filter -> inbound'); setActive(); renderADMActivities(admActivitiesData||[]); renderOppCharts(admOppData||[]); renderThisWeekSection(); });
  outboundBtn.addEventListener('click', ()=>{ admTeamFilter='outbound'; console.log('[ADM] Team filter -> outbound'); setActive(); renderADMActivities(admActivitiesData||[]); renderOppCharts(admOppData||[]); renderThisWeekSection(); });
  allBtn.addEventListener('click', ()=>{ admTeamFilter='all'; console.log('[ADM] Team filter -> all'); setActive(); renderADMActivities(admActivitiesData||[]); renderOppCharts(admOppData||[]); renderThisWeekSection(); });

  // Initial state
  setActive();
}

// Fetch OppData (ADM-sourced opportunities)
async function fetchOppDataOAuth(){
  if(!gapi.client.sheets) await gapi.client.load('sheets','v4');
  // Fetch broad range to include all columns up to AN
  const range='OppData!A:AN';
  const resp=await gapi.client.sheets.spreadsheets.values.get({ spreadsheetId: CONFIG.SPREADSHEET_ID, range, valueRenderOption:'UNFORMATTED_VALUE', dateTimeRenderOption:'SERIAL_NUMBER' });
  const values=resp.result.values||[]; if(values.length<=1) return []; return values.slice(1);
}

// Render OppData charts focused on ADM-sourced opps (R not blank)
function ensureOppCanvas(id){
  const canvas = document.getElementById(id);
  if (canvas) return canvas;
  // Rebuild canvas if it was removed
  const selector = `#${id}`;
  // Try to find the container: the layout uses canvas directly in chart-container
  // We recreate the canvas inside the corresponding container by known order
  // As a fallback, append to the first chart-container not having a canvas with that id
  const containers = Array.from(document.querySelectorAll('.chart-container'));
  let parent = containers.find(c=>!c.querySelector(selector));
  if (!parent) parent = containers[0];
  const el = document.createElement('canvas');
  el.id = id;
  parent.appendChild(el);
  return el;
}

function isClosedWonValue(v){
  const s = String(v||'').toLowerCase().trim();
  return s==='true' || s==='yes' || s==='y' || s==='1' || v===true || v===1;
}

function renderOppDeeperInsights(dataset){
  try{
    const COL_SOURCE=15, COL_SUBSOURCE=16, COL_ADM=17, COL_STAGE=3, COL_AMOUNT=37, COL_CLOSED_WON=34, COL_LOST_REASON=31;
    const classify = (row)=>{
      const srcRaw = String(row[COL_SOURCE]||'').trim().toLowerCase();
      const subRaw = String(row[COL_SUBSOURCE]||'').trim().toLowerCase();
      const adm = String(row[COL_ADM]||'').trim();
      const inboundRep = isInboundAssigned(adm);
      const isMarketing = srcRaw.includes('marketing');
      const isSales = srcRaw.includes('sales');
      const isProspectScheduled = subRaw.includes('prospect') && subRaw.includes('sched');
      if (isMarketing && isProspectScheduled) return 'MarketingInbound';
      if (isSales && inboundRep) return 'InboundADM';
      if (isMarketing || isSales) return 'Outbound';
      return 'Other';
    };
    const isLost = (row)=>{ const st=String(row[COL_STAGE]||'').toLowerCase(); return st.includes('lost'); };

    // Early empty-state handling
    if(!Array.isArray(dataset) || dataset.length===0){
      const ids=['oppClosedLostReasonBySourceChart','oppAvgWonBySourceChart','oppAvgWonByADMChart'];
      ids.forEach(id=>{ const c=document.getElementById(id); const parent=c&&c.parentElement; if(parent){ parent.innerHTML='<div class="flex items-center justify-center h-full text-sm text-gray-500">No ADM-sourced opportunities found for this selection</div>'; }});
      const teamBody=document.getElementById('oppTeamSummaryTableBody'); if(teamBody){ teamBody.innerHTML='<tr><td colspan="7" class="px-6 py-4 text-center text-sm text-gray-500">No data for selection</td></tr>'; }
      const inBody=document.getElementById('oppAdmInboundSummaryTableBody'); if(inBody){ inBody.innerHTML='<tr><td colspan="8" class="px-6 py-4 text-center text-sm text-gray-500">No data for selection</td></tr>'; }
      const outBody=document.getElementById('oppAdmOutboundSummaryTableBody'); if(outBody){ outBody.innerHTML='<tr><td colspan="8" class="px-6 py-4 text-center text-sm text-gray-500">No data for selection</td></tr>'; }
      if(oppClosedLostReasonBySourceChart){ oppClosedLostReasonBySourceChart.destroy(); oppClosedLostReasonBySourceChart=null; }
      if(oppAvgWonBySourceChart){ oppAvgWonBySourceChart.destroy(); oppAvgWonBySourceChart=null; }
      if(oppAvgWonByADMChart){ oppAvgWonByADMChart.destroy(); oppAvgWonByADMChart=null; }
      return;
    }

    // Aggregations
    const lostReasons = {}; // reason -> {MarketingInbound, InboundADM, Outbound, total}
    const byTeam = {
      MarketingInbound:{opps:0, won:0, wonSum:0, wonCnt:0, lostReasons:{}},
      InboundADM:{opps:0, won:0, wonSum:0, wonCnt:0, lostReasons:{}},
      Outbound:{opps:0, won:0, wonSum:0, wonCnt:0, lostReasons:{}}
    };
    const byAdm = {}; // adm -> {MarketingInbound:{wonSum,wonCnt}, InboundADM:{wonSum,wonCnt}, Outbound:{wonSum,wonCnt}, totalOpps, won, inboundOpps, outboundOpps}

    dataset.forEach(r=>{
      const team = classify(r);
      if(team!=='MarketingInbound' && team!=='InboundADM' && team!=='Outbound') return; // ignore Other
      const adm = String(r[COL_ADM]||'Unknown').trim()||'Unknown';
      const amt = parseSheetNumber(r[COL_AMOUNT]);
      const won = isClosedWonValue(r[COL_CLOSED_WON]);

      // byTeam
      byTeam[team].opps += 1;
      if(won){ byTeam[team].won += 1; byTeam[team].wonSum += amt; byTeam[team].wonCnt += 1; }

      // lost reasons
      const reason = String(r[COL_LOST_REASON]||'').trim();
      if(!won && (isLost(r) || reason)){ // count if lost stage or reason present
        if(!lostReasons[reason]) lostReasons[reason] = {MarketingInbound:0, InboundADM:0, Outbound:0, total:0};
        lostReasons[reason][team] = (lostReasons[reason][team]||0)+1;
        lostReasons[reason].total += 1;
        byTeam[team].lostReasons[reason] = (byTeam[team].lostReasons[reason]||0)+1;
      }

      // by ADM
      if(!byAdm[adm]) byAdm[adm] = { MarketingInbound:{wonSum:0, wonCnt:0}, InboundADM:{wonSum:0, wonCnt:0}, Outbound:{wonSum:0, wonCnt:0}, totalOpps:0, won:0, inboundOpps:0, outboundOpps:0 };
      byAdm[adm].totalOpps += 1; if(won) byAdm[adm].won += 1;
      if(team==='MarketingInbound' || team==='InboundADM') byAdm[adm].inboundOpps += 1; else if(team==='Outbound') byAdm[adm].outboundOpps += 1;
      if(won){ byAdm[adm][team].wonSum += amt; byAdm[adm][team].wonCnt += 1; }
    });

    // Build Closed Lost Reason by Source (top 12)
    if(oppClosedLostReasonBySourceChart){ oppClosedLostReasonBySourceChart.destroy(); oppClosedLostReasonBySourceChart=null; }
    const cR = ensureOppCanvas('oppClosedLostReasonBySourceChart');
    if(cR){
      const top = Object.entries(lostReasons).filter(([k])=>k && k!=='').sort((a,b)=>b[1].total - a[1].total).slice(0,12);
      const labels = top.map(e=>e[0]);
      const mInboundData = top.map(([,v])=>v.MarketingInbound||0);
      const iAdmData = top.map(([,v])=>v.InboundADM||0);
      const outboundData = top.map(([,v])=>v.Outbound||0);
      oppClosedLostReasonBySourceChart = new Chart(cR.getContext('2d'),{
        type:'bar', data:{ labels, datasets:[
          {label:'Marketing Inbound', data: mInboundData, backgroundColor:'rgba(16,185,129,0.8)'},
          {label:'Inbound ADM', data: iAdmData, backgroundColor:'rgba(34,197,94,0.8)'},
          {label:'Outbound', data: outboundData, backgroundColor:'rgba(59,130,246,0.8)'}
        ]},
        options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, plugins:{ title:{display:true,text:'Top Closed Lost Reasons by Source'}, datalabels:{ color:'#111', font:{size:10,weight:'600'}, formatter:(v)=>v||'' } }, scales:{ x:{ beginAtZero:true, stacked:true }, y:{ stacked:true } } }
      });
    }

    // Avg Won by Source
    if(oppAvgWonBySourceChart){ oppAvgWonBySourceChart.destroy(); oppAvgWonBySourceChart=null; }
    const cS = ensureOppCanvas('oppAvgWonBySourceChart');
    if(cS){
      const labels=['Marketing Inbound','Inbound ADM','Outbound'];
      const avgMarketingInbound = byTeam.MarketingInbound.wonCnt? byTeam.MarketingInbound.wonSum/byTeam.MarketingInbound.wonCnt : 0;
      const avgInboundADM = byTeam.InboundADM.wonCnt? byTeam.InboundADM.wonSum/byTeam.InboundADM.wonCnt : 0;
      const avgOutbound = byTeam.Outbound.wonCnt? byTeam.Outbound.wonSum/byTeam.Outbound.wonCnt : 0;
      oppAvgWonBySourceChart = new Chart(cS.getContext('2d'),{
        type:'bar', data:{ labels, datasets:[{ label:'Avg Won Amount', data:[avgMarketingInbound, avgInboundADM, avgOutbound], backgroundColor:['rgba(16,185,129,0.8)','rgba(34,197,94,0.8)','rgba(59,130,246,0.8)'], borderRadius:6 }] },
        options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, title:{display:true, text:'Average Closed Won Amount by Source'}, datalabels:{ color:'#111', font:{weight:'600',size:11}, formatter:(v)=>toCurrency(v) } }, scales:{ y:{ beginAtZero:true } } }
      });
    }

    // Avg Won by ADM (top 10 by won count)
    if(oppAvgWonByADMChart){ oppAvgWonByADMChart.destroy(); oppAvgWonByADMChart=null; }
    const cA = ensureOppCanvas('oppAvgWonByADMChart');
    if(cA){
      const entries = Object.entries(byAdm).map(([adm,v])=>{
        const am = v.MarketingInbound.wonCnt? v.MarketingInbound.wonSum/v.MarketingInbound.wonCnt : 0;
        const ai = v.InboundADM.wonCnt? v.InboundADM.wonSum/v.InboundADM.wonCnt : 0;
        const ao = v.Outbound.wonCnt? v.Outbound.wonSum/v.Outbound.wonCnt : 0;
        return { adm, am, ai, ao, won:v.won };
      }).sort((a,b)=>b.won - a.won).slice(0,10);
      const labels = entries.map(e=>e.adm);
      const mInboundData = entries.map(e=>e.am);
      const iAdmData = entries.map(e=>e.ai);
      const outboundData = entries.map(e=>e.ao);
      oppAvgWonByADMChart = new Chart(cA.getContext('2d'),{
        type:'bar', data:{ labels, datasets:[
          { label:'Marketing Inbound Avg', data: mInboundData, backgroundColor:'rgba(16,185,129,0.8)' },
          { label:'Inbound ADM Avg', data: iAdmData, backgroundColor:'rgba(34,197,94,0.8)' },
          { label:'Outbound Avg', data: outboundData, backgroundColor:'rgba(59,130,246,0.8)' }
        ] },
        options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, layout:{ padding:{ right:24 } }, plugins:{ title:{display:true,text:'Average Closed Won Amount by ADM'}, datalabels:{ anchor:'end', align:'right', offset:4, clamp:true, color:'#111', backgroundColor:'rgba(255,255,255,0.85)', borderColor:'rgba(0,0,0,0.08)', borderWidth:1, borderRadius:4, padding:3, font:{size:9,weight:'600'}, formatter:(v)=>toCurrency(v) } }, scales:{ x:{ beginAtZero:true }, y:{ stacked:false } } }
      });
    }

    // Fill tables
    const teamBody = document.getElementById('oppTeamSummaryTableBody');
    if(teamBody){
      teamBody.innerHTML='';
      const teams = ['MarketingInbound','InboundADM','Outbound'];
      teams.forEach(t=>{
        const m = byTeam[t]; if(!m) return;
        const winRate = m.opps? (m.won/m.opps)*100 : 0;
        // Top lost reason for team
        const lrEntries = Object.entries(m.lostReasons||{}).sort((a,b)=>b[1]-a[1]);
        const topLR = lrEntries.length? lrEntries[0][0] : '';
        const tr=document.createElement('tr');
        tr.innerHTML=`
          <td class="px-6 py-3 text-sm text-gray-900">${t==='MarketingInbound'?'Marketing Inbound':(t==='InboundADM'?'Inbound ADM':'Outbound')}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${m.opps.toLocaleString()}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${m.won.toLocaleString()}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${to2(winRate).toFixed(2)}%</td>
          <td class="px-6 py-3 text-sm text-gray-900">${toCurrency(m.wonCnt? m.wonSum/m.wonCnt : 0)}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${toCurrency(m.wonSum)}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${topLR||'—'}</td>`;
        teamBody.appendChild(tr);
      });
      if(!teamBody.children.length){ const tr=document.createElement('tr'); tr.innerHTML='<td colspan="7" class="px-6 py-4 text-center text-sm text-gray-500">No data for selection</td>'; teamBody.appendChild(tr); }
    }

    const inboundBody = document.getElementById('oppAdmInboundSummaryTableBody');
    const outboundBody = document.getElementById('oppAdmOutboundSummaryTableBody');
    if(inboundBody || outboundBody){
      if(inboundBody) inboundBody.innerHTML='';
      if(outboundBody) outboundBody.innerHTML='';
      const entries = Object.entries(byAdm).map(([adm,v])=>{
        const inboundCount = (v.inboundOpps||0);
        const outboundCount = v.outboundOpps||0;
        const avgWonInbound = (function(){
          const sum = (v.MarketingInbound.wonSum||0) + (v.InboundADM.wonSum||0);
          const cnt = (v.MarketingInbound.wonCnt||0) + (v.InboundADM.wonCnt||0);
          return cnt? sum/cnt : 0;
        })();
        const avgWonOutbound = v.Outbound.wonCnt? v.Outbound.wonSum/v.Outbound.wonCnt:0;
        const totalWonSum = (v.MarketingInbound.wonSum||0) + (v.InboundADM.wonSum||0) + (v.Outbound.wonSum||0);
        return { adm, inboundCount, outboundCount, won:v.won, totalOpps:v.totalOpps, avgWonInbound, avgWonOutbound, totalWonSum };
      }).sort((a,b)=> b.won - a.won);

      // Compute top lost reason per ADM with second pass
      const topLostByAdm = {};
      dataset.forEach(r=>{
        const team = classify(r); if(team!=='MarketingInbound' && team!=='InboundADM' && team!=='Outbound') return;
        const adm = String(r[COL_ADM]||'Unknown').trim() || 'Unknown';
        const won = isClosedWonValue(r[COL_CLOSED_WON]);
        if(won) return; const reason = String(r[COL_LOST_REASON]||'').trim();
        if(!reason) return;
        if(!topLostByAdm[adm]) topLostByAdm[adm] = {};
        topLostByAdm[adm][reason] = (topLostByAdm[adm][reason]||0)+1;
      });

      const inboundEntries = entries.filter(e=> isInboundAssigned(e.adm));
      const outboundEntries = entries.filter(e=> !isInboundAssigned(e.adm));

      const renderRows = (body, list) => {
        if(!body) return;
        if(list.length===0){ const tr=document.createElement('tr'); tr.innerHTML='<td colspan="8" class="px-6 py-4 text-center text-sm text-gray-500">No data for selection</td>'; body.appendChild(tr); return; }
        list.forEach((e,idx)=>{
          const tr=document.createElement('tr'); tr.className=idx%2===0?'bg-white':'bg-gray-50';
          const tmap = topLostByAdm[e.adm]||{};
          const topLR = Object.entries(tmap).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—';
          tr.innerHTML=`
            <td class="px-6 py-3 text-sm text-gray-900">${e.adm}</td>
            <td class="px-6 py-3 text-sm text-gray-900">${(e.inboundCount||0).toLocaleString()}</td>
            <td class="px-6 py-3 text-sm text-gray-900">${(e.outboundCount||0).toLocaleString()}</td>
            <td class="px-6 py-3 text-sm text-gray-900">${(e.won||0).toLocaleString()}</td>
            <td class="px-6 py-3 text-sm text-gray-900">${(e.totalOpps? (e.won/e.totalOpps)*100:0).toFixed(2)}%</td>
            <td class="px-6 py-3 text-sm text-gray-900">${toCurrency(e.avgWonInbound||0)} / ${toCurrency(e.avgWonOutbound||0)}</td>
            <td class="px-6 py-3 text-sm text-gray-900">${toCurrency(e.totalWonSum||0)}</td>
            <td class="px-6 py-3 text-sm text-gray-900">${topLR}</td>`;
          body.appendChild(tr);
        });
      };

      renderRows(inboundBody, inboundEntries);
      renderRows(outboundBody, outboundEntries);
    }
  }catch(e){ console.error('Error rendering OppData deeper insights', e); }
}

function renderOppCharts(rows){
  try{
    const COL_SOURCE=15, COL_SUBSOURCE=16, COL_ADM=17, COL_CREATED=8, COL_STAGE=3, COL_AMOUNT_PROJ=5, COL_AMOUNT=37;
    const isMarketingProspectScheduled = r => {
      const srcRaw = String(r?.[COL_SOURCE]||'').trim().toLowerCase();
      const subRaw = String(r?.[COL_SUBSOURCE]||'').trim().toLowerCase();
      return srcRaw.includes('marketing') && (subRaw.includes('prospect') && subRaw.includes('sched'));
    };
    const isAdmSourced = r => {
      if(!r) return false;
      const rName = String(r[COL_ADM]||'').trim();
      return rName !== '' || isMarketingProspectScheduled(r);
    };
    const datasetAll = (rows||[]).filter(isAdmSourced);
    console.log('[ADM] OppData ADM-sourced rows:', datasetAll.length);

    // Classify inbound/outbound using Opportunity Source (P)
    function classify(row){
      const srcRaw = String(row[COL_SOURCE]||'').trim().toLowerCase();
      const subRaw = String(row[COL_SUBSOURCE]||'').trim().toLowerCase();
      const adm = String(row[COL_ADM]||'').trim();
      const inboundRep = isInboundAssigned(adm);
      const isMarketing = srcRaw.includes('marketing');
      const isSales = srcRaw.includes('sales');
      const isProspectScheduled = subRaw.includes('prospect') && subRaw.includes('sched');
      if (isMarketing && isProspectScheduled) return 'MarketingInbound';
      if (isSales && inboundRep) return 'InboundADM';
      if (isMarketing || isSales) return 'Outbound';
      return 'Other';
    }

    // Apply date range and allowed years (using Created Date column)
    const start = admActivitiesDateFilter.startDate;
    const end = admActivitiesDateFilter.endDate;
    const dateFiltered = datasetAll.filter(r => {
      const d = parseSheetDate(r[COL_CREATED]);
      if (!d) return false;
      if (!inAllowedYears(d)) return false;
      const a = !start || d >= start;
      const b = !end || d <= end;
      return a && b;
    });

    // Apply team filter to OppData using P/Q/R classification
    const dataset = (admTeamFilter === 'all')
      ? dateFiltered
      : dateFiltered.filter(r => {
          const t = classify(r);
          return admTeamFilter === 'inbound' ? (t==='MarketingInbound' || t==='InboundADM') : t==='Outbound';
        });

    if (dataset.length === 0) {
      [oppInboundOutboundCountChart, oppAmountBySourceChart, oppByADMChart, oppStageChart, oppCreatedByMonthChart].forEach(ch=>{ if(ch){ ch.destroy(); }});
      oppInboundOutboundCountChart = oppAmountBySourceChart = oppByADMChart = oppStageChart = oppCreatedByMonthChart = null;

      // Build placeholder charts with titles
      const c1 = ensureOppCanvas('oppInboundOutboundCountChart');
      if(c1){ oppInboundOutboundCountChart = new Chart(c1.getContext('2d'),{ type:'bar', data:{ labels:['No Data'], datasets:[{ label:'Count', data:[0], backgroundColor:['rgba(148,163,184,0.6)'], borderRadius:6 }] }, options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, title:{display:true,text:'No data for selection'}, datalabels:{ display:false } }, scales:{ y:{ beginAtZero:true } } } }); }
      const c2 = ensureOppCanvas('oppAmountBySourceChart');
      if(c2){ oppAmountBySourceChart = new Chart(c2.getContext('2d'),{ type:'bar', data:{ labels:['No Data'], datasets:[{ label:'Amount', data:[0], backgroundColor:['rgba(148,163,184,0.6)'], borderRadius:6 }] }, options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, title:{display:true,text:'No data for selection'}, datalabels:{ display:false } }, scales:{ y:{ beginAtZero:true } } } }); }
      const c3 = ensureOppCanvas('oppByADMChart');
      if(c3){ oppByADMChart = new Chart(c3.getContext('2d'),{ type:'bar', data:{ labels:['No Data'], datasets:[{ label:'Count', data:[0], backgroundColor:['rgba(148,163,184,0.6)'] }] }, options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, title:{display:true,text:'No data for selection'}, datalabels:{ display:false } }, scales:{ y:{ beginAtZero:true } } } }); }
      const c4 = ensureOppCanvas('oppStageChart');
      if(c4){ oppStageChart = new Chart(c4.getContext('2d'),{ type:'bar', data:{ labels:['No Data'], datasets:[{ label:'Count', data:[0], backgroundColor:['rgba(148,163,184,0.6)'] }] }, options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, title:{display:true,text:'No data for selection'}, datalabels:{ display:false } } } }); }
      const c5 = ensureOppCanvas('oppCreatedByMonthChart');
      if(c5){ oppCreatedByMonthChart = new Chart(c5.getContext('2d'),{ type:'line', data:{ labels:[], datasets:[{ label:'No Data', data:[], borderColor:'rgba(148,163,184,0.9)' }] }, options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, title:{display:true,text:'No data for selection'}, datalabels:{ display:false } } } }); }
      // Also render deeper insights placeholders
      renderOppDeeperInsights([]);
      return;
    }

    // 1) Counts by category
    const counts = { MarketingInbound:0, InboundADM:0, Outbound:0, Other:0 };
    // 2) Amount by category (use Amount Projected if present else Amount)
    const amounts = { MarketingInbound:0, InboundADM:0, Outbound:0, Other:0 };
    // 3) Per-ADM counts split by category
    const perAdm = {};
    // 4) Stage distribution split
    const perStage = {};
    // 5) Created by month counts
    const byMonth = { MarketingInbound:{}, InboundADM:{}, Outbound:{}, Other:{} };

    dataset.forEach(r=>{
      const cls = classify(r);
      counts[cls] = (counts[cls]||0)+1;
      const amtProj = parseSheetNumber(r[COL_AMOUNT_PROJ]);
      const amt = parseSheetNumber(r[COL_AMOUNT]);
      amounts[cls] = (amounts[cls]||0) + (amtProj>0? amtProj : amt);

      const admName = String(r[COL_ADM]||'Unknown').trim() || 'Unknown';
      if(!perAdm[admName]) perAdm[admName] = { MarketingInbound:0, InboundADM:0, Outbound:0, Other:0, total:0 };
      perAdm[admName][cls] += 1; perAdm[admName].total += 1;

      const stage = String(r[COL_STAGE]||'Unknown').trim() || 'Unknown';
      if(!perStage[stage]) perStage[stage] = { MarketingInbound:0, InboundADM:0, Outbound:0, Other:0, total:0 };
      perStage[stage][cls] += 1; perStage[stage].total += 1;

      const cd = parseSheetDate(r[COL_CREATED]);
      const d = cd ? new Date(cd.getFullYear(), cd.getMonth(), 1) : null;
      const key = d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}` : 'Unknown';
      byMonth[cls][key] = (byMonth[cls][key]||0)+1;
    });

    // Even if totals are zero, we will still draw placeholder charts with a 'No data' title

    // Destroy existing opp charts
    [oppInboundOutboundCountChart, oppAmountBySourceChart, oppByADMChart, oppStageChart, oppCreatedByMonthChart].forEach(ch=>{ if(ch){ ch.destroy(); }});
    oppInboundOutboundCountChart = oppAmountBySourceChart = oppByADMChart = oppStageChart = oppCreatedByMonthChart = null;

    // Build 1) Opp count by source (three categories)
    const c1 = ensureOppCanvas('oppInboundOutboundCountChart');
    if(c1){
      const ctx=c1.getContext('2d');
      const labels=['Marketing Inbound','Inbound ADM','Outbound'];
      const data=[counts.MarketingInbound||0, counts.InboundADM||0, counts.Outbound||0];
      const cTitle = (data[0]+data[1]+data[2])>0 ? 'Opp Count by Source' : 'No data for selection';
      oppInboundOutboundCountChart = new Chart(ctx,{
        type:'bar', data:{ labels, datasets:[{ label:'Count', data, backgroundColor:['rgba(16,185,129,0.8)','rgba(34,197,94,0.8)','rgba(59,130,246,0.8)'], borderRadius:6 }] },
        options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, title:{display:true,text:cTitle}, datalabels:{ color:'#111', font:{weight:'600',size:11}, formatter:v=>v||'' } }, scales:{ y:{ beginAtZero:true } } }
      });
    }

    // 2) Amount by source (three categories)
    const c2 = ensureOppCanvas('oppAmountBySourceChart');
    if(c2){
      const ctx=c2.getContext('2d');
      const labels=['Marketing Inbound','Inbound ADM','Outbound'];
      const data=[amounts.MarketingInbound||0, amounts.InboundADM||0, amounts.Outbound||0];
      const aTitle = (data[0]+data[1]+data[2])>0 ? 'Amount by Source' : 'No data for selection';
      oppAmountBySourceChart = new Chart(ctx,{
        type:'bar', data:{ labels, datasets:[{ label:'Amount Projected', data, backgroundColor:['rgba(16,185,129,0.8)','rgba(34,197,94,0.8)','rgba(59,130,246,0.8)'], borderRadius:6 }] },
        options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, title:{display:true,text:aTitle}, datalabels:{ color:'#111', font:{weight:'600',size:11}, formatter:v=>new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',maximumFractionDigits:0}).format(v||0) } }, scales:{ y:{ beginAtZero:true } } }
      });
    }

    // 3) Per-ADM stacked counts (top 8 by total) with 3 categories
    const c3 = ensureOppCanvas('oppByADMChart');
    if(c3){
      const ctx=c3.getContext('2d');
      const entries = Object.entries(perAdm).sort((a,b)=>b[1].total - a[1].total).slice(0,8);
      const labels = entries.length? entries.map(e=>e[0]) : ['No Data'];
      const mInboundData = entries.length? entries.map(([,v])=>v.MarketingInbound||0) : [0];
      const iAdmData = entries.length? entries.map(([,v])=>v.InboundADM||0) : [0];
      const outboundData = entries.length? entries.map(([,v])=>v.Outbound||0) : [0];
      const pTitle = entries.length? 'ADM-Sourced Opps by ADM (Top 8)' : 'No data for selection';
      oppByADMChart = new Chart(ctx,{
        type:'bar', data:{ labels, datasets:[
          { label:'Marketing Inbound', data: mInboundData, backgroundColor:'rgba(16,185,129,0.8)' },
          { label:'Inbound ADM', data: iAdmData, backgroundColor:'rgba(34,197,94,0.8)' },
          { label:'Outbound', data: outboundData, backgroundColor:'rgba(59,130,246,0.8)' }
        ] },
        options:{ responsive:true, maintainAspectRatio:false, plugins:{ title:{display:true,text:pTitle}, datalabels:{ color:'#111', anchor:'end', align:'top', font:{size:10,weight:'600'}, formatter:(v)=>v||'' } }, scales:{ x:{ stacked:true }, y:{ stacked:true, beginAtZero:true } } }
      });
    }

    // 4) Stage distribution stacked by source (top 10 stages, 3 categories)
    const c4 = ensureOppCanvas('oppStageChart');
    if(c4){
      const ctx=c4.getContext('2d');
      const stages = Object.entries(perStage).sort((a,b)=>b[1].total - a[1].total).slice(0,10);
      const labels = stages.length? stages.map(e=>e[0]) : ['No Data'];
      const mInboundData = stages.length? stages.map(([,v])=>v.MarketingInbound||0) : [0];
      const iAdmData = stages.length? stages.map(([,v])=>v.InboundADM||0) : [0];
      const outboundData = stages.length? stages.map(([,v])=>v.Outbound||0) : [0];
      const sTitle = stages.length? 'Opp Stage Distribution (ADM-sourced)' : 'No data for selection';
      oppStageChart = new Chart(ctx,{
        type:'bar', data:{ labels, datasets:[
          { label:'Marketing Inbound', data: mInboundData, backgroundColor:'rgba(16,185,129,0.8)' },
          { label:'Inbound ADM', data: iAdmData, backgroundColor:'rgba(34,197,94,0.8)' },
          { label:'Outbound', data: outboundData, backgroundColor:'rgba(59,130,246,0.8)' }
        ] },
        options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, plugins:{ title:{display:true,text:sTitle}, datalabels:{ color:'#111', font:{size:10,weight:'600'}, formatter:(v)=>v||'' } }, scales:{ x:{ stacked:true, beginAtZero:true }, y:{ stacked:true } } }
      });
    }

    // 5) Created by month (3 categories)
    const c5 = ensureOppCanvas('oppCreatedByMonthChart');
    if(c5){
      const ctx=c5.getContext('2d');
      const months = Array.from(new Set([ ...Object.keys(byMonth.MarketingInbound), ...Object.keys(byMonth.InboundADM), ...Object.keys(byMonth.Outbound) ])).filter(k=>k!=='Unknown').sort();
      const mInboundData = months.length? months.map(m=> (byMonth.MarketingInbound[m]||0)) : [];
      const iAdmData = months.length? months.map(m=> (byMonth.InboundADM[m]||0)) : [];
      const outboundData = months.length? months.map(m=> (byMonth.Outbound[m]||0)) : [];
      const mTitle = months.length? 'Opps Created by Month (ADM-sourced)' : 'No data for selection';
      oppCreatedByMonthChart = new Chart(ctx,{
        type:'line', data:{ labels: months, datasets:[
          { label:'Marketing Inbound', data: mInboundData, borderColor:'rgba(16,185,129,1)', backgroundColor:'rgba(16,185,129,0.2)', tension:0.25, fill:true },
          { label:'Inbound ADM', data: iAdmData, borderColor:'rgba(34,197,94,1)', backgroundColor:'rgba(34,197,94,0.2)', tension:0.25, fill:true },
          { label:'Outbound', data: outboundData, borderColor:'rgba(59,130,246,1)', backgroundColor:'rgba(59,130,246,0.2)', tension:0.25, fill:true }
        ] },
        options:{ responsive:true, maintainAspectRatio:false, plugins:{ title:{display:true,text:mTitle}, datalabels:{ display:false } } }
      });
    }
    // Render deeper insights with the same filtered dataset
    renderOppDeeperInsights(dataset);
  }catch(e){
    console.error('Error rendering OppData charts', e);
  }
}

function getThisWeekRange(){
  const now=new Date();
  const day=now.getDay();
  const diff=(day+6)%7;
  const start=new Date(now.getFullYear(), now.getMonth(), now.getDate()-diff);
  const end=new Date(start.getFullYear(), start.getMonth(), start.getDate()+6);
  return {start,end};
}

function getLastWeekRange(){
  const w = getThisWeekRange();
  const start = new Date(w.start); start.setDate(start.getDate()-7);
  const end = new Date(w.end); end.setDate(end.getDate()-7);
  return {start,end};
}

function renderThisWeekSection(){
  try{
    const ids={
      pipeTotal:document.getElementById('twPipelineTotal'),
      pipeComp:document.getElementById('twPipelineComparison'),
      pipeIn:document.getElementById('twPipelineInbound'),
      pipeOut:document.getElementById('twPipelineOutbound'),
      oppTotal:document.getElementById('twOppsTotal'),
      oppComp:document.getElementById('twOppsComparison'),
      oppIn:document.getElementById('twOppsInbound'),
      oppOut:document.getElementById('twOppsOutbound'),
      table:document.getElementById('thisWeekTeamHighlightsBody')
    };
    if(!ids.pipeTotal||!ids.pipeIn||!ids.pipeOut||!ids.oppTotal||!ids.oppIn||!ids.oppOut||!ids.table) return;
    
    const w=getThisWeekRange();
    const lw=getLastWeekRange();
    
    // Adjust Last Week to match current week's progress (Same Point in Time)
    const now = new Date();
    const daysIntoWeek = (now.getDay() + 6) % 7 + 1;
    
    const lwEndAdjusted = new Date(lw.start);
    lwEndAdjusted.setDate(lwEndAdjusted.getDate() + daysIntoWeek - 1);
    lwEndAdjusted.setHours(23,59,59,999);

    const lwAdjusted = { start: lw.start, end: lwEndAdjusted };
    
    const COL_SOURCE=15, COL_SUBSOURCE=16, COL_ADM=17, COL_CREATED=8, COL_AMOUNT_PROJ=5;
    const isMarketingProspectScheduled = r => {
      const srcRaw=String(r?.[COL_SOURCE]||'').trim().toLowerCase();
      const subRaw=String(r?.[COL_SUBSOURCE]||'').trim().toLowerCase();
      return srcRaw.includes('marketing') && (subRaw.includes('prospect') && subRaw.includes('sched'));
    };
    const isAdmSourced = r => {
      if(!r) return false;
      const rName=String(r[COL_ADM]||'').trim();
      return rName!=='' || isMarketingProspectScheduled(r);
    };
    const classify = (row)=>{
      const srcRaw=String(row[COL_SOURCE]||'').trim().toLowerCase();
      const subRaw=String(row[COL_SUBSOURCE]||'').trim().toLowerCase();
      const adm=String(row[COL_ADM]||'').trim();
      const inboundRep=isInboundAssigned(adm);
      const isMarketing=srcRaw.includes('marketing');
      const isSales=srcRaw.includes('sales');
      const isProspectScheduled=subRaw.includes('prospect') && subRaw.includes('sched');
      if(isMarketing && isProspectScheduled) return 'MarketingInbound';
      if(isSales && inboundRep) return 'InboundADM';
      if(isMarketing || isSales) return 'Outbound';
      return 'Other';
    };

    const filterByTeam = (r) => {
      const t = classify(r);
      if(admTeamFilter==='all') return true;
      return admTeamFilter==='inbound' ? (t==='MarketingInbound' || t==='InboundADM') : t==='Outbound';
    };

    const allOpps = (admOppData||[]).filter(isAdmSourced);

    // Calculate metrics for a given date range
    const calcRangeMetrics = (range) => {
      const rows = allOpps.filter(r=>{
        const d=parseSheetDate(r[COL_CREATED]); if(!d) return false; if(!inAllowedYears(d)) return false; 
        return d>=range.start && d<=range.end;
      }).filter(filterByTeam);

      let oppTotal=0, oppIn=0, oppOut=0, pipeTotal=0, pipeIn=0, pipeOut=0;
      const perUser={};

      rows.forEach(r=>{
        const cls=classify(r);
        const amtProj=parseSheetNumber(r[COL_AMOUNT_PROJ]);
        const val=amtProj>0? amtProj : 0;
        const admName=String(r[COL_ADM]||'Unknown').trim()||'Unknown';
        
        oppTotal++;
        pipeTotal+=val;
        
        if(!perUser[admName]) perUser[admName]={oppIn:0,oppOut:0,pipeIn:0,pipeOut:0};

        if(cls==='MarketingInbound' || cls==='InboundADM'){ 
          oppIn++; pipeIn+=val; 
          perUser[admName].oppIn++; perUser[admName].pipeIn+=val; 
        }
        else if(cls==='Outbound'){ 
          oppOut++; pipeOut+=val; 
          perUser[admName].oppOut++; perUser[admName].pipeOut+=val; 
        }
      });
      return { oppTotal, oppIn, oppOut, pipeTotal, pipeIn, pipeOut, perUser };
    };

    const curr = calcRangeMetrics(w);
    const prev = calcRangeMetrics(lwAdjusted);

    // Update Main Metrics
    ids.pipeTotal.textContent=toCurrency(curr.pipeTotal);
    ids.pipeIn.textContent=toCurrency(curr.pipeIn);
    ids.pipeOut.textContent=toCurrency(curr.pipeOut);
    
    ids.oppTotal.textContent=curr.oppTotal.toLocaleString();
    ids.oppIn.textContent=curr.oppIn.toLocaleString();
    ids.oppOut.textContent=curr.oppOut.toLocaleString();

    // Update Comparisons
    const updateComp = (el, current, previous, isCurrency) => {
        if(!el) return;
        const diff = current - previous;
        const format = (v) => isCurrency ? toCurrency(v) : Math.abs(v).toLocaleString();
        const arrow = diff > 0 ? '▲' : (diff < 0 ? '▼' : '—');
        const color = diff > 0 ? 'text-green-600' : (diff < 0 ? 'text-red-600' : 'text-gray-500');
        
        let text = '';
        if(diff === 0) text = 'No change vs last week (same day)';
        else text = `${arrow} ${format(Math.abs(diff))} vs last week (same day)`;
        
        el.className = `text-xs font-medium mb-2 ${color}`;
        el.textContent = text;
    };

    updateComp(ids.pipeComp, curr.pipeTotal, prev.pipeTotal, true);
    updateComp(ids.oppComp, curr.oppTotal, prev.oppTotal, false);

    // Update Team Highlights Table (Only for This Week)
    // We also need Activity data for This Week to populate the table
    const perAdm = curr.perUser;
    
    const aRows=(admActivitiesData||[]);
    const COL_A_DATE=17, COL_A_ASSIGNED=1, COL_A_ROLE=2, COL_A_TYPE=7, COL_A_CALL_RESULT=9;
    
    const actRows=aRows.filter(r=>{
      if(String(r[COL_A_ROLE]||'').toLowerCase().trim()!=='adm') return false;
      const d=parseSheetDate(r[COL_A_DATE]); if(!d) return false; 
      // Check date range
      const inRange = d>=w.start && d<=w.end;
      if(!inRange) return false;
      
      if(!inAllowedYears(d)) return false;
      
      if(admTeamFilter==='all') return true;
      const assigned = r[COL_A_ASSIGNED];
      const inbound = isInboundAssigned(assigned);
      return admTeamFilter==='inbound' ? inbound : !inbound;
    });
    if(actRows.length > 0) console.log('[ThisWeek] Sample Activity:', actRows[0]);

    actRows.forEach(r=>{
      const user=String(r[COL_A_ASSIGNED]||'Unknown').trim()||'Unknown';
      const typeRaw=String(r[COL_A_TYPE]||'').toLowerCase();
      const isCall=typeRaw.includes('call')||typeRaw.includes('phone');
      const isEmail=typeRaw.includes('email')||typeRaw.includes('e-mail');
      const callRes=String(r[COL_A_CALL_RESULT]||'').toLowerCase();
      
      if(!perAdm[user]) perAdm[user]={oppIn:0,oppOut:0,pipeIn:0,pipeOut:0, calls:0, emails:0, demos:0}; // Initialize if not present from opps
      if(!perAdm[user].calls) perAdm[user].calls=0;
      if(!perAdm[user].emails) perAdm[user].emails=0;
      if(!perAdm[user].demos) perAdm[user].demos=0;

      if(isCall) perAdm[user].calls++;
      if(isEmail) perAdm[user].emails++;
      if(callRes==='demo set' || callRes.includes('demo')) perAdm[user].demos++;
    });

    ids.table.innerHTML='';
    const entries=Object.entries(perAdm).map(([u,v])=>({ user:u, ...v, totalOpp:(v.oppIn||0)+(v.oppOut||0), totalAct:(v.calls||0)+(v.emails||0) }))
      .sort((a,b)=> (b.totalOpp - a.totalOpp) || (b.totalAct - a.totalAct));
      
    if(entries.length===0){
      const tr=document.createElement('tr'); tr.innerHTML='<td colspan="8" class="px-6 py-4 text-center text-sm text-gray-500">No data for this week</td>'; ids.table.appendChild(tr);
    }else{
      entries.forEach((e,idx)=>{
        const tr=document.createElement('tr'); tr.className=idx%2===0?'bg-white':'bg-gray-50';
        tr.innerHTML=`
          <td class="px-6 py-3 text-sm text-gray-900">${e.user}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${(e.calls||0).toLocaleString()}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${(e.emails||0).toLocaleString()}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${(e.demos||0).toLocaleString()}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${(e.oppIn||0).toLocaleString()}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${(e.oppOut||0).toLocaleString()}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${toCurrency(e.pipeIn||0)}</td>
          <td class="px-6 py-3 text-sm text-gray-900">${toCurrency(e.pipeOut||0)}</td>`;
        ids.table.appendChild(tr);
      });
    }
  }catch(e){ console.error('Error rendering This Week section', e); }
}
