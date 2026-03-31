const S={
  excelFiles:[],datFiles:[],shiftMap:{},
  records:[],filtered:[],
  sortCol:'date',sortDir:1,
  page:1,perPage:100,
  lateRecs:[],absentRecs:[],earlyRecs:[],
  holidays:[
    {name:"Makarsankranti",d:"2026-01-14",b:["Gujarat"]},
    {name:"Republic Day",d:"2026-01-26",b:["Mumbai","Borivali","Nagpur","Gujarat","Goa"]},
    {name:"Rang Panchami",d:"2026-03-03",b:["Mumbai","Borivali","Nagpur","Gujarat","Goa"]},
    {name:"Gudi Padwa",d:"2026-03-19",b:["Mumbai","Borivali","Nagpur"]},
    {name:"Maharashtra/Gujarat Day",d:"2026-05-01",b:["Mumbai","Borivali","Nagpur","Gujarat"]},
    {name:"Independence Day",d:"2026-08-15",b:["Mumbai","Borivali","Nagpur","Gujarat","Goa"]},
    {name:"Rakshabandhan",d:"2026-08-28",b:["Mumbai","Borivali","Nagpur","Gujarat"]},
    {name:"Gokulashtami",d:"2026-09-04",b:["Mumbai","Borivali","Nagpur"]},
    {name:"Ganesh Chaturthi",d:"2026-09-14",d2:"2026-09-15",b:["Mumbai","Borivali","Nagpur","Gujarat","Goa"]},
    {name:"Anant Chaturthi",d:"2026-09-25",b:["Mumbai","Borivali","Nagpur"]},
    {name:"Gandhi Jayanti",d:"2026-10-02",b:["Mumbai","Borivali","Nagpur","Gujarat","Goa"]},
    {name:"Dussehra",d:"2026-10-20",b:["Mumbai","Borivali","Nagpur","Gujarat","Goa"]},
    {name:"Jinharsh Diwali Celebration",d:"2026-11-07",b:["Mumbai","Borivali","Nagpur","Gujarat","Goa"]},
    {name:"Narak Chaturdashi/Laxmi Pujan",d:"2026-11-08",b:["Mumbai","Borivali","Gujarat"]},
    {name:"Diwali",d:"2026-11-09",b:["Mumbai","Borivali","Nagpur","Gujarat","Goa"]},
    {name:"Gujarati New Year / Padwa",d:"2026-11-10",b:["Mumbai","Borivali","Nagpur","Gujarat"]},
    {name:"Bhai Duj",d:"2026-11-11",b:["Mumbai","Borivali","Nagpur","Gujarat"]},
    {name:"Diwali Holidays",d:"2026-11-12",d2:"2026-11-14",b:["Gujarat"]},
    {name:"Goa Liberation Day",d:"2026-12-19",b:["Goa"]},
    {name:"Christmas",d:"2026-12-25",b:["Goa"]}
  ],
  failureDates:[]
};

function syncStickyLayout(){
  const root=document.documentElement;
  const filterBar=document.querySelector('#tab-daily .filter-bar');
  const desktop=window.innerWidth>768;
  const topbarOffset=desktop?0:0;
  const filterHeight=desktop&&filterBar?Math.ceil(filterBar.getBoundingClientRect().height):0;
  root.style.setProperty('--sticky-panel-top',topbarOffset+'px');
  root.style.setProperty('--sticky-filter-height',filterHeight+'px');
}

function queueStickyLayoutSync(){
  window.requestAnimationFrame(syncStickyLayout);
}

function syncAppMode(){
  const uploadSection=document.getElementById('upload-section');
  const isReport=uploadSection && uploadSection.style.display==='none';
  document.body.dataset.mode=isReport?'report':'start';
}

function getShiftDurationMinutes(shiftStart,shiftEnd){
  if(shiftStart==null || shiftEnd==null)return 480;
  return shiftEnd>=shiftStart ? shiftEnd-shiftStart : (24*60-shiftStart)+shiftEnd;
}

function getLateMinutes(shiftStart,inM){
  if(shiftStart==null || inM==null)return 0;
  return Math.max(0,inM-shiftStart);
}

function getEarlyMinutes(shiftStart,shiftEnd,outM){
  if(shiftEnd==null || outM==null)return 0;
  if(shiftStart!=null && shiftEnd<shiftStart){
    const normalizedOut=outM<shiftStart ? outM+1440 : outM;
    return Math.max(0,(shiftEnd+1440)-normalizedOut);
  }
  return Math.max(0,shiftEnd-outM);
}

function getOvertimeMinutes(shiftStart,shiftEnd,outM,workedMins){
  if(outM==null)return 0;
  if(shiftStart==null || shiftEnd==null){
    return Math.max(0,workedMins-480);
  }
  if(shiftEnd<shiftStart){
    const normalizedOut=outM<shiftStart ? outM+1440 : outM;
    return Math.max(0,normalizedOut-(shiftEnd+1440));
  }
  return Math.max(0,outM-shiftEnd);
}

window.onload=()=>{
  if(localStorage.getItem('theme')==='dark'){
    document.documentElement.setAttribute('data-theme','dark');
    document.getElementById('theme-btn').textContent='Light';
  }
  const saved=localStorage.getItem('hr_att_v3');
  if(saved){
    try{
      S.records=JSON.parse(saved);
      document.getElementById('upload-section').style.display='none';
      document.getElementById('btn-gen').style.display='none';
      document.getElementById('btn-clear').style.display='flex';
      buildReport();
    }catch(e){localStorage.removeItem('hr_att_v3')}
  }
  syncAppMode();
  queueStickyLayoutSync();
};

window.addEventListener('resize',queueStickyLayoutSync);

function toggleTheme(){
  const d=document.documentElement.getAttribute('data-theme')==='dark';
  document.documentElement.setAttribute('data-theme',d?'light':'dark');
  document.getElementById('theme-btn').textContent=d?'Dark':'Light';
  localStorage.setItem('theme',d?'light':'dark');
  syncAppMode();
  queueStickyLayoutSync();
}
function clearData(){
  if(!confirm('Clear all data and restart?'))return;
  localStorage.removeItem('hr_att_v3');location.reload();
}

function handleFiles(files,type){
  (type==='excel'?S.excelFiles:S.datFiles).push(...files);
  renderPills(type);checkReady();
}
function dzDrag(e,id){e.preventDefault();document.getElementById(id).classList.add('drag-over')}
function dzLeave(id){document.getElementById(id).classList.remove('drag-over')}
function dzDrop(e,type){
  e.preventDefault();
  document.getElementById(type==='excel'?'dz-excel':'dz-dat').classList.remove('drag-over');
  handleFiles(e.dataTransfer.files,type);
}
function removeFile(type,idx){
  (type==='excel'?S.excelFiles:S.datFiles).splice(idx,1);
  renderPills(type);checkReady();
}
function renderPills(type){
  const arr=type==='excel'?S.excelFiles:S.datFiles;
  document.getElementById('fl-'+type).innerHTML=arr.map((f,i)=>
    `<div class="pill"><span>${f.name}</span><button class="pill-rm" onclick="removeFile('${type}',${i})">x</button></div>`
  ).join('');
}
function checkReady(){
  document.getElementById('btn-gen').disabled=!(S.excelFiles.length&&S.datFiles.length);
}

function setProg(pct,msg){
  document.getElementById('prog-wrap').style.display='block';
  document.getElementById('prog-fill').style.width=pct+'%';
  document.getElementById('prog-pct').textContent=pct+'%';
  document.getElementById('prog-msg').textContent=msg;
}
function hideProg(){document.getElementById('prog-wrap').style.display='none'}
function showToast(msg){
  const t=document.getElementById('toast');
  document.getElementById('toast-msg').textContent=msg;
  t.style.display='flex';setTimeout(()=>t.style.display='none',5000);
}

async function parseShift(file){
  return new Promise((res,rej)=>{
    const r=new FileReader();
    r.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:'array'});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1});
        const map={};let hi=0;
        for(let i=0;i<Math.min(6,rows.length);i++){
          const r2=rows[i].map(c=>String(c||'').toLowerCase());
          if(r2.some(c=>c.includes('userid')||c.includes('user id')||c.includes('emp'))){hi=i;break}
        }
        const hdr=rows[hi].map(c=>String(c||'').toLowerCase().trim());
        const col=k=>hdr.findIndex(h=>h.includes(k));
        const idC=col('userid')!==-1?col('userid'):col('user')!==-1?col('user'):col('emp');
        const nmC=col('particular')!==-1?col('particular'):col('name');
        const brC=col('branch'),dpC=col('department')!==-1?col('department'):col('dept');
        const stC=col('shift start')!==-1?col('shift start'):col('start');
        const enC=col('shift end')!==-1?col('shift end'):col('end');
        const toMins=v=>{
          if(v==null)return null;
          if(v instanceof Date)return v.getHours()*60+v.getMinutes();
          if(typeof v==='number')return Math.round(v*24*60);
          const m=String(v).match(/(\d+):(\d+)/);
          return m?parseInt(m[1])*60+parseInt(m[2]):null;
        };
        for(let i=hi+1;i<rows.length;i++){
          const row=rows[i];const uid=row[idC];
          if(!uid&&!row[nmC])continue;
          if(uid)map[String(uid).trim()]={
            name:row[nmC]?String(row[nmC]).trim():'User '+uid,
            branch:row[brC]?String(row[brC]).trim():'',
            department:row[dpC]?String(row[dpC]).trim():'',
            shiftStart:toMins(row[stC]),shiftEnd:toMins(row[enC])
          };
        }
        res(map);
      }catch(err){rej(err)}
    };
    r.readAsArrayBuffer(file);
  });
}

async function parseDat(file){
  return new Promise((res,rej)=>{
    const r=new FileReader();
    r.onload=e=>{
      const lines=e.target.result.split(/\r?\n/);const ps=[];
      for(const line of lines){
        if(!line.trim())continue;
        const pts=line.trim().split(/\s+/);if(pts.length<2)continue;
        const ds=(pts[1]&&pts[2]&&!pts[1].includes(' '))?pts[1]+' '+pts[2]:pts[1];
        const d=new Date(ds);
        if(!isNaN(d.getTime()))ps.push({uid:String(pts[0]).trim(),dt:d});
      }
      res(ps);
    };
    r.onerror=rej;
    r.readAsText(file);
  });
}

async function generateReport(){
  document.getElementById('toast').style.display='none';
  document.getElementById('btn-gen').disabled=true;
  try{
    setProg(10,'Reading shift master...');
    S.shiftMap={};
    for(const f of S.excelFiles)Object.assign(S.shiftMap,await parseShift(f));
    setProg(30,'Parsing attendance logs...');
    let punches=[];
    for(const f of S.datFiles)punches.push(...await parseDat(f));
    if(!punches.length)throw new Error('No valid punch records found in the uploaded files.');
    setProg(55,'Calculating attendance...');
    const grouped={};
    for(const p of punches){
      const dk=p.dt.toISOString().split('T')[0];
      const key=p.uid+'|'+dk;
      if(!grouped[key])grouped[key]={uid:p.uid,date:dk,punches:[]};
      grouped[key].punches.push(p.dt);
    }
    const allPunchedDates=punches.map(p=>p.dt.getTime());
    let minD=new Date(Math.min(...allPunchedDates)), maxD=new Date(Math.max(...allPunchedDates));
    
    // SAFETY GUARD: If logs span too many years, limit "Full Calendar Filling" to current month only.
    if((maxD.getTime()-minD.getTime()) > (62*24*60*60*1000)){
       minD=new Date(maxD.getFullYear(), maxD.getMonth(), 1);
    }
    
    const dts=[];
    for(let d=new Date(minD); d<=maxD; d.setDate(d.getDate()+1)){
      dts.push(d.toISOString().split('T')[0]);
    }
    const uids=[...new Set([...Object.keys(S.shiftMap), ...punches.map(p=>p.uid)])];
    const pad2=n=>String(n).padStart(2,'0');
    const m2t=m=>m==null||isNaN(m)?'--':pad2(Math.floor(m/60))+':'+pad2(m%60);
    const fmtD=m=>m<=0?'--':(Math.floor(m/60)?Math.floor(m/60)+'h ':'')+((m%60)?m%60+'m':'');
    
    S.records=[];
    for(const dt of dts){
      for(const uid of uids){
        const k=uid+'|'+dt;
        const g=grouped[k];
        const info=S.shiftMap[uid]||{name:'User '+uid,branch:'',department:'',shiftStart:null,shiftEnd:null};
        const ps=g?g.punches.sort((a,b)=>a-b):[];
        
        const dy=new Date(dt+'T12:00:00').toLocaleDateString('en-US',{weekday:'short'});
        const sd=(info.shiftStart!==null&&info.shiftEnd!==null)?m2t(info.shiftStart)+' - '+m2t(info.shiftEnd):'--';
        
        let first=null,last=null,hrs=0,inM=null,outM=null,status='Absent',lateMins=0,earlyMins=0,otMins=0;
        const sDur=getShiftDurationMinutes(info.shiftStart,info.shiftEnd);
        
        const isHol=S.holidays.find(h=>{
          const bMatch=h.b.some(b=>info.branch.includes(b));
          if(!bMatch)return false;
          if(h.d2)return dt>=h.d&&dt<=h.d2;
          return dt===h.d;
        });
        
        if(ps.length===0){
          if(isHol)status='Holiday';
          else if(dy==='Sun')status='Week Off';
        }

        if(ps.length>0){
          first=ps[0];last=ps[ps.length-1];
          hrs=(last-first)/3600000;
          const workedMins=Math.max(0,Math.round((last-first)/60000));
          inM=first.getHours()*60+first.getMinutes();
          outM=last.getHours()*60+last.getMinutes();
          status='Present';
          lateMins=getLateMinutes(info.shiftStart,inM);
          if(lateMins>15)status='Late';
          earlyMins=getEarlyMinutes(info.shiftStart,info.shiftEnd,outM);
          if(hrs<4.5)status='Half Day';
          if(hrs<0.25 || ps.length===1)status='Missed Punch';
          if(status!=='Missed Punch'){
            // Overtime should reflect time worked beyond the scheduled shift end.
            otMins=getOvertimeMinutes(info.shiftStart,info.shiftEnd,outM,workedMins);
          }
        }
        
        // Gap: For Absent it's full shift, for WeekOff/Holiday it's 0, for Present it's Dur - Actual
        let gapMins=0;
        if(status==='Absent')gapMins=sDur;
        else if(status==='Present'||status==='Late'||status==='Late (Comp)'||status==='Half Day')gapMins=sDur-Math.round(hrs*60);
        
        if(status==='Late'&&gapMins<=0)status='Late (Comp)';
        
        const gapFmt=gapMins===0?'0m':fmtD(Math.abs(gapMins));
        const gapClass=gapMins<=0?'g-ok':'g-err';
        
        S.records.push({
          uid,name:info.name,branch:info.branch,department:info.department,
          date:dt,day:dy,shiftDisplay:sd,
          firstIn:ps.length?m2t(inM):'--',lastOut:ps.length?m2t(outM):'--',
          hoursWorked:Math.round(hrs*100)/100,status,
          lateMins,earlyMins,lateBy:fmtD(lateMins),earlyBy:fmtD(earlyMins),
          otMins,overtime:fmtD(otMins),punchCount:ps.length,
          gapMins,gapFmt,gapClass
        });
      }
    }
    setProg(90,'Saving session...');
    try{localStorage.setItem('hr_att_v3',JSON.stringify(S.records))}catch(e){}
    setProg(100,`Done - ${S.records.length} records processed.`);
    setTimeout(hideProg,1200);
    document.getElementById('upload-section').style.display='none';
    document.getElementById('btn-gen').style.display='none';
    document.getElementById('btn-clear').style.display='flex';
    buildReport();
    syncAppMode();
    queueStickyLayoutSync();
  }catch(err){
    showToast(err.message||'Processing failed.');
    console.error(err);
    document.getElementById('btn-gen').disabled=false;
    hideProg();
  }
}

function buildReport(){
  detectMachineFailures();
  const uniq=k=>[...new Set(S.records.map(r=>r[k]).filter(Boolean))].sort();
  const fill=(id,vals)=>{
    const el=document.getElementById(id);
    el.innerHTML='<option value="">'+el.options[0].text+'</option>'+vals.map(v=>`<option value="${v}">${v}</option>`).join('');
  };
  fill('f-branch',uniq('branch'));fill('f-dept',uniq('department'));
  const emps=[...new Set(S.records.map(r=>`${r.uid}|${r.name}`))].sort().map(e=>{const[uid,nm]=e.split('|');return{uid,name:nm}});
  document.getElementById('f-emp').innerHTML='<option value="">All Employees</option>'+emps.map(e=>`<option value="${e.uid}">${e.name} (${e.uid})</option>`).join('');
  
  S.lateRecs=S.records.filter(r=>r.lateMins>0);
  S.absentRecs=S.records.filter(r=>r.status==='Absent');
  S.earlyRecs=S.records.filter(r=>r.earlyMins>0);
  document.getElementById('dl-wrap').style.display='flex';
  document.getElementById('stats-row').style.display='grid';
  document.getElementById('tabs-row').style.display='flex';
  document.getElementById('tab-daily').style.display='block';
  document.getElementById('tb-daily').textContent=S.records.length;
  document.getElementById('tb-summary').textContent=[...new Set(S.records.map(r=>r.uid))].size;
  document.getElementById('tb-late').textContent=S.lateRecs.length;
  document.getElementById('tb-absent').textContent=S.absentRecs.length;
  document.getElementById('tb-early').textContent=S.earlyRecs.length;
  S.filtered=[...S.records];S.page=1;
  renderStats(S.records);renderTable();renderSubTables();renderInsights();
  syncAppMode();
  queueStickyLayoutSync();
}

function detectMachineFailures(){
  const r=S.records; if(!r.length)return;

  // Ã¢â€â‚¬Ã¢â€â‚¬ Helper: is this date a holiday for this branch? Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
  function isHolidayForBranch(date, branch){
    return S.holidays.some(h=>{
      const bMatch=h.b.some(b=>(branch||'').includes(b));
      if(!bMatch)return false;
      if(h.d2)return date>=h.d&&date<=h.d2;
      return date===h.d;
    });
  }

  // Ã¢â€â‚¬Ã¢â€â‚¬ Step 1: Group by date|branch, tracking absent + punched counts Ã¢â€â‚¬
  const grouped={};
  r.forEach(x=>{
    const bk=x.date+'|'+(x.branch||'Default');
    if(!grouped[bk])grouped[bk]={
      total:0, absent:0, punched:0,
      date:x.date, branch:x.branch
    };
    grouped[bk].total++;
    if(x.status==='Absent') grouped[bk].absent++;
    // "punched" = anyone who actually had at least 1 swipe (not Absent/Holiday/WeekOff)
    if(!['Absent','Week Off','Holiday','System Error'].includes(x.status)) grouped[bk].punched++;
  });

  // Ã¢â€â‚¬Ã¢â€â‚¬ Step 2: Evaluate each branch-day for anomaly Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
  S.failureDates=[];
  Object.keys(grouped).forEach(bk=>{
    const g=grouped[bk];
    const dy=new Date(g.date+'T12:00:00').toLocaleDateString('en-US',{weekday:'short'});

    // Guard 1: Skip weekends (Sun & Sat handled by Week Off status, but check Sunday explicitly)
    if(dy==='Sun') return;

    // Guard 2: Skip recognized holidays for this branch
    if(isHolidayForBranch(g.date, g.branch)) return;

    // Guard 3: Need at least 3 people in the branch to avoid false positives
    if(g.total < 3) return;

    const absentRate = g.absent / g.total;

    // Ã¢â€â‚¬Ã¢â€â‚¬ TIER 1 Ã¢â‚¬â€ Zero Punch Day Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
    // If NOBODY in the branch punched at all Ã¢â€ â€™ 100% machine failure
    const isZeroPunchDay = g.punched === 0 && g.absent >= 3;

    // Ã¢â€â‚¬Ã¢â€â‚¬ TIER 2 Ã¢â‚¬â€ Near-Zero Punch Day Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
    // 1 or 2 people punched but the rest are absent Ã¢â€ â€™ almost certainly a glitch
    // (e.g. manager arrived early before machine died, or used manual entry)
    const isNearZeroDay = g.punched <= 2 && g.absent >= 3 && absentRate >= 0.80;

    // Ã¢â€â‚¬Ã¢â€â‚¬ TIER 3 Ã¢â‚¬â€ High Absence Spike Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
    // Ã¢â€°Â¥40% absent AND at least 3 people affected.
    // 40% is chosen because genuine mass-absence on a normal working day is very rare.
    const isHighSpike = absentRate >= 0.40 && g.absent >= 3;

    // Ã¢â€â‚¬Ã¢â€â‚¬ TIER 4 Ã¢â‚¬â€ Moderate Spike with large branch Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
    // Branch has 10+ people and Ã¢â€°Â¥30% absent Ã¢â‚¬â€ statistically improbable without a glitch
    const isModerateSpike = g.total >= 10 && absentRate >= 0.30 && g.absent >= 4;

    const isAnomaly = isZeroPunchDay || isNearZeroDay || isHighSpike || isModerateSpike;

    if(isAnomaly){
      // Determine the confidence reason for the insight panel
      let reason='high absence spike';
      if(isZeroPunchDay)       reason='zero punch day (total machine failure)';
      else if(isNearZeroDay)   reason='near-zero punches (likely machine failure)';
      else if(isHighSpike)     reason=`${Math.round(absentRate*100)}% absence spike`;
      else if(isModerateSpike) reason=`${Math.round(absentRate*100)}% absence in large branch`;

      S.failureDates.push({date:g.date, branch:g.branch, reason, affected:g.absent, total:g.total});

      // Reclassify all Absent records for this branch+date to System Error
      r.forEach(x=>{
        if(x.date===g.date && x.branch===g.branch && x.status==='Absent'){
          x.status='System Error';
          x.gapMins=0; x.gapFmt='0m'; x.gapClass='g-ok';
          // Clear late/early metrics Ã¢â‚¬â€ these aren't meaningful for a system error
          x.lateMins=0; x.earlyMins=0; x.lateBy='--'; x.earlyBy='--';
        }
      });
    }
  });
}

function quickFilter(s){
  document.getElementById('f-status').value=s;
  switchTab('daily', document.querySelector('[onclick*="switchTab(\'daily\'"]'));
  applyFilters();
}

function renderInsights(){
  const r=S.records; if(!r.length)return;
  const ins=document.getElementById('insights-panel'); ins.style.display='grid';
  const dash=v=>v||'<span class="dash">--</span>';
  
  // Logic: Top Branch (Highest Attendance %)
  const brData={}; 
  r.forEach(x=>{
    if(!x.branch)return;
    if(!brData[x.branch])brData[x.branch]={days:0,present:0};
    brData[x.branch].days++;
    if(x.status==='Present'||x.status==='Late')brData[x.branch].present++;
  });
  let topBr='', topPct=-1;
  Object.keys(brData).forEach(b=>{
    const pct=Math.round((brData[b].present/brData[b].days)*100);
    if(pct>topPct){topPct=pct; topBr=b;}
  });

  // Logic: Lateness Trend
  const latePct=Math.round((S.lateRecs.length/r.length)*100);
  const totalAtt=Math.round((r.filter(x=>x.status==='Present'||x.status==='Late').length/r.length)*100);

  ins.innerHTML=`
    <div class="insight-item">
      <div class="insight-icon">Top</div>
      <div class="insight-txt"><strong>Top Branch:</strong> <strong>${dash(topBr)}</strong> is leading with <strong>${topPct}%</strong> attendance stability.</div>
    </div>
    <div class="insight-item">
      <div class="insight-icon">KPI</div>
      <div class="insight-txt"><strong>Stability Score:</strong> Total workplace attendance is at <strong>${totalAtt}%</strong>. ${totalAtt>85?'Very healthy!':'Check for blockages.'}</div>
    </div>
    ${S.failureDates.length ? `
    <div class="insight-item" onclick="quickFilter('System Error')" style="cursor:pointer">
      <div class="insight-icon">Alert</div>
      <div class="insight-txt">
        <strong>Machine Failures Detected:</strong> <strong>${S.failureDates.length} incident${S.failureDates.length>1?'s':''}</strong> across <strong>${[...new Set(S.failureDates.map(f=>f.branch))].length} branch${[...new Set(S.failureDates.map(f=>f.branch))].length>1?'es':''}</strong> - reclassified as System Error.
        <br><span style="font-size:11px;color:var(--ink3);margin-top:3px;display:block">${S.failureDates.slice(0,3).map(f=>`${f.date} | ${f.branch||'Unknown'} (${f.reason})`).join(' | ')}${S.failureDates.length>3?` | +${S.failureDates.length-3} more`:''}. Click to view.</span>
      </div>
    </div>` : `
    <div class="insight-item">
      <div class="insight-icon">Time</div>
      <div class="insight-txt"><strong>Punctuality:</strong> <strong>${latePct}%</strong> of records are late. ${latePct<10?'Excellent discipline!':'Review shift overlaps.'}</div>
    </div>
    `}
  `;
}

function switchTab(name,btn){
  ['daily','summary','late','absent','early'].forEach(t=>document.getElementById('tab-'+t).style.display=t===name?'block':'none');
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  btn.classList.add('active');
  if(name==='summary')renderSummary();
  queueStickyLayoutSync();
}

function applyFilters(){
  const q=document.getElementById('search').value.toLowerCase();
  const branch=document.getElementById('f-branch').value;
  const dept=document.getElementById('f-dept').value;
  const status=document.getElementById('f-status').value;
  const empId=document.getElementById('f-emp').value;
  const fv=document.getElementById('date-from').value;
  const tv=document.getElementById('date-to').value;
  const hasDates=fv&&tv;
  document.getElementById('btn-dl').style.display=hasDates?'inline-flex':'none';
  document.getElementById('dl-gate').style.display=hasDates?'none':'inline-flex';
  let from=-Infinity,to=Infinity;
  if(fv)from=new Date(fv).setHours(0,0,0,0);
  if(tv)to=new Date(tv).setHours(23,59,59,999);
  
  S.filtered=S.records.filter(r=>{
    if(q&&!(r.name.toLowerCase().includes(q)||r.uid.includes(q)))return false;
    if(branch&&r.branch!==branch)return false;
    if(dept&&r.department!==dept)return false;
    if(status&&r.status!==status)return false;
    if(empId&&r.uid!==empId)return false;
    const rt=new Date(r.date).getTime();
    if(fv&&rt<from)return false;
    if(tv&&rt>to)return false;
    return true;
  });
  
  S.lateRecs=S.filtered.filter(r=>r.lateMins>0);
  S.absentRecs=S.filtered.filter(r=>r.status==='Absent');
  S.earlyRecs=S.filtered.filter(r=>r.earlyMins>0);
  
  document.getElementById('tb-daily').textContent=S.filtered.length;
  document.getElementById('tb-summary').textContent=[...new Set(S.filtered.map(r=>r.uid))].size;
  document.getElementById('tb-late').textContent=S.lateRecs.length;
  document.getElementById('tb-absent').textContent=S.absentRecs.length;
  document.getElementById('tb-early').textContent=S.earlyRecs.length;
  
  S.page=1;renderStats(S.filtered);renderTable();
  renderSubTables();renderSummary();
  
  document.getElementById('daily-title').textContent=
    S.filtered.length===S.records.length?'Daily Attendance Records':`Daily Attendance Records (${S.filtered.length} filtered)`;
}

function sortBy(col){
  S.sortDir=S.sortCol===col?S.sortDir*-1:1;S.sortCol=col;
  S.filtered.sort((a,b)=>{
    const av=a[col]??'',bv=b[col]??'';
    if(['hoursWorked','lateMins','earlyMins','gapMins'].includes(col))return(Number(av)-Number(bv))*S.sortDir;
    return String(av).localeCompare(String(bv))*S.sortDir;
  });
  renderTable();
}
function sortSub(type,col){
  const arr=type==='late'?S.lateRecs:type==='absent'?S.absentRecs:S.earlyRecs;
  arr.sort((a,b)=>['lateMins','earlyMins'].includes(col)?(Number(a[col]||0)-Number(b[col]||0)):String(a[col]||'').localeCompare(String(b[col]||'')));
  renderSubTables();
}

function renderTable(){
  const total=S.filtered.length,start=(S.page-1)*S.perPage;
  const slice=S.filtered.slice(start,start+S.perPage);
  const bdg=s=>{const m={Present:'present',Late:'late',Absent:'absent','Half Day':'half','Missed Punch':'missed','Week Off':'weekoff','Holiday':'holiday','Late (Comp)':'late-comp','System Error':'syserr'};return`<span class="badge b-${m[s]||'present'}">${s}</span>`};
  const hc=h=>h>=8?'h-ok':h>=4?'h-low':'h-zero';
  document.getElementById('table-body').innerHTML=slice.map((r,i)=>`
<tr onclick="toggleExp(${start+i})" id="row-${start+i}" style="animation-delay: ${i*0.04}s">
  <td class="td-emp" data-label="Employee"><strong>${r.name}</strong><small>ID ${r.uid}</small></td>
  <td title="${r.branch}" data-label="Branch">${r.branch||'<span class="dash">--</span>'}</td>
  <td title="${r.department}" data-label="Department">${r.department||'<span class="dash">--</span>'}</td>
  <td class="mono" data-label="Date">${r.date}</td>
  <td style="color:var(--ink3);font-size:12px" data-label="Day">${r.day}</td>
  <td class="mono" data-label="In">${r.firstIn}</td>
  <td class="mono" data-label="Out">${r.lastOut}</td>
  <td class="mono ${hc(r.hoursWorked)}" data-label="Hours">${r.hoursWorked}h</td>
  <td class="overtime-v" data-label="Overtime">${r.otMins>0?r.overtime:'<span class="dash">--</span>'}</td>
  <td class="mono" data-label="Punches">${r.punchCount}</td>
  <td data-label="Status">${bdg(r.status)}</td>
  <td class="${r.gapClass}" data-label="Gap">${r.gapMins<=0?'-':''}${r.gapFmt}</td>
  <td class="late-v" data-label="Late By">${r.lateMins>0?r.lateBy:'<span class="dash">--</span>'}</td>
  <td class="early-v" data-label="Early Out">${r.earlyMins>0?r.earlyBy:'<span class="dash">--</span>'}</td>
</tr>
<tr class="exp-row" id="exp-${start+i}">
  <td class="exp-cell" colspan="14">
    <div class="exp-inner">
      <div class="exp-stat"><label>Shift</label><span>${r.shiftDisplay}</span></div>
      <div class="exp-stat"><label>Total Punches</label><span>${r.punchCount}</span></div>
      <div class="exp-stat"><label>Overtime</label><span class="overtime-v">${r.otMins>0?r.overtime:'0m'}</span></div>
      <div class="exp-stat"><label>Hours Worked</label><span class="${hc(r.hoursWorked)}">${r.hoursWorked}h</span></div>
      <div class="exp-stat"><label>Department</label><span>${r.department||'--'}</span></div>
      <div class="exp-stat"><label>Branch</label><span>${r.branch||'--'}</span></div>
    </div>
  </td>
</tr>`).join('');
  const end=Math.min(start+S.perPage,total);
  document.getElementById('page-info').textContent=total===0?'No records':`Showing ${start+1}-${end} of ${total} records`;
  const pages=Math.ceil(total/S.perPage);let html='';
  if(pages>1){
    html+=`<button class="pBtn" onclick="goPage(${S.page-1})" ${S.page===1?'disabled':''}>&lt;</button>`;
    for(let i=1;i<=pages;i++){
      if(i===1||i===pages||Math.abs(i-S.page)<=1)html+=`<button class="pBtn ${i===S.page?'active':''}" onclick="goPage(${i})">${i}</button>`;
      else if(i===S.page-2||i===S.page+2)html+=`<button class="pBtn" style="pointer-events:none">...</button>`;
    }
    html+=`<button class="pBtn" onclick="goPage(${S.page+1})" ${S.page===pages?'disabled':''}>&gt;</button>`;
  }
  document.getElementById('page-ctrls').innerHTML=html;
}

function toggleExp(idx){
  if(window.innerWidth<=768)return;
  const row=document.getElementById('row-'+idx);
  const exp=document.getElementById('exp-'+idx);
  const open=exp.classList.contains('open');
  document.querySelectorAll('.exp-row.open').forEach(r=>r.classList.remove('open'));
  document.querySelectorAll('tbody tr.expanded').forEach(r=>r.classList.remove('expanded'));
  if(!open){exp.classList.add('open');row.classList.add('expanded')}
}
function goPage(p){const pages=Math.ceil(S.filtered.length/S.perPage);if(p>=1&&p<=pages){S.page=p;renderTable()}}

function renderStats(recs){
  const empSet=new Set(recs.map(r=>r.uid));
  const dateSet=new Set(recs.map(r=>r.date));
  const present=recs.filter(r=>['Present','Late','Late (Comp)'].includes(r.status)).length;
  const late=recs.filter(r=>r.lateMins>0).length;
  const absent=recs.filter(r=>r.status==='Absent').length;
  // Use non-system-error total for accurate attendance KPI
  const validRecs=recs.filter(r=>r.status!=='System Error').length;
  const pct=validRecs?Math.round(present/validRecs*100):0;
  document.getElementById('st-emp').textContent=empSet.size;
  document.getElementById('st-emp-sub').textContent=recs.length+' total records';
  document.getElementById('st-days').textContent=dateSet.size;
  document.getElementById('st-days-sub').textContent='unique dates';
  document.getElementById('st-att').textContent=pct+'%';
  document.getElementById('st-att-sub').textContent=present+' present days';
  document.getElementById('st-late').textContent=late;
  document.getElementById('st-late-sub').textContent=recs.length?Math.round(late/recs.length*100)+'% of records':'';
  document.getElementById('st-abs').textContent=absent;
  document.getElementById('st-abs-sub').textContent=recs.length?Math.round(absent/recs.length*100)+'% of records':'';
}

const CLR=['#1B4FD8','#0B7B60','#8A5A00','#5B3FA6','#C0280C','#3E4B66'];
function avatarCol(n){return CLR[n.charCodeAt(0)%CLR.length]}
function initials(n){return n.split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase()}

function renderSummary(){
  const q=(document.getElementById('search-summary').value||'').toLowerCase();
  const map={};
  S.records.forEach(r=>{
    if(!map[r.uid])map[r.uid]={uid:r.uid,name:r.name,branch:r.branch,dept:r.department,
      present:0,late:0,half:0,absent:0,days:0,hrs:0};
    const s=map[r.uid]; s.hrs+=r.hoursWorked;
    if(r.status==='System Error')return;
    s.days++;
    if(r.status==='Present')s.present++;
    else if(r.status==='Late'||r.status==='Late (Comp)'){s.present++;s.late++}
    else if(r.status==='Half Day')s.half++;
    else if(r.status==='Absent')s.absent++;
  });
  let emps=Object.values(map);
  if(q)emps=emps.filter(e=>e.name.toLowerCase().includes(q)||e.uid.includes(q));
  emps.sort((a,b)=>a.name.localeCompare(b.name));
  document.getElementById('summary-grid').innerHTML=emps.map((e,i)=>{
    const p=e.days?Math.round(e.present/e.days*100):0;
    const avg=e.days?Math.round(e.hrs/e.days*10)/10:0;
    const col=avatarCol(e.name);
    return `<div class="emp-card anim-fade" style="animation-delay: ${Math.min(i*0.05, 1)}s">
      <div class="emp-card-head">
        <div class="avatar" style="background:${col}22;color:${col}">${initials(e.name)}</div>
        <div class="emp-card-info"><h4 title="${e.name}">${e.name}</h4><small>ID ${e.uid}</small></div>
      </div>
      <div class="emp-meta">${e.dept||'--'} | ${e.branch||'--'}</div>
      <div class="att-bar-wrap"><div class="att-bar ${p<70?'low':''}" style="width:${p}%"></div></div>
      <div class="emp-pct-row"><span>Attendance <strong style="color:var(--ink)">${p}%</strong></span><span>Avg <strong style="color:var(--ink)">${avg}h/day</strong></span></div>
      <div class="emp-stats">
        <div class="emp-stat"><span class="sv sv-p">${e.present}</span><span class="sl">Present</span></div>
        <div class="emp-stat"><span class="sv sv-l">${e.late}</span><span class="sl">Late</span></div>
        <div class="emp-stat"><span class="sv sv-h">${e.half}</span><span class="sl">Half</span></div>
        <div class="emp-stat"><span class="sv sv-a">${e.absent}</span><span class="sl">Absent</span></div>
      </div>
    </div>`}).join('');
}

function renderSubTables(){
  const row=(x,i,type)=>`
    <tr class="anim-fade" style="animation-delay: ${i*0.04}s">
      <td class="td-emp" data-label="Employee"><strong>${x.name}</strong><small>ID ${x.uid}</small></td>
      <td title="${x.branch}" data-label="Branch">${x.branch||'--'}</td>
      <td title="${x.department}" data-label="Department">${x.department||'--'}</td>
      <td class="mono" data-label="Date">${x.date}</td>
      <td style="color:var(--ink3);font-size:12px" data-label="Day">${x.day}</td>
      <td data-label="Shift"><span class="shift-chip">${x.shiftDisplay}</span></td>
      ${type==='late'?`
        <td class="mono late-v" data-label="Arrived">${x.firstIn}</td>
        <td class="late-v" data-label="Late By">${x.lateBy}</td>
      `:type==='early'?`
        <td class="mono early-v" data-label="Left At">${x.lastOut}</td>
        <td class="early-v" data-label="Left Early By">${x.earlyBy}</td>
      `:`
        <td data-label="Status"><span class="badge b-absent">Absent</span></td>
      `}
    </tr>`;

  document.getElementById('late-body').innerHTML=S.lateRecs.length?S.lateRecs.slice(0,50).map((x,i)=>row(x,i,'late')).join(''):'<tr><td colspan="8" style="padding:24px;text-align:center;color:var(--ink3)">No late arrivals</td></tr>';
  document.getElementById('absent-body').innerHTML=S.absentRecs.length?S.absentRecs.slice(0,50).map((x,i)=>row(x,i,'absent')).join(''):'<tr><td colspan="7" style="padding:24px;text-align:center;color:var(--ink3)">No absences</td></tr>';
  document.getElementById('early-body').innerHTML=S.earlyRecs.length?S.earlyRecs.slice(0,50).map((x,i)=>row(x,i,'early')).join(''):'<tr><td colspan="8" style="padding:24px;text-align:center;color:var(--ink3)">No early departures</td></tr>';
}

function downloadExcel(){
  const data=S.filtered;
  if(!data.length)return showToast('No records to export.');
  const wb=XLSX.utils.book_new();
  const sm={};
  data.forEach(r=>{
    if(!sm[r.uid])sm[r.uid]={'Emp ID':r.uid,'Name':r.name,'Branch':r.branch,'Department':r.department,
      'Total Days':0,'Present':0,'Late':0,'Half Day':0,'Absent':0,'Avg Hrs/Day':0,'Att %':'','_h':0};
    const s=sm[r.uid];s['Total Days']++;s['_h']+=r.hoursWorked;
    if(r.status==='Present')s['Present']++;
    else if(r.status==='Late'){s['Present']++;s['Late']++}
    else if(r.status==='Half Day')s['Half Day']++;
    else s['Absent']++;
  });
  const sumRows=Object.values(sm).map(s=>{
    s['Avg Hrs/Day']=Math.round(s['_h']/s['Total Days']*100)/100;
    s['Att %']=Math.round(s['Present']/s['Total Days']*100)+'%';
    delete s['_h'];return s;
  });
  const ws1=XLSX.utils.json_to_sheet(sumRows);ws1['!cols']=[9,22,16,16,11,9,9,10,9,12,8].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb,ws1,'Summary');
  const ws2=XLSX.utils.json_to_sheet(data.map(r=>({'Emp ID':r.uid,'Name':r.name,'Branch':r.branch,'Department':r.department,
    'Date':r.date,'Day':r.day,'Shift':r.shiftDisplay,'First In':r.firstIn,'Last Out':r.lastOut,
    'Hours':r.hoursWorked,'Overtime':r.overtime,'Status':r.status,'Late By':r.lateBy,'Early Out':r.earlyBy,'Punches':r.punchCount})));
  ws2['!cols']=[9,22,16,16,12,6,16,9,9,7,10,10,9,9,8].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb,ws2,'Daily Detail');
  const lr=data.filter(r=>r.lateMins>0);
  if(lr.length){const ws3=XLSX.utils.json_to_sheet(lr.map(r=>({'Emp ID':r.uid,'Name':r.name,'Branch':r.branch,'Department':r.department,'Date':r.date,'Day':r.day,'Shift':r.shiftDisplay,'Arrived':r.firstIn,'Late By':r.lateBy})));ws3['!cols']=[9,22,16,16,12,6,16,9,9].map(w=>({wch:w}));XLSX.utils.book_append_sheet(wb,ws3,'Late Report')}
  const ar=data.filter(r=>r.status==='Absent');
  if(ar.length){const ws4=XLSX.utils.json_to_sheet(ar.map(r=>({'Emp ID':r.uid,'Name':r.name,'Branch':r.branch,'Department':r.department,'Date':r.date,'Day':r.day,'Shift':r.shiftDisplay})));ws4['!cols']=[9,22,16,16,12,6,16].map(w=>({wch:w}));XLSX.utils.book_append_sheet(wb,ws4,'Absent Report')}
  const er=data.filter(r=>r.earlyMins>0);
  if(er.length){const ws5=XLSX.utils.json_to_sheet(er.map(r=>({'Emp ID':r.uid,'Name':r.name,'Branch':r.branch,'Department':r.department,'Date':r.date,'Day':r.day,'Shift':r.shiftDisplay,'Left At':r.lastOut,'Left Early By':r.earlyBy})));ws5['!cols']=[9,22,16,16,12,6,16,9,12].map(w=>({wch:w}));XLSX.utils.book_append_sheet(wb,ws5,'Early Departure')}
  const from=document.getElementById('date-from').value,to=document.getElementById('date-to').value;
  XLSX.writeFile(wb,`HR_Attendance_${from}_to_${to}.xlsx`);
}
