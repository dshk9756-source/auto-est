/** 카탈로그: data/catalog.xlsx — 시트/컬럼은 scripts/build-catalog.mjs 참고 */
const CATALOG_XLSX = "data/catalog.xlsx";

let SETS = [];
let ITEMS = [];

/* 연락처 자동 하이픈 포맷 (숫자만 허용, 010-0000-0000) */
function formatPhone(el){
  const nums = el.value.replace(/[^0-9]/g,'').slice(0,11);
  let formatted = nums;
  if(nums.length<=3) formatted = nums;
  else if(nums.length<=7) formatted = nums.slice(0,3)+'-'+nums.slice(3);
  else formatted = nums.slice(0,3)+'-'+nums.slice(3,7)+'-'+nums.slice(7);
  el.value = formatted;
}
let NOTES_DEFAULT = [];

function normRowKeys(row) {
  const o = {};
  for (const k of Object.keys(row)) {
    const key = String(k).trim().toLowerCase();
    const v = row[k];

    // 한글/영문 헤더를 모두 지원(앱 파서가 기대하는 키로 정규화)
    if (key === "세트id" || key === "setid" || key === "세트_id" || key === "set_id") {
      o.set_id = v;
      o.id = v;
      continue;
    }
    // 단품ID/품목ID 헤더도 앱이 기대하는 `id`로 매핑
    if (
      key === "단품id" ||
      key === "품목id" ||
      key === "itemid" ||
      key === "item_id" ||
      key === "items_id" ||
      key === "id"
    ) {
      o.id = v;
      continue;
    }
    if (key === "세트명" || key === "setname") {
      o.name = v;
      continue;
    }
    if (key === "카테고리" || key === "category" || key === "cat") {
      o.cat = v;
      continue;
    }
    if (key === "품목명" || key === "itemname" || key === "name") {
      // SetMembers는 n, Items는 name을 기대
      o.name = v;
      o.n = v;
      continue;
    }
    if (key === "사양" || key === "spec") {
      o.spec = v;
      continue;
    }
    if (key === "설명" || key === "desc" || key === "description") {
      o.desc = v;
      continue;
    }
    if (key === "단위" || key === "unit" || key === "u") {
      o.u = v;
      continue;
    }
    if (key === "수량" || key === "qty" || key === "q") {
      o.q = v;
      continue;
    }
    if (key === "단가" || key === "price" || key === "p") {
      o.p = v;
      continue;
    }
    if (key === "비고" || key === "note") {
      o.note = v;
      continue;
    }
    if (key === "순서" || key === "sort") {
      o.sort = v;
      continue;
    }

    o[key] = v;
  }
  return o;
}
function cellStr(v) {
  return String(v == null ? "" : v).replace(/\r\n/g, "\n");
}
function cellNum(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function catalogFromWorkbook(wb) {
  const X = typeof XLSX !== "undefined" ? XLSX : null;
  if (!X || !wb || !wb.Sheets) throw new Error("XLSX 워크북이 없습니다");
  const setsSh = wb.Sheets.Sets;
  const memSh = wb.Sheets.SetMembers;
  const itemsSh = wb.Sheets.Items;
  const notesSh = wb.Sheets.Notes_Default;
  if (!setsSh || !memSh || !itemsSh) throw new Error("시트 Sets, SetMembers, Items 가 필요합니다");

  const setRows = X.utils.sheet_to_json(setsSh, { defval: "" }).map(normRowKeys);
  const memRows = X.utils.sheet_to_json(memSh, { defval: "" }).map(normRowKeys);
  const itemRows = X.utils.sheet_to_json(itemsSh, { defval: "" }).map(normRowKeys);
  const noteRows = notesSh ? X.utils.sheet_to_json(notesSh, { defval: "" }).map(normRowKeys) : [];

  // 단품 ID → 단품 데이터 맵 (세트 멤버에서 단가 참조용)
  const itemsMap = {};
  for (const r of itemRows) {
    const iid = cellStr(r.id).trim();
    if (iid) itemsMap[iid] = r;
  }

  const bySet = {};
  for (const r of memRows) {
    const sid = cellStr(r.set_id).trim();
    if (!sid) continue;
    const sort = cellNum(r.sort);
    const note = cellStr(r.note).trim();

    // 단품ID가 있으면 Items 시트에서 이름/단가/사양 자동 참조
    const linkedId = cellStr(r.id).trim();
    const linked = linkedId ? itemsMap[linkedId] : null;

    const linkedNote = linked ? cellStr(linked.note).trim() : "";
    const m = {
      itemId: linkedId,
      n:    linked ? cellStr(linked.name) : cellStr(r.n),
      spec: linked ? cellStr(linked.spec) : cellStr(r.spec),
      u:    cellStr(r.u) || (linked ? cellStr(linked.u) : "") || "EA",
      q:    cellNum(r.q) || 0,
      p:    linked ? cellNum(linked.p) : cellNum(r.p),
    };
    const resolvedNote = note || linkedNote;
    if (resolvedNote) m.note = resolvedNote;
    (bySet[sid] = bySet[sid] || []).push({ sort, m });
  }
  for (const k of Object.keys(bySet)) {
    bySet[k].sort((a, b) => a.sort - b.sort);
    bySet[k] = bySet[k].map((x) => x.m);
  }

  const sets = setRows
    .filter((r) => cellStr(r.id).trim())
    .map((r) => {
      const id = cellStr(r.id).trim();
      return {
        id,
        name: cellStr(r.name),
        desc: cellStr(r.desc),
        cat: cellStr(r.cat),
        members: bySet[id] || [],
      };
    });

  const items = itemRows
    .filter((r) => cellStr(r.id).trim())
    .map((r) => {
      const it = {
        id: cellStr(r.id).trim(),
        cat: cellStr(r.cat),
        name: cellStr(r.name),
        spec: cellStr(r.spec),
        u: cellStr(r.u) || "EA",
        p: cellNum(r.p) || 0,
      };
      const note = cellStr(r.note).trim();
      if (note) it.note = note;
      return it;
    });

  let notesDefault = noteRows
    .map((r) => cellStr(r.text).trim() || cellStr(r.Text).trim())
    .filter(Boolean);
  if (!notesDefault.length) {
    for (const r of noteRows) {
      const vals = Object.values(r).filter((v) => v != null && String(v).trim() !== "");
      if (vals.length) notesDefault.push(cellStr(vals[0]));
    }
    notesDefault = notesDefault.filter(Boolean);
  }

  return { SETS: sets, ITEMS: items, NOTES_DEFAULT: notesDefault };
}

function showCatalogErr(msg) {
  const html =
    '<div class="catalog-err"><strong>카탈로그를 불러올 수 없습니다.</strong><br><span style="font-size:11px;color:var(--mist)">' +
    String(msg) +
    "</span><br><span style=\"font-size:10px;color:var(--mist2)\">로컬 HTTP 서버로 열어 주세요(file:// 에서는 fetch 가 막힐 수 있습니다). `data/catalog.xlsx` 경로와 헤더를 확인하세요.</span></div>";
  const sg = document.getElementById("set-grid");
  const ig = document.getElementById("item-grid");
  if (sg) sg.innerHTML = html;
  if (ig) ig.innerHTML = html;
}

function initAfterCatalog() {
  document.getElementById("f-date").value = new Date().toISOString().slice(0, 10);
  renderCatTabs();
  renderSets();
  renderItems();
  renderCart();
  updateHistCount();
}

async function loadCatalog() {
  const loading = '<div class="catalog-loading">카탈로그 불러오는 중…</div>';
  const sg = document.getElementById("set-grid");
  const ig = document.getElementById("item-grid");
  if (sg) sg.innerHTML = loading;
  if (ig) ig.innerHTML = loading;

  let data = null;
  try {
    const xRes = await fetch(CATALOG_XLSX, { cache: "no-store" });
    if (!xRes.ok) throw new Error("HTTP " + xRes.status + " (" + CATALOG_XLSX + ")");
    if (typeof XLSX === "undefined") throw new Error("XLSX 라이브러리가 로드되지 않았습니다");
    const buf = await xRes.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    data = catalogFromWorkbook(wb);
  } catch (e) {
    console.error(e);
    showCatalogErr(e.message || e);
    SETS = [];
    ITEMS = [];
    NOTES_DEFAULT = [];
    initAfterCatalog();
    return;
  }

  SETS = data.SETS || [];
  ITEMS = data.ITEMS || [];
  NOTES_DEFAULT = data.NOTES_DEFAULT || [];
  if (!SETS.length && !ITEMS.length) {
    showCatalogErr("세트/단품 데이터가 비어 있습니다.");
    SETS = [];
    ITEMS = [];
    NOTES_DEFAULT = [];
  }
  initAfterCatalog();
}
const STATUS_CFG={
  review:{label:'검토중',cls:'st-review',dot:'#F5A623'},
  active:{label:'진행중',cls:'st-active',dot:'#00C9A7'},
  done:{label:'완 료',cls:'st-done',dot:'#4A9EF5'},
  cancel:{label:'취 소',cls:'st-cancel',dot:'#4F6070'},
};

let cart={};
let histFilter='all';
const N=n=>Math.round(n).toLocaleString('ko-KR');
const NW=n=>N(n)+' 원';
const stop=e=>e&&e.stopPropagation();

function loadHistory(){try{return JSON.parse(localStorage.getItem('ow_hist_v2')||'[]')}catch{return[]}}
function saveHistory(h){localStorage.setItem('ow_hist_v2',JSON.stringify(h))}

function showView(id,btn){
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('on'));
  document.querySelectorAll('.h-tab').forEach(b=>b.classList.remove('on'));
  document.getElementById('v'+id).classList.add('on');
  btn.classList.add('on');
  if(id==='b')renderHistory();
}

function syncSteps(){
  const hasInfo=(document.getElementById('f-client').value||document.getElementById('f-site').value).trim();
  const hasCart=Object.keys(cart).length>0;
  setSt('step1',hasInfo?'done':'active');
  setSt('step2',hasCart?'done':hasInfo?'active':'');
  setSt('step3',hasCart?'active':'');
}
function setSt(id,s){document.getElementById(id).className='si'+(s?' '+s:'')}

function switchTab(tab,btn){
  document.querySelectorAll('.ctabs .ct').forEach(b=>b.classList.remove('on'));
  btn.classList.add('on');
  document.getElementById('tab-sets').style.display=tab==='sets'?'':'none';
  document.getElementById('tab-items').style.display=tab==='items'?'':'none';
}

function toggleSet(id,e){
  stop(e);
  const el=e?.currentTarget||document.getElementById('pc-'+id);
  if(cart[id]){delete cart[id];el.classList.remove('sel');}
  else{
    const s=SETS.find(x=>x.id===id);
    cart[id]={type:'set',name:s.name,cat:s.cat,mult:1,
      members:s.members.map(m=>({...m})),customPrices:{}};
    el.classList.add('sel');
  }
  renderCart();syncSteps();
}

function toggleItem(id,e){
  stop(e);
  const el=e?.currentTarget||document.getElementById('pc-'+id);
  if(cart[id]){delete cart[id];el&&el.classList.remove('sel');}
  else{
    // 이미 카트의 세트 멤버로 존재하면 해당 멤버 수량만 증가
    let merged=false;
    for(const c of Object.values(cart)){
      if(c.type!=='set')continue;
      const mem=c.members.find(m=>m.itemId===id);
      if(mem){mem.q+=1;merged=true;break;}
    }
    if(!merged){
      const it=ITEMS.find(x=>x.id===id);
      cart[id]={type:'item',name:it.name,cat:it.cat,qty:1,price:it.p,
        spec:it.spec,unit:it.u,note:it.note||'',customPrice:null};
      el&&el.classList.add('sel');
    }
  }
  renderCart();syncSteps();
}

function adj(id,d,e){
  stop(e);
  if(!cart[id])return;
  const c=cart[id];
  if(c.type==='set')c.mult=Math.max(1,(c.mult||1)+d);
  else c.qty=Math.max(1,(c.qty||1)+d);
  renderCart();
}

function onItemPriceChange(id,val){
  if(!cart[id])return;
  const num=parseInt(val.replace(/,/g,''))||0;
  cart[id].customPrice=num;
  recalcTotals();
}
function onMemberPriceChange(id,idx,val){
  if(!cart[id])return;
  const num=parseInt(val.replace(/,/g,''))||0;
  cart[id].customPrices=cart[id].customPrices||{};
  cart[id].customPrices[idx]=num;
  recalcTotals();
}
function formatPriceInput(el){
  const raw=el.value.replace(/,/g,'');
  if(raw&&!isNaN(raw))el.value=parseInt(raw).toLocaleString('ko-KR');
}

function effectiveItemPrice(c){return c.customPrice!=null?c.customPrice:c.price}
function effectiveMemberPrice(c,idx){
  return(c.customPrices&&c.customPrices[idx]!=null)?c.customPrices[idx]:c.members[idx].p;
}
function cartSupply(){
  let s=0;
  for(const[,c]of Object.entries(cart)){
    const mult=c.type==='set'?(c.mult||1):(c.qty||1);
    if(c.type==='set'){
      c.members.forEach((m,i)=>{s+=effectiveMemberPrice(c,i)*m.q*mult;});
    }else{
      s+=effectiveItemPrice(c)*mult;
    }
  }
  return s;
}

function renderCart(){
  const scroll=document.getElementById('cart-scroll');
  const ids=Object.keys(cart);
  const saveBtn=document.getElementById('save-btn');
  if(!ids.length){
    scroll.innerHTML=`<div class="cart-empty"><svg viewBox="0 0 44 44"><rect x="4" y="12" width="36" height="26" rx="3"/><path d="M14 12V8a8 8 0 0116 0v4"/><circle cx="16" cy="25" r="2"/><circle cx="28" cy="25" r="2"/></svg><p>세트 패키지 또는 단품을<br>선택하면 여기에 담깁니다<br><br><span style="font-family:var(--mono);font-size:9px;color:var(--mist2)">단가는 직접 수정 가능합니다</span></p></div>`;
    ['t-supply','t-vat','t-total'].forEach(id=>document.getElementById(id).textContent='—');
    document.getElementById('rp-meta').textContent='항목을 선택하세요';
    saveBtn.disabled=true;
    const excelBtnE=document.getElementById('excel-btn');if(excelBtnE)excelBtnE.disabled=true;
    return;
  }
  let html='';
  for(const id of ids){
    const c=cart[id];
    const mult=c.type==='set'?(c.mult||1):(c.qty||1);
    if(c.type==='set'){
      const totalLine=c.members.reduce((a,m,i)=>a+effectiveMemberPrice(c,i)*m.q*mult,0);
      const memRows=c.members.map((m,i)=>{
        const ep=effectiveMemberPrice(c,i);
        const la=ep*m.q*mult;
        return`<div class="cim-row" style="flex-direction:column;align-items:stretch;gap:3px">
          <div style="display:flex;justify-content:space-between"><span class="cim-n">${m.n}${m.note?` (${m.note})`:''} ×${m.q*mult}</span><span class="cim-q">${N(la)} 원</span></div>
          <div class="mem-price-row">
            <span class="mem-price-lbl">단가 수정</span>
            <input class="mem-price-in" value="${N(ep)}" onchange="onMemberPriceChange('${id}',${i},this.value);formatPriceInput(this)" oninput="onMemberPriceChange('${id}',${i},this.value)" onfocus="this.select()" onclick="event.stopPropagation()">
            <span style="font-family:var(--mono);font-size:9px;color:var(--mist)">원</span>
          </div>
        </div>`;
      }).join('');
      html+=`<div class="ci ci-set" data-setid="${id}" ondragover="onCartSetDragOver(event)" ondragleave="onCartSetDragLeave(event)">
        <div class="ci-top"><div class="ci-name">${c.name}</div><button class="ci-rm" onclick="rmCart('${id}')">×</button></div>
        <div class="ci-row"><span class="ci-tag">세트</span><span style="font-family:var(--mono);font-size:12px;color:var(--ice)">${N(totalLine)} 원</span></div>
        <div class="ci-stepper"><span class="ci-slbl">배수</span><button class="sbtn" onclick="adj('${id}',-1,event)">−</button><span class="sval">×${mult}</span><button class="sbtn" onclick="adj('${id}',1,event)">+</button></div>
        <div class="ci-drop-hint">단품을 여기에 드롭</div>
        <div class="ci-members" style="margin-top:8px;padding-top:8px;border-top:1px solid var(--wire)">${memRows}</div>
      </div>`;
    }else{
      const ep=effectiveItemPrice(c);
      const la=ep*mult;
      html+=`<div class="ci">
        <div class="ci-top"><div class="ci-name">${c.name}${c.note?` <span style="font-size:10px;color:var(--amber)">(${c.note})</span>`:''}</div><button class="ci-rm" onclick="rmCart('${id}')">×</button></div>
        <div class="ci-row"><span class="ci-tag item">단품</span><span style="font-family:var(--mono);font-size:12px;color:var(--ice)">${N(la)} 원</span></div>
        <div class="ci-price-row"><span class="ci-price-lbl">단가 수정</span><input class="ci-price-in" value="${N(ep)}" onchange="onItemPriceChange('${id}',this.value);formatPriceInput(this)" oninput="onItemPriceChange('${id}',this.value)" onfocus="this.select()" onclick="event.stopPropagation()"><span class="ci-price-unit">원</span></div>
        <div class="ci-stepper"><span class="ci-slbl">수량</span><button class="sbtn" onclick="adj('${id}',-1,event)">−</button><span class="sval">×${mult}</span><button class="sbtn" onclick="adj('${id}',1,event)">+</button></div>
      </div>`;
    }
  }
  scroll.innerHTML=html;
  recalcTotals();
  saveBtn.disabled=false;
  const excelBtn=document.getElementById('excel-btn');
  if(excelBtn)excelBtn.disabled=false;
}

function recalcTotals(){
  const supply=cartSupply();
  const vat=supply*.1;
  document.getElementById('t-supply').textContent=NW(supply);
  document.getElementById('t-vat').textContent=NW(vat);
  document.getElementById('t-total').textContent=NW(supply+vat);
  document.getElementById('rp-meta').textContent=`${Object.keys(cart).length}개 선택 · 합계 ${NW(supply+vat)}`;
}

function rmCart(id){
  delete cart[id];
  // HTML에서는 id가 유니크여야 하지만, 혹시 중복된 경우에도 클릭한 항목만 정상 해제되도록 처리
  document.querySelectorAll('[id="pc-'+id+'"]').forEach(el=>el.classList.remove('sel'));
  renderCart();syncSteps();
}
function clearCart(){cart={};document.querySelectorAll('.pc.sel').forEach(el=>el.classList.remove('sel'));renderCart();syncSteps();}

/* ── 카탈로그 → 견적 구성 드래그 앤 드롭 ────────────────────────── */
function onCatalogDragStart(e, type, id){
  e.dataTransfer.setData('dragType', type);
  e.dataTransfer.setData('dragId', id);
  e.dataTransfer.effectAllowed='copy';
  e.currentTarget.classList.add('dragging');
  document.getElementById('cart-scroll').classList.add('drop-zone');
}
function onCatalogDragEnd(e){
  e.currentTarget.classList.remove('dragging');
  document.getElementById('cart-scroll').classList.remove('drop-zone','drop-active');
  document.querySelectorAll('.ci-set').forEach(el=>el.classList.remove('drop-target'));
}
function onCartDragOver(e){
  e.preventDefault();
  e.dataTransfer.dropEffect='copy';
  e.currentTarget.classList.add('drop-active');
  // 단품을 세트 위로 드래그 중이면 세트 카드 하이라이트
  const setCard=e.target.closest('.ci-set');
  document.querySelectorAll('.ci-set').forEach(el=>el.classList.remove('drop-target'));
  if(setCard && e.dataTransfer.types.includes('dragtype')) setCard.classList.add('drop-target');
}
function onCartDragLeave(e){
  if(e.currentTarget.contains(e.relatedTarget)) return;
  e.currentTarget.classList.remove('drop-active');
  document.querySelectorAll('.ci-set').forEach(el=>el.classList.remove('drop-target'));
}
function onCartDrop(e){
  e.preventDefault();
  e.currentTarget.classList.remove('drop-active');
  document.querySelectorAll('.ci-set').forEach(el=>el.classList.remove('drop-target'));
  const type=e.dataTransfer.getData('dragType');
  const id=e.dataTransfer.getData('dragId');
  if(!type||!id) return;
  // 단품을 세트 카드 위에 드롭 → 세트 멤버에 추가
  const setCard=e.target.closest('.ci-set');
  if(setCard && type==='item'){
    const setId=setCard.dataset.setid;
    if(!cart[setId]) return;
    const src=ITEMS.find(i=>i.id===id);
    if(!src) return;
    const set=cart[setId];
    const existing=set.members.findIndex(m=>m.n===src.name);
    if(existing>=0){ set.members[existing].q+=1; }
    else { set.members.push({n:src.name,q:1,p:src.p,spec:src.spec||''}); }
    renderCart();syncSteps();
    return;
  }
  // 일반 드롭 → 카트에 추가 (이미 있으면 무시)
  if(type==='set') { if(!cart[id]) toggleSet(id); }
  else if(type==='item') { if(!cart[id]) toggleItem(id); }
}
function onCartSetDragOver(e){
  e.preventDefault();
  if(e.currentTarget.contains(e.relatedTarget)) return;
  e.currentTarget.classList.add('drop-target');
}
function onCartSetDragLeave(e){
  if(e.currentTarget.contains(e.relatedTarget)) return;
  e.currentTarget.classList.remove('drop-target');
}


function saveQuote(){
  const client=document.getElementById('f-client').value||'(미입력)';
  const site=document.getElementById('f-site').value||'(미입력)';
  const manager=document.getElementById('f-manager').value||'';
  const contact=document.getElementById('f-contact').value||'';
  const date=document.getElementById('f-date').value||new Date().toISOString().slice(0,10);
  const valid=document.getElementById('f-valid').value||'30일';
  const notes=getNotes();
  const supply=cartSupply();
  const snapshot=Object.entries(cart).map(([id,c])=>{
    if(c.type==='set'){
      return{type:'set',name:c.name,cat:c.cat,mult:c.mult||1,
        members:c.members.map((m,i)=>({...m,effectiveP:effectiveMemberPrice(c,i)}))};
    }else{
      return{type:'item',name:c.name,cat:c.cat,qty:c.qty||1,
        effectiveP:effectiveItemPrice(c),unit:c.unit||'식',note:c.note||''};
    }
  });
  const h=loadHistory();
  h.unshift({id:'q'+Date.now(),client,site,manager,contact,date,valid,notes,
    supply,vat:supply*.1,total:supply*1.1,status:'review',items:snapshot,
    createdAt:new Date().toISOString()});
  saveHistory(h);updateHistCount();
  const btn=document.getElementById('save-btn');
  btn.textContent='✓ 저장 완료';btn.style.borderColor='var(--sigbdr)';btn.style.color='var(--sig)';
  setTimeout(()=>{btn.textContent='이력에 저장';btn.style.borderColor='';btn.style.color='';},1800);
}

function updateHistCount(){
  const h=loadHistory();
  document.getElementById('hist-cnt').textContent=h.length?`(${h.length})`:'';
}

function filterHist(f,btn){
  histFilter=f;
  document.querySelectorAll('.vb-flt').forEach(b=>b.classList.remove('on'));
  btn.classList.add('on');
  renderHistory();
}

function renderHistory(){
  const h=loadHistory();
  const filtered=histFilter==='all'?h:h.filter(q=>q.status===histFilter);
  const body=document.getElementById('hist-body');
  if(!filtered.length){
    body.innerHTML=`<div class="hist-empty"><svg viewBox="0 0 48 48"><rect x="8" y="8" width="32" height="36" rx="3"/><line x1="16" y1="18" x2="32" y2="18"/><line x1="16" y1="24" x2="32" y2="24"/><line x1="16" y1="30" x2="24" y2="30"/></svg><p>${histFilter==='all'?'저장된 견적이 없습니다':'해당 상태의 견적이 없습니다'}</p></div>`;
    return;
  }
  let html=`<table class="hist-tbl"><thead><tr><th style="width:200px">거래처 / 현장</th><th>합계 금액</th><th>견적일</th><th style="width:100px">상태</th><th style="width:160px"></th></tr></thead><tbody>`;
  filtered.forEach(q=>{
    const sc=STATUS_CFG[q.status]||STATUS_CFG.review;
    html+=`<tr id="hr-${q.id}">
      <td><div class="hist-client">${q.client}</div><div class="hist-site">${q.site}</div></td>
      <td><div class="hist-amt">${NW(q.total)}</div><div style="font-family:var(--mono);font-size:9px;color:var(--mist);margin-top:2px">공급가 ${NW(q.supply)}</div></td>
      <td><div class="hist-date">${q.date}</div><div style="font-family:var(--mono);font-size:9px;color:var(--mist2);margin-top:2px">${q.manager||''}</div></td>
      <td><div class="st-wrap" id="stw-${q.id}">
        <span class="status-badge ${sc.cls}" onclick="toggleStDrop('${q.id}')">${sc.label}</span>
        <div class="st-drop">${Object.entries(STATUS_CFG).map(([k,v])=>`<div class="st-opt" onclick="changeStatus('${q.id}','${k}')"><div class="st-dot" style="background:${v.dot}"></div><span style="font-size:11px;color:var(--ice)">${v.label}</span></div>`).join('')}</div>
      </div></td>
      <td><div class="hist-actions"><button class="a-btn" onclick="toggleDetail('${q.id}')">상세</button><button class="a-btn" onclick="downloadFromHist('${q.id}')">다운로드</button><button class="a-btn del" onclick="deleteQuote('${q.id}')">삭제</button></div></td>
    </tr>
    <tr><td colspan="5" style="padding:0"><div class="hist-detail" id="hd-${q.id}">
      <div class="hd-title">견적 구성 품목</div>
      ${q.items.map(it=>{
        if(it.type==='set'){
          const tot=it.members.reduce((a,m)=>a+m.effectiveP*m.q*(it.mult||1),0);
          return`<div class="hd-row"><span class="hd-name">[세트] ${it.name} ×${it.mult}</span><span class="hd-amt">${NW(tot)}</span></div>`+
            it.members.map(m=>`<div class="hd-row" style="padding-left:16px"><span class="hd-name" style="color:var(--mist2)">↳ ${m.n} ×${m.q*(it.mult||1)}</span><span class="hd-amt" style="color:var(--mist)">${NW(m.effectiveP*m.q*(it.mult||1))}</span></div>`).join('');
        }else{
          return`<div class="hd-row"><span class="hd-name">${it.name} ×${it.qty}</span><span class="hd-amt">${NW(it.effectiveP*it.qty)}</span></div>`;
        }
      }).join('')}
    </div></td></tr>`;
  });
  html+='</tbody></table>';
  body.innerHTML=html;
  document.addEventListener('click',closeAllDrops,{once:true});
}

function toggleDetail(id){document.getElementById('hd-'+id).classList.toggle('open')}
function toggleStDrop(id){const w=document.getElementById('stw-'+id);const was=w.classList.contains('open');closeAllDrops();if(!was)w.classList.add('open');}
function closeAllDrops(){document.querySelectorAll('.st-wrap.open').forEach(w=>w.classList.remove('open'))}
function changeStatus(id,status){const h=loadHistory();const q=h.find(x=>x.id===id);if(q){q.status=status;saveHistory(h);}closeAllDrops();renderHistory();}
function deleteQuote(id){if(!confirm('이 견적 이력을 삭제할까요?'))return;saveHistory(loadHistory().filter(x=>x.id!==id));updateHistCount();renderHistory();}
/* ── 특기사항 동적 리스트 ────────────────────────────────────────── */
function addNote(value=''){
  const li=document.createElement('div');
  li.className='note-item';
  li.innerHTML=`<input type="text" placeholder="납품기한, 특이사항 등 추가 입력" value="${value.replace(/"/g,'&quot;')}">
    <button class="note-del-btn" onclick="removeNote(this)" title="삭제">×</button>`;
  document.getElementById('note-list').appendChild(li);
}
function removeNote(btn){btn.closest('.note-item').remove();}
function getNotes(){
  return [...document.querySelectorAll('#note-list .note-item input')]
    .map(i=>i.value.trim()).filter(Boolean);
}

/* ── 엑셀 페이로드 빌드 ─────────────────────────────────────────── */
function _buildPayload(src){
  if(src) return {
    client:src.client||'',site:src.site||'',
    manager:src.manager||'',contact:src.contact||'',
    date:src.date||'',valid:src.valid||'30일',
    notes:src.notes||(src.note?[src.note]:[]),items:src.items||[]
  };
  return {
    client:  document.getElementById('f-client').value||'(미입력)',
    site:    document.getElementById('f-site').value||'(미입력)',
    manager: document.getElementById('f-manager').value||'',
    contact: document.getElementById('f-contact').value||'',
    date:    document.getElementById('f-date').value||new Date().toISOString().slice(0,10),
    valid:   document.getElementById('f-valid').value||'30일',
    template: selectedTemplate,
    notes:   getNotes(),
    items:   Object.entries(cart).map(([,c])=>{
      if(c.type==='set') return{type:'set',name:c.name,cat:c.cat,mult:c.mult||1,
        members:c.members.map((m,i)=>({...m,effectiveP:effectiveMemberPrice(c,i)}))};
      return{type:'item',name:c.name,cat:c.cat,qty:c.qty||1,
        effectiveP:effectiveItemPrice(c),unit:c.unit||'식',spec:c.spec||'',note:c.note||''};
    })
  };
}

async function downloadFromHist(id){
  const q=loadHistory().find(x=>x.id===id);
  if(!q)return;
  const p=_buildPayload(q);
  try{
    const resp=await fetch('/api/generate-excel',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});
    if(!resp.ok){const e=await resp.json().catch(()=>({error:'서버 오류'}));throw new Error(e.error);}
    const blob=await resp.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;a.download=`${p.client}_${p.date}.xlsx`;
    document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
  }catch(e){alert('오류: '+e.message);}
}

function renderCatTabs(){
  const cats=['sets',...[...new Set(SETS.map(s=>s.cat))],'items'];
  const catLabels={sets:'세트 패키지',items:'단품 선택'};
  SETS.forEach(s=>catLabels[s.cat]=s.cat);
  const tabs=document.getElementById('cat-tabs');
  tabs.innerHTML=`<button class="ct on" onclick="switchTab('sets',this)">세트 패키지</button>
  <button class="ct" onclick="switchTab('items',this)">단품 선택</button>`;
}

function renderSets(){
  document.getElementById('set-grid').innerHTML=SETS.map(s=>{
    const tot=s.members.reduce((a,m)=>a+m.p*m.q,0);
    const mems=s.members.map(m=>`<div class="pm-row"><span class="pm-n">${m.n}</span><span class="pm-q">×${m.q}</span></div>`).join('');
    return`<div class="pc" id="pc-${s.id}" onclick="toggleSet('${s.id}',event)" draggable="true" ondragstart="onCatalogDragStart(event,'set','${s.id}')" ondragend="onCatalogDragEnd(event)">
      <div class="pc-chk"><svg viewBox="0 0 12 12"><polyline points="2,6 5,9 10,3"/></svg></div>
      <div class="pc-cat">${s.cat}</div>
      <div class="pc-name">${s.name}</div>
      <div class="pc-spec">${s.desc}</div>
      <div class="pc-foot"><span class="pc-price">${tot>0?N(tot)+' 원~':'별도협의'}</span><span class="pc-badge">${s.members.length}개 품목</span></div>
      <div class="pc-stepper" onclick="event.stopPropagation()">
        <span class="stlbl">배수</span>
        <button class="sbtn" onclick="adj('${s.id}',-1,event)">−</button>
        <span class="sval">×1</span>
        <button class="sbtn" onclick="adj('${s.id}',1,event)">+</button>
      </div>
      <div class="pc-mems">${mems}</div>
    </div>`;
  }).join('');
}

function renderItems(){
  const cats=[...new Set(ITEMS.map(i=>i.cat))];
  const sel=document.getElementById('item-cat');
  sel.innerHTML='<option value="">전체</option>'+cats.map(c=>`<option>${c}</option>`).join('');
  filterItems();
}

function filterItems(){
  const q=document.getElementById('item-q').value.toLowerCase();
  const cat=document.getElementById('item-cat').value;
  const list=ITEMS.filter(i=>(!q||i.name.toLowerCase().includes(q)||i.spec.toLowerCase().includes(q))&&(!cat||i.cat===cat));
  document.getElementById('item-grid').innerHTML=list.map(i=>{
    const sel=cart[i.id]?'sel':'';
    return`<div class="pc ${sel}" id="pc-${i.id}" onclick="toggleItem('${i.id}',event)" draggable="true" ondragstart="onCatalogDragStart(event,'item','${i.id}')" ondragend="onCatalogDragEnd(event)">
      <div class="pc-chk"><svg viewBox="0 0 12 12"><polyline points="2,6 5,9 10,3"/></svg></div>
      <div class="pc-cat">${i.cat}</div>
      <div class="pc-name">${i.name}</div>
      <div class="pc-foot"><span class="pc-price">${i.p>0?N(i.p)+' 원':'별도협의'}</span><span class="pc-badge item">단품</span></div>
      <div class="pc-stepper" onclick="event.stopPropagation()">
        <span class="stlbl">수량</span>
        <button class="sbtn" onclick="adj('${i.id}',-1,event)">−</button>
        <span class="sval">${cart[i.id]?.qty||1}</span>
        <button class="sbtn" onclick="adj('${i.id}',1,event)">+</button>
      </div>
    </div>`;
  }).join('');
}

loadCatalog();
loadTemplates();

/* ── 템플릿 선택 ─────────────────────────────────────────────────── */
let selectedTemplate = '';

async function loadTemplates() {
  try {
    const res = await fetch('/api/templates');
    const list = await res.json();
    const container = document.getElementById('tpl-list');
    if (!list.length) return;
    selectedTemplate = list[0].filename;
    container.innerHTML = list.map((t, i) =>
      `<button class="tpl-btn${i===0?' on':''}" onclick="selectTemplate(this,'${t.filename}')">${t.label}</button>`
    ).join('');
  } catch(e) { console.warn('템플릿 목록 로드 실패', e); }
}

function selectTemplate(btn, filename) {
  document.querySelectorAll('.tpl-btn').forEach(b => b.classList.remove('on'));
  btn.classList.add('on');
  selectedTemplate = filename;
}

/* ────────────────────────────────────────────────
   엑셀 견적서 다운로드
──────────────────────────────────────────────── */
async function generateExcel(){
  const btn=document.getElementById('excel-btn');
  const orig=btn.textContent;
  btn.textContent='생성 중…';btn.disabled=true;
  const p=_buildPayload();
  try{
    const resp=await fetch('/api/generate-excel',{
      method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)
    });
    if(!resp.ok){const e=await resp.json().catch(()=>({error:'서버 오류'}));throw new Error(e.error);}
    const blob=await resp.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;a.download=`${p.client}_${p.date}.xlsx`;
    document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
    btn.textContent='✓ 완료';btn.style.borderColor='var(--sigbdr)';btn.style.color='var(--sig)';
    setTimeout(()=>{btn.textContent=orig;btn.style.borderColor='';btn.style.color='';btn.disabled=false;},2000);
  }catch(e){
    alert('오류: '+e.message);btn.textContent=orig;btn.disabled=false;
  }
}
