const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const CATALOG_XLSX = path.join(__dirname, '..', 'data', 'catalog.xlsx');
const CATALOG_JSON = path.join(__dirname, '..', 'data', 'catalog.json');

function cellStr(v) { return String(v == null ? '' : v).replace(/\r\n/g, '\n'); }
function cellNum(v) { const n = Number(v); return Number.isFinite(n) ? n : 0; }

function normKey(k) { return String(k).trim().toLowerCase(); }

function parseItems(sh) {
  return XLSX.utils.sheet_to_json(sh, { defval: '' }).map(row => {
    const o = {};
    for (const k of Object.keys(row)) {
      const key = normKey(k);
      const v = row[k];
      if (['카테고리', 'category', 'cat'].includes(key)) { o.cat = cellStr(v); continue; }
      if (['품목id', 'itemid', 'item_id', 'id'].includes(key)) { o.id = cellStr(v).trim(); continue; }
      if (['품목명', 'itemname', 'name'].includes(key)) { o.name = cellStr(v); continue; }
      if (['사양', 'spec'].includes(key)) { o.spec = cellStr(v); continue; }
      if (['단위', 'unit', 'u'].includes(key)) { o.u = cellStr(v) || 'EA'; continue; }
      if (['단가', 'price', 'p'].includes(key)) { o.p = cellNum(v); continue; }
      if (['비고', 'note'].includes(key)) { o.note = cellStr(v).trim(); continue; }
    }
    return o;
  }).filter(r => r.id);
}

function parseSets(setsSh, memSh) {
  const setRows = XLSX.utils.sheet_to_json(setsSh, { defval: '' });
  const memRows = XLSX.utils.sheet_to_json(memSh, { defval: '' });

  const sets = setRows.map(row => {
    let id = '', name = '';
    for (const k of Object.keys(row)) {
      const key = normKey(k);
      if (['세트id', 'setid', 'set_id', 'id'].includes(key)) id = cellStr(row[k]).trim();
      if (['세트명', 'setname', 'ui name', 'ui_name', 'uiname', 'name'].includes(key)) name = cellStr(row[k]);
    }
    return { id, name, members: [] };
  }).filter(s => s.id);

  const bySet = {};
  for (const row of memRows) {
    let setId = '', itemId = '', itemName = '', unit = '', qty = 0, note = '', sort = 0;
    for (const k of Object.keys(row)) {
      const key = normKey(k);
      if (['세트id', 'setid', 'set_id'].includes(key)) setId = cellStr(row[k]).trim();
      if (['품목id', 'itemid', 'item_id', 'id', '단품id', 'items_id'].includes(key)) itemId = cellStr(row[k]).trim();
      if (['품목명', 'itemname', 'name'].includes(key)) itemName = cellStr(row[k]);
      if (['단위', 'unit', 'u'].includes(key)) unit = cellStr(row[k]);
      if (['수량', 'qty', 'quantity', 'q'].includes(key)) qty = cellNum(row[k]);
      if (['비고', 'note'].includes(key)) note = cellStr(row[k]).trim();
      if (['순서', 'sort'].includes(key)) sort = cellNum(row[k]);
    }
    if (!setId) continue;
    (bySet[setId] = bySet[setId] || []).push({ itemId, sort, q: qty || 1, note });
  }

  for (const s of sets) {
    const members = bySet[s.id] || [];
    members.sort((a, b) => a.sort - b.sort);
    s.members = members.map(m => ({ itemId: m.itemId, sort: m.sort, q: m.q, note: m.note }));
  }
  return sets;
}

function parseSmartGroups(sgSh) {
  return XLSX.utils.sheet_to_json(sgSh, { defval: '' }).map(row => {
    const o = {};
    for (const k of Object.keys(row)) {
      const key = normKey(k);
      if (['group_id', '그룹id', '그룹_id'].includes(key)) o.id = cellStr(row[k]).trim();
      if (['group_name', '그룹명'].includes(key)) o.name = cellStr(row[k]);
      if (['cat', '카테고리'].includes(key)) o.cat = cellStr(row[k]);
      if (['sort', '순서'].includes(key)) o.sort = cellNum(row[k]);
    }
    return o;
  }).filter(r => r.id);
}

function parseSmartItems(siSh) {
  return XLSX.utils.sheet_to_json(siSh, { defval: '' }).map(row => {
    const o = {};
    for (const k of Object.keys(row)) {
      const key = normKey(k);
      if (['group_id', '그룹id', '그룹_id'].includes(key)) o.group_id = cellStr(row[k]).trim();
      if (['items_id', 'item_id', '품목id', '품목_id'].includes(key)) o.item_id = cellStr(row[k]).trim();
      if (['소분류', 'sub_cat', 'subcat', 'sub_group'].includes(key)) o.cat = cellStr(row[k]).trim();
      if (['name', '품목명'].includes(key)) o.name = cellStr(row[k]);
      if (['규격', 'spec'].includes(key)) o.spec = cellStr(row[k]);
      if (['단위', 'unit'].includes(key)) o.unit = cellStr(row[k]).trim() || 'EA';
      if (['수량', 'quantity', 'qty'].includes(key)) o.qty = cellNum(row[k]);
      if (['가격', 'price', '단가'].includes(key)) o.price = cellNum(row[k]);
      if (['비고', 'note'].includes(key)) o.note = cellStr(row[k]);
      if (['설명', 'description'].includes(key)) o.desc = cellStr(row[k]);
    }
    return o;
  }).filter(r => r.group_id);
}

function parsePurchaseMap(prSh) {
  return XLSX.utils.sheet_to_json(prSh, { defval: '' }).map(row => {
    const o = {};
    for (const k of Object.keys(row)) {
      const key = normKey(k);
      if (['품목id', 'id', 'item_id'].includes(key)) o.id = cellStr(row[k]).trim();
      if (['구매품목명', 'name'].includes(key)) o.name = cellStr(row[k]).trim();
      if (['기본구매처', 'supplier', '구매처'].includes(key)) o.supplier = cellStr(row[k]).trim();
      if (['비고', 'note'].includes(key)) o.note = cellStr(row[k]).trim();
    }
    return o;
  }).filter(r => r.id);
}

function parseNotes(notesSh) {
  const rows = XLSX.utils.sheet_to_json(notesSh, { defval: '' });
  let notes = rows.map(r => {
    for (const k of Object.keys(r)) {
      const key = normKey(k);
      if (['text', '텍스트', '내용'].includes(key)) return cellStr(r[k]).trim();
    }
    const vals = Object.values(r).filter(v => v != null && String(v).trim() !== '');
    return vals.length ? cellStr(vals[0]).trim() : '';
  }).filter(Boolean);
  return notes;
}

// ── Main ──
if (!fs.existsSync(CATALOG_XLSX)) {
  console.error('catalog.xlsx not found:', CATALOG_XLSX);
  process.exit(1);
}

if (fs.existsSync(CATALOG_JSON)) {
  console.log('catalog.json already exists. Overwriting...');
}

const wb = XLSX.readFile(CATALOG_XLSX);
const S = wb.Sheets;

const itemsSh = S['단품목록'] || S.Items;
const setsSh = S['세트목록'] || S.Sets;
const memSh = S['세트구성'] || S.SetMembers;
const sgSh = S['스마트그룹'] || S.SmartGroups;
const siSh = S['스마트항목'] || S.SmartItems;
const prSh = S['구매요청'] || S.PurchaseRequest;
const notesSh = S['비고기본값'] || S.Notes_Default;

const catalog = {
  items: itemsSh ? parseItems(itemsSh) : [],
  sets: (setsSh && memSh) ? parseSets(setsSh, memSh) : [],
  smartGroups: sgSh ? parseSmartGroups(sgSh) : [],
  smartItems: siSh ? parseSmartItems(siSh) : [],
  purchaseMap: prSh ? parsePurchaseMap(prSh) : [],
  notesDefault: notesSh ? parseNotes(notesSh) : [],
};

fs.writeFileSync(CATALOG_JSON, JSON.stringify(catalog, null, 2), 'utf8');

console.log('Migration complete!');
console.log(`  Items:        ${catalog.items.length}`);
console.log(`  Sets:         ${catalog.sets.length}`);
console.log(`  Smart Groups: ${catalog.smartGroups.length}`);
console.log(`  Smart Items:  ${catalog.smartItems.length}`);
console.log(`  Purchase Map: ${catalog.purchaseMap.length}`);
console.log(`  Notes:        ${catalog.notesDefault.length}`);
console.log(`  Output: ${CATALOG_JSON}`);
