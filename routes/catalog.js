const express = require('express');
const fs = require('fs');
const path = require('path');
const router = express.Router();

const CATALOG_PATH = path.join(__dirname, '..', 'data', 'catalog.json');
const BACKUPS_DIR = path.join(__dirname, '..', 'data', 'backups');

function readCatalog() {
  return JSON.parse(fs.readFileSync(CATALOG_PATH, 'utf8'));
}

function writeCatalog(data) {
  const tmp = CATALOG_PATH + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(data, null, 2), 'utf8');
  fs.renameSync(tmp, CATALOG_PATH);
}

function cellStr(v) { return String(v == null ? '' : v); }
function cellNum(v) { const n = Number(v); return Number.isFinite(n) ? n : 0; }

function genId(prefix, existingIds) {
  let id;
  do {
    id = prefix + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
  } while (existingIds.has(id));
  return id;
}

const CAT_PREFIX = {
  'CCTV 카메라': '카메라', '기타 IT기기': '기타IT', '공사비': '공사비',
  '통신 함체': '함체', 'SUS 304 Pole': '폴', '잡자재': '잡자재',
  '소프트웨어': '소프트웨어', '태양광 시스템': '태양광', 'A/S': 'AS',
  '녹화기(NVR)': '녹화기', 'HDD 디스크': 'HDD', '지능형 영상분석 시스템': '영상분석',
  'PC': 'PC', 'TV': 'TV', '이동식 스탠드형 거치대': '거치대',
  'VPN': 'VPN', '무선장비(AP)': '무선장비', '안전관리': '안전', '회선관리': '회선',
};

function sortItems(items) {
  return items.sort((a, b) => {
    if (a.cat !== b.cat) return a.cat < b.cat ? -1 : 1;
    return a.id < b.id ? -1 : a.id > b.id ? 1 : 0;
  });
}

function genItemId(category, existingIds) {
  const prefix = CAT_PREFIX[category] || category || '기타';
  let num = 1;
  // 해당 접두사의 기존 최대 번호를 찾아서 +1
  for (const id of existingIds) {
    const m = id.match(new RegExp('^' + prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '-(\\d+)$'));
    if (m) num = Math.max(num, parseInt(m[1], 10) + 1);
  }
  let id = prefix + '-' + String(num).padStart(3, '0');
  // 혹시 충돌 시 번호 증가
  while (existingIds.has(id)) {
    num++;
    id = prefix + '-' + String(num).padStart(3, '0');
  }
  return id;
}

// GET /api/catalog — 프론트엔드용 전체 카탈로그 (조인된 형태)
router.get('/catalog', (req, res) => {
  try {
    const raw = readCatalog();
    const itemsMap = {};
    for (const it of raw.items) itemsMap[it.id] = it;

    const ITEMS = raw.items.map(it => ({
      id: it.id, cat: it.cat, name: it.name, spec: it.spec,
      u: it.u || 'EA', p: it.p || 0, note: it.note || '',
    }));

    const SETS = raw.sets.map(s => ({
      id: s.id,
      name: s.name,
      members: s.members.map(m => {
        const linked = itemsMap[m.itemId];
        return {
          itemId: m.itemId,
          n: linked ? linked.name : m.itemId,
          spec: linked ? linked.spec : '',
          u: linked ? (linked.u || 'EA') : 'EA',
          q: m.q || 1,
          p: linked ? (linked.p || 0) : 0,
          cat: linked ? (linked.cat || '') : '',
          note: m.note || (linked ? (linked.note || '') : ''),
        };
      }),
    }));

    const byGroup = {};
    for (const si of raw.smartItems) {
      const gid = si.group_id;
      (byGroup[gid] = byGroup[gid] || []).push(si);
    }
    const SMART_GROUPS = raw.smartGroups
      .sort((a, b) => (a.sort || 0) - (b.sort || 0))
      .map(g => {
        const items = byGroup[g.id] || [];
        const totalPrice = items.reduce((s, it) => s + (it.price || 0) * Math.max(it.qty || 1, 1), 0);
        return { id: g.id, name: g.name, cat: g.cat, sort: g.sort || 0, items, totalPrice };
      });

    const PURCHASE_MAP = {};
    for (const pm of raw.purchaseMap) {
      PURCHASE_MAP[pm.id] = { name: pm.name, supplier: pm.supplier, note: pm.note };
    }

    res.json({
      SETS, ITEMS,
      NOTES_DEFAULT: raw.notesDefault || [],
      SMART_GROUPS,
      PURCHASE_MAP,
    });
  } catch (e) {
    console.error('[catalog ERROR]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── CRUD: Items ──

router.get('/catalog/items', (req, res) => {
  res.json(readCatalog().items);
});

router.post('/catalog/items', (req, res) => {
  const cat = readCatalog();
  const existingIds = new Set(cat.items.map(i => i.id));
  const item = {
    id: cellStr(req.body.id).trim() || genItemId(cellStr(req.body.cat), existingIds),
    cat: cellStr(req.body.cat),
    name: cellStr(req.body.name),
    spec: cellStr(req.body.spec),
    u: cellStr(req.body.u) || 'EA',
    p: cellNum(req.body.p),
    note: cellStr(req.body.note).trim(),
  };
  if (existingIds.has(item.id)) return res.status(409).json({ error: 'Duplicate ID' });
  cat.items.push(item);
  sortItems(cat.items);
  writeCatalog(cat);
  res.status(201).json(item);
});

router.put('/catalog/items/:id', (req, res) => {
  const cat = readCatalog();
  const idx = cat.items.findIndex(i => i.id === req.params.id);
  if (idx < 0) return res.status(404).json({ error: 'Not found' });
  const item = cat.items[idx];
  if (req.body.cat !== undefined) item.cat = cellStr(req.body.cat);
  if (req.body.name !== undefined) item.name = cellStr(req.body.name);
  if (req.body.spec !== undefined) item.spec = cellStr(req.body.spec);
  if (req.body.u !== undefined) item.u = cellStr(req.body.u) || 'EA';
  if (req.body.p !== undefined) item.p = cellNum(req.body.p);
  if (req.body.note !== undefined) item.note = cellStr(req.body.note).trim();
  sortItems(cat.items);
  writeCatalog(cat);
  res.json(item);
});

router.delete('/catalog/items/:id', (req, res) => {
  const cat = readCatalog();
  const before = cat.items.length;
  cat.items = cat.items.filter(i => i.id !== req.params.id);
  if (cat.items.length === before) return res.status(404).json({ error: 'Not found' });
  writeCatalog(cat);
  res.json({ ok: true });
});

// ── CRUD: Sets ──

router.get('/catalog/sets', (req, res) => {
  res.json(readCatalog().sets);
});

router.post('/catalog/sets', (req, res) => {
  const cat = readCatalog();
  const existingIds = new Set(cat.sets.map(s => s.id));
  const set = {
    id: cellStr(req.body.id).trim() || genId('세트-', existingIds),
    name: cellStr(req.body.name),
    members: Array.isArray(req.body.members) ? req.body.members.map(m => ({
      itemId: cellStr(m.itemId).trim(),
      sort: cellNum(m.sort),
      q: cellNum(m.q) || 1,
      note: cellStr(m.note || '').trim(),
    })) : [],
  };
  if (existingIds.has(set.id)) return res.status(409).json({ error: 'Duplicate ID' });
  cat.sets.push(set);
  writeCatalog(cat);
  res.status(201).json(set);
});

router.put('/catalog/sets/:id', (req, res) => {
  const cat = readCatalog();
  const idx = cat.sets.findIndex(s => s.id === req.params.id);
  if (idx < 0) return res.status(404).json({ error: 'Not found' });
  const set = cat.sets[idx];
  if (req.body.name !== undefined) set.name = cellStr(req.body.name);
  if (req.body.members !== undefined) {
    set.members = req.body.members.map(m => ({
      itemId: cellStr(m.itemId).trim(),
      sort: cellNum(m.sort),
      q: cellNum(m.q) || 1,
      note: cellStr(m.note || '').trim(),
    }));
  }
  writeCatalog(cat);
  res.json(set);
});

router.delete('/catalog/sets/:id', (req, res) => {
  const cat = readCatalog();
  const before = cat.sets.length;
  cat.sets = cat.sets.filter(s => s.id !== req.params.id);
  if (cat.sets.length === before) return res.status(404).json({ error: 'Not found' });
  writeCatalog(cat);
  res.json({ ok: true });
});

// ── CRUD: Smart Groups ──

router.get('/catalog/smart-groups', (req, res) => {
  const cat = readCatalog();
  res.json({ groups: cat.smartGroups, items: cat.smartItems });
});

router.post('/catalog/smart-groups', (req, res) => {
  const cat = readCatalog();
  const existingIds = new Set(cat.smartGroups.map(g => g.id));
  const group = {
    id: cellStr(req.body.id).trim() || genId('스마트-', existingIds),
    name: cellStr(req.body.name),
    cat: cellStr(req.body.cat),
    sort: cellNum(req.body.sort),
  };
  if (existingIds.has(group.id)) return res.status(409).json({ error: 'Duplicate ID' });
  cat.smartGroups.push(group);
  if (Array.isArray(req.body.items)) {
    for (const si of req.body.items) {
      cat.smartItems.push({
        group_id: group.id, item_id: cellStr(si.item_id).trim(),
        cat: cellStr(si.cat), name: cellStr(si.name), spec: cellStr(si.spec),
        unit: cellStr(si.unit) || 'EA', qty: cellNum(si.qty), price: cellNum(si.price),
        note: cellStr(si.note || ''),
      });
    }
  }
  writeCatalog(cat);
  res.status(201).json(group);
});

router.put('/catalog/smart-groups/:id', (req, res) => {
  const cat = readCatalog();
  const idx = cat.smartGroups.findIndex(g => g.id === req.params.id);
  if (idx < 0) return res.status(404).json({ error: 'Not found' });
  const group = cat.smartGroups[idx];
  if (req.body.name !== undefined) group.name = cellStr(req.body.name);
  if (req.body.cat !== undefined) group.cat = cellStr(req.body.cat);
  if (req.body.sort !== undefined) group.sort = cellNum(req.body.sort);
  if (Array.isArray(req.body.items)) {
    cat.smartItems = cat.smartItems.filter(si => si.group_id !== req.params.id);
    for (const si of req.body.items) {
      cat.smartItems.push({
        group_id: req.params.id, item_id: cellStr(si.item_id).trim(),
        cat: cellStr(si.cat), name: cellStr(si.name), spec: cellStr(si.spec),
        unit: cellStr(si.unit) || 'EA', qty: cellNum(si.qty), price: cellNum(si.price),
        note: cellStr(si.note || ''),
      });
    }
  }
  writeCatalog(cat);
  res.json(group);
});

router.delete('/catalog/smart-groups/:id', (req, res) => {
  const cat = readCatalog();
  const before = cat.smartGroups.length;
  cat.smartGroups = cat.smartGroups.filter(g => g.id !== req.params.id);
  cat.smartItems = cat.smartItems.filter(si => si.group_id !== req.params.id);
  if (cat.smartGroups.length === before) return res.status(404).json({ error: 'Not found' });
  writeCatalog(cat);
  res.json({ ok: true });
});

// ── CRUD: Purchase Map ──

router.get('/catalog/purchase-map', (req, res) => {
  res.json(readCatalog().purchaseMap);
});

router.post('/catalog/purchase-map', (req, res) => {
  const cat = readCatalog();
  const pm = {
    id: cellStr(req.body.id).trim(),
    name: cellStr(req.body.name).trim(),
    supplier: cellStr(req.body.supplier).trim(),
    note: cellStr(req.body.note).trim(),
  };
  if (!pm.id) return res.status(400).json({ error: 'ID is required' });
  if (cat.purchaseMap.find(p => p.id === pm.id)) return res.status(409).json({ error: 'Duplicate ID' });
  cat.purchaseMap.push(pm);
  writeCatalog(cat);
  res.status(201).json(pm);
});

router.put('/catalog/purchase-map/:id', (req, res) => {
  const cat = readCatalog();
  const idx = cat.purchaseMap.findIndex(p => p.id === req.params.id);
  if (idx < 0) return res.status(404).json({ error: 'Not found' });
  const pm = cat.purchaseMap[idx];
  if (req.body.name !== undefined) pm.name = cellStr(req.body.name).trim();
  if (req.body.supplier !== undefined) pm.supplier = cellStr(req.body.supplier).trim();
  if (req.body.note !== undefined) pm.note = cellStr(req.body.note).trim();
  writeCatalog(cat);
  res.json(pm);
});

router.delete('/catalog/purchase-map/:id', (req, res) => {
  const cat = readCatalog();
  const before = cat.purchaseMap.length;
  cat.purchaseMap = cat.purchaseMap.filter(p => p.id !== req.params.id);
  if (cat.purchaseMap.length === before) return res.status(404).json({ error: 'Not found' });
  writeCatalog(cat);
  res.json({ ok: true });
});

// ── Notes ──

router.get('/catalog/notes', (req, res) => {
  res.json(readCatalog().notesDefault);
});

router.put('/catalog/notes', (req, res) => {
  const cat = readCatalog();
  if (!Array.isArray(req.body)) return res.status(400).json({ error: 'Array expected' });
  cat.notesDefault = req.body.map(n => cellStr(n).trim()).filter(Boolean);
  writeCatalog(cat);
  res.json(cat.notesDefault);
});

// ── Backup ──

router.post('/catalog/backup', (req, res) => {
  if (!fs.existsSync(BACKUPS_DIR)) fs.mkdirSync(BACKUPS_DIR, { recursive: true });
  const now = new Date();
  const ts = now.toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const dest = path.join(BACKUPS_DIR, `catalog-${ts}.json`);
  fs.copyFileSync(CATALOG_PATH, dest);
  res.json({ ok: true, file: path.basename(dest) });
});

module.exports = router;
