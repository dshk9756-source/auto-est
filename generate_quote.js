'use strict';
/**
 * generate_quote.js
 * ─────────────────────────────────────────────────────────────────────────
 * 설계 원칙
 *  - 템플릿 파일은 절대 수정하지 않는다. 항상 버퍼 복사본에서 작업한다.
 *  - 코드는 {{마커}}의 위치만 파악하고, 디자인(스타일·병합·서식)은
 *    템플릿이 정의한 그대로 유지한다.
 *  - 스타일 필터링 / 높이 강제 조정 등 템플릿을 변형하는 코드 금지.
 *
 * 마커 규칙 (구조 마커는 C7에 배치)
 *  {{ITEM_START}}    아이템 반복 구역 시작
 *  {{ITEM_END}}      아이템 반복 구역 끝
 *  {{HEADER_ROW}}    컬럼 헤더 행 — 그룹마다 반복 (C1~C6 텍스트 그대로 사용)
 *  {{GROUP_TITLE}}   그룹 타이틀 행
 *  {{ITEM_ROW}}      아이템 데이터 행
 *  {{SUBTOTAL_ROW}}  소계 행 — 그룹마다 반복 (C1 텍스트 그대로, C5에 SUM 수식)
 *  {{EMPTY_ROW}}     그룹 간 빈 구분 행
 *  {{GRAND_TOTAL}}   합계 행 — C5에 소계 합산 SUM 수식 자동 생성
 *
 * 헤더 플레이스홀더
 *  {{거래처명 현장명}} {{담당자}} {{연락처}} {{견적일자}} {{견적명}}
 *  {{TOP_GROUP_1}} {{TOP_GROUP_2}}
 */

const ExcelJS = require('exceljs');
const JSZip   = require('jszip');
const fs      = require('fs');
const path    = require('path');

const TEMPLATE_PATH = path.join(__dirname, 'template', '견적서.xlsx');
const NUM_FMT       = '#,##0';

/* ═══════════════════════════════════════════════════════════════════════
   § 1. 행 스냅샷 / 복원
   ─ 템플릿 행의 스타일과 병합 정보를 완전히 캡처하고,
     새 행에 그대로 재현한다. 필터링·변형 없음.
═══════════════════════════════════════════════════════════════════════ */

/**
 * 지정 행의 모든 셀 스타일(font·fill·border·alignment·numFmt)과
 * 행 높이, 그리고 해당 행에 걸린 병합 범위를 캡처한다.
 */
function resolveThemeFill(fill, themeColors) {
  if (!fill || !themeColors) return fill;
  const f = JSON.parse(JSON.stringify(fill));
  const resolve = col => {
    if (!col || col.theme == null) return col;
    const base = themeColors[col.theme];
    if (!base) return col;
    return { argb: applyExcelTint(base, col.tint || 0) };
  };
  if (f.fgColor) f.fgColor = resolve(f.fgColor);
  if (f.bgColor && f.bgColor.theme != null) f.bgColor = resolve(f.bgColor);
  return f;
}

function snapshotRow(ws, rowNum, maxCol = 8, themeColors = null) {
  const row   = ws.getRow(rowNum);
  const cells = {};

  for (let c = 1; c <= maxCol; c++) {
    const cell = row.getCell(c);
    const rawFill = cell.fill ? JSON.parse(JSON.stringify(cell.fill)) : undefined;
    cells[c] = {
      font:      cell.font      ? JSON.parse(JSON.stringify(cell.font))      : undefined,
      fill:      resolveThemeFill(rawFill, themeColors),
      border:    cell.border    ? JSON.parse(JSON.stringify(cell.border))    : undefined,
      alignment: cell.alignment ? JSON.parse(JSON.stringify(cell.alignment)) : undefined,
      numFmt:    cell.numFmt    || undefined,
      value:     cell.value,    // 원본 셀 값 캡처 (텍스트·수식 모두)
    };
  }

  // 해당 행에 걸린 단일-행 병합 범위 수집 (ws._merges : ExcelJS 내부 맵)
  const merges = [];
  if (ws._merges) {
    Object.values(ws._merges).forEach(m => {
      const { left, top, right, bottom } = m.model;
      if (top === rowNum && bottom === rowNum && right > left) {
        merges.push({ left, right }); // 열 범위만 저장 (행은 대상 행으로 동적 결정)
      }
    });
  }

  return { height: row.height, cells, merges };
}

/* ── 행 높이 계산 헬퍼 ────────────────────────────────────────────────
   맑은고딕 10pt 기준 1줄 높이 (pt).
   font 10pt × 1.35 ≈ 13.5pt (위아래 여백 포함)
──────────────────────────────────────────────────────────────────── */
const LINE_HEIGHT_PT = 15.5; // 맑은고딕 10pt 기준

/**
 * 문자열이 주어진 열 너비(character unit)에서 몇 줄을 차지하는지 계산.
 * - CJK(한글·한자·전각) 문자는 2칸, ASCII 는 1칸
 * - \r\n · \r · \n 모두 명시적 줄바꿈으로 처리
 */
function estimateLines(text, colWidthChars) {
  if (!text) return 1;
  const segments = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  let total = 0;
  for (const seg of segments) {
    let w = 0;
    for (const ch of seg) w += ch.charCodeAt(0) > 0xFF ? 2 : 1;
    total += Math.max(1, Math.ceil(w / Math.max(colWidthChars, 1)));
  }
  return Math.max(1, total);
}

/**
 * rowNum 행의 wrapText=true 셀을 검사해 필요한 행 높이(pt)를 반환한다.
 * 열 너비는 읽기만 하며 절대 변경하지 않는다.
 */
function calcRowHeight(ws, rowNum, maxCol = 8) {
  const row = ws.getRow(rowNum);
  let maxLines = 1;
  for (let c = 1; c <= maxCol; c++) {
    const cell = row.getCell(c);
    if (!cell.alignment || !cell.alignment.wrapText) continue;
    const v = cell.value;
    if (!v || typeof v === 'object') continue; // 수식·null 제외
    const colWidth = ws.getColumn(c).width || 8;
    maxLines = Math.max(maxLines, estimateLines(String(v), colWidth));
  }
  return Math.round(maxLines * LINE_HEIGHT_PT);
}

/**
 * dstRn 행과 겹치는 모든 병합 항목을 _merges 에서 직접 제거한다.
 *
 * ws.unMergeCells() API를 쓰지 않는 이유:
 *   unMergeCells()는 내부적으로 findCell()을 호출하는데,
 *   spliceRows()로 새로 삽입된 빈 행의 셀은 findCell()이 null을 반환해
 *   실제 _merges 항목이 삭제되지 않는다.
 *   그 결과 _mergeCellsInternal()의 intersects() 체크에서 충돌이 발생한다.
 */
function unmergeRow(ws, rowNum) {
  if (!ws._merges) return;
  // 키 목록을 먼저 수집한 뒤 삭제 (반복 도중 객체 변경 방지)
  const keys = Object.keys(ws._merges).filter(k => {
    const { top, bottom } = ws._merges[k].model;
    return top <= rowNum && bottom >= rowNum;
  });
  keys.forEach(k => delete ws._merges[k]);
}

/**
 * 스냅샷을 dstRn 행에 복원한다.
 * - 기존 병합 해제 → 스타일 적용 → 병합 재적용 순서를 지킨다.
 * - 스타일: 템플릿 그대로 복사
 * - 병합:   템플릿 행에 있던 병합을 새 행에 재적용
 * - 값:     values 맵 {colNum: value} 으로만 채움, 나머지는 비움
 */
function stampRow(ws, dstRn, snap, values = {}, maxCol = 8) {
  // ① 해당 행의 기존 병합 범위를 먼저 모두 해제
  unmergeRow(ws, dstRn);

  const row = ws.getRow(dstRn);

  // ② 행 높이: 템플릿 값을 기본으로 적용 (item 행은 호출부에서 별도 재계산)
  // snap.height 가 0·undefined 인 경우도 안전하게 처리
  if (snap.height != null && snap.height > 0) row.height = snap.height;

  for (let c = 1; c <= maxCol; c++) {
    const cell = row.getCell(c);
    const s    = snap.cells[c] || {};

    if (s.font)      cell.font      = JSON.parse(JSON.stringify(s.font));
    if (s.fill)      cell.fill      = JSON.parse(JSON.stringify(s.fill));
    if (s.border)    cell.border    = JSON.parse(JSON.stringify(s.border));
    if (s.alignment) cell.alignment = JSON.parse(JSON.stringify(s.alignment));
    if (s.numFmt)    cell.numFmt    = s.numFmt;

    // 값 결정: 명시적 override > 스냅샷 원본값 > null
    // 스냅샷 값이 {{마커}}인 경우 자동 제거
    if (Object.prototype.hasOwnProperty.call(values, c)) {
      cell.value = values[c];
    } else {
      const sv = s.value;
      const isMarker = sv && typeof sv === 'string' && sv.includes('{{');
      cell.value = isMarker ? null : (sv ?? null);
    }
  }

  // ② 템플릿 행의 병합 패턴을 대상 행에 적용
  snap.merges.forEach(({ left, right }) => {
    ws.mergeCells(dstRn, left, dstRn, right);
  });
}

/* ═══════════════════════════════════════════════════════════════════════
   § 2. 마커 탐색  ─  모든 열 검색, 하드코딩 없음
═══════════════════════════════════════════════════════════════════════ */
function findMarkers(ws) {
  const m = {
    start:           -1,  // {{ITEM_START}} 행
    end:             -1,  // {{ITEM_END}} 행
    headerRow:       -1,  // {{HEADER_ROW}} → 컬럼 헤더 스타일 기준
    groupTitleFirst: -1,  // 첫 {{GROUP_TITLE}} → 삭제 시작점
    groupTitle:      -1,  // 마지막 {{GROUP_TITLE}} → 스타일 기준
    itemRow:         -1,  // {{ITEM_ROW}} → 아이템 행 스타일 기준
    subtotalRow:     -1,  // {{SUBTOTAL_ROW}} → 소계 스타일 기준
    emptyRow:        -1,  // {{EMPTY_ROW}} → 빈 구분 행 스타일 기준
    grandTotal:      -1,  // {{GRAND_TOTAL}} → 합계 행 (footer)
    topGroupFirst:   -1,  // 첫 {{TOP_GROUP_N}} 행
    topGroupLast:    -1,  // 마지막 {{TOP_GROUP_N}} 행 (반복 템플릿 기준)
    safetyNoteRows:  [],  // {{SAFETY_NOTE_ROW}} 행 목록 (안전관리 그룹 전용)
  };
  ws.eachRow((row, rn) => {
    row.eachCell({ includeEmpty: false }, cell => {
      const v = String(cell.value ?? '');
      if (v.includes('{{ITEM_START}}'))  m.start = rn;
      if (v.includes('{{ITEM_END}}'))    m.end   = rn;
      if (v.includes('{{HEADER_ROW}}')   && m.headerRow   < 0) m.headerRow   = rn;
      if (v.includes('{{SUBTOTAL_ROW}}') && m.subtotalRow  < 0) m.subtotalRow = rn;
      if (v.includes('{{EMPTY_ROW}}')    && m.emptyRow     < 0) m.emptyRow    = rn;
      if (v.includes('{{GRAND_TOTAL}}')  && m.grandTotal   < 0) m.grandTotal  = rn;
      if (v.includes('{{GROUP_TITLE}}')) {
        if (m.groupTitleFirst < 0) m.groupTitleFirst = rn;
        m.groupTitle = rn;
      }
      if (v.includes('{{ITEM_ROW}}'))    m.itemRow = rn;
      if (v.includes('{{TOP_GROUP_')) {
        if (m.topGroupFirst < 0) m.topGroupFirst = rn;
        m.topGroupLast = rn;
      }
      if (v.includes('{{SAFETY_NOTE_ROW}}')) m.safetyNoteRows.push(rn);
    });
  });
  return m;
}

/* ═══════════════════════════════════════════════════════════════════════
   § 3. 헤더 치환  ─  ITEM_START 행까지만, 행 번호 하드코딩 없음
   ※ {{TOP_GROUP_N}} 마커는 § 3.5 expandTopGroups 에서 처리
═══════════════════════════════════════════════════════════════════════ */
function fillHeader(ws, data, markers) {
  const stopRow = markers.start > 0 ? markers.start : 99;

  const phMap = {
    '{{거래처명 현장명}}': data.client  || '',
    '{{담당자}}':          data.manager || '',
    '{{연락처}}':          data.contact || '',
    '{{견적일자}}':         data.date    || '',
    '{{견적명}}':           data.site    || '',
  };

  ws.eachRow((row, rn) => {
    if (rn > stopRow) return;

    row.eachCell({ includeEmpty: false }, cell => {
      if (typeof cell.value !== 'string') return;
      let v = cell.value;

      // {{TOP_GROUP_N}} 은 expandTopGroups() 가 처리 — 여기서는 건드리지 않음
      if (v.includes('{{TOP_GROUP_')) return;

      for (const [ph, val] of Object.entries(phMap))
        if (v.includes(ph)) v = v.replace(ph, val);

      if (v !== null && v.includes('{{EMPTY_ROW}}'))
        v = null;
      if (v !== null && v.includes('{{ITEM_START}}'))
        v = null; // 마커 텍스트만 제거 (같은 행 다른 셀은 그대로)

      cell.value = (v === null || (typeof v === 'string' && v.trim() === ''))
        ? null : v;
    });
  });
}

/* ═══════════════════════════════════════════════════════════════════════
   § 3.5. TOP_GROUP 동적 확장
   ─ 템플릿의 {{TOP_GROUP_N}} 행 전체를 삭제하고,
     마지막 행의 스냅샷 스타일로 groupKeys 수만큼 전부 재생성한다.
   ─ 반환값 : 헤더에서 증감된 행 수 (delta). 이후 ITEM 구역 마커 위치 보정에 사용.
═══════════════════════════════════════════════════════════════════════ */
function expandTopGroups(ws, markers, groupKeys) {
  const { topGroupFirst, topGroupLast } = markers;
  if (topGroupFirst < 0) return 0; // 마커 없음 — 변화 없음

  const templateCount = topGroupLast - topGroupFirst + 1;

  // 재생성에 쓸 스타일 기준: 마지막 {{TOP_GROUP_N}} 행을 스냅샷
  const snap = snapshotRow(ws, topGroupLast);

  // 1. 템플릿 TOP_GROUP 행 전부 삭제
  if (ws._merges) {
    Object.keys(ws._merges)
      .filter(k => {
        const { top, bottom } = ws._merges[k].model;
        return top >= topGroupFirst && bottom <= topGroupLast;
      })
      .forEach(k => delete ws._merges[k]);
  }
  ws.spliceRows(topGroupFirst, templateCount);

  // 2. groupKeys 수만큼 새 행 삽입 후 스냅샷 스타일로 전부 재생성
  if (groupKeys.length > 0) {
    ws.spliceRows(topGroupFirst, 0, ...Array(groupKeys.length).fill([]));
    groupKeys.forEach((name, i) => {
      const overrides = { 1: `(${i + 1})${name}` };
      for (let c = 2; c <= 8; c++) overrides[c] = null;
      stampRow(ws, topGroupFirst + i, snap, overrides);
    });
  }

  // delta = 최종 행 수 - 원래 행 수
  return groupKeys.length - templateCount;
}

/* ═══════════════════════════════════════════════════════════════════════
   § 4. 행 계획 빌드  ─  장바구니 순서 유지, 카테고리 변경 시만 GROUP_TITLE
═══════════════════════════════════════════════════════════════════════ */
function buildPlan(items, hasEmptyRow, hasSafetyNote) {
  const safeGroups = new Set();
  if (hasSafetyNote) {
    for (const item of items) {
      const cat = item.cat || '';
      if (cat.includes('안전관리')) safeGroups.add(cat);
    }
  }

  const plan = [];
  let curGroup = null;
  let groupIdx = 0;

  for (const item of items) {
    const group = item.cat;

    if (group !== curGroup) {
      if (curGroup !== null) {
        plan.push({ type: 'subtotal' });
        if (hasEmptyRow) plan.push({ type: 'empty' });
      }
      groupIdx++;
      plan.push({ type: 'header' });
      plan.push({ type: 'group_title', name: `(${groupIdx})${group}` });
      if (safeGroups.has(group)) {
        plan.push({ type: 'safety_note', idx: 0 });
        plan.push({ type: 'safety_note', idx: 1 });
      }
      curGroup = group;
    }

    plan.push({
      type:  'item',
      col_a: item.name || '',
      col_b: item.spec || '',
      col_c: typeof item.effectiveP === 'number' ? item.effectiveP : 0,
      col_d: item.qty  || 1,
      col_f: item.note || '',
    });
  }

  if (curGroup !== null) {
    plan.push({ type: 'subtotal' });
    if (hasEmptyRow) plan.push({ type: 'empty' });
  }

  const seen = new Set(), groupKeys = [];
  for (const item of items) {
    const k = item.cat;
    if (!seen.has(k)) { seen.add(k); groupKeys.push(k); }
  }

  return { plan, groupKeys };
}

/* ═══════════════════════════════════════════════════════════════════════
   § 5. 핵심 파이프라인
   ─ 템플릿 버퍼 로드 → (복사본에서 작업) → 마커 탐색 → 헤더 치환
     → 템플릿 스냅샷 → 마커 구역 삭제 → 아이템 행 삽입 → 버퍼 반환
═══════════════════════════════════════════════════════════════════════ */
async function generateQuote(data) {
  /* ── 1. 템플릿을 버퍼로 읽어 인메모리 복사본 생성
         data.template 이 지정되면 해당 파일 사용, 없으면 기본 템플릿 사용 */
  const tplFile = data.template
    ? path.join(__dirname, 'template', path.basename(data.template))
    : TEMPLATE_PATH;
  const templateBuf = fs.readFileSync(tplFile);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(templateBuf);          // readFile 대신 load(buffer) — 명시적 복사

  const ws = wb.worksheets[0];
  if (!ws) throw new Error('템플릿에서 시트를 찾을 수 없습니다.');

  /* ── 2. 마커 탐색 */
  const markers = findMarkers(ws);
  const { start, end, headerRow, groupTitleFirst, groupTitle, itemRow, subtotalRow, emptyRow, grandTotal, safetyNoteRows } = markers;
  if (start < 0 || end < 0 || groupTitleFirst < 0 || itemRow < 0)
    throw new Error(`마커를 찾지 못했습니다: ${JSON.stringify(markers)}`);

  /* ── 2.5. footer 병합 스냅샷 ────────────────────────────────────────
        ITEM_END 행(end) 아래의 모든 병합(합계·특수조건 등)을 미리 기록한다.
        spliceRows 는 이 병합들을 깨뜨리거나(slave 셀을 String 으로 변환)
        _merges 항목을 잘못된 위치로 이동시키므로,
        나중에 올바른 새 위치에 재적용하기 위해 원본 좌표를 보존한다.
        expandTopGroups 전에 찍어야 원본 행 번호 기준으로 계산할 수 있다. */
  const footerMergesOrig = [];
  if (ws._merges) {
    Object.values(ws._merges).forEach(m => {
      const { top, left, bottom, right } = m.model;
      if (top > end) footerMergesOrig.push({ top, left, bottom, right });
    });
  }

  /* ── 3. 템플릿 행 스냅샷 — 삭제 전에 반드시 캡처
         스타일 + 병합 정보를 한 번에 저장, 코드는 해석·변형 안 함 */
  const headerRowSnap   = headerRow > 0   ? snapshotRow(ws, headerRow)   : null;
  const groupTitleSnap  = snapshotRow(ws, groupTitle);
  const itemRowSnap     = snapshotRow(ws, itemRow);
  const subtotalRowSnap = subtotalRow > 0 ? snapshotRow(ws, subtotalRow) : null;
  const emptyRowSnap    = emptyRow > 0    ? snapshotRow(ws, emptyRow)    : null;
  const safetyNoteSnaps = safetyNoteRows.map(rn => snapshotRow(ws, rn));

  /* ── 5. 행 계획 빌드 */
  const { plan, groupKeys } = buildPlan(data.items || [], emptyRow > 0, safetyNoteSnaps.length > 0);

  /* ── 6. 헤더 플레이스홀더 치환 ({{TOP_GROUP_N}} 제외) */
  fillHeader(ws, data, markers);

  /* ── 6a. TOP_GROUP 동적 확장
          헤더에서 행이 추가·삭제되므로, 이후 마커 행 번호를 delta만큼 보정 */
  const topDelta    = expandTopGroups(ws, markers, groupKeys);

  /* ── 7. 마커 구역 삭제 (HEADER_ROW 또는 첫 GROUP_TITLE ~ ITEM_END 포함)
          topDelta 만큼 보정된 행 번호를 사용 */
  const deleteFrom  = (headerRow > 0 && headerRow < groupTitleFirst) ? headerRow : groupTitleFirst;
  const insertAt    = deleteFrom + topDelta;
  const deleteCount = (end + topDelta) - insertAt + 1;

  // spliceRows 전에 삭제 범위 내 병합 항목을 _merges 에서 선제 제거한다.
  // ExcelJS spliceRows 는 삭제된 행의 _merges 항목을 자동 정리하지 않아,
  // 이후 stampRow 에서 mergeCells 호출 시 intersects() 충돌이 발생한다.
  if (ws._merges) {
    const lastDelRow = insertAt + deleteCount - 1;
    const staleKeys = Object.keys(ws._merges).filter(k => {
      const { top, bottom } = ws._merges[k].model;
      return top >= insertAt && bottom <= lastDelRow;
    });
    staleKeys.forEach(k => delete ws._merges[k]);
  }

  ws.spliceRows(insertAt, deleteCount);

  /* ── 8. 새 행 삽입 (빈 행) */
  if (plan.length > 0)
    ws.spliceRows(insertAt, 0, ...plan.map(() => []));

  /* ── 9. 각 행에 스타일·병합·값 적용 ─────────────────────────────────
        소계 행의 SUM 범위를 계산하기 위해, 현재 그룹의 첫 아이템 행 번호를
        추적한다. subtotalRows 배열에 소계 행 번호를 기록하여 합계 수식에 사용. */
  let groupItemStart = -1;  // 현재 그룹의 첫 item 행 번호
  const subtotalRows = [];  // 소계 행 번호 목록 (합계 수식용)

  plan.forEach((p, i) => {
    const rn = insertAt + i;

    if (p.type === 'header' && headerRowSnap) {
      // 컬럼 헤더 행 — 템플릿 원본 텍스트 그대로 사용 (override 없음)
      stampRow(ws, rn, headerRowSnap);
      groupItemStart = -1; // 새 그룹 시작 전 리셋

    } else if (p.type === 'group_title') {
      // 템플릿 GROUP_TITLE 행의 스타일·병합 그대로 + 그룹명
      stampRow(ws, rn, groupTitleSnap, { 1: ' ' + p.name });

    } else if (p.type === 'item') {
      if (groupItemStart < 0) groupItemStart = rn; // 첫 아이템 행 기록
      // 템플릿 ITEM_ROW 행의 스타일·병합 그대로 + 데이터
      stampRow(ws, rn, itemRowSnap, {
        1: p.col_a,
        2: p.col_b,
        3: p.col_c > 0 ? p.col_c : 0,
        4: p.col_d,
        5: { formula: `C${rn}*D${rn}` },
        6: p.col_f,
      });
      // 숫자 서식 (stampRow 이후 덮어쓰기)
      ws.getRow(rn).getCell(3).numFmt = NUM_FMT;
      ws.getRow(rn).getCell(4).numFmt = '#,##0';
      ws.getRow(rn).getCell(5).numFmt = NUM_FMT;
      // wrapText 명시 적용 (품목명·사양·비고) — 열 너비는 건드리지 않음
      [1, 2, 6].forEach(c => {
        const cell = ws.getRow(rn).getCell(c);
        cell.alignment = { ...(cell.alignment || {}), wrapText: true };
      });
      // 맑은고딕 10pt 기준 행 높이 자동 계산 (줄바꿈 \r\n 포함)
      // 템플릿 기본 높이보다 작아지지 않도록 max 처리
      ws.getRow(rn).height = Math.max(calcRowHeight(ws, rn), itemRowSnap.height || 0);

    } else if (p.type === 'subtotal' && subtotalRowSnap) {
      // 소계 행 — 템플릿 원본 텍스트 그대로 + C5에 SUM 수식만 override
      const sumFormula = groupItemStart > 0
        ? `SUM(E${groupItemStart}:E${rn - 1})`
        : 'SUM(E0:E0)'; // fallback (발생하지 않아야 함)
      stampRow(ws, rn, subtotalRowSnap, {
        5: { formula: sumFormula },
      });
      ws.getRow(rn).getCell(5).numFmt = NUM_FMT;
      subtotalRows.push(rn);
      groupItemStart = -1; // 리셋

    } else if (p.type === 'safety_note' && safetyNoteSnaps.length > 0) {
      const snap = safetyNoteSnaps[p.idx] || safetyNoteSnaps[0];
      stampRow(ws, rn, snap);

    } else if (p.type === 'empty' && emptyRowSnap) {
      // 빈 구분 행 — 템플릿 스타일 그대로 사용
      // value=null 이면 ExcelJS가 스타일을 누락할 수 있으므로 공백(' ')을 채운다
      const emptyVals = {};
      for (let c = 1; c <= 6; c++) emptyVals[c] = ' ';
      stampRow(ws, rn, emptyRowSnap, emptyVals);

    } else {
      // fallback 빈 행
      ws.getRow(rn).getCell(1).value = '';
    }
  });

  /* ── 9.5. footer 병합 복원 ───────────────────────────────────────────
        spliceRows 가 footer 행(합계·특수조건 등)의 slave 셀을 String 타입으로
        변환하고 _merges 항목을 잘못된 위치로 이동시켜 병합이 깨지는 문제를 수정.

        새 위치 = 원본 행 + netShift
          netShift = topDelta(헤더 증감) − deleteCount(아이템 구역 삭제) + plan.length(새 행 삽입)

        ws.mergeCells() 재호출로 _merges 항목과 slave 셀 타입을 동시에 복원한다.
        footerMergesOrig 는 위의 §2.5 에서 expandTopGroups 전에 캡처한 값.      */
  const footerNetShift = topDelta - deleteCount + plan.length;
  const footerStart    = insertAt + plan.length; // 첫 번째 footer 행

  // ① 잘못된 위치에 남아 있는 기존 footer 병합 항목 제거
  if (ws._merges) {
    Object.keys(ws._merges)
      .filter(k => ws._merges[k].model.top >= footerStart)
      .forEach(k => delete ws._merges[k]);
  }

  // ② 올바른 새 위치에 병합 재적용
  footerMergesOrig.forEach(({ top, left, bottom, right }) => {
    ws.mergeCells(top + footerNetShift, left, bottom + footerNetShift, right);
  });

  /* ── 9.6. 특기사항 처리 ─────────────────────────────────────────────────
        notes 배열의 각 항목을 {{SPECIAL_NOTE}} 행 스타일로 순서대로 삽입.
        항목이 없으면 {{SPECIAL_NOTE}} 행 삭제.
        {{AFTER_NOTE_NUM}} = 7 + notes.length                               */
  {
    const notes = Array.isArray(data.notes)
      ? data.notes.filter(Boolean)
      : (data.note && data.note.trim() ? [data.note.trim()] : []);

    let specialNoteRowNum = -1;
    let specialNoteCol    = 1;
    let specialNoteStart  = 7; // {{SPECIAL_NOTE:N}} 에서 파싱

    ws.eachRow((row, rn) => {
      row.eachCell({ includeEmpty: false }, cell => {
        if (typeof cell.value !== 'string') return;
        if (cell.value.includes('{{SPECIAL_NOTE:')) {
          specialNoteRowNum = rn;
          const numMatch = cell.value.match(/\{\{SPECIAL_NOTE:(\d+)\}\}/);
          specialNoteStart = numMatch ? parseInt(numMatch[1], 10) : 7;
          const letters = cell.address.replace(/[0-9]/g, '');
          let col = 0;
          for (const ch of letters) col = col * 26 + ch.charCodeAt(0) - 64;
          specialNoteCol = col || 1;
        }
        if (cell.value.includes('{{AFTER_NOTE_NUM}}'))
          cell.value = cell.value.replace('{{AFTER_NOTE_NUM}}', String(specialNoteStart + notes.length));
      });
    });

    if (specialNoteRowNum > 0) {
      const netShift = notes.length === 0 ? -1 : notes.length - 1;

      // spliceRows 후 ExcelJS가 아래 행 병합을 올바르게 갱신하지 못하므로
      // 미리 캡처 후 수동 재적용
      const belowMerges = [];
      if (ws._merges) {
        Object.values(ws._merges).forEach(m => {
          if (m.model.top > specialNoteRowNum) belowMerges.push({ ...m.model });
        });
        Object.keys(ws._merges)
          .filter(k => ws._merges[k].model.top > specialNoteRowNum)
          .forEach(k => delete ws._merges[k]);
      }

      if (notes.length === 0) {
        unmergeRow(ws, specialNoteRowNum);
        ws.spliceRows(specialNoteRowNum, 1);
      } else {
        const snap = snapshotRow(ws, specialNoteRowNum);
        unmergeRow(ws, specialNoteRowNum);
        ws.spliceRows(specialNoteRowNum, 1);
        ws.spliceRows(specialNoteRowNum, 0, ...notes.map(() => []));
        notes.forEach((txt, i) => {
          const rn  = specialNoteRowNum + i;
          const val = ` ${specialNoteStart + i}. ${txt}`;
          stampRow(ws, rn, snap, {});
          // mergeCells 이후 master 셀 값이 초기화될 수 있으므로 병합 적용 후 명시 재설정
          ws.getRow(rn).getCell(specialNoteCol).value = val;
        });
      }

      // 아래 행 병합을 netShift 만큼 이동해 재적용
      belowMerges.forEach(({ top, left, bottom, right }) => {
        ws.mergeCells(top + netShift, left, bottom + netShift, right);
      });
    }
  }

  /* ── 10. 합계 행 SUM 수식 갱신  ─  {{GRAND_TOTAL}} 마커 기반
         마커가 있던 행을 찾아 C5에 소계 합산 수식을 넣고, C7 마커를 제거한다. */
  if (grandTotal > 0) {
    const gtRow = grandTotal + footerNetShift;
    const gtCell = ws.getRow(gtRow).getCell(5);

    if (subtotalRows.length > 0) {
      const refs = subtotalRows.map(sr => `E${sr}`).join(',');
      gtCell.value = { formula: `SUM(${refs})` };
    } else {
      gtCell.value = { formula: `SUM(E${insertAt}:E${gtRow - 1})` };
    }
    gtCell.numFmt = NUM_FMT;

    // C7 마커 텍스트 제거
    const markerCell = ws.getRow(gtRow).getCell(7);
    if (markerCell.value && String(markerCell.value).includes('{{GRAND_TOTAL}}')) {
      markerCell.value = null;
    }
  }

  /* ── 11. printArea 제거 ───────────────────────────────────────────────
        템플릿의 고정 printArea(예: "A1:F42")를 삭제해 행 수 제한을 없앤다. */
  if (ws.pageSetup) {
    delete ws.pageSetup.printArea;
  }

  const generated = await wb.xlsx.writeBuffer();
  return copyDrawings(templateBuf, generated);
}

/* ═══════════════════════════════════════════════════════════════════════
   § 6. 도형·이미지 복원
   ─ ExcelJS 의 spliceRows 는 drawing XML 앵커를 잘못 수정하여
     일부 도형을 삭제한다. 생성 완료 후 템플릿 원본 drawing 으로 덮어써
     손상을 복원한다.
   ─ xl/drawings/* · xl/media/* 전체를 동적으로 탐색하므로
     템플릿이 바뀌어도 코드 수정 불필요.
═══════════════════════════════════════════════════════════════════════ */
/* ── 테마 색상 → ARGB 변환 헬퍼 ───────────────────────────────────── */
function parseThemeColors(themeXml) {
  const order = ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6'];
  return order.map(name => {
    const srgb = themeXml.match(new RegExp(`<a:${name}[^>]*>[^<]*<a:srgbClr val="([0-9A-Fa-f]{6})"`, 'i'));
    if (srgb) return 'FF' + srgb[1].toUpperCase();
    const sys = themeXml.match(new RegExp(`<a:${name}[^>]*>[^<]*<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"`, 'i'));
    if (sys) return 'FF' + sys[1].toUpperCase();
    return null;
  });
}

function applyExcelTint(argbHex, tint) {
  if (!tint) return argbHex;
  const r = parseInt(argbHex.slice(2,4), 16) / 255;
  const g = parseInt(argbHex.slice(4,6), 16) / 255;
  const b = parseInt(argbHex.slice(6,8), 16) / 255;
  const max = Math.max(r,g,b), min = Math.min(r,g,b);
  let h = 0, l = (max+min)/2, s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d/(2-max-min) : d/(max+min);
    if (max === r)      h = ((g-b)/d + (g<b?6:0)) / 6;
    else if (max === g) h = ((b-r)/d + 2) / 6;
    else                h = ((r-g)/d + 4) / 6;
  }
  l = tint < 0 ? l*(1+tint) : l*(1-tint)+tint;
  l = Math.max(0, Math.min(1, l));
  const hue2rgb = (p,q,t) => {
    if (t<0) t+=1; if (t>1) t-=1;
    if (t<1/6) return p+(q-p)*6*t;
    if (t<1/2) return q;
    if (t<2/3) return p+(q-p)*(2/3-t)*6;
    return p;
  };
  let nr, ng, nb;
  if (s === 0) { nr=ng=nb=l; }
  else {
    const q = l<0.5 ? l*(1+s) : l+s-l*s, p = 2*l-q;
    nr = hue2rgb(p,q,h+1/3);
    ng = hue2rgb(p,q,h);
    nb = hue2rgb(p,q,h-1/3);
  }
  const hex = v => Math.round(Math.min(1,Math.max(0,v))*255).toString(16).padStart(2,'0').toUpperCase();
  return 'FF' + hex(nr) + hex(ng) + hex(nb);
}

async function copyDrawings(templateBuf, generatedBuf) {
  const [tplZip, outZip] = await Promise.all([
    JSZip.loadAsync(templateBuf),
    JSZip.loadAsync(generatedBuf),
  ]);

  /* ── 1. xl/drawings/* 와 xl/media/* : 템플릿 원본으로 덮어쓰기 ──
     ExcelJS 가 spliceRows 과정에서 앵커를 잘못 수정한 drawing XML 을
     템플릿 원본으로 복원한다. 파일명은 하드코딩하지 않고 동적 탐색.    */
  const toCopy = Object.keys(tplZip.files).filter(f =>
    (f.startsWith('xl/drawings/') || f.startsWith('xl/media/') || f.startsWith('xl/theme/')) &&
    !tplZip.files[f].dir
  );
  for (const p of toCopy) {
    outZip.file(p, await tplZip.files[p].async('nodebuffer'));
  }

  /* ── 2. xl/workbook.xml — definedNames 완전 제거 ──
     ExcelJS 가 명명된 범위를 잘못 기록하거나, 행 삽입/삭제로 참조가 깨져
     Excel 열 때 복구 경고가 뜨는 문제 수정.
     행 번호가 바뀐 상태에서 원본을 그대로 이식하면 여전히 깨지므로
     생성 파일의 definedNames 블록을 완전히 제거한다.                         */
  const wbPath = 'xl/workbook.xml';
  if (outZip.files[wbPath]) {
    let outWb = await outZip.files[wbPath].async('string');
    outWb = outWb.replace(/<definedNames[\s\S]*?<\/definedNames>\s*/g, '');
    outZip.file(wbPath, outWb);
  }

  /* ── 3. [Content_Types].xml — drawing · 이미지 타입 누락 시 보완 ──
     ExcelJS 가 Content_Types 에서 drawing/이미지 항목을 누락시키는
     경우를 대비해 템플릿에 있는 항목을 추가한다.                        */
  const ctPath = '[Content_Types].xml';
  if (tplZip.files[ctPath] && outZip.files[ctPath]) {
    const tplCt = await tplZip.files[ctPath].async('string');
    let   outCt = await outZip.files[ctPath].async('string');

    const entries = [
      ...[...tplCt.matchAll(/<Default[^>]+(png|jpg|jpeg|gif|bmp|emf|wmf)[^>]+\/>/gi)].map(m => m[0]),
      ...[...tplCt.matchAll(/<Override[^>]+drawing[^>]+\/>/g)].map(m => m[0]),
    ];
    for (const entry of entries) {
      const ext      = entry.match(/Extension="([^"]+)"/)?.[1];
      const partName = entry.match(/PartName="([^"]+)"/)?.[1];
      const alreadyIn = (ext && outCt.includes(`Extension="${ext}"`)) ||
                        (partName && outCt.includes(`PartName="${partName}"`));
      if (!alreadyIn) outCt = outCt.replace('</Types>', entry + '\n</Types>');
    }
    outZip.file(ctPath, outCt);
  }

  return outZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/* ═══════════════════════════════════════════════════════════════════════
   § 7. 스마트시스템 내역서 전용 생성기
   ─ 스마트 견적서 템플릿 (2)상세내역서 시트) 전용
   ─ 마커: {{ITEM_START}} {{ITEM_END}} {{MAIN_GROUP_ROW}} {{SUB_GROUP_ROW}}
           {{ITEM_ROW}} {{GRAND_TOTAL}}
   ─ 컬럼: A=코드 B=공종 C=규격 D=단위 E=수량 F=공급단가 G=공급가액 H=비고 I=마커
═══════════════════════════════════════════════════════════════════════ */
async function generateSmartQuote(data) {
  const tplFile = path.join(__dirname, 'template', path.basename(data.template));
  const templateBuf = fs.readFileSync(tplFile);

  /* ── 테마 색상 사전 파싱 (스냅샷 시 theme 색상 → ARGB 변환용) */
  let themeColors = null;
  try {
    const tplZipForTheme = await JSZip.loadAsync(templateBuf);
    const themeFile = tplZipForTheme.files['xl/theme/theme1.xml'];
    if (themeFile) themeColors = parseThemeColors(await themeFile.async('string'));
  } catch(e) { /* 테마 파싱 실패 시 무시 */ }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(templateBuf);

  const ws = wb.getWorksheet('2)상세내역서');
  if (!ws) throw new Error('2)상세내역서 시트를 찾을 수 없습니다.');

  /* ── 마커 탐색 */
  const m = { start: -1, end: -1, mainGroup: -1, subGroup: -1, itemRow: -1, grandTotal: -1 };
  ws.eachRow((row, rn) => {
    row.eachCell({ includeEmpty: false }, cell => {
      const v = String(cell.value ?? '');
      if (v.includes('{{ITEM_START}}'))     m.start      = rn;
      if (v.includes('{{ITEM_END}}'))       m.end        = rn;
      if (v.includes('{{MAIN_GROUP_ROW}}')) m.mainGroup  = rn;
      if (v.includes('{{SUB_GROUP_ROW}}'))  m.subGroup   = rn;
      if (v.includes('{{ITEM_ROW}}'))       m.itemRow    = rn;
      if (v.includes('{{GRAND_TOTAL}}'))    m.grandTotal = rn;
    });
  });
  if (m.start < 0 || m.end < 0 || m.mainGroup < 0 || m.itemRow < 0)
    throw new Error(`스마트 마커를 찾지 못했습니다: ${JSON.stringify(m)}`);

  /* ── 템플릿 행 스냅샷 (삭제 전) — themeColors 전달로 theme 색상 즉시 ARGB 변환 */
  const mainGroupSnap = snapshotRow(ws, m.mainGroup, 9, themeColors);
  const subGroupSnap  = m.subGroup > 0 ? snapshotRow(ws, m.subGroup, 9, themeColors) : null;
  const itemRowSnap   = snapshotRow(ws, m.itemRow, 9, themeColors);

  /* ── 헤더 플레이스홀더 치환 (전 시트, ITEM_START 이전 행만) */
  const phMap = {
    '{{현장명}}':          data.client  || '',
    '{{거래처명 현장명}}': data.client  || '',
    '{{담당자}}':          data.manager || '',
    '{{연락처}}':          data.contact || '',
    '{{견적일자}}':         data.date    || '',
    '{{견적명}}':           data.site    || '',
  };
  for (const sheet of wb.worksheets) {
    const stop = sheet === ws ? m.start - 1 : 9999;
    sheet.eachRow((row, rn) => {
      if (rn > stop) return;
      row.eachCell({ includeEmpty: false }, cell => {
        if (typeof cell.value !== 'string') return;
        let v = cell.value;
        for (const [ph, val] of Object.entries(phMap)) v = v.replace(ph, val);
        cell.value = v;
      });
    });
  }

  /* ── 헤더 행 높이 사전 캡처 (spliceRows가 헤더 행 높이를 초기화하는 버그 방지) */
  const headerHeights = {};
  for (let rn = 1; rn < m.start; rn++) {
    const row = ws.getRow(rn);
    if (row.height != null && row.height > 0) headerHeights[rn] = row.height;
  }

  /* ── footer 병합 사전 캡처 */
  const footerMergesOrig = [];
  if (ws._merges) {
    Object.values(ws._merges).forEach(mg => {
      const { top, left, bottom, right } = mg.model;
      if (top > m.end) footerMergesOrig.push({ top, left, bottom, right });
    });
  }

  /* ── 행 계획 빌드 */
  const smartGroups = (data.items || []).filter(it => it.type === 'smart_group');
  const plan = [];
  for (const group of smartGroups) {
    plan.push({ type: 'main_group', name: group.name });
    const items = group.items || [];
    const hasCats = subGroupSnap && items.some(it => it.cat && String(it.cat).trim());
    if (hasCats) {
      const catOrder = [], byCat = {};
      for (const it of items) {
        const c = String(it.cat || '').trim() || '기타';
        if (!byCat[c]) { byCat[c] = []; catOrder.push(c); }
        byCat[c].push(it);
      }
      for (const cat of catOrder) {
        plan.push({ type: 'sub_group', name: cat });
        for (const it of byCat[cat]) plan.push({ type: 'item', item: it });
      }
    } else {
      for (const it of items) plan.push({ type: 'item', item: it });
    }
  }

  /* ── 마커 구역 삭제 (ITEM_START ~ ITEM_END)
       ITEM_START 행에 헤더 내용(수량·공급단가 등)이 함께 있는 경우,
       그 행은 보존하고 마커 셀만 지운 뒤 다음 행부터 삭제한다. */
  const itemStartRow = ws.getRow(m.start);
  let itemStartHasContent = false;
  itemStartRow.eachCell({ includeEmpty: false }, cell => {
    const v = String(cell.value ?? '');
    if (!v.includes('{{') && v.trim() !== '') itemStartHasContent = true;
  });
  if (itemStartHasContent) {
    itemStartRow.eachCell({ includeEmpty: false }, cell => {
      if (String(cell.value ?? '').includes('{{ITEM_START}}')) cell.value = null;
    });
  }

  const insertAt    = itemStartHasContent ? m.start + 1 : m.start;
  const deleteCount = m.end - insertAt + 1;
  if (ws._merges) {
    Object.keys(ws._merges)
      .filter(k => { const { top, bottom } = ws._merges[k].model; return top >= insertAt && bottom <= insertAt + deleteCount - 1; })
      .forEach(k => delete ws._merges[k]);
  }
  ws.spliceRows(insertAt, deleteCount);

  /* ── 새 행 삽입 */
  if (plan.length > 0) ws.spliceRows(insertAt, 0, ...plan.map(() => []));

  /* ── 행 스탬프 */
  const itemGRows = [];
  let mainGroupCounter = 0;
  const mainGroups2 = [];
  const subGroups2 = [];
  let curMain2 = null, curSub2 = null;
  plan.forEach((p, i) => {
    const rn = insertAt + i;
    if (p.type === 'main_group') {
      mainGroupCounter++;
      stampRow(ws, rn, mainGroupSnap, { 1: mainGroupCounter, 2: p.name }, 9);
      curMain2 = { rn, subRns: [], itemRns: [] };
      mainGroups2.push(curMain2);
      curSub2 = null;
    } else if (p.type === 'sub_group' && subGroupSnap) {
      stampRow(ws, rn, subGroupSnap, { 1: null, 2: p.name }, 9);
      curSub2 = { rn, itemRns: [] };
      subGroups2.push(curSub2);
      if (curMain2) curMain2.subRns.push(rn);
    } else if (p.type === 'item') {
      const it    = p.item;
      const qty   = typeof it.qty === 'number' ? it.qty : (typeof it.quantity === 'number' ? it.quantity : 1);
      const price = typeof it.price === 'number' ? it.price : 0;
      stampRow(ws, rn, itemRowSnap, {
        1: null,
        2: it.name    || '',
        3: it.spec    || '',
        4: it.unit    || '식',
        5: qty,
        6: price > 0 ? price : null,
        7: price > 0 ? { formula: `E${rn}*F${rn}` } : null,
        8: it.note    || null,
      }, 9);
      ws.getRow(rn).getCell(5).numFmt = '#,##0';
      if (price > 0) {
        ws.getRow(rn).getCell(6).numFmt = NUM_FMT;
        ws.getRow(rn).getCell(7).numFmt = NUM_FMT;
        itemGRows.push(rn);
      }
      if (curSub2) curSub2.itemRns.push(rn);
      else if (curMain2) curMain2.itemRns.push(rn);
    }
  });

  /* ── 소분류 SUM 수식 (해당 소분류 아이템들의 G열 합계) */
  for (const sg of subGroups2) {
    if (sg.itemRns.length > 0) {
      const cell = ws.getRow(sg.rn).getCell(7);
      cell.value = { formula: `SUM(G${sg.itemRns[0]}:G${sg.itemRns[sg.itemRns.length-1]})` };
      cell.numFmt = NUM_FMT;
    }
  }
  /* ── 대분류 SUM 수식 (소분류 G열 합계, 소분류 없으면 아이템 직접 합계) */
  for (const mg of mainGroups2) {
    const cell = ws.getRow(mg.rn).getCell(7);
    if (mg.subRns.length > 0) {
      cell.value = { formula: `SUM(${mg.subRns.map(r => `G${r}`).join(',')})` };
    } else if (mg.itemRns.length > 0) {
      cell.value = { formula: `SUM(G${mg.itemRns[0]}:G${mg.itemRns[mg.itemRns.length-1]})` };
    }
    cell.numFmt = NUM_FMT;
  }

  /* ── footer 병합 복원 + 합계 행 수식 */
  const netShift = plan.length - deleteCount;
  if (ws._merges) {
    Object.keys(ws._merges)
      .filter(k => ws._merges[k].model.top >= insertAt + plan.length)
      .forEach(k => delete ws._merges[k]);
  }
  footerMergesOrig.forEach(({ top, left, bottom, right }) => {
    ws.mergeCells(top + netShift, left, bottom + netShift, right);
  });

  if (m.grandTotal > 0) {
    const gtRn    = m.grandTotal + netShift;
    const gtCell  = ws.getRow(gtRn).getCell(7);
    gtCell.value  = itemGRows.length > 0
      ? { formula: `SUM(${itemGRows.map(r => `G${r}`).join(',')})` }
      : 0;
    gtCell.numFmt = NUM_FMT;
    ws.getRow(gtRn).getCell(9).value = null; // 마커 제거
  }

  if (ws.pageSetup) delete ws.pageSetup.printArea;

  /* ── 헤더 행 높이 복원 (spliceRows로 변경된 헤더 높이를 템플릿 원본으로 되돌림) */
  for (const [rn, h] of Object.entries(headerHeights)) {
    ws.getRow(Number(rn)).height = h;
  }

  /* ── 잔여 마커 행 정리 ─────────────────────────────────────────────
     spliceRows 이후에도 남아있을 수 있는 {{마커}} 텍스트를 가진 행을
     역순으로 찾아 삭제한다.                                           */
  const markerRows = [];
  ws.eachRow((row, rn) => {
    let hasMarker = false;
    row.eachCell({ includeEmpty: false }, cell => {
      const v = String(cell.value ?? '');
      if (v.includes('{{') && v.includes('}}')) hasMarker = true;
    });
    if (hasMarker) markerRows.push(rn);
  });
  for (let i = markerRows.length - 1; i >= 0; i--) {
    unmergeRow(ws, markerRows[i]);
    ws.spliceRows(markerRows[i], 1);
  }

  /* ── 마커 열 숨김 (I열=col9: 프린트 범위에서 제외) */
  ws.getColumn(9).hidden = true;

  /* ═══════════════════════════════════════════════════════════════════
     § 1)내역집계표 처리
     ─ 대분류·소분류별 합계 금액을 집계해 내역집계표에 삽입한다.
     ─ 컬럼: A=번호 B=공종 C=규격 D=단위 E=수량 F=공급단가 G=공급가액 H=비고 I=마커
  ═══════════════════════════════════════════════════════════════════ */
  const ws1 = wb.getWorksheet('1)내역집계표');
  if (ws1) {
    /* ── 그룹/소분류별 합계 사전 계산 */
    const groupTotals = {};
    for (const group of smartGroups) {
      let groupTotal = 0;
      const subs = {};
      for (const it of (group.items || [])) {
        const qty   = typeof it.qty === 'number' ? it.qty : (typeof it.quantity === 'number' ? it.quantity : 1);
        const price = typeof it.price === 'number' ? it.price : 0;
        const amt   = qty * price;
        groupTotal += amt;
        const cat = String(it.cat || '').trim() || '기타';
        subs[cat] = (subs[cat] || 0) + amt;
      }
      groupTotals[group.name] = { total: groupTotal, subs };
    }

    /* ── 마커 탐색 */
    const m1 = { start: -1, end: -1, mainGroup: -1, subGroup: -1, grandTotal: -1 };
    ws1.eachRow((row, rn) => {
      row.eachCell({ includeEmpty: false }, cell => {
        const v = String(cell.value ?? '');
        if (v.includes('{{ITEM_START}}'))     m1.start     = rn;
        if (v.includes('{{ITEM_END}}'))       m1.end       = rn;
        if (v.includes('{{MAIN_GROUP_ROW}}')) m1.mainGroup = rn;
        if (v.includes('{{SUB_GROUP_ROW}}'))  m1.subGroup  = rn;
        if (v.includes('{{GRAND_TOTAL}}'))    m1.grandTotal = rn;
      });
    });

    if (m1.start > 0 && m1.end > 0 && m1.mainGroup > 0) {
      /* ── 열 너비 사전 캡처 (spliceRows 후 리셋 방지) */
      const colWidths1 = {};
      ws1.columns.forEach((col, idx) => { if (col.width) colWidths1[idx + 1] = col.width; });

      /* ── 템플릿 스냅샷 */
      const mainSnap1 = snapshotRow(ws1, m1.mainGroup, 9, themeColors);
      const subSnap1  = m1.subGroup > 0 ? snapshotRow(ws1, m1.subGroup, 9, themeColors) : null;

      /* ── 소분류 스냅샷 fill 보정: theme=0(dk1=검정)은 "배경 없음"으로 처리 */
      if (subSnap1) {
        for (let c = 1; c <= 9; c++) {
          const cell = subSnap1.cells[c];
          if (cell && cell.fill && cell.fill.fgColor && cell.fill.fgColor.argb === 'FF000000') {
            cell.fill = undefined;
          }
        }
      }

      /* ── 헤더 행 높이 캡처 */
      const headerHeights1 = {};
      for (let rn = 1; rn <= m1.start; rn++) {
        const r = ws1.getRow(rn);
        if (r.height != null && r.height > 0) headerHeights1[rn] = r.height;
      }

      /* ── ITEM_START 행(m1.start) theme 색상 → 명시적 ARGB 변환
           rows 1~(start-1)는 copyDrawings theme XML 복원으로 처리하지만,
           ITEM_START 행은 theme 참조가 Excel에서 흰색으로 렌더링되는 문제가 있어
           resolveThemeFill로 명시적 ARGB를 직접 지정한다. */
      if (themeColors) {
        const startR = ws1.getRow(m1.start);
        for (let c = 1; c <= 9; c++) {
          const cell = startR.getCell(c);
          if (cell.fill && cell.fill.fgColor && cell.fill.fgColor.theme != null) {
            cell.fill = resolveThemeFill(cell.fill, themeColors);
          }
        }
      }

      /* ── footer 병합 캡처 */
      const footerMerges1 = [];
      if (ws1._merges) {
        Object.values(ws1._merges).forEach(mg => {
          const { top, left, bottom, right } = mg.model;
          if (top > m1.end) footerMerges1.push({ top, left, bottom, right });
        });
      }

      /* ── 행 계획 빌드 (대분류·소분류만, 아이템 없음) */
      const plan1 = [];
      for (const group of smartGroups) {
        plan1.push({ type: 'main_group', name: group.name });
        const items = group.items || [];
        const hasCats = subSnap1 && items.some(it => it.cat && String(it.cat).trim());
        if (hasCats) {
          const seen = [], catMap = {};
          for (const it of items) {
            const c = String(it.cat || '').trim() || '기타';
            if (!catMap[c]) { catMap[c] = true; seen.push(c); }
          }
          for (const cat of seen) plan1.push({ type: 'sub_group', name: cat });
        }
      }

      /* ── ITEM_START 행에 헤더 내용이 있으면 보존 */
      const startRow1 = ws1.getRow(m1.start);
      let startHasContent1 = false;
      startRow1.eachCell({ includeEmpty: false }, cell => {
        const v = String(cell.value ?? '');
        if (!v.includes('{{') && v.trim() !== '') startHasContent1 = true;
      });
      if (startHasContent1) {
        startRow1.eachCell({ includeEmpty: false }, cell => {
          if (String(cell.value ?? '').includes('{{ITEM_START}}')) cell.value = null;
        });
      }

      const insertAt1    = startHasContent1 ? m1.start + 1 : m1.start;
      const deleteCount1 = m1.end - insertAt1 + 1;

      if (ws1._merges) {
        Object.keys(ws1._merges)
          .filter(k => { const { top, bottom } = ws1._merges[k].model; return top >= insertAt1 && bottom <= insertAt1 + deleteCount1 - 1; })
          .forEach(k => delete ws1._merges[k]);
      }
      ws1.spliceRows(insertAt1, deleteCount1);
      if (plan1.length > 0) ws1.spliceRows(insertAt1, 0, ...plan1.map(() => []));

      /* ── 행 스탬프 */
      let mainCounter1 = 0;
      const grandRefs1 = [];
      let curMain1 = null;
      const mainGroups1 = [];
      plan1.forEach((p, i) => {
        const rn = insertAt1 + i;
        if (p.type === 'main_group') {
          mainCounter1++;
          stampRow(ws1, rn, mainSnap1, {
            1: mainCounter1,
            2: p.name,
          }, 9);
          curMain1 = { rn, name: p.name, subRns: [] };
          mainGroups1.push(curMain1);
          grandRefs1.push(rn);
        } else if (p.type === 'sub_group' && subSnap1) {
          let parentName = null;
          for (let j = i - 1; j >= 0; j--) {
            if (plan1[j].type === 'main_group') { parentName = plan1[j].name; break; }
          }
          const subTotal = groupTotals[parentName]?.subs[p.name] || 0;
          stampRow(ws1, rn, subSnap1, {
            1: null,
            2: p.name,
            5: 1,
            6: subTotal > 0 ? subTotal : null,
            7: subTotal > 0 ? { formula: `F${rn}*E${rn}` } : null,
          }, 9);
          if (subTotal > 0) { ws1.getRow(rn).getCell(6).numFmt = NUM_FMT; ws1.getRow(rn).getCell(7).numFmt = NUM_FMT; }
          if (curMain1) curMain1.subRns.push(rn);
        }
      });

      /* ── 대분류 G열 = 소분류 G열 합계 수식 */
      for (const mg of mainGroups1) {
        const cell = ws1.getRow(mg.rn).getCell(7);
        if (mg.subRns.length > 0) {
          cell.value = { formula: `SUM(${mg.subRns.map(r => `G${r}`).join(',')})` };
        } else {
          const total = groupTotals[mg.name]?.total || 0;
          cell.value = total > 0 ? total : null;
        }
        cell.numFmt = NUM_FMT;
      }

      /* ── footer 병합 복원 + 합계 수식 */
      const netShift1 = plan1.length - deleteCount1;
      if (ws1._merges) {
        Object.keys(ws1._merges)
          .filter(k => ws1._merges[k].model.top >= insertAt1 + plan1.length)
          .forEach(k => delete ws1._merges[k]);
      }
      footerMerges1.forEach(({ top, left, bottom, right }) => {
        ws1.mergeCells(top + netShift1, left, bottom + netShift1, right);
      });

      if (m1.grandTotal > 0) {
        const gtRn1 = m1.grandTotal + netShift1;
        const gtCell1 = ws1.getRow(gtRn1).getCell(7);
        gtCell1.value  = grandRefs1.length > 0
          ? { formula: `SUM(${grandRefs1.map(r => `G${r}`).join(',')})` }
          : 0;
        gtCell1.numFmt = NUM_FMT;
        ws1.getRow(gtRn1).getCell(9).value = null;
      }

      /* ── 헤더 행 높이 복원 */
      for (const [rn, h] of Object.entries(headerHeights1)) {
        ws1.getRow(Number(rn)).height = h;
      }

      /* ── 잔여 마커 정리 */
      const markerRows1 = [];
      ws1.eachRow((row, rn) => {
        let has = false;
        row.eachCell({ includeEmpty: false }, cell => {
          if (String(cell.value ?? '').includes('{{')) has = true;
        });
        if (has) markerRows1.push(rn);
      });
      for (let i = markerRows1.length - 1; i >= 0; i--) {
        unmergeRow(ws1, markerRows1[i]);
        ws1.spliceRows(markerRows1[i], 1);
      }

      /* ── 열 너비 복원 (spliceRows 이후에 실행해야 리셋되지 않음) */
      for (const [c, w] of Object.entries(colWidths1)) {
        ws1.getColumn(Number(c)).width = w;
      }

      /* ── 인쇄 범위: H열까지만 (I열 마커 제외) */
      if (!ws1.pageSetup) ws1.pageSetup = {};
      ws1.pageSetup.printArea = `A1:H${ws1.rowCount || 9999}`;
    }
  }

  const generated = await wb.xlsx.writeBuffer();
  return copyDrawings(templateBuf, generated);
}

module.exports = { generateQuote, generateSmartQuote };
