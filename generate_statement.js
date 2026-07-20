'use strict';
/**
 * generate_statement.js
 * ─────────────────────────────────────────────────────────────────────────
 * 거래명세서 생성기. template/거래명세서.xlsx 의 {{마커}}만 해석하고
 * 서식(병합·폰트·테두리)은 템플릿 그대로 유지한다.
 *
 * 마커
 *  {{견적일자}} {{문서번호}}
 *  {{거래처_등록번호}} {{거래처_상호}} {{거래처_대표자}}
 *  {{거래처_주소}} {{거래처_업태}} {{거래처_종목}}
 *  {{ITEM_START}} {{ITEM_ROW}} {{ITEM_END}}
 *
 * 품목 행 컬럼 (병합 기준)
 *  A:C=품목명  D=단위(EA·식 등)  E=수량  F:G=단가  H:I=금액(수식 E*F)  J=비고
 */
const ExcelJS = require('exceljs');
const fs      = require('fs');
const path    = require('path');

const TEMPLATE_PATH = path.join(__dirname, 'template', '거래명세서.xlsx');
const NUM_FMT        = '#,##0';
const MAX_COL         = 10; // A~J

function snapshotRow(ws, rowNum, maxCol = MAX_COL) {
  const row   = ws.getRow(rowNum);
  const cells = {};
  for (let c = 1; c <= maxCol; c++) {
    const cell = row.getCell(c);
    cells[c] = {
      font:      cell.font      ? JSON.parse(JSON.stringify(cell.font))      : undefined,
      fill:      cell.fill      ? JSON.parse(JSON.stringify(cell.fill))      : undefined,
      border:    cell.border    ? JSON.parse(JSON.stringify(cell.border))    : undefined,
      alignment: cell.alignment ? JSON.parse(JSON.stringify(cell.alignment)) : undefined,
      numFmt:    cell.numFmt    || undefined,
      value:     cell.value,
    };
  }
  const merges = [];
  if (ws._merges) {
    Object.values(ws._merges).forEach(m => {
      const { left, top, right, bottom } = m.model;
      if (top === rowNum && bottom === rowNum && right > left) merges.push({ left, right });
    });
  }
  return { height: row.height, cells, merges };
}

function unmergeRow(ws, rowNum) {
  if (!ws._merges) return;
  const keys = Object.keys(ws._merges).filter(k => {
    const { top, bottom } = ws._merges[k].model;
    return top <= rowNum && bottom >= rowNum;
  });
  keys.forEach(k => delete ws._merges[k]);
}

function stampRow(ws, dstRn, snap, values = {}, maxCol = MAX_COL) {
  unmergeRow(ws, dstRn);
  const row = ws.getRow(dstRn);
  if (snap.height != null && snap.height > 0) row.height = snap.height;

  for (let c = 1; c <= maxCol; c++) {
    const cell = row.getCell(c);
    const s    = snap.cells[c] || {};
    if (s.font)      cell.font      = JSON.parse(JSON.stringify(s.font));
    if (s.fill)      cell.fill      = JSON.parse(JSON.stringify(s.fill));
    if (s.border)    cell.border    = JSON.parse(JSON.stringify(s.border));
    if (s.alignment) cell.alignment = JSON.parse(JSON.stringify(s.alignment));
    if (s.numFmt)    cell.numFmt    = s.numFmt;

    if (Object.prototype.hasOwnProperty.call(values, c)) {
      cell.value = values[c];
    } else {
      const sv = s.value;
      const isMarker = sv && typeof sv === 'string' && sv.includes('{{');
      cell.value = isMarker ? null : (sv ?? null);
    }
  }
  snap.merges.forEach(({ left, right }) => ws.mergeCells(dstRn, left, dstRn, right));
}

/**
 * 병합/값은 건드리지 않고 셀 스타일(테두리 등)만 다시 입힌다.
 * ExcelJS의 mergeCells()는 병합 범위 안의 슬레이브 셀 스타일을 마스터 셀
 * 것으로 덮어써버리므로, 병합을 복원한 뒤 이 함수로 각 셀의 원래 테두리를
 * 다시 씌워야 한다(merge → style 순서, stampRow처럼 unmerge하면 안 됨).
 */
function applyRowStyle(ws, rn, snap, maxCol = MAX_COL) {
  const row = ws.getRow(rn);
  for (let c = 1; c <= maxCol; c++) {
    const cell = row.getCell(c);
    const s    = snap.cells[c] || {};
    if (s.font)      cell.font      = JSON.parse(JSON.stringify(s.font));
    if (s.fill)      cell.fill      = JSON.parse(JSON.stringify(s.fill));
    if (s.border)    cell.border    = JSON.parse(JSON.stringify(s.border));
    if (s.alignment) cell.alignment = JSON.parse(JSON.stringify(s.alignment));
    if (s.numFmt)    cell.numFmt    = s.numFmt;
  }
}

function findMarkers(ws) {
  const m = { start: -1, itemRow: -1, end: -1 };
  ws.eachRow((row, rn) => {
    row.eachCell({ includeEmpty: false }, cell => {
      const v = String(cell.value ?? '');
      if (v.includes('{{ITEM_START}}')) m.start   = rn;
      if (v.includes('{{ITEM_ROW}}'))   m.itemRow = rn;
      if (v.includes('{{ITEM_END}}'))   m.end     = rn;
    });
  });
  return m;
}

function genDocNo(date) {
  const d = (date || '').replace(/-/g, '');
  return `${d || 'DATE'}-001`;
}

/**
 * 세트(패키지)의 구성 품목은 세트명 하나로 합치고, 단품으로 담은 품목은
 * 서로 카테고리가 같아도 각자 자기 이름으로 별도 줄을 유지한다.
 * 품목 개수와 무관하게 항상 수량 1 · 단위 "식" · 금액은 합계로 표시한다.
 */
function aggregateByGroup(items) {
  const order = [];
  const map = new Map();
  for (const it of items) {
    const key   = it.isSet ? (it.group || it.name || '기타') : (it.itemId || it.name || '기타');
    const label = it.isSet ? (it.group || it.name || '기타') : (it.name || '기타');
    if (!map.has(key)) { map.set(key, { label, members: [] }); order.push(key); }
    map.get(key).members.push(it);
  }
  const result = [];
  for (const key of order) {
    const { label, members } = map.get(key);
    const amount = members.reduce((sum, it) => {
      const qty   = it.qty || 1;
      const price = typeof it.effectiveP === 'number' ? it.effectiveP : (it.price || 0);
      return sum + price * qty;
    }, 0);
    result.push({ name: label, unit: '식', qty: 1, effectiveP: amount, note: '' });
  }
  return result;
}

async function generateStatement(data) {
  const templateBuf = fs.readFileSync(TEMPLATE_PATH);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(templateBuf);

  const ws = wb.worksheets[0];
  if (!ws) throw new Error('거래명세서 템플릿에서 시트를 찾을 수 없습니다.');

  /* ── 마커 탐색 */
  const m = findMarkers(ws);
  if (m.start < 0 || m.itemRow < 0 || m.end < 0)
    throw new Error(`거래명세서 마커를 찾지 못했습니다: ${JSON.stringify(m)}`);

  /* ── 헤더 플레이스홀더 치환 (ITEM_START 행 이전) */
  const buyer = data.buyer || {};
  const phMap = {
    '{{견적일자}}':         data.date || '',
    '{{문서번호}}':         genDocNo(data.date),
    '{{거래처_등록번호}}':  buyer.bizNo   || '',
    '{{거래처_상호}}':      buyer.name    || data.company || data.client || '',
    '{{거래처_대표자}}':    buyer.ceo     || '',
    '{{거래처_주소}}':      buyer.address || '',
    '{{거래처_업태}}':      buyer.bizType || '',
    '{{거래처_종목}}':      buyer.bizItem || '',
  };
  ws.eachRow((row, rn) => {
    if (rn >= m.start) return;
    row.eachCell({ includeEmpty: false }, cell => {
      if (typeof cell.value !== 'string') return;
      let v = cell.value;
      for (const [ph, val] of Object.entries(phMap)) if (v.includes(ph)) v = v.replace(ph, val);
      cell.value = v;
    });
  });

  /* ── footer(소계·공급가액·부가세·총계) 병합 사전 캡처 */
  const footerMergesOrig = [];
  if (ws._merges) {
    Object.values(ws._merges).forEach(mg => {
      const { top, left, bottom, right } = mg.model;
      if (top > m.end) footerMergesOrig.push({ top, left, bottom, right });
    });
  }

  /* ── footer 4개 행(소계·특기사항+공급가액·부가세·총계) 스타일 스냅샷
        spliceRows는 이동되는 행의 병합뿐 아니라 바깥쪽 굵은 테두리 같은
        스타일도 깨뜨릴 수 있어, 원본 템플릿 스타일을 미리 찍어둔다. */
  const footerRowSnaps = [1, 2, 3, 4].map(off => snapshotRow(ws, m.end + off));

  /* ── 아이템 행 스타일 스냅샷 (삭제 전) */
  const itemRowSnap = snapshotRow(ws, m.itemRow);

  /* ── 마커 구역(ITEM_START~ITEM_END) 삭제 후 재생성
        품목 수가 원본 템플릿 행 수(deleteCount)보다 적으면 템플릿 그대로의
        빈 행으로 채워 명세서가 줄어들지 않게 하고, 더 많으면 그만큼 늘린다. */
  const rawItems = (data.items || []).filter(it => it.type !== 'smart_group');
  const items = aggregateByGroup(rawItems);
  const insertAt    = m.start;
  const deleteCount = m.end - m.start + 1;
  const rowCount    = Math.max(items.length, deleteCount);

  if (ws._merges) {
    const lastDel = insertAt + deleteCount - 1;
    Object.keys(ws._merges)
      .filter(k => { const { top, bottom } = ws._merges[k].model; return top >= insertAt && bottom <= lastDel; })
      .forEach(k => delete ws._merges[k]);
  }
  ws.spliceRows(insertAt, deleteCount);
  if (rowCount > 0) ws.spliceRows(insertAt, 0, ...Array(rowCount).fill(0).map(() => []));

  for (let i = 0; i < rowCount; i++) {
    const rn = insertAt + i;
    const it = items[i];
    if (!it) {
      // value가 전부 null인 행은 ExcelJS가 테두리를 누락시킬 수 있어 공백으로 채운다
      const blankVals = {};
      for (let c = 1; c <= MAX_COL; c++) blankVals[c] = ' ';
      stampRow(ws, rn, itemRowSnap, blankVals);
      continue;
    }
    const qty   = it.qty || 1;
    const price = typeof it.effectiveP === 'number' ? it.effectiveP : (it.price || 0);
    stampRow(ws, rn, itemRowSnap, {
      1: it.name || '',
      4: it.unit || '',
      5: qty,
      6: price > 0 ? price : 0,
      8: { formula: `E${rn}*F${rn}` },
      10: it.note || '',
    });
    ws.getRow(rn).getCell(5).numFmt = '#,##0';
    ws.getRow(rn).getCell(6).numFmt = NUM_FMT;
    ws.getRow(rn).getCell(8).numFmt = NUM_FMT;
  }

  /* ── footer 병합 복원 (netShift 만큼 이동) */
  const netShift = rowCount - deleteCount;
  if (ws._merges) {
    Object.keys(ws._merges)
      .filter(k => ws._merges[k].model.top >= insertAt + rowCount)
      .forEach(k => delete ws._merges[k]);
  }
  footerMergesOrig.forEach(({ top, left, bottom, right }) => {
    ws.mergeCells(top + netShift, left, bottom + netShift, right);
  });

  /* ── footer 4개 행 테두리 재적용 ──────────────────────────────────────
        mergeCells()가 병합 범위의 슬레이브 셀 테두리를 마스터 셀 것으로
        덮어쓰므로, 병합을 먼저 복원한 뒤 각 셀의 원래 테두리를 다시 씌운다. */
  [1, 2, 3, 4].forEach((off, idx) => {
    applyRowStyle(ws, m.end + off + netShift, footerRowSnaps[idx]);
  });

  /* ── 소계 / 공급가액 / 부가가치세 / 총계 수식 재계산
        원본 템플릿 기준 위치: 소계 = ITEM_END+1, 공급가액 = +2, 부가세 = +3, 총계 = +4 */
  const subtotalRow = m.end + 1 + netShift;
  const supplyRow    = subtotalRow + 1;
  const vatRow        = subtotalRow + 2;
  const totalRow      = subtotalRow + 3;

  const subCell = ws.getRow(subtotalRow).getCell(8); // H열
  subCell.value = rowCount > 0
    ? { formula: `SUM(H${insertAt}:I${insertAt + rowCount - 1})` }
    : 0;
  subCell.numFmt = NUM_FMT;

  const supplyCell = ws.getRow(supplyRow).getCell(9); // I열
  supplyCell.value = { formula: `H${subtotalRow}` };
  supplyCell.numFmt = NUM_FMT;

  const vatCell = ws.getRow(vatRow).getCell(9);
  vatCell.value = { formula: `I${supplyRow}*0.1` };
  vatCell.numFmt = NUM_FMT;

  const totalCell = ws.getRow(totalRow).getCell(9);
  totalCell.value = { formula: `I${supplyRow}+I${vatRow}` };
  totalCell.numFmt = NUM_FMT;

  /* ── 유령 행 정리 ─────────────────────────────────────────────────────
        spliceRows(start, count)는 count>1로 시트 끝부분을 한 번에 삭제하면
        내부 배열을 제대로 축소하지 못해 스타일만 남은 빈 행 · 마커 잔재가
        꼬리에 남는다. 끝에서부터 한 행씩 삭제하면 정상 동작한다. */
  for (let r = ws.rowCount; r > totalRow; r--) ws.spliceRows(r, 1);

  /* ── 인쇄 범위 (A1 ~ J{총계행}) */
  ws.pageSetup = {
    ...(ws.pageSetup || {}),
    printArea: `A1:J${totalRow}`,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    paperSize: 9,
    orientation: 'portrait',
  };

  return wb.xlsx.writeBuffer();
}

module.exports = { generateStatement };
