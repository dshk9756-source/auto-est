const XLSX = require('xlsx');
const path = require('path');

const filePath = String.raw`C:\Users\user\Desktop\담당현장\#영업(입찰 확정 안 된 현장)\대우건설 홍천양수발전소\04. 공내역서_홍천양수발전소_R2_rev2.xlsx`;

const wb = XLSX.readFile(filePath, { cellFormula: true, cellStyles: true, bookVBA: true });

function setFormula(ws, addr, formula) {
  const existing = ws[addr];
  if (existing) {
    existing.f = formula;
    existing.t = 'n';
  } else {
    ws[addr] = { t: 'n', f: formula, v: 0 };
  }
}

function hasFormula(ws, addr) {
  const cell = ws[addr];
  return cell && cell.f && !cell.f.includes('#REF!');
}

function hasSumFormula(ws, addr) {
  const cell = ws[addr];
  return cell && cell.f && cell.f.toUpperCase().includes('SUM');
}

function fixTypo(ws, addr, oldText, newText) {
  const cell = ws[addr];
  if (cell && typeof cell.v === 'string' && cell.v.includes(oldText)) {
    cell.v = cell.v.split(oldText).join(newText);
    if (cell.w) delete cell.w;
    if (cell.r) delete cell.r;
    if (cell.h) delete cell.h;
    console.log(`  TYPO FIX ${addr}: "${oldText}" → "${newText}"`);
  }
}

let fixCount = 0;
function trackFix(sheet, addr, desc) {
  fixCount++;
  console.log(`  [${fixCount}] ${sheet} ${addr}: ${desc}`);
}

// =============================================
// SHEET 1: 공내역서
// =============================================
console.log('\n=== 공내역서 수정 ===');
const ws1 = wb.Sheets['1) 공내역서'];

// R8 (1.1.1 상부지): 참조 행 틀림 F9→F10, H9→H10, J9→J10
setFormula(ws1, 'E8', "'2) 상세내역서'!F10");
setFormula(ws1, 'G8', "'2) 상세내역서'!H10");
setFormula(ws1, 'I8', "'2) 상세내역서'!J10");
trackFix('공내역서', 'E8,G8,I8', '1.1.1 상부지 참조 수정 (F9→F10, H9→H10, J9→J10)');

// R9 (1.1.2 지상): 참조 행 틀림 F10→F25, H10→H25, J10→J25
setFormula(ws1, 'E9', "'2) 상세내역서'!F25");
setFormula(ws1, 'G9', "'2) 상세내역서'!H25");
setFormula(ws1, 'I9', "'2) 상세내역서'!J25");
trackFix('공내역서', 'E9,G9,I9', '1.1.2 지상 참조 수정 (F10→F25, H10→H25, J10→J25)');

// R10 (1.1.3 터널): #REF! → F37, H37, J37
setFormula(ws1, 'E10', "'2) 상세내역서'!F37");
setFormula(ws1, 'G10', "'2) 상세내역서'!H37");
setFormula(ws1, 'I10', "'2) 상세내역서'!J37");
trackFix('공내역서', 'E10,G10,I10', '1.1.3 터널 #REF! 수정 → F37, H37, J37');

// =============================================
// SHEET 2: 상세내역서
// =============================================
console.log('\n=== 상세내역서 수정 ===');
const ws2 = wb.Sheets['2) 상세내역서'];

// ---------- 소계 수식 수정 ----------
console.log('\n--- 소계 SUM 범위 수정 ---');

// R9 (1.1 통신네트워크 구축) - #REF! 제거
setFormula(ws2, 'F9', 'SUM(F10,F25,F37)');
setFormula(ws2, 'H9', 'SUM(H10,H25,H37)');
setFormula(ws2, 'J9', 'SUM(J10,J25,J37)');
setFormula(ws2, 'L9', 'SUM(L10,L25,L37)');
trackFix('상세내역서', 'R9', '1.1 통신네트워크 #REF! 제거 → SUM(10,25,37)');

// R37 (1.1.3 터널 구역) - SUM(F38) → SUM(F38:F48)
setFormula(ws2, 'F37', 'SUM(F38:F48)');
setFormula(ws2, 'H37', 'SUM(H38:H48)');
setFormula(ws2, 'J37', 'SUM(J38:J48)');
setFormula(ws2, 'L37', 'SUM(L38:L48)');
trackFix('상세내역서', 'R37', '1.1.3 터널 SUM 범위 수정 (38→38:48)');

// R49 (1.2 통신 요금) - SUM(F50) → SUM(F50:F53)
setFormula(ws2, 'F49', 'SUM(F50:F53)');
setFormula(ws2, 'H49', 'SUM(H50:H53)');
setFormula(ws2, 'J49', 'SUM(J50:J53)');
setFormula(ws2, 'L49', 'SUM(L50:L53)');
trackFix('상세내역서', 'R49', '1.2 통신요금 SUM 범위 수정 (50→50:53)');

// R54 (1.3 전기공사) - SUM(F55) → SUM(F55,F62)
setFormula(ws2, 'F54', 'SUM(F55,F62)');
setFormula(ws2, 'H54', 'SUM(H55,H62)');
setFormula(ws2, 'J54', 'SUM(J55,J62)');
setFormula(ws2, 'L54', 'SUM(L55,L62)');
trackFix('상세내역서', 'R54', '1.3 전기공사 SUM 범위 수정 (55→55,62)');

// R55 (1.3.1 태양광 구성) - SUM(F56) → SUM(F56:F61)
setFormula(ws2, 'F55', 'SUM(F56:F61)');
setFormula(ws2, 'H55', 'SUM(H56:H61)');
setFormula(ws2, 'J55', 'SUM(J56:J61)');
setFormula(ws2, 'L55', 'SUM(L56:L61)');
trackFix('상세내역서', 'R55', '1.3.1 태양광 SUM 범위 수정 (56→56:61)');

// R64 (1.4 전기요금) - SUM(F65) → SUM(F65:F66)
setFormula(ws2, 'F64', 'SUM(F65:F66)');
setFormula(ws2, 'H64', 'SUM(H65:H66)');
setFormula(ws2, 'J64', 'SUM(J65:J66)');
setFormula(ws2, 'L64', 'SUM(L65:L66)');
trackFix('상세내역서', 'R64', '1.4 전기요금 SUM 범위 수정 (65→65:66)');

// R70 (2.2 통합관제센터) - 선행 콤마 제거
setFormula(ws2, 'F70', 'SUM(F71,F78,F83,F86)');
setFormula(ws2, 'H70', 'SUM(H71,H78,H83,H86)');
setFormula(ws2, 'J70', 'SUM(J71,J78,J83,J86)');
setFormula(ws2, 'L70', 'SUM(L71,L78,L83,L86)');
trackFix('상세내역서', 'R70', '2.2 통합관제센터 SUM 선행콤마 제거');

// R71 (2.2.1 모니터링 장비) - SUM(F72) → SUM(F72:F77)
setFormula(ws2, 'F71', 'SUM(F72:F77)');
setFormula(ws2, 'H71', 'SUM(H72:H77)');
setFormula(ws2, 'J71', 'SUM(J72:J77)');
setFormula(ws2, 'L71', 'SUM(L72:L77)');
trackFix('상세내역서', 'R71', '2.2.1 모니터링장비 SUM 범위 수정 (72→72:77)');

// R78 (2.2.2 운영장치) - SUM(F79) → SUM(F79:F82)
setFormula(ws2, 'F78', 'SUM(F79:F82)');
setFormula(ws2, 'H78', 'SUM(H79:H82)');
setFormula(ws2, 'J78', 'SUM(J79:J82)');
setFormula(ws2, 'L78', 'SUM(L79:L82)');
trackFix('상세내역서', 'R78', '2.2.2 운영장치 SUM 범위 수정 (79→79:82)');

// R83 (2.2.3 관제 데스크) - SUM(F84) → SUM(F84:F85)
setFormula(ws2, 'F83', 'SUM(F84:F85)');
setFormula(ws2, 'H83', 'SUM(H84:H85)');
setFormula(ws2, 'J83', 'SUM(J84:J85)');
setFormula(ws2, 'L83', 'SUM(L84:L85)');
trackFix('상세내역서', 'R83', '2.2.3 관제데스크 SUM 범위 수정 (84→84:85)');

// R86 (2.2.4 영상저장장치) - SUM(F87) → SUM(F87:F94)
setFormula(ws2, 'F86', 'SUM(F87:F94)');
setFormula(ws2, 'H86', 'SUM(H87:H94)');
setFormula(ws2, 'J86', 'SUM(J87:J94)');
setFormula(ws2, 'L86', 'SUM(L87:L94)');
trackFix('상세내역서', 'R86', '2.2.4 영상저장장치 SUM 범위 수정 (87→87:94)');

// R95 (2.3 방송시스템) - H, J 범위 수정
// F95=SUM(F96:F111) ✓, L95=SUM(L96:L111) ✓
setFormula(ws2, 'H95', 'SUM(H96:H111)');
setFormula(ws2, 'J95', 'SUM(J96:J111)');
trackFix('상세내역서', 'R95', '2.3 방송시스템 H,J SUM 범위 수정 (96:97→96:111)');

// R112 (2.4 VPN장비) - SUM(F113) → SUM(F113:F114)
setFormula(ws2, 'F112', 'SUM(F113:F114)');
setFormula(ws2, 'H112', 'SUM(H113:H114)');
setFormula(ws2, 'J112', 'SUM(J113:J114)');
setFormula(ws2, 'L112', 'SUM(L113:L114)');
trackFix('상세내역서', 'R112', '2.4 VPN SUM 범위 수정 (113→113:114)');

// R163 (3.4 이동식 CCTV 태양광) - 소계 수식 신규 추가
setFormula(ws2, 'F163', 'SUM(F164:F171)');
setFormula(ws2, 'H163', 'SUM(H164:H171)');
setFormula(ws2, 'J163', 'SUM(J164:J171)');
setFormula(ws2, 'L163', 'SUM(L164:L171)');
trackFix('상세내역서', 'R163', '3.4 태양광CCTV 소계 수식 신규 추가');

// R315 (10.1 보건관리용 키오스크) - F,H,J 범위 수정
setFormula(ws2, 'F315', 'SUM(F316:F317)');
setFormula(ws2, 'H315', 'SUM(H316:H317)');
setFormula(ws2, 'J315', 'SUM(J316:J317)');
trackFix('상세내역서', 'R315', '10.1 키오스크 F,H,J SUM 범위 수정 (316→316:317)');

// R318 (10.2 스마트 혈압계) - F,H,J 범위 수정
setFormula(ws2, 'F318', 'SUM(F319:F320)');
setFormula(ws2, 'H318', 'SUM(H319:H320)');
setFormula(ws2, 'J318', 'SUM(J319:J320)');
trackFix('상세내역서', 'R318', '10.2 혈압계 F,H,J SUM 범위 수정 (319→319:320)');

// ---------- 개별 항목 누락 수식 추가 ----------
console.log('\n--- 개별 항목 누락 수식 추가 ---');

const subtotalRows = new Set([
  1, 2, 3, 4, 5, 6, 7,
  8, 9, 10, 25, 37, 49, 54, 55, 62, 64, 67, 68, 70, 71, 78, 83, 86, 95, 112, 115,
  122, 123, 138, 156, 163, 172, 175, 176, 191, 201, 210, 219, 226, 227, 235, 253,
  261, 262, 269, 273, 274, 282, 293, 300, 301, 303, 304, 306, 308, 310, 312, 314,
  315, 318, 321, 323, 324, 327, 332, 335, 336, 338, 340, 343
]);

const range = XLSX.utils.decode_range(ws2['!ref']);
let itemFixCount = 0;

for (let r = 6; r <= range.e.r; r++) {
  const rowNum = r + 1;
  if (subtotalRows.has(rowNum)) continue;

  const cCell = ws2[XLSX.utils.encode_cell({r, c: 2})];
  if (!cCell || !cCell.v) continue;

  const dCell = ws2[XLSX.utils.encode_cell({r, c: 3})];
  const hasD = dCell && dCell.v !== undefined && dCell.v !== null && dCell.v !== '';

  const eCell = ws2[`E${rowNum}`];
  const gCell = ws2[`G${rowNum}`];
  const iCell = ws2[`I${rowNum}`];
  const hasE = eCell && (eCell.v !== undefined && eCell.v !== null);
  const hasG = gCell && (gCell.v !== undefined && gCell.v !== null);
  const hasI = iCell && (iCell.v !== undefined && iCell.v !== null);

  let rowFixed = false;
  const fixes = [];

  // F = D*E (재료비 금액)
  if (!hasFormula(ws2, `F${rowNum}`)) {
    setFormula(ws2, `F${rowNum}`, `D${rowNum}*E${rowNum}`);
    fixes.push('F');
    rowFixed = true;
  }

  // H = D*G (노무비 금액)
  if (!hasFormula(ws2, `H${rowNum}`)) {
    setFormula(ws2, `H${rowNum}`, `D${rowNum}*G${rowNum}`);
    fixes.push('H');
    rowFixed = true;
  }

  // J = D*I (경비 금액)
  if (!hasFormula(ws2, `J${rowNum}`)) {
    setFormula(ws2, `J${rowNum}`, `D${rowNum}*I${rowNum}`);
    fixes.push('J');
    rowFixed = true;
  }

  // K = E+G+I (합계 단가)
  if (!hasFormula(ws2, `K${rowNum}`)) {
    setFormula(ws2, `K${rowNum}`, `E${rowNum}+G${rowNum}+I${rowNum}`);
    fixes.push('K');
    rowFixed = true;
  }

  // L = D*K (합계 금액)
  if (!hasFormula(ws2, `L${rowNum}`)) {
    setFormula(ws2, `L${rowNum}`, `D${rowNum}*K${rowNum}`);
    fixes.push('L');
    rowFixed = true;
  }

  if (rowFixed) {
    const aCell = ws2[`A${rowNum}`];
    const itemName = aCell ? String(aCell.v).substring(0, 25) : `Row${rowNum}`;
    trackFix('상세내역서', `R${rowNum}`, `${itemName} → 수식 추가: ${fixes.join(',')}`);
    itemFixCount++;
  }
}

console.log(`\n  총 ${itemFixCount}개 항목행 수식 보완 완료`);

// ---------- 오타 수정 ----------
console.log('\n--- 오타 수정 ---');

// "양뱡항" → "양방향"
for (let r = 0; r <= range.e.r; r++) {
  const addr = `B${r + 1}`;
  fixTypo(ws2, addr, '양뱡항', '양방향');
}

// "Deceter" → "Detector"
for (let r = 0; r <= range.e.r; r++) {
  const addr = `B${r + 1}`;
  fixTypo(ws2, addr, 'Deceter', 'Detector');
}

// "휠터" → "필터"
for (let r = 0; r <= range.e.r; r++) {
  const addr = `B${r + 1}`;
  fixTypo(ws2, addr, '휠터', '필터');
}

// ---------- 저장 ----------
console.log('\n=== 파일 저장 ===');

const outputPath = filePath.replace('.xlsx', '_수정완료.xlsx');
XLSX.writeFile(wb, outputPath, { cellStyles: true, bookVBA: true });
console.log(`저장 완료: ${outputPath}`);
console.log(`총 수정 건수: ${fixCount}`);
