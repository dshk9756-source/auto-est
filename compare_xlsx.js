const XLSX = require('xlsx');
const path = require('path');

const base = 'C:\\Users\\user\\Desktop\\담당현장\\#영업(입찰 확정 안 된 현장)\\대우건설 홍천양수발전소';
const f1 = path.join(base, '04. 공내역서_홍천양수발전소_R2_rev2.xlsx');
const f2 = path.join(base, '대우건설(홍천 양수발전소)260601.xlsx');

const wb1 = XLSX.readFile(f1);
const wb2 = XLSX.readFile(f2);

console.log('=== rev2 시트 ===');
console.log(wb1.SheetNames);
console.log('=== 260601 시트 ===');
console.log(wb2.SheetNames);

// Compare each matching sheet
for (const sheetName of wb1.SheetNames) {
  if (!wb2.SheetNames.includes(sheetName)) {
    console.log(`\n[!] 시트 "${sheetName}" - rev2에만 존재`);
    continue;
  }

  const s1 = wb1.Sheets[sheetName];
  const s2 = wb2.Sheets[sheetName];

  const range1 = XLSX.utils.decode_range(s1['!ref'] || 'A1');
  const range2 = XLSX.utils.decode_range(s2['!ref'] || 'A1');

  const maxRow = Math.max(range1.e.r, range2.e.r);
  const maxCol = Math.max(range1.e.c, range2.e.c);

  const diffs = [];

  for (let r = 0; r <= maxRow; r++) {
    for (let c = 0; c <= maxCol; c++) {
      const addr = XLSX.utils.encode_cell({r, c});
      const c1 = s1[addr];
      const c2 = s2[addr];
      const v1 = c1 ? (c1.v !== undefined ? c1.v : '') : '';
      const v2 = c2 ? (c2.v !== undefined ? c2.v : '') : '';

      if (String(v1) !== String(v2)) {
        diffs.push({cell: addr, row: r+1, col: XLSX.utils.encode_col(c), rev2: v1, new260601: v2});
      }
    }
  }

  console.log(`\n========== 시트: "${sheetName}" ==========`);
  console.log(`변경된 셀 수: ${diffs.length}`);

  if (diffs.length > 0 && diffs.length <= 500) {
    for (const d of diffs) {
      console.log(`  [${d.cell}] rev2: "${d.rev2}" → 260601: "${d.new260601}"`);
    }
  } else if (diffs.length > 500) {
    console.log('  (변경사항이 너무 많아 주요 변경만 표시)');
    // Show first 100
    for (let i = 0; i < 100; i++) {
      const d = diffs[i];
      console.log(`  [${d.cell}] rev2: "${d.rev2}" → 260601: "${d.new260601}"`);
    }
    console.log(`  ... 외 ${diffs.length - 100}건`);
  }
}

// Check sheets only in 260601
for (const sheetName of wb2.SheetNames) {
  if (!wb1.SheetNames.includes(sheetName)) {
    console.log(`\n[!] 시트 "${sheetName}" - 260601에만 존재`);
  }
}
