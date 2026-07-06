const XLSX = require('xlsx');
const filePath = 'C:\\Users\\user\\Desktop\\담당현장\\#영업(입찰 확정 안 된 현장)\\대우건설 홍천양수발전소\\04. 공내역서_홍천양수발전소_R2_rev2_수정완료.xlsx';
const wb = XLSX.readFile(filePath, {cellFormula: true});

console.log('=== 공내역서 검증 ===');
const ws1 = wb.Sheets['1) 공내역서'];
for (const r of [8,9,10]) {
  for (const c of ['E','G','I']) {
    const cell = ws1[c+r];
    console.log(c+r + ': ' + (cell ? cell.f : 'MISSING'));
  }
}

console.log('\n=== 상세내역서 검증 ===');
const ws2 = wb.Sheets['2) 상세내역서'];

console.log('\n--- 소계 SUM 수식 ---');
const subtotals = [9,37,49,54,55,64,70,71,78,83,86,95,112,163,315,318];
for (const r of subtotals) {
  const fmls = ['F','H','J','L'].map(c => {
    const cell = ws2[c+r];
    return c+': '+(cell ? cell.f : 'MISSING');
  });
  console.log('R'+r+': '+fmls.join(' | '));
}

console.log('\n--- #REF! 검사 ---');
let refErrors = 0;
const range = XLSX.utils.decode_range(ws2['!ref']);
for (let r2=0; r2<=range.e.r; r2++) {
  for (let c2=0; c2<=range.e.c; c2++) {
    const addr = XLSX.utils.encode_cell({r:r2,c:c2});
    const cell = ws2[addr];
    if (cell && cell.f && cell.f.includes('#REF!')) {
      console.log('  #REF! at '+addr+': '+cell.f);
      refErrors++;
    }
  }
}
console.log(refErrors === 0 ? '  #REF! 에러 없음 OK' : '  #REF! 에러 '+refErrors+'건!');

console.log('\n--- K,L열 누락 검사 ---');
let missingKL = 0;
const skipRows = new Set([1,2,3,4,5,6,7,8,9,10,25,37,49,54,55,62,64,67,68,70,71,78,83,86,95,112,115,122,123,138,156,163,172,175,176,191,201,210,219,226,227,235,253,261,262,269,273,274,282,293,300,301,303,304,306,308,310,312,314,315,318,321,323,324,327,332,335,336,338,340,343]);
for (let r2=6; r2<=range.e.r; r2++) {
  const rowNum = r2+1;
  if (skipRows.has(rowNum)) continue;
  const cCell = ws2[XLSX.utils.encode_cell({r:r2,c:2})];
  const dCell = ws2[XLSX.utils.encode_cell({r:r2,c:3})];
  if (!cCell || !cCell.v) continue;
  if (!dCell || dCell.v === undefined) continue;
  const kCell = ws2['K'+rowNum];
  const lCell = ws2['L'+rowNum];
  if (!kCell || !kCell.f) { console.log('  K'+rowNum+' missing'); missingKL++; }
  if (!lCell || !lCell.f) { console.log('  L'+rowNum+' missing'); missingKL++; }
}
console.log(missingKL === 0 ? '  K,L열 모두 정상 OK' : '  누락 '+missingKL+'건');

console.log('\n--- 오타 검사 ---');
let typos = 0;
for (let r2=0; r2<=range.e.r; r2++) {
  const cell = ws2['B'+(r2+1)];
  if (cell && typeof cell.v === 'string') {
    ['양뱡항','Deceter','휠터'].forEach(t => {
      if (cell.v.includes(t)) { console.log('  B'+(r2+1)+': '+t+' 잔존'); typos++; }
    });
  }
}
console.log(typos === 0 ? '  오타 없음 OK' : '  오타 '+typos+'건');

console.log('\n=== 검증 완료 ===');
