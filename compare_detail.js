const XLSX = require('xlsx');
const path = require('path');

const base = 'C:\\Users\\user\\Desktop\\담당현장\\#영업(입찰 확정 안 된 현장)\\대우건설 홍천양수발전소';
const f1 = path.join(base, '04. 공내역서_홍천양수발전소_R2_rev2.xlsx');
const f2 = path.join(base, '대우건설(홍천 양수발전소)260601.xlsx');

const wb1 = XLSX.readFile(f1);
const wb2 = XLSX.readFile(f2);

// ===== 1) 공내역서 분석 =====
console.log('╔════════════════════════════════════════════════════════════╗');
console.log('║            1) 공내역서 주요 변경사항 비교                 ║');
console.log('╚════════════════════════════════════════════════════════════╝');

const s1 = wb1.Sheets['1) 공내역서'];
const s2 = wb2.Sheets['1) 공내역서'];

// Read as JSON to understand structure
const d1 = XLSX.utils.sheet_to_json(s1, {header:1, defval:''});
const d2 = XLSX.utils.sheet_to_json(s2, {header:1, defval:''});

console.log('\n[공내역서 구조 - rev2 헤더]');
for (let i = 0; i < Math.min(5, d1.length); i++) {
  console.log(`Row ${i+1}: ${JSON.stringify(d1[i].slice(0, 13))}`);
}

console.log('\n[공내역서 구조 - 260601 헤더]');
for (let i = 0; i < Math.min(5, d2.length); i++) {
  console.log(`Row ${i+1}: ${JSON.stringify(d2[i].slice(0, 13))}`);
}

// Compare item names and amounts
console.log('\n\n--- 공내역서: 항목명 변경 ---');
for (let i = 0; i < Math.max(d1.length, d2.length); i++) {
  const r1 = d1[i] || [];
  const r2 = d2[i] || [];
  const name1 = String(r1[0] || '').trim();
  const name2 = String(r2[0] || '').trim();
  if (name1 && name2 && name1 !== name2) {
    console.log(`  Row ${i+1}: "${name1}" → "${name2}"`);
  }
}

console.log('\n\n--- 공내역서: 금액 변경 (합계열 L열 기준) ---');
for (let i = 0; i < Math.max(d1.length, d2.length); i++) {
  const r1 = d1[i] || [];
  const r2 = d2[i] || [];
  const name = String(r1[0] || r2[0] || '').trim();
  const amt1 = Number(r1[11]) || 0; // L열 (index 11) = 합계금액
  const amt2 = Number(r2[11]) || 0;
  if (name && amt1 !== amt2 && (amt1 > 0 || amt2 > 0)) {
    const diff = amt2 - amt1;
    const pct = amt1 !== 0 ? ((diff / amt1) * 100).toFixed(1) : 'NEW';
    console.log(`  Row ${i+1} [${name}]`);
    console.log(`    rev2: ${amt1.toLocaleString()}원 → 260601: ${amt2.toLocaleString()}원 (차이: ${diff > 0 ? '+' : ''}${diff.toLocaleString()}원, ${pct}%)`);
  }
}

// Total comparison
console.log('\n\n--- 최종 합계 비교 ---');
const lastRow1 = d1[d1.length - 1] || [];
const lastRow2 = d2[d2.length - 1] || [];
// Find total rows
for (let i = d1.length - 5; i < d1.length; i++) {
  const r1 = d1[i] || [];
  const name = String(r1[0] || '').trim();
  if (name) {
    const r2 = d2[i] || [];
    console.log(`  ${name}: rev2=${(Number(r1[11])||0).toLocaleString()} → 260601=${(Number(r2[11])||0).toLocaleString()}`);
  }
}

// ===== 2) 상세내역서 분석 =====
console.log('\n\n╔════════════════════════════════════════════════════════════╗');
console.log('║            2) 상세내역서 주요 변경사항 비교               ║');
console.log('╚════════════════════════════════════════════════════════════╝');

const ss1 = wb1.Sheets['2) 상세내역서'];
const ss2 = wb2.Sheets['2) 상세내역서'];
const dd1 = XLSX.utils.sheet_to_json(ss1, {header:1, defval:''});
const dd2 = XLSX.utils.sheet_to_json(ss2, {header:1, defval:''});

console.log('\n[상세내역서 구조 - rev2 헤더]');
for (let i = 0; i < Math.min(5, dd1.length); i++) {
  console.log(`Row ${i+1}: ${JSON.stringify(dd1[i].slice(0, 13))}`);
}
console.log('\n[상세내역서 구조 - 260601 헤더]');
for (let i = 0; i < Math.min(5, dd2.length); i++) {
  console.log(`Row ${i+1}: ${JSON.stringify(dd2[i].slice(0, 13))}`);
}

// Find items where name changed
console.log('\n\n--- 상세내역서: 항목명 변경 ---');
for (let i = 0; i < Math.max(dd1.length, dd2.length); i++) {
  const r1 = dd1[i] || [];
  const r2 = dd2[i] || [];
  const name1 = String(r1[0] || '').trim();
  const name2 = String(r2[0] || '').trim();
  if (name1 && name2 && name1 !== name2) {
    console.log(`  Row ${i+1}: "${name1}" → "${name2}"`);
  }
}

// Find items where total amount (L col) changed
console.log('\n\n--- 상세내역서: 금액 변경 (합계금액 L열 기준) ---');
for (let i = 0; i < Math.max(dd1.length, dd2.length); i++) {
  const r1 = dd1[i] || [];
  const r2 = dd2[i] || [];
  const name = String(r1[0] || r2[0] || '').trim();
  const amt1 = Number(r1[11]) || 0;
  const amt2 = Number(r2[11]) || 0;
  if (name && amt1 !== amt2 && (amt1 > 0 || amt2 > 0)) {
    const diff = amt2 - amt1;
    const pct = amt1 !== 0 ? ((diff / amt1) * 100).toFixed(1) : 'NEW';
    console.log(`  Row ${i+1} [${name}]`);
    console.log(`    rev2: ${amt1.toLocaleString()}원 → 260601: ${amt2.toLocaleString()}원 (차이: ${diff > 0 ? '+' : ''}${diff.toLocaleString()}원, ${pct}%)`);
  }
}

// Summary of cost categories
console.log('\n\n--- 상세내역서: 소계/합계 행 비교 ---');
for (let i = 0; i < dd1.length; i++) {
  const r1 = dd1[i] || [];
  const name = String(r1[0] || '').trim();
  if (name.includes('소계') || name.includes('합계') || name.includes('총합계') || name.includes('계')) {
    const r2 = dd2[i] || [];
    const amt1 = Number(r1[11]) || 0;
    const amt2 = Number(r2[11]) || 0;
    if (amt1 !== amt2) {
      console.log(`  Row ${i+1} [${name}]: rev2=${amt1.toLocaleString()} → 260601=${amt2.toLocaleString()} (${amt2 > amt1 ? '+' : ''}${(amt2-amt1).toLocaleString()})`);
    }
  }
}
