const ExcelJS = require('exceljs');

async function create() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('시공일정표');

  ws.getColumn(1).width = 12;
  ws.getColumn(2).width = 32;
  ws.getColumn(3).width = 6;
  for (let c = 4; c <= 13; c++) ws.getColumn(c).width = 7;

  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9999CC' } };
  const headerFont = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
  const boldFont = { name: '맑은 고딕', bold: true, size: 10 };
  const normalFont = { name: '맑은 고딕', size: 10 };
  const centerAlign = { horizontal: 'center', vertical: 'middle', wrapText: true };
  const thinBorder = {
    top: { style: 'thin' }, bottom: { style: 'thin' },
    left: { style: 'thin' }, right: { style: 'thin' }
  };
  const redFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };

  // === HEADERS ===
  ws.mergeCells('A1:B3');
  ws.getCell('A1').value = '구분';
  ws.mergeCells('C1:C3');
  ws.getCell('C1').value = '수량';
  ws.mergeCells('D1:M1');
  ws.getCell('D1').value = '2025';
  ws.mergeCells('D2:M2');
  ws.getCell('D2').value = '4월';

  const days = [9, 10, 11, 12, 13, 14, 15, 16, 17, 18];
  days.forEach((d, i) => { ws.getCell(3, 4 + i).value = d; });

  for (let r = 1; r <= 3; r++) {
    for (let c = 1; c <= 13; c++) {
      const cell = ws.getCell(r, c);
      cell.fill = headerFill;
      cell.font = headerFont;
      cell.alignment = centerAlign;
      cell.border = thinBorder;
    }
  }

  // === DATA ===
  const rows = [
    { row: 4, a: '시점부\nTBM\n(상선)', b: '터널 내부 안전관리 시스템\n(비상인터폰/유해가스/비상WiFi/\n비상방송/위치관제)', q: 1, gs: 15, ge: 17 },
    { row: 5, b: '출입통제 시스템', q: 1, gs: 13, ge: 17 },
    { row: 6, b: '옥외 DID 시스템', q: 1, gs: 14, ge: 16 },
    { row: 7, a: '주출입구', b: '터널 내부 안전관리 시스템\n(비상인터폰/유해가스/비상WiFi/\n비상방송/위치관제)', q: 1, gs: 17, ge: 18 },
    { row: 8, b: '출입통제 시스템', q: 1, gs: 13, ge: 16 },
    { row: 9, b: '옥외 DID 시스템', q: 1, gs: 14, ge: 16 },
    { row: 10, b: '터널내부 CCTV', q: 1, gs: 16, ge: 17 },
    { row: 11, a: '#3환기구', b: '크람셸 하역알림 시스템', q: 1, gs: 9, ge: 10 },
    { row: 12, a: '#5환기구', b: '출입통제 시스템', q: 1, gs: 13, ge: 15 },
    { row: 13, b: '옥외 DID 시스템', q: 1, gs: 14, ge: 15 },
    { row: 14, a: '부출입구', b: '크람셸 하역알림 시스템', q: 1, gs: 16, ge: 17 },
    { row: 15, a: '#6환기구', b: '출입통제 시스템', q: 1, gs: 14, ge: 16 },
    { row: 16, b: '옥외 DID 시스템', q: 1, gs: 15, ge: 16 },
    { row: 17, a: '공통', b: '네트워크 인프라 인입', q: 1 },
  ];

  ws.mergeCells('A4:A6');
  ws.mergeCells('A7:A10');
  ws.mergeCells('A12:A13');
  ws.mergeCells('A15:A16');

  [4, 7].forEach(r => { ws.getRow(r).height = 50; });
  [5, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17].forEach(r => { ws.getRow(r).height = 28; });

  rows.forEach(d => {
    if (d.a) {
      const cellA = ws.getCell(d.row, 1);
      cellA.value = d.a;
      cellA.font = boldFont;
      cellA.alignment = centerAlign;
    }
    const cellB = ws.getCell(d.row, 2);
    cellB.value = d.b;
    cellB.font = normalFont;
    cellB.alignment = centerAlign;

    const cellC = ws.getCell(d.row, 3);
    cellC.value = d.q;
    cellC.font = normalFont;
    cellC.alignment = centerAlign;

    if (d.gs && d.ge) {
      for (let day = d.gs; day <= d.ge; day++) {
        const col = 4 + (day - 9);
        ws.getCell(d.row, col).fill = redFill;
      }
    }
  });

  for (let r = 4; r <= 17; r++) {
    for (let c = 1; c <= 13; c++) {
      const cell = ws.getCell(r, c);
      cell.border = thinBorder;
      if (!cell.alignment) cell.alignment = centerAlign;
    }
  }

  await wb.xlsx.writeFile('시공일정표.xlsx');
  console.log('완료: 시공일정표.xlsx');
}

create().catch(console.error);
