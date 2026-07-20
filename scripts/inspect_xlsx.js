const ExcelJS = require('exceljs');
const path = require('path');

const target = process.argv[2];

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(target);
  wb.worksheets.forEach(ws => {
    console.log('=== SHEET:', ws.name, 'dims:', ws.dimensions ? ws.dimensions.address : '', 'rowCount:', ws.rowCount, 'colCount:', ws.columnCount);
    console.log('--- merges ---');
    if (ws._merges) {
      Object.values(ws._merges).forEach(m => {
        console.log(JSON.stringify(m.model));
      });
    }
    console.log('--- column widths ---');
    ws.columns.forEach((c, i) => {
      if (c && c.width) console.log('col', i+1, 'width', c.width);
    });
    console.log('--- cells ---');
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      const vals = [];
      row.eachCell({ includeEmpty: true }, (cell, cn) => {
        let v = cell.value;
        if (v && typeof v === 'object') {
          if (v.richText) v = v.richText.map(r => r.text).join('');
          else if (v.formula) v = '=FORMULA:' + v.formula + (v.result != null ? ('|res:' + v.result) : '');
          else if (v.result != null) v = v.result;
          else v = JSON.stringify(v);
        }
        if (v !== null && v !== undefined && v !== '') vals.push(`${cell.address}=${JSON.stringify(v)}`);
      });
      if (vals.length) console.log(`R${rn}:`, vals.join(' | '));
    });
  });
})().catch(e => { console.error(e); process.exit(1); });
