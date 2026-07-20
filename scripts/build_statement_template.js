'use strict';
/**
 * 한신공영 거래명세서 파일을 기반으로 마커를 삽입해 template/거래명세서.xlsx 를 생성한다.
 * 서식(폰트·병합·테두리)은 절대 건드리지 않고 cell.value 만 교체한다.
 */
const ExcelJS = require('exceljs');
const path = require('path');

const SRC  = 'C:\\Users\\user\\Desktop\\한신공영 인주-염치 2공구 거래명세서_260713.xlsx';
const DEST = path.join(__dirname, '..', 'template', '거래명세서.xlsx');
const DESKTOP_COPY = 'C:\\Users\\user\\Desktop\\거래명세서_템플릿_확인용.xlsx';

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(SRC);
  const ws = wb.worksheets[0];

  // 작성일 / 문서번호
  ws.getCell('A2').value = '작성일자 : {{견적일자}}';
  ws.getCell('I2').value = 'No : {{문서번호}}';

  // 공급받는자(건설사) 정보 — 견적 대상 건설사에 따라 바뀌는 값
  ws.getCell('H3').value = '{{거래처_등록번호}}';
  ws.getCell('H4').value = '{{거래처_상호}}';
  ws.getCell('J4').value = '{{거래처_대표자}}';
  ws.getCell('H5').value = '{{거래처_주소}}';
  ws.getCell('H6').value = '{{거래처_업태}}';
  ws.getCell('J6').value = '{{거래처_종목}}';

  // 품목 영역 마커 (8~17행). 8행 = 아이템 스타일 템플릿, 17행 = 끝
  ws.getCell('J8').value  = '{{ITEM_START}}{{ITEM_ROW}}';
  ws.getCell('J17').value = '{{ITEM_END}}';

  // 8~17행에 남아있던 예시 데이터 정리 (서식은 그대로, 값만 비움)
  for (let r = 8; r <= 17; r++) {
    ['A','B','C','D','E','F','G','H','I'].forEach(col => {
      if (col === 'J' && (r === 8 || r === 17)) return; // 마커 유지
      const cell = ws.getCell(`${col}${r}`);
      cell.value = null;
    });
  }

  // 특기사항
  ws.getCell('B19').value = '{{특기사항}}';

  await wb.xlsx.writeFile(DEST);
  await wb.xlsx.writeFile(DESKTOP_COPY);
  console.log('저장 완료:', DEST);
  console.log('바탕화면 확인용 사본:', DESKTOP_COPY);
})().catch(e => { console.error(e); process.exit(1); });
