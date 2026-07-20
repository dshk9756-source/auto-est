const { generateStatement } = require('../generate_statement');
const fs = require('fs');
const path = require('path');

(async () => {
  const buf = await generateStatement({
    date: '2026-07-14',
    client: '한신공영 인주-염치 2공구',
    company: '한신공영',
    siteName: '인주-염치 2공구',
    buyer: {
      bizNo: '114-81-04605',
      name: '한신공영(주)',
      ceo: '최문규, 전재식',
      address: '서울시 강남구 테헤란로 000',
      bizType: '건설',
      bizItem: '토목건축,산업설비 등',
    },
    items: [
      { name: 'CCTV 카메라', spec: '실외형', unit: 'EA', qty: 4, effectiveP: 250000, note: '' },
      { name: 'NVR 녹화장치', spec: '16채널', unit: 'EA', qty: 1, effectiveP: 900000, note: '' },
      { name: '설치공사비', spec: '', unit: '식', qty: 1, effectiveP: 500000, note: '' },
    ],
  });
  const out = path.join(__dirname, '..', 'test_output_거래명세서.xlsx');
  fs.writeFileSync(out, buf);
  console.log('생성 완료:', out);
})().catch(e => { console.error(e); process.exit(1); });
