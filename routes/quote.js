const express = require('express');
const router = express.Router();
const { generateQuote, generateSmartQuote, extractPreviewData, generatePurchaseRequest } = require('../generate_quote');

router.post('/generate-excel', async (req, res) => {
  try {
    const data = req.body;
    const isSmartTpl = data.template && data.template.includes('스마트');
    const buffer = await (isSmartTpl ? generateSmartQuote(data) : generateQuote(data));
    const cmpSuffix = (data.template && data.template.includes('비교')) ? '_비교견적' : '';
    const fname = encodeURIComponent(
      `${data.client || 'OPENWORKS'}_${data.date || ''}${cmpSuffix}.xlsx`
    );
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename*=UTF-8''${fname}`,
      'Content-Length': buffer.length,
    });
    res.end(buffer);
  } catch (e) {
    console.error('[generate-excel ERROR]', e.message, e.stack);
    res.status(500).json({ error: e.message });
  }
});

router.post('/preview', async (req, res) => {
  try {
    const data = req.body;
    const isSmartTpl = data.template && data.template.includes('스마트');
    const buffer = await (isSmartTpl ? generateSmartQuote(data) : generateQuote(data));
    const preview = await extractPreviewData(buffer);
    res.json(preview);
  } catch (e) {
    console.error('[preview ERROR]', e.message, e.stack);
    res.status(500).json({ error: e.message });
  }
});

router.post('/generate-purchase-request', async (req, res) => {
  try {
    const data = req.body;
    const buffer = await generatePurchaseRequest(data);
    const fname = encodeURIComponent(
      `구매요청서_${data.client || 'OPENWORKS'}_${data.date || ''}.xlsx`
    );
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename*=UTF-8''${fname}`,
      'Content-Length': buffer.length,
    });
    res.end(buffer);
  } catch (e) {
    console.error('[purchase-request ERROR]', e.message, e.stack);
    res.status(500).json({ error: e.message });
  }
});

module.exports = router;
