const express = require('express');
const fs = require('fs');
const path = require('path');
const router = express.Router();

const ROOT = path.join(__dirname, '..');
const EXCLUDE = ['구매요청서.xlsx'];

router.get('/templates', (req, res) => {
  const templateDir = path.join(ROOT, 'template');
  const files = fs.readdirSync(templateDir)
    .filter(f => f.endsWith('.xlsx') && !f.startsWith('~$') && !EXCLUDE.includes(f))
    .map(f => ({ filename: f, label: f.replace(/\.xlsx$/i, '') }));
  res.json(files);
});

module.exports = router;
