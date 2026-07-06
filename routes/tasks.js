const express = require('express');
const router = express.Router();
const fs = require('fs');
const path = require('path');

const TASKS_DIR = path.join(__dirname, '..', 'data', 'tasks');

if (!fs.existsSync(TASKS_DIR)) fs.mkdirSync(TASKS_DIR, { recursive: true });

router.get('/tasks', (req, res) => {
  const files = fs.readdirSync(TASKS_DIR)
    .filter(f => f.endsWith('.json'))
    .sort()
    .reverse();

  const allTasks = [];
  for (const file of files) {
    try {
      const raw = fs.readFileSync(path.join(TASKS_DIR, file), 'utf-8');
      allTasks.push(JSON.parse(raw));
    } catch (e) { /* skip broken files */ }
  }
  res.json(allTasks);
});

router.get('/tasks/:date', (req, res) => {
  const file = path.join(TASKS_DIR, `${req.params.date}.json`);
  if (!fs.existsSync(file)) return res.status(404).json({ error: '해당 날짜 데이터 없음' });
  try {
    res.json(JSON.parse(fs.readFileSync(file, 'utf-8')));
  } catch (e) {
    res.status(500).json({ error: '파일 읽기 오류' });
  }
});

router.post('/tasks', (req, res) => {
  const data = req.body;
  if (!data.date) return res.status(400).json({ error: 'date 필드 필요' });
  const file = path.join(TASKS_DIR, `${data.date}.json`);
  fs.writeFileSync(file, JSON.stringify(data, null, 2), 'utf-8');
  res.json({ ok: true, date: data.date });
});

router.delete('/tasks/:date', (req, res) => {
  const file = path.join(TASKS_DIR, `${req.params.date}.json`);
  if (fs.existsSync(file)) fs.unlinkSync(file);
  res.json({ ok: true });
});

module.exports = router;
