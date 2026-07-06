const express = require('express');
const bcrypt = require('bcryptjs');
const fs = require('fs');
const path = require('path');
const router = express.Router();

const USERS_PATH = path.join(__dirname, '..', 'data', 'users.json');

function readUsers() {
  return JSON.parse(fs.readFileSync(USERS_PATH, 'utf8'));
}
function writeUsers(data) {
  fs.writeFileSync(USERS_PATH, JSON.stringify(data, null, 2), 'utf8');
}

// 회원가입 (누구나 접근 가능)
router.post('/register', (req, res) => {
  const { id, pw, name, team } = req.body;
  if (!id || !pw) return res.status(400).json({ error: '아이디와 비밀번호를 입력하세요.' });
  if (id.length < 2) return res.status(400).json({ error: '아이디는 2자 이상이어야 합니다.' });
  if (pw.length < 4) return res.status(400).json({ error: '비밀번호는 4자 이상이어야 합니다.' });
  if (!name || !name.trim()) return res.status(400).json({ error: '사용자 이름을 입력하세요.' });

  const users = readUsers();
  if (users.find(u => u.id === id)) return res.status(409).json({ error: '이미 존재하는 아이디입니다.' });

  users.push({
    id,
    pw: bcrypt.hashSync(pw, 10),
    role: 'user',
    name: name.trim(),
    team: (team || '').trim(),
  });
  writeUsers(users);
  res.status(201).json({ ok: true, id });
});

// 아이디 찾기 (이름으로 검색, 누구나 접근 가능)
router.post('/find-id', (req, res) => {
  const { name } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: '사용자 이름을 입력하세요.' });

  const users = readUsers();
  const found = users.filter(u => u.name === name.trim());
  if (!found.length) return res.status(404).json({ error: '해당 이름으로 등록된 계정이 없습니다.' });

  const results = found.map(u => ({ id: u.id, name: u.name, team: u.team || '' }));
  res.json(results);
});

// 아이디 중복 확인 (누구나 접근 가능)
router.get('/check-id', (req, res) => {
  const id = (req.query.id || '').trim();
  if (!id) return res.status(400).json({ error: '아이디를 입력하세요.' });
  const users = readUsers();
  const exists = users.some(u => u.id === id);
  res.json({ exists });
});

// 로그인
router.post('/login', (req, res) => {
  const { id, pw } = req.body;
  if (!id || !pw) return res.status(400).json({ error: '아이디와 비밀번호를 입력하세요.' });

  const users = readUsers();
  const user = users.find(u => u.id === id);
  if (!user || !bcrypt.compareSync(pw, user.pw)) {
    return res.status(401).json({ error: '아이디 또는 비밀번호가 일치하지 않습니다.' });
  }

  req.session.user = { id: user.id, role: user.role, name: user.name, team: user.team || '' };
  res.json({ id: user.id, role: user.role, name: user.name, team: user.team || '' });
});

// 로그아웃
router.post('/logout', (req, res) => {
  req.session.destroy(() => {
    res.json({ ok: true });
  });
});

// 현재 세션 확인
router.get('/me', (req, res) => {
  if (!req.session.user) return res.status(401).json({ error: '로그인 필요' });
  res.json(req.session.user);
});

// 비밀번호 변경
router.post('/change-pw', (req, res) => {
  if (!req.session.user) return res.status(401).json({ error: '로그인 필요' });
  const { currentPw, newPw } = req.body;
  if (!currentPw || !newPw) return res.status(400).json({ error: '현재/새 비밀번호를 입력하세요.' });
  if (newPw.length < 4) return res.status(400).json({ error: '비밀번호는 4자 이상이어야 합니다.' });

  const users = readUsers();
  const user = users.find(u => u.id === req.session.user.id);
  if (!user || !bcrypt.compareSync(currentPw, user.pw)) {
    return res.status(401).json({ error: '현재 비밀번호가 일치하지 않습니다.' });
  }

  user.pw = bcrypt.hashSync(newPw, 10);
  writeUsers(users);
  res.json({ ok: true });
});

// ── 계정 관리 (관리자 전용) ──

function requireAdmin(req, res, next) {
  if (!req.session || !req.session.user) return res.status(401).json({ error: '로그인 필요' });
  if (req.session.user.role !== 'admin') return res.status(403).json({ error: '관리자 권한이 필요합니다.' });
  next();
}

// 사용자 목록 조회
router.get('/users', requireAdmin, (req, res) => {
  const users = readUsers().map(u => ({ id: u.id, role: u.role, name: u.name, team: u.team || '' }));
  res.json(users);
});

// 사용자 추가
router.post('/users', requireAdmin, (req, res) => {
  const { id, pw, role, name, team } = req.body;
  if (!id || !pw) return res.status(400).json({ error: '아이디와 비밀번호를 입력하세요.' });
  if (pw.length < 4) return res.status(400).json({ error: '비밀번호는 4자 이상이어야 합니다.' });
  if (!['admin', 'user'].includes(role)) return res.status(400).json({ error: '역할은 admin 또는 user만 가능합니다.' });

  const users = readUsers();
  if (users.find(u => u.id === id)) return res.status(409).json({ error: '이미 존재하는 아이디입니다.' });

  users.push({
    id,
    pw: bcrypt.hashSync(pw, 10),
    role: role || 'user',
    name: name || id,
    team: (team || '').trim(),
  });
  writeUsers(users);
  res.status(201).json({ id, role: role || 'user', name: name || id, team: (team || '').trim() });
});

// 사용자 수정 (이름, 역할, 비밀번호 재설정)
router.put('/users/:id', requireAdmin, (req, res) => {
  const users = readUsers();
  const user = users.find(u => u.id === req.params.id);
  if (!user) return res.status(404).json({ error: '사용자를 찾을 수 없습니다.' });

  if (req.body.name !== undefined) user.name = req.body.name;
  if (req.body.team !== undefined) user.team = req.body.team;
  if (req.body.role !== undefined) {
    if (!['admin', 'user'].includes(req.body.role)) return res.status(400).json({ error: '역할은 admin 또는 user만 가능합니다.' });
    user.role = req.body.role;
  }
  if (req.body.pw) {
    if (req.body.pw.length < 4) return res.status(400).json({ error: '비밀번호는 4자 이상이어야 합니다.' });
    user.pw = bcrypt.hashSync(req.body.pw, 10);
  }
  writeUsers(users);
  res.json({ id: user.id, role: user.role, name: user.name, team: user.team || '' });
});

// 사용자 삭제
router.delete('/users/:id', requireAdmin, (req, res) => {
  // 자기 자신은 삭제 불가
  if (req.session.user && req.session.user.id === req.params.id) {
    return res.status(400).json({ error: '현재 로그인된 계정은 삭제할 수 없습니다.' });
  }
  const users = readUsers();
  const before = users.length;
  const filtered = users.filter(u => u.id !== req.params.id);
  if (filtered.length === before) return res.status(404).json({ error: '사용자를 찾을 수 없습니다.' });
  writeUsers(filtered);
  res.json({ ok: true });
});

module.exports = router;
