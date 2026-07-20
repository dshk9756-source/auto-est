const express = require('express');
const session = require('express-session');
const path = require('path');

const app = express();
const PORT = process.env.PORT ? Number(process.env.PORT) : 5173;

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

app.use(express.json({ limit: '10mb' }));

// 세션 설정
app.use(session({
  secret: 'openworks-estimate-secret-key-2026',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 }  // 8시간
}));

// 인증 API (로그인/로그아웃은 누구나 접근, 계정관리는 라우터 내부에서 권한 체크)
app.use('/api/auth', require('./routes/auth'));

// 로그인 페이지는 인증 없이 접근 허용
app.get('/login.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'login.html'));
});


// 로그인 여부 체크 미들웨어 (login.html, 정적 리소스 제외)
function requireLogin(req, res, next) {
  // CSS, JS, 폰트, 이미지 등 정적 리소스는 통과
  if (/\.(css|js|woff2?|ttf|eot|svg|png|jpg|ico)$/i.test(req.path)) return next();
  // 이미 로그인된 상태
  if (req.session && req.session.user) return next();
  // API 요청이면 401
  if (req.path.startsWith('/api/')) return res.status(401).json({ error: '로그인 필요' });
  // 그 외 페이지 요청이면 로그인 페이지로 리다이렉트
  return res.redirect('/login.html');
}

// 관리자 전용 체크 미들웨어
function requireAdmin(req, res, next) {
  if (!req.session || !req.session.user) {
    if (req.path.startsWith('/api/')) return res.status(401).json({ error: '로그인 필요' });
    return res.redirect('/login.html');
  }
  if (req.session.user.role !== 'admin') {
    if (req.path.startsWith('/api/')) return res.status(403).json({ error: '관리자 권한이 필요합니다.' });
    return res.status(403).send('<h3>관리자 권한이 필요합니다.</h3><a href="/">돌아가기</a>');
  }
  next();
}

// 관리자 전용: admin 페이지 & 계정 관리 페이지 & 카탈로그 수정 API
app.get('/admin.html', requireAdmin, (req, res) => {
  res.sendFile(path.join(__dirname, 'admin.html'));
});
app.get('/accounts.html', requireAdmin, (req, res) => {
  res.sendFile(path.join(__dirname, 'accounts.html'));
});
// 계정 관리 API 권한 체크는 routes/auth.js 내부에서 처리
app.post('/api/catalog/items', requireAdmin);
app.put('/api/catalog/items/:id', requireAdmin);
app.delete('/api/catalog/items/:id', requireAdmin);
app.post('/api/catalog/sets', requireAdmin);
app.put('/api/catalog/sets/:id', requireAdmin);
app.delete('/api/catalog/sets/:id', requireAdmin);
app.post('/api/catalog/smart-groups', requireAdmin);
app.put('/api/catalog/smart-groups/:id', requireAdmin);
app.delete('/api/catalog/smart-groups/:id', requireAdmin);
app.post('/api/catalog/purchase-map', requireAdmin);
app.put('/api/catalog/purchase-map/:id', requireAdmin);
app.delete('/api/catalog/purchase-map/:id', requireAdmin);
app.post('/api/catalog/companies', requireAdmin);
app.put('/api/catalog/companies/:id', requireAdmin);
app.delete('/api/catalog/companies/:id', requireAdmin);
app.put('/api/catalog/notes', requireAdmin);
app.post('/api/catalog/backup', requireAdmin);

// 나머지 모든 요청은 로그인 필요
app.use(requireLogin);

app.use('/api', require('./routes/catalog'));
app.use('/api', require('./routes/quote'));
app.use('/api', require('./routes/templates'));

app.use(express.static(__dirname));

app.listen(PORT, () => {
  console.log(`Server running: http://localhost:${PORT}`);
});
