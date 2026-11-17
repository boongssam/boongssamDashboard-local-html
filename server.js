// 교실 서버: 선생님 컴퓨터에서 켜 두면 같은 학교망의 학생 폰이 학생 화면(읽기 전용)을 볼 수 있음
// 실행:  node server.js                        (데이터 폴더 기본값: 바탕화면\data-dashboard)
//        node server.js "D:\내 데이터 폴더" 8080  (폴더와 포트를 직접 지정)
// 데이터는 인터넷에 올라가지 않음. 학생 화면에는 student-snapshot.js 가 고른 항목만 나감.
import http from 'node:http';
import fs from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const DATA_DIR = process.argv[2] || path.join(os.homedir(), 'Desktop', 'data-dashboard');
const PORT = Number(process.argv[3]) || 8080;
const DB_FILE = path.join(DATA_DIR, 'db.json');
// 학생이 폰으로 누른 '활동 완료'는 db.json 을 건드리지 않고 이 폴더에 한 건씩 파일로 남김.
// 선생님 앱이 몇 초마다 읽어 학습 관리에 반영한 뒤 지움 (선생님 앱이 db.json 을 덮어써도 체크가 사라지지 않게)
const CHECK_DIR = path.join(DATA_DIR, 'checks');
const safeId = (v) => String(v).replace(/[^A-Za-z0-9_-]/g, '_').slice(0, 80);

async function pendingChecks(activityId) {
  const names = await fs.readdir(CHECK_DIR).catch(() => []);
  const prefix = safeId(activityId) + '__';
  const out = [];
  for (const n of names) {
    if (!n.startsWith(prefix) || !n.endsWith('.json')) continue;
    try { out.push(JSON.parse(await fs.readFile(path.join(CHECK_DIR, n), 'utf8')).studentId); } catch {}
  }
  return out;
}

// 학생 폰 투표도 같은 방식: votes/ 에 한 표씩 파일로 남기고, 선생님 앱이 가져가 더한 뒤 지움
const VOTE_DIR = path.join(DATA_DIR, 'votes');
const votedBy = new Map(); // 투표id → 이미 낸 폰(IP) — 같은 폰에서 여러 번 못 내게 (서버를 다시 켜면 초기화)

async function pendingVotes(pollId, optionCount) {
  const counts = Array(optionCount).fill(0);
  const names = await fs.readdir(VOTE_DIR).catch(() => []);
  const prefix = safeId(pollId) + '__';
  for (const n of names) {
    if (!n.startsWith(prefix) || !n.endsWith('.json')) continue;
    try { const v = JSON.parse(await fs.readFile(path.join(VOTE_DIR, n), 'utf8')); if (Number.isInteger(v.option) && v.option >= 0 && v.option < optionCount) counts[v.option]++; } catch {}
  }
  return counts;
}

function readBody(req, limit = 2048) {
  return new Promise((resolve, reject) => {
    let s = '';
    req.on('data', (c) => { s += c; if (s.length > limit) { reject(new Error('too_large')); req.destroy(); } });
    req.on('end', () => resolve(s));
    req.on('error', reject);
  });
}

// 학생에게 내보내는 규칙은 브라우저와 같은 파일을 그대로 씀
const sandbox = { window: {} };
vm.runInNewContext(readFileSync(path.join(HERE, 'student-snapshot.js'), 'utf8'), sandbox);
const { build } = sandbox.window.StudentSnapshot;

// 구글 달력 일정: 선생님 설정의 Google API 키·달력으로 서버가 읽어 학생 '다가오는 날'에 섞음
// (선생님 대시보드와 같은 함수·같은 20일 기준. 학생 생일은 넣지 않음. 10분마다 한 번만 새로 읽음)
const lib = { window: {}, location: { search: '', protocol: 'http:' }, fetch, AbortSignal, URLSearchParams, console, setTimeout, clearTimeout };
vm.runInNewContext(readFileSync(path.join(HERE, 'local.js'), 'utf8'), lib);
const { fetchCalendarEvents } = lib.window.GoogleCalApi;
const { upcomingAutoDays } = lib.window.Helpers;
const CAL_DAYS = 20, CAL_TTL = 10 * 60 * 1000;
let calCache = { key: '', at: 0, days: [] };
async function calendarDays(settings = {}) {
  const ids = [settings.calendarId1, settings.calendarId2, settings.calendarId3].filter(Boolean);
  const key = (settings.googleApiKey || '').trim();
  if (!ids.length || !key) return [];
  const cacheKey = ids.join('|') + '|' + key;
  if (calCache.key === cacheKey && Date.now() - calCache.at < CAL_TTL) return calCache.days;
  try {
    const from = new Date(); from.setHours(0, 0, 0, 0);
    const to = new Date(from); to.setDate(to.getDate() + CAL_DAYS + 1);
    const evs = await Promise.race([fetchCalendarEvents(ids, key, from, to), new Promise((_, rej) => setTimeout(() => rej(new Error('시간 초과')), 8000))]);
    const days = upcomingAutoDays([], evs, new Date(), CAL_DAYS).map(d => ({ name: d.name, date: d.date, emoji: d.emoji, auto: '달력' }));
    calCache = { key: cacheKey, at: Date.now(), days };
  } catch (e) {
    console.error('구글 달력을 읽지 못했어요:', e.message);
    calCache = { key: cacheKey, at: Date.now(), days: calCache.key === cacheKey ? calCache.days : [] }; // 실패해도 10분은 다시 안 부름
  }
  return calCache.days;
}

// 이 목록에 있는 파일만 보냄 (server.js, api/, .git, package.json 등은 절대 안 보냄)
const FILES = {
  '/index.html': 'text/html; charset=utf-8',
  '/app.js': 'text/javascript; charset=utf-8',
  '/local.js': 'text/javascript; charset=utf-8',
  '/student-snapshot.js': 'text/javascript; charset=utf-8',
  '/styles.css': 'text/css; charset=utf-8',
  '/theme.css': 'text/css; charset=utf-8',
};

// 학생 Wi-Fi에서 외부 CDN(unpkg)이 막혀도 열리도록 라이브러리를 이 서버가 직접 보냄 (npm install 로 받은 같은 버전)
const VENDOR = {
  '/vendor/react.js': ['https://unpkg.com/react@18.3.1/umd/react.production.min.js', 'node_modules/react/umd/react.production.min.js'],
  '/vendor/react-dom.js': ['https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js', 'node_modules/react-dom/umd/react-dom.production.min.js'],
  '/vendor/xlsx.js': ['https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js', 'node_modules/xlsx/dist/xlsx.full.min.js'],
};
const hasVendor = Object.values(VENDOR).every(([, f]) => { try { readFileSync(path.join(HERE, f)); return true; } catch { return false; } });

const isLocal = (req) => ['127.0.0.1', '::1', '::ffff:127.0.0.1'].includes(req.socket.remoteAddress);

function send(res, code, body, type = 'text/plain; charset=utf-8', extra = {}) {
  res.writeHead(code, { 'Content-Type': type, 'X-Content-Type-Options': 'nosniff', ...extra });
  res.end(body);
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, 'http://x');
    const isCheck = req.method === 'POST' && url.pathname === '/api/student/check';
    const isVote = req.method === 'POST' && url.pathname === '/api/student/vote';
    if (req.method !== 'GET' && req.method !== 'HEAD' && !(req.method === 'PUT' && url.pathname === '/api/student') && !isCheck && !isVote) return send(res, 405, '허용되지 않는 요청');
    const json = (code, obj) => send(res, code, JSON.stringify(obj), 'application/json', { 'Cache-Control': 'no-store' });
    const readDb = async () => JSON.parse(await fs.readFile(DB_FILE, 'utf8'));

    // 학생 폰: "활동 끝냈어요" — 지금 활동 + 반 명단에 있는 학생만 받음
    if (isCheck) {
      let body;
      try { body = JSON.parse(await readBody(req)); } catch { return json(400, { error: 'bad_body' }); }
      let db;
      try { db = await readDb(); } catch { return json(503, { error: 'no_data' }); }
      // 지난 활동도 늦게 끝내면 체크할 수 있게, 학습 관리에 남아 있는 '활동'이면 받음
      const act = (db.assignments || []).find(a => a.subject === '활동' && String(a.id) === String(body?.activityId));
      const studentId = String(body?.studentId ?? '');
      if (!act) return json(409, { error: 'activity_changed' });
      if (!(db.students || []).some(s => String(s.id) === studentId)) return json(400, { error: 'unknown_student' });
      await fs.mkdir(CHECK_DIR, { recursive: true });
      await fs.writeFile(path.join(CHECK_DIR, `${safeId(act.id)}__${safeId(studentId)}.json`), JSON.stringify({ activityId: String(act.id), studentId, at: new Date().toISOString() }));
      return json(200, { ok: true });
    }

    // 학생 폰: 투표 한 표 — 진행 중인 투표, 있는 항목만, 같은 폰은 한 번만
    if (isVote) {
      let body;
      try { body = JSON.parse(await readBody(req)); } catch { return json(400, { error: 'bad_body' }); }
      let db;
      try { db = await readDb(); } catch { return json(503, { error: 'no_data' }); }
      const poll = (db.polls || []).find(p => String(p.id) === String(body?.pollId));
      if (!poll || poll.isActive === false) return json(409, { error: 'poll_closed' });
      const option = body?.option;
      if (!Number.isInteger(option) || option < 0 || option >= (poll.options || []).length) return json(400, { error: 'bad_option' });
      const who = req.socket.remoteAddress || '?';
      const voters = votedBy.get(String(poll.id)) || new Set();
      if (voters.has(who)) return json(409, { error: 'already_voted' });
      voters.add(who); votedBy.set(String(poll.id), voters);
      await fs.mkdir(VOTE_DIR, { recursive: true });
      await fs.writeFile(path.join(VOTE_DIR, `${safeId(poll.id)}__${Date.now()}_${Math.random().toString(36).slice(2, 8)}.json`), JSON.stringify({ pollId: String(poll.id), option, at: new Date().toISOString() }));
      return json(200, { ok: true });
    }

    if (url.pathname === '/api/student') {
      // 선생님 앱이 Vercel용으로 올리려 해도 여기서는 db.json 을 직접 읽으므로 받기만 하고 무시
      if (req.method === 'PUT') { req.resume(); return send(res, 200, '{"ok":true,"local":true}', 'application/json'); }
      let db;
      try { db = await readDb(); }
      catch { return json(404, { error: 'not_published' }); }
      const snap = { ...build(db, new Date(), { withActivity: true }), canCheck: true, updatedAt: (await fs.stat(DB_FILE)).mtime.toISOString() };
      // 선생님 앱이 아직 가져가지 않은 체크도 바로 보이게 합침
      // 구글 달력 일정을 직접 넣은 D-Day와 섞어 날짜순으로 (같은 날 같은 이름은 하나만)
      const seen = new Set(snap.ddays.map(d => d.name + '|' + d.date));
      snap.ddays = [...snap.ddays, ...(await calendarDays(db.settings)).filter(d => !seen.has(d.name + '|' + d.date))]
        .sort((a, b) => String(a.date).localeCompare(String(b.date)));
      for (const a of snap.activities || []) a.done = [...new Set([...a.done, ...(await pendingChecks(a.id))])];
      if (snap.activities?.length) snap.activity = snap.activities[0];
      for (const p of snap.polls || []) {
        const extra = await pendingVotes(p.id, p.options.length);
        p.votes = p.votes.map((v, i) => v + extra[i]);
        p.myVote = (votedBy.get(p.id) || new Set()).has(req.socket.remoteAddress || '?');
      }
      return json(200, snap);
    }

    // 다른 기기(학생 폰)에서 주소만 치고 들어오면 바로 학생 화면으로
    if (url.pathname === '/' && !isLocal(req) && !url.searchParams.has('student')) return send(res, 302, '', 'text/plain', { Location: '/?student' });

    if (VENDOR[url.pathname] && hasVendor) {
      const body = await fs.readFile(path.join(HERE, VENDOR[url.pathname][1]));
      return send(res, 200, req.method === 'HEAD' ? '' : body, 'text/javascript; charset=utf-8', { 'Cache-Control': 'public, max-age=86400' });
    }

    const file = url.pathname === '/' ? '/index.html' : url.pathname;
    if (!FILES[file]) return send(res, 404, '없는 페이지예요');
    let body = await fs.readFile(path.join(HERE, file));
    if (file === '/index.html' && hasVendor) {
      let html = body.toString('utf8');
      for (const [local, [cdn]] of Object.entries(VENDOR)) html = html.split(cdn).join(local);
      body = html;
    }
    return send(res, 200, req.method === 'HEAD' ? '' : body, FILES[file], { 'Cache-Control': 'no-cache' });
  } catch (e) {
    console.error(e);
    return send(res, 500, '서버 오류');
  }
});

server.on('error', (e) => {
  console.error(e.code === 'EADDRINUSE' ? `포트 ${PORT} 를 이미 다른 프로그램이 쓰고 있어요. 다른 번호로 실행해 보세요: node server.js "${DATA_DIR}" 8081` : e);
  process.exit(1);
});

server.listen(PORT, '0.0.0.0', async () => {
  const ok = await fs.access(DB_FILE).then(() => true, () => false);
  const ips = Object.values(os.networkInterfaces()).flat().filter(a => a && a.family === 'IPv4' && !a.internal).map(a => a.address);
  console.log('\n학급 대시보드 교실 서버가 켜졌어요.');
  console.log(`데이터: ${DB_FILE} ${ok ? '(찾음)' : '(없음! 폴더 경로를 확인하세요)'}`);
  if (!hasVendor) console.log('라이브러리: 인터넷(unpkg)에서 받음 — 학생 Wi-Fi에서 막히면 이 폴더에서 `npm install` 을 한 번 실행하세요');
  console.log(`\n선생님(이 컴퓨터):  http://localhost:${PORT}`);
  ips.forEach(ip => console.log(`학생 QR 주소:      http://${ip}:${PORT}/?student`));
  console.log('\n끄려면 Ctrl + C');
});
