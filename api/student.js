// 학생 화면용 공개 스냅샷 (Vercel 서버리스 함수)
// GET  /api/student  → 선생님이 마지막으로 올린 내용 (누구나 읽기)
// PUT  /api/student  → 선생님 앱이 내용 올리기 (X-Publish-Key 가 STUDENT_PUBLISH_KEY 와 같아야 함)
// 필요한 환경 변수: BLOB_READ_WRITE_TOKEN (Vercel Blob 저장소 연결 시 자동), STUDENT_PUBLISH_KEY (직접 설정)
import { put, head } from '@vercel/blob';
import { timingSafeEqual } from 'node:crypto';

const PATH = 'student/snapshot.json';
const MAX_BYTES = 500_000;

function cors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, PUT, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Publish-Key');
}

function sameKey(a, b) {
  const x = Buffer.from(String(a || '')), y = Buffer.from(String(b || ''));
  return x.length === y.length && timingSafeEqual(x, y);
}

export default async function handler(req, res) {
  cors(res);
  if (req.method === 'OPTIONS') return res.status(204).end();

  if (req.method === 'GET') {
    res.setHeader('Cache-Control', 'no-store');
    try {
      const meta = await head(PATH);
      const r = await fetch(meta.url);
      if (!r.ok) throw new Error(`blob ${r.status}`);
      const data = await r.json();
      // 학생 여러 명이 1분마다 확인해도 Blob 조회는 3분에 한 번꼴로만 일어나게 Vercel CDN에 잠깐 보관 (Hobby 무료 한도 보호)
      res.setHeader('Cache-Control', 'public, s-maxage=180, stale-while-revalidate=600');
      return res.status(200).json(data);
    } catch (e) {
      if (e && (e.name === 'BlobNotFoundError' || /not.?found/i.test(e.message || ''))) return res.status(404).json({ error: 'not_published' });
      console.error('student snapshot read failed:', e);
      return res.status(500).json({ error: 'read_failed' });
    }
  }

  if (req.method === 'PUT') {
    const serverKey = process.env.STUDENT_PUBLISH_KEY;
    if (!serverKey) return res.status(500).json({ error: 'no_server_key' });
    if (!sameKey(req.headers['x-publish-key'], serverKey)) return res.status(401).json({ error: 'bad_key' });
    let body = req.body;
    try { if (typeof body === 'string') body = JSON.parse(body); } catch { return res.status(400).json({ error: 'bad_json' }); }
    if (!body || body.v !== 1) return res.status(400).json({ error: 'bad_body' });
    const json = JSON.stringify(body);
    if (json.length > MAX_BYTES) return res.status(413).json({ error: 'too_large' });
    try {
      await put(PATH, json, { access: 'public', addRandomSuffix: false, allowOverwrite: true, contentType: 'application/json', cacheControlMaxAge: 60 });
      return res.status(200).json({ ok: true, updatedAt: body.updatedAt });
    } catch (e) {
      console.error('student snapshot write failed:', e);
      return res.status(500).json({ error: 'write_failed' });
    }
  }

  res.setHeader('Allow', 'GET, PUT, OPTIONS');
  return res.status(405).json({ error: 'method_not_allowed' });
}
