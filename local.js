/* 서버 없이 브라우저에서만 동작하는 로컬 저장/외부 API 레이어.
   File System Access API로 사용자가 고른 폴더에 db.json을 읽고 씁니다.
   IIFE로 감싸서 내부 함수/상수 이름이 다른 <script>(app.js 등)와 전역에서 충돌하지 않게 함. */
(function () {

/* ===================== 기본 데이터 모양 ===================== */
const DEFAULT_DB = {
  students: [],
  attendance: {},
  announcements: [],
  assignments: [],
  ddays: [],
  polls: [],
  newsletters: [],
  quickLinks: [],
  weeklyPlans: {},
  studentRecords: {},
  settings: {
    darkMode: false,
    fontSize: 'medium',
    calendarId1: '',
    calendarId2: '',
    calendarId3: '',
    thermometerGoal: 1500,
    thermometerReward: '',
    dashboardLayouts: null,
    adminPassword: '',
    geminiApiKey: '',
  },
};

/* ===================== IndexedDB: 폴더 핸들 저장 ===================== */
const IDB_NAME = 'boongssam-local';
const IDB_STORE = 'handles';
const IDB_KEY = 'dataDir';

function openHandleDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(IDB_STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveHandle(handle) {
  const db = await openHandleDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, 'readwrite');
    tx.objectStore(IDB_STORE).put(handle, IDB_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function loadSavedHandle() {
  const db = await openHandleDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, 'readonly');
    const req = tx.objectStore(IDB_STORE).get(IDB_KEY);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

/* ===================== 폴더 권한 ===================== */
async function ensurePermission(handle, requestIfNeeded) {
  const opts = { mode: 'readwrite' };
  if ((await handle.queryPermission(opts)) === 'granted') return true;
  if (!requestIfNeeded) return false;
  return (await handle.requestPermission(opts)) === 'granted';
}

/* 폴더를 처음 고를 때(또는 다시 고를 때) 호출 */
async function pickFolder() {
  const handle = await window.showDirectoryPicker();
  await ensurePermission(handle, true);
  try { await navigator.storage.persist(); } catch (e) { /* 지원 안 해도 무방 */ }
  await saveHandle(handle);
  return handle;
}

/* 저장해둔 폴더가 있으면 반환, 권한이 끊겼으면 재요청(사용자 클릭 필요) */
async function getActiveFolder() {
  const handle = await loadSavedHandle();
  if (!handle) return null;
  const ok = await ensurePermission(handle, false);
  if (ok) return handle;
  return { needsReconnect: true, handle };
}

async function reconnectFolder(handle) {
  const ok = await ensurePermission(handle, true);
  return ok ? handle : null;
}

/* ===================== db.json 읽기/쓰기 ===================== */
async function readDb(folderHandle) {
  try {
    const fileHandle = await folderHandle.getFileHandle('db.json', { create: false });
    const file = await fileHandle.getFile();
    const text = await file.text();
    return { ...DEFAULT_DB, ...JSON.parse(text) };
  } catch (e) {
    // 파일이 아직 없으면 기본값으로 새로 만들어줌
    await writeDb(folderHandle, DEFAULT_DB);
    return { ...DEFAULT_DB };
  }
}

async function writeDb(folderHandle, data) {
  const fileHandle = await folderHandle.getFileHandle('db.json', { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(JSON.stringify(data, null, 2));
  await writable.close();
}

window.LocalStore = { DEFAULT_DB, pickFolder, getActiveFolder, reconnectFolder, readDb, writeDb };

/* ===================== 데모/테스트 모드: 주소 끝에 ?mock=1 을 붙이면 실제 폴더 없이 메모리상의 가짜 데이터로 둘러볼 수 있음 ===================== */
if (location.search.includes('mock=1')) {
  let MOCK_DATA = JSON.parse(JSON.stringify(DEFAULT_DB));
  MOCK_DATA.students = [
    { id: 1, number: 1, name: '강민준', gender: 'M', birthday: '2014-03-15', allergies: '', healthNotes: '', parentPhone: '010-1234-5678', points: 45, group: 1 },
    { id: 2, number: 2, name: '고서연', gender: 'F', birthday: '2014-07-22', allergies: '땅콩', healthNotes: '', parentPhone: '010-2345-6789', points: 52, group: 1 },
  ];
  MOCK_DATA.ddays = [{ id: 'd1', name: '현장체험학습', date: '2026-12-20', emoji: '🚌' }];
  MOCK_DATA.announcements = [{ id: 'a1', title: '테스트 공지', content: '목(mock) 모드 테스트용 공지입니다.', type: 'general', date: '2026-11-13' }];
  MOCK_DATA.settings.thermometerGoal = 100;

  const fakeHandle = {
    queryPermission: async () => 'granted',
    requestPermission: async () => 'granted',
    getFileHandle: async () => ({
      getFile: async () => ({ text: async () => JSON.stringify(MOCK_DATA) }),
      createWritable: async () => ({
        write: async (text) => { MOCK_DATA = JSON.parse(text); },
        close: async () => {},
      }),
    }),
  };

  window.LocalStore.pickFolder = async () => fakeHandle;
  window.LocalStore.getActiveFolder = async () => fakeHandle;
  window.LocalStore.reconnectFolder = async () => fakeHandle;
  console.warn('[mock 모드] 실제 폴더 대신 메모리 안의 가짜 데이터를 사용합니다. 새로고침하면 초기화됩니다.');
}

/* ===================== 날짜/기타 헬퍼 (utils/helpers.js 포팅) ===================== */
const DAYS_KR = ['일', '월', '화', '수', '목', '금', '토'];

function getDayName(date = new Date()) { return DAYS_KR[date.getDay()]; }

function formatDate(date, format = 'full') {
  const d = typeof date === 'string' ? new Date(date) : date;
  const y = d.getFullYear(), m = d.getMonth() + 1, day = d.getDate(), dow = DAYS_KR[d.getDay()];
  switch (format) {
    case 'full': return `${y}년 ${m}월 ${day}일 (${dow})`;
    case 'short': return `${m}/${day} (${dow})`;
    case 'month': return `${y}년 ${m}월`;
    case 'iso': return `${y}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    default: return `${m}월 ${day}일`;
  }
}

function getDday(targetDate) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const target = new Date(targetDate); target.setHours(0, 0, 0, 0);
  const diff = Math.ceil((target - today) / (1000 * 60 * 60 * 24));
  if (diff === 0) return 'D-Day';
  if (diff > 0) return `D-${diff}`;
  return `D+${Math.abs(diff)}`;
}

function isTodayBirthday(birthday) {
  if (!birthday) return false;
  const today = new Date(); const bd = new Date(birthday);
  return today.getMonth() === bd.getMonth() && today.getDate() === bd.getDate();
}

function getAvatarColor(gender) { return gender === 'M' ? 'male' : 'female'; }
function getInitials(name) { return name.slice(0, 1); }

function getCurrentPeriod() {
  const now = new Date();
  const time = now.getHours() * 60 + now.getMinutes();
  const periods = [
    { p: 1, start: 540, end: 580 }, { p: 2, start: 590, end: 630 },
    { p: 3, start: 650, end: 690 }, { p: 4, start: 750, end: 790 },
    { p: 5, start: 790, end: 830 }, { p: 6, start: 840, end: 880 },
  ];
  const found = periods.find(p => time >= p.start && time <= p.end);
  return found ? found.p : null;
}

function shuffleArray(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function generateUUID() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') return crypto.randomUUID();
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
    const r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

function getWeekStart(date = new Date()) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  d.setDate(diff); d.setHours(0, 0, 0, 0);
  return d;
}

function toISODate(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

window.Helpers = {
  getDayName, formatDate, getDday, isTodayBirthday, getAvatarColor, getInitials,
  getCurrentPeriod, shuffleArray, generateUUID, getWeekStart, toISODate,
};

/* ===================== NEIS API (급식 / 학사일정) ===================== */
const ATPT_OFCDC_SC_CODE = 'F10';
const SD_SCHUL_CODE = '7401158';
const MEAL_URL = 'https://open.neis.go.kr/hub/mealServiceDietInfo';
const SCHEDULE_URL = 'https://open.neis.go.kr/hub/SchoolSchedule';

function neisTodayString() {
  const t = new Date();
  return `${t.getFullYear()}${String(t.getMonth() + 1).padStart(2, '0')}${String(t.getDate()).padStart(2, '0')}`;
}

async function fetchTodayLunch(date = neisTodayString()) {
  try {
    const url = `${MEAL_URL}?Type=json&pIndex=1&pSize=1&ATPT_OFCDC_SC_CODE=${ATPT_OFCDC_SC_CODE}&SD_SCHUL_CODE=${SD_SCHUL_CODE}&MLSV_YMD=${date}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error('네트워크 응답이 정상이 아닙니다.');
    const data = await res.json();
    if (!data.mealServiceDietInfo) return null;
    const row = data.mealServiceDietInfo[1].row[0];
    return {
      date: row.MLSV_YMD, type: row.MMEAL_SC_NM,
      menu: row.DDISH_NM.replace(/<br\/>/g, '\n').replace(/\([0-9.,]+\)/g, '').trim(),
      calories: row.CAL_INFO,
    };
  } catch (e) { console.error('급식 정보를 가져오는 중 오류 발생:', e); return null; }
}

async function fetchSchoolHoliday(date = neisTodayString()) {
  try {
    const url = `${SCHEDULE_URL}?Type=json&pIndex=1&pSize=10&ATPT_OFCDC_SC_CODE=${ATPT_OFCDC_SC_CODE}&SD_SCHUL_CODE=${SD_SCHUL_CODE}&AA_YMD=${date}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error('네트워크 응답이 정상이 아닙니다.');
    const data = await res.json();
    if (data.SchoolSchedule) {
      const row = data.SchoolSchedule[1].row.find(r => r.SBTR_DD_SC_NM === '휴업일');
      if (row) return { isHoliday: true, name: row.EVENT_NM || '휴업일' };
    }
    return { isHoliday: false, name: '' };
  } catch (e) { console.error('학사일정 정보를 가져오는 중 오류 발생:', e); return null; }
}

window.NeisApi = { fetchTodayLunch, fetchSchoolHoliday };

/* ===================== Gemini (브라우저에서 직접 호출) ===================== */
async function generateBehaviorSummary(studentData, apiKey) {
  if (!apiKey) throw new Error('Gemini API 키가 설정되지 않았습니다. 설정에서 입력해주세요.');

  const recordsText = (studentData.records || []).map(r => `- [${r.date}] [${r.category}] ${r.content}`).join('\n') || '기록 없음';
  const evaluationsText = (studentData.evaluations || []).map(e => `- [${e.title}] ${e.grade || '제출'}`).join('\n') || '기록 없음';

  const prompt = `
당신은 초등학교 교사입니다. 다음 학생의 기록을 바탕으로 생활기록부용 행동발달 종합의견을 작성해주세요.

[학생 정보]
이름: ${studentData.name}
성별: ${studentData.gender === 'M' ? '남' : '여'}
포인트: ${studentData.points || 0}점

[관찰/상담 기록]
${recordsText}

[과제/수행평가 결과]
${evaluationsText}

[요청 사항]
1. 이 학생의 학교생활, 교우관계, 학업태도를 종합적으로 요약해주세요.
2. 장점을 부각하되, 발전할 점도 부드럽게 포함해주세요.
3. 생활기록부에 들어갈 만한 정중한 교사 어투(~함, ~임 형태)로 300자 내외로 작성해 주세요.
`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=${apiKey}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message || 'AI 종합의견 생성 중 오류가 발생했습니다.');
  }
  const data = await res.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error('AI 응답을 읽을 수 없습니다.');
  return text;
}

window.GeminiApi = { generateBehaviorSummary };

/* ===================== 엑셀(학생 명단) ===================== */
function downloadStudentTemplate() {
  const templateData = [{ 이름: '홍길동', 성별: 'M', 생일: '2014-01-01', 연락처: '010-0000-0000' }];
  const ws = XLSX.utils.json_to_sheet(templateData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '학생명단');
  XLSX.writeFile(wb, '학생명단양식.xlsx');
}

function parseStudentExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        resolve(XLSX.utils.sheet_to_json(ws));
      } catch (e) { reject(e); }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

window.ExcelApi = { downloadStudentTemplate, parseStudentExcel };

})();
