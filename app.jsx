const { useState, useEffect, useRef, useCallback, useReducer, useContext, useMemo, createContext } = React;
const {
  getDayName, formatDate, getDday, isTodayBirthday, getAvatarColor, getInitials,
  getCurrentPeriod, shuffleArray, generateUUID, getWeekStart, toISODate, normalizeBirthday, upcomingAutoDays,
} = window.Helpers;
const { fetchTodayLunch, fetchSchoolHoliday } = window.NeisApi;
const { generateBehaviorSummary } = window.GeminiApi;
const { extractCalendarId, fetchCalendarEvents, placeEventsInWeek } = window.GoogleCalApi;
const { downloadStudentTemplate, parseStudentExcel } = window.ExcelApi;
const { pickFolder, getActiveFolder, reconnectFolder, readDb, writeDb } = window.LocalStore;

/* ===================================================================
   AppContext (서버 대신 로컬 폴더에 읽고 쓰는 버전)
   =================================================================== */
const AppContext = createContext(null);

const initialState = {
  students: [],
  attendance: {},
  selectedAttendanceDate: toISODate(new Date()),
  monthlyAttendance: {},
  announcements: [],
  assignments: [],
  newsletters: [],
  polls: [],
  toasts: [],
  weeklyPlans: {},
  quickLinks: [],
  isAdmin: false,
  settings: { ...window.LocalStore.DEFAULT_DB.settings },
  ddays: [],
  studentRecords: {},
};

function reducer(state, action) {
  switch (action.type) {
    case 'SET_STUDENTS': return { ...state, students: action.payload };
    case 'SET_ANNOUNCEMENTS': return { ...state, announcements: action.payload };
    case 'SET_ASSIGNMENTS': return { ...state, assignments: action.payload };
    case 'SET_ATTENDANCE_MAP': return { ...state, attendance: action.payload };
    case 'SET_SELECTED_ATTENDANCE_DATE': return { ...state, selectedAttendanceDate: action.payload };
    case 'SET_MONTHLY_ATTENDANCE': {
      const newMonthly = { ...state.monthlyAttendance };
      for (const [date, records] of Object.entries(action.payload)) {
        newMonthly[date] = { ...newMonthly[date], ...records };
      }
      return { ...state, monthlyAttendance: newMonthly };
    }
    case 'SAVE_ATTENDANCE_DAY': {
      const { date, records } = action.payload;
      return {
        ...state,
        monthlyAttendance: { ...state.monthlyAttendance, [date]: { ...(state.monthlyAttendance[date] || {}), ...records } },
        attendance: date === state.selectedAttendanceDate ? { ...state.attendance, ...records } : state.attendance,
      };
    }
    case 'UPDATE_MONTHLY_ATTENDANCE': {
      const { date, studentId, status, id } = action.payload;
      const dateRecords = state.monthlyAttendance[date] || {};
      return {
        ...state,
        monthlyAttendance: {
          ...state.monthlyAttendance,
          [date]: { ...dateRecords, [studentId]: { ...dateRecords[studentId], status, note: dateRecords[studentId]?.note || '', id: id || dateRecords[studentId]?.id } },
        },
      };
    }
    case 'SET_ATTENDANCE':
      return {
        ...state,
        attendance: {
          ...state.attendance,
          [action.payload.studentId]: { ...state.attendance[action.payload.studentId], status: action.payload.status, note: action.payload.note || '', id: action.payload.id || state.attendance[action.payload.studentId]?.id },
        },
      };
    case 'UPDATE_POINTS':
      return { ...state, students: state.students.map(s => s.id === action.payload.studentId ? { ...s, points: Math.max(0, s.points + action.payload.amount) } : s) };
    case 'RESET_POINTS':
      return { ...state, students: state.students.map(s => s.id === action.payload.studentId ? { ...s, points: 0 } : s) };
    case 'ADD_ANNOUNCEMENT': return { ...state, announcements: [action.payload, ...state.announcements] };
    case 'UPDATE_ANNOUNCEMENT': return { ...state, announcements: state.announcements.map(a => a.id === action.payload.id ? action.payload : a) };
    case 'DELETE_ANNOUNCEMENT': return { ...state, announcements: state.announcements.filter(a => a.id !== action.payload) };
    case 'ADD_ASSIGNMENT': return { ...state, assignments: [...state.assignments, action.payload] };
    case 'DELETE_ASSIGNMENT': return { ...state, assignments: state.assignments.filter(a => a.id !== action.payload) };
    case 'UPDATE_ASSIGNMENT_SUBMISSIONS':
      return { ...state, assignments: state.assignments.map(a => a.id === action.payload.id ? { ...a, submissions: { ...a.submissions, ...action.payload.submissions } } : a) };
    case 'EVALUATE_ASSIGNMENT':
      return {
        ...state,
        assignments: state.assignments.map(a => {
          if (a.id === action.payload.id) {
            const currentSub = a.submissions?.[action.payload.studentId] || {};
            return { ...a, submissions: { ...a.submissions, [action.payload.studentId]: { ...currentSub, grade: action.payload.grade } } };
          }
          return a;
        }),
      };
    case 'ADD_DDAY': return { ...state, ddays: [...state.ddays, action.payload] };
    case 'DELETE_DDAY': return { ...state, ddays: state.ddays.filter(d => d.id !== action.payload) };
    case 'SET_DDAYS': return { ...state, ddays: action.payload };
    case 'ADD_POLL': return { ...state, polls: [...state.polls, action.payload] };
    case 'DELETE_POLL': return { ...state, polls: state.polls.filter(p => p.id !== action.payload) };
    case 'SET_POLLS': return { ...state, polls: action.payload };
    case 'VOTE_POLL': return { ...state, polls: state.polls.map(p => p.id === action.payload.pollId ? { ...p, votes: action.payload.newVotes } : p) };
    case 'ADD_TOAST': return { ...state, toasts: [...state.toasts, action.payload] };
    case 'REMOVE_TOAST': return { ...state, toasts: state.toasts.filter(t => t.id !== action.payload) };
    case 'ADD_STUDENT': return { ...state, students: [...state.students, action.payload] };
    case 'UPDATE_STUDENT': return { ...state, students: state.students.map(s => s.id === action.payload.id ? { ...s, ...action.payload } : s) };
    case 'BULK_UPDATE_STUDENTS':
      return { ...state, students: state.students.map(s => { const u = action.payload.find(x => x.id === s.id); return u ? { ...s, ...u } : s; }) };
    case 'SET_NEWSLETTERS': return { ...state, newsletters: action.payload };
    case 'COLLECT_NEWSLETTER': return { ...state, newsletters: state.newsletters.map(n => n.id === action.payload.id ? { ...n, collected: action.payload.collected } : n) };
    case 'TOGGLE_ADMIN': return { ...state, isAdmin: action.payload };
    case 'UPDATE_SETTINGS': return { ...state, settings: { ...state.settings, ...action.payload } };
    case 'SET_WEEKLY_PLAN': {
      const { weekKey, day, period, subject, content } = action.payload;
      const existingWeek = state.weeklyPlans[weekKey] || {};
      return { ...state, weeklyPlans: { ...state.weeklyPlans, [weekKey]: { ...existingWeek, [day]: { ...(existingWeek[day] || {}), [period]: { subject, content } } } } };
    }
    case 'SET_WEEKLY_PLANS_BULK': return { ...state, weeklyPlans: action.payload };
    case 'SET_QUICK_LINKS': return { ...state, quickLinks: action.payload };
    case 'ADD_LINK': return { ...state, quickLinks: [...state.quickLinks, action.payload] };
    case 'UPDATE_LINK': return { ...state, quickLinks: state.quickLinks.map(l => l.id === action.payload.id ? action.payload : l) };
    case 'DELETE_LINK': return { ...state, quickLinks: state.quickLinks.filter(l => l.id !== action.payload) };
    case 'SET_STUDENT_RECORDS': return { ...state, studentRecords: action.payload };
    case 'ADD_STUDENT_RECORD': {
      const { studentId, record } = action.payload;
      const existing = state.studentRecords[studentId] || [];
      return { ...state, studentRecords: { ...state.studentRecords, [studentId]: [...existing, record] } };
    }
    case 'DELETE_STUDENT_RECORD': {
      const { studentId, recordId } = action.payload;
      const existing = state.studentRecords[studentId] || [];
      return { ...state, studentRecords: { ...state.studentRecords, [studentId]: existing.filter(r => r.id !== recordId) } };
    }
    default: return state;
  }
}

function AppProvider({ children }) {
  const [state, dispatch] = useReducer(reducer, initialState);
  const [isDbLoading, setIsDbLoading] = useState(true);
  const [folder, setFolder] = useState(null);
  const [folderStatus, setFolderStatus] = useState('loading'); // loading | none | needs-reconnect | connected
  const saveTimer = useRef(null);
  const writeQueue = useRef(Promise.resolve()); // db.json 쓰기를 한 줄로 세워서 동시에 쓰다 깨지지 않게
  const pendingSave = useRef(null); // 아직 안 쓴 저장 — 창을 닫거나 다른 창으로 가면 바로 씀
  useEffect(() => {
    const flush = () => { if (document.visibilityState === 'hidden' && pendingSave.current) { clearTimeout(saveTimer.current); pendingSave.current(); } };
    document.addEventListener('visibilitychange', flush);
    window.addEventListener('pagehide', flush);
    return () => { document.removeEventListener('visibilitychange', flush); window.removeEventListener('pagehide', flush); };
  }, []);

  const showToast = useCallback((message, type = 'info') => {
    const id = Date.now() + Math.random();
    dispatch({ type: 'ADD_TOAST', payload: { id, message, type } });
    setTimeout(() => dispatch({ type: 'REMOVE_TOAST', payload: id }), 3000);
  }, []);

  const loadFromFolder = useCallback(async (handle) => {
    const db = await readDb(handle);
    // 예전에 엑셀 숫자 그대로 저장된 생일도 여기서 날짜로 고쳐 읽음(다음 저장 때 고친 값으로 저장됨)
    dispatch({ type: 'SET_STUDENTS', payload: (db.students || []).map(s => ({ ...s, birthday: normalizeBirthday(s.birthday) })) });
    dispatch({ type: 'SET_ANNOUNCEMENTS', payload: db.announcements || [] });
    dispatch({ type: 'SET_ASSIGNMENTS', payload: db.assignments || [] });
    dispatch({ type: 'SET_DDAYS', payload: db.ddays || [] });
    dispatch({ type: 'SET_QUICK_LINKS', payload: db.quickLinks || [] });
    dispatch({ type: 'SET_POLLS', payload: db.polls || [] });
    dispatch({ type: 'SET_NEWSLETTERS', payload: db.newsletters || [] });
    dispatch({ type: 'SET_WEEKLY_PLANS_BULK', payload: db.weeklyPlans || {} });
    dispatch({ type: 'SET_STUDENT_RECORDS', payload: db.studentRecords || {} });

    const attendanceByDate = db.attendance || {};
    dispatch({ type: 'SET_MONTHLY_ATTENDANCE', payload: attendanceByDate });
    const today = toISODate(new Date());
    const attMap = {};
    (db.students || []).forEach(s => { attMap[s.id] = { status: 'present', note: '' }; });
    Object.assign(attMap, attendanceByDate[today] || {});
    dispatch({ type: 'SET_ATTENDANCE_MAP', payload: attMap });

    dispatch({ type: 'UPDATE_SETTINGS', payload: db.settings || {} });
    setIsDbLoading(false);
  }, []);

  // 최초 마운트: 저장해둔 폴더가 있는지 확인
  useEffect(() => {
    (async () => {
      try {
        const result = await getActiveFolder();
        if (!result) { setFolderStatus('none'); setIsDbLoading(false); return; }
        if (result.needsReconnect) { setFolder(result.handle); setFolderStatus('needs-reconnect'); setIsDbLoading(false); return; }
        setFolder(result);
        setFolderStatus('connected');
        await loadFromFolder(result);
      } catch (e) {
        console.error('폴더 연결 확인 중 오류:', e);
        setFolderStatus('none');
        setIsDbLoading(false);
      }
    })();

    if (localStorage.getItem('isAdmin') === 'true') dispatch({ type: 'TOGGLE_ADMIN', payload: true });
  }, [loadFromFolder]);

  const connectFolder = async () => {
    try {
      const handle = await pickFolder();
      setFolder(handle);
      setFolderStatus('connected');
      setIsDbLoading(true);
      await loadFromFolder(handle);
    } catch (e) {
      if (e.name !== 'AbortError') console.error('폴더 선택 오류:', e);
    }
  };

  const reconnect = async () => {
    const handle = await reconnectFolder(folder);
    if (handle) { setFolderStatus('connected'); setIsDbLoading(true); await loadFromFolder(handle); }
    else showToast('폴더 접근이 거부되었습니다.', 'error');
  };

  // 상태가 바뀔 때마다 로컬 폴더에 저장 (디바운스)
  useEffect(() => {
    if (isDbLoading || folderStatus !== 'connected' || !folder) return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    const save = () => {
      pendingSave.current = null;
      const payload = {
        students: state.students, attendance: state.monthlyAttendance, announcements: state.announcements,
        assignments: state.assignments, ddays: state.ddays, polls: state.polls, newsletters: state.newsletters,
        quickLinks: state.quickLinks, weeklyPlans: state.weeklyPlans, studentRecords: state.studentRecords,
        settings: state.settings,
      };
      writeQueue.current = writeQueue.current
        .then(() => writeDb(folder, payload))
        .catch(err => { console.error('로컬 저장 오류:', err); showToast('저장하지 못했어요. 폴더 연결을 확인해 주세요.', 'error'); });
    };
    pendingSave.current = save;
    saveTimer.current = setTimeout(save, 400);
    return () => clearTimeout(saveTimer.current);
  }, [state, isDbLoading, folderStatus, folder]);

  const saveWeeklyPlan = useCallback((weekKey, day, period, subject, content) => {
    dispatch({ type: 'SET_WEEKLY_PLAN', payload: { weekKey, day, period, subject, content } });
  }, []);

  const addDDay = (dday) => dispatch({ type: 'ADD_DDAY', payload: { id: generateUUID(), name: dday.name, date: dday.date, emoji: dday.emoji } });
  const deleteDDayFunc = (id) => dispatch({ type: 'DELETE_DDAY', payload: id });
  const updatePointsFunc = (studentId, _cur, amount) => dispatch({ type: 'UPDATE_POINTS', payload: { studentId, amount } });
  const resetPointsFunc = (studentId) => dispatch({ type: 'RESET_POINTS', payload: { studentId } });


  const loadAttendanceByDateFunc = useCallback((date) => {
    dispatch({ type: 'SET_SELECTED_ATTENDANCE_DATE', payload: date });
    const attMap = {};
    state.students.forEach(s => { attMap[s.id] = { status: 'present', note: '' }; });
    Object.assign(attMap, state.monthlyAttendance[date] || {});
    dispatch({ type: 'SET_ATTENDANCE_MAP', payload: attMap });
  }, [state.students, state.monthlyAttendance]);

  const loadMonthlyAttendanceFunc = useCallback(() => {}, []);
  const saveAttendanceDayFunc = (date, records) => dispatch({ type: 'SAVE_ATTENDANCE_DAY', payload: { date, records } });

  const addStudentRecordFunc = (studentId, type, content, category = '') => {
    const date = toISODate(new Date());
    dispatch({ type: 'ADD_STUDENT_RECORD', payload: { studentId, record: { id: generateUUID(), type, content, category, date } } });
    showToast('기록이 저장되었습니다.', 'success');
  };
  const deleteStudentRecordFunc = (studentId, recordId) => dispatch({ type: 'DELETE_STUDENT_RECORD', payload: { studentId, recordId } });
  const updateStudentFunc = (studentData) => dispatch({ type: 'UPDATE_STUDENT', payload: studentData });
  const bulkUpdateStudentsFunc = (updates) => dispatch({ type: 'BULK_UPDATE_STUDENTS', payload: updates });

  const votePollFunc = (pollId, optionIndex) => {
    const poll = state.polls.find(p => p.id === pollId);
    if (!poll) return;
    const newVotes = [...poll.votes];
    newVotes[optionIndex] += 1;
    dispatch({ type: 'VOTE_POLL', payload: { pollId, newVotes } });
  };

  const collectNewsletterFunc = (id, newCollected) => dispatch({ type: 'COLLECT_NEWSLETTER', payload: { id, collected: newCollected } });
  const addLinkFunc = (linkData) => dispatch({ type: 'ADD_LINK', payload: linkData });
  const updateLinkFunc = (linkData) => dispatch({ type: 'UPDATE_LINK', payload: linkData });
  const deleteLinkFunc = (id) => dispatch({ type: 'DELETE_LINK', payload: id });
  const addAssignmentFunc = (assignment) => dispatch({ type: 'ADD_ASSIGNMENT', payload: assignment });
  const updateAssignmentSubmissionsFunc = (id, submissions) => dispatch({ type: 'UPDATE_ASSIGNMENT_SUBMISSIONS', payload: { id, submissions } });
  const deleteAssignmentFunc = (id) => dispatch({ type: 'DELETE_ASSIGNMENT', payload: id });
  const addAnnouncementFunc = (a) => dispatch({ type: 'ADD_ANNOUNCEMENT', payload: a });
  const updateAnnouncementFunc = (a) => dispatch({ type: 'UPDATE_ANNOUNCEMENT', payload: a });
  const deleteAnnouncementFunc = (id) => dispatch({ type: 'DELETE_ANNOUNCEMENT', payload: id });
  const addPollFunc = (p) => dispatch({ type: 'ADD_POLL', payload: p });
  const deletePollFunc = (id) => dispatch({ type: 'DELETE_POLL', payload: id });
  const updateSettingsFunc = (newSettings) => dispatch({ type: 'UPDATE_SETTINGS', payload: newSettings });

  const loginFunc = (password) => {
    if (!state.settings.adminPassword) {
      dispatch({ type: 'UPDATE_SETTINGS', payload: { adminPassword: password } });
      localStorage.setItem('isAdmin', 'true');
      dispatch({ type: 'TOGGLE_ADMIN', payload: true });
      showToast('선생님 비밀번호가 설정되고, 교사 모드로 전환되었습니다.', 'success');
      return true;
    }
    if (password === state.settings.adminPassword) {
      localStorage.setItem('isAdmin', 'true');
      dispatch({ type: 'TOGGLE_ADMIN', payload: true });
      showToast('교사 모드로 전환되었습니다.', 'success');
      return true;
    }
    showToast('비밀번호가 틀렸습니다.', 'error');
    return false;
  };
  const logoutFunc = () => {
    localStorage.removeItem('isAdmin');
    dispatch({ type: 'TOGGLE_ADMIN', payload: false });
    showToast('학생 모드로 전환되었습니다.', 'info');
  };

  const value = {
    state, dispatch, showToast,
    updatePoints: updatePointsFunc, resetPoints: resetPointsFunc,
    loadAttendanceByDate: loadAttendanceByDateFunc, loadMonthlyAttendance: loadMonthlyAttendanceFunc, saveAttendanceDay: saveAttendanceDayFunc,
    updateStudent: updateStudentFunc, bulkUpdateStudents: bulkUpdateStudentsFunc, addStudentRecord: addStudentRecordFunc, deleteStudentRecord: deleteStudentRecordFunc,
    saveWeeklyPlan, updateSettings: updateSettingsFunc, addDDay, deleteDDay: deleteDDayFunc,
    votePoll: votePollFunc, collectNewsletter: collectNewsletterFunc,
    addLink: addLinkFunc, updateLink: updateLinkFunc, deleteLink: deleteLinkFunc,
    addAssignment: addAssignmentFunc, updateAssignmentSubmissions: updateAssignmentSubmissionsFunc, deleteAssignment: deleteAssignmentFunc,
    addAnnouncement: addAnnouncementFunc, updateAnnouncement: updateAnnouncementFunc, deleteAnnouncement: deleteAnnouncementFunc,
    addPoll: addPollFunc, deletePoll: deletePollFunc,
    login: loginFunc, logout: logoutFunc,
    isDbLoading, getWeekStart, toISODate,
  };

  if (folderStatus === 'loading') {
    return <div className="folder-gate"><p>불러오는 중...</p></div>;
  }
  if (folderStatus !== 'connected') {
    return <FolderGate status={folderStatus} onConnect={connectFolder} onReconnect={reconnect} />;
  }

  return <AppContext.Provider value={value}>{children}</AppContext.Provider>;
}

function useApp() {
  const ctx = useContext(AppContext);
  if (!ctx) throw new Error('useApp must be used within AppProvider');
  return ctx;
}

function FolderGate({ status, onConnect, onReconnect }) {
  return (
    <div className="folder-gate">
      <div className="folder-gate-card">
        <div className="folder-gate-icon">📁</div>
        {status === 'needs-reconnect' ? (
          <>
            <h2>데이터 폴더 접근이 끊겼습니다</h2>
            <p>브라우저를 새로 시작하면 보안상 폴더 접근을 다시 허용해야 합니다.</p>
            <button className="btn btn-primary btn-lg" onClick={onReconnect}>폴더 다시 연결하기</button>
          </>
        ) : (
          <>
            <h2>학급 데이터를 저장할 폴더를 선택해주세요</h2>
            <p>선택한 폴더 안에 db.json 파일로 모든 데이터가 저장됩니다. 한 번만 선택하면 됩니다.</p>
            <button className="btn btn-primary btn-lg" onClick={onConnect}>폴더 선택하기</button>
            <p className="folder-gate-note">⚠️ 크롬 또는 엣지 브라우저에서만 지원됩니다.</p>
          </>
        )}
      </div>
    </div>
  );
}

/* ===================================================================
   라우팅 (해시 기반)
   =================================================================== */
function useHashRoute() {
  const [hash, setHash] = useState(window.location.hash.slice(1) || '/');
  useEffect(() => {
    const onChange = () => setHash(window.location.hash.slice(1) || '/');
    window.addEventListener('hashchange', onChange);
    return () => window.removeEventListener('hashchange', onChange);
  }, []);
  return hash;
}

/* ===================================================================
   Toast
   =================================================================== */
function Toast() {
  const { state } = useApp();
  return (
    <div className="toast-container">
      {state.toasts.map(t => (
        <div key={t.id} className={`toast ${t.type}`}>
          <span>{t.type === 'success' ? '✅' : t.type === 'error' ? '❌' : t.type === 'warning' ? '⚠️' : 'ℹ️'}</span>
          <span>{t.message}</span>
        </div>
      ))}
    </div>
  );
}

/* ===================================================================
   Icon (이모지 대신 쓰는 선 아이콘)
   =================================================================== */
const ICON_PATHS = {
  dashboard: <><rect x="3" y="3" width="7" height="9" rx="1.5" /><rect x="14" y="3" width="7" height="5" rx="1.5" /><rect x="14" y="12" width="7" height="9" rx="1.5" /><rect x="3" y="16" width="7" height="5" rx="1.5" /></>,
  weekly: <path d="M4 6h16M4 12h16M4 18h10" />,
  attendance: <><rect x="4" y="4" width="16" height="17" rx="2" /><path d="M9 2.5v3M15 2.5v3M8.5 13l2.5 2.5 4.5-5" /></>,
  learning: <><path d="M4 5.5C4 4.7 4.7 4 5.5 4H11v16H5.5c-.8 0-1.5-.7-1.5-1.5z" /><path d="M20 5.5c0-.8-.7-1.5-1.5-1.5H13v16h5.5c.8 0 1.5-.7 1.5-1.5z" /></>,
  calendar: <><rect x="3.5" y="5" width="17" height="15.5" rx="2" /><path d="M3.5 10h17M8 3v4M16 3v4" /></>,
  tools: <path d="M14.5 6.5l3 3M4 20l4.5-1L19 8.5a2.1 2.1 0 0 0-3-3L5.5 16z" />,
  links: <><path d="M10 14a4.5 4.5 0 0 0 6.4 0l3-3a4.5 4.5 0 0 0-6.4-6.4l-1 1" /><path d="M14 10a4.5 4.5 0 0 0-6.4 0l-3 3a4.5 4.5 0 0 0 6.4 6.4l1-1" /></>,
  students: <><circle cx="9" cy="8" r="3.5" /><path d="M2.5 20c.8-3.6 3.4-5.5 6.5-5.5s5.7 1.9 6.5 5.5" /><circle cx="17" cy="9" r="2.5" /><path d="M17.5 14.5c2.2.3 3.6 1.9 4 4.5" /></>,
  gear: <><path d="M4 7h9M17 7h3M4 17h3M11 17h9" /><circle cx="15" cy="7" r="2" /><circle cx="9" cy="17" r="2" /></>,
  plus: <path d="M12 5v14M5 12h14" />,
  edit: <path d="M14.5 6.5l3 3M4 20l4.5-1L19 8.5a2.1 2.1 0 0 0-3-3L5.5 16z" />,
  trash: <path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3" />,
  move: <path d="M12 3v18M3 12h18M12 3l-3 3M12 3l3 3M12 21l-3-3M12 21l3-3M3 12l3-3M3 12l3 3M21 12l-3-3M21 12l-3 3" />,
  resize: <path d="M20 14v6h-6M20 20l-7-7M10 4H4v6M4 4l7 7" />,
  menu: <path d="M4 7h16M4 12h16M4 17h16" />,
  chevron: <path d="M15 5l-7 7 7 7" />,
  cake: <path d="M4 21h16M5 21v-7a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2v7M5 16c1.5 1 3 1 4.5 0s3-1 4.5 0 3 1 4.5 0M12 12V8M12 5.5c.8-.6.8-1.6 0-2.5-.8.9-.8 1.9 0 2.5z" />,
  search: <><circle cx="11" cy="11" r="6.5" /><path d="M16 16l4.5 4.5" /></>,
  minus: <path d="M5 12h14" />,
  lock: <><rect x="5" y="11" width="14" height="9.5" rx="2" /><path d="M8.5 11V8a3.5 3.5 0 0 1 7 0v3" /></>,
  expand: <path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5" />,
  shrink: <path d="M9 4v5H4M15 4v5h5M9 20v-5H4M15 20v-5h5" />,
  sun: <><circle cx="12" cy="12" r="4" /><path d="M12 2.5v2M12 19.5v2M2.5 12h2M19.5 12h2M5.3 5.3l1.4 1.4M17.3 17.3l1.4 1.4M5.3 18.7l1.4-1.4M17.3 6.7l1.4-1.4" /></>,
};
function Icon({ name, size = 20 }) {
  return <svg aria-hidden="true" width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">{ICON_PATHS[name]}</svg>;
}

/* ===================================================================
   Thermometer
   =================================================================== */
function Thermometer({ current, goal, reward, isAdmin, onUpdate }) {
  const percentage = Math.min((current / goal) * 100, 100);
  const isGoalReached = current >= goal;
  const [isEditing, setIsEditing] = useState(false);
  const [editGoal, setEditGoal] = useState(goal);
  const [editReward, setEditReward] = useState(reward || '');

  const handleSave = () => { onUpdate && onUpdate(parseInt(editGoal) || 1500, editReward); setIsEditing(false); };

  return (
    <>
      <div className="wcard-head">
        <h2 className="wcard-title">우리 반 칭찬 온도계</h2>
        {isAdmin && !isEditing && <button type="button" className="chip-btn" onClick={() => { setEditGoal(goal); setEditReward(reward || ''); setIsEditing(true); }}>목표 바꾸기</button>}
      </div>
      <div className="wcard-body thermo">
        {isEditing && isAdmin && (
          <div className="thermo-form">
            <label className="thermo-field"><span>목표 점수</span><input type="number" min="1" value={editGoal} onChange={e => setEditGoal(e.target.value)} className="form-input" /></label>
            <label className="thermo-field thermo-field--wide"><span>보상</span><input type="text" value={editReward} onChange={e => setEditReward(e.target.value)} placeholder="예: 금요일 영화 감상" className="form-input" /></label>
            <div className="thermo-form-actions">
              <button type="button" className="btn btn-secondary btn-sm" onClick={() => setIsEditing(false)}>취소</button>
              <button type="button" className="btn btn-primary btn-sm" onClick={handleSave}>저장</button>
            </div>
          </div>
        )}
        <div className="thermo-score"><span className="num-xl">{current}</span><span className="thermo-goal">/ {goal}점</span></div>
        <div>
          <div className="thermo-track" role="progressbar" aria-label="칭찬 점수" aria-valuemin={0} aria-valuemax={goal} aria-valuenow={current}>
            <span className={isGoalReached ? 'is-done' : ''} style={{ width: `${percentage}%` }} />
          </div>
          <div className="thermo-ticks">{[0, 0.25, 0.5, 0.75, 1].map(f => <span key={f}>{Math.round(goal * f)}</span>)}</div>
        </div>
        <p className="thermo-note">
          {isGoalReached
            ? <><strong>목표 달성!</strong> {reward ? <>이제 <strong>{reward}</strong> 시간이에요.</> : '다 같이 축하해요.'}</>
            : <><strong>{goal - current}점만 더!</strong> {reward ? <>목표를 채우면 <strong>{reward}</strong>을(를) 해요.</> : '보상을 정하면 여기에 보여요.'}</>}
        </p>
      </div>
    </>
  );
}

/* ===================================================================
   AdminSettingsModal (캘린더 ID + Gemini 키)
   =================================================================== */
function AdminSettingsModal({ onClose }) {
  const { state, updateSettings, showToast } = useApp();
  const { settings } = state;
  const [calendarId1, setCalendarId1] = useState(settings.calendarId1 || '');
  const [calendarId2, setCalendarId2] = useState(settings.calendarId2 || '');
  const [calendarId3, setCalendarId3] = useState(settings.calendarId3 || '');
  const [geminiApiKey, setGeminiApiKey] = useState(settings.geminiApiKey || '');
  const [googleApiKey, setGoogleApiKey] = useState(settings.googleApiKey || '');

  const handleSave = () => {
    updateSettings({ ...settings, calendarId1, calendarId2, calendarId3, geminiApiKey, googleApiKey: googleApiKey.trim() });
    showToast('설정이 저장되었습니다.', 'success');
    onClose();
  };

  return (
    <div className="settings-modal-overlay" onClick={onClose}>
      <div className="settings-modal-content" onClick={e => e.stopPropagation()}>
        <div className="settings-modal-header"><h3>⚙️ 관리자 설정</h3><button className="settings-close-btn" onClick={onClose}>×</button></div>
        <div className="settings-modal-body">
          <div className="settings-form-group">
            <label>📅 구글 캘린더 ID 1 (필수)</label>
            <input type="text" className="settings-form-input" placeholder="예: ...@group.calendar.google.com" value={calendarId1} onChange={e => setCalendarId1(e.target.value)} />
          </div>
          <div className="settings-form-group">
            <label>📅 구글 캘린더 ID 2 (선택)</label>
            <input type="text" className="settings-form-input" value={calendarId2} onChange={e => setCalendarId2(e.target.value)} />
          </div>
          <div className="settings-form-group">
            <label>📅 구글 캘린더 ID 3 (선택)</label>
            <input type="text" className="settings-form-input" value={calendarId3} onChange={e => setCalendarId3(e.target.value)} />
            <small className="settings-help-text">일정 및 공지 페이지에 표시할 구글 캘린더 ID를 입력하세요. (최대 3개)</small>
          </div>
          <div className="settings-form-group">
            <label>📅 Google API 키 (주간 일정에 달력 일정 넣기)</label>
            <input type="password" className="settings-form-input" placeholder="AIza…" value={googleApiKey} onChange={e => setGoogleApiKey(e.target.value)} />
            <small className="settings-help-text">Google Cloud 콘솔에서 Calendar API를 켜고 만든 API 키를 넣으면, 위 달력들의 일정이 주간 일정 표의 해당 교시에 함께 나옵니다. 달력이 ‘공개’로 공유되어 있어야 읽을 수 있어요.</small>
          </div>
          <div className="settings-form-group">
            <label>✨ Gemini API 키</label>
            <input type="password" className="settings-form-input" placeholder="AI 종합의견 생성에 사용됩니다" value={geminiApiKey} onChange={e => setGeminiApiKey(e.target.value)} />
            <small className="settings-help-text">서버가 없어서 이 키는 선택한 데이터 폴더의 db.json에 그대로 저장됩니다. 이 컴퓨터를 다른 사람과 같이 쓴다면 주의하세요.</small>
          </div>
        </div>
        <div className="settings-modal-footer">
          <button className="settings-btn settings-btn-secondary" onClick={onClose}>취소</button>
          <button className="settings-btn settings-btn-primary" onClick={handleSave}>저장</button>
        </div>
      </div>
    </div>
  );
}

/* ===================================================================
   Layout
   =================================================================== */
const STUDENT_ALLOWED_PATHS = ['/', '/weekly', '/learning', '/calendar', '/tools', '/links', '/students'];
const NAV_ITEMS = [
  { path: '/', tone: 'carrot', label: '대시보드', icon: <Icon name="dashboard" /> },
  { path: '/weekly', tone: 'apricot', label: '주간 일정', icon: <Icon name="weekly" /> },
  { path: '/attendance', tone: 'tangerine', label: '출석 관리', icon: <Icon name="attendance" /> },
  { path: '/learning', tone: 'cocoa', label: '학습 관리', icon: <Icon name="learning" /> },
  { path: '/calendar', tone: 'coral', label: '일정 & 공지', icon: <Icon name="calendar" /> },
  { path: '/tools', tone: 'carrot', label: '학급 도구', icon: <Icon name="tools" /> },
  { path: '/links', tone: 'apricot', label: '바로가기', icon: <Icon name="links" /> },
  { path: '/students', tone: 'tangerine', label: '학생 관리', icon: <Icon name="students" /> },
];
const TONES = ['carrot', 'apricot', 'tangerine', 'cocoa', 'coral'];
// 과목마다 늘 같은 색 (모르는 과목은 이름으로 골고루 나눔)
const SUBJECT_TONES = { 국어: 'carrot', 수학: 'apricot', 과학: 'tangerine', 사회: 'cocoa', 영어: 'coral', 도덕: 'cocoa', 체육: 'tangerine', 음악: 'coral', 미술: 'carrot', 실과: 'tangerine', 창체: 'apricot' };
const subjectTone = (name) => { const n = (name || '').trim(); if (!n) return undefined; const hit = Object.keys(SUBJECT_TONES).find(k => n.startsWith(k)); return hit ? SUBJECT_TONES[hit] : TONES[[...n].reduce((a, c) => a + c.charCodeAt(0), 0) % TONES.length]; };
const genderTone = (gender) => (gender === 'F' ? 'coral' : 'tangerine'); // 학생 카드 색은 남녀로만 구분
const WIDGET_TONES = { schedule: 'apricot', thermometer: 'carrot', lunch: 'cocoa', dday: 'coral', ranking: 'tangerine', notices: 'apricot' };
const PAGE_TITLES = { '/': '대시보드', '/weekly': '주간 일정', '/attendance': '출석 관리', '/learning': '학습 관리', '/calendar': '일정 & 공지', '/tools': '학급 도구', '/links': '바로가기', '/students': '학생 관리' };
const MOBILE_LABELS = { '/': '대시보드', '/weekly': '주간', '/attendance': '출석', '/learning': '학습', '/calendar': '일정', '/tools': '도구', '/links': '링크', '/students': '학생' };

function Layout({ pathname, children }) {
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [collapsed, setCollapsed] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showPinModal, setShowPinModal] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const { state, login, logout, updateSettings } = useApp();
  const { isAdmin, settings } = state;

  // 메뉴 순서: settings.navOrder(경로 목록) 순서대로, 모르는 메뉴는 뒤에 붙임
  const navItems = useMemo(() => {
    const order = Array.isArray(settings?.navOrder) ? settings.navOrder : [];
    const rank = (p) => { const i = order.indexOf(p); return i < 0 ? 99 + NAV_ITEMS.findIndex(n => n.path === p) : i; };
    return [...NAV_ITEMS].sort((a, b) => rank(a.path) - rank(b.path));
  }, [settings?.navOrder]);
  const [navDraft, setNavDraft] = useState(null);
  const navDraftRef = useRef(null);
  const dragPath = useRef(null);
  const shownNav = navDraft || navItems;
  const isCustomOrder = navItems.some((n, i) => n.path !== NAV_ITEMS[i].path);
  const moveNav = (list, from, to) => { const next = [...list]; const [m] = next.splice(from, 1); next.splice(to, 0, m); return next; };
  const saveNav = (list) => updateSettings({ navOrder: list.map(n => n.path) });
  const endNavDrag = (save) => {
    if (save && navDraftRef.current) saveNav(navDraftRef.current);
    dragPath.current = null; navDraftRef.current = null; setNavDraft(null);
  };
  const navDragProps = (item) => (!isAdmin ? {} : {
    draggable: true,
    onDragStart: (e) => { dragPath.current = item.path; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', item.path); navDraftRef.current = navItems; setNavDraft(navItems); },
    onDragOver: (e) => {
      if (!dragPath.current) return;
      e.preventDefault();
      const list = navDraftRef.current || navItems;
      const from = list.findIndex(n => n.path === dragPath.current), to = list.findIndex(n => n.path === item.path);
      if (from !== to) { const next = moveNav(list, from, to); navDraftRef.current = next; setNavDraft(next); }
    },
    onDrop: (e) => { e.preventDefault(); endNavDrag(true); },
    onDragEnd: () => endNavDrag(true),
    // 키보드: Alt+↑/↓ 로 한 칸씩
    onKeyDown: (e) => {
      const d = e.altKey ? { ArrowUp: -1, ArrowDown: 1 }[e.key] : 0;
      if (!d) return;
      e.preventDefault();
      const i = navItems.findIndex(n => n.path === item.path), j = i + d;
      if (j >= 0 && j < navItems.length) saveNav(moveNav(navItems, i, j));
    },
  });

  const doLogin = () => { if (login(passwordInput)) { setShowPinModal(false); setPasswordInput(''); } };

  const today = new Date();
  const isDashboard = pathname === '/';
  const pageTitle = PAGE_TITLES[pathname] || '대시보드';
  const currentNav = NAV_ITEMS.find(n => n.path === pathname);
  const isReadOnlyPath = !isAdmin && !STUDENT_ALLOWED_PATHS.includes(pathname);

  return (
    <div className="app-layout">
      {sidebarOpen && <div className="sidebar-overlay" onClick={() => setSidebarOpen(false)} />}
      <aside className={`sidebar ${collapsed ? 'collapsed' : ''} ${sidebarOpen ? 'open' : ''}`}>
        <div className="sidebar-header">
          <div className="sidebar-logo">
            <span className="logo-mark">4·2</span>
            <div className="logo-text"><h1>하남중앙초등학교</h1><span className="logo-sub">4학년 2반 · 김대붕 선생님</span></div>
          </div>
          <button type="button" className="sidebar-toggle" aria-label={collapsed ? '메뉴 펼치기' : '메뉴 접기'} onClick={() => { setCollapsed(!collapsed); setSidebarOpen(false); }}><Icon name="chevron" size={18} /></button>
        </div>
        <nav className={`sidebar-nav ${navDraft ? 'is-sorting' : ''}`} aria-label="주 메뉴">
          {shownNav.map(item => {
            if (item.path === '/attendance' && !isAdmin) return null;
            const isAllowed = isAdmin || STUDENT_ALLOWED_PATHS.includes(item.path);
            const isActive = pathname === item.path;
            return (
              <a key={item.path} href={`#${item.path}`} aria-current={isActive ? 'page' : undefined} data-tone={item.tone} className={`nav-item ${isActive ? 'active' : ''} ${!isAllowed ? 'nav-disabled' : ''} ${navDraft && dragPath.current === item.path ? 'is-dragging' : ''}`} onClick={() => setSidebarOpen(false)} {...navDragProps(item)}>
                <span className="nav-icon">{item.icon}</span>
                <span className="nav-label">{item.label}</span>
                {!isAllowed && <span className="nav-lock">선생님</span>}
              </a>
            );
          })}
        </nav>
        {isAdmin && !collapsed && (
          <div className="nav-order-hint">
            <span>메뉴를 끌어서 순서를 바꿀 수 있어요</span>
            {isCustomOrder && <button type="button" onClick={() => updateSettings({ navOrder: [] })}>기본 순서로</button>}
          </div>
        )}
      </aside>

      <main className="main-content" data-tone={currentNav?.tone}>
        <header className="page-header">
          <button type="button" className="menu-btn" aria-label="메뉴 열기" onClick={() => setSidebarOpen(true)}><Icon name="menu" /></button>
          <div className="header-content">
            {!isDashboard && currentNav && <span className="page-icon" aria-hidden="true">{currentNav.icon}</span>}
            <div className="header-titles">
              {isDashboard
                ? <h2 className="page-title">{formatDate(today, 'md')} {getDayName(today)}요일</h2>
                : <><span className="page-date">{formatDate(today)}</span><h2 className="page-title">{pageTitle}</h2></>}
            </div>
            <div className="header-actions">
              <div className="mode-switch" role="group" aria-label="보기 모드">
                <button type="button" aria-pressed={!isAdmin} onClick={() => { if (isAdmin) logout(); }}>학생 화면</button>
                <button type="button" aria-pressed={isAdmin} onClick={() => { if (!isAdmin) setShowPinModal(true); }}>선생님</button>
              </div>
              {isAdmin && <button type="button" className="round-btn" aria-label="설정" onClick={() => setShowSettings(true)}><Icon name="gear" /></button>}
            </div>
          </div>
        </header>
        <div className={`pages-container ${isReadOnlyPath ? 'student-read-only' : ''}`}>{children}</div>
      </main>

      <nav className="mobile-nav" aria-label="주 메뉴">
        {navItems.map(item => {
          if (item.path === '/attendance' && !isAdmin) return null;
          const isAllowed = isAdmin || STUDENT_ALLOWED_PATHS.includes(item.path);
          const isActive = pathname === item.path;
          return (
            <a key={item.path} href={`#${item.path}`} aria-current={isActive ? 'page' : undefined} data-tone={item.tone} className={`mobile-nav-item ${isActive ? 'active' : ''} ${!isAllowed ? 'nav-disabled' : ''}`}>
              {item.icon}<span>{MOBILE_LABELS[item.path]}</span>
            </a>
          );
        })}
      </nav>

      {showPinModal && (
        <div className="modal-overlay" onClick={() => setShowPinModal(false)}>
          <div className="modal-content" onClick={e => e.stopPropagation()} style={{ maxWidth: '350px' }}>
            <div className="modal-header">
              <h3>{settings.adminPassword ? '선생님 로그인' : '선생님 비밀번호 설정'}</h3>
              <button type="button" className="close-btn" aria-label="닫기" onClick={() => setShowPinModal(false)}>×</button>
            </div>
            <div className="modal-body" style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
              {!settings.adminPassword && <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)' }}>처음 사용하시는군요! 앞으로 선생님 모드 전환에 쓸 비밀번호를 정해주세요.</p>}
              <div>
                <label htmlFor="admin-password" style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.9rem', color: 'var(--text-secondary)' }}>비밀번호</label>
                <input id="admin-password" type="password" className="form-input" placeholder="비밀번호" value={passwordInput} onChange={e => setPasswordInput(e.target.value)} onKeyDown={e => e.key === 'Enter' && doLogin()} autoFocus />
              </div>
            </div>
            <div className="modal-footer">
              <button type="button" className="btn btn-primary" style={{ width: '100%' }} onClick={doLogin}>{settings.adminPassword ? '로그인' : '설정하고 시작하기'}</button>
            </div>
          </div>
        </div>
      )}

      {showSettings && isAdmin && <AdminSettingsModal onClose={() => setShowSettings(false)} />}
    </div>
  );
}

/* ===================================================================
   DashboardGrid — 위젯 이동/크기 조절 (라이브러리 없이 12열 격자)
   =================================================================== */
const GRID_COLS = 12, GRID_ROW_H = 72, GRID_GAP = 20, GRID_MIN_W = 3, GRID_MIN_H = 3, GRID_MAX_H = 12;
const GRID_STACK_BELOW = 900; // 이보다 좁으면 한 줄로 쌓고 편집은 끔
const DEFAULT_DASHBOARD_LAYOUT = [
  { id: 'schedule', x: 0, y: 0, w: 8, h: 4 },
  { id: 'thermometer', x: 8, y: 0, w: 4, h: 4 },
  { id: 'lunch', x: 0, y: 4, w: 4, h: 5 },
  { id: 'dday', x: 4, y: 4, w: 4, h: 5 },
  { id: 'ranking', x: 8, y: 4, w: 4, h: 5 },
  { id: 'notices', x: 0, y: 9, w: 12, h: 4 },
];

function clampItem(it) {
  const w = Math.max(GRID_MIN_W, Math.min(GRID_COLS, Math.round(it.w)));
  const h = Math.max(GRID_MIN_H, Math.min(GRID_MAX_H, Math.round(it.h)));
  return { ...it, w, h, x: Math.max(0, Math.min(GRID_COLS - w, Math.round(it.x))), y: Math.max(0, Math.round(it.y)) };
}
const gridCollides = (a, b) => a.id !== b.id && a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

// 움직인 위젯은 그 자리에 두고, 나머지는 위에서부터 빈 자리를 찾아 채운 뒤 움직인 위젯도 위로 당김
function settleLayout(layout, activeId) {
  const active = layout.find(i => i.id === activeId);
  const placed = active ? [active] : [];
  const rest = layout.filter(i => i !== active).sort((a, b) => a.y - b.y || a.x - b.x);
  for (const it of rest) {
    let y = 0;
    while (placed.some(p => gridCollides({ ...it, y }, p))) y++;
    placed.push({ ...it, y });
  }
  if (active) {
    let y = active.y;
    while (y > 0 && !placed.some(p => gridCollides({ ...active, y: y - 1 }, p))) y--;
    placed[0] = { ...active, y };
  }
  return layout.map(i => placed.find(p => p.id === i.id));
}

// 저장된 배치(예전 Next.js 버전의 {lg:[{i,...}]} 형식 포함)를 읽고, 모르는 값은 기본 배치로
function normalizeLayout(saved) {
  const list = Array.isArray(saved) ? saved : (saved && Array.isArray(saved.lg) ? saved.lg.map(i => ({ ...i, id: i.i })) : []);
  const merged = DEFAULT_DASHBOARD_LAYOUT.map(d => {
    const s = list.find(x => x && x.id === d.id);
    return s && [s.x, s.y, s.w, s.h].every(Number.isFinite) ? clampItem({ ...d, x: s.x, y: s.y, w: s.w, h: s.h }) : d;
  });
  // 겹치는 데가 없으면 저장한 그대로 보여줌(다시 정렬하면 새로고침 뒤 배치가 달라 보이던 문제)
  const overlaps = merged.some(a => merged.some(b => gridCollides(a, b)));
  return overlaps ? settleLayout(merged, null) : merged;
}

function DashboardGrid({ layout, editing, onChange, widgets }) {
  const boxRef = useRef(null);
  const dragRef = useRef(null);
  const [width, setWidth] = useState(0);
  const [draft, setDraft] = useState(null);
  const [activeId, setActiveId] = useState(null);
  const [drag, setDrag] = useState(null); // 끄는 동안 위젯이 마우스를 픽셀 단위로 따라가게 {id, mode, dx, dy, start}
  const rafRef = useRef(0);
  useEffect(() => () => cancelAnimationFrame(rafRef.current), []);

  useEffect(() => {
    const el = boxRef.current;
    if (!el) return;
    const ro = new ResizeObserver(entries => setWidth(entries[0].contentRect.width));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const shown = draft || layout;
  const stacked = width > 0 && width < GRID_STACK_BELOW;
  const colW = (width - GRID_GAP * (GRID_COLS - 1)) / GRID_COLS;
  const unitX = colW + GRID_GAP, unitY = GRID_ROW_H + GRID_GAP;

  // 마우스가 가리키는 격자 칸 (위젯 모서리가 아니라 손가락 위치로 판단)
  const cellAt = (clientX, clientY) => {
    const r = boxRef.current.getBoundingClientRect();
    return { c: Math.floor((clientX - r.left) / unitX), r: Math.floor((clientY - r.top) / unitY) };
  };
  const apply = (id, mode, start, dc, dr, cell) => {
    if (mode === 'move' && cell) {
      // 마우스 아래에 다른 위젯이 있으면 그 위젯과 자리를 맞바꿈
      const target = layout.find(i => i.id !== id && cell.c >= i.x && cell.c < i.x + i.w && cell.r >= i.y && cell.r < i.y + i.h);
      if (target) {
        const a = clampItem({ ...start, x: target.x, y: target.y });
        const b = clampItem({ ...target, x: start.x, y: start.y });
        return settleLayout(layout.map(i => (i.id === id ? a : i.id === target.id ? b : i)), id);
      }
    }
    const cand = clampItem(mode === 'move' ? { ...start, x: start.x + dc, y: start.y + dr } : { ...start, w: start.w + dc, h: start.h + dr });
    return settleLayout(layout.map(i => (i.id === id ? cand : i)), id);
  };

  const onPointerDown = (e, id, mode) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.currentTarget.setPointerCapture(e.pointerId);
    const start = layout.find(i => i.id === id);
    dragRef.current = { id, mode, sx: e.clientX, sy: e.clientY, x: e.clientX, y: e.clientY, start };
    setActiveId(id);
    setDrag({ id, mode, dx: 0, dy: 0, start });
  };
  const onPointerMove = (e) => {
    const d = dragRef.current;
    if (!d) return;
    d.x = e.clientX; d.y = e.clientY;
    if (rafRef.current) return; // 화면 한 프레임에 한 번만 계산
    rafRef.current = requestAnimationFrame(() => {
      rafRef.current = 0;
      const cur = dragRef.current;
      if (!cur) return;
      const dx = cur.x - cur.sx, dy = cur.y - cur.sy;
      const next = apply(cur.id, cur.mode, cur.start, Math.round(dx / unitX), Math.round(dy / unitY), cellAt(cur.x, cur.y));
      setDraft(next);
      setDrag({ id: cur.id, mode: cur.mode, dx, dy, start: cur.start });
    });
  };
  const onPointerUp = (e) => {
    const d = dragRef.current;
    if (!d) return;
    cancelAnimationFrame(rafRef.current); rafRef.current = 0;
    dragRef.current = null;
    if (e?.type === 'pointerup') { d.x = e.clientX; d.y = e.clientY; }
    // 마지막 위치는 손을 뗀 지점에서 바로 계산 — 빠르게 끌면 프레임 갱신 전에 손을 떼서 저장이 빠지던 문제
    const moved = d.x !== d.sx || d.y !== d.sy;
    const next = moved ? apply(d.id, d.mode, d.start, Math.round((d.x - d.sx) / unitX), Math.round((d.y - d.sy) / unitY), cellAt(d.x, d.y)) : null;
    setActiveId(null); setDrag(null); setDraft(null);
    if (next && JSON.stringify(next) !== JSON.stringify(layout)) onChange(next);
  };
  // 키보드: 화살표 = 이동, Shift+화살표 = 크기 조절
  const onKeyDown = (e, id) => {
    const delta = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
    if (!delta) return;
    e.preventDefault();
    onChange(apply(id, e.shiftKey ? 'resize' : 'move', layout.find(i => i.id === id), delta[0], delta[1]));
  };

  // 격자 모드에서는 화면 요소 순서를 고정 — 끄는 중에 순서가 바뀌면 브라우저가 포인터를 놓쳐서 저장이 안 되던 원인
  const ordered = stacked ? [...shown].sort((a, b) => a.y - b.y || a.x - b.x) : shown;
  const rows = Math.max(...shown.map(i => i.y + i.h));
  const rectOf = (it) => ({ left: it.x * unitX, top: it.y * unitY, width: it.w * colW + (it.w - 1) * GRID_GAP, height: it.h * GRID_ROW_H + (it.h - 1) * GRID_GAP });
  const ghost = drag && draft ? draft.find(i => i.id === drag.id) : null;

  return (
    <div ref={boxRef} className={`dgrid ${stacked ? 'is-stacked' : ''} ${editing && !stacked ? 'is-editing' : ''}`}
      style={stacked || !width ? undefined : { height: rows * unitY - GRID_GAP }}>
      {ghost && !stacked && <div className="dgrid-ghost" style={rectOf(ghost)} aria-hidden="true" />}
      {ordered.map(it => {
        const w = widgets[it.id];
        if (!w) return null;
        let style = stacked || !width ? undefined : rectOf(it);
        if (style && drag && drag.id === it.id) {
          const s = rectOf(drag.start);
          style = drag.mode === 'move'
            ? { ...s, left: s.left + drag.dx, top: Math.max(0, s.top + drag.dy) }
            : { ...s, width: Math.max(GRID_MIN_W * unitX - GRID_GAP, s.width + drag.dx), height: Math.max(GRID_MIN_H * unitY - GRID_GAP, s.height + drag.dy) };
        }
        return (
          <section key={it.id} data-tone={WIDGET_TONES[it.id]} className={`wcard ${w.className || ''} ${activeId === it.id ? 'is-active' : ''}`} style={style} aria-label={w.label}>
            {w.content}
            {editing && !stacked && (
              <>
                <button type="button" className="grid-handle grid-handle--move" aria-label={`${w.label} 옮기기 (화살표 키로 이동, Shift+화살표로 크기 조절)`}
                  onPointerDown={e => onPointerDown(e, it.id, 'move')} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={onPointerUp} onLostPointerCapture={onPointerUp}
                  onKeyDown={e => onKeyDown(e, it.id)}><Icon name="move" size={16} />옮기기</button>
                <button type="button" className="grid-handle grid-handle--resize" aria-label={`${w.label} 크기 조절`} tabIndex={-1}
                  onPointerDown={e => onPointerDown(e, it.id, 'resize')} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={onPointerUp} onLostPointerCapture={onPointerUp}><Icon name="resize" size={14} /></button>
              </>
            )}
          </section>
        );
      })}
    </div>
  );
}

/* ===================================================================
   DashboardPage
   =================================================================== */
const PERIOD_SLOTS = { 1: [540, 580], 2: [590, 630], 3: [650, 690], 4: [750, 790], 5: [790, 830], 6: [840, 880] };
const minToClock = (m) => `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;

function DashboardPage() {
  const { state, showToast, updateSettings, addDDay, deleteDDay, addAnnouncement, updateAnnouncement, deleteAnnouncement } = useApp();
  const { students, announcements, settings, weeklyPlans, ddays = [], isAdmin } = state;

  const [lunchInfo, setLunchInfo] = useState(null);
  const [isLoadingLunch, setIsLoadingLunch] = useState(true);
  const [holiday, setHoliday] = useState(null);
  const [showDdayForm, setShowDdayForm] = useState(false);
  const [newDday, setNewDday] = useState({ name: '', date: '', emoji: '' });
  const [newNotice, setNewNotice] = useState({ title: '', content: '', type: 'general' });
  const [showNoticeForm, setShowNoticeForm] = useState(false);
  const [editNoticeId, setEditNoticeId] = useState(null);
  const [editingLayout, setEditingLayout] = useState(false);
  const [now, setNow] = useState(() => new Date());

  useEffect(() => {
    (async () => { setIsLoadingLunch(true); setLunchInfo(await fetchTodayLunch()); setIsLoadingLunch(false); })();
    (async () => { setHoliday(await fetchSchoolHoliday()); })();
    const t = setInterval(() => setNow(new Date()), 30000);
    return () => clearInterval(t);
  }, []);
  useEffect(() => { if (!isAdmin) setEditingLayout(false); }, [isAdmin]);

  const today = now;
  const dayName = getDayName(today);
  const nowMin = today.getHours() * 60 + today.getMinutes();

  const closeNoticeForm = () => { setShowNoticeForm(false); setEditNoticeId(null); setNewNotice({ title: '', content: '', type: 'general' }); };
  const handleAddNotice = () => {
    if (!newNotice.title.trim()) return;
    const dateStr = toISODate(today);
    if (editNoticeId) { updateAnnouncement({ id: editNoticeId, ...newNotice, date: dateStr }); showToast('공지사항이 수정되었습니다', 'success'); }
    else { addAnnouncement({ id: generateUUID(), ...newNotice, date: dateStr }); showToast('공지사항이 등록되었습니다', 'success'); }
    closeNoticeForm();
  };
  const handleEditNotice = (n) => { setNewNotice({ title: n.title, content: n.content, type: n.type }); setEditNoticeId(n.id); setShowNoticeForm(true); };
  const handleDeleteNotice = (id) => { if (confirm('이 공지사항을 삭제하시겠습니까?')) { deleteAnnouncement(id); showToast('공지사항이 삭제되었습니다', 'success'); } };

  const weekKey = toISODate(getWeekStart(today));
  const todayPlan = weeklyPlans[weekKey]?.[dayName] || {};
  const todaySchedule = Object.keys(todayPlan).map(Number).sort((a, b) => a - b).filter(p => todayPlan[p]?.subject && PERIOD_SLOTS[p])
    .map(p => ({ period: p, subject: todayPlan[p].subject, content: todayPlan[p].content || '', start: PERIOD_SLOTS[p][0], end: PERIOD_SLOTS[p][1] }));
  const currentClass = todaySchedule.find(c => nowMin >= c.start && nowMin <= c.end);
  const nextClass = todaySchedule.find(c => c.start > nowMin);

  let headline;
  if (holiday?.isHoliday) headline = <>오늘은 <em>{holiday.name}</em>, 쉬는 날이에요</>;
  else if (currentClass) headline = <>지금은 <em>{currentClass.period}교시 {currentClass.subject}</em>, {currentClass.end - nowMin}분 남았어요</>;
  else if (nextClass && nextClass === todaySchedule[0]) headline = <>오늘 첫 수업은 <em>{nextClass.period}교시 {nextClass.subject}</em>예요</>;
  else if (nextClass) headline = <>다음은 <em>{nextClass.period}교시 {nextClass.subject}</em>, {nextClass.start - nowMin}분 뒤에 시작해요</>;
  else if (todaySchedule.length) headline = <>오늘 수업이 모두 끝났어요. 수고했어요!</>;
  else headline = <>오늘은 등록된 수업이 없어요</>;

  const birthdayStudents = students.filter(s => isTodayBirthday(s.birthday));
  // 20일 안의 학생 생일·구글 달력 일정을 자동으로 함께 보여줌 (열 때마다 계산, 직접 넣은 D-Day와 섞어 날짜순)
  const AUTO_DDAY_LIMIT = 20;
  const [calSoon, setCalSoon] = useState([]);
  const [calProblem, setCalProblem] = useState(''); // 달력 일정을 못 가져온 이유 (조용히 빠지지 않게 위젯에 보여줌)
  const calInputs = [settings?.calendarId1, settings?.calendarId2, settings?.calendarId3].filter(Boolean);
  const calKey = (settings?.googleApiKey || '').trim();
  useEffect(() => {
    if (!calInputs.length) { setCalSoon([]); setCalProblem(''); return; }
    if (!calKey) { setCalSoon([]); setCalProblem('구글 달력 일정은 ⚙️ 설정에 Google API 키를 넣어야 함께 나와요.'); return; }
    let alive = true;
    const from = new Date(); from.setHours(0, 0, 0, 0);
    const to = new Date(from); to.setDate(to.getDate() + AUTO_DDAY_LIMIT + 1);
    fetchCalendarEvents(calInputs, calKey, from, to)
      .then(evs => { if (alive) { setCalSoon(evs); setCalProblem(''); } })
      .catch(err => { if (alive) { setCalSoon([]); setCalProblem(`구글 달력을 불러오지 못했어요 (${err.message}). API 키와 달력 ‘공개’ 설정을 확인해 주세요.`); } });
    return () => { alive = false; };
  }, [calKey, calInputs.join('|')]); // eslint-disable-line react-hooks/exhaustive-deps
  const autoDays = useMemo(() => upcomingAutoDays(students, calSoon, new Date(), AUTO_DDAY_LIMIT), [students, calSoon]);
  const ddaySorted = [...ddays, ...autoDays].sort((a, b) => new Date(a.date) - new Date(b.date));
  const firstUpcoming = ddaySorted.find(ev => !getDday(ev.date).startsWith('D+'));

  const handleAddDday = () => {
    if (!newDday.name || !newDday.date) { showToast('이름과 날짜를 입력해주세요.', 'error'); return; }
    addDDay(newDday); setNewDday({ name: '', date: '', emoji: '' }); setShowDdayForm(false); showToast('D-Day가 추가되었습니다.', 'success');
  };
  const handleUpdateThermometer = (goal, reward) => { updateSettings({ ...settings, thermometerGoal: goal, thermometerReward: reward }); showToast('우리반 온도계 설정이 변경되었습니다.', 'success'); };

  const topStudents = [...students].sort((a, b) => (b.points || 0) - (a.points || 0)).slice(0, 5);
  const topScore = Math.max(1, topStudents[0]?.points || 0);
  const totalPoints = students.reduce((sum, s) => sum + (s.points || 0), 0);
  const goalPoints = settings?.thermometerGoal || 1500;
  const lunchItems = lunchInfo ? lunchInfo.menu.split('\n').map(s => s.trim()).filter(Boolean) : [];

  const layout = useMemo(() => normalizeLayout(settings?.dashboardLayouts), [settings?.dashboardLayouts]);
  const saveLayout = (next) => updateSettings({ dashboardLayouts: next.map(({ id, x, y, w, h }) => ({ id, x, y, w, h })) });
  const resetLayout = () => { saveLayout(DEFAULT_DASHBOARD_LAYOUT); showToast('기본 배치로 되돌렸습니다.', 'success'); };

  const shortDate = (iso) => { const d = new Date(iso); return `${d.getMonth() + 1}월 ${d.getDate()}일 ${getDayName(d)}요일`; };
  const noticeTypeLabel = { important: '중요', newsletter: '통신문', general: '일반' };

  const widgets = {
    schedule: {
      label: '오늘의 시간표',
      content: (<>
        <div className="wcard-head"><h2 className="wcard-title">오늘의 시간표</h2><a href="#/weekly" className="wcard-link">주간 일정에서 수정</a></div>
        <div className="wcard-body">
          {todaySchedule.length > 0 ? (
            <ol className="periods">
              {todaySchedule.map(c => {
                const status = nowMin > c.end ? 'done' : (currentClass === c ? 'now' : 'next');
                const progress = status === 'now' ? Math.round(((nowMin - c.start) / (c.end - c.start)) * 100) : status === 'done' ? 100 : 0;
                return (
                  <li key={c.period} data-tone={subjectTone(c.subject)} className={`period is-${status}`} aria-current={status === 'now' ? 'time' : undefined}>
                    <span className="period-time">{c.period}교시 · {minToClock(c.start)}</span>
                    <span className="period-subject">{c.subject}</span>
                    {c.content && <span className="period-content">{c.content}</span>}
                    <span className="period-bar"><span style={{ width: `${progress}%` }} /></span>
                  </li>
                );
              })}
            </ol>
          ) : <div className="wcard-empty">오늘 시간표가 비어 있어요. 주간 일정에서 입력해 주세요.</div>}
        </div>
      </>),
    },
    thermometer: {
      label: '우리 반 칭찬 온도계', className: 'wcard--accent',
      content: <Thermometer current={totalPoints} goal={goalPoints} reward={settings?.thermometerReward || ''} isAdmin={isAdmin} onUpdate={handleUpdateThermometer} />,
    },
    lunch: {
      label: '오늘의 급식',
      content: (<>
        <div className="wcard-head"><h2 className="wcard-title">오늘의 급식</h2>{lunchInfo?.calories && <span className="wcard-meta num">{lunchInfo.calories}</span>}</div>
        <div className="wcard-body">
          {isLoadingLunch ? <div className="wcard-empty">식단표를 불러오는 중이에요…</div>
            : lunchItems.length ? <ul className="lunch-list">{lunchItems.map((m, i) => <li key={i}>{m}</li>)}</ul>
            : <div className="wcard-empty">오늘은 급식 정보가 없어요.</div>}
        </div>
      </>),
    },
    dday: {
      label: '다가오는 날',
      content: (<>
        <div className="wcard-head">
          <h2 className="wcard-title">다가오는 날</h2>
          {isAdmin && <button type="button" className="round-btn round-btn--sm" aria-label={showDdayForm ? 'D-Day 추가 닫기' : 'D-Day 추가'} aria-expanded={showDdayForm} onClick={() => setShowDdayForm(!showDdayForm)}><Icon name="plus" size={16} /></button>}
        </div>
        <div className="wcard-body">
          {showDdayForm && isAdmin && (
            <div className="dday-form">
              <input type="text" aria-label="이름" placeholder="예: 현장체험학습" value={newDday.name} onChange={e => setNewDday({ ...newDday, name: e.target.value })} className="form-input" />
              <input type="date" aria-label="날짜" value={newDday.date} onChange={e => setNewDday({ ...newDday, date: e.target.value })} className="form-input" />
              <button type="button" className="btn btn-primary btn-sm" onClick={handleAddDday}>추가</button>
            </div>
          )}
          {isAdmin && calProblem && <p className="dday-cal-problem">{calProblem}</p>}
          {ddaySorted.length === 0 ? <div className="wcard-empty">등록된 D-Day가 없어요.</div> : (
            <ul className="dday-rows">
              {ddaySorted.map((ev, i) => (
                <li key={ev.id || i} className={`dday-row ${ev === firstUpcoming ? 'is-next' : ''}`}>
                  <span className="dday-num num">{getDday(ev.date)}</span>
                  <span className="dday-text"><span className="dday-title">{ev.emoji ? `${ev.emoji} ` : ''}{ev.name}</span><span className="dday-when">{shortDate(ev.date)}{ev.auto && <span className="dday-auto">{ev.auto}</span>}</span></span>
                  {isAdmin && !ev.auto && <button type="button" className="icon-ghost" aria-label={`${ev.name} 삭제`} onClick={() => deleteDDay(ev.id)}><Icon name="trash" size={16} /></button>}
                </li>
              ))}
            </ul>
          )}
        </div>
      </>),
    },
    ranking: {
      label: '칭찬 랭킹',
      content: (<>
        <div className="wcard-head"><h2 className="wcard-title">칭찬 랭킹</h2><span className="wcard-meta">상위 5명</span></div>
        <div className="wcard-body">
          {topStudents.length === 0 ? <div className="wcard-empty">학생 관리에서 학생을 등록하면 순위가 보여요.</div> : (
            <ol className="rank-list">
              {topStudents.map((s, idx) => (
                <li key={s.id} className={idx === 0 ? 'is-first' : ''}>
                  <span className="rank-no num">{idx + 1}</span>
                  <span className="rank-name">{s.name}</span>
                  <span className="rank-score num">{s.points || 0}</span>
                  <span className="rank-bar"><span style={{ width: `${Math.round(((s.points || 0) / topScore) * 100)}%` }} /></span>
                </li>
              ))}
            </ol>
          )}
        </div>
      </>),
    },
    notices: {
      label: '공지사항',
      content: (<>
        <div className="wcard-head">
          <h2 className="wcard-title">공지사항</h2>
          {isAdmin && <button type="button" className="btn btn-primary btn-sm" onClick={() => setShowNoticeForm(true)}><Icon name="plus" size={16} />새 공지</button>}
        </div>
        <div className="wcard-body">
          {announcements.length === 0 ? <div className="wcard-empty">등록된 공지사항이 없어요.</div> : (
            <div className="notice-grid">
              {announcements.map(n => (
                <article key={n.id} className={`notice-card is-${n.type || 'general'}`}>
                  <div className="notice-card-top">
                    <span className={`notice-tag is-${n.type || 'general'}`}>{noticeTypeLabel[n.type] || '일반'}</span>
                    <span className="notice-date num">{(n.date || '').slice(5).replace('-', '.')}</span>
                    {isAdmin && (<span className="notice-actions">
                      <button type="button" className="icon-ghost" aria-label={`${n.title} 수정`} onClick={() => handleEditNotice(n)}><Icon name="edit" size={16} /></button>
                      <button type="button" className="icon-ghost" aria-label={`${n.title} 삭제`} onClick={() => handleDeleteNotice(n.id)}><Icon name="trash" size={16} /></button>
                    </span>)}
                  </div>
                  <h3 className="notice-card-title">{n.title}</h3>
                  {n.content && <p className="notice-card-body">{n.content}</p>}
                </article>
              ))}
            </div>
          )}
        </div>
      </>),
    },
  };

  return (
    <div className="dash">
      <div className="dash-top">
        <h1 className="dash-headline">{headline}</h1>
        {isAdmin && (
          <div className="dash-tools">
            {editingLayout && <button type="button" className="btn btn-secondary btn-sm" onClick={resetLayout}>기본 배치로</button>}
            <button type="button" className={`btn btn-sm ${editingLayout ? 'btn-primary' : 'btn-secondary'}`} aria-pressed={editingLayout} onClick={() => setEditingLayout(!editingLayout)}>
              {editingLayout ? '배치 완료' : <><Icon name="move" size={16} />배치 편집</>}
            </button>
          </div>
        )}
      </div>
      {editingLayout && <p className="dash-edit-hint">위젯 오른쪽 위 <strong>옮기기</strong>를 끌어 자리를 바꾸고, 오른쪽 아래 모서리를 끌어 크기를 조절하세요. 바꾼 배치는 바로 저장돼요.</p>}

      {birthdayStudents.length > 0 && (
        <div className="dash-banner"><Icon name="cake" size={22} /><span className="dash-banner-main">오늘은 <strong>{birthdayStudents.map(s => s.name).join(', ')}</strong> 학생의 생일이에요</span><span className="dash-banner-sub">다 같이 축하해 주세요</span></div>
      )}
      {holiday?.isHoliday && (
        <div className="dash-banner"><Icon name="sun" size={22} /><span className="dash-banner-main">학사일정상 휴업일: <strong>{holiday.name}</strong></span><span className="dash-banner-sub">출석 체크가 필요 없는 날일 수 있어요</span></div>
      )}

      <DashboardGrid layout={layout} editing={editingLayout} onChange={saveLayout} widgets={widgets} />

      {isAdmin && showNoticeForm && (
        <div className="notice-modal-overlay" onClick={closeNoticeForm}>
          <div className="notice-modal-content" onClick={e => e.stopPropagation()}>
            <div className="notice-modal-header"><h3>{editNoticeId ? '공지사항 수정' : '새 공지사항'}</h3><button type="button" className="notice-close-btn" aria-label="닫기" onClick={closeNoticeForm}>&times;</button></div>
            <div className="notice-modal-body">
              <input className="form-input" aria-label="제목" placeholder="제목을 입력하세요" value={newNotice.title} onChange={e => setNewNotice({ ...newNotice, title: e.target.value })} />
              <textarea className="form-input" aria-label="내용" placeholder="내용을 입력하세요" value={newNotice.content} onChange={e => setNewNotice({ ...newNotice, content: e.target.value })} rows="5" style={{ resize: 'vertical' }} />
              <select className="form-input" aria-label="종류" value={newNotice.type} onChange={e => setNewNotice({ ...newNotice, type: e.target.value })}>
                <option value="general">일반</option><option value="important">중요</option><option value="newsletter">통신문</option>
              </select>
            </div>
            <div className="notice-modal-footer">
              <button type="button" className="btn btn-secondary" onClick={closeNoticeForm}>취소</button>
              <button type="button" className="btn btn-primary" onClick={handleAddNotice}>{editNoticeId ? '수정' : '등록'}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ===================================================================
   StudentsPage
   =================================================================== */
function StudentsPage() {
  const { state, dispatch, showToast, updatePoints, resetPoints, updateStudent, bulkUpdateStudents, addStudentRecord, deleteStudentRecord } = useApp();
  const { students, isAdmin, studentRecords, assignments, settings } = state;
  const [activeTab, setActiveTab] = useState('list');
  useEffect(() => { if (!isAdmin) setActiveTab('list'); }, [isAdmin]); // 학생 화면으로 바꾸면 선생님 전용 탭을 닫음
  const [search, setSearch] = useState('');
  const [view, setView] = useState('card');
  const [pointsView, setPointsView] = useState('individual');
  const [selectedStudent, setSelectedStudent] = useState(null);
  const [isEditingName, setIsEditingName] = useState(false);
  const [editNameValue, setEditNameValue] = useState('');
  const [showBulkModal, setShowBulkModal] = useState(false);
  const [bulkNames, setBulkNames] = useState('');
  const [generatingAiId, setGeneratingAiId] = useState(null);
  const [draggedSeat, setDraggedSeat] = useState(null);
  const [draggedGroupStudent, setDraggedGroupStudent] = useState(null);
  const [editGender, setEditGender] = useState('');
  const [editBirthday, setEditBirthday] = useState('');
  const [editParentPhone, setEditParentPhone] = useState('');
  const [newRecordType, setNewRecordType] = useState('observation');
  const [newRecordContent, setNewRecordContent] = useState('');
  const [newRecordCategory, setNewRecordCategory] = useState('생활');

  const filtered = students.filter(s => s.name.includes(search) || String(s.number).includes(search));
  const sorted = [...students].sort((a, b) => b.points - a.points);

  const handleAddPoints = (id, currentPoints, amount) => { updatePoints(id, currentPoints, amount); showToast(`${amount > 0 ? '+' : ''}${amount}점이 부여되었습니다.`, 'success'); };

  const groupPoints = [1, 2, 3, 4, 5, 6].map(groupId => {
    const groupStudents = students.filter(s => s.group === groupId);
    const totalPoints = groupStudents.reduce((sum, s) => sum + (s.points || 0), 0);
    return { groupId, totalPoints, students: groupStudents };
  }).sort((a, b) => b.totalPoints - a.totalPoints);

  const seatingList = [...students].sort((a, b) => {
    const aIdx = a.seatIndex !== undefined ? a.seatIndex : a.number;
    const bIdx = b.seatIndex !== undefined ? b.seatIndex : b.number;
    return aIdx - bIdx;
  });

  const handleDragStartSeat = (e, student) => { if (!isAdmin) return; setDraggedSeat(student); e.dataTransfer.effectAllowed = 'move'; e.target.style.opacity = '0.5'; };
  const handleDragEndSeat = (e) => { if (!isAdmin) return; e.target.style.opacity = '1'; setDraggedSeat(null); };
  const handleDragOverSeat = (e) => { if (!isAdmin) return; e.preventDefault(); e.dataTransfer.dropEffect = 'move'; };
  const handleDropSeat = (e, targetStudent) => {
    if (!isAdmin) return; e.preventDefault();
    if (!draggedSeat || draggedSeat.id === targetStudent.id) return;
    const dragIndex = draggedSeat.seatIndex !== undefined ? draggedSeat.seatIndex : draggedSeat.number;
    const targetIndex = targetStudent.seatIndex !== undefined ? targetStudent.seatIndex : targetStudent.number;
    bulkUpdateStudents([{ id: draggedSeat.id, seatIndex: targetIndex }, { id: targetStudent.id, seatIndex: dragIndex }]);
  };

  const handleDragStartTeam = (e, student) => { if (!isAdmin) return; setDraggedGroupStudent(student); e.dataTransfer.effectAllowed = 'move'; e.target.style.opacity = '0.5'; };
  const handleDragEndTeam = (e) => { if (!isAdmin) return; e.target.style.opacity = '1'; setDraggedGroupStudent(null); };
  const handleDragOverTeam = (e) => { if (!isAdmin) return; e.preventDefault(); e.currentTarget.classList.add('is-over'); };
  const handleDragLeaveTeam = (e) => { if (!isAdmin) return; e.currentTarget.classList.remove('is-over'); };
  const handleDropTeam = (e, targetGroupId) => {
    if (!isAdmin) return; e.preventDefault();
    e.currentTarget.classList.remove('is-over');
    if (!draggedGroupStudent || draggedGroupStudent.group === targetGroupId) return;
    updateStudent({ id: draggedGroupStudent.id, group: targetGroupId });
  };

  const handleSelectStudent = (s) => { setSelectedStudent(s); setEditGender(s.gender || 'M'); setEditBirthday(s.birthday || ''); setEditParentPhone(s.parentPhone || ''); setIsEditingName(false); setNewRecordContent(''); };
  const handleSaveStudentInfo = () => {
    updateStudent({ id: selectedStudent.id, gender: editGender, birthday: editBirthday, parentPhone: editParentPhone });
    setSelectedStudent({ ...selectedStudent, gender: editGender, birthday: editBirthday, parentPhone: editParentPhone });
    showToast('학생 정보가 수정되었습니다.', 'success');
  };
  const handleResetPoints = () => {
    if (!window.confirm(`${selectedStudent.name} 학생의 포인트를 초기화하시겠습니까?`)) return;
    resetPoints(selectedStudent.id); setSelectedStudent({ ...selectedStudent, points: 0 }); showToast('포인트가 초기화되었습니다.', 'success');
  };
  const handleAddRecord = () => {
    if (!newRecordContent.trim()) { showToast('내용을 입력해주세요.', 'error'); return; }
    addStudentRecord(selectedStudent.id, newRecordType, newRecordContent, newRecordCategory); setNewRecordContent('');
  };

  const handleBulkUpdate = () => {
    const lines = bulkNames.split('\n').map(l => l.trim()).filter(l => l);
    if (lines.length === 0) { showToast('입력된 내용이 없습니다.', 'error'); return; }
    const newStudents = lines.map((line, i) => {
      const parts = line.split(/[\t,]+/).map(p => p.trim());
      const existing = students.find(s => s.number === i + 1);
      let gender = 'M';
      if (parts[1]) { const g = parts[1].toUpperCase(); if (g === 'F' || g === '여') gender = 'F'; }
      return { ...existing, id: existing ? existing.id : generateUUID(), number: i + 1, name: parts[0] || `학생${i + 1}`, gender, birthday: normalizeBirthday(parts[2]) || (existing ? existing.birthday : ''), points: existing ? existing.points : 0, group: existing ? existing.group : Math.ceil((i + 1) / 4) };
    });
    dispatch({ type: 'SET_STUDENTS', payload: newStudents });
    bulkUpdateStudents(newStudents);
    showToast(`총 ${newStudents.length}명으로 명단이 갱신되었습니다.`, 'success');
    setShowBulkModal(false); setBulkNames('');
  };

  const handleExcelUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const data = await parseStudentExcel(file);
      if (data.length === 0) { showToast('데이터가 없습니다.', 'error'); return; }
      const newStudents = data.map((row, i) => {
        const existing = students.find(s => s.number === i + 1);
        let gender = 'M';
        if (row['성별']) { const g = String(row['성별']).toUpperCase(); if (g === 'F' || g === '여') gender = 'F'; }
        return { ...existing, id: existing ? existing.id : generateUUID(), number: i + 1, name: row['이름'] || `학생${i + 1}`, gender, birthday: normalizeBirthday(row['생일']) || (existing ? existing.birthday : ''), parentPhone: row['연락처'] || (existing ? existing.parentPhone : ''), points: existing ? existing.points : 0, group: existing ? existing.group : Math.ceil((i + 1) / 4) };
      });
      dispatch({ type: 'SET_STUDENTS', payload: newStudents });
      bulkUpdateStudents(newStudents);
      showToast(`총 ${newStudents.length}명으로 명단이 갱신되었습니다.`, 'success');
    } catch (err) { showToast('엑셀 파일을 읽는 중 오류가 발생했습니다.', 'error'); }
    e.target.value = null;
  };

  const handleGenerateAI = async (student) => {
    const recs = studentRecords[student.id] || [];
    const evals = assignments.filter(a => a.submissions?.[student.id]).map(a => ({ title: a.title, grade: a.submissions[student.id].grade }));
    const studentData = { name: student.name, gender: student.gender, points: student.points || 0, records: recs, evaluations: evals };
    try {
      setGeneratingAiId(student.id);
      const summary = await generateBehaviorSummary(studentData, settings.geminiApiKey);
      updateStudent({ ...student, aiSummary: summary });
      showToast('AI가 종합의견을 생성했습니다.', 'success');
    } catch (err) { showToast('AI 생성 중 오류가 발생했습니다: ' + err.message, 'error'); }
    finally { setGeneratingAiId(null); }
  };

  const handleEvaluate = (assignmentId, studentId, grade) => { dispatch({ type: 'EVALUATE_ASSIGNMENT', payload: { id: assignmentId, studentId, grade } }); showToast('평가가 저장되었습니다.', 'success'); };

  const records = selectedStudent ? (studentRecords[selectedStudent.id] || []) : [];

  const TABS = [{ k: 'list', label: '학생 목록' }, { k: 'behavior', label: '행동발달', admin: true }, { k: 'evaluation', label: '학습평가', admin: true }];
  const VIEWS = [{ v: 'card', label: '카드' }, { v: 'table', label: '표' }, { v: 'seating', label: '자리' }, { v: 'points', label: '칭찬' }];
  const fmtBirthday = (b) => { const m = /^\d{4}-(\d{2})-(\d{2})$/.exec(b || ''); return m ? `${+m[1]}월 ${+m[2]}일` : ''; };
  const rankClass = (i) => (i === 0 ? 'is-first' : '');

  return (
    <div className="stu">
      <div className="wk-top">
        <div className="wk-title">
          <h1 className="dash-headline">우리 반 <em>{students.length}명</em></h1>
        </div>
        <div className="mode-switch" role="group" aria-label="학생 관리 메뉴">
          {TABS.map(t => {
            const locked = t.admin && !isAdmin;
            return <button key={t.k} type="button" aria-pressed={activeTab === t.k} aria-disabled={locked || undefined} className={locked ? 'is-locked' : ''} onClick={() => locked ? showToast('선생님 모드에서 볼 수 있어요.', 'error') : setActiveTab(t.k)}>{t.label}{locked && <Icon name="lock" size={14} />}</button>;
          })}
        </div>
      </div>

      {activeTab === 'list' && (
        <>
          <div className="stu-bar">
            <label className="stu-search">
              <Icon name="search" size={18} />
              <span className="sr-only">학생 검색</span>
              <input type="search" placeholder="이름이나 번호로 찾기" value={search} onChange={e => setSearch(e.target.value)} />
            </label>
            <div className="mode-switch" role="group" aria-label="보기 방식">
              {VIEWS.map(b => <button key={b.v} type="button" aria-pressed={view === b.v} onClick={() => setView(b.v)}>{b.label}</button>)}
            </div>
            {isAdmin && (
              <div className="stu-admin">
                <button type="button" className="btn btn-secondary btn-sm" onClick={downloadStudentTemplate}>양식 받기</button>
                <label className="btn btn-secondary btn-sm stu-upload">엑셀 올리기<input type="file" accept=".xlsx, .xls" className="sr-only" onChange={handleExcelUpload} /></label>
                <button type="button" className="btn btn-secondary btn-sm" onClick={() => { const existingNames = students.map(s => `${s.name}\t${s.gender || ''}\t${s.birthday || ''}`).join('\n'); setBulkNames(existingNames); setShowBulkModal(true); }}>명단 한 번에 고치기</button>
              </div>
            )}
          </div>

          {view === 'card' && (
            filtered.length ? (
              <ul className="stu-grid">
                {filtered.map(s => (
                  <li key={s.id} className="stu-card" data-tone={genderTone(s.gender)}>
                    <button type="button" className="stu-card-main" onClick={() => handleSelectStudent(s)}>
                      <span className="stu-card-top"><span className="att-no num">{s.number}</span><span className="stu-name">{s.name}</span></span>
                      <span className="stu-meta">{[s.gender === 'F' ? '여' : '남', fmtBirthday(s.birthday)].filter(Boolean).join(' · ')}</span>
                      {(s.allergies || s.healthNotes) && (
                        <span className="stu-tags">
                          {s.allergies && <span className="stu-tag is-warn">알레르기 {s.allergies}</span>}
                          {s.healthNotes && <span className="stu-tag">{s.healthNotes}</span>}
                        </span>
                      )}
                    </button>
                    <div className="stu-card-foot">
                      <span className="stu-points"><strong className="num">{s.points || 0}</strong>점</span>
                      {isAdmin && (
                        <span className="stu-steppers">
                          <button type="button" className="icon-ghost" aria-label={`${s.name} 1점 빼기`} onClick={() => handleAddPoints(s.id, s.points, -1)}><Icon name="minus" size={16} /></button>
                          <button type="button" className="icon-ghost is-plus" aria-label={`${s.name} 1점 더하기`} onClick={() => handleAddPoints(s.id, s.points, 1)}><Icon name="plus" size={16} /></button>
                        </span>
                      )}
                    </div>
                  </li>
                ))}
              </ul>
            ) : <div className="wcard wcard-empty stu-empty">{search ? `“${search}”에 맞는 학생이 없어요.` : '아직 학생이 없어요. 엑셀로 명단을 올려 주세요.'}</div>
          )}

          {view === 'table' && (
            <div className="wk-card">
              <table className="stu-table">
                <thead><tr><th className="stu-col-no">번호</th><th>이름</th><th>성별</th><th>생일</th><th>알레르기</th><th>건강</th><th className="stu-col-num">포인트</th></tr></thead>
                <tbody>
                  {filtered.map(s => (
                    <tr key={s.id} tabIndex={0} onClick={() => handleSelectStudent(s)} onKeyDown={e => e.key === 'Enter' && handleSelectStudent(s)}>
                      <td className="stu-col-no num">{s.number}</td><td className="stu-table-name">{s.name}</td>
                      <td>{s.gender === 'F' ? '여' : '남'}</td><td>{fmtBirthday(s.birthday) || '—'}</td>
                      <td className={s.allergies ? 'stu-warn' : 'stu-dim'}>{s.allergies || '—'}</td><td className={s.healthNotes ? '' : 'stu-dim'}>{s.healthNotes || '—'}</td>
                      <td className="stu-col-num num">{s.points || 0}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {view === 'seating' && (
            <section className="wcard">
              <div className="wcard-head"><h2 className="wcard-title">자리 배치도</h2><span className="seat-legend"><span data-tone="tangerine">남</span><span data-tone="coral">여</span>{isAdmin && <span className="wcard-meta">이름표를 끌어다 놓으면 자리가 바뀌어요</span>}</span></div>
              <div className="wcard-body seating-container">
                <div className="teacher-desk">교탁</div>
                <div className="seating-grid-dynamic">
                  {seatingList.map((s) => (
                    <div key={s.id} className="seat occupied" data-tone={genderTone(s.gender)} draggable={isAdmin} onDragStart={(e) => handleDragStartSeat(e, s)} onDragEnd={handleDragEndSeat} onDragOver={handleDragOverSeat} onDrop={(e) => handleDropSeat(e, s)} onClick={() => handleSelectStudent(s)} style={{ cursor: isAdmin ? 'grab' : 'pointer' }}>
                      <div className="seat-number num">{s.number} · {s.gender === 'F' ? '여' : '남'}</div><div className="seat-name">{s.name}</div>
                    </div>
                  ))}
                </div>
              </div>
            </section>
          )}

          {view === 'points' && (
            <section className="wcard">
              <div className="wcard-head">
                <h2 className="wcard-title">칭찬 포인트 순위</h2>
                <div className="mode-switch" role="group" aria-label="순위 기준">
                  <button type="button" aria-pressed={pointsView === 'individual'} onClick={() => setPointsView('individual')}>개인별</button>
                  <button type="button" aria-pressed={pointsView === 'group'} onClick={() => setPointsView('group')}>팀별</button>
                </div>
              </div>
              <div className="wcard-body">
                {pointsView === 'individual' ? (
                  <ol className="stu-rank">
                    {sorted.map((s, i) => (
                      <li key={s.id} className={rankClass(i)}>
                        <span className="rank-no num">{i + 1}</span>
                        <span className="stu-rank-name">{s.name}<span className="stu-team">{s.group}팀</span></span>
                        <span className="stu-rank-score"><strong className="num">{s.points || 0}</strong>점</span>
                        {isAdmin && (
                          <span className="stu-rank-actions">
                            <button type="button" className="chip-btn" onClick={() => handleAddPoints(s.id, s.points, 5)}>+5</button>
                            <button type="button" className="chip-btn" onClick={() => handleAddPoints(s.id, s.points, 10)}>+10</button>
                            <button type="button" className="chip-btn is-minus" onClick={() => handleAddPoints(s.id, s.points, -5)}>−5</button>
                          </span>
                        )}
                      </li>
                    ))}
                  </ol>
                ) : (
                  <>
                    {isAdmin && <p className="lrn-hint stu-hint">학생 이름표를 다른 팀으로 끌어다 놓으면 팀이 바뀌어요.</p>}
                    <ol className="stu-teams">
                      {groupPoints.map((g, i) => (
                        <li key={g.groupId} className={`stu-team-card ${rankClass(i)}`} onDragOver={handleDragOverTeam} onDragLeave={handleDragLeaveTeam} onDrop={(e) => handleDropTeam(e, g.groupId)}>
                          <div className="stu-team-head">
                            <span className="rank-no num">{i + 1}</span>
                            <span className="stu-team-name">{g.groupId}팀</span>
                            <span className="stu-rank-score"><strong className="num">{g.totalPoints}</strong>점</span>
                          </div>
                          <div className="stu-pills">
                            {g.students.map(s => (<span key={s.id} className="stu-pill" draggable={isAdmin} onDragStart={(e) => handleDragStartTeam(e, s)} onDragEnd={handleDragEndTeam} style={{ cursor: isAdmin ? 'grab' : 'default' }}>{s.name}</span>))}
                            {g.students.length === 0 && <span className="lrn-none">아직 아무도 없어요</span>}
                          </div>
                        </li>
                      ))}
                    </ol>
                  </>
                )}
              </div>
            </section>
          )}
        </>
      )}

      {activeTab === 'behavior' && (
        <ul className="stu-behavior">
          {students.map(s => {
            const recs = studentRecords[s.id] || [];
            return (
              <li key={s.id} className="wcard stu-bh">
                <div className="stu-bh-who">
                  <span className="stu-card-top"><span className="att-no num">{s.number}</span><span className="stu-name">{s.name}</span></span>
                  {isAdmin && (
                    <div className="stu-bh-actions">
                      <button type="button" className="btn btn-secondary btn-sm" onClick={() => handleSelectStudent(s)}><Icon name="plus" size={16} /> 기록 추가</button>
                      <button type="button" className="btn btn-primary btn-sm" onClick={() => handleGenerateAI(s)} disabled={generatingAiId === s.id}>{generatingAiId === s.id ? 'AI가 쓰는 중…' : s.aiSummary ? 'AI 의견 다시 쓰기' : 'AI 의견 쓰기'}</button>
                    </div>
                  )}
                </div>
                <div className="stu-bh-col">
                  <h3 className="stu-bh-title">누적 기록 <span className="num">{recs.length}</span></h3>
                  {recs.length > 0 ? (
                    <ul className="stu-records">
                      {recs.map((r, i) => (
                        <li key={r.id || i}>
                          <div className="stu-record-meta"><span className="num">{r.date}</span>{r.category && <span className="notice-tag">{r.category}</span>}</div>
                          <p className="stu-record-text">{r.content}</p>
                          {isAdmin && <button type="button" className="icon-ghost" aria-label="이 기록 삭제" onClick={() => { if (confirm('이 기록을 삭제하시겠습니까?')) deleteStudentRecord(s.id, r.id); }}><Icon name="trash" size={16} /></button>}
                        </li>
                      ))}
                    </ul>
                  ) : <p className="stu-bh-empty">아직 기록이 없어요.</p>}
                </div>
                <div className="stu-bh-col">
                  <h3 className="stu-bh-title">종합의견</h3>
                  {s.aiSummary ? <p className="stu-summary">{s.aiSummary}</p> : <p className="stu-bh-empty">기록을 모은 뒤 “AI 의견 쓰기”를 누르면 초안이 생겨요.</p>}
                </div>
              </li>
            );
          })}
        </ul>
      )}

      {activeTab === 'evaluation' && (
        assignments.length === 0 ? (<div className="wcard wcard-empty stu-empty">등록된 과제·수행평가가 없어요. 학습 관리에서 활동을 추가해 주세요.</div>) : (
          <div className="wk-card">
            <table className="stu-eval">
              <thead><tr>
                <th className="stu-eval-who">학생</th>
                {assignments.map(a => (
                  <th key={a.id}>
                    <div className="stu-eval-head">
                      <span className={`lrn-type ${a.type === '수행평가' ? 'is-eval' : ''}`}>{a.type}</span>
                      {isAdmin && <button type="button" className="icon-ghost" aria-label={`${a.title} 삭제`} onClick={() => { if (confirm('이 항목을 정말 삭제하시겠습니까?')) { dispatch({ type: 'DELETE_ASSIGNMENT', payload: a.id }); showToast('삭제되었습니다.', 'success'); } }}><Icon name="trash" size={16} /></button>}
                    </div>
                    <div className="stu-eval-title">{a.title}</div>
                    <div className="stu-eval-date num">{a.dueDate}</div>
                  </th>
                ))}
              </tr></thead>
              <tbody>
                {students.map(s => (
                  <tr key={s.id}>
                    <th scope="row" className="stu-eval-who"><span className="att-no num">{s.number}</span>{s.name}</th>
                    {assignments.map(a => {
                      const submission = a.submissions?.[s.id];
                      const grade = submission?.grade;
                      return (
                        <td key={`${s.id}-${a.id}`} className={submission ? '' : 'is-missing'}>
                          {submission ? (isAdmin ? (
                            <div className="stu-grade" role="group" aria-label={`${s.name} ${a.title} 평가`}>
                              {['상', '중', '하'].map(g => (<button key={g} type="button" aria-pressed={grade === g} onClick={() => handleEvaluate(a.id, s.id, g)}>{g}</button>))}
                            </div>
                          ) : (<span className={`stu-grade-chip ${grade ? '' : 'is-wait'}`}>{grade || '채점 전'}</span>)
                          ) : (<span className="stu-missing">미제출</span>)}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )
      )}

      {selectedStudent && (
        <div className="modal-overlay" onClick={() => { setSelectedStudent(null); setIsEditingName(false); }}>
          <div className="modal" onClick={e => e.stopPropagation()}>
            <div className="modal-header"><h3 className="modal-title">학생 정보</h3><button className="modal-close" onClick={() => { setSelectedStudent(null); setIsEditingName(false); }}>&times;</button></div>
            <div className="modal-body" style={{ maxHeight: '70vh', overflowY: 'auto' }}>
              <div className="student-detail-header">
                <div className={`avatar xl ${getAvatarColor(selectedStudent.gender)}`}>{getInitials(selectedStudent.name)}</div>
                <div className="student-detail-info">
                  <div className="student-name-edit-wrapper">
                    {isEditingName ? (
                      <div className="edit-name-form">
                        <input type="text" className="form-input" value={editNameValue} onChange={(e) => setEditNameValue(e.target.value)} autoFocus style={{ width: '120px', padding: '0.25rem 0.5rem', minHeight: '30px' }} />
                        <button className="btn btn-primary btn-sm" onClick={() => { if (editNameValue.trim()) { updateStudent({ id: selectedStudent.id, name: editNameValue.trim() }); setSelectedStudent({ ...selectedStudent, name: editNameValue.trim() }); showToast('이름이 수정되었습니다.', 'success'); } setIsEditingName(false); }}>저장</button>
                        <button className="btn btn-secondary btn-sm" onClick={() => setIsEditingName(false)}>취소</button>
                      </div>
                    ) : (
                      <div className="name-display" style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                        <h3>{selectedStudent.name}</h3>
                        {isAdmin && <button type="button" onClick={() => { setEditNameValue(selectedStudent.name); setIsEditingName(true); }} aria-label="이름 고치기" className="icon-ghost"><Icon name="edit" size={16} /></button>}
                      </div>
                    )}
                  </div>
                  <p>{selectedStudent.number}번 · {selectedStudent.gender === 'M' ? '남학생' : '여학생'} · 생일: {selectedStudent.birthday}</p>
                </div>
              </div>
              {isAdmin && (
                <div className="student-edit-section">
                  <h4 className="stu-modal-h">정보 수정</h4>
                  <div className="edit-fields-grid">
                    <div className="edit-field"><label>성별</label><select className="form-input" value={editGender} onChange={e => setEditGender(e.target.value)}><option value="M">남</option><option value="F">여</option></select></div>
                    <div className="edit-field"><label>생일</label><input type="date" className="form-input" value={editBirthday} onChange={e => setEditBirthday(e.target.value)} /></div>
                    <div className="edit-field"><label>연락처</label><input type="text" className="form-input" value={editParentPhone} onChange={e => setEditParentPhone(e.target.value)} placeholder="010-0000-0000" /></div>
                    <div className="edit-field" style={{ display: 'flex', gap: '0.5rem', alignItems: 'flex-end' }}><button className="btn btn-primary btn-sm" onClick={handleSaveStudentInfo}>저장</button><button className="btn btn-danger btn-sm" onClick={handleResetPoints}>포인트 초기화</button></div>
                  </div>
                </div>
              )}
              <p className="stu-modal-line"><span>학부모 연락처</span><strong className="num">{selectedStudent.parentPhone || '미등록'}</strong></p>
              {isAdmin && (
                <div className="record-input-section">
                  <h4 className="stu-modal-h">관찰·상담 기록 추가</h4>
                  <div className="record-input-row">
                    <select className="form-input" value={newRecordType} onChange={e => setNewRecordType(e.target.value)} style={{ width: 'auto', minWidth: '100px' }}><option value="observation">관찰</option><option value="counseling">상담</option></select>
                    <select className="form-input" value={newRecordCategory} onChange={e => setNewRecordCategory(e.target.value)} style={{ width: 'auto', minWidth: '80px' }}>{['생활', '학습', '교우', '건강', '기타'].map(c => <option key={c} value={c}>{c}</option>)}</select>
                  </div>
                  <div className="record-input-row" style={{ marginTop: '0.5rem' }}>
                    <textarea className="form-input" placeholder="관찰/상담 내용을 입력하세요..." value={newRecordContent} onChange={e => setNewRecordContent(e.target.value)} rows="4" style={{ resize: 'vertical', flex: 1 }} />
                    <button className="btn btn-primary" onClick={handleAddRecord} style={{ alignSelf: 'flex-end' }}>저장</button>
                  </div>
                </div>
              )}
              {records.length > 0 && (
                <div className="records-list-section">
                  <h4 className="stu-modal-h">기록 <span className="num">{records.length}</span>건</h4>
                  <div className="counseling-list" style={{ marginTop: '0.5rem' }}>
                    {records.map((c, i) => (
                      <div key={c.id || i} className="counseling-item">
                        <div className="counseling-date">{c.date} · <span className={`record-type-badge ${c.type}`}>{c.type === 'observation' ? '관찰' : c.type === 'counseling' ? '상담' : c.type}</span> · {c.category}
                          {isAdmin && <button type="button" className="icon-ghost" aria-label="이 기록 삭제" onClick={() => { if (confirm('이 기록을 삭제하시겠습니까?')) deleteStudentRecord(selectedStudent.id, c.id); }}><Icon name="trash" size={16} /></button>}
                        </div>
                        <div className="counseling-content">{c.content}</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {showBulkModal && (
        <div className="modal-overlay" onClick={() => setShowBulkModal(false)}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ maxWidth: '450px' }}>
            <div className="modal-header"><h3 className="modal-title">명단 한 번에 고치기</h3><button className="modal-close" onClick={() => setShowBulkModal(false)}>&times;</button></div>
            <div className="modal-body">
              <p style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', marginBottom: '0.75rem' }}>한 줄에 한 명씩 탭(Tab)이나 쉼표(,)로 구분하여 <strong>이름, 성별(M/F), 생일(YYYY-MM-DD)</strong>을 입력하세요.<br />(예: <code>김학생, M, 2014-05-01</code>)</p>
              <textarea className="form-input" value={bulkNames} onChange={e => setBulkNames(e.target.value)} rows="15" style={{ resize: 'vertical', fontFamily: 'monospace', lineHeight: '1.8' }} placeholder={'이름\t성별\t생일'} />
              <div style={{ display: 'flex', gap: '0.5rem', marginTop: '0.75rem', justifyContent: 'flex-end' }}><button className="btn btn-secondary" onClick={() => setShowBulkModal(false)}>취소</button><button className="btn btn-primary" onClick={handleBulkUpdate}>저장</button></div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ===================================================================
   AttendancePage
   =================================================================== */
const STATUS_MAP = { present: { label: '출석', mark: '○' }, absent: { label: '결석', mark: '✕' }, late: { label: '지각', mark: '△' }, 'early-leave': { label: '조퇴', mark: '▽' } };
const STATUS_KEYS = Object.keys(STATUS_MAP);
const sameAtt = (a, b) => (a?.status || 'present') === (b?.status || 'present') && (a?.note || '') === (b?.note || '') && (a?.reason || '') === (b?.reason || '');
// NEIS 출결 구분 — 결석·지각·조퇴마다 이 넷 중 하나로 입력함
const ATT_REASONS = ['질병', '미인정', '기타', '출석인정'];
const NEIS_TYPES = ['absent', 'late', 'early-leave'];

function AttendancePage() {
  const { state, loadAttendanceByDate, loadMonthlyAttendance, saveAttendanceDay, showToast } = useApp();
  const { students, attendance, selectedAttendanceDate, monthlyAttendance, isAdmin } = state;
  const [monthlyActiveCell, setMonthlyActiveCell] = useState(null);
  const [viewMode, setViewMode] = useState('daily');
  const [monthYear, setMonthYear] = useState({ year: new Date().getFullYear(), month: new Date().getMonth() });
  // 입력 중인 출석(저장 전). 완료 버튼을 눌러야 db에 반영됨
  const [draft, setDraft] = useState({});
  const [selected, setSelected] = useState(() => new Set());
  const [noteFor, setNoteFor] = useState(null);
  const [noteText, setNoteText] = useState('');
  const [noteReason, setNoteReason] = useState('');
  // 월별 현황에서 고친 칸(저장 전): { 'YYYY-MM-DD': { 학생id: 상태 } }
  const [monthDraft, setMonthDraft] = useState({});
  const monthChanges = Object.entries(monthDraft).flatMap(([d, recs]) => Object.entries(recs).filter(([id, st]) => (monthlyAttendance[d]?.[id]?.status || null) !== st).map(([id, st]) => ({ date: d, id, status: st })));
  const monthDirty = monthChanges.length > 0;

  const baseOf = (id) => attendance[id] || { status: 'present', note: '' };
  const current = (id) => draft[id] || baseOf(id);
  const changedIds = students.filter(s => draft[s.id] && !sameAtt(draft[s.id], baseOf(s.id))).map(s => s.id);
  const dirty = changedIds.length > 0;

  useEffect(() => { setDraft({}); setSelected(new Set()); setNoteFor(null); }, [selectedAttendanceDate]);
  useEffect(() => {
    if (!dirty && !monthDirty) return;
    const warn = (e) => { e.preventDefault(); e.returnValue = ''; };
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, [dirty, monthDirty]);

  // 드롭다운 밖을 누르거나 Esc 를 누르면 닫힘
  useEffect(() => {
    if (!monthlyActiveCell) return;
    const onDown = (e) => { if (!e.target.closest('.monthly-cell.active')) setMonthlyActiveCell(null); };
    const onKey = (e) => { if (e.key === 'Escape') setMonthlyActiveCell(null); };
    document.addEventListener('pointerdown', onDown);
    document.addEventListener('keydown', onKey);
    return () => { document.removeEventListener('pointerdown', onDown); document.removeEventListener('keydown', onKey); };
  }, [monthlyActiveCell]);

  const handleDateChange = (e) => {
    if (dirty && !confirm('저장하지 않은 출석 변경이 있어요. 날짜를 바꾸면 사라집니다. 계속할까요?')) return;
    loadAttendanceByDate(e.target.value);
  };
  // 칸을 고르면 표시만 바뀌고, 아래 '저장'을 눌러야 db에 들어감
  const handleMonthlyStatus = (studentId, date, status) => {
    if (!isAdmin) { showToast('선생님 모드에서만 수정할 수 있습니다.', 'error'); return; }
    setMonthDraft(prev => ({ ...prev, [date]: { ...(prev[date] || {}), [studentId]: status } }));
    setMonthlyActiveCell(null);
  };
  const saveMonth = () => {
    const byDate = {};
    monthChanges.forEach(({ date, id, status }) => {
      const old = monthlyAttendance[date]?.[id] || {};
      const absent = status !== 'present';
      (byDate[date] = byDate[date] || {})[id] = { ...old, status, note: absent ? (old.note || '') : '', reason: absent ? (old.reason || '') : '' };
    });
    Object.entries(byDate).forEach(([date, records]) => saveAttendanceDay(date, records));
    showToast(`출석 ${monthChanges.length}칸을 저장했습니다.`, 'success');
    setMonthDraft({});
  };
  const confirmLeaveMonth = () => !monthDirty || confirm('저장하지 않은 월별 출석 변경이 있어요. 이동하면 사라집니다. 계속할까요?');
  const changeMonth = (delta) => {
    if (!confirmLeaveMonth()) return;
    setMonthDraft({}); setMonthlyActiveCell(null);
    setMonthYear(({ year, month }) => { const m = month + delta; return { year: year + Math.floor(m / 12), month: ((m % 12) + 12) % 12 }; });
  };

  useEffect(() => { if (viewMode === 'monthly') loadMonthlyAttendance(monthYear.year, monthYear.month); }, [viewMode, monthYear, loadMonthlyAttendance]);

  const counts = Object.fromEntries(STATUS_KEYS.map(k => [k, students.filter(s => current(s.id).status === k).length]));

  const toggleSelect = (id) => setSelected(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const allSelected = students.length > 0 && selected.size === students.length;
  const toggleAll = () => setSelected(allSelected ? new Set() : new Set(students.map(s => s.id)));

  const applyStatus = (status) => {
    if (!selected.size) return;
    setDraft(prev => {
      const next = { ...prev };
      selected.forEach(id => { const cur = next[id] || baseOf(id); next[id] = status === 'present' ? { status, note: '', reason: '' } : { status, note: cur.note || '', reason: cur.reason || '' }; });
      return next;
    });
    setSelected(new Set());
  };
  const openNote = (id) => { setNoteFor(id); setNoteText(current(id).note || ''); setNoteReason(current(id).reason || ''); };
  const saveNote = () => { setDraft(prev => ({ ...prev, [noteFor]: { ...current(noteFor), note: noteText.trim(), reason: noteReason } })); setNoteFor(null); };

  const discard = () => { setDraft({}); setSelected(new Set()); setNoteFor(null); };
  const handleSubmit = () => {
    if (!isAdmin) return;
    const records = {};
    students.forEach(s => { const a = current(s.id); const absent = a.status !== 'present'; records[s.id] = { status: a.status, note: absent ? (a.note || '') : '', reason: absent ? (a.reason || '') : '', id: baseOf(s.id).id }; });
    saveAttendanceDay(selectedAttendanceDate, records);
    setDraft({}); setSelected(new Set()); setNoteFor(null);
    showToast(`${selectedAttendanceDate.slice(5).replace('-', '/')} 출석을 저장했습니다.`, 'success');
  };

  const getDaysInMonth = (year, month) => {
    const days = [];
    const lastDay = new Date(year, month + 1, 0).getDate();
    for (let d = 1; d <= lastDay; d++) {
      const dow = new Date(year, month, d).getDay();
      if (dow !== 0 && dow !== 6) days.push(`${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`);
    }
    return days;
  };
  const monthDays = getDaysInMonth(monthYear.year, monthYear.month);
  const selDate = new Date(selectedAttendanceDate);
  const isSavedDay = !!monthlyAttendance[selectedAttendanceDate];
  const noteStudent = students.find(s => s.id === noteFor);

  return (
    <div className="att">
      <div className="wk-top">
        <div className="wk-title">
          <h1 className="dash-headline">{viewMode === 'daily' ? `${selDate.getMonth() + 1}월 ${selDate.getDate()}일 ${getDayName(selDate)}요일` : `${monthYear.year}년 ${monthYear.month + 1}월`}</h1>
          {viewMode === 'daily' && <span className={`att-saved ${isSavedDay && !dirty ? 'is-saved' : ''}`}>{dirty ? `저장 안 됨 · ${changedIds.length}명 변경` : isSavedDay ? '저장됨' : '아직 입력 전'}</span>}
        </div>
        <div className="mode-switch" role="group" aria-label="보기">
          <button type="button" aria-pressed={viewMode === 'daily'} onClick={() => { if (viewMode === 'daily' || !confirmLeaveMonth()) return; setMonthDraft({}); setMonthlyActiveCell(null); setViewMode('daily'); }}>일별 출석</button>
          <button type="button" aria-pressed={viewMode === 'monthly'} onClick={() => setViewMode('monthly')}>월별 현황</button>
        </div>
      </div>

      {viewMode === 'daily' ? (
        <>
          <div className="att-bar">
            <label className="att-date"><Icon name="calendar" size={18} /><span className="sr-only">날짜</span><input type="date" value={selectedAttendanceDate} onChange={handleDateChange} /></label>
            <div className="att-counts">
              {STATUS_KEYS.map(k => <span key={k} className={`att-count is-${k}`}><span className="att-dot" />{STATUS_MAP[k].label}<strong className="num">{counts[k]}</strong></span>)}
            </div>
          </div>

          {isAdmin && (
            <div className="att-tools">
              <button type="button" className={`att-check-all ${allSelected ? 'is-on' : ''}`} aria-pressed={allSelected} onClick={toggleAll} disabled={!students.length}>
                <span className="att-box" aria-hidden="true" />{allSelected ? '선택 해제' : '모두 선택'}
              </button>
              <span className="att-tools-label">{selected.size ? <><strong className="num">{selected.size}</strong>명을</> : '학생을 고른 뒤'}</span>
              <div className="att-apply">
                {STATUS_KEYS.map(k => <button key={k} type="button" className={`att-apply-btn is-${k}`} disabled={!selected.size} onClick={() => applyStatus(k)}>{STATUS_MAP[k].label}</button>)}
              </div>
            </div>
          )}

          {students.length === 0 ? <div className="wcard-empty att-empty">학생 관리에서 학생을 등록하면 출석부가 만들어져요.</div> : (
            <ul className="att-grid">
              {students.map(s => {
                const a = current(s.id);
                const isSel = selected.has(s.id);
                const changed = changedIds.includes(s.id);
                const needsNote = a.status !== 'present';
                return (
                  <li key={s.id} className={`att-card is-${a.status} ${isSel ? 'is-selected' : ''}`}>
                    <button type="button" className="att-card-main" aria-pressed={isSel} disabled={!isAdmin} onClick={() => toggleSelect(s.id)}>
                      {isAdmin && <span className="att-box" aria-hidden="true" />}
                      <span className="att-no num">{s.number}</span>
                      <span className="att-name">{s.name}</span>
                      <span className="att-status">{STATUS_MAP[a.status].label}{changed && <span className="att-changed" aria-label="변경됨" />}</span>
                    </button>
                    {needsNote && (
                      isAdmin
                        ? <button type="button" className={`att-note-btn ${a.note || a.reason ? 'has-note' : ''}`} onClick={() => openNote(s.id)}><Icon name="edit" size={14} />{a.reason ? <span className="att-reason">{a.reason}</span> : <span className="att-reason is-missing">구분 선택</span>}{a.note || (a.reason ? '' : '사유 적기')}</button>
                        : (a.note || a.reason) ? <span className="att-note-btn has-note">{a.reason && <span className="att-reason">{a.reason}</span>}{a.note}</span> : null
                    )}
                  </li>
                );
              })}
            </ul>
          )}

          {isAdmin && (
            <div className={`att-submit ${dirty ? 'is-dirty' : ''}`}>
              <span className="att-submit-text">{dirty ? <><strong className="num">{changedIds.length}</strong>명의 출석이 바뀌었어요. 완료를 눌러야 저장돼요.</> : isSavedDay ? '이 날의 출석이 저장되어 있어요.' : '출석 상태를 확인하고 완료를 눌러 주세요.'}</span>
              <div className="att-submit-actions">
                {dirty && <button type="button" className="btn btn-secondary" onClick={discard}>되돌리기</button>}
                <button type="button" className="btn btn-primary" onClick={handleSubmit} disabled={!dirty && isSavedDay}>출석 입력 완료</button>
              </div>
            </div>
          )}

          {noteStudent && (
            <div className="modal-overlay" onClick={() => setNoteFor(null)}>
              <div className="modal-content att-note-modal" role="dialog" aria-modal="true" aria-labelledby="att-note-title" onClick={e => e.stopPropagation()}>
                <div className="modal-header"><h3 id="att-note-title">{noteStudent.name} · {STATUS_MAP[current(noteFor).status].label} 사유</h3><button type="button" className="close-btn" aria-label="닫기" onClick={() => setNoteFor(null)}>×</button></div>
                <div className="modal-body">
                  <div className="att-reason-label" id="att-reason-label">NEIS 구분</div>
                  <div className="mode-switch att-reason-pick" role="group" aria-labelledby="att-reason-label">
                    {ATT_REASONS.map(r => <button key={r} type="button" aria-pressed={noteReason === r} onClick={() => setNoteReason(r)}>{r}</button>)}
                  </div>
                  <label htmlFor="att-note-text" className="att-reason-label">사유 (특기사항)</label>
                  <textarea id="att-note-text" className="form-input" rows="3" autoFocus placeholder="예: 감기로 병원 진료" value={noteText} onChange={e => setNoteText(e.target.value)} />
                  <div className="att-note-presets">{['질병', '병원 진료', '가정 사정', '체험학습', '늦잠'].map(p => <button key={p} type="button" className="chip-btn" onClick={() => setNoteText(p)}>{p}</button>)}</div>
                </div>
                <div className="modal-footer att-note-footer"><button type="button" className="btn btn-secondary" onClick={() => setNoteFor(null)}>취소</button><button type="button" className="btn btn-primary" onClick={saveNote}>저장</button></div>
              </div>
            </div>
          )}
        </>
      ) : (
        <div className="monthly-attendance">
          <div className="wk-nav att-month-nav">
            <button type="button" className="round-btn" aria-label="이전 달" onClick={() => changeMonth(-1)}><Icon name="chevron" size={18} /></button>
            <button type="button" className="round-btn wk-next" aria-label="다음 달" onClick={() => changeMonth(1)}><Icon name="chevron" size={18} /></button>
            <div className="monthly-legend">
              {STATUS_KEYS.map(k => <span key={k} className="legend-item"><span className={`legend-icon ${k}`}>{STATUS_MAP[k].mark}</span>{STATUS_MAP[k].label}</span>)}
              <span className="legend-item"><span className="legend-icon no-data">-</span>기록 없음</span>
            </div>
          </div>
          <div className="monthly-table-wrapper">
            <table className="monthly-table">
              <thead><tr><th className="sticky-col">번호</th><th className="sticky-col-name">이름</th>
                {monthDays.map(d => { const date = new Date(d); return (<th key={d} className="date-col"><div className="date-header"><span className="num">{date.getDate()}</span><span className="date-dow">{getDayName(date)}</span></div></th>); })}
              </tr></thead>
              <tbody>
                {students.map(s => (
                  <tr key={s.id}>
                    <td className="sticky-col num">{s.number}</td><td className="sticky-col-name">{s.name}</td>
                    {monthDays.map(d => {
                      const record = monthlyAttendance[d]?.[s.id];
                      const saved = record?.status || null;
                      const pending = monthDraft[d]?.[s.id];
                      const status = pending || saved;
                      const isChanged = !!pending && pending !== saved;
                      const isCellActive = monthlyActiveCell?.studentId === s.id && monthlyActiveCell?.date === d;
                      return (
                        <td key={d} className={`monthly-cell ${status || 'no-data'} ${isCellActive ? 'active' : ''} ${isChanged ? 'is-changed' : ''}`} title={[record?.reason, record?.note].filter(Boolean).join(' · ')} onClick={() => { if (isAdmin) setMonthlyActiveCell(isCellActive ? null : { studentId: s.id, date: d }); }} style={{ position: 'relative', cursor: isAdmin ? 'pointer' : 'default' }}>
                          <span className={`monthly-status ${status || 'no-data'}`}>{status ? STATUS_MAP[status].mark : '-'}</span>
                          {record?.note && <span className="monthly-note-dot" aria-label={`사유: ${record.note}`} />}
                          {isAdmin && isCellActive && (<div className="attendance-status-selector show monthly-selector" onClick={e => e.stopPropagation()}>{STATUS_KEYS.map(k => (<button key={k} type="button" className={`status-option ${status === k ? 'active' : ''}`} onClick={() => handleMonthlyStatus(s.id, d, k)}>{STATUS_MAP[k].label}</button>))}</div>)}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {isAdmin && (
            <div className={`att-submit ${monthDirty ? 'is-dirty' : ''}`}>
              <span className="att-submit-text">{monthDirty ? <><strong className="num">{monthChanges.length}</strong>칸이 바뀌었어요. 저장을 눌러야 반영돼요.</> : '칸을 눌러 출석을 고친 뒤 저장을 누르세요.'}</span>
              <div className="att-submit-actions">
                {monthDirty && <button type="button" className="btn btn-secondary" onClick={() => setMonthDraft({})}>되돌리기</button>}
                <button type="button" className="btn btn-primary" onClick={saveMonth} disabled={!monthDirty}>저장</button>
              </div>
            </div>
          )}
          <NeisSummary students={students} monthlyAttendance={monthlyAttendance} monthDays={monthDays} year={monthYear.year} month={monthYear.month} />
        </div>
      )}
    </div>
  );
}

/* NEIS 입력용 월별 출결 요약 — NEIS에 자동으로 넣지는 않고, 보면서 옮겨 적기 쉽게 정리 + 엑셀 */
function NeisSummary({ students, monthlyAttendance, monthDays, year, month }) {
  const rows = students.map(s => {
    const counts = {}; const notes = []; let missing = 0;
    NEIS_TYPES.forEach(t => ATT_REASONS.forEach(r => { counts[`${t}:${r}`] = 0; }));
    monthDays.forEach(d => {
      const rec = monthlyAttendance[d]?.[s.id];
      if (!rec || !NEIS_TYPES.includes(rec.status)) return;
      if (rec.reason) counts[`${rec.status}:${rec.reason}`]++; else missing++;
      notes.push(`${+d.slice(5, 7)}/${+d.slice(8)} ${STATUS_MAP[rec.status].label}(${rec.reason || '구분 없음'})${rec.note ? ' ' + rec.note : ''}`);
    });
    return { s, counts, notes, missing };
  }).filter(r => r.notes.length);
  const missingTotal = rows.reduce((n, r) => n + r.missing, 0);
  const title = `${year}년 ${month + 1}월 출결 요약`;

  const downloadExcel = () => {
    const head1 = ['번호', '이름', ...NEIS_TYPES.flatMap(t => ATT_REASONS.map((r, i) => (i === 0 ? STATUS_MAP[t].label : ''))), '특기사항'];
    const head2 = ['', '', ...NEIS_TYPES.flatMap(() => ATT_REASONS), ''];
    const body = rows.map(r => [r.s.number, r.s.name, ...NEIS_TYPES.flatMap(t => ATT_REASONS.map(x => r.counts[`${t}:${x}`] || '')), r.notes.join(', ')]);
    const ws = XLSX.utils.aoa_to_sheet([[title], head1, head2, ...body]);
    ws['!merges'] = NEIS_TYPES.map((t, i) => ({ s: { r: 1, c: 2 + i * 4 }, e: { r: 1, c: 5 + i * 4 } }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '출결요약');
    XLSX.writeFile(wb, `출결요약_${year}-${String(month + 1).padStart(2, '0')}.xlsx`);
  };

  return (
    <section className="neis-sum" aria-labelledby="neis-sum-title">
      <div className="neis-sum-head">
        <div>
          <h2 id="neis-sum-title" className="wcard-title">NEIS 입력용 요약</h2>
          <p className="neis-sum-sub">결석·지각·조퇴를 NEIS 구분(질병·미인정·기타·출석인정)별로 셌어요. NEIS 출결 화면에 보면서 옮겨 적으세요.</p>
        </div>
        <div className="neis-sum-actions">
          <button type="button" className="btn btn-secondary btn-sm" onClick={() => window.print()}>인쇄</button>
          <button type="button" className="btn btn-primary btn-sm" onClick={downloadExcel} disabled={!rows.length}>엑셀 받기</button>
        </div>
      </div>
      {missingTotal > 0 && <p className="neis-sum-warn">NEIS 구분이 비어 있는 기록이 <strong className="num">{missingTotal}</strong>건 있어요. 일별 출석에서 사유를 눌러 구분을 골라 주세요.</p>}
      {rows.length === 0 ? <p className="neis-sum-empty">이번 달에는 결석·지각·조퇴 기록이 없어요.</p> : (
        <div className="neis-sum-scroll">
          <table className="neis-sum-table">
            <thead>
              <tr><th rowSpan="2">번호</th><th rowSpan="2">이름</th>{NEIS_TYPES.map(t => <th key={t} colSpan="4" className={`is-${t}`}>{STATUS_MAP[t].label}</th>)}<th rowSpan="2" className="neis-sum-note">특기사항</th></tr>
              <tr>{NEIS_TYPES.flatMap(t => ATT_REASONS.map(r => <th key={`${t}-${r}`} className="neis-sum-sub-h">{r}</th>))}</tr>
            </thead>
            <tbody>
              {rows.map(r => (
                <tr key={r.s.id}>
                  <td className="num">{r.s.number}</td><td className="neis-sum-name">{r.s.name}</td>
                  {NEIS_TYPES.flatMap(t => ATT_REASONS.map(x => { const n = r.counts[`${t}:${x}`]; return <td key={`${t}-${x}`} className={`num ${n ? 'has-n' : ''}`}>{n || ''}</td>; }))}
                  <td className="neis-sum-note">{r.notes.join(', ')}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}

/* ===================================================================
   LearningPage
   =================================================================== */
function LearningPage() {
  const { state, showToast, addAssignment, updateAssignmentSubmissions, deleteAssignment } = useApp();
  const { students, assignments, isAdmin } = state;
  const currentActivity = [...assignments].filter(a => a.subject === '활동').sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0))[0];
  const [isEditingActivity, setIsEditingActivity] = useState(false);
  const [tempActivityContent, setTempActivityContent] = useState('');
  const [tempActivityType, setTempActivityType] = useState('과제');
  const [selectedId, setSelectedAssignment] = useState(null);
  const selectedAssignment = assignments.find(a => a.id === selectedId) || null; // id로 찾아야 체크 후에도 최신 상태가 보임
  // 아래 목록에서 과제를 선택하면 체크 그리드도 그 과제 기준으로 바뀌고, 선택을 풀면 최신 활동으로 돌아감
  const activeActivity = selectedAssignment || currentActivity;
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 5;
  const timerRef = useRef(null);

  const handlePointerDown = (studentId) => {
    if (!activeActivity) return;
    timerRef.current = setTimeout(() => {
      if (activeActivity.submissions?.[studentId]) updateAssignmentSubmissions(activeActivity.id, { [studentId]: null });
      timerRef.current = null;
    }, 500);
  };
  const handlePointerUp = (studentId) => {
    if (!activeActivity) return;
    if (timerRef.current) { clearTimeout(timerRef.current); timerRef.current = null; if (!activeActivity.submissions?.[studentId]) updateAssignmentSubmissions(activeActivity.id, { [studentId]: { timestamp: new Date().toISOString() } }); }
  };
  const handlePointerLeave = () => { if (timerRef.current) { clearTimeout(timerRef.current); timerRef.current = null; } };

  // 키보드(Enter/Space)로도 체크·해제할 수 있게
  const toggleByKey = (studentId) => {
    if (!activeActivity) return;
    updateAssignmentSubmissions(activeActivity.id, { [studentId]: activeActivity.submissions?.[studentId] ? null : { timestamp: new Date().toISOString() } });
  };
  const closeForm = () => { setTempActivityContent(''); setTempActivityType('과제'); setIsEditingActivity(false); };
  const saveActivity = () => {
    if (tempActivityContent.trim() !== '') { addAssignment({ id: generateUUID(), title: tempActivityContent, subject: '활동', type: tempActivityType, dueDate: toISODate(new Date()), createdAt: new Date().toISOString(), submissions: {}, total: students.length }); showToast('새로운 활동이 추가되었습니다.', 'success'); setSelectedAssignment(null); setCurrentPage(1); }
    closeForm();
  };

  const sortedAssignments = [...assignments].reverse();
  const totalPages = Math.ceil(sortedAssignments.length / itemsPerPage) || 1;
  const startIndex = (currentPage - 1) * itemsPerPage;
  const displayAssignments = sortedAssignments.slice(startIndex, startIndex + itemsPerPage);
  const doneCount = activeActivity ? students.filter(s => activeActivity.submissions?.[s.id]).length : 0;
  const donePct = students.length ? Math.round((doneCount / students.length) * 100) : 0;

  return (
    <div className="lrn">
      <div className="wk-top">
        <div className="wk-title">
          <h1 className="dash-headline">{activeActivity ? <>{activeActivity.title}</> : '오늘 할 활동을 정해요'}</h1>
          {activeActivity && <span className={`lrn-type ${activeActivity.type === '수행평가' ? 'is-eval' : ''}`}>{activeActivity.type || '수행평가'}</span>}
          {selectedAssignment && <button type="button" className="chip-btn" onClick={() => setSelectedAssignment(null)}>최신 활동으로</button>}
        </div>
        {isAdmin && !isEditingActivity && <button type="button" className="btn btn-primary" onClick={() => { closeForm(); setIsEditingActivity(true); }}><Icon name="plus" size={18} /> 활동 추가</button>}
      </div>

      {isAdmin && isEditingActivity && (
        <form className="lrn-form" onSubmit={e => { e.preventDefault(); saveActivity(); }}>
          <div className="mode-switch" role="group" aria-label="종류">
            {['과제', '수행평가'].map(t => <button key={t} type="button" aria-pressed={tempActivityType === t} onClick={() => setTempActivityType(t)}>{t}</button>)}
          </div>
          <input type="text" className="form-input" aria-label="활동 내용" value={tempActivityContent} onChange={e => setTempActivityContent(e.target.value)} placeholder="예: 수학 익힘책 34쪽" autoFocus />
          <div className="lrn-form-actions">
            <button type="button" className="btn btn-secondary btn-sm" onClick={closeForm}>취소</button>
            <button type="submit" className="btn btn-primary btn-sm">저장</button>
          </div>
        </form>
      )}

      <section className="wcard lrn-check">
        <div className="wcard-head">
          <h2 className="wcard-title">활동 완료 체크</h2>
          {activeActivity && <span className="wcard-meta"><strong className="num lrn-done">{doneCount}</strong> / <span className="num">{students.length}</span>명 완료</span>}
        </div>
        <div className="wcard-body">
          {activeActivity ? (
            <>
              <div className="lrn-track" role="progressbar" aria-label="완료율" aria-valuenow={donePct} aria-valuemin={0} aria-valuemax={100}><span style={{ width: `${donePct}%` }} className={donePct === 100 ? 'is-done' : ''} /></div>
              <p className="lrn-hint">이름을 누르면 완료, 길게 누르면 취소돼요.</p>
              <ul className="lrn-grid">
                {students.map(student => {
                  const isCompleted = !!activeActivity.submissions?.[student.id];
                  return (
                    <li key={student.id}>
                      <button type="button" aria-pressed={isCompleted} className={`lrn-student ${isCompleted ? 'is-done' : ''}`} onPointerDown={() => handlePointerDown(student.id)} onPointerUp={() => handlePointerUp(student.id)} onPointerLeave={handlePointerLeave} onContextMenu={e => e.preventDefault()} onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleByKey(student.id); } }}>
                        <span className="att-box" />
                        <span className="att-no num">{student.number}</span>
                        <span className="lrn-name">{student.name}</span>
                      </button>
                    </li>
                  );
                })}
              </ul>
            </>
          ) : (
            <div className="wcard-empty">{isAdmin ? '오른쪽 위 “활동 추가”로 새 활동을 만들어 주세요.' : '아직 등록된 활동이 없어요.'}</div>
          )}
        </div>
      </section>

      <section className="wcard">
        <div className="wcard-head">
          <h2 className="wcard-title">과제·수행평가 현황</h2>
          <span className="wcard-meta">눌러서 제출 명단 보기</span>
        </div>
        <div className="wcard-body">
          {displayAssignments.length > 0 ? (
            <ul className="lrn-list">
              {displayAssignments.map(a => {
                const total = a.total || students.length;
                const submittedCount = typeof a.submitted === 'number' ? a.submitted : Object.keys(a.submissions || {}).length;
                const pct = total ? Math.round((submittedCount / total) * 100) : 0;
                const remaining = total - submittedCount;
                const isOpen = selectedAssignment?.id === a.id;
                const done = students.filter(s => a.submissions?.[s.id]);
                const notDone = students.filter(s => !a.submissions?.[s.id]);
                return (
                  <li key={a.id} className={`lrn-item ${isOpen ? 'is-open' : ''}`}>
                    <div className="lrn-row">
                      <button type="button" className="lrn-row-main" aria-expanded={isOpen} onClick={() => setSelectedAssignment(isOpen ? null : a.id)}>
                        <span className={`lrn-type ${a.type === '수행평가' ? 'is-eval' : ''}`}>{a.type}</span>
                        <span className="lrn-row-text">
                          <span className="lrn-row-title">{a.title}</span>
                          <span className="lrn-row-meta"><span className="num">{a.dueDate}</span>{remaining > 0 ? <span className="lrn-left">미제출 {remaining}명</span> : <span className="lrn-all">모두 제출</span>}</span>
                        </span>
                        <span className="lrn-row-bar"><span style={{ width: `${pct}%` }} className={pct === 100 ? 'is-done' : ''} /></span>
                        <span className="lrn-pct num">{pct}<small>%</small></span>
                      </button>
                      {isAdmin && <button type="button" className="icon-ghost" aria-label={`${a.title} 삭제`} onClick={() => { if (confirm('이 항목을 정말 삭제하시겠습니까?')) { deleteAssignment(a.id); if (isOpen) setSelectedAssignment(null); showToast('삭제되었습니다.', 'success'); } }}><Icon name="trash" size={18} /></button>}
                    </div>
                    {isOpen && (
                      <div className="lrn-who">
                        <div className="lrn-who-group">
                          <span className="lrn-who-label is-done">제출 <span className="num">{done.length}</span></span>
                          <div className="lrn-chips">{done.length ? done.map(s => <span key={s.id} className="lrn-chip is-done"><span className="num">{s.number}</span>{s.name}</span>) : <span className="lrn-none">아직 없어요</span>}</div>
                        </div>
                        <div className="lrn-who-group">
                          <span className="lrn-who-label is-left">미제출 <span className="num">{notDone.length}</span></span>
                          <div className="lrn-chips">{notDone.length ? notDone.map(s => <span key={s.id} className="lrn-chip is-left"><span className="num">{s.number}</span>{s.name}</span>) : <span className="lrn-none">모두 냈어요</span>}</div>
                        </div>
                      </div>
                    )}
                  </li>
                );
              })}
            </ul>
          ) : (<div className="wcard-empty">등록된 과제가 없어요. 활동을 추가하면 여기에 쌓여요.</div>)}
          {totalPages > 1 && (
            <div className="lrn-pager">
              <button type="button" className="round-btn round-btn--sm" aria-label="이전 쪽" disabled={currentPage === 1} onClick={() => setCurrentPage(p => Math.max(1, p - 1))}><Icon name="chevron" size={18} /></button>
              <span className="num">{currentPage} / {totalPages}</span>
              <button type="button" className="round-btn round-btn--sm wk-next" aria-label="다음 쪽" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}><Icon name="chevron" size={18} /></button>
            </div>
          )}
        </div>
      </section>
    </div>
  );
}

/* ===================================================================
   CalendarPage
   =================================================================== */
function CalendarPage() {
  const { state } = useApp();
  const { isAdmin, settings } = state;
  const calendars = [settings?.calendarId1, settings?.calendarId2, settings?.calendarId3].filter(Boolean);

  const getCombinedCalendarSrc = (cals) => {
    if (cals.length === 0) return '';
    let baseUrl = 'https://calendar.google.com/calendar/embed?ctz=Asia%2FSeoul&showPrint=0&showCalendars=0&showTz=0&mode=MONTH';
    const colors = ['%23039BE5', '%23D50000', '%230B8043'];
    cals.forEach((input, index) => {
      const calId = extractCalendarId(input);
      if (calId) baseUrl += `&src=${encodeURIComponent(calId)}&color=${colors[index % colors.length]}`;
    });
    return baseUrl;
  };

  return (
    <div className="calendar-page" style={{ height: '100%' }}>
      {calendars.length > 0 ? (
        <div className="card google-calendar-embed" style={{ height: '100%', padding: '0' }}><iframe src={getCombinedCalendarSrc(calendars)} width="100%" height="800" frameBorder="0" title="Google Calendar Combined" style={{ display: 'block', border: 'none' }}></iframe></div>
      ) : (
        <div className="card google-calendar-embed" style={{ padding: '3rem', textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: '800px' }}>
          <h3>📅 구글 캘린더가 연동되지 않았습니다.</h3>
          {isAdmin ? (<p style={{ color: 'var(--text-secondary)', marginTop: '0.5rem' }}>우측 상단의 ⚙️ 설정 메뉴에서 구글 캘린더 ID를 입력해주세요.</p>) : (<p style={{ color: 'var(--text-secondary)', marginTop: '0.5rem' }}>관리자가 캘린더를 연동하면 이 곳에 표시됩니다.</p>)}
        </div>
      )}
    </div>
  );
}

/* ===================================================================
   ToolsPage
   =================================================================== */
function ToolsPage() {
  const { state, showToast, addPoll, deletePoll, votePoll } = useApp();
  const { students, polls, isAdmin } = state;
  const [pickerResult, setPickerResult] = useState('?');
  const [pickerHistory, setPickerHistory] = useState([]);
  const [isSpinning, setIsSpinning] = useState(false);
  const [pickerCount, setPickerCount] = useState(1);
  const spinRef = useRef(null);

  const handlePick = useCallback(() => {
    if (isSpinning) return;
    const available = students.filter(s => !pickerHistory.find(h => h.id === s.id));
    if (available.length === 0) { showToast('모든 학생이 이미 뽑혔습니다!', 'warning'); return; }
    setIsSpinning(true);
    let count = 0;
    spinRef.current = setInterval(() => {
      const rand = available[Math.floor(Math.random() * available.length)];
      setPickerResult(rand.name);
      count++;
      if (count >= 20) {
        clearInterval(spinRef.current);
        const picks = shuffleArray(available).slice(0, Math.min(pickerCount, available.length));
        setPickerResult(picks.map(p => p.name).join(', '));
        setPickerHistory(prev => [...prev, ...picks]);
        setIsSpinning(false);
      }
    }, 80);
  }, [isSpinning, students, pickerHistory, pickerCount, showToast]);

  const resetPicker = () => { setPickerHistory([]); setPickerResult('?'); };

  const [timerMode, setTimerMode] = useState('timer');
  const [timerActivity, setTimerActivity] = useState(''); // 지금 할 활동 이름 (타이머 위에 크게)
  const [timerSeconds, setTimerSeconds] = useState(300);
  const [timerRunning, setTimerRunning] = useState(false);
  const [timerElapsed, setTimerElapsed] = useState(0);
  const [isAlarm, setIsAlarm] = useState(false);
  const [customMin, setCustomMin] = useState('');
  const [customSec, setCustomSec] = useState('');
  const timerRef = useRef(null);
  const presets = [{ label: '1분', sec: 60 }, { label: '3분', sec: 180 }, { label: '5분', sec: 300 }, { label: '10분', sec: 600 }, { label: '15분', sec: 900 }];
  const formatTimer = (sec) => `${String(Math.floor(sec / 60)).padStart(2, '0')}:${String(sec % 60).padStart(2, '0')}`;
  const formatStopwatch = (ms) => { const totalSec = Math.floor(ms / 1000); return `${String(Math.floor(totalSec / 60)).padStart(2, '0')}:${String(totalSec % 60).padStart(2, '0')}.${String(Math.floor((ms % 1000) / 10)).padStart(2, '0')}`; };

  const startTimer = () => {
    if (timerRunning) return;
    setIsAlarm(false); setTimerRunning(true);
    if (timerMode === 'timer') {
      timerRef.current = setInterval(() => { setTimerSeconds(prev => { if (prev <= 1) { clearInterval(timerRef.current); setTimerRunning(false); setIsAlarm(true); showToast('⏰ 시간 종료!', 'warning'); return 0; } return prev - 1; }); }, 1000);
    } else {
      const startTime = Date.now() - timerElapsed;
      timerRef.current = setInterval(() => setTimerElapsed(Date.now() - startTime), 10);
    }
  };
  const pauseTimer = () => { clearInterval(timerRef.current); setTimerRunning(false); };
  const resetTimer = () => { clearInterval(timerRef.current); setTimerRunning(false); setTimerSeconds(300); setTimerElapsed(0); setIsAlarm(false); };
  const applyCustomTime = () => {
    const m = parseInt(customMin) || 0, s = parseInt(customSec) || 0;
    if (m === 0 && s === 0) { showToast('시간을 입력해주세요', 'error'); return; }
    setTimerSeconds(m * 60 + s); showToast(`${m}분 ${s}초로 설정되었습니다`, 'success');
  };
  useEffect(() => () => { clearInterval(timerRef.current); clearInterval(spinRef.current); }, []);

  const [groupCount, setGroupCount] = useState(4);
  const [groups, setGroups] = useState([]);
  const makeGroups = () => {
    const shuffled = shuffleArray(students);
    const result = Array.from({ length: groupCount }, () => []);
    shuffled.forEach((s, i) => result[i % groupCount].push(s));
    setGroups(result); showToast(`${groupCount}개 팀이 편성되었습니다!`, 'success');
  };

  const [newPollQ, setNewPollQ] = useState('');
  const [newPollOpts, setNewPollOpts] = useState(['', '']);
  const getVotedPolls = () => { try { return JSON.parse(sessionStorage.getItem('votedPolls') || '{}'); } catch { return {}; } };
  const markPollVoted = (pollId) => { const voted = getVotedPolls(); voted[pollId] = true; sessionStorage.setItem('votedPolls', JSON.stringify(voted)); };
  const hasVoted = (pollId) => !!getVotedPolls()[pollId];

  const handleAddPoll = () => {
    if (!newPollQ.trim() || newPollOpts.filter(o => o.trim()).length < 2) { showToast('질문과 최소 2개 항목을 입력하세요', 'error'); return; }
    const opts = newPollOpts.filter(o => o.trim());
    addPoll({ id: generateUUID(), question: newPollQ, options: opts, votes: opts.map(() => 0), isActive: true });
    setNewPollQ(''); setNewPollOpts(['', '']); showToast('투표가 생성되었습니다!', 'success');
  };
  const handleVote = (pollId, optIdx) => {
    if (hasVoted(pollId)) { showToast('이미 투표하셨습니다!', 'warning'); return; }
    votePoll(pollId, optIdx); markPollVoted(pollId); showToast('투표가 완료되었습니다!', 'success');
  };
  const handleDeletePoll = (pollId) => { deletePoll(pollId); showToast('투표가 삭제되었습니다.', 'success'); };

  // 도구 카드 하나를 창 전체로 키움(같은 화면 요소라 타이머·뽑기가 끊기지 않음). 브라우저 전체화면은 되면 덤으로 켬.
  const [fsTool, setFsTool] = useState(null);
  const openFull = (id) => { setFsTool(id); document.documentElement.requestFullscreen?.().catch(() => {}); };
  const closeFull = () => { setFsTool(null); if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); };
  useEffect(() => {
    if (!fsTool) return;
    const onKey = (e) => { if (e.key === 'Escape') closeFull(); };
    const onFs = () => { if (!document.fullscreenElement) setFsTool(null); }; // 브라우저 전체화면을 Esc로 끄면 같이 작게
    document.addEventListener('keydown', onKey);
    document.addEventListener('fullscreenchange', onFs);
    document.body.classList.add('has-full-tool');
    return () => { document.removeEventListener('keydown', onKey); document.removeEventListener('fullscreenchange', onFs); document.body.classList.remove('has-full-tool'); };
  }, [fsTool]);
  const fsButton = (id) => (
    <button type="button" className="tool-fs-btn" aria-pressed={fsTool === id} onClick={() => (fsTool === id ? closeFull() : openFull(id))}>
      <Icon name={fsTool === id ? 'shrink' : 'expand'} size={16} />{fsTool === id ? '작게 보기' : '전체화면'}
    </button>
  );

  return (
    <div className="tools-page">
      <div className="tools-grid">
        <div data-tool="picker" className={`tool-card ${!isAdmin ? 'disabled-tool' : ''} ${fsTool === 'picker' ? 'is-full' : ''}`}>
          <div className="tool-card-header"><span className="tool-emoji">🎲</span><div><div className="tool-title">랜덤 뽑기</div><div className="tool-desc">발표자를 랜덤으로 선택해요</div></div>{isAdmin && fsButton('picker')}</div>
          <div className="picker-container">
            <div className="roulette-wheel"><div className="roulette-pointer" /><div className={`roulette-circle ${isSpinning ? 'spinning' : pickerResult !== '?' ? 'reveal' : ''}`}><div className="roulette-result">{pickerResult}</div></div></div>
            {isAdmin && (<div className="picker-count"><label>뽑을 인원:</label><input type="number" min="1" max="28" value={pickerCount} onChange={e => setPickerCount(Math.max(1, parseInt(e.target.value) || 1))} /><span>명</span></div>)}
            {isAdmin && (<div style={{ display: 'flex', gap: '0.5rem', justifyContent: 'center' }}><button className="btn btn-primary btn-lg" onClick={handlePick} disabled={isSpinning}>{isSpinning ? '🎰 뽑는 중...' : '🎲 뽑기!'}</button><button className="btn btn-secondary" onClick={resetPicker}>초기화</button></div>)}
            {pickerHistory.length > 0 && (<div className="picker-history">{pickerHistory.map((s, i) => <span key={i} className="picked-badge">{s.name}</span>)}</div>)}
          </div>
        </div>

        <div data-tool="timer" className={`tool-card ${!isAdmin ? 'disabled-tool' : ''} ${fsTool === 'timer' ? 'is-full' : ''}`}>
          <div className="tool-card-header"><span className="tool-emoji">⏱️</span><div><div className="tool-title">타이머 / 스톱워치</div><div className="tool-desc">활동 시간을 관리해요</div></div>{isAdmin && fsButton('timer')}</div>
          <div className="timer-container">
            <input type="text" className="timer-activity" aria-label="활동 이름" placeholder="무엇을 할 시간인가요? (예: 모둠 토의)" maxLength={40} value={timerActivity} onChange={e => setTimerActivity(e.target.value)} />
            <div className="timer-mode-toggle"><button className={`timer-mode-btn ${timerMode === 'timer' ? 'active' : ''}`} onClick={() => { resetTimer(); setTimerMode('timer'); }}>⏳ 타이머</button><button className={`timer-mode-btn ${timerMode === 'stopwatch' ? 'active' : ''}`} onClick={() => { resetTimer(); setTimerMode('stopwatch'); }}>⏱️ 스톱워치</button></div>
            <div className={`timer-display ${timerRunning ? 'running' : ''} ${isAlarm ? 'alarm' : ''}`}>{timerMode === 'timer' ? formatTimer(timerSeconds) : formatStopwatch(timerElapsed)}</div>
            {timerMode === 'timer' && (<>
              <div className="timer-presets">{presets.map(p => (<button key={p.sec} className={`timer-preset ${timerSeconds === p.sec && !timerRunning ? 'active' : ''}`} onClick={() => { if (!timerRunning) setTimerSeconds(p.sec); }}>{p.label}</button>))}</div>
              {!timerRunning && (<div className="timer-custom-input"><input type="number" className="custom-time-input" placeholder="분" min="0" max="99" value={customMin} onChange={e => setCustomMin(e.target.value)} /><span className="time-separator">:</span><input type="number" className="custom-time-input" placeholder="초" min="0" max="59" value={customSec} onChange={e => setCustomSec(e.target.value)} /><button className="btn btn-sm btn-primary" onClick={applyCustomTime}>적용</button></div>)}
            </>)}
            <div className="timer-controls">{!timerRunning ? <button className="btn btn-primary" onClick={startTimer}>▶ 시작</button> : <button className="btn btn-secondary" onClick={pauseTimer}>⏸ 일시정지</button>}<button className="btn btn-danger" onClick={resetTimer} style={{ background: 'var(--danger)', color: 'white' }}>↺ 초기화</button></div>
          </div>
        </div>

        <div data-tool="groups" className={`tool-card ${!isAdmin ? 'disabled-tool' : ''} ${fsTool === 'groups' ? 'is-full' : ''}`}>
          <div className="tool-card-header"><span className="tool-emoji">👥</span><div><div className="tool-title">팀 편성</div><div className="tool-desc">랜덤으로 팀을 나눠요</div></div>{isAdmin && fsButton('groups')}</div>
          <div className="group-maker">
            {isAdmin && (<div className="group-settings"><div className="group-setting"><label>팀 수:</label><input type="number" min="2" max="14" value={groupCount} onChange={e => setGroupCount(Math.max(2, parseInt(e.target.value) || 2))} /><span>개</span></div><button className="btn btn-primary" onClick={makeGroups}>🔀 팀 편성하기</button></div>)}
            {groups.length > 0 && (<div className="groups-result">{groups.map((g, i) => (<div key={i} className="group-card" style={{ animationDelay: `${i * 0.1}s` }}><div className="group-card-title">{i + 1}팀</div><div className="group-members">{g.map(s => (<div key={s.id} className="group-member"><span>{s.gender === 'M' ? '👦' : '👧'}</span><span>{s.number}. {s.name}</span></div>))}</div></div>))}</div>)}
          </div>
        </div>

        <div data-tool="poll" className={`tool-card ${fsTool === 'poll' ? 'is-full' : ''}`}>
          <div className="tool-card-header"><span className="tool-emoji">📊</span><div><div className="tool-title">투표 / 설문</div><div className="tool-desc">학급 의견을 모아요</div></div>{fsButton('poll')}</div>
          <div className="poll-container">
            {isAdmin && (
              <div className="poll-form">
                <input className="form-input" placeholder="투표 질문을 입력하세요" value={newPollQ} onChange={e => setNewPollQ(e.target.value)} style={{ marginBottom: '0.5rem' }} />
                <div className="poll-options-list">{newPollOpts.map((opt, i) => (<div key={i} className="poll-option-input"><input className="form-input" placeholder={`항목 ${i + 1}`} value={opt} onChange={e => { const o = [...newPollOpts]; o[i] = e.target.value; setNewPollOpts(o); }} />{newPollOpts.length > 2 && <button className="poll-remove" onClick={() => setNewPollOpts(newPollOpts.filter((_, j) => j !== i))}>×</button>}</div>))}</div>
                <div style={{ display: 'flex', gap: '0.5rem' }}><button className="btn btn-secondary btn-sm" onClick={() => setNewPollOpts([...newPollOpts, ''])}>+ 항목 추가</button><button className="btn btn-primary btn-sm" onClick={handleAddPoll}>투표 생성</button></div>
              </div>
            )}
            {polls.map(p => {
              const totalVotes = p.votes.reduce((a, b) => a + b, 0);
              const voted = hasVoted(p.id);
              return (
                <div key={p.id} className="poll-active">
                  <div className="poll-header-row"><h4 className="poll-question">{p.question}</h4>{isAdmin && <button className="poll-delete-btn" onClick={() => handleDeletePoll(p.id)}>🗑️</button>}</div>
                  {voted && <div className="poll-voted-badge">✅ 투표 완료</div>}
                  <div className="poll-results">{p.options.map((opt, i) => { const pct = totalVotes > 0 ? Math.round((p.votes[i] / totalVotes) * 100) : 0; return (<div key={i} className={`poll-result-item ${voted ? 'voted' : ''}`} onClick={() => !voted && handleVote(p.id, i)} style={{ cursor: voted ? 'default' : 'pointer' }}><span className="poll-result-label">{opt}</span><div className="poll-result-bar"><div className="poll-result-fill" style={{ width: `${pct}%` }}>{pct > 10 ? `${pct}%` : ''}</div></div><span className="poll-result-count">{p.votes[i]}표</span></div>); })}</div>
                  <div className="poll-total">총 {totalVotes}표</div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ===================================================================
   WeeklyPlanPage
   =================================================================== */
const WEEK_DAYS = ['월', '화', '수', '목', '금'];
const WEEK_PERIODS = [1, 2, 3, 4, 5, 6];

function WeeklyPlanPage() {
  const { state, showToast, saveWeeklyPlan } = useApp();
  const { weeklyPlans, isAdmin, settings } = state;
  const [currentDate, setCurrentDate] = useState(new Date());
  const [editing, setEditing] = useState(null);
  const [editSubject, setEditSubject] = useState('');
  const [editContent, setEditContent] = useState('');

  const weekStart = useMemo(() => getWeekStart(currentDate), [currentDate]);
  const weekKey = toISODate(weekStart);
  const dayDates = useMemo(() => WEEK_DAYS.map((d, i) => { const date = new Date(weekStart); date.setDate(date.getDate() + i); return { day: d, dateStr: `${date.getMonth() + 1}/${date.getDate()}`, isoDate: toISODate(date) }; }), [weekStart]);
  const weekEndDate = new Date(weekStart); weekEndDate.setDate(weekEndDate.getDate() + 4);
  const currentWeekPlans = weeklyPlans[weekKey] || {};
  const todayStr = toISODate(new Date());
  const isCurrentWeek = dayDates.some(dd => dd.isoDate === todayStr);

  const goToPrevWeek = () => { const d = new Date(currentDate); d.setDate(d.getDate() - 7); setCurrentDate(d); setEditing(null); };
  const goToNextWeek = () => { const d = new Date(currentDate); d.setDate(d.getDate() + 7); setCurrentDate(d); setEditing(null); };
  const goToToday = () => { setCurrentDate(new Date()); setEditing(null); };
  const handleEdit = (day, period, subject, content) => { if (!isAdmin) return; setEditing({ day, period }); setEditSubject(subject || ''); setEditContent(content || ''); };
  const handleSave = useCallback(() => { saveWeeklyPlan(weekKey, editing.day, editing.period, editSubject, editContent); setEditing(null); showToast('주간 일정이 저장되었습니다.', 'success'); }, [saveWeeklyPlan, weekKey, editing, editSubject, editContent, showToast]);

  // 구글 달력 일정: 이 주(월~금)를 읽어 교시 칸·요일 머리에 함께 표시 (주간 일정 데이터는 건드리지 않음)
  const calInputs = [settings?.calendarId1, settings?.calendarId2, settings?.calendarId3].filter(Boolean);
  const calKey = (settings?.googleApiKey || '').trim();
  const [cal, setCal] = useState({ status: 'off', cells: {}, dayNotes: {}, count: 0 });
  useEffect(() => {
    if (!calKey || !calInputs.length) { setCal({ status: 'off', cells: {}, dayNotes: {}, count: 0 }); return; }
    let alive = true;
    const end = new Date(weekStart); end.setDate(end.getDate() + 5);
    setCal(c => ({ ...c, status: 'loading' }));
    fetchCalendarEvents(calInputs, calKey, weekStart, end)
      .then(evs => { if (alive) setCal({ status: 'ok', count: evs.length, ...placeEventsInWeek(evs, dayDates.map(d => ({ day: d.day, iso: d.isoDate })), PERIOD_SLOTS) }); })
      .catch(err => { if (alive) setCal({ status: 'error', error: err.message, cells: {}, dayNotes: {}, count: 0 }); });
    return () => { alive = false; };
  }, [weekKey, calKey, calInputs.join('|')]); // eslint-disable-line react-hooks/exhaustive-deps
  const calLines = (key) => (cal.cells[key] || []).map((t, i) => <span key={i} className="wk-cal">{t}</span>);

  const sameMonth = weekStart.getMonth() === weekEndDate.getMonth();
  const rangeLabel = `${weekStart.getMonth() + 1}월 ${weekStart.getDate()}일 – ${sameMonth ? '' : `${weekEndDate.getMonth() + 1}월 `}${weekEndDate.getDate()}일`;

  const renderRow = (p) => (
    <tr key={p}>
      <th scope="row" className="wk-period"><span className="wk-period-no">{p}교시</span><span className="wk-period-time num">{minToClock(PERIOD_SLOTS[p][0])}</span></th>
      {dayDates.map(dd => {
        const plan = currentWeekPlans[dd.day]?.[p] || {};
        const isEditing = editing?.day === dd.day && editing?.period === p;
        const isToday = dd.isoDate === todayStr;
        const isEmpty = !plan.subject && !plan.content;
        return (
          <td key={dd.day} className={`wk-cell ${isToday ? 'is-today' : ''} ${isEditing ? 'is-editing' : ''}`}>
            {isEditing ? (
              <div className="wk-form">
                <input className="form-input" aria-label={`${dd.day}요일 ${p}교시 과목`} placeholder="과목 (예: 국어)" value={editSubject} onChange={e => setEditSubject(e.target.value)} autoFocus />
                <input className="form-input" aria-label={`${dd.day}요일 ${p}교시 학습 내용`} placeholder="학습 내용" value={editContent} onChange={e => setEditContent(e.target.value)} onKeyDown={e => e.key === 'Enter' && handleSave()} />
                <div className="wk-form-actions"><button type="button" className="btn btn-sm btn-secondary" onClick={() => setEditing(null)}>취소</button><button type="button" className="btn btn-sm btn-primary" onClick={handleSave}>저장</button></div>
              </div>
            ) : isAdmin ? (
              <button type="button" data-tone={subjectTone(plan.subject)} className={`wk-plan ${isEmpty ? 'is-empty' : ''}`} aria-label={`${dd.day}요일 ${p}교시 ${isEmpty ? '입력' : '수정'}`} onClick={() => handleEdit(dd.day, p, plan.subject, plan.content)}>
                {isEmpty ? <span className="wk-add"><Icon name="plus" size={14} />입력</span> : <><span className="wk-subject">{plan.subject}</span>{plan.content && <span className="wk-content">{plan.content}</span>}</>}
                {calLines(`${dd.day}-${p}`)}
              </button>
            ) : (
              <div data-tone={subjectTone(plan.subject)} className={`wk-plan ${isEmpty ? 'is-empty' : ''}`}>{plan.subject && <span className="wk-subject">{plan.subject}</span>}{plan.content && <span className="wk-content">{plan.content}</span>}{calLines(`${dd.day}-${p}`)}</div>
            )}
          </td>
        );
      })}
    </tr>
  );

  return (
    <div className="wk">
      <div className="wk-top">
        <div className="wk-title">
          <h1 className="dash-headline">{rangeLabel}</h1>
          {isCurrentWeek && <span className="wk-chip">이번 주</span>}
        </div>
        <div className="wk-nav">
          <button type="button" className="round-btn" aria-label="이전 주" onClick={goToPrevWeek}><Icon name="chevron" size={18} /></button>
          <button type="button" className="btn btn-sm btn-secondary" onClick={goToToday} disabled={isCurrentWeek}>오늘</button>
          <button type="button" className="round-btn wk-next" aria-label="다음 주" onClick={goToNextWeek}><Icon name="chevron" size={18} /></button>
        </div>
      </div>
      {isAdmin && <p className="dash-edit-hint">칸을 눌러 과목과 학습 내용을 적고, Enter로 저장하세요.</p>}
      {cal.status === 'ok' && <p className="wk-cal-status"><Icon name="calendar" size={16} />{cal.count ? `구글 달력 일정 ${cal.count}개를 함께 보여 주고 있어요.` : '이번 주에는 구글 달력 일정이 없어요.'}</p>}
      {cal.status === 'loading' && <p className="wk-cal-status"><Icon name="calendar" size={16} />구글 달력 일정을 불러오는 중…</p>}
      {cal.status === 'error' && <p className="wk-cal-status is-error"><Icon name="calendar" size={16} />구글 달력을 불러오지 못했어요 ({cal.error}). 설정의 API 키와 달력 공개 여부를 확인해 주세요.</p>}
      {cal.status === 'off' && isAdmin && <p className="wk-cal-status"><Icon name="calendar" size={16} />⚙️ 설정에 구글 달력과 Google API 키를 넣으면 달력 일정이 이 표에 함께 나와요.</p>}
      <div className="wk-card">
        <table className="wk-table">
          <thead>
            <tr>
              <th scope="col" className="wk-corner">교시</th>
              {dayDates.map(dd => (
                <th scope="col" key={dd.day} className={dd.isoDate === todayStr ? 'is-today' : ''}>
                  <span className="wk-day">{dd.day}</span><span className="wk-date num">{dd.dateStr}</span>
                  {(cal.dayNotes[dd.day] || []).map((t, i) => <span key={i} className="wk-cal-day">{t}</span>)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {WEEK_PERIODS.map(renderRow)}
          </tbody>
        </table>
      </div>
    </div>
  );
}


/* ===================================================================
   LinksPage
   =================================================================== */
const LINK_EMOJIS = [['🔗', '링크'], ['🌐', '웹사이트'], ['🏫', '학교'], ['📚', '책·자료'], ['📝', '과제·문서'], ['📊', '통계·표'], ['📅', '일정'], ['💬', '소통·메신저'], ['🎬', '영상'], ['🎵', '음악'], ['🎨', '미술'], ['🧮', '수학'], ['🔬', '과학'], ['🌏', '지도·사회'], ['🎮', '게임']];

function LinksPage() {
  const { state, showToast, addLink, updateLink, deleteLink } = useApp();
  const { quickLinks, isAdmin } = state;
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [formData, setFormData] = useState({ title: '', url: '', icon: '🔗', desc: '' });

  const handleEdit = (link) => { setEditingId(link.id); setFormData(link); setShowForm(true); };
  const handleDelete = (id) => { if (confirm('정말 삭제하시겠습니까?')) { deleteLink(id); showToast('삭제되었습니다.', 'success'); } };
  const handleSave = () => {
    if (!formData.title || !formData.url) { showToast('이름과 주소를 모두 입력해주세요.', 'error'); return; }
    const url = formData.url.startsWith('http') ? formData.url : `https://${formData.url}`;
    if (editingId) { updateLink({ ...formData, url, id: editingId }); showToast('수정되었습니다.', 'success'); }
    else { addLink({ ...formData, url, id: generateUUID() }); showToast('추가되었습니다.', 'success'); }
    setShowForm(false); setEditingId(null); setFormData({ title: '', url: '', icon: '🔗', desc: '' });
  };
  const handleCancel = () => { setShowForm(false); setEditingId(null); setFormData({ title: '', url: '', icon: '🔗', desc: '' }); };

  return (
    <div className="links-page">
      <div className="links-header"><h2 className="section-title">🔗 즐겨찾기 & 바로가기</h2>{isAdmin && <button className="btn btn-primary" onClick={() => setShowForm(true)} disabled={showForm}>+ 링크 추가</button>}</div>
      {isAdmin && showForm && (
        <div className="card links-form">
          <h3 style={{ marginBottom: '1rem', fontSize: '1.125rem' }}>{editingId ? '링크 수정' : '새 링크 추가'}</h3>
          <div className="form-group">
            <label id="link-icon-label">아이콘 고르기</label>
            <div className="emoji-picker" role="group" aria-labelledby="link-icon-label">
              {LINK_EMOJIS.map(([e, name]) => <button key={e} type="button" className="emoji-opt" aria-pressed={formData.icon === e} aria-label={name} title={name} onClick={() => setFormData({ ...formData, icon: e })}>{e}</button>)}
            </div>
          </div>
          <div className="form-row">
            <div className="form-group" style={{ flex: 1 }}><label>직접 입력</label><input className="form-input" value={formData.icon} onChange={e => setFormData({ ...formData, icon: e.target.value })} maxLength="2" /></div>
            <div className="form-group" style={{ flex: 4 }}><label>이름</label><input className="form-input" placeholder="예: 구글 번역" value={formData.title} onChange={e => setFormData({ ...formData, title: e.target.value })} /></div>
          </div>
          <div className="form-group"><label>주소 (URL)</label><input className="form-input" placeholder="https://..." value={formData.url} onChange={e => setFormData({ ...formData, url: e.target.value })} /></div>
          <div className="form-group"><label>설명 (선택)</label><input className="form-input" placeholder="어떤 사이트인지 설명..." value={formData.desc} onChange={e => setFormData({ ...formData, desc: e.target.value })} /></div>
          <div className="form-actions"><button className="btn btn-secondary" onClick={handleCancel}>취소</button><button className="btn btn-primary" onClick={handleSave}>저장</button></div>
        </div>
      )}
      <div className="links-grid">
        {quickLinks.map(link => (
          <div key={link.id} className="link-card-wrapper">
            <a href={link.url} target="_blank" rel="noopener noreferrer" className="link-card card"><div className="link-icon">{link.icon}</div><div className="link-content"><h3 className="link-title">{link.title}</h3>{link.desc && <p className="link-desc">{link.desc}</p>}</div></a>
            {isAdmin && (<div className="link-actions"><button className="btn-icon link-action-btn" onClick={() => handleEdit(link)}>✏️</button><button className="btn-icon link-action-btn" onClick={() => handleDelete(link.id)}>🗑️</button></div>)}
          </div>
        ))}
        {quickLinks.length === 0 && (<div className="empty-state" style={{ gridColumn: '1 / -1', textAlign: 'center', padding: '3rem', color: 'var(--text-secondary)' }}>등록된 즐겨찾기가 없습니다.</div>)}
      </div>
    </div>
  );
}

/* ===================================================================
   Root
   =================================================================== */
const ROUTE_PAGES = { '/': DashboardPage, '/students': StudentsPage, '/weekly': WeeklyPlanPage, '/attendance': AttendancePage, '/learning': LearningPage, '/calendar': CalendarPage, '/tools': ToolsPage, '/links': LinksPage };

function AppShell() {
  const pathname = useHashRoute();
  const Page = ROUTE_PAGES[pathname] || DashboardPage;
  return (
    <>
      <Layout pathname={pathname}><Page /></Layout>
      <Toast />
    </>
  );
}

function Root() {
  return (<AppProvider><AppShell /></AppProvider>);
}

ReactDOM.createRoot(document.getElementById('root')).render(<Root />);
