const { useState, useEffect, useRef, useCallback, useReducer, useContext, useMemo, createContext } = React;
const {
  getDayName, formatDate, getDday, isTodayBirthday, getAvatarColor, getInitials,
  getCurrentPeriod, shuffleArray, generateUUID, getWeekStart, toISODate,
} = window.Helpers;
const { fetchTodayLunch, fetchSchoolHoliday } = window.NeisApi;
const { generateBehaviorSummary } = window.GeminiApi;
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
          [date]: { ...dateRecords, [studentId]: { status, note: dateRecords[studentId]?.note || '', id: id || dateRecords[studentId]?.id } },
        },
      };
    }
    case 'SET_ATTENDANCE':
      return {
        ...state,
        attendance: {
          ...state.attendance,
          [action.payload.studentId]: { status: action.payload.status, note: action.payload.note || '', id: action.payload.id || state.attendance[action.payload.studentId]?.id },
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

  const showToast = useCallback((message, type = 'info') => {
    const id = Date.now() + Math.random();
    dispatch({ type: 'ADD_TOAST', payload: { id, message, type } });
    setTimeout(() => dispatch({ type: 'REMOVE_TOAST', payload: id }), 3000);
  }, []);

  const loadFromFolder = useCallback(async (handle) => {
    const db = await readDb(handle);
    dispatch({ type: 'SET_STUDENTS', payload: db.students || [] });
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
    saveTimer.current = setTimeout(() => {
      const payload = {
        students: state.students, attendance: state.monthlyAttendance, announcements: state.announcements,
        assignments: state.assignments, ddays: state.ddays, polls: state.polls, newsletters: state.newsletters,
        quickLinks: state.quickLinks, weeklyPlans: state.weeklyPlans, studentRecords: state.studentRecords,
        settings: state.settings,
      };
      writeDb(folder, payload).catch(err => console.error('로컬 저장 오류:', err));
    }, 500);
    return () => clearTimeout(saveTimer.current);
  }, [state, isDbLoading, folderStatus, folder]);

  const saveWeeklyPlan = useCallback((weekKey, day, period, subject, content) => {
    dispatch({ type: 'SET_WEEKLY_PLAN', payload: { weekKey, day, period, subject, content } });
  }, []);

  const addDDay = (dday) => dispatch({ type: 'ADD_DDAY', payload: { id: generateUUID(), name: dday.name, date: dday.date, emoji: dday.emoji } });
  const deleteDDayFunc = (id) => dispatch({ type: 'DELETE_DDAY', payload: id });
  const updatePointsFunc = (studentId, _cur, amount) => dispatch({ type: 'UPDATE_POINTS', payload: { studentId, amount } });
  const resetPointsFunc = (studentId) => dispatch({ type: 'RESET_POINTS', payload: { studentId } });

  const updateAttendanceFunc = (studentId, status, note = '', date = null) => {
    const targetDate = date || toISODate(new Date());
    dispatch({ type: 'SET_ATTENDANCE', payload: { studentId, status, note, id: '' } });
    dispatch({ type: 'UPDATE_MONTHLY_ATTENDANCE', payload: { date: targetDate, studentId, status, id: '' } });
  };

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
    updateAttendance: updateAttendanceFunc, loadAttendanceByDate: loadAttendanceByDateFunc, loadMonthlyAttendance: loadMonthlyAttendanceFunc, saveAttendanceDay: saveAttendanceDayFunc,
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

  const handleSave = () => {
    updateSettings({ ...settings, calendarId1, calendarId2, calendarId3, geminiApiKey });
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
  { path: '/', label: '대시보드', icon: <Icon name="dashboard" /> },
  { path: '/weekly', label: '주간 일정', icon: <Icon name="weekly" /> },
  { path: '/attendance', label: '출석 관리', icon: <Icon name="attendance" /> },
  { path: '/learning', label: '학습 관리', icon: <Icon name="learning" /> },
  { path: '/calendar', label: '일정 & 공지', icon: <Icon name="calendar" /> },
  { path: '/tools', label: '학급 도구', icon: <Icon name="tools" /> },
  { path: '/links', label: '바로가기', icon: <Icon name="links" /> },
  { path: '/students', label: '학생 관리', icon: <Icon name="students" /> },
];
const PAGE_TITLES = { '/': '대시보드', '/weekly': '주간 일정', '/attendance': '출석 관리', '/learning': '학습 관리', '/calendar': '일정 & 공지', '/tools': '학급 도구', '/links': '바로가기', '/students': '학생 관리' };
const MOBILE_LABELS = { '/': '대시보드', '/weekly': '주간', '/attendance': '출석', '/learning': '학습', '/calendar': '일정', '/tools': '도구', '/links': '링크', '/students': '학생' };

function Layout({ pathname, children }) {
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [collapsed, setCollapsed] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showPinModal, setShowPinModal] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const { state, login, logout } = useApp();
  const { isAdmin, settings } = state;

  const doLogin = () => { if (login(passwordInput)) { setShowPinModal(false); setPasswordInput(''); } };

  const today = new Date();
  const isDashboard = pathname === '/';
  const pageTitle = PAGE_TITLES[pathname] || '대시보드';
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
        <nav className="sidebar-nav" aria-label="주 메뉴">
          {NAV_ITEMS.map(item => {
            if (item.path === '/attendance' && !isAdmin) return null;
            const isAllowed = isAdmin || STUDENT_ALLOWED_PATHS.includes(item.path);
            const isActive = pathname === item.path;
            return (
              <a key={item.path} href={`#${item.path}`} aria-current={isActive ? 'page' : undefined} className={`nav-item ${isActive ? 'active' : ''} ${!isAllowed ? 'nav-disabled' : ''}`} onClick={() => setSidebarOpen(false)}>
                <span className="nav-icon">{item.icon}</span>
                <span className="nav-label">{item.label}</span>
                {!isAllowed && <span className="nav-lock">선생님</span>}
              </a>
            );
          })}
        </nav>
      </aside>

      <main className="main-content">
        <header className="page-header">
          <button type="button" className="menu-btn" aria-label="메뉴 열기" onClick={() => setSidebarOpen(true)}><Icon name="menu" /></button>
          <div className="header-content">
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
        {NAV_ITEMS.map(item => {
          if (item.path === '/attendance' && !isAdmin) return null;
          const isAllowed = isAdmin || STUDENT_ALLOWED_PATHS.includes(item.path);
          const isActive = pathname === item.path;
          return (
            <a key={item.path} href={`#${item.path}`} aria-current={isActive ? 'page' : undefined} className={`mobile-nav-item ${isActive ? 'active' : ''} ${!isAllowed ? 'nav-disabled' : ''}`}>
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
  return settleLayout(merged, null);
}

function DashboardGrid({ layout, editing, onChange, widgets }) {
  const boxRef = useRef(null);
  const dragRef = useRef(null);
  const [width, setWidth] = useState(0);
  const [draft, setDraft] = useState(null);
  const [activeId, setActiveId] = useState(null);

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

  const apply = (id, mode, start, dc, dr) => {
    const cand = clampItem(mode === 'move' ? { ...start, x: start.x + dc, y: start.y + dr } : { ...start, w: start.w + dc, h: start.h + dr });
    return settleLayout(layout.map(i => (i.id === id ? cand : i)), id);
  };

  const onPointerDown = (e, id, mode) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.currentTarget.setPointerCapture(e.pointerId);
    dragRef.current = { id, mode, sx: e.clientX, sy: e.clientY, start: layout.find(i => i.id === id) };
    setActiveId(id);
  };
  const onPointerMove = (e) => {
    const d = dragRef.current;
    if (!d) return;
    setDraft(apply(d.id, d.mode, d.start, Math.round((e.clientX - d.sx) / unitX), Math.round((e.clientY - d.sy) / unitY)));
  };
  const onPointerUp = () => {
    if (!dragRef.current) return;
    dragRef.current = null;
    setActiveId(null);
    if (draft) onChange(draft);
    setDraft(null);
  };
  // 키보드: 화살표 = 이동, Shift+화살표 = 크기 조절
  const onKeyDown = (e, id) => {
    const delta = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
    if (!delta) return;
    e.preventDefault();
    onChange(apply(id, e.shiftKey ? 'resize' : 'move', layout.find(i => i.id === id), delta[0], delta[1]));
  };

  const ordered = [...shown].sort((a, b) => a.y - b.y || a.x - b.x);
  const rows = Math.max(...shown.map(i => i.y + i.h));

  return (
    <div ref={boxRef} className={`dgrid ${stacked ? 'is-stacked' : ''} ${editing && !stacked ? 'is-editing' : ''}`}
      style={stacked || !width ? undefined : { height: rows * unitY - GRID_GAP }}>
      {ordered.map(it => {
        const w = widgets[it.id];
        if (!w) return null;
        const style = stacked || !width ? undefined : {
          left: it.x * unitX, top: it.y * unitY,
          width: it.w * colW + (it.w - 1) * GRID_GAP, height: it.h * GRID_ROW_H + (it.h - 1) * GRID_GAP,
        };
        return (
          <section key={it.id} className={`wcard ${w.className || ''} ${activeId === it.id ? 'is-active' : ''}`} style={style} aria-label={w.label}>
            {w.content}
            {editing && !stacked && (
              <>
                <button type="button" className="grid-handle grid-handle--move" aria-label={`${w.label} 옮기기 (화살표 키로 이동, Shift+화살표로 크기 조절)`}
                  onPointerDown={e => onPointerDown(e, it.id, 'move')} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={onPointerUp}
                  onKeyDown={e => onKeyDown(e, it.id)}><Icon name="move" size={16} />옮기기</button>
                <button type="button" className="grid-handle grid-handle--resize" aria-label={`${w.label} 크기 조절`} tabIndex={-1}
                  onPointerDown={e => onPointerDown(e, it.id, 'resize')} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={onPointerUp}><Icon name="resize" size={14} /></button>
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
  const ddaySorted = [...ddays].sort((a, b) => new Date(a.date) - new Date(b.date));
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
                  <li key={c.period} className={`period is-${status}`} aria-current={status === 'now' ? 'time' : undefined}>
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
          {ddaySorted.length === 0 ? <div className="wcard-empty">등록된 D-Day가 없어요.</div> : (
            <ul className="dday-rows">
              {ddaySorted.map((ev, i) => (
                <li key={ev.id || i} className={`dday-row ${ev === firstUpcoming ? 'is-next' : ''}`}>
                  <span className="dday-num num">{getDday(ev.date)}</span>
                  <span className="dday-text"><span className="dday-title">{ev.emoji ? `${ev.emoji} ` : ''}{ev.name}</span><span className="dday-when">{shortDate(ev.date)}</span></span>
                  {isAdmin && <button type="button" className="icon-ghost" aria-label={`${ev.name} 삭제`} onClick={() => deleteDDay(ev.id)}><Icon name="trash" size={16} /></button>}
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
  const handleDragOverTeam = (e) => { if (!isAdmin) return; e.preventDefault(); e.currentTarget.style.borderColor = 'var(--primary)'; e.currentTarget.style.background = 'var(--primary-50)'; };
  const handleDragLeaveTeam = (e) => { if (!isAdmin) return; e.currentTarget.style.borderColor = 'var(--glass-border)'; e.currentTarget.style.background = 'var(--bg-card)'; };
  const handleDropTeam = (e, targetGroupId) => {
    if (!isAdmin) return; e.preventDefault();
    e.currentTarget.style.borderColor = 'var(--glass-border)'; e.currentTarget.style.background = 'var(--bg-card)';
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
      return { ...existing, id: existing ? existing.id : generateUUID(), number: i + 1, name: parts[0] || `학생${i + 1}`, gender, birthday: parts[2] || (existing ? existing.birthday : ''), points: existing ? existing.points : 0, group: existing ? existing.group : Math.ceil((i + 1) / 4) };
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
        return { ...existing, id: existing ? existing.id : generateUUID(), number: i + 1, name: row['이름'] || `학생${i + 1}`, gender, birthday: row['생일'] || (existing ? existing.birthday : ''), parentPhone: row['연락처'] || (existing ? existing.parentPhone : ''), points: existing ? existing.points : 0, group: existing ? existing.group : Math.ceil((i + 1) / 4) };
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

  return (
    <div className="students-page">
      <div className="students-tabs">
        <button className={`tab-btn ${activeTab === 'list' ? 'active' : ''}`} onClick={() => setActiveTab('list')}>👥 학생 목록</button>
        <button className={`tab-btn ${activeTab === 'behavior' ? 'active' : ''} ${!isAdmin ? 'tab-locked' : ''}`} onClick={() => isAdmin ? setActiveTab('behavior') : showToast('관리자 전용 기능입니다.', 'error')}>🌱 행동발달 {!isAdmin && <span className="lock-icon">🔒</span>}</button>
        <button className={`tab-btn ${activeTab === 'evaluation' ? 'active' : ''} ${!isAdmin ? 'tab-locked' : ''}`} onClick={() => isAdmin ? setActiveTab('evaluation') : showToast('관리자 전용 기능입니다.', 'error')}>📝 학습평가 {!isAdmin && <span className="lock-icon">🔒</span>}</button>
      </div>

      {activeTab === 'list' && (
        <>
          <div className="students-toolbar">
            <div className="search-wrapper">
              <span>🔍</span>
              <input className="search-input" placeholder="학생 이름 또는 번호 검색..." value={search} onChange={e => setSearch(e.target.value)} />
            </div>
            <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
              {isAdmin && (
                <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
                  <button className="btn-admin-action btn-admin-outline" onClick={downloadStudentTemplate}>📥 양식 다운로드</button>
                  <label className="btn-admin-action btn-admin-excel" style={{ cursor: 'pointer', margin: 0 }}>📤 엑셀 업로드<input type="file" accept=".xlsx, .xls" style={{ display: 'none' }} onChange={handleExcelUpload} /></label>
                  <button className="btn-admin-action btn-admin-primary" onClick={() => { const existingNames = students.map(s => `${s.name}\t${s.gender || ''}\t${s.birthday || ''}`).join('\n'); setBulkNames(existingNames); setShowBulkModal(true); }}>📝 일괄 수정</button>
                </div>
              )}
              <div className="view-toggle">
                {[{ v: 'card', icon: '▦' }, { v: 'table', icon: '☰' }, { v: 'seating', icon: '⊞' }, { v: 'points', icon: '⭐' }].map(b => (
                  <button key={b.v} className={`view-toggle-btn ${view === b.v ? 'active' : ''}`} onClick={() => setView(b.v)}>{b.icon}</button>
                ))}
              </div>
            </div>
          </div>

          {view === 'card' && (
            <div className="students-grid">
              {filtered.map(s => (
                <div key={s.id} className="student-card" onClick={() => handleSelectStudent(s)}>
                  <div className="student-card-header">
                    <div className={`avatar ${getAvatarColor(s.gender)}`}>{getInitials(s.name)}</div>
                    <div><div className="student-number">{s.number}번</div><div className="student-name-main">{s.name}</div></div>
                  </div>
                  <div className="student-card-body">
                    <div className="student-meta"><span>{s.gender === 'M' ? '👦' : '👧'} {s.birthday?.slice(5) || ''}</span></div>
                    {s.allergies && <div className="student-meta allergy">⚠️ {s.allergies}</div>}
                    {s.healthNotes && <div className="student-meta">🏥 {s.healthNotes}</div>}
                  </div>
                  <div className="student-card-footer">
                    <span className="student-points">⭐ {s.points}점</span>
                    {isAdmin && (<div className="student-actions"><button className="point-add plus" onClick={e => { e.stopPropagation(); handleAddPoints(s.id, s.points, 1); }}>+</button><button className="point-add minus" onClick={e => { e.stopPropagation(); handleAddPoints(s.id, s.points, -1); }}>-</button></div>)}
                  </div>
                </div>
              ))}
            </div>
          )}

          {view === 'table' && (
            <div className="table-container">
              <table>
                <thead><tr><th>번호</th><th>이름</th><th>성별</th><th>생일</th><th>알레르기</th><th>건강</th><th>포인트</th></tr></thead>
                <tbody>
                  {filtered.map(s => (
                    <tr key={s.id} onClick={() => handleSelectStudent(s)} style={{ cursor: 'pointer' }}>
                      <td>{s.number}</td><td><div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}><div className={`avatar sm ${getAvatarColor(s.gender)}`}>{getInitials(s.name)}</div>{s.name}</div></td>
                      <td>{s.gender === 'M' ? '남' : '여'}</td><td>{s.birthday?.slice(5) || '-'}</td><td>{s.allergies || '-'}</td><td>{s.healthNotes || '-'}</td><td><span className="student-points">⭐ {s.points}</span></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {view === 'seating' && (
            <div className="seating-container">
              <div className="seating-header"><h3>📐 자리 배치도</h3>{isAdmin && <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', marginLeft: '1rem' }}>💡 이름표를 드래그 앤 드롭하여 자리를 바꿀 수 있습니다.</span>}</div>
              <div className="seating-board">
                <div className="teacher-desk">교탁</div>
                <div className="seating-grid-dynamic">
                  {seatingList.map((s) => (
                    <div key={s.id} className="seat occupied" draggable={isAdmin} onDragStart={(e) => handleDragStartSeat(e, s)} onDragEnd={handleDragEndSeat} onDragOver={handleDragOverSeat} onDrop={(e) => handleDropSeat(e, s)} onClick={() => handleSelectStudent(s)} style={{ cursor: isAdmin ? 'grab' : 'pointer' }}>
                      <div className="seat-name">{s.name}</div><div className="seat-number">{s.number}번</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {view === 'points' && (
            <div className="points-section">
              <div className="points-header" style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '1rem', alignItems: 'center' }}>
                <h3 style={{ margin: 0 }}>🏆 칭찬 포인트 랭킹</h3>
                <div className="points-tabs" style={{ display: 'flex', gap: '0.5rem' }}>
                  <button className={`admin-btn ${pointsView === 'individual' ? 'active' : ''}`} style={pointsView === 'individual' ? { background: 'var(--primary)', color: 'white', borderColor: 'var(--primary)' } : {}} onClick={() => setPointsView('individual')}>개인별</button>
                  <button className={`admin-btn ${pointsView === 'group' ? 'active' : ''}`} style={pointsView === 'group' ? { background: 'var(--primary)', color: 'white', borderColor: 'var(--primary)' } : {}} onClick={() => setPointsView('group')}>팀별</button>
                </div>
              </div>
              {pointsView === 'individual' ? (
                <div className="points-leaderboard">
                  {sorted.map((s, i) => (
                    <div key={s.id} className="points-row glass-card" style={{ background: 'var(--bg-card)', border: '1px solid var(--glass-border)', borderRadius: 'var(--radius-md)', padding: '0.75rem 1rem', marginBottom: '0.5rem', display: 'flex', alignItems: 'center', gap: '1rem' }}>
                      <div className={`points-rank ${i === 0 ? 'gold' : i === 1 ? 'silver' : i === 2 ? 'bronze' : ''}`}>{i + 1}</div>
                      <div className={`avatar sm ${getAvatarColor(s.gender)}`}>{getInitials(s.name)}</div>
                      <div className="points-student-name" style={{ flex: 1, fontWeight: 600 }}>{s.name} <span style={{ fontSize: '0.75rem', color: 'var(--text-secondary)', marginLeft: '0.5rem', fontWeight: 500 }}>{s.group}팀</span></div>
                      <div className="points-score" style={{ fontWeight: 800, fontSize: '1.1rem', color: 'var(--primary)' }}>⭐ {s.points}</div>
                      {isAdmin && (<div className="points-actions" style={{ display: 'flex', gap: '0.25rem' }}><button className="btn btn-primary" onClick={() => handleAddPoints(s.id, s.points, 5)}>+5점</button><button className="btn" style={{ background: 'var(--success)', color: 'white' }} onClick={() => handleAddPoints(s.id, s.points, 10)}>+10점</button><button className="btn btn-danger" onClick={() => handleAddPoints(s.id, s.points, -5)}>-5점</button></div>)}
                    </div>
                  ))}
                </div>
              ) : (
                <div className="points-leaderboard group-leaderboard">
                  <div style={{ marginBottom: '1rem', fontSize: '0.85rem', color: 'var(--text-secondary)' }}>{isAdmin ? '💡 학생 이름표를 드래그하여 다른 팀으로 옮겨보세요.' : ''}</div>
                  {groupPoints.map((g, i) => (
                    <div key={g.groupId} className="points-row glass-card group-card-dnd" onDragOver={handleDragOverTeam} onDragLeave={handleDragLeaveTeam} onDrop={(e) => handleDropTeam(e, g.groupId)} style={{ background: 'var(--bg-card)', border: '2px solid var(--glass-border)', borderRadius: 'var(--radius-lg)', padding: '1.25rem', marginBottom: '0.75rem', display: 'flex', alignItems: 'flex-start', gap: '1.5rem' }}>
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '0.5rem' }}>
                        <div className={`points-rank ${i === 0 ? 'gold' : i === 1 ? 'silver' : i === 2 ? 'bronze' : ''}`} style={{ fontSize: '1.2rem', width: '30px', height: '30px', textAlign: 'center' }}>{i + 1}</div>
                        <div className="group-avatar" style={{ width: '40px', height: '40px', borderRadius: 'var(--radius-full)', background: 'var(--primary-50)', color: 'var(--primary-dark)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800, fontSize: '1.2rem' }}>{g.groupId}</div>
                        <div className="points-score" style={{ fontWeight: 800, fontSize: '1.1rem', color: 'var(--primary-dark)' }}>⭐ {g.totalPoints}</div>
                      </div>
                      <div className="points-student-name" style={{ flex: 1 }}>
                        <div style={{ fontWeight: 800, fontSize: '1.1rem', marginBottom: '0.75rem', color: 'var(--text-primary)' }}>{g.groupId}팀</div>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.5rem' }}>
                          {g.students.map(s => (
                            <div key={s.id} className="team-student-pill" draggable={isAdmin} onDragStart={(e) => handleDragStartTeam(e, s)} onDragEnd={handleDragEndTeam} style={{ padding: '0.4rem 0.75rem', background: 'var(--gray-50)', border: '1px solid var(--border-light)', borderRadius: 'var(--radius-xl)', fontSize: '0.85rem', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '0.4rem', cursor: isAdmin ? 'grab' : 'default' }}>
                              <div className={`avatar sm ${getAvatarColor(s.gender)}`} style={{ width: '20px', height: '20px', fontSize: '0.6rem' }}>{getInitials(s.name)}</div>{s.name}
                            </div>
                          ))}
                          {g.students.length === 0 && <span style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>편성된 학생이 없습니다. 드래그하여 추가하세요.</span>}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </>
      )}

      {activeTab === 'behavior' && (
        <div className="behavior-tab">
          <div className="behavior-grid" style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
            {students.map(s => {
              const recs = studentRecords[s.id] || [];
              return (
                <div key={s.id} className="behavior-student-card">
                  <div style={{ width: '160px', flexShrink: 0, textAlign: 'center', borderRight: '1px solid var(--border-light)', paddingRight: '1rem', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                    <div className={`avatar xl ${getAvatarColor(s.gender)}`} style={{ margin: '0 auto 0.75rem auto' }}>{getInitials(s.name)}</div>
                    <h3 style={{ margin: '0 0 0.5rem 0' }}>{s.number}. {s.name}</h3>
                    {isAdmin && <button className="btn-ai-generate" onClick={() => handleGenerateAI(s)} disabled={generatingAiId === s.id}>{generatingAiId === s.id ? '⏳ 생성 중...' : '✨ AI 생성'}</button>}
                    {isAdmin && <button className="btn btn-sm btn-secondary" onClick={() => handleSelectStudent(s)} style={{ marginTop: '0.5rem', width: '100%', fontWeight: 700 }}>➕ 기록 추가</button>}
                  </div>
                  <div style={{ flex: 1, display: 'flex', flexDirection: 'column' }}>
                    <h4 style={{ marginBottom: '0.5rem', color: 'var(--text-secondary)', display: 'flex', justifyContent: 'space-between', fontWeight: 700 }}><span>📝 누적 기록</span><span className="badge">{recs.length}건</span></h4>
                    <div className="behavior-record-box">
                      {recs.length > 0 ? (
                        <ul style={{ paddingLeft: '1.25rem', margin: 0, fontSize: '0.875rem', display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                          {recs.map((r, i) => (<li key={i}><strong style={{ color: 'var(--primary)' }}>[{r.date}]</strong> {r.content}{isAdmin && <button style={{ background: 'transparent', border: 'none', color: '#DC2626', fontSize: '0.75rem', cursor: 'pointer', marginLeft: '0.5rem', fontWeight: 600 }} onClick={() => { if (confirm('이 기록을 삭제하시겠습니까?')) deleteStudentRecord(s.id, r.id); }}>🗑️</button>}</li>))}
                        </ul>
                      ) : (<div className="empty-state" style={{ minHeight: '100px', padding: 0 }}>기록이 없습니다.</div>)}
                    </div>
                  </div>
                  <div style={{ flex: 1, display: 'flex', flexDirection: 'column' }}>
                    <h4 style={{ marginBottom: '0.5rem', color: 'var(--primary)', display: 'flex', alignItems: 'center', gap: '0.5rem', fontWeight: 800 }}>✨ 종합의견 (AI)</h4>
                    <div className="behavior-ai-box">
                      {s.aiSummary ? (<div style={{ whiteSpace: 'pre-wrap' }}>{s.aiSummary}</div>) : (<div className="empty-state" style={{ minHeight: '100px', padding: 0, opacity: 0.6, fontSize: '0.85rem' }}>AI 생성 버튼을 눌러 종합의견을 작성해보세요.</div>)}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {activeTab === 'evaluation' && (
        <div className="evaluation-tab">
          {assignments.length === 0 ? (<div className="empty-state"><p>등록된 과제 및 수행평가가 없습니다.</p></div>) : (
            <div className="table-container" style={{ overflowX: 'auto', background: 'var(--bg-card)', borderRadius: 'var(--radius-lg)' }}>
              <table style={{ minWidth: '100%', whiteSpace: 'nowrap' }}>
                <thead><tr><th style={{ position: 'sticky', left: 0, zIndex: 10, background: 'var(--gray-50)', minWidth: '120px' }}>학생 이름</th>
                  {assignments.map(a => (<th key={a.id} style={{ minWidth: '160px', textAlign: 'center', padding: '1rem' }}>
                    <div style={{ display: 'inline-block', padding: '0.2rem 0.5rem', background: a.type === '수행평가' ? '#FFF4ED' : '#ECFDF5', color: a.type === '수행평가' ? '#E8590C' : '#059669', borderRadius: '4px', fontSize: '0.75rem', marginBottom: '0.4rem' }}>{a.type}</div>
                    <div style={{ fontSize: '0.9rem', color: 'var(--text-primary)', fontWeight: 700, marginBottom: '0.25rem', whiteSpace: 'normal', wordBreak: 'keep-all' }}>{a.title}</div>
                    <div style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>{a.dueDate}</div>
                    {isAdmin && <button style={{ background: 'transparent', border: 'none', color: '#DC2626', fontSize: '0.75rem', cursor: 'pointer', marginTop: '0.5rem', fontWeight: 600 }} onClick={() => { if (confirm('이 항목을 정말 삭제하시겠습니까?')) { dispatch({ type: 'DELETE_ASSIGNMENT', payload: a.id }); showToast('삭제되었습니다.', 'success'); } }}>🗑️ 삭제</button>}
                  </th>))}
                </tr></thead>
                <tbody>
                  {students.map(s => (
                    <tr key={s.id}>
                      <td style={{ position: 'sticky', left: 0, zIndex: 5, background: 'var(--bg-card)', fontWeight: 600, borderRight: '1px solid var(--border-light)' }}>{s.number}번 <span style={{ marginLeft: '0.25rem' }}>{s.name}</span></td>
                      {assignments.map(a => {
                        const submission = a.submissions?.[s.id];
                        const grade = submission?.grade;
                        return (
                          <td key={`${s.id}-${a.id}`} style={{ textAlign: 'center', background: submission ? 'transparent' : '#FEF2F2' }}>
                            {submission ? (isAdmin ? (
                              <div style={{ display: 'flex', gap: '0.25rem', justifyContent: 'center' }}>
                                {['상', '중', '하'].map(g => (<button key={g} className={`grade-btn ${grade === g ? 'active' : ''}`} style={{ padding: '4px 10px', borderRadius: '4px', border: '1px solid var(--border-light)', background: grade === g ? 'var(--primary)' : 'white', color: grade === g ? 'white' : 'var(--text-secondary)', cursor: 'pointer', fontWeight: 600, fontSize: '0.75rem' }} onClick={() => handleEvaluate(a.id, s.id, g)}>{g}</button>))}
                              </div>
                            ) : (<span style={{ display: 'inline-block', padding: '0.25rem 0.75rem', background: '#D1FAE5', color: '#065F46', borderRadius: '4px', fontWeight: 700, fontSize: '0.8rem' }}>{grade || '채점 대기'}</span>)
                            ) : (<span style={{ color: '#DC2626', fontSize: '0.8rem', fontWeight: 600 }}>미제출</span>)}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
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
                        {isAdmin && <button className="edit-name-btn" onClick={() => { setEditNameValue(selectedStudent.name); setIsEditingName(true); }} style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: 'var(--text-muted)', fontSize: '1.2rem' }}>✎</button>}
                      </div>
                    )}
                  </div>
                  <p>{selectedStudent.number}번 · {selectedStudent.gender === 'M' ? '남학생' : '여학생'} · 생일: {selectedStudent.birthday}</p>
                </div>
              </div>
              {isAdmin && (
                <div className="student-edit-section">
                  <h4 style={{ fontSize: '0.875rem', fontWeight: 700, marginBottom: '0.75rem' }}>📝 정보 수정</h4>
                  <div className="edit-fields-grid">
                    <div className="edit-field"><label>성별</label><select className="form-input" value={editGender} onChange={e => setEditGender(e.target.value)}><option value="M">남</option><option value="F">여</option></select></div>
                    <div className="edit-field"><label>생일</label><input type="date" className="form-input" value={editBirthday} onChange={e => setEditBirthday(e.target.value)} /></div>
                    <div className="edit-field"><label>연락처</label><input type="text" className="form-input" value={editParentPhone} onChange={e => setEditParentPhone(e.target.value)} placeholder="010-0000-0000" /></div>
                    <div className="edit-field" style={{ display: 'flex', gap: '0.5rem', alignItems: 'flex-end' }}><button className="btn btn-primary btn-sm" onClick={handleSaveStudentInfo}>저장</button><button className="btn btn-danger btn-sm" onClick={handleResetPoints}>포인트 초기화</button></div>
                  </div>
                </div>
              )}
              <div style={{ marginBottom: '0.75rem' }}><strong style={{ fontSize: '0.875rem' }}>📞 학부모 연락처:</strong><span style={{ marginLeft: '0.5rem', fontSize: '0.875rem' }}>{selectedStudent.parentPhone || '미등록'}</span></div>
              {isAdmin && (
                <div className="record-input-section">
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.75rem' }}><h4 style={{ fontSize: '0.875rem', fontWeight: 700, margin: 0 }}>📋 관찰/상담 기록 추가</h4></div>
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
                  <h4 style={{ fontSize: '0.875rem', fontWeight: 700, marginBottom: '0.75rem' }}>📖 기록 목록 ({records.length}건)</h4>
                  <div className="counseling-list" style={{ marginTop: '0.5rem' }}>
                    {records.map((c, i) => (
                      <div key={c.id || i} className="counseling-item">
                        <div className="counseling-date">{c.date} · <span className={`record-type-badge ${c.type}`}>{c.type === 'observation' ? '관찰' : c.type === 'counseling' ? '상담' : c.type}</span> · {c.category}
                          {isAdmin && <button style={{ background: 'transparent', border: 'none', color: '#DC2626', fontSize: '0.75rem', cursor: 'pointer', marginLeft: '0.5rem', fontWeight: 600 }} onClick={() => { if (confirm('이 기록을 삭제하시겠습니까?')) deleteStudentRecord(selectedStudent.id, c.id); }}>🗑️ 삭제</button>}
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
            <div className="modal-header"><h3 className="modal-title">📝 학생 이름 일괄 수정</h3><button className="modal-close" onClick={() => setShowBulkModal(false)}>&times;</button></div>
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
const sameAtt = (a, b) => (a?.status || 'present') === (b?.status || 'present') && (a?.note || '') === (b?.note || '');

function AttendancePage() {
  const { state, updateAttendance, loadAttendanceByDate, loadMonthlyAttendance, saveAttendanceDay, showToast } = useApp();
  const { students, attendance, selectedAttendanceDate, monthlyAttendance, isAdmin } = state;
  const [monthlyActiveCell, setMonthlyActiveCell] = useState(null);
  const [viewMode, setViewMode] = useState('daily');
  const [monthYear, setMonthYear] = useState({ year: new Date().getFullYear(), month: new Date().getMonth() });
  // 입력 중인 출석(저장 전). 완료 버튼을 눌러야 db에 반영됨
  const [draft, setDraft] = useState({});
  const [selected, setSelected] = useState(() => new Set());
  const [noteFor, setNoteFor] = useState(null);
  const [noteText, setNoteText] = useState('');

  const baseOf = (id) => attendance[id] || { status: 'present', note: '' };
  const current = (id) => draft[id] || baseOf(id);
  const changedIds = students.filter(s => draft[s.id] && !sameAtt(draft[s.id], baseOf(s.id))).map(s => s.id);
  const dirty = changedIds.length > 0;

  useEffect(() => { setDraft({}); setSelected(new Set()); setNoteFor(null); }, [selectedAttendanceDate]);
  useEffect(() => {
    if (!dirty) return;
    const warn = (e) => { e.preventDefault(); e.returnValue = ''; };
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, [dirty]);

  const handleDateChange = (e) => {
    if (dirty && !confirm('저장하지 않은 출석 변경이 있어요. 날짜를 바꾸면 사라집니다. 계속할까요?')) return;
    loadAttendanceByDate(e.target.value);
  };
  const handleMonthlyStatus = (studentId, date, status) => {
    if (!isAdmin) { showToast('선생님 모드에서만 수정할 수 있습니다.', 'error'); return; }
    updateAttendance(studentId, status, monthlyAttendance[date]?.[studentId]?.note || '', date);
    const s = students.find(st => st.id === studentId);
    showToast(`${s.name} (${date.slice(5).replace('-', '/')}) → ${STATUS_MAP[status].label}`, 'success');
    setMonthlyActiveCell(null);
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
      selected.forEach(id => { const cur = next[id] || baseOf(id); next[id] = { status, note: status === 'present' ? '' : cur.note || '' }; });
      return next;
    });
    setSelected(new Set());
  };
  const openNote = (id) => { setNoteFor(id); setNoteText(current(id).note || ''); };
  const saveNote = () => { setDraft(prev => ({ ...prev, [noteFor]: { ...current(noteFor), note: noteText.trim() } })); setNoteFor(null); };

  const discard = () => { setDraft({}); setSelected(new Set()); setNoteFor(null); };
  const handleSubmit = () => {
    if (!isAdmin) return;
    const records = {};
    students.forEach(s => { const a = current(s.id); records[s.id] = { status: a.status, note: a.status === 'present' ? '' : (a.note || ''), id: baseOf(s.id).id }; });
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
          <button type="button" aria-pressed={viewMode === 'daily'} onClick={() => setViewMode('daily')}>일별 출석</button>
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
                        ? <button type="button" className={`att-note-btn ${a.note ? 'has-note' : ''}`} onClick={() => openNote(s.id)}><Icon name="edit" size={14} />{a.note || '사유 쓰기'}</button>
                        : a.note ? <span className="att-note-btn has-note">{a.note}</span> : null
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
                  <label htmlFor="att-note-text" className="sr-only">사유</label>
                  <textarea id="att-note-text" className="form-input" rows="3" autoFocus placeholder="예: 감기로 병원 진료" value={noteText} onChange={e => setNoteText(e.target.value)} />
                  <div className="att-note-presets">{['질병', '병원 진료', '가정 사정', '체험학습', '늦잠'].map(p => <button key={p} type="button" className="chip-btn" onClick={() => setNoteText(p)}>{p}</button>)}</div>
                </div>
                <div className="modal-footer att-note-footer"><button type="button" className="btn btn-secondary" onClick={() => setNoteFor(null)}>취소</button><button type="button" className="btn btn-primary" onClick={saveNote}>사유 적기</button></div>
              </div>
            </div>
          )}
        </>
      ) : (
        <div className="monthly-attendance">
          <div className="wk-nav att-month-nav">
            <button type="button" className="round-btn" aria-label="이전 달" onClick={() => { if (monthYear.month === 0) setMonthYear({ year: monthYear.year - 1, month: 11 }); else setMonthYear({ ...monthYear, month: monthYear.month - 1 }); }}><Icon name="chevron" size={18} /></button>
            <button type="button" className="round-btn wk-next" aria-label="다음 달" onClick={() => { if (monthYear.month === 11) setMonthYear({ year: monthYear.year + 1, month: 0 }); else setMonthYear({ ...monthYear, month: monthYear.month + 1 }); }}><Icon name="chevron" size={18} /></button>
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
                      const status = record?.status || null;
                      const isCellActive = monthlyActiveCell?.studentId === s.id && monthlyActiveCell?.date === d;
                      return (
                        <td key={d} className={`monthly-cell ${status || 'no-data'} ${isCellActive ? 'active' : ''}`} title={record?.note || ''} onClick={() => { if (isAdmin) setMonthlyActiveCell(isCellActive ? null : { studentId: s.id, date: d }); }} style={{ position: 'relative', cursor: isAdmin ? 'pointer' : 'default' }}>
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
        </div>
      )}
    </div>
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
      const srcMatch = input.match(/src="([^"]+)"/);
      let extracted = srcMatch ? srcMatch[1] : input.trim();
      let calId = '';
      if (extracted.includes('calendar.google.com') && extracted.includes('src=')) { const m = extracted.match(/src=([^&]+)/); if (m) calId = decodeURIComponent(m[1]); }
      else calId = extracted;
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

  return (
    <div className="tools-page">
      <div className="tools-grid">
        <div className={`tool-card ${!isAdmin ? 'disabled-tool' : ''}`}>
          <div className="tool-card-header"><span className="tool-emoji">🎲</span><div><div className="tool-title">랜덤 뽑기</div><div className="tool-desc">발표자를 랜덤으로 선택해요</div></div></div>
          <div className="picker-container">
            <div className="roulette-wheel"><div className="roulette-pointer" /><div className={`roulette-circle ${isSpinning ? 'spinning' : pickerResult !== '?' ? 'reveal' : ''}`}><div className="roulette-result">{pickerResult}</div></div></div>
            {isAdmin && (<div className="picker-count"><label>뽑을 인원:</label><input type="number" min="1" max="28" value={pickerCount} onChange={e => setPickerCount(Math.max(1, parseInt(e.target.value) || 1))} /><span>명</span></div>)}
            {isAdmin && (<div style={{ display: 'flex', gap: '0.5rem', justifyContent: 'center' }}><button className="btn btn-primary btn-lg" onClick={handlePick} disabled={isSpinning}>{isSpinning ? '🎰 뽑는 중...' : '🎲 뽑기!'}</button><button className="btn btn-secondary" onClick={resetPicker}>초기화</button></div>)}
            {pickerHistory.length > 0 && (<div className="picker-history">{pickerHistory.map((s, i) => <span key={i} className="picked-badge">{s.name}</span>)}</div>)}
          </div>
        </div>

        <div className={`tool-card ${!isAdmin ? 'disabled-tool' : ''}`}>
          <div className="tool-card-header"><span className="tool-emoji">⏱️</span><div><div className="tool-title">타이머 / 스톱워치</div><div className="tool-desc">활동 시간을 관리해요</div></div></div>
          <div className="timer-container">
            <div className="timer-mode-toggle"><button className={`timer-mode-btn ${timerMode === 'timer' ? 'active' : ''}`} onClick={() => { resetTimer(); setTimerMode('timer'); }}>⏳ 타이머</button><button className={`timer-mode-btn ${timerMode === 'stopwatch' ? 'active' : ''}`} onClick={() => { resetTimer(); setTimerMode('stopwatch'); }}>⏱️ 스톱워치</button></div>
            <div className={`timer-display ${timerRunning ? 'running' : ''} ${isAlarm ? 'alarm' : ''}`}>{timerMode === 'timer' ? formatTimer(timerSeconds) : formatStopwatch(timerElapsed)}</div>
            {timerMode === 'timer' && (<>
              <div className="timer-presets">{presets.map(p => (<button key={p.sec} className={`timer-preset ${timerSeconds === p.sec && !timerRunning ? 'active' : ''}`} onClick={() => { if (!timerRunning) setTimerSeconds(p.sec); }}>{p.label}</button>))}</div>
              {!timerRunning && (<div className="timer-custom-input"><input type="number" className="custom-time-input" placeholder="분" min="0" max="99" value={customMin} onChange={e => setCustomMin(e.target.value)} /><span className="time-separator">:</span><input type="number" className="custom-time-input" placeholder="초" min="0" max="59" value={customSec} onChange={e => setCustomSec(e.target.value)} /><button className="btn btn-sm btn-primary" onClick={applyCustomTime}>적용</button></div>)}
            </>)}
            <div className="timer-controls">{!timerRunning ? <button className="btn btn-primary" onClick={startTimer}>▶ 시작</button> : <button className="btn btn-secondary" onClick={pauseTimer}>⏸ 일시정지</button>}<button className="btn btn-danger" onClick={resetTimer} style={{ background: 'var(--danger)', color: 'white' }}>↺ 초기화</button></div>
          </div>
        </div>

        <div className={`tool-card ${!isAdmin ? 'disabled-tool' : ''}`}>
          <div className="tool-card-header"><span className="tool-emoji">👥</span><div><div className="tool-title">팀 편성</div><div className="tool-desc">랜덤으로 팀을 나눠요</div></div></div>
          <div className="group-maker">
            {isAdmin && (<div className="group-settings"><div className="group-setting"><label>팀 수:</label><input type="number" min="2" max="14" value={groupCount} onChange={e => setGroupCount(Math.max(2, parseInt(e.target.value) || 2))} /><span>개</span></div><button className="btn btn-primary" onClick={makeGroups}>🔀 팀 편성하기</button></div>)}
            {groups.length > 0 && (<div className="groups-result">{groups.map((g, i) => (<div key={i} className="group-card" style={{ animationDelay: `${i * 0.1}s` }}><div className="group-card-title">{i + 1}팀</div><div className="group-members">{g.map(s => (<div key={s.id} className="group-member"><span>{s.gender === 'M' ? '👦' : '👧'}</span><span>{s.number}. {s.name}</span></div>))}</div></div>))}</div>)}
          </div>
        </div>

        <div className="tool-card">
          <div className="tool-card-header"><span className="tool-emoji">📊</span><div><div className="tool-title">투표 / 설문</div><div className="tool-desc">학급 의견을 모아요</div></div></div>
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
  const { weeklyPlans, isAdmin } = state;
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
              <button type="button" className={`wk-plan ${isEmpty ? 'is-empty' : ''}`} aria-label={`${dd.day}요일 ${p}교시 ${isEmpty ? '입력' : '수정'}`} onClick={() => handleEdit(dd.day, p, plan.subject, plan.content)}>
                {isEmpty ? <span className="wk-add"><Icon name="plus" size={14} />입력</span> : <><span className="wk-subject">{plan.subject}</span>{plan.content && <span className="wk-content">{plan.content}</span>}</>}
              </button>
            ) : (
              <div className={`wk-plan ${isEmpty ? 'is-empty' : ''}`}>{plan.subject && <span className="wk-subject">{plan.subject}</span>}{plan.content && <span className="wk-content">{plan.content}</span>}</div>
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
      <div className="wk-card">
        <table className="wk-table">
          <thead>
            <tr>
              <th scope="col" className="wk-corner">교시</th>
              {dayDates.map(dd => (
                <th scope="col" key={dd.day} className={dd.isoDate === todayStr ? 'is-today' : ''}>
                  <span className="wk-day">{dd.day}</span><span className="wk-date num">{dd.dateStr}</span>
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
