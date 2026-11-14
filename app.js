const {
  useState,
  useEffect,
  useRef,
  useCallback,
  useReducer,
  useContext,
  useMemo,
  createContext
} = React;
const {
  getDayName,
  formatDate,
  getDday,
  isTodayBirthday,
  getAvatarColor,
  getInitials,
  getCurrentPeriod,
  shuffleArray,
  generateUUID,
  getWeekStart,
  toISODate
} = window.Helpers;
const {
  fetchTodayLunch,
  fetchSchoolHoliday
} = window.NeisApi;
const {
  generateBehaviorSummary
} = window.GeminiApi;
const {
  downloadStudentTemplate,
  parseStudentExcel
} = window.ExcelApi;
const {
  pickFolder,
  getActiveFolder,
  reconnectFolder,
  readDb,
  writeDb
} = window.LocalStore;

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
  settings: {
    ...window.LocalStore.DEFAULT_DB.settings
  },
  ddays: [],
  studentRecords: {}
};
function reducer(state, action) {
  switch (action.type) {
    case 'SET_STUDENTS':
      return {
        ...state,
        students: action.payload
      };
    case 'SET_ANNOUNCEMENTS':
      return {
        ...state,
        announcements: action.payload
      };
    case 'SET_ASSIGNMENTS':
      return {
        ...state,
        assignments: action.payload
      };
    case 'SET_ATTENDANCE_MAP':
      return {
        ...state,
        attendance: action.payload
      };
    case 'SET_SELECTED_ATTENDANCE_DATE':
      return {
        ...state,
        selectedAttendanceDate: action.payload
      };
    case 'SET_MONTHLY_ATTENDANCE':
      {
        const newMonthly = {
          ...state.monthlyAttendance
        };
        for (const [date, records] of Object.entries(action.payload)) {
          newMonthly[date] = {
            ...newMonthly[date],
            ...records
          };
        }
        return {
          ...state,
          monthlyAttendance: newMonthly
        };
      }
    case 'SAVE_ATTENDANCE_DAY':
      {
        const {
          date,
          records
        } = action.payload;
        return {
          ...state,
          monthlyAttendance: {
            ...state.monthlyAttendance,
            [date]: {
              ...(state.monthlyAttendance[date] || {}),
              ...records
            }
          },
          attendance: date === state.selectedAttendanceDate ? {
            ...state.attendance,
            ...records
          } : state.attendance
        };
      }
    case 'UPDATE_MONTHLY_ATTENDANCE':
      {
        const {
          date,
          studentId,
          status,
          id
        } = action.payload;
        const dateRecords = state.monthlyAttendance[date] || {};
        return {
          ...state,
          monthlyAttendance: {
            ...state.monthlyAttendance,
            [date]: {
              ...dateRecords,
              [studentId]: {
                status,
                note: dateRecords[studentId]?.note || '',
                id: id || dateRecords[studentId]?.id
              }
            }
          }
        };
      }
    case 'SET_ATTENDANCE':
      return {
        ...state,
        attendance: {
          ...state.attendance,
          [action.payload.studentId]: {
            status: action.payload.status,
            note: action.payload.note || '',
            id: action.payload.id || state.attendance[action.payload.studentId]?.id
          }
        }
      };
    case 'UPDATE_POINTS':
      return {
        ...state,
        students: state.students.map(s => s.id === action.payload.studentId ? {
          ...s,
          points: Math.max(0, s.points + action.payload.amount)
        } : s)
      };
    case 'RESET_POINTS':
      return {
        ...state,
        students: state.students.map(s => s.id === action.payload.studentId ? {
          ...s,
          points: 0
        } : s)
      };
    case 'ADD_ANNOUNCEMENT':
      return {
        ...state,
        announcements: [action.payload, ...state.announcements]
      };
    case 'UPDATE_ANNOUNCEMENT':
      return {
        ...state,
        announcements: state.announcements.map(a => a.id === action.payload.id ? action.payload : a)
      };
    case 'DELETE_ANNOUNCEMENT':
      return {
        ...state,
        announcements: state.announcements.filter(a => a.id !== action.payload)
      };
    case 'ADD_ASSIGNMENT':
      return {
        ...state,
        assignments: [...state.assignments, action.payload]
      };
    case 'DELETE_ASSIGNMENT':
      return {
        ...state,
        assignments: state.assignments.filter(a => a.id !== action.payload)
      };
    case 'UPDATE_ASSIGNMENT_SUBMISSIONS':
      return {
        ...state,
        assignments: state.assignments.map(a => a.id === action.payload.id ? {
          ...a,
          submissions: {
            ...a.submissions,
            ...action.payload.submissions
          }
        } : a)
      };
    case 'EVALUATE_ASSIGNMENT':
      return {
        ...state,
        assignments: state.assignments.map(a => {
          if (a.id === action.payload.id) {
            const currentSub = a.submissions?.[action.payload.studentId] || {};
            return {
              ...a,
              submissions: {
                ...a.submissions,
                [action.payload.studentId]: {
                  ...currentSub,
                  grade: action.payload.grade
                }
              }
            };
          }
          return a;
        })
      };
    case 'ADD_DDAY':
      return {
        ...state,
        ddays: [...state.ddays, action.payload]
      };
    case 'DELETE_DDAY':
      return {
        ...state,
        ddays: state.ddays.filter(d => d.id !== action.payload)
      };
    case 'SET_DDAYS':
      return {
        ...state,
        ddays: action.payload
      };
    case 'ADD_POLL':
      return {
        ...state,
        polls: [...state.polls, action.payload]
      };
    case 'DELETE_POLL':
      return {
        ...state,
        polls: state.polls.filter(p => p.id !== action.payload)
      };
    case 'SET_POLLS':
      return {
        ...state,
        polls: action.payload
      };
    case 'VOTE_POLL':
      return {
        ...state,
        polls: state.polls.map(p => p.id === action.payload.pollId ? {
          ...p,
          votes: action.payload.newVotes
        } : p)
      };
    case 'ADD_TOAST':
      return {
        ...state,
        toasts: [...state.toasts, action.payload]
      };
    case 'REMOVE_TOAST':
      return {
        ...state,
        toasts: state.toasts.filter(t => t.id !== action.payload)
      };
    case 'ADD_STUDENT':
      return {
        ...state,
        students: [...state.students, action.payload]
      };
    case 'UPDATE_STUDENT':
      return {
        ...state,
        students: state.students.map(s => s.id === action.payload.id ? {
          ...s,
          ...action.payload
        } : s)
      };
    case 'BULK_UPDATE_STUDENTS':
      return {
        ...state,
        students: state.students.map(s => {
          const u = action.payload.find(x => x.id === s.id);
          return u ? {
            ...s,
            ...u
          } : s;
        })
      };
    case 'SET_NEWSLETTERS':
      return {
        ...state,
        newsletters: action.payload
      };
    case 'COLLECT_NEWSLETTER':
      return {
        ...state,
        newsletters: state.newsletters.map(n => n.id === action.payload.id ? {
          ...n,
          collected: action.payload.collected
        } : n)
      };
    case 'TOGGLE_ADMIN':
      return {
        ...state,
        isAdmin: action.payload
      };
    case 'UPDATE_SETTINGS':
      return {
        ...state,
        settings: {
          ...state.settings,
          ...action.payload
        }
      };
    case 'SET_WEEKLY_PLAN':
      {
        const {
          weekKey,
          day,
          period,
          subject,
          content
        } = action.payload;
        const existingWeek = state.weeklyPlans[weekKey] || {};
        return {
          ...state,
          weeklyPlans: {
            ...state.weeklyPlans,
            [weekKey]: {
              ...existingWeek,
              [day]: {
                ...(existingWeek[day] || {}),
                [period]: {
                  subject,
                  content
                }
              }
            }
          }
        };
      }
    case 'SET_WEEKLY_PLANS_BULK':
      return {
        ...state,
        weeklyPlans: action.payload
      };
    case 'SET_QUICK_LINKS':
      return {
        ...state,
        quickLinks: action.payload
      };
    case 'ADD_LINK':
      return {
        ...state,
        quickLinks: [...state.quickLinks, action.payload]
      };
    case 'UPDATE_LINK':
      return {
        ...state,
        quickLinks: state.quickLinks.map(l => l.id === action.payload.id ? action.payload : l)
      };
    case 'DELETE_LINK':
      return {
        ...state,
        quickLinks: state.quickLinks.filter(l => l.id !== action.payload)
      };
    case 'SET_STUDENT_RECORDS':
      return {
        ...state,
        studentRecords: action.payload
      };
    case 'ADD_STUDENT_RECORD':
      {
        const {
          studentId,
          record
        } = action.payload;
        const existing = state.studentRecords[studentId] || [];
        return {
          ...state,
          studentRecords: {
            ...state.studentRecords,
            [studentId]: [...existing, record]
          }
        };
      }
    case 'DELETE_STUDENT_RECORD':
      {
        const {
          studentId,
          recordId
        } = action.payload;
        const existing = state.studentRecords[studentId] || [];
        return {
          ...state,
          studentRecords: {
            ...state.studentRecords,
            [studentId]: existing.filter(r => r.id !== recordId)
          }
        };
      }
    default:
      return state;
  }
}
function AppProvider({
  children
}) {
  const [state, dispatch] = useReducer(reducer, initialState);
  const [isDbLoading, setIsDbLoading] = useState(true);
  const [folder, setFolder] = useState(null);
  const [folderStatus, setFolderStatus] = useState('loading'); // loading | none | needs-reconnect | connected
  const saveTimer = useRef(null);
  const showToast = useCallback((message, type = 'info') => {
    const id = Date.now() + Math.random();
    dispatch({
      type: 'ADD_TOAST',
      payload: {
        id,
        message,
        type
      }
    });
    setTimeout(() => dispatch({
      type: 'REMOVE_TOAST',
      payload: id
    }), 3000);
  }, []);
  const loadFromFolder = useCallback(async handle => {
    const db = await readDb(handle);
    dispatch({
      type: 'SET_STUDENTS',
      payload: db.students || []
    });
    dispatch({
      type: 'SET_ANNOUNCEMENTS',
      payload: db.announcements || []
    });
    dispatch({
      type: 'SET_ASSIGNMENTS',
      payload: db.assignments || []
    });
    dispatch({
      type: 'SET_DDAYS',
      payload: db.ddays || []
    });
    dispatch({
      type: 'SET_QUICK_LINKS',
      payload: db.quickLinks || []
    });
    dispatch({
      type: 'SET_POLLS',
      payload: db.polls || []
    });
    dispatch({
      type: 'SET_NEWSLETTERS',
      payload: db.newsletters || []
    });
    dispatch({
      type: 'SET_WEEKLY_PLANS_BULK',
      payload: db.weeklyPlans || {}
    });
    dispatch({
      type: 'SET_STUDENT_RECORDS',
      payload: db.studentRecords || {}
    });
    const attendanceByDate = db.attendance || {};
    dispatch({
      type: 'SET_MONTHLY_ATTENDANCE',
      payload: attendanceByDate
    });
    const today = toISODate(new Date());
    const attMap = {};
    (db.students || []).forEach(s => {
      attMap[s.id] = {
        status: 'present',
        note: ''
      };
    });
    Object.assign(attMap, attendanceByDate[today] || {});
    dispatch({
      type: 'SET_ATTENDANCE_MAP',
      payload: attMap
    });
    dispatch({
      type: 'UPDATE_SETTINGS',
      payload: db.settings || {}
    });
    setIsDbLoading(false);
  }, []);

  // 최초 마운트: 저장해둔 폴더가 있는지 확인
  useEffect(() => {
    (async () => {
      try {
        const result = await getActiveFolder();
        if (!result) {
          setFolderStatus('none');
          setIsDbLoading(false);
          return;
        }
        if (result.needsReconnect) {
          setFolder(result.handle);
          setFolderStatus('needs-reconnect');
          setIsDbLoading(false);
          return;
        }
        setFolder(result);
        setFolderStatus('connected');
        await loadFromFolder(result);
      } catch (e) {
        console.error('폴더 연결 확인 중 오류:', e);
        setFolderStatus('none');
        setIsDbLoading(false);
      }
    })();
    if (localStorage.getItem('isAdmin') === 'true') dispatch({
      type: 'TOGGLE_ADMIN',
      payload: true
    });
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
    if (handle) {
      setFolderStatus('connected');
      setIsDbLoading(true);
      await loadFromFolder(handle);
    } else showToast('폴더 접근이 거부되었습니다.', 'error');
  };

  // 상태가 바뀔 때마다 로컬 폴더에 저장 (디바운스)
  useEffect(() => {
    if (isDbLoading || folderStatus !== 'connected' || !folder) return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      const payload = {
        students: state.students,
        attendance: state.monthlyAttendance,
        announcements: state.announcements,
        assignments: state.assignments,
        ddays: state.ddays,
        polls: state.polls,
        newsletters: state.newsletters,
        quickLinks: state.quickLinks,
        weeklyPlans: state.weeklyPlans,
        studentRecords: state.studentRecords,
        settings: state.settings
      };
      writeDb(folder, payload).catch(err => console.error('로컬 저장 오류:', err));
    }, 500);
    return () => clearTimeout(saveTimer.current);
  }, [state, isDbLoading, folderStatus, folder]);
  const saveWeeklyPlan = useCallback((weekKey, day, period, subject, content) => {
    dispatch({
      type: 'SET_WEEKLY_PLAN',
      payload: {
        weekKey,
        day,
        period,
        subject,
        content
      }
    });
  }, []);
  const addDDay = dday => dispatch({
    type: 'ADD_DDAY',
    payload: {
      id: generateUUID(),
      name: dday.name,
      date: dday.date,
      emoji: dday.emoji
    }
  });
  const deleteDDayFunc = id => dispatch({
    type: 'DELETE_DDAY',
    payload: id
  });
  const updatePointsFunc = (studentId, _cur, amount) => dispatch({
    type: 'UPDATE_POINTS',
    payload: {
      studentId,
      amount
    }
  });
  const resetPointsFunc = studentId => dispatch({
    type: 'RESET_POINTS',
    payload: {
      studentId
    }
  });
  const updateAttendanceFunc = (studentId, status, note = '', date = null) => {
    const targetDate = date || toISODate(new Date());
    dispatch({
      type: 'SET_ATTENDANCE',
      payload: {
        studentId,
        status,
        note,
        id: ''
      }
    });
    dispatch({
      type: 'UPDATE_MONTHLY_ATTENDANCE',
      payload: {
        date: targetDate,
        studentId,
        status,
        id: ''
      }
    });
  };
  const loadAttendanceByDateFunc = useCallback(date => {
    dispatch({
      type: 'SET_SELECTED_ATTENDANCE_DATE',
      payload: date
    });
    const attMap = {};
    state.students.forEach(s => {
      attMap[s.id] = {
        status: 'present',
        note: ''
      };
    });
    Object.assign(attMap, state.monthlyAttendance[date] || {});
    dispatch({
      type: 'SET_ATTENDANCE_MAP',
      payload: attMap
    });
  }, [state.students, state.monthlyAttendance]);
  const loadMonthlyAttendanceFunc = useCallback(() => {}, []);
  const saveAttendanceDayFunc = (date, records) => dispatch({
    type: 'SAVE_ATTENDANCE_DAY',
    payload: {
      date,
      records
    }
  });
  const addStudentRecordFunc = (studentId, type, content, category = '') => {
    const date = toISODate(new Date());
    dispatch({
      type: 'ADD_STUDENT_RECORD',
      payload: {
        studentId,
        record: {
          id: generateUUID(),
          type,
          content,
          category,
          date
        }
      }
    });
    showToast('기록이 저장되었습니다.', 'success');
  };
  const deleteStudentRecordFunc = (studentId, recordId) => dispatch({
    type: 'DELETE_STUDENT_RECORD',
    payload: {
      studentId,
      recordId
    }
  });
  const updateStudentFunc = studentData => dispatch({
    type: 'UPDATE_STUDENT',
    payload: studentData
  });
  const bulkUpdateStudentsFunc = updates => dispatch({
    type: 'BULK_UPDATE_STUDENTS',
    payload: updates
  });
  const votePollFunc = (pollId, optionIndex) => {
    const poll = state.polls.find(p => p.id === pollId);
    if (!poll) return;
    const newVotes = [...poll.votes];
    newVotes[optionIndex] += 1;
    dispatch({
      type: 'VOTE_POLL',
      payload: {
        pollId,
        newVotes
      }
    });
  };
  const collectNewsletterFunc = (id, newCollected) => dispatch({
    type: 'COLLECT_NEWSLETTER',
    payload: {
      id,
      collected: newCollected
    }
  });
  const addLinkFunc = linkData => dispatch({
    type: 'ADD_LINK',
    payload: linkData
  });
  const updateLinkFunc = linkData => dispatch({
    type: 'UPDATE_LINK',
    payload: linkData
  });
  const deleteLinkFunc = id => dispatch({
    type: 'DELETE_LINK',
    payload: id
  });
  const addAssignmentFunc = assignment => dispatch({
    type: 'ADD_ASSIGNMENT',
    payload: assignment
  });
  const updateAssignmentSubmissionsFunc = (id, submissions) => dispatch({
    type: 'UPDATE_ASSIGNMENT_SUBMISSIONS',
    payload: {
      id,
      submissions
    }
  });
  const deleteAssignmentFunc = id => dispatch({
    type: 'DELETE_ASSIGNMENT',
    payload: id
  });
  const addAnnouncementFunc = a => dispatch({
    type: 'ADD_ANNOUNCEMENT',
    payload: a
  });
  const updateAnnouncementFunc = a => dispatch({
    type: 'UPDATE_ANNOUNCEMENT',
    payload: a
  });
  const deleteAnnouncementFunc = id => dispatch({
    type: 'DELETE_ANNOUNCEMENT',
    payload: id
  });
  const addPollFunc = p => dispatch({
    type: 'ADD_POLL',
    payload: p
  });
  const deletePollFunc = id => dispatch({
    type: 'DELETE_POLL',
    payload: id
  });
  const updateSettingsFunc = newSettings => dispatch({
    type: 'UPDATE_SETTINGS',
    payload: newSettings
  });
  const loginFunc = password => {
    if (!state.settings.adminPassword) {
      dispatch({
        type: 'UPDATE_SETTINGS',
        payload: {
          adminPassword: password
        }
      });
      localStorage.setItem('isAdmin', 'true');
      dispatch({
        type: 'TOGGLE_ADMIN',
        payload: true
      });
      showToast('선생님 비밀번호가 설정되고, 교사 모드로 전환되었습니다.', 'success');
      return true;
    }
    if (password === state.settings.adminPassword) {
      localStorage.setItem('isAdmin', 'true');
      dispatch({
        type: 'TOGGLE_ADMIN',
        payload: true
      });
      showToast('교사 모드로 전환되었습니다.', 'success');
      return true;
    }
    showToast('비밀번호가 틀렸습니다.', 'error');
    return false;
  };
  const logoutFunc = () => {
    localStorage.removeItem('isAdmin');
    dispatch({
      type: 'TOGGLE_ADMIN',
      payload: false
    });
    showToast('학생 모드로 전환되었습니다.', 'info');
  };
  const value = {
    state,
    dispatch,
    showToast,
    updatePoints: updatePointsFunc,
    resetPoints: resetPointsFunc,
    updateAttendance: updateAttendanceFunc,
    loadAttendanceByDate: loadAttendanceByDateFunc,
    loadMonthlyAttendance: loadMonthlyAttendanceFunc,
    saveAttendanceDay: saveAttendanceDayFunc,
    updateStudent: updateStudentFunc,
    bulkUpdateStudents: bulkUpdateStudentsFunc,
    addStudentRecord: addStudentRecordFunc,
    deleteStudentRecord: deleteStudentRecordFunc,
    saveWeeklyPlan,
    updateSettings: updateSettingsFunc,
    addDDay,
    deleteDDay: deleteDDayFunc,
    votePoll: votePollFunc,
    collectNewsletter: collectNewsletterFunc,
    addLink: addLinkFunc,
    updateLink: updateLinkFunc,
    deleteLink: deleteLinkFunc,
    addAssignment: addAssignmentFunc,
    updateAssignmentSubmissions: updateAssignmentSubmissionsFunc,
    deleteAssignment: deleteAssignmentFunc,
    addAnnouncement: addAnnouncementFunc,
    updateAnnouncement: updateAnnouncementFunc,
    deleteAnnouncement: deleteAnnouncementFunc,
    addPoll: addPollFunc,
    deletePoll: deletePollFunc,
    login: loginFunc,
    logout: logoutFunc,
    isDbLoading,
    getWeekStart,
    toISODate
  };
  if (folderStatus === 'loading') {
    return /*#__PURE__*/React.createElement("div", {
      className: "folder-gate"
    }, /*#__PURE__*/React.createElement("p", null, "\uBD88\uB7EC\uC624\uB294 \uC911..."));
  }
  if (folderStatus !== 'connected') {
    return /*#__PURE__*/React.createElement(FolderGate, {
      status: folderStatus,
      onConnect: connectFolder,
      onReconnect: reconnect
    });
  }
  return /*#__PURE__*/React.createElement(AppContext.Provider, {
    value: value
  }, children);
}
function useApp() {
  const ctx = useContext(AppContext);
  if (!ctx) throw new Error('useApp must be used within AppProvider');
  return ctx;
}
function FolderGate({
  status,
  onConnect,
  onReconnect
}) {
  return /*#__PURE__*/React.createElement("div", {
    className: "folder-gate"
  }, /*#__PURE__*/React.createElement("div", {
    className: "folder-gate-card"
  }, /*#__PURE__*/React.createElement("div", {
    className: "folder-gate-icon"
  }, "\uD83D\uDCC1"), status === 'needs-reconnect' ? /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("h2", null, "\uB370\uC774\uD130 \uD3F4\uB354 \uC811\uADFC\uC774 \uB04A\uACBC\uC2B5\uB2C8\uB2E4"), /*#__PURE__*/React.createElement("p", null, "\uBE0C\uB77C\uC6B0\uC800\uB97C \uC0C8\uB85C \uC2DC\uC791\uD558\uBA74 \uBCF4\uC548\uC0C1 \uD3F4\uB354 \uC811\uADFC\uC744 \uB2E4\uC2DC \uD5C8\uC6A9\uD574\uC57C \uD569\uB2C8\uB2E4."), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary btn-lg",
    onClick: onReconnect
  }, "\uD3F4\uB354 \uB2E4\uC2DC \uC5F0\uACB0\uD558\uAE30")) : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("h2", null, "\uD559\uAE09 \uB370\uC774\uD130\uB97C \uC800\uC7A5\uD560 \uD3F4\uB354\uB97C \uC120\uD0DD\uD574\uC8FC\uC138\uC694"), /*#__PURE__*/React.createElement("p", null, "\uC120\uD0DD\uD55C \uD3F4\uB354 \uC548\uC5D0 db.json \uD30C\uC77C\uB85C \uBAA8\uB4E0 \uB370\uC774\uD130\uAC00 \uC800\uC7A5\uB429\uB2C8\uB2E4. \uD55C \uBC88\uB9CC \uC120\uD0DD\uD558\uBA74 \uB429\uB2C8\uB2E4."), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary btn-lg",
    onClick: onConnect
  }, "\uD3F4\uB354 \uC120\uD0DD\uD558\uAE30"), /*#__PURE__*/React.createElement("p", {
    className: "folder-gate-note"
  }, "\u26A0\uFE0F \uD06C\uB86C \uB610\uB294 \uC5E3\uC9C0 \uBE0C\uB77C\uC6B0\uC800\uC5D0\uC11C\uB9CC \uC9C0\uC6D0\uB429\uB2C8\uB2E4."))));
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
  const {
    state
  } = useApp();
  return /*#__PURE__*/React.createElement("div", {
    className: "toast-container"
  }, state.toasts.map(t => /*#__PURE__*/React.createElement("div", {
    key: t.id,
    className: `toast ${t.type}`
  }, /*#__PURE__*/React.createElement("span", null, t.type === 'success' ? '✅' : t.type === 'error' ? '❌' : t.type === 'warning' ? '⚠️' : 'ℹ️'), /*#__PURE__*/React.createElement("span", null, t.message))));
}

/* ===================================================================
   Icon (이모지 대신 쓰는 선 아이콘)
   =================================================================== */
const ICON_PATHS = {
  dashboard: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("rect", {
    x: "3",
    y: "3",
    width: "7",
    height: "9",
    rx: "1.5"
  }), /*#__PURE__*/React.createElement("rect", {
    x: "14",
    y: "3",
    width: "7",
    height: "5",
    rx: "1.5"
  }), /*#__PURE__*/React.createElement("rect", {
    x: "14",
    y: "12",
    width: "7",
    height: "9",
    rx: "1.5"
  }), /*#__PURE__*/React.createElement("rect", {
    x: "3",
    y: "16",
    width: "7",
    height: "5",
    rx: "1.5"
  })),
  weekly: /*#__PURE__*/React.createElement("path", {
    d: "M4 6h16M4 12h16M4 18h10"
  }),
  attendance: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("rect", {
    x: "4",
    y: "4",
    width: "16",
    height: "17",
    rx: "2"
  }), /*#__PURE__*/React.createElement("path", {
    d: "M9 2.5v3M15 2.5v3M8.5 13l2.5 2.5 4.5-5"
  })),
  learning: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
    d: "M4 5.5C4 4.7 4.7 4 5.5 4H11v16H5.5c-.8 0-1.5-.7-1.5-1.5z"
  }), /*#__PURE__*/React.createElement("path", {
    d: "M20 5.5c0-.8-.7-1.5-1.5-1.5H13v16h5.5c.8 0 1.5-.7 1.5-1.5z"
  })),
  calendar: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("rect", {
    x: "3.5",
    y: "5",
    width: "17",
    height: "15.5",
    rx: "2"
  }), /*#__PURE__*/React.createElement("path", {
    d: "M3.5 10h17M8 3v4M16 3v4"
  })),
  tools: /*#__PURE__*/React.createElement("path", {
    d: "M14.5 6.5l3 3M4 20l4.5-1L19 8.5a2.1 2.1 0 0 0-3-3L5.5 16z"
  }),
  links: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
    d: "M10 14a4.5 4.5 0 0 0 6.4 0l3-3a4.5 4.5 0 0 0-6.4-6.4l-1 1"
  }), /*#__PURE__*/React.createElement("path", {
    d: "M14 10a4.5 4.5 0 0 0-6.4 0l-3 3a4.5 4.5 0 0 0 6.4 6.4l1-1"
  })),
  students: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("circle", {
    cx: "9",
    cy: "8",
    r: "3.5"
  }), /*#__PURE__*/React.createElement("path", {
    d: "M2.5 20c.8-3.6 3.4-5.5 6.5-5.5s5.7 1.9 6.5 5.5"
  }), /*#__PURE__*/React.createElement("circle", {
    cx: "17",
    cy: "9",
    r: "2.5"
  }), /*#__PURE__*/React.createElement("path", {
    d: "M17.5 14.5c2.2.3 3.6 1.9 4 4.5"
  })),
  gear: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
    d: "M4 7h9M17 7h3M4 17h3M11 17h9"
  }), /*#__PURE__*/React.createElement("circle", {
    cx: "15",
    cy: "7",
    r: "2"
  }), /*#__PURE__*/React.createElement("circle", {
    cx: "9",
    cy: "17",
    r: "2"
  })),
  plus: /*#__PURE__*/React.createElement("path", {
    d: "M12 5v14M5 12h14"
  }),
  edit: /*#__PURE__*/React.createElement("path", {
    d: "M14.5 6.5l3 3M4 20l4.5-1L19 8.5a2.1 2.1 0 0 0-3-3L5.5 16z"
  }),
  trash: /*#__PURE__*/React.createElement("path", {
    d: "M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"
  }),
  move: /*#__PURE__*/React.createElement("path", {
    d: "M12 3v18M3 12h18M12 3l-3 3M12 3l3 3M12 21l-3-3M12 21l3-3M3 12l3-3M3 12l3 3M21 12l-3-3M21 12l-3 3"
  }),
  resize: /*#__PURE__*/React.createElement("path", {
    d: "M20 14v6h-6M20 20l-7-7M10 4H4v6M4 4l7 7"
  }),
  menu: /*#__PURE__*/React.createElement("path", {
    d: "M4 7h16M4 12h16M4 17h16"
  }),
  chevron: /*#__PURE__*/React.createElement("path", {
    d: "M15 5l-7 7 7 7"
  }),
  cake: /*#__PURE__*/React.createElement("path", {
    d: "M4 21h16M5 21v-7a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2v7M5 16c1.5 1 3 1 4.5 0s3-1 4.5 0 3 1 4.5 0M12 12V8M12 5.5c.8-.6.8-1.6 0-2.5-.8.9-.8 1.9 0 2.5z"
  }),
  search: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("circle", {
    cx: "11",
    cy: "11",
    r: "6.5"
  }), /*#__PURE__*/React.createElement("path", {
    d: "M16 16l4.5 4.5"
  })),
  minus: /*#__PURE__*/React.createElement("path", {
    d: "M5 12h14"
  }),
  lock: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("rect", {
    x: "5",
    y: "11",
    width: "14",
    height: "9.5",
    rx: "2"
  }), /*#__PURE__*/React.createElement("path", {
    d: "M8.5 11V8a3.5 3.5 0 0 1 7 0v3"
  })),
  sun: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("circle", {
    cx: "12",
    cy: "12",
    r: "4"
  }), /*#__PURE__*/React.createElement("path", {
    d: "M12 2.5v2M12 19.5v2M2.5 12h2M19.5 12h2M5.3 5.3l1.4 1.4M17.3 17.3l1.4 1.4M5.3 18.7l1.4-1.4M17.3 6.7l1.4-1.4"
  }))
};
function Icon({
  name,
  size = 20
}) {
  return /*#__PURE__*/React.createElement("svg", {
    "aria-hidden": "true",
    width: size,
    height: size,
    viewBox: "0 0 24 24",
    fill: "none",
    stroke: "currentColor",
    strokeWidth: "1.8",
    strokeLinecap: "round",
    strokeLinejoin: "round"
  }, ICON_PATHS[name]);
}

/* ===================================================================
   Thermometer
   =================================================================== */
function Thermometer({
  current,
  goal,
  reward,
  isAdmin,
  onUpdate
}) {
  const percentage = Math.min(current / goal * 100, 100);
  const isGoalReached = current >= goal;
  const [isEditing, setIsEditing] = useState(false);
  const [editGoal, setEditGoal] = useState(goal);
  const [editReward, setEditReward] = useState(reward || '');
  const handleSave = () => {
    onUpdate && onUpdate(parseInt(editGoal) || 1500, editReward);
    setIsEditing(false);
  };
  return /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    className: "wcard-head"
  }, /*#__PURE__*/React.createElement("h2", {
    className: "wcard-title"
  }, "\uC6B0\uB9AC \uBC18 \uCE6D\uCC2C \uC628\uB3C4\uACC4"), isAdmin && !isEditing && /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "chip-btn",
    onClick: () => {
      setEditGoal(goal);
      setEditReward(reward || '');
      setIsEditing(true);
    }
  }, "\uBAA9\uD45C \uBC14\uAFB8\uAE30")), /*#__PURE__*/React.createElement("div", {
    className: "wcard-body thermo"
  }, isEditing && isAdmin && /*#__PURE__*/React.createElement("div", {
    className: "thermo-form"
  }, /*#__PURE__*/React.createElement("label", {
    className: "thermo-field"
  }, /*#__PURE__*/React.createElement("span", null, "\uBAA9\uD45C \uC810\uC218"), /*#__PURE__*/React.createElement("input", {
    type: "number",
    min: "1",
    value: editGoal,
    onChange: e => setEditGoal(e.target.value),
    className: "form-input"
  })), /*#__PURE__*/React.createElement("label", {
    className: "thermo-field thermo-field--wide"
  }, /*#__PURE__*/React.createElement("span", null, "\uBCF4\uC0C1"), /*#__PURE__*/React.createElement("input", {
    type: "text",
    value: editReward,
    onChange: e => setEditReward(e.target.value),
    placeholder: "\uC608: \uAE08\uC694\uC77C \uC601\uD654 \uAC10\uC0C1",
    className: "form-input"
  })), /*#__PURE__*/React.createElement("div", {
    className: "thermo-form-actions"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-secondary btn-sm",
    onClick: () => setIsEditing(false)
  }, "\uCDE8\uC18C"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-primary btn-sm",
    onClick: handleSave
  }, "\uC800\uC7A5"))), /*#__PURE__*/React.createElement("div", {
    className: "thermo-score"
  }, /*#__PURE__*/React.createElement("span", {
    className: "num-xl"
  }, current), /*#__PURE__*/React.createElement("span", {
    className: "thermo-goal"
  }, "/ ", goal, "\uC810")), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "thermo-track",
    role: "progressbar",
    "aria-label": "\uCE6D\uCC2C \uC810\uC218",
    "aria-valuemin": 0,
    "aria-valuemax": goal,
    "aria-valuenow": current
  }, /*#__PURE__*/React.createElement("span", {
    className: isGoalReached ? 'is-done' : '',
    style: {
      width: `${percentage}%`
    }
  })), /*#__PURE__*/React.createElement("div", {
    className: "thermo-ticks"
  }, [0, 0.25, 0.5, 0.75, 1].map(f => /*#__PURE__*/React.createElement("span", {
    key: f
  }, Math.round(goal * f))))), /*#__PURE__*/React.createElement("p", {
    className: "thermo-note"
  }, isGoalReached ? /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("strong", null, "\uBAA9\uD45C \uB2EC\uC131!"), " ", reward ? /*#__PURE__*/React.createElement(React.Fragment, null, "\uC774\uC81C ", /*#__PURE__*/React.createElement("strong", null, reward), " \uC2DC\uAC04\uC774\uC5D0\uC694.") : '다 같이 축하해요.') : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("strong", null, goal - current, "\uC810\uB9CC \uB354!"), " ", reward ? /*#__PURE__*/React.createElement(React.Fragment, null, "\uBAA9\uD45C\uB97C \uCC44\uC6B0\uBA74 ", /*#__PURE__*/React.createElement("strong", null, reward), "\uC744(\uB97C) \uD574\uC694.") : '보상을 정하면 여기에 보여요.'))));
}

/* ===================================================================
   AdminSettingsModal (캘린더 ID + Gemini 키)
   =================================================================== */
function AdminSettingsModal({
  onClose
}) {
  const {
    state,
    updateSettings,
    showToast
  } = useApp();
  const {
    settings
  } = state;
  const [calendarId1, setCalendarId1] = useState(settings.calendarId1 || '');
  const [calendarId2, setCalendarId2] = useState(settings.calendarId2 || '');
  const [calendarId3, setCalendarId3] = useState(settings.calendarId3 || '');
  const [geminiApiKey, setGeminiApiKey] = useState(settings.geminiApiKey || '');
  const handleSave = () => {
    updateSettings({
      ...settings,
      calendarId1,
      calendarId2,
      calendarId3,
      geminiApiKey
    });
    showToast('설정이 저장되었습니다.', 'success');
    onClose();
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "settings-modal-overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "settings-modal-content",
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "settings-modal-header"
  }, /*#__PURE__*/React.createElement("h3", null, "\u2699\uFE0F \uAD00\uB9AC\uC790 \uC124\uC815"), /*#__PURE__*/React.createElement("button", {
    className: "settings-close-btn",
    onClick: onClose
  }, "\xD7")), /*#__PURE__*/React.createElement("div", {
    className: "settings-modal-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "settings-form-group"
  }, /*#__PURE__*/React.createElement("label", null, "\uD83D\uDCC5 \uAD6C\uAE00 \uCE98\uB9B0\uB354 ID 1 (\uD544\uC218)"), /*#__PURE__*/React.createElement("input", {
    type: "text",
    className: "settings-form-input",
    placeholder: "\uC608: ...@group.calendar.google.com",
    value: calendarId1,
    onChange: e => setCalendarId1(e.target.value)
  })), /*#__PURE__*/React.createElement("div", {
    className: "settings-form-group"
  }, /*#__PURE__*/React.createElement("label", null, "\uD83D\uDCC5 \uAD6C\uAE00 \uCE98\uB9B0\uB354 ID 2 (\uC120\uD0DD)"), /*#__PURE__*/React.createElement("input", {
    type: "text",
    className: "settings-form-input",
    value: calendarId2,
    onChange: e => setCalendarId2(e.target.value)
  })), /*#__PURE__*/React.createElement("div", {
    className: "settings-form-group"
  }, /*#__PURE__*/React.createElement("label", null, "\uD83D\uDCC5 \uAD6C\uAE00 \uCE98\uB9B0\uB354 ID 3 (\uC120\uD0DD)"), /*#__PURE__*/React.createElement("input", {
    type: "text",
    className: "settings-form-input",
    value: calendarId3,
    onChange: e => setCalendarId3(e.target.value)
  }), /*#__PURE__*/React.createElement("small", {
    className: "settings-help-text"
  }, "\uC77C\uC815 \uBC0F \uACF5\uC9C0 \uD398\uC774\uC9C0\uC5D0 \uD45C\uC2DC\uD560 \uAD6C\uAE00 \uCE98\uB9B0\uB354 ID\uB97C \uC785\uB825\uD558\uC138\uC694. (\uCD5C\uB300 3\uAC1C)")), /*#__PURE__*/React.createElement("div", {
    className: "settings-form-group"
  }, /*#__PURE__*/React.createElement("label", null, "\u2728 Gemini API \uD0A4"), /*#__PURE__*/React.createElement("input", {
    type: "password",
    className: "settings-form-input",
    placeholder: "AI \uC885\uD569\uC758\uACAC \uC0DD\uC131\uC5D0 \uC0AC\uC6A9\uB429\uB2C8\uB2E4",
    value: geminiApiKey,
    onChange: e => setGeminiApiKey(e.target.value)
  }), /*#__PURE__*/React.createElement("small", {
    className: "settings-help-text"
  }, "\uC11C\uBC84\uAC00 \uC5C6\uC5B4\uC11C \uC774 \uD0A4\uB294 \uC120\uD0DD\uD55C \uB370\uC774\uD130 \uD3F4\uB354\uC758 db.json\uC5D0 \uADF8\uB300\uB85C \uC800\uC7A5\uB429\uB2C8\uB2E4. \uC774 \uCEF4\uD4E8\uD130\uB97C \uB2E4\uB978 \uC0AC\uB78C\uACFC \uAC19\uC774 \uC4F4\uB2E4\uBA74 \uC8FC\uC758\uD558\uC138\uC694."))), /*#__PURE__*/React.createElement("div", {
    className: "settings-modal-footer"
  }, /*#__PURE__*/React.createElement("button", {
    className: "settings-btn settings-btn-secondary",
    onClick: onClose
  }, "\uCDE8\uC18C"), /*#__PURE__*/React.createElement("button", {
    className: "settings-btn settings-btn-primary",
    onClick: handleSave
  }, "\uC800\uC7A5"))));
}

/* ===================================================================
   Layout
   =================================================================== */
const STUDENT_ALLOWED_PATHS = ['/', '/weekly', '/learning', '/calendar', '/tools', '/links', '/students'];
const NAV_ITEMS = [{
  path: '/',
  tone: 'carrot',
  label: '대시보드',
  icon: /*#__PURE__*/React.createElement(Icon, {
    name: "dashboard"
  })
}, {
  path: '/weekly',
  tone: 'apricot',
  label: '주간 일정',
  icon: /*#__PURE__*/React.createElement(Icon, {
    name: "weekly"
  })
}, {
  path: '/attendance',
  tone: 'tangerine',
  label: '출석 관리',
  icon: /*#__PURE__*/React.createElement(Icon, {
    name: "attendance"
  })
}, {
  path: '/learning',
  tone: 'cocoa',
  label: '학습 관리',
  icon: /*#__PURE__*/React.createElement(Icon, {
    name: "learning"
  })
}, {
  path: '/calendar',
  tone: 'coral',
  label: '일정 & 공지',
  icon: /*#__PURE__*/React.createElement(Icon, {
    name: "calendar"
  })
}, {
  path: '/tools',
  tone: 'carrot',
  label: '학급 도구',
  icon: /*#__PURE__*/React.createElement(Icon, {
    name: "tools"
  })
}, {
  path: '/links',
  tone: 'apricot',
  label: '바로가기',
  icon: /*#__PURE__*/React.createElement(Icon, {
    name: "links"
  })
}, {
  path: '/students',
  tone: 'tangerine',
  label: '학생 관리',
  icon: /*#__PURE__*/React.createElement(Icon, {
    name: "students"
  })
}];
const TONES = ['carrot', 'apricot', 'tangerine', 'cocoa', 'coral'];
// 과목마다 늘 같은 색 (모르는 과목은 이름으로 골고루 나눔)
const SUBJECT_TONES = {
  국어: 'carrot',
  수학: 'apricot',
  과학: 'tangerine',
  사회: 'cocoa',
  영어: 'coral',
  도덕: 'cocoa',
  체육: 'tangerine',
  음악: 'coral',
  미술: 'carrot',
  실과: 'tangerine',
  창체: 'apricot'
};
const subjectTone = name => {
  const n = (name || '').trim();
  if (!n) return undefined;
  const hit = Object.keys(SUBJECT_TONES).find(k => n.startsWith(k));
  return hit ? SUBJECT_TONES[hit] : TONES[[...n].reduce((a, c) => a + c.charCodeAt(0), 0) % TONES.length];
};
const teamTone = group => group ? TONES[(group - 1) % TONES.length] : undefined;
const WIDGET_TONES = {
  schedule: 'apricot',
  thermometer: 'carrot',
  lunch: 'cocoa',
  dday: 'coral',
  ranking: 'tangerine',
  notices: 'apricot'
};
const PAGE_TITLES = {
  '/': '대시보드',
  '/weekly': '주간 일정',
  '/attendance': '출석 관리',
  '/learning': '학습 관리',
  '/calendar': '일정 & 공지',
  '/tools': '학급 도구',
  '/links': '바로가기',
  '/students': '학생 관리'
};
const MOBILE_LABELS = {
  '/': '대시보드',
  '/weekly': '주간',
  '/attendance': '출석',
  '/learning': '학습',
  '/calendar': '일정',
  '/tools': '도구',
  '/links': '링크',
  '/students': '학생'
};
function Layout({
  pathname,
  children
}) {
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [collapsed, setCollapsed] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showPinModal, setShowPinModal] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const {
    state,
    login,
    logout
  } = useApp();
  const {
    isAdmin,
    settings
  } = state;
  const doLogin = () => {
    if (login(passwordInput)) {
      setShowPinModal(false);
      setPasswordInput('');
    }
  };
  const today = new Date();
  const isDashboard = pathname === '/';
  const pageTitle = PAGE_TITLES[pathname] || '대시보드';
  const currentNav = NAV_ITEMS.find(n => n.path === pathname);
  const isReadOnlyPath = !isAdmin && !STUDENT_ALLOWED_PATHS.includes(pathname);
  return /*#__PURE__*/React.createElement("div", {
    className: "app-layout"
  }, sidebarOpen && /*#__PURE__*/React.createElement("div", {
    className: "sidebar-overlay",
    onClick: () => setSidebarOpen(false)
  }), /*#__PURE__*/React.createElement("aside", {
    className: `sidebar ${collapsed ? 'collapsed' : ''} ${sidebarOpen ? 'open' : ''}`
  }, /*#__PURE__*/React.createElement("div", {
    className: "sidebar-header"
  }, /*#__PURE__*/React.createElement("div", {
    className: "sidebar-logo"
  }, /*#__PURE__*/React.createElement("span", {
    className: "logo-mark"
  }, "4\xB72"), /*#__PURE__*/React.createElement("div", {
    className: "logo-text"
  }, /*#__PURE__*/React.createElement("h1", null, "\uD558\uB0A8\uC911\uC559\uCD08\uB4F1\uD559\uAD50"), /*#__PURE__*/React.createElement("span", {
    className: "logo-sub"
  }, "4\uD559\uB144 2\uBC18 \xB7 \uAE40\uB300\uBD95 \uC120\uC0DD\uB2D8"))), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "sidebar-toggle",
    "aria-label": collapsed ? '메뉴 펼치기' : '메뉴 접기',
    onClick: () => {
      setCollapsed(!collapsed);
      setSidebarOpen(false);
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "chevron",
    size: 18
  }))), /*#__PURE__*/React.createElement("nav", {
    className: "sidebar-nav",
    "aria-label": "\uC8FC \uBA54\uB274"
  }, NAV_ITEMS.map(item => {
    if (item.path === '/attendance' && !isAdmin) return null;
    const isAllowed = isAdmin || STUDENT_ALLOWED_PATHS.includes(item.path);
    const isActive = pathname === item.path;
    return /*#__PURE__*/React.createElement("a", {
      key: item.path,
      href: `#${item.path}`,
      "aria-current": isActive ? 'page' : undefined,
      "data-tone": item.tone,
      className: `nav-item ${isActive ? 'active' : ''} ${!isAllowed ? 'nav-disabled' : ''}`,
      onClick: () => setSidebarOpen(false)
    }, /*#__PURE__*/React.createElement("span", {
      className: "nav-icon"
    }, item.icon), /*#__PURE__*/React.createElement("span", {
      className: "nav-label"
    }, item.label), !isAllowed && /*#__PURE__*/React.createElement("span", {
      className: "nav-lock"
    }, "\uC120\uC0DD\uB2D8"));
  }))), /*#__PURE__*/React.createElement("main", {
    className: "main-content",
    "data-tone": currentNav?.tone
  }, /*#__PURE__*/React.createElement("header", {
    className: "page-header"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "menu-btn",
    "aria-label": "\uBA54\uB274 \uC5F4\uAE30",
    onClick: () => setSidebarOpen(true)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "menu"
  })), /*#__PURE__*/React.createElement("div", {
    className: "header-content"
  }, !isDashboard && currentNav && /*#__PURE__*/React.createElement("span", {
    className: "page-icon",
    "aria-hidden": "true"
  }, currentNav.icon), /*#__PURE__*/React.createElement("div", {
    className: "header-titles"
  }, isDashboard ? /*#__PURE__*/React.createElement("h2", {
    className: "page-title"
  }, formatDate(today, 'md'), " ", getDayName(today), "\uC694\uC77C") : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("span", {
    className: "page-date"
  }, formatDate(today)), /*#__PURE__*/React.createElement("h2", {
    className: "page-title"
  }, pageTitle))), /*#__PURE__*/React.createElement("div", {
    className: "header-actions"
  }, /*#__PURE__*/React.createElement("div", {
    className: "mode-switch",
    role: "group",
    "aria-label": "\uBCF4\uAE30 \uBAA8\uB4DC"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    "aria-pressed": !isAdmin,
    onClick: () => {
      if (isAdmin) logout();
    }
  }, "\uD559\uC0DD \uD654\uBA74"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    "aria-pressed": isAdmin,
    onClick: () => {
      if (!isAdmin) setShowPinModal(true);
    }
  }, "\uC120\uC0DD\uB2D8")), isAdmin && /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "round-btn",
    "aria-label": "\uC124\uC815",
    onClick: () => setShowSettings(true)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "gear"
  }))))), /*#__PURE__*/React.createElement("div", {
    className: `pages-container ${isReadOnlyPath ? 'student-read-only' : ''}`
  }, children)), /*#__PURE__*/React.createElement("nav", {
    className: "mobile-nav",
    "aria-label": "\uC8FC \uBA54\uB274"
  }, NAV_ITEMS.map(item => {
    if (item.path === '/attendance' && !isAdmin) return null;
    const isAllowed = isAdmin || STUDENT_ALLOWED_PATHS.includes(item.path);
    const isActive = pathname === item.path;
    return /*#__PURE__*/React.createElement("a", {
      key: item.path,
      href: `#${item.path}`,
      "aria-current": isActive ? 'page' : undefined,
      "data-tone": item.tone,
      className: `mobile-nav-item ${isActive ? 'active' : ''} ${!isAllowed ? 'nav-disabled' : ''}`
    }, item.icon, /*#__PURE__*/React.createElement("span", null, MOBILE_LABELS[item.path]));
  })), showPinModal && /*#__PURE__*/React.createElement("div", {
    className: "modal-overlay",
    onClick: () => setShowPinModal(false)
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-content",
    onClick: e => e.stopPropagation(),
    style: {
      maxWidth: '350px'
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-header"
  }, /*#__PURE__*/React.createElement("h3", null, settings.adminPassword ? '선생님 로그인' : '선생님 비밀번호 설정'), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "close-btn",
    "aria-label": "\uB2EB\uAE30",
    onClick: () => setShowPinModal(false)
  }, "\xD7")), /*#__PURE__*/React.createElement("div", {
    className: "modal-body",
    style: {
      display: 'flex',
      flexDirection: 'column',
      gap: '1rem'
    }
  }, !settings.adminPassword && /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: '0.85rem',
      color: 'var(--text-secondary)'
    }
  }, "\uCC98\uC74C \uC0AC\uC6A9\uD558\uC2DC\uB294\uAD70\uC694! \uC55E\uC73C\uB85C \uC120\uC0DD\uB2D8 \uBAA8\uB4DC \uC804\uD658\uC5D0 \uC4F8 \uBE44\uBC00\uBC88\uD638\uB97C \uC815\uD574\uC8FC\uC138\uC694."), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("label", {
    htmlFor: "admin-password",
    style: {
      display: 'block',
      marginBottom: '0.5rem',
      fontSize: '0.9rem',
      color: 'var(--text-secondary)'
    }
  }, "\uBE44\uBC00\uBC88\uD638"), /*#__PURE__*/React.createElement("input", {
    id: "admin-password",
    type: "password",
    className: "form-input",
    placeholder: "\uBE44\uBC00\uBC88\uD638",
    value: passwordInput,
    onChange: e => setPasswordInput(e.target.value),
    onKeyDown: e => e.key === 'Enter' && doLogin(),
    autoFocus: true
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-footer"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-primary",
    style: {
      width: '100%'
    },
    onClick: doLogin
  }, settings.adminPassword ? '로그인' : '설정하고 시작하기')))), showSettings && isAdmin && /*#__PURE__*/React.createElement(AdminSettingsModal, {
    onClose: () => setShowSettings(false)
  }));
}

/* ===================================================================
   DashboardGrid — 위젯 이동/크기 조절 (라이브러리 없이 12열 격자)
   =================================================================== */
const GRID_COLS = 12,
  GRID_ROW_H = 72,
  GRID_GAP = 20,
  GRID_MIN_W = 3,
  GRID_MIN_H = 3,
  GRID_MAX_H = 12;
const GRID_STACK_BELOW = 900; // 이보다 좁으면 한 줄로 쌓고 편집은 끔
const DEFAULT_DASHBOARD_LAYOUT = [{
  id: 'schedule',
  x: 0,
  y: 0,
  w: 8,
  h: 4
}, {
  id: 'thermometer',
  x: 8,
  y: 0,
  w: 4,
  h: 4
}, {
  id: 'lunch',
  x: 0,
  y: 4,
  w: 4,
  h: 5
}, {
  id: 'dday',
  x: 4,
  y: 4,
  w: 4,
  h: 5
}, {
  id: 'ranking',
  x: 8,
  y: 4,
  w: 4,
  h: 5
}, {
  id: 'notices',
  x: 0,
  y: 9,
  w: 12,
  h: 4
}];
function clampItem(it) {
  const w = Math.max(GRID_MIN_W, Math.min(GRID_COLS, Math.round(it.w)));
  const h = Math.max(GRID_MIN_H, Math.min(GRID_MAX_H, Math.round(it.h)));
  return {
    ...it,
    w,
    h,
    x: Math.max(0, Math.min(GRID_COLS - w, Math.round(it.x))),
    y: Math.max(0, Math.round(it.y))
  };
}
const gridCollides = (a, b) => a.id !== b.id && a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

// 움직인 위젯은 그 자리에 두고, 나머지는 위에서부터 빈 자리를 찾아 채운 뒤 움직인 위젯도 위로 당김
function settleLayout(layout, activeId) {
  const active = layout.find(i => i.id === activeId);
  const placed = active ? [active] : [];
  const rest = layout.filter(i => i !== active).sort((a, b) => a.y - b.y || a.x - b.x);
  for (const it of rest) {
    let y = 0;
    while (placed.some(p => gridCollides({
      ...it,
      y
    }, p))) y++;
    placed.push({
      ...it,
      y
    });
  }
  if (active) {
    let y = active.y;
    while (y > 0 && !placed.some(p => gridCollides({
      ...active,
      y: y - 1
    }, p))) y--;
    placed[0] = {
      ...active,
      y
    };
  }
  return layout.map(i => placed.find(p => p.id === i.id));
}

// 저장된 배치(예전 Next.js 버전의 {lg:[{i,...}]} 형식 포함)를 읽고, 모르는 값은 기본 배치로
function normalizeLayout(saved) {
  const list = Array.isArray(saved) ? saved : saved && Array.isArray(saved.lg) ? saved.lg.map(i => ({
    ...i,
    id: i.i
  })) : [];
  const merged = DEFAULT_DASHBOARD_LAYOUT.map(d => {
    const s = list.find(x => x && x.id === d.id);
    return s && [s.x, s.y, s.w, s.h].every(Number.isFinite) ? clampItem({
      ...d,
      x: s.x,
      y: s.y,
      w: s.w,
      h: s.h
    }) : d;
  });
  return settleLayout(merged, null);
}
function DashboardGrid({
  layout,
  editing,
  onChange,
  widgets
}) {
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
  const unitX = colW + GRID_GAP,
    unitY = GRID_ROW_H + GRID_GAP;
  const apply = (id, mode, start, dc, dr) => {
    const cand = clampItem(mode === 'move' ? {
      ...start,
      x: start.x + dc,
      y: start.y + dr
    } : {
      ...start,
      w: start.w + dc,
      h: start.h + dr
    });
    return settleLayout(layout.map(i => i.id === id ? cand : i), id);
  };
  const onPointerDown = (e, id, mode) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.currentTarget.setPointerCapture(e.pointerId);
    dragRef.current = {
      id,
      mode,
      sx: e.clientX,
      sy: e.clientY,
      start: layout.find(i => i.id === id)
    };
    setActiveId(id);
  };
  const onPointerMove = e => {
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
    const delta = {
      ArrowLeft: [-1, 0],
      ArrowRight: [1, 0],
      ArrowUp: [0, -1],
      ArrowDown: [0, 1]
    }[e.key];
    if (!delta) return;
    e.preventDefault();
    onChange(apply(id, e.shiftKey ? 'resize' : 'move', layout.find(i => i.id === id), delta[0], delta[1]));
  };
  const ordered = [...shown].sort((a, b) => a.y - b.y || a.x - b.x);
  const rows = Math.max(...shown.map(i => i.y + i.h));
  return /*#__PURE__*/React.createElement("div", {
    ref: boxRef,
    className: `dgrid ${stacked ? 'is-stacked' : ''} ${editing && !stacked ? 'is-editing' : ''}`,
    style: stacked || !width ? undefined : {
      height: rows * unitY - GRID_GAP
    }
  }, ordered.map(it => {
    const w = widgets[it.id];
    if (!w) return null;
    const style = stacked || !width ? undefined : {
      left: it.x * unitX,
      top: it.y * unitY,
      width: it.w * colW + (it.w - 1) * GRID_GAP,
      height: it.h * GRID_ROW_H + (it.h - 1) * GRID_GAP
    };
    return /*#__PURE__*/React.createElement("section", {
      key: it.id,
      "data-tone": WIDGET_TONES[it.id],
      className: `wcard ${w.className || ''} ${activeId === it.id ? 'is-active' : ''}`,
      style: style,
      "aria-label": w.label
    }, w.content, editing && !stacked && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: "grid-handle grid-handle--move",
      "aria-label": `${w.label} 옮기기 (화살표 키로 이동, Shift+화살표로 크기 조절)`,
      onPointerDown: e => onPointerDown(e, it.id, 'move'),
      onPointerMove: onPointerMove,
      onPointerUp: onPointerUp,
      onPointerCancel: onPointerUp,
      onKeyDown: e => onKeyDown(e, it.id)
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "move",
      size: 16
    }), "\uC62E\uAE30\uAE30"), /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: "grid-handle grid-handle--resize",
      "aria-label": `${w.label} 크기 조절`,
      tabIndex: -1,
      onPointerDown: e => onPointerDown(e, it.id, 'resize'),
      onPointerMove: onPointerMove,
      onPointerUp: onPointerUp,
      onPointerCancel: onPointerUp
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "resize",
      size: 14
    }))));
  }));
}

/* ===================================================================
   DashboardPage
   =================================================================== */
const PERIOD_SLOTS = {
  1: [540, 580],
  2: [590, 630],
  3: [650, 690],
  4: [750, 790],
  5: [790, 830],
  6: [840, 880]
};
const minToClock = m => `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
function DashboardPage() {
  const {
    state,
    showToast,
    updateSettings,
    addDDay,
    deleteDDay,
    addAnnouncement,
    updateAnnouncement,
    deleteAnnouncement
  } = useApp();
  const {
    students,
    announcements,
    settings,
    weeklyPlans,
    ddays = [],
    isAdmin
  } = state;
  const [lunchInfo, setLunchInfo] = useState(null);
  const [isLoadingLunch, setIsLoadingLunch] = useState(true);
  const [holiday, setHoliday] = useState(null);
  const [showDdayForm, setShowDdayForm] = useState(false);
  const [newDday, setNewDday] = useState({
    name: '',
    date: '',
    emoji: ''
  });
  const [newNotice, setNewNotice] = useState({
    title: '',
    content: '',
    type: 'general'
  });
  const [showNoticeForm, setShowNoticeForm] = useState(false);
  const [editNoticeId, setEditNoticeId] = useState(null);
  const [editingLayout, setEditingLayout] = useState(false);
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    (async () => {
      setIsLoadingLunch(true);
      setLunchInfo(await fetchTodayLunch());
      setIsLoadingLunch(false);
    })();
    (async () => {
      setHoliday(await fetchSchoolHoliday());
    })();
    const t = setInterval(() => setNow(new Date()), 30000);
    return () => clearInterval(t);
  }, []);
  useEffect(() => {
    if (!isAdmin) setEditingLayout(false);
  }, [isAdmin]);
  const today = now;
  const dayName = getDayName(today);
  const nowMin = today.getHours() * 60 + today.getMinutes();
  const closeNoticeForm = () => {
    setShowNoticeForm(false);
    setEditNoticeId(null);
    setNewNotice({
      title: '',
      content: '',
      type: 'general'
    });
  };
  const handleAddNotice = () => {
    if (!newNotice.title.trim()) return;
    const dateStr = toISODate(today);
    if (editNoticeId) {
      updateAnnouncement({
        id: editNoticeId,
        ...newNotice,
        date: dateStr
      });
      showToast('공지사항이 수정되었습니다', 'success');
    } else {
      addAnnouncement({
        id: generateUUID(),
        ...newNotice,
        date: dateStr
      });
      showToast('공지사항이 등록되었습니다', 'success');
    }
    closeNoticeForm();
  };
  const handleEditNotice = n => {
    setNewNotice({
      title: n.title,
      content: n.content,
      type: n.type
    });
    setEditNoticeId(n.id);
    setShowNoticeForm(true);
  };
  const handleDeleteNotice = id => {
    if (confirm('이 공지사항을 삭제하시겠습니까?')) {
      deleteAnnouncement(id);
      showToast('공지사항이 삭제되었습니다', 'success');
    }
  };
  const weekKey = toISODate(getWeekStart(today));
  const todayPlan = weeklyPlans[weekKey]?.[dayName] || {};
  const todaySchedule = Object.keys(todayPlan).map(Number).sort((a, b) => a - b).filter(p => todayPlan[p]?.subject && PERIOD_SLOTS[p]).map(p => ({
    period: p,
    subject: todayPlan[p].subject,
    content: todayPlan[p].content || '',
    start: PERIOD_SLOTS[p][0],
    end: PERIOD_SLOTS[p][1]
  }));
  const currentClass = todaySchedule.find(c => nowMin >= c.start && nowMin <= c.end);
  const nextClass = todaySchedule.find(c => c.start > nowMin);
  let headline;
  if (holiday?.isHoliday) headline = /*#__PURE__*/React.createElement(React.Fragment, null, "\uC624\uB298\uC740 ", /*#__PURE__*/React.createElement("em", null, holiday.name), ", \uC26C\uB294 \uB0A0\uC774\uC5D0\uC694");else if (currentClass) headline = /*#__PURE__*/React.createElement(React.Fragment, null, "\uC9C0\uAE08\uC740 ", /*#__PURE__*/React.createElement("em", null, currentClass.period, "\uAD50\uC2DC ", currentClass.subject), ", ", currentClass.end - nowMin, "\uBD84 \uB0A8\uC558\uC5B4\uC694");else if (nextClass && nextClass === todaySchedule[0]) headline = /*#__PURE__*/React.createElement(React.Fragment, null, "\uC624\uB298 \uCCAB \uC218\uC5C5\uC740 ", /*#__PURE__*/React.createElement("em", null, nextClass.period, "\uAD50\uC2DC ", nextClass.subject), "\uC608\uC694");else if (nextClass) headline = /*#__PURE__*/React.createElement(React.Fragment, null, "\uB2E4\uC74C\uC740 ", /*#__PURE__*/React.createElement("em", null, nextClass.period, "\uAD50\uC2DC ", nextClass.subject), ", ", nextClass.start - nowMin, "\uBD84 \uB4A4\uC5D0 \uC2DC\uC791\uD574\uC694");else if (todaySchedule.length) headline = /*#__PURE__*/React.createElement(React.Fragment, null, "\uC624\uB298 \uC218\uC5C5\uC774 \uBAA8\uB450 \uB05D\uB0AC\uC5B4\uC694. \uC218\uACE0\uD588\uC5B4\uC694!");else headline = /*#__PURE__*/React.createElement(React.Fragment, null, "\uC624\uB298\uC740 \uB4F1\uB85D\uB41C \uC218\uC5C5\uC774 \uC5C6\uC5B4\uC694");
  const birthdayStudents = students.filter(s => isTodayBirthday(s.birthday));
  const ddaySorted = [...ddays].sort((a, b) => new Date(a.date) - new Date(b.date));
  const firstUpcoming = ddaySorted.find(ev => !getDday(ev.date).startsWith('D+'));
  const handleAddDday = () => {
    if (!newDday.name || !newDday.date) {
      showToast('이름과 날짜를 입력해주세요.', 'error');
      return;
    }
    addDDay(newDday);
    setNewDday({
      name: '',
      date: '',
      emoji: ''
    });
    setShowDdayForm(false);
    showToast('D-Day가 추가되었습니다.', 'success');
  };
  const handleUpdateThermometer = (goal, reward) => {
    updateSettings({
      ...settings,
      thermometerGoal: goal,
      thermometerReward: reward
    });
    showToast('우리반 온도계 설정이 변경되었습니다.', 'success');
  };
  const topStudents = [...students].sort((a, b) => (b.points || 0) - (a.points || 0)).slice(0, 5);
  const topScore = Math.max(1, topStudents[0]?.points || 0);
  const totalPoints = students.reduce((sum, s) => sum + (s.points || 0), 0);
  const goalPoints = settings?.thermometerGoal || 1500;
  const lunchItems = lunchInfo ? lunchInfo.menu.split('\n').map(s => s.trim()).filter(Boolean) : [];
  const layout = useMemo(() => normalizeLayout(settings?.dashboardLayouts), [settings?.dashboardLayouts]);
  const saveLayout = next => updateSettings({
    dashboardLayouts: next.map(({
      id,
      x,
      y,
      w,
      h
    }) => ({
      id,
      x,
      y,
      w,
      h
    }))
  });
  const resetLayout = () => {
    saveLayout(DEFAULT_DASHBOARD_LAYOUT);
    showToast('기본 배치로 되돌렸습니다.', 'success');
  };
  const shortDate = iso => {
    const d = new Date(iso);
    return `${d.getMonth() + 1}월 ${d.getDate()}일 ${getDayName(d)}요일`;
  };
  const noticeTypeLabel = {
    important: '중요',
    newsletter: '통신문',
    general: '일반'
  };
  const widgets = {
    schedule: {
      label: '오늘의 시간표',
      content: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
        className: "wcard-head"
      }, /*#__PURE__*/React.createElement("h2", {
        className: "wcard-title"
      }, "\uC624\uB298\uC758 \uC2DC\uAC04\uD45C"), /*#__PURE__*/React.createElement("a", {
        href: "#/weekly",
        className: "wcard-link"
      }, "\uC8FC\uAC04 \uC77C\uC815\uC5D0\uC11C \uC218\uC815")), /*#__PURE__*/React.createElement("div", {
        className: "wcard-body"
      }, todaySchedule.length > 0 ? /*#__PURE__*/React.createElement("ol", {
        className: "periods"
      }, todaySchedule.map(c => {
        const status = nowMin > c.end ? 'done' : currentClass === c ? 'now' : 'next';
        const progress = status === 'now' ? Math.round((nowMin - c.start) / (c.end - c.start) * 100) : status === 'done' ? 100 : 0;
        return /*#__PURE__*/React.createElement("li", {
          key: c.period,
          "data-tone": subjectTone(c.subject),
          className: `period is-${status}`,
          "aria-current": status === 'now' ? 'time' : undefined
        }, /*#__PURE__*/React.createElement("span", {
          className: "period-time"
        }, c.period, "\uAD50\uC2DC \xB7 ", minToClock(c.start)), /*#__PURE__*/React.createElement("span", {
          className: "period-subject"
        }, c.subject), c.content && /*#__PURE__*/React.createElement("span", {
          className: "period-content"
        }, c.content), /*#__PURE__*/React.createElement("span", {
          className: "period-bar"
        }, /*#__PURE__*/React.createElement("span", {
          style: {
            width: `${progress}%`
          }
        })));
      })) : /*#__PURE__*/React.createElement("div", {
        className: "wcard-empty"
      }, "\uC624\uB298 \uC2DC\uAC04\uD45C\uAC00 \uBE44\uC5B4 \uC788\uC5B4\uC694. \uC8FC\uAC04 \uC77C\uC815\uC5D0\uC11C \uC785\uB825\uD574 \uC8FC\uC138\uC694.")))
    },
    thermometer: {
      label: '우리 반 칭찬 온도계',
      className: 'wcard--accent',
      content: /*#__PURE__*/React.createElement(Thermometer, {
        current: totalPoints,
        goal: goalPoints,
        reward: settings?.thermometerReward || '',
        isAdmin: isAdmin,
        onUpdate: handleUpdateThermometer
      })
    },
    lunch: {
      label: '오늘의 급식',
      content: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
        className: "wcard-head"
      }, /*#__PURE__*/React.createElement("h2", {
        className: "wcard-title"
      }, "\uC624\uB298\uC758 \uAE09\uC2DD"), lunchInfo?.calories && /*#__PURE__*/React.createElement("span", {
        className: "wcard-meta num"
      }, lunchInfo.calories)), /*#__PURE__*/React.createElement("div", {
        className: "wcard-body"
      }, isLoadingLunch ? /*#__PURE__*/React.createElement("div", {
        className: "wcard-empty"
      }, "\uC2DD\uB2E8\uD45C\uB97C \uBD88\uB7EC\uC624\uB294 \uC911\uC774\uC5D0\uC694\u2026") : lunchItems.length ? /*#__PURE__*/React.createElement("ul", {
        className: "lunch-list"
      }, lunchItems.map((m, i) => /*#__PURE__*/React.createElement("li", {
        key: i
      }, m))) : /*#__PURE__*/React.createElement("div", {
        className: "wcard-empty"
      }, "\uC624\uB298\uC740 \uAE09\uC2DD \uC815\uBCF4\uAC00 \uC5C6\uC5B4\uC694.")))
    },
    dday: {
      label: '다가오는 날',
      content: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
        className: "wcard-head"
      }, /*#__PURE__*/React.createElement("h2", {
        className: "wcard-title"
      }, "\uB2E4\uAC00\uC624\uB294 \uB0A0"), isAdmin && /*#__PURE__*/React.createElement("button", {
        type: "button",
        className: "round-btn round-btn--sm",
        "aria-label": showDdayForm ? 'D-Day 추가 닫기' : 'D-Day 추가',
        "aria-expanded": showDdayForm,
        onClick: () => setShowDdayForm(!showDdayForm)
      }, /*#__PURE__*/React.createElement(Icon, {
        name: "plus",
        size: 16
      }))), /*#__PURE__*/React.createElement("div", {
        className: "wcard-body"
      }, showDdayForm && isAdmin && /*#__PURE__*/React.createElement("div", {
        className: "dday-form"
      }, /*#__PURE__*/React.createElement("input", {
        type: "text",
        "aria-label": "\uC774\uB984",
        placeholder: "\uC608: \uD604\uC7A5\uCCB4\uD5D8\uD559\uC2B5",
        value: newDday.name,
        onChange: e => setNewDday({
          ...newDday,
          name: e.target.value
        }),
        className: "form-input"
      }), /*#__PURE__*/React.createElement("input", {
        type: "date",
        "aria-label": "\uB0A0\uC9DC",
        value: newDday.date,
        onChange: e => setNewDday({
          ...newDday,
          date: e.target.value
        }),
        className: "form-input"
      }), /*#__PURE__*/React.createElement("button", {
        type: "button",
        className: "btn btn-primary btn-sm",
        onClick: handleAddDday
      }, "\uCD94\uAC00")), ddaySorted.length === 0 ? /*#__PURE__*/React.createElement("div", {
        className: "wcard-empty"
      }, "\uB4F1\uB85D\uB41C D-Day\uAC00 \uC5C6\uC5B4\uC694.") : /*#__PURE__*/React.createElement("ul", {
        className: "dday-rows"
      }, ddaySorted.map((ev, i) => /*#__PURE__*/React.createElement("li", {
        key: ev.id || i,
        className: `dday-row ${ev === firstUpcoming ? 'is-next' : ''}`
      }, /*#__PURE__*/React.createElement("span", {
        className: "dday-num num"
      }, getDday(ev.date)), /*#__PURE__*/React.createElement("span", {
        className: "dday-text"
      }, /*#__PURE__*/React.createElement("span", {
        className: "dday-title"
      }, ev.emoji ? `${ev.emoji} ` : '', ev.name), /*#__PURE__*/React.createElement("span", {
        className: "dday-when"
      }, shortDate(ev.date))), isAdmin && /*#__PURE__*/React.createElement("button", {
        type: "button",
        className: "icon-ghost",
        "aria-label": `${ev.name} 삭제`,
        onClick: () => deleteDDay(ev.id)
      }, /*#__PURE__*/React.createElement(Icon, {
        name: "trash",
        size: 16
      })))))))
    },
    ranking: {
      label: '칭찬 랭킹',
      content: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
        className: "wcard-head"
      }, /*#__PURE__*/React.createElement("h2", {
        className: "wcard-title"
      }, "\uCE6D\uCC2C \uB7AD\uD0B9"), /*#__PURE__*/React.createElement("span", {
        className: "wcard-meta"
      }, "\uC0C1\uC704 5\uBA85")), /*#__PURE__*/React.createElement("div", {
        className: "wcard-body"
      }, topStudents.length === 0 ? /*#__PURE__*/React.createElement("div", {
        className: "wcard-empty"
      }, "\uD559\uC0DD \uAD00\uB9AC\uC5D0\uC11C \uD559\uC0DD\uC744 \uB4F1\uB85D\uD558\uBA74 \uC21C\uC704\uAC00 \uBCF4\uC5EC\uC694.") : /*#__PURE__*/React.createElement("ol", {
        className: "rank-list"
      }, topStudents.map((s, idx) => /*#__PURE__*/React.createElement("li", {
        key: s.id,
        className: idx === 0 ? 'is-first' : ''
      }, /*#__PURE__*/React.createElement("span", {
        className: "rank-no num"
      }, idx + 1), /*#__PURE__*/React.createElement("span", {
        className: "rank-name"
      }, s.name), /*#__PURE__*/React.createElement("span", {
        className: "rank-score num"
      }, s.points || 0), /*#__PURE__*/React.createElement("span", {
        className: "rank-bar"
      }, /*#__PURE__*/React.createElement("span", {
        style: {
          width: `${Math.round((s.points || 0) / topScore * 100)}%`
        }
      })))))))
    },
    notices: {
      label: '공지사항',
      content: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
        className: "wcard-head"
      }, /*#__PURE__*/React.createElement("h2", {
        className: "wcard-title"
      }, "\uACF5\uC9C0\uC0AC\uD56D"), isAdmin && /*#__PURE__*/React.createElement("button", {
        type: "button",
        className: "btn btn-primary btn-sm",
        onClick: () => setShowNoticeForm(true)
      }, /*#__PURE__*/React.createElement(Icon, {
        name: "plus",
        size: 16
      }), "\uC0C8 \uACF5\uC9C0")), /*#__PURE__*/React.createElement("div", {
        className: "wcard-body"
      }, announcements.length === 0 ? /*#__PURE__*/React.createElement("div", {
        className: "wcard-empty"
      }, "\uB4F1\uB85D\uB41C \uACF5\uC9C0\uC0AC\uD56D\uC774 \uC5C6\uC5B4\uC694.") : /*#__PURE__*/React.createElement("div", {
        className: "notice-grid"
      }, announcements.map(n => /*#__PURE__*/React.createElement("article", {
        key: n.id,
        className: `notice-card is-${n.type || 'general'}`
      }, /*#__PURE__*/React.createElement("div", {
        className: "notice-card-top"
      }, /*#__PURE__*/React.createElement("span", {
        className: `notice-tag is-${n.type || 'general'}`
      }, noticeTypeLabel[n.type] || '일반'), /*#__PURE__*/React.createElement("span", {
        className: "notice-date num"
      }, (n.date || '').slice(5).replace('-', '.')), isAdmin && /*#__PURE__*/React.createElement("span", {
        className: "notice-actions"
      }, /*#__PURE__*/React.createElement("button", {
        type: "button",
        className: "icon-ghost",
        "aria-label": `${n.title} 수정`,
        onClick: () => handleEditNotice(n)
      }, /*#__PURE__*/React.createElement(Icon, {
        name: "edit",
        size: 16
      })), /*#__PURE__*/React.createElement("button", {
        type: "button",
        className: "icon-ghost",
        "aria-label": `${n.title} 삭제`,
        onClick: () => handleDeleteNotice(n.id)
      }, /*#__PURE__*/React.createElement(Icon, {
        name: "trash",
        size: 16
      })))), /*#__PURE__*/React.createElement("h3", {
        className: "notice-card-title"
      }, n.title), n.content && /*#__PURE__*/React.createElement("p", {
        className: "notice-card-body"
      }, n.content))))))
    }
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "dash"
  }, /*#__PURE__*/React.createElement("div", {
    className: "dash-top"
  }, /*#__PURE__*/React.createElement("h1", {
    className: "dash-headline"
  }, headline), isAdmin && /*#__PURE__*/React.createElement("div", {
    className: "dash-tools"
  }, editingLayout && /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-secondary btn-sm",
    onClick: resetLayout
  }, "\uAE30\uBCF8 \uBC30\uCE58\uB85C"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: `btn btn-sm ${editingLayout ? 'btn-primary' : 'btn-secondary'}`,
    "aria-pressed": editingLayout,
    onClick: () => setEditingLayout(!editingLayout)
  }, editingLayout ? '배치 완료' : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(Icon, {
    name: "move",
    size: 16
  }), "\uBC30\uCE58 \uD3B8\uC9D1")))), editingLayout && /*#__PURE__*/React.createElement("p", {
    className: "dash-edit-hint"
  }, "\uC704\uC82F \uC624\uB978\uCABD \uC704 ", /*#__PURE__*/React.createElement("strong", null, "\uC62E\uAE30\uAE30"), "\uB97C \uB04C\uC5B4 \uC790\uB9AC\uB97C \uBC14\uAFB8\uACE0, \uC624\uB978\uCABD \uC544\uB798 \uBAA8\uC11C\uB9AC\uB97C \uB04C\uC5B4 \uD06C\uAE30\uB97C \uC870\uC808\uD558\uC138\uC694. \uBC14\uAFBC \uBC30\uCE58\uB294 \uBC14\uB85C \uC800\uC7A5\uB3FC\uC694."), birthdayStudents.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "dash-banner"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "cake",
    size: 22
  }), /*#__PURE__*/React.createElement("span", {
    className: "dash-banner-main"
  }, "\uC624\uB298\uC740 ", /*#__PURE__*/React.createElement("strong", null, birthdayStudents.map(s => s.name).join(', ')), " \uD559\uC0DD\uC758 \uC0DD\uC77C\uC774\uC5D0\uC694"), /*#__PURE__*/React.createElement("span", {
    className: "dash-banner-sub"
  }, "\uB2E4 \uAC19\uC774 \uCD95\uD558\uD574 \uC8FC\uC138\uC694")), holiday?.isHoliday && /*#__PURE__*/React.createElement("div", {
    className: "dash-banner"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "sun",
    size: 22
  }), /*#__PURE__*/React.createElement("span", {
    className: "dash-banner-main"
  }, "\uD559\uC0AC\uC77C\uC815\uC0C1 \uD734\uC5C5\uC77C: ", /*#__PURE__*/React.createElement("strong", null, holiday.name)), /*#__PURE__*/React.createElement("span", {
    className: "dash-banner-sub"
  }, "\uCD9C\uC11D \uCCB4\uD06C\uAC00 \uD544\uC694 \uC5C6\uB294 \uB0A0\uC77C \uC218 \uC788\uC5B4\uC694")), /*#__PURE__*/React.createElement(DashboardGrid, {
    layout: layout,
    editing: editingLayout,
    onChange: saveLayout,
    widgets: widgets
  }), isAdmin && showNoticeForm && /*#__PURE__*/React.createElement("div", {
    className: "notice-modal-overlay",
    onClick: closeNoticeForm
  }, /*#__PURE__*/React.createElement("div", {
    className: "notice-modal-content",
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "notice-modal-header"
  }, /*#__PURE__*/React.createElement("h3", null, editNoticeId ? '공지사항 수정' : '새 공지사항'), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "notice-close-btn",
    "aria-label": "\uB2EB\uAE30",
    onClick: closeNoticeForm
  }, "\xD7")), /*#__PURE__*/React.createElement("div", {
    className: "notice-modal-body"
  }, /*#__PURE__*/React.createElement("input", {
    className: "form-input",
    "aria-label": "\uC81C\uBAA9",
    placeholder: "\uC81C\uBAA9\uC744 \uC785\uB825\uD558\uC138\uC694",
    value: newNotice.title,
    onChange: e => setNewNotice({
      ...newNotice,
      title: e.target.value
    })
  }), /*#__PURE__*/React.createElement("textarea", {
    className: "form-input",
    "aria-label": "\uB0B4\uC6A9",
    placeholder: "\uB0B4\uC6A9\uC744 \uC785\uB825\uD558\uC138\uC694",
    value: newNotice.content,
    onChange: e => setNewNotice({
      ...newNotice,
      content: e.target.value
    }),
    rows: "5",
    style: {
      resize: 'vertical'
    }
  }), /*#__PURE__*/React.createElement("select", {
    className: "form-input",
    "aria-label": "\uC885\uB958",
    value: newNotice.type,
    onChange: e => setNewNotice({
      ...newNotice,
      type: e.target.value
    })
  }, /*#__PURE__*/React.createElement("option", {
    value: "general"
  }, "\uC77C\uBC18"), /*#__PURE__*/React.createElement("option", {
    value: "important"
  }, "\uC911\uC694"), /*#__PURE__*/React.createElement("option", {
    value: "newsletter"
  }, "\uD1B5\uC2E0\uBB38"))), /*#__PURE__*/React.createElement("div", {
    className: "notice-modal-footer"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-secondary",
    onClick: closeNoticeForm
  }, "\uCDE8\uC18C"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-primary",
    onClick: handleAddNotice
  }, editNoticeId ? '수정' : '등록')))));
}

/* ===================================================================
   StudentsPage
   =================================================================== */
function StudentsPage() {
  const {
    state,
    dispatch,
    showToast,
    updatePoints,
    resetPoints,
    updateStudent,
    bulkUpdateStudents,
    addStudentRecord,
    deleteStudentRecord
  } = useApp();
  const {
    students,
    isAdmin,
    studentRecords,
    assignments,
    settings
  } = state;
  const [activeTab, setActiveTab] = useState('list');
  useEffect(() => {
    if (!isAdmin) setActiveTab('list');
  }, [isAdmin]); // 학생 화면으로 바꾸면 선생님 전용 탭을 닫음
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
  const handleAddPoints = (id, currentPoints, amount) => {
    updatePoints(id, currentPoints, amount);
    showToast(`${amount > 0 ? '+' : ''}${amount}점이 부여되었습니다.`, 'success');
  };
  const groupPoints = [1, 2, 3, 4, 5, 6].map(groupId => {
    const groupStudents = students.filter(s => s.group === groupId);
    const totalPoints = groupStudents.reduce((sum, s) => sum + (s.points || 0), 0);
    return {
      groupId,
      totalPoints,
      students: groupStudents
    };
  }).sort((a, b) => b.totalPoints - a.totalPoints);
  const seatingList = [...students].sort((a, b) => {
    const aIdx = a.seatIndex !== undefined ? a.seatIndex : a.number;
    const bIdx = b.seatIndex !== undefined ? b.seatIndex : b.number;
    return aIdx - bIdx;
  });
  const handleDragStartSeat = (e, student) => {
    if (!isAdmin) return;
    setDraggedSeat(student);
    e.dataTransfer.effectAllowed = 'move';
    e.target.style.opacity = '0.5';
  };
  const handleDragEndSeat = e => {
    if (!isAdmin) return;
    e.target.style.opacity = '1';
    setDraggedSeat(null);
  };
  const handleDragOverSeat = e => {
    if (!isAdmin) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
  };
  const handleDropSeat = (e, targetStudent) => {
    if (!isAdmin) return;
    e.preventDefault();
    if (!draggedSeat || draggedSeat.id === targetStudent.id) return;
    const dragIndex = draggedSeat.seatIndex !== undefined ? draggedSeat.seatIndex : draggedSeat.number;
    const targetIndex = targetStudent.seatIndex !== undefined ? targetStudent.seatIndex : targetStudent.number;
    bulkUpdateStudents([{
      id: draggedSeat.id,
      seatIndex: targetIndex
    }, {
      id: targetStudent.id,
      seatIndex: dragIndex
    }]);
  };
  const handleDragStartTeam = (e, student) => {
    if (!isAdmin) return;
    setDraggedGroupStudent(student);
    e.dataTransfer.effectAllowed = 'move';
    e.target.style.opacity = '0.5';
  };
  const handleDragEndTeam = e => {
    if (!isAdmin) return;
    e.target.style.opacity = '1';
    setDraggedGroupStudent(null);
  };
  const handleDragOverTeam = e => {
    if (!isAdmin) return;
    e.preventDefault();
    e.currentTarget.classList.add('is-over');
  };
  const handleDragLeaveTeam = e => {
    if (!isAdmin) return;
    e.currentTarget.classList.remove('is-over');
  };
  const handleDropTeam = (e, targetGroupId) => {
    if (!isAdmin) return;
    e.preventDefault();
    e.currentTarget.classList.remove('is-over');
    if (!draggedGroupStudent || draggedGroupStudent.group === targetGroupId) return;
    updateStudent({
      id: draggedGroupStudent.id,
      group: targetGroupId
    });
  };
  const handleSelectStudent = s => {
    setSelectedStudent(s);
    setEditGender(s.gender || 'M');
    setEditBirthday(s.birthday || '');
    setEditParentPhone(s.parentPhone || '');
    setIsEditingName(false);
    setNewRecordContent('');
  };
  const handleSaveStudentInfo = () => {
    updateStudent({
      id: selectedStudent.id,
      gender: editGender,
      birthday: editBirthday,
      parentPhone: editParentPhone
    });
    setSelectedStudent({
      ...selectedStudent,
      gender: editGender,
      birthday: editBirthday,
      parentPhone: editParentPhone
    });
    showToast('학생 정보가 수정되었습니다.', 'success');
  };
  const handleResetPoints = () => {
    if (!window.confirm(`${selectedStudent.name} 학생의 포인트를 초기화하시겠습니까?`)) return;
    resetPoints(selectedStudent.id);
    setSelectedStudent({
      ...selectedStudent,
      points: 0
    });
    showToast('포인트가 초기화되었습니다.', 'success');
  };
  const handleAddRecord = () => {
    if (!newRecordContent.trim()) {
      showToast('내용을 입력해주세요.', 'error');
      return;
    }
    addStudentRecord(selectedStudent.id, newRecordType, newRecordContent, newRecordCategory);
    setNewRecordContent('');
  };
  const handleBulkUpdate = () => {
    const lines = bulkNames.split('\n').map(l => l.trim()).filter(l => l);
    if (lines.length === 0) {
      showToast('입력된 내용이 없습니다.', 'error');
      return;
    }
    const newStudents = lines.map((line, i) => {
      const parts = line.split(/[\t,]+/).map(p => p.trim());
      const existing = students.find(s => s.number === i + 1);
      let gender = 'M';
      if (parts[1]) {
        const g = parts[1].toUpperCase();
        if (g === 'F' || g === '여') gender = 'F';
      }
      return {
        ...existing,
        id: existing ? existing.id : generateUUID(),
        number: i + 1,
        name: parts[0] || `학생${i + 1}`,
        gender,
        birthday: parts[2] || (existing ? existing.birthday : ''),
        points: existing ? existing.points : 0,
        group: existing ? existing.group : Math.ceil((i + 1) / 4)
      };
    });
    dispatch({
      type: 'SET_STUDENTS',
      payload: newStudents
    });
    bulkUpdateStudents(newStudents);
    showToast(`총 ${newStudents.length}명으로 명단이 갱신되었습니다.`, 'success');
    setShowBulkModal(false);
    setBulkNames('');
  };
  const handleExcelUpload = async e => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const data = await parseStudentExcel(file);
      if (data.length === 0) {
        showToast('데이터가 없습니다.', 'error');
        return;
      }
      const newStudents = data.map((row, i) => {
        const existing = students.find(s => s.number === i + 1);
        let gender = 'M';
        if (row['성별']) {
          const g = String(row['성별']).toUpperCase();
          if (g === 'F' || g === '여') gender = 'F';
        }
        return {
          ...existing,
          id: existing ? existing.id : generateUUID(),
          number: i + 1,
          name: row['이름'] || `학생${i + 1}`,
          gender,
          birthday: row['생일'] || (existing ? existing.birthday : ''),
          parentPhone: row['연락처'] || (existing ? existing.parentPhone : ''),
          points: existing ? existing.points : 0,
          group: existing ? existing.group : Math.ceil((i + 1) / 4)
        };
      });
      dispatch({
        type: 'SET_STUDENTS',
        payload: newStudents
      });
      bulkUpdateStudents(newStudents);
      showToast(`총 ${newStudents.length}명으로 명단이 갱신되었습니다.`, 'success');
    } catch (err) {
      showToast('엑셀 파일을 읽는 중 오류가 발생했습니다.', 'error');
    }
    e.target.value = null;
  };
  const handleGenerateAI = async student => {
    const recs = studentRecords[student.id] || [];
    const evals = assignments.filter(a => a.submissions?.[student.id]).map(a => ({
      title: a.title,
      grade: a.submissions[student.id].grade
    }));
    const studentData = {
      name: student.name,
      gender: student.gender,
      points: student.points || 0,
      records: recs,
      evaluations: evals
    };
    try {
      setGeneratingAiId(student.id);
      const summary = await generateBehaviorSummary(studentData, settings.geminiApiKey);
      updateStudent({
        ...student,
        aiSummary: summary
      });
      showToast('AI가 종합의견을 생성했습니다.', 'success');
    } catch (err) {
      showToast('AI 생성 중 오류가 발생했습니다: ' + err.message, 'error');
    } finally {
      setGeneratingAiId(null);
    }
  };
  const handleEvaluate = (assignmentId, studentId, grade) => {
    dispatch({
      type: 'EVALUATE_ASSIGNMENT',
      payload: {
        id: assignmentId,
        studentId,
        grade
      }
    });
    showToast('평가가 저장되었습니다.', 'success');
  };
  const records = selectedStudent ? studentRecords[selectedStudent.id] || [] : [];
  const TABS = [{
    k: 'list',
    label: '학생 목록'
  }, {
    k: 'behavior',
    label: '행동발달',
    admin: true
  }, {
    k: 'evaluation',
    label: '학습평가',
    admin: true
  }];
  const VIEWS = [{
    v: 'card',
    label: '카드'
  }, {
    v: 'table',
    label: '표'
  }, {
    v: 'seating',
    label: '자리'
  }, {
    v: 'points',
    label: '칭찬'
  }];
  const fmtBirthday = b => {
    const m = /^\d{4}-(\d{2})-(\d{2})$/.exec(b || '');
    return m ? `${+m[1]}월 ${+m[2]}일` : '';
  };
  const rankClass = i => i === 0 ? 'is-first' : '';
  return /*#__PURE__*/React.createElement("div", {
    className: "stu"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wk-top"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wk-title"
  }, /*#__PURE__*/React.createElement("h1", {
    className: "dash-headline"
  }, "\uC6B0\uB9AC \uBC18 ", /*#__PURE__*/React.createElement("em", null, students.length, "\uBA85"))), /*#__PURE__*/React.createElement("div", {
    className: "mode-switch",
    role: "group",
    "aria-label": "\uD559\uC0DD \uAD00\uB9AC \uBA54\uB274"
  }, TABS.map(t => {
    const locked = t.admin && !isAdmin;
    return /*#__PURE__*/React.createElement("button", {
      key: t.k,
      type: "button",
      "aria-pressed": activeTab === t.k,
      "aria-disabled": locked || undefined,
      className: locked ? 'is-locked' : '',
      onClick: () => locked ? showToast('선생님 모드에서 볼 수 있어요.', 'error') : setActiveTab(t.k)
    }, t.label, locked && /*#__PURE__*/React.createElement(Icon, {
      name: "lock",
      size: 14
    }));
  }))), activeTab === 'list' && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    className: "stu-bar"
  }, /*#__PURE__*/React.createElement("label", {
    className: "stu-search"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "search",
    size: 18
  }), /*#__PURE__*/React.createElement("span", {
    className: "sr-only"
  }, "\uD559\uC0DD \uAC80\uC0C9"), /*#__PURE__*/React.createElement("input", {
    type: "search",
    placeholder: "\uC774\uB984\uC774\uB098 \uBC88\uD638\uB85C \uCC3E\uAE30",
    value: search,
    onChange: e => setSearch(e.target.value)
  })), /*#__PURE__*/React.createElement("div", {
    className: "mode-switch",
    role: "group",
    "aria-label": "\uBCF4\uAE30 \uBC29\uC2DD"
  }, VIEWS.map(b => /*#__PURE__*/React.createElement("button", {
    key: b.v,
    type: "button",
    "aria-pressed": view === b.v,
    onClick: () => setView(b.v)
  }, b.label))), isAdmin && /*#__PURE__*/React.createElement("div", {
    className: "stu-admin"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-secondary btn-sm",
    onClick: downloadStudentTemplate
  }, "\uC591\uC2DD \uBC1B\uAE30"), /*#__PURE__*/React.createElement("label", {
    className: "btn btn-secondary btn-sm stu-upload"
  }, "\uC5D1\uC140 \uC62C\uB9AC\uAE30", /*#__PURE__*/React.createElement("input", {
    type: "file",
    accept: ".xlsx, .xls",
    className: "sr-only",
    onChange: handleExcelUpload
  })), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-secondary btn-sm",
    onClick: () => {
      const existingNames = students.map(s => `${s.name}\t${s.gender || ''}\t${s.birthday || ''}`).join('\n');
      setBulkNames(existingNames);
      setShowBulkModal(true);
    }
  }, "\uBA85\uB2E8 \uD55C \uBC88\uC5D0 \uACE0\uCE58\uAE30"))), view === 'card' && (filtered.length ? /*#__PURE__*/React.createElement("ul", {
    className: "stu-grid"
  }, filtered.map(s => /*#__PURE__*/React.createElement("li", {
    key: s.id,
    className: "stu-card",
    "data-tone": teamTone(s.group)
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "stu-card-main",
    onClick: () => handleSelectStudent(s)
  }, /*#__PURE__*/React.createElement("span", {
    className: "stu-card-top"
  }, /*#__PURE__*/React.createElement("span", {
    className: "att-no num"
  }, s.number), /*#__PURE__*/React.createElement("span", {
    className: "stu-name"
  }, s.name)), /*#__PURE__*/React.createElement("span", {
    className: "stu-meta"
  }, [s.group && `${s.group}팀`, s.gender === 'F' ? '여' : '남', fmtBirthday(s.birthday)].filter(Boolean).join(' · ')), (s.allergies || s.healthNotes) && /*#__PURE__*/React.createElement("span", {
    className: "stu-tags"
  }, s.allergies && /*#__PURE__*/React.createElement("span", {
    className: "stu-tag is-warn"
  }, "\uC54C\uB808\uB974\uAE30 ", s.allergies), s.healthNotes && /*#__PURE__*/React.createElement("span", {
    className: "stu-tag"
  }, s.healthNotes))), /*#__PURE__*/React.createElement("div", {
    className: "stu-card-foot"
  }, /*#__PURE__*/React.createElement("span", {
    className: "stu-points"
  }, /*#__PURE__*/React.createElement("strong", {
    className: "num"
  }, s.points || 0), "\uC810"), isAdmin && /*#__PURE__*/React.createElement("span", {
    className: "stu-steppers"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "icon-ghost",
    "aria-label": `${s.name} 1점 빼기`,
    onClick: () => handleAddPoints(s.id, s.points, -1)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "minus",
    size: 16
  })), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "icon-ghost is-plus",
    "aria-label": `${s.name} 1점 더하기`,
    onClick: () => handleAddPoints(s.id, s.points, 1)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 16
  }))))))) : /*#__PURE__*/React.createElement("div", {
    className: "wcard wcard-empty stu-empty"
  }, search ? `“${search}”에 맞는 학생이 없어요.` : '아직 학생이 없어요. 엑셀로 명단을 올려 주세요.')), view === 'table' && /*#__PURE__*/React.createElement("div", {
    className: "wk-card"
  }, /*#__PURE__*/React.createElement("table", {
    className: "stu-table"
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", {
    className: "stu-col-no"
  }, "\uBC88\uD638"), /*#__PURE__*/React.createElement("th", null, "\uC774\uB984"), /*#__PURE__*/React.createElement("th", null, "\uC131\uBCC4"), /*#__PURE__*/React.createElement("th", null, "\uC0DD\uC77C"), /*#__PURE__*/React.createElement("th", null, "\uC54C\uB808\uB974\uAE30"), /*#__PURE__*/React.createElement("th", null, "\uAC74\uAC15"), /*#__PURE__*/React.createElement("th", {
    className: "stu-col-num"
  }, "\uD3EC\uC778\uD2B8"))), /*#__PURE__*/React.createElement("tbody", null, filtered.map(s => /*#__PURE__*/React.createElement("tr", {
    key: s.id,
    tabIndex: 0,
    onClick: () => handleSelectStudent(s),
    onKeyDown: e => e.key === 'Enter' && handleSelectStudent(s)
  }, /*#__PURE__*/React.createElement("td", {
    className: "stu-col-no num"
  }, s.number), /*#__PURE__*/React.createElement("td", {
    className: "stu-table-name"
  }, s.name), /*#__PURE__*/React.createElement("td", null, s.gender === 'F' ? '여' : '남'), /*#__PURE__*/React.createElement("td", null, fmtBirthday(s.birthday) || '—'), /*#__PURE__*/React.createElement("td", {
    className: s.allergies ? 'stu-warn' : 'stu-dim'
  }, s.allergies || '—'), /*#__PURE__*/React.createElement("td", {
    className: s.healthNotes ? '' : 'stu-dim'
  }, s.healthNotes || '—'), /*#__PURE__*/React.createElement("td", {
    className: "stu-col-num num"
  }, s.points || 0)))))), view === 'seating' && /*#__PURE__*/React.createElement("section", {
    className: "wcard"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wcard-head"
  }, /*#__PURE__*/React.createElement("h2", {
    className: "wcard-title"
  }, "\uC790\uB9AC \uBC30\uCE58\uB3C4"), isAdmin && /*#__PURE__*/React.createElement("span", {
    className: "wcard-meta"
  }, "\uC774\uB984\uD45C\uB97C \uB04C\uC5B4\uB2E4 \uB193\uC73C\uBA74 \uC790\uB9AC\uAC00 \uBC14\uB00C\uC5B4\uC694")), /*#__PURE__*/React.createElement("div", {
    className: "wcard-body seating-container"
  }, /*#__PURE__*/React.createElement("div", {
    className: "teacher-desk"
  }, "\uAD50\uD0C1"), /*#__PURE__*/React.createElement("div", {
    className: "seating-grid-dynamic"
  }, seatingList.map(s => /*#__PURE__*/React.createElement("div", {
    key: s.id,
    className: "seat occupied",
    draggable: isAdmin,
    onDragStart: e => handleDragStartSeat(e, s),
    onDragEnd: handleDragEndSeat,
    onDragOver: handleDragOverSeat,
    onDrop: e => handleDropSeat(e, s),
    onClick: () => handleSelectStudent(s),
    style: {
      cursor: isAdmin ? 'grab' : 'pointer'
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "seat-number num"
  }, s.number), /*#__PURE__*/React.createElement("div", {
    className: "seat-name"
  }, s.name)))))), view === 'points' && /*#__PURE__*/React.createElement("section", {
    className: "wcard"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wcard-head"
  }, /*#__PURE__*/React.createElement("h2", {
    className: "wcard-title"
  }, "\uCE6D\uCC2C \uD3EC\uC778\uD2B8 \uC21C\uC704"), /*#__PURE__*/React.createElement("div", {
    className: "mode-switch",
    role: "group",
    "aria-label": "\uC21C\uC704 \uAE30\uC900"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    "aria-pressed": pointsView === 'individual',
    onClick: () => setPointsView('individual')
  }, "\uAC1C\uC778\uBCC4"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    "aria-pressed": pointsView === 'group',
    onClick: () => setPointsView('group')
  }, "\uD300\uBCC4"))), /*#__PURE__*/React.createElement("div", {
    className: "wcard-body"
  }, pointsView === 'individual' ? /*#__PURE__*/React.createElement("ol", {
    className: "stu-rank"
  }, sorted.map((s, i) => /*#__PURE__*/React.createElement("li", {
    key: s.id,
    className: rankClass(i)
  }, /*#__PURE__*/React.createElement("span", {
    className: "rank-no num"
  }, i + 1), /*#__PURE__*/React.createElement("span", {
    className: "stu-rank-name"
  }, s.name, /*#__PURE__*/React.createElement("span", {
    className: "stu-team",
    "data-tone": teamTone(s.group)
  }, s.group, "\uD300")), /*#__PURE__*/React.createElement("span", {
    className: "stu-rank-score"
  }, /*#__PURE__*/React.createElement("strong", {
    className: "num"
  }, s.points || 0), "\uC810"), isAdmin && /*#__PURE__*/React.createElement("span", {
    className: "stu-rank-actions"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "chip-btn",
    onClick: () => handleAddPoints(s.id, s.points, 5)
  }, "+5"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "chip-btn",
    onClick: () => handleAddPoints(s.id, s.points, 10)
  }, "+10"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "chip-btn is-minus",
    onClick: () => handleAddPoints(s.id, s.points, -5)
  }, "\u22125"))))) : /*#__PURE__*/React.createElement(React.Fragment, null, isAdmin && /*#__PURE__*/React.createElement("p", {
    className: "lrn-hint stu-hint"
  }, "\uD559\uC0DD \uC774\uB984\uD45C\uB97C \uB2E4\uB978 \uD300\uC73C\uB85C \uB04C\uC5B4\uB2E4 \uB193\uC73C\uBA74 \uD300\uC774 \uBC14\uB00C\uC5B4\uC694."), /*#__PURE__*/React.createElement("ol", {
    className: "stu-teams"
  }, groupPoints.map((g, i) => /*#__PURE__*/React.createElement("li", {
    key: g.groupId,
    "data-tone": teamTone(g.groupId),
    className: `stu-team-card ${rankClass(i)}`,
    onDragOver: handleDragOverTeam,
    onDragLeave: handleDragLeaveTeam,
    onDrop: e => handleDropTeam(e, g.groupId)
  }, /*#__PURE__*/React.createElement("div", {
    className: "stu-team-head"
  }, /*#__PURE__*/React.createElement("span", {
    className: "rank-no num"
  }, i + 1), /*#__PURE__*/React.createElement("span", {
    className: "stu-team-name"
  }, g.groupId, "\uD300"), /*#__PURE__*/React.createElement("span", {
    className: "stu-rank-score"
  }, /*#__PURE__*/React.createElement("strong", {
    className: "num"
  }, g.totalPoints), "\uC810")), /*#__PURE__*/React.createElement("div", {
    className: "stu-pills"
  }, g.students.map(s => /*#__PURE__*/React.createElement("span", {
    key: s.id,
    className: "stu-pill",
    draggable: isAdmin,
    onDragStart: e => handleDragStartTeam(e, s),
    onDragEnd: handleDragEndTeam,
    style: {
      cursor: isAdmin ? 'grab' : 'default'
    }
  }, s.name)), g.students.length === 0 && /*#__PURE__*/React.createElement("span", {
    className: "lrn-none"
  }, "\uC544\uC9C1 \uC544\uBB34\uB3C4 \uC5C6\uC5B4\uC694"))))))))), activeTab === 'behavior' && /*#__PURE__*/React.createElement("ul", {
    className: "stu-behavior"
  }, students.map(s => {
    const recs = studentRecords[s.id] || [];
    return /*#__PURE__*/React.createElement("li", {
      key: s.id,
      className: "wcard stu-bh"
    }, /*#__PURE__*/React.createElement("div", {
      className: "stu-bh-who"
    }, /*#__PURE__*/React.createElement("span", {
      className: "stu-card-top"
    }, /*#__PURE__*/React.createElement("span", {
      className: "att-no num"
    }, s.number), /*#__PURE__*/React.createElement("span", {
      className: "stu-name"
    }, s.name)), isAdmin && /*#__PURE__*/React.createElement("div", {
      className: "stu-bh-actions"
    }, /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: "btn btn-secondary btn-sm",
      onClick: () => handleSelectStudent(s)
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "plus",
      size: 16
    }), " \uAE30\uB85D \uCD94\uAC00"), /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: "btn btn-primary btn-sm",
      onClick: () => handleGenerateAI(s),
      disabled: generatingAiId === s.id
    }, generatingAiId === s.id ? 'AI가 쓰는 중…' : s.aiSummary ? 'AI 의견 다시 쓰기' : 'AI 의견 쓰기'))), /*#__PURE__*/React.createElement("div", {
      className: "stu-bh-col"
    }, /*#__PURE__*/React.createElement("h3", {
      className: "stu-bh-title"
    }, "\uB204\uC801 \uAE30\uB85D ", /*#__PURE__*/React.createElement("span", {
      className: "num"
    }, recs.length)), recs.length > 0 ? /*#__PURE__*/React.createElement("ul", {
      className: "stu-records"
    }, recs.map((r, i) => /*#__PURE__*/React.createElement("li", {
      key: r.id || i
    }, /*#__PURE__*/React.createElement("div", {
      className: "stu-record-meta"
    }, /*#__PURE__*/React.createElement("span", {
      className: "num"
    }, r.date), r.category && /*#__PURE__*/React.createElement("span", {
      className: "notice-tag"
    }, r.category)), /*#__PURE__*/React.createElement("p", {
      className: "stu-record-text"
    }, r.content), isAdmin && /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: "icon-ghost",
      "aria-label": "\uC774 \uAE30\uB85D \uC0AD\uC81C",
      onClick: () => {
        if (confirm('이 기록을 삭제하시겠습니까?')) deleteStudentRecord(s.id, r.id);
      }
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "trash",
      size: 16
    }))))) : /*#__PURE__*/React.createElement("p", {
      className: "stu-bh-empty"
    }, "\uC544\uC9C1 \uAE30\uB85D\uC774 \uC5C6\uC5B4\uC694.")), /*#__PURE__*/React.createElement("div", {
      className: "stu-bh-col"
    }, /*#__PURE__*/React.createElement("h3", {
      className: "stu-bh-title"
    }, "\uC885\uD569\uC758\uACAC"), s.aiSummary ? /*#__PURE__*/React.createElement("p", {
      className: "stu-summary"
    }, s.aiSummary) : /*#__PURE__*/React.createElement("p", {
      className: "stu-bh-empty"
    }, "\uAE30\uB85D\uC744 \uBAA8\uC740 \uB4A4 \u201CAI \uC758\uACAC \uC4F0\uAE30\u201D\uB97C \uB204\uB974\uBA74 \uCD08\uC548\uC774 \uC0DD\uACA8\uC694.")));
  })), activeTab === 'evaluation' && (assignments.length === 0 ? /*#__PURE__*/React.createElement("div", {
    className: "wcard wcard-empty stu-empty"
  }, "\uB4F1\uB85D\uB41C \uACFC\uC81C\xB7\uC218\uD589\uD3C9\uAC00\uAC00 \uC5C6\uC5B4\uC694. \uD559\uC2B5 \uAD00\uB9AC\uC5D0\uC11C \uD65C\uB3D9\uC744 \uCD94\uAC00\uD574 \uC8FC\uC138\uC694.") : /*#__PURE__*/React.createElement("div", {
    className: "wk-card"
  }, /*#__PURE__*/React.createElement("table", {
    className: "stu-eval"
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", {
    className: "stu-eval-who"
  }, "\uD559\uC0DD"), assignments.map(a => /*#__PURE__*/React.createElement("th", {
    key: a.id
  }, /*#__PURE__*/React.createElement("div", {
    className: "stu-eval-head"
  }, /*#__PURE__*/React.createElement("span", {
    className: `lrn-type ${a.type === '수행평가' ? 'is-eval' : ''}`
  }, a.type), isAdmin && /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "icon-ghost",
    "aria-label": `${a.title} 삭제`,
    onClick: () => {
      if (confirm('이 항목을 정말 삭제하시겠습니까?')) {
        dispatch({
          type: 'DELETE_ASSIGNMENT',
          payload: a.id
        });
        showToast('삭제되었습니다.', 'success');
      }
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "stu-eval-title"
  }, a.title), /*#__PURE__*/React.createElement("div", {
    className: "stu-eval-date num"
  }, a.dueDate))))), /*#__PURE__*/React.createElement("tbody", null, students.map(s => /*#__PURE__*/React.createElement("tr", {
    key: s.id
  }, /*#__PURE__*/React.createElement("th", {
    scope: "row",
    className: "stu-eval-who"
  }, /*#__PURE__*/React.createElement("span", {
    className: "att-no num"
  }, s.number), s.name), assignments.map(a => {
    const submission = a.submissions?.[s.id];
    const grade = submission?.grade;
    return /*#__PURE__*/React.createElement("td", {
      key: `${s.id}-${a.id}`,
      className: submission ? '' : 'is-missing'
    }, submission ? isAdmin ? /*#__PURE__*/React.createElement("div", {
      className: "stu-grade",
      role: "group",
      "aria-label": `${s.name} ${a.title} 평가`
    }, ['상', '중', '하'].map(g => /*#__PURE__*/React.createElement("button", {
      key: g,
      type: "button",
      "aria-pressed": grade === g,
      onClick: () => handleEvaluate(a.id, s.id, g)
    }, g))) : /*#__PURE__*/React.createElement("span", {
      className: `stu-grade-chip ${grade ? '' : 'is-wait'}`
    }, grade || '채점 전') : /*#__PURE__*/React.createElement("span", {
      className: "stu-missing"
    }, "\uBBF8\uC81C\uCD9C"));
  }))))))), selectedStudent && /*#__PURE__*/React.createElement("div", {
    className: "modal-overlay",
    onClick: () => {
      setSelectedStudent(null);
      setIsEditingName(false);
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-header"
  }, /*#__PURE__*/React.createElement("h3", {
    className: "modal-title"
  }, "\uD559\uC0DD \uC815\uBCF4"), /*#__PURE__*/React.createElement("button", {
    className: "modal-close",
    onClick: () => {
      setSelectedStudent(null);
      setIsEditingName(false);
    }
  }, "\xD7")), /*#__PURE__*/React.createElement("div", {
    className: "modal-body",
    style: {
      maxHeight: '70vh',
      overflowY: 'auto'
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "student-detail-header"
  }, /*#__PURE__*/React.createElement("div", {
    className: `avatar xl ${getAvatarColor(selectedStudent.gender)}`
  }, getInitials(selectedStudent.name)), /*#__PURE__*/React.createElement("div", {
    className: "student-detail-info"
  }, /*#__PURE__*/React.createElement("div", {
    className: "student-name-edit-wrapper"
  }, isEditingName ? /*#__PURE__*/React.createElement("div", {
    className: "edit-name-form"
  }, /*#__PURE__*/React.createElement("input", {
    type: "text",
    className: "form-input",
    value: editNameValue,
    onChange: e => setEditNameValue(e.target.value),
    autoFocus: true,
    style: {
      width: '120px',
      padding: '0.25rem 0.5rem',
      minHeight: '30px'
    }
  }), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary btn-sm",
    onClick: () => {
      if (editNameValue.trim()) {
        updateStudent({
          id: selectedStudent.id,
          name: editNameValue.trim()
        });
        setSelectedStudent({
          ...selectedStudent,
          name: editNameValue.trim()
        });
        showToast('이름이 수정되었습니다.', 'success');
      }
      setIsEditingName(false);
    }
  }, "\uC800\uC7A5"), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-secondary btn-sm",
    onClick: () => setIsEditingName(false)
  }, "\uCDE8\uC18C")) : /*#__PURE__*/React.createElement("div", {
    className: "name-display",
    style: {
      display: 'flex',
      alignItems: 'center',
      gap: '0.5rem'
    }
  }, /*#__PURE__*/React.createElement("h3", null, selectedStudent.name), isAdmin && /*#__PURE__*/React.createElement("button", {
    type: "button",
    onClick: () => {
      setEditNameValue(selectedStudent.name);
      setIsEditingName(true);
    },
    "aria-label": "\uC774\uB984 \uACE0\uCE58\uAE30",
    className: "icon-ghost"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "edit",
    size: 16
  })))), /*#__PURE__*/React.createElement("p", null, selectedStudent.number, "\uBC88 \xB7 ", selectedStudent.gender === 'M' ? '남학생' : '여학생', " \xB7 \uC0DD\uC77C: ", selectedStudent.birthday))), isAdmin && /*#__PURE__*/React.createElement("div", {
    className: "student-edit-section"
  }, /*#__PURE__*/React.createElement("h4", {
    className: "stu-modal-h"
  }, "\uC815\uBCF4 \uC218\uC815"), /*#__PURE__*/React.createElement("div", {
    className: "edit-fields-grid"
  }, /*#__PURE__*/React.createElement("div", {
    className: "edit-field"
  }, /*#__PURE__*/React.createElement("label", null, "\uC131\uBCC4"), /*#__PURE__*/React.createElement("select", {
    className: "form-input",
    value: editGender,
    onChange: e => setEditGender(e.target.value)
  }, /*#__PURE__*/React.createElement("option", {
    value: "M"
  }, "\uB0A8"), /*#__PURE__*/React.createElement("option", {
    value: "F"
  }, "\uC5EC"))), /*#__PURE__*/React.createElement("div", {
    className: "edit-field"
  }, /*#__PURE__*/React.createElement("label", null, "\uC0DD\uC77C"), /*#__PURE__*/React.createElement("input", {
    type: "date",
    className: "form-input",
    value: editBirthday,
    onChange: e => setEditBirthday(e.target.value)
  })), /*#__PURE__*/React.createElement("div", {
    className: "edit-field"
  }, /*#__PURE__*/React.createElement("label", null, "\uC5F0\uB77D\uCC98"), /*#__PURE__*/React.createElement("input", {
    type: "text",
    className: "form-input",
    value: editParentPhone,
    onChange: e => setEditParentPhone(e.target.value),
    placeholder: "010-0000-0000"
  })), /*#__PURE__*/React.createElement("div", {
    className: "edit-field",
    style: {
      display: 'flex',
      gap: '0.5rem',
      alignItems: 'flex-end'
    }
  }, /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary btn-sm",
    onClick: handleSaveStudentInfo
  }, "\uC800\uC7A5"), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-danger btn-sm",
    onClick: handleResetPoints
  }, "\uD3EC\uC778\uD2B8 \uCD08\uAE30\uD654")))), /*#__PURE__*/React.createElement("p", {
    className: "stu-modal-line"
  }, /*#__PURE__*/React.createElement("span", null, "\uD559\uBD80\uBAA8 \uC5F0\uB77D\uCC98"), /*#__PURE__*/React.createElement("strong", {
    className: "num"
  }, selectedStudent.parentPhone || '미등록')), isAdmin && /*#__PURE__*/React.createElement("div", {
    className: "record-input-section"
  }, /*#__PURE__*/React.createElement("h4", {
    className: "stu-modal-h"
  }, "\uAD00\uCC30\xB7\uC0C1\uB2F4 \uAE30\uB85D \uCD94\uAC00"), /*#__PURE__*/React.createElement("div", {
    className: "record-input-row"
  }, /*#__PURE__*/React.createElement("select", {
    className: "form-input",
    value: newRecordType,
    onChange: e => setNewRecordType(e.target.value),
    style: {
      width: 'auto',
      minWidth: '100px'
    }
  }, /*#__PURE__*/React.createElement("option", {
    value: "observation"
  }, "\uAD00\uCC30"), /*#__PURE__*/React.createElement("option", {
    value: "counseling"
  }, "\uC0C1\uB2F4")), /*#__PURE__*/React.createElement("select", {
    className: "form-input",
    value: newRecordCategory,
    onChange: e => setNewRecordCategory(e.target.value),
    style: {
      width: 'auto',
      minWidth: '80px'
    }
  }, ['생활', '학습', '교우', '건강', '기타'].map(c => /*#__PURE__*/React.createElement("option", {
    key: c,
    value: c
  }, c)))), /*#__PURE__*/React.createElement("div", {
    className: "record-input-row",
    style: {
      marginTop: '0.5rem'
    }
  }, /*#__PURE__*/React.createElement("textarea", {
    className: "form-input",
    placeholder: "\uAD00\uCC30/\uC0C1\uB2F4 \uB0B4\uC6A9\uC744 \uC785\uB825\uD558\uC138\uC694...",
    value: newRecordContent,
    onChange: e => setNewRecordContent(e.target.value),
    rows: "4",
    style: {
      resize: 'vertical',
      flex: 1
    }
  }), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary",
    onClick: handleAddRecord,
    style: {
      alignSelf: 'flex-end'
    }
  }, "\uC800\uC7A5"))), records.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "records-list-section"
  }, /*#__PURE__*/React.createElement("h4", {
    className: "stu-modal-h"
  }, "\uAE30\uB85D ", /*#__PURE__*/React.createElement("span", {
    className: "num"
  }, records.length), "\uAC74"), /*#__PURE__*/React.createElement("div", {
    className: "counseling-list",
    style: {
      marginTop: '0.5rem'
    }
  }, records.map((c, i) => /*#__PURE__*/React.createElement("div", {
    key: c.id || i,
    className: "counseling-item"
  }, /*#__PURE__*/React.createElement("div", {
    className: "counseling-date"
  }, c.date, " \xB7 ", /*#__PURE__*/React.createElement("span", {
    className: `record-type-badge ${c.type}`
  }, c.type === 'observation' ? '관찰' : c.type === 'counseling' ? '상담' : c.type), " \xB7 ", c.category, isAdmin && /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "icon-ghost",
    "aria-label": "\uC774 \uAE30\uB85D \uC0AD\uC81C",
    onClick: () => {
      if (confirm('이 기록을 삭제하시겠습니까?')) deleteStudentRecord(selectedStudent.id, c.id);
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "counseling-content"
  }, c.content)))))))), showBulkModal && /*#__PURE__*/React.createElement("div", {
    className: "modal-overlay",
    onClick: () => setShowBulkModal(false)
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    onClick: e => e.stopPropagation(),
    style: {
      maxWidth: '450px'
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-header"
  }, /*#__PURE__*/React.createElement("h3", {
    className: "modal-title"
  }, "\uBA85\uB2E8 \uD55C \uBC88\uC5D0 \uACE0\uCE58\uAE30"), /*#__PURE__*/React.createElement("button", {
    className: "modal-close",
    onClick: () => setShowBulkModal(false)
  }, "\xD7")), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: '0.8rem',
      color: 'var(--text-secondary)',
      marginBottom: '0.75rem'
    }
  }, "\uD55C \uC904\uC5D0 \uD55C \uBA85\uC529 \uD0ED(Tab)\uC774\uB098 \uC27C\uD45C(,)\uB85C \uAD6C\uBD84\uD558\uC5EC ", /*#__PURE__*/React.createElement("strong", null, "\uC774\uB984, \uC131\uBCC4(M/F), \uC0DD\uC77C(YYYY-MM-DD)"), "\uC744 \uC785\uB825\uD558\uC138\uC694.", /*#__PURE__*/React.createElement("br", null), "(\uC608: ", /*#__PURE__*/React.createElement("code", null, "\uAE40\uD559\uC0DD, M, 2014-05-01"), ")"), /*#__PURE__*/React.createElement("textarea", {
    className: "form-input",
    value: bulkNames,
    onChange: e => setBulkNames(e.target.value),
    rows: "15",
    style: {
      resize: 'vertical',
      fontFamily: 'monospace',
      lineHeight: '1.8'
    },
    placeholder: '이름\t성별\t생일'
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      display: 'flex',
      gap: '0.5rem',
      marginTop: '0.75rem',
      justifyContent: 'flex-end'
    }
  }, /*#__PURE__*/React.createElement("button", {
    className: "btn btn-secondary",
    onClick: () => setShowBulkModal(false)
  }, "\uCDE8\uC18C"), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary",
    onClick: handleBulkUpdate
  }, "\uC800\uC7A5"))))));
}

/* ===================================================================
   AttendancePage
   =================================================================== */
const STATUS_MAP = {
  present: {
    label: '출석',
    mark: '○'
  },
  absent: {
    label: '결석',
    mark: '✕'
  },
  late: {
    label: '지각',
    mark: '△'
  },
  'early-leave': {
    label: '조퇴',
    mark: '▽'
  }
};
const STATUS_KEYS = Object.keys(STATUS_MAP);
const sameAtt = (a, b) => (a?.status || 'present') === (b?.status || 'present') && (a?.note || '') === (b?.note || '');
function AttendancePage() {
  const {
    state,
    updateAttendance,
    loadAttendanceByDate,
    loadMonthlyAttendance,
    saveAttendanceDay,
    showToast
  } = useApp();
  const {
    students,
    attendance,
    selectedAttendanceDate,
    monthlyAttendance,
    isAdmin
  } = state;
  const [monthlyActiveCell, setMonthlyActiveCell] = useState(null);
  const [viewMode, setViewMode] = useState('daily');
  const [monthYear, setMonthYear] = useState({
    year: new Date().getFullYear(),
    month: new Date().getMonth()
  });
  // 입력 중인 출석(저장 전). 완료 버튼을 눌러야 db에 반영됨
  const [draft, setDraft] = useState({});
  const [selected, setSelected] = useState(() => new Set());
  const [noteFor, setNoteFor] = useState(null);
  const [noteText, setNoteText] = useState('');
  const baseOf = id => attendance[id] || {
    status: 'present',
    note: ''
  };
  const current = id => draft[id] || baseOf(id);
  const changedIds = students.filter(s => draft[s.id] && !sameAtt(draft[s.id], baseOf(s.id))).map(s => s.id);
  const dirty = changedIds.length > 0;
  useEffect(() => {
    setDraft({});
    setSelected(new Set());
    setNoteFor(null);
  }, [selectedAttendanceDate]);
  useEffect(() => {
    if (!dirty) return;
    const warn = e => {
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, [dirty]);
  const handleDateChange = e => {
    if (dirty && !confirm('저장하지 않은 출석 변경이 있어요. 날짜를 바꾸면 사라집니다. 계속할까요?')) return;
    loadAttendanceByDate(e.target.value);
  };
  const handleMonthlyStatus = (studentId, date, status) => {
    if (!isAdmin) {
      showToast('선생님 모드에서만 수정할 수 있습니다.', 'error');
      return;
    }
    updateAttendance(studentId, status, monthlyAttendance[date]?.[studentId]?.note || '', date);
    const s = students.find(st => st.id === studentId);
    showToast(`${s.name} (${date.slice(5).replace('-', '/')}) → ${STATUS_MAP[status].label}`, 'success');
    setMonthlyActiveCell(null);
  };
  useEffect(() => {
    if (viewMode === 'monthly') loadMonthlyAttendance(monthYear.year, monthYear.month);
  }, [viewMode, monthYear, loadMonthlyAttendance]);
  const counts = Object.fromEntries(STATUS_KEYS.map(k => [k, students.filter(s => current(s.id).status === k).length]));
  const toggleSelect = id => setSelected(prev => {
    const n = new Set(prev);
    n.has(id) ? n.delete(id) : n.add(id);
    return n;
  });
  const allSelected = students.length > 0 && selected.size === students.length;
  const toggleAll = () => setSelected(allSelected ? new Set() : new Set(students.map(s => s.id)));
  const applyStatus = status => {
    if (!selected.size) return;
    setDraft(prev => {
      const next = {
        ...prev
      };
      selected.forEach(id => {
        const cur = next[id] || baseOf(id);
        next[id] = {
          status,
          note: status === 'present' ? '' : cur.note || ''
        };
      });
      return next;
    });
    setSelected(new Set());
  };
  const openNote = id => {
    setNoteFor(id);
    setNoteText(current(id).note || '');
  };
  const saveNote = () => {
    setDraft(prev => ({
      ...prev,
      [noteFor]: {
        ...current(noteFor),
        note: noteText.trim()
      }
    }));
    setNoteFor(null);
  };
  const discard = () => {
    setDraft({});
    setSelected(new Set());
    setNoteFor(null);
  };
  const handleSubmit = () => {
    if (!isAdmin) return;
    const records = {};
    students.forEach(s => {
      const a = current(s.id);
      records[s.id] = {
        status: a.status,
        note: a.status === 'present' ? '' : a.note || '',
        id: baseOf(s.id).id
      };
    });
    saveAttendanceDay(selectedAttendanceDate, records);
    setDraft({});
    setSelected(new Set());
    setNoteFor(null);
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
  return /*#__PURE__*/React.createElement("div", {
    className: "att"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wk-top"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wk-title"
  }, /*#__PURE__*/React.createElement("h1", {
    className: "dash-headline"
  }, viewMode === 'daily' ? `${selDate.getMonth() + 1}월 ${selDate.getDate()}일 ${getDayName(selDate)}요일` : `${monthYear.year}년 ${monthYear.month + 1}월`), viewMode === 'daily' && /*#__PURE__*/React.createElement("span", {
    className: `att-saved ${isSavedDay && !dirty ? 'is-saved' : ''}`
  }, dirty ? `저장 안 됨 · ${changedIds.length}명 변경` : isSavedDay ? '저장됨' : '아직 입력 전')), /*#__PURE__*/React.createElement("div", {
    className: "mode-switch",
    role: "group",
    "aria-label": "\uBCF4\uAE30"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    "aria-pressed": viewMode === 'daily',
    onClick: () => setViewMode('daily')
  }, "\uC77C\uBCC4 \uCD9C\uC11D"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    "aria-pressed": viewMode === 'monthly',
    onClick: () => setViewMode('monthly')
  }, "\uC6D4\uBCC4 \uD604\uD669"))), viewMode === 'daily' ? /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    className: "att-bar"
  }, /*#__PURE__*/React.createElement("label", {
    className: "att-date"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "calendar",
    size: 18
  }), /*#__PURE__*/React.createElement("span", {
    className: "sr-only"
  }, "\uB0A0\uC9DC"), /*#__PURE__*/React.createElement("input", {
    type: "date",
    value: selectedAttendanceDate,
    onChange: handleDateChange
  })), /*#__PURE__*/React.createElement("div", {
    className: "att-counts"
  }, STATUS_KEYS.map(k => /*#__PURE__*/React.createElement("span", {
    key: k,
    className: `att-count is-${k}`
  }, /*#__PURE__*/React.createElement("span", {
    className: "att-dot"
  }), STATUS_MAP[k].label, /*#__PURE__*/React.createElement("strong", {
    className: "num"
  }, counts[k]))))), isAdmin && /*#__PURE__*/React.createElement("div", {
    className: "att-tools"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: `att-check-all ${allSelected ? 'is-on' : ''}`,
    "aria-pressed": allSelected,
    onClick: toggleAll,
    disabled: !students.length
  }, /*#__PURE__*/React.createElement("span", {
    className: "att-box",
    "aria-hidden": "true"
  }), allSelected ? '선택 해제' : '모두 선택'), /*#__PURE__*/React.createElement("span", {
    className: "att-tools-label"
  }, selected.size ? /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("strong", {
    className: "num"
  }, selected.size), "\uBA85\uC744") : '학생을 고른 뒤'), /*#__PURE__*/React.createElement("div", {
    className: "att-apply"
  }, STATUS_KEYS.map(k => /*#__PURE__*/React.createElement("button", {
    key: k,
    type: "button",
    className: `att-apply-btn is-${k}`,
    disabled: !selected.size,
    onClick: () => applyStatus(k)
  }, STATUS_MAP[k].label)))), students.length === 0 ? /*#__PURE__*/React.createElement("div", {
    className: "wcard-empty att-empty"
  }, "\uD559\uC0DD \uAD00\uB9AC\uC5D0\uC11C \uD559\uC0DD\uC744 \uB4F1\uB85D\uD558\uBA74 \uCD9C\uC11D\uBD80\uAC00 \uB9CC\uB4E4\uC5B4\uC838\uC694.") : /*#__PURE__*/React.createElement("ul", {
    className: "att-grid"
  }, students.map(s => {
    const a = current(s.id);
    const isSel = selected.has(s.id);
    const changed = changedIds.includes(s.id);
    const needsNote = a.status !== 'present';
    return /*#__PURE__*/React.createElement("li", {
      key: s.id,
      className: `att-card is-${a.status} ${isSel ? 'is-selected' : ''}`
    }, /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: "att-card-main",
      "aria-pressed": isSel,
      disabled: !isAdmin,
      onClick: () => toggleSelect(s.id)
    }, isAdmin && /*#__PURE__*/React.createElement("span", {
      className: "att-box",
      "aria-hidden": "true"
    }), /*#__PURE__*/React.createElement("span", {
      className: "att-no num"
    }, s.number), /*#__PURE__*/React.createElement("span", {
      className: "att-name"
    }, s.name), /*#__PURE__*/React.createElement("span", {
      className: "att-status"
    }, STATUS_MAP[a.status].label, changed && /*#__PURE__*/React.createElement("span", {
      className: "att-changed",
      "aria-label": "\uBCC0\uACBD\uB428"
    }))), needsNote && (isAdmin ? /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: `att-note-btn ${a.note ? 'has-note' : ''}`,
      onClick: () => openNote(s.id)
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "edit",
      size: 14
    }), a.note || '사유 쓰기') : a.note ? /*#__PURE__*/React.createElement("span", {
      className: "att-note-btn has-note"
    }, a.note) : null));
  })), isAdmin && /*#__PURE__*/React.createElement("div", {
    className: `att-submit ${dirty ? 'is-dirty' : ''}`
  }, /*#__PURE__*/React.createElement("span", {
    className: "att-submit-text"
  }, dirty ? /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("strong", {
    className: "num"
  }, changedIds.length), "\uBA85\uC758 \uCD9C\uC11D\uC774 \uBC14\uB00C\uC5C8\uC5B4\uC694. \uC644\uB8CC\uB97C \uB20C\uB7EC\uC57C \uC800\uC7A5\uB3FC\uC694.") : isSavedDay ? '이 날의 출석이 저장되어 있어요.' : '출석 상태를 확인하고 완료를 눌러 주세요.'), /*#__PURE__*/React.createElement("div", {
    className: "att-submit-actions"
  }, dirty && /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-secondary",
    onClick: discard
  }, "\uB418\uB3CC\uB9AC\uAE30"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-primary",
    onClick: handleSubmit,
    disabled: !dirty && isSavedDay
  }, "\uCD9C\uC11D \uC785\uB825 \uC644\uB8CC"))), noteStudent && /*#__PURE__*/React.createElement("div", {
    className: "modal-overlay",
    onClick: () => setNoteFor(null)
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-content att-note-modal",
    role: "dialog",
    "aria-modal": "true",
    "aria-labelledby": "att-note-title",
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-header"
  }, /*#__PURE__*/React.createElement("h3", {
    id: "att-note-title"
  }, noteStudent.name, " \xB7 ", STATUS_MAP[current(noteFor).status].label, " \uC0AC\uC720"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "close-btn",
    "aria-label": "\uB2EB\uAE30",
    onClick: () => setNoteFor(null)
  }, "\xD7")), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("label", {
    htmlFor: "att-note-text",
    className: "sr-only"
  }, "\uC0AC\uC720"), /*#__PURE__*/React.createElement("textarea", {
    id: "att-note-text",
    className: "form-input",
    rows: "3",
    autoFocus: true,
    placeholder: "\uC608: \uAC10\uAE30\uB85C \uBCD1\uC6D0 \uC9C4\uB8CC",
    value: noteText,
    onChange: e => setNoteText(e.target.value)
  }), /*#__PURE__*/React.createElement("div", {
    className: "att-note-presets"
  }, ['질병', '병원 진료', '가정 사정', '체험학습', '늦잠'].map(p => /*#__PURE__*/React.createElement("button", {
    key: p,
    type: "button",
    className: "chip-btn",
    onClick: () => setNoteText(p)
  }, p)))), /*#__PURE__*/React.createElement("div", {
    className: "modal-footer att-note-footer"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-secondary",
    onClick: () => setNoteFor(null)
  }, "\uCDE8\uC18C"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-primary",
    onClick: saveNote
  }, "\uC0AC\uC720 \uC801\uAE30"))))) : /*#__PURE__*/React.createElement("div", {
    className: "monthly-attendance"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wk-nav att-month-nav"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "round-btn",
    "aria-label": "\uC774\uC804 \uB2EC",
    onClick: () => {
      if (monthYear.month === 0) setMonthYear({
        year: monthYear.year - 1,
        month: 11
      });else setMonthYear({
        ...monthYear,
        month: monthYear.month - 1
      });
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "chevron",
    size: 18
  })), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "round-btn wk-next",
    "aria-label": "\uB2E4\uC74C \uB2EC",
    onClick: () => {
      if (monthYear.month === 11) setMonthYear({
        year: monthYear.year + 1,
        month: 0
      });else setMonthYear({
        ...monthYear,
        month: monthYear.month + 1
      });
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "chevron",
    size: 18
  })), /*#__PURE__*/React.createElement("div", {
    className: "monthly-legend"
  }, STATUS_KEYS.map(k => /*#__PURE__*/React.createElement("span", {
    key: k,
    className: "legend-item"
  }, /*#__PURE__*/React.createElement("span", {
    className: `legend-icon ${k}`
  }, STATUS_MAP[k].mark), STATUS_MAP[k].label)), /*#__PURE__*/React.createElement("span", {
    className: "legend-item"
  }, /*#__PURE__*/React.createElement("span", {
    className: "legend-icon no-data"
  }, "-"), "\uAE30\uB85D \uC5C6\uC74C"))), /*#__PURE__*/React.createElement("div", {
    className: "monthly-table-wrapper"
  }, /*#__PURE__*/React.createElement("table", {
    className: "monthly-table"
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", {
    className: "sticky-col"
  }, "\uBC88\uD638"), /*#__PURE__*/React.createElement("th", {
    className: "sticky-col-name"
  }, "\uC774\uB984"), monthDays.map(d => {
    const date = new Date(d);
    return /*#__PURE__*/React.createElement("th", {
      key: d,
      className: "date-col"
    }, /*#__PURE__*/React.createElement("div", {
      className: "date-header"
    }, /*#__PURE__*/React.createElement("span", {
      className: "num"
    }, date.getDate()), /*#__PURE__*/React.createElement("span", {
      className: "date-dow"
    }, getDayName(date))));
  }))), /*#__PURE__*/React.createElement("tbody", null, students.map(s => /*#__PURE__*/React.createElement("tr", {
    key: s.id
  }, /*#__PURE__*/React.createElement("td", {
    className: "sticky-col num"
  }, s.number), /*#__PURE__*/React.createElement("td", {
    className: "sticky-col-name"
  }, s.name), monthDays.map(d => {
    const record = monthlyAttendance[d]?.[s.id];
    const status = record?.status || null;
    const isCellActive = monthlyActiveCell?.studentId === s.id && monthlyActiveCell?.date === d;
    return /*#__PURE__*/React.createElement("td", {
      key: d,
      className: `monthly-cell ${status || 'no-data'} ${isCellActive ? 'active' : ''}`,
      title: record?.note || '',
      onClick: () => {
        if (isAdmin) setMonthlyActiveCell(isCellActive ? null : {
          studentId: s.id,
          date: d
        });
      },
      style: {
        position: 'relative',
        cursor: isAdmin ? 'pointer' : 'default'
      }
    }, /*#__PURE__*/React.createElement("span", {
      className: `monthly-status ${status || 'no-data'}`
    }, status ? STATUS_MAP[status].mark : '-'), record?.note && /*#__PURE__*/React.createElement("span", {
      className: "monthly-note-dot",
      "aria-label": `사유: ${record.note}`
    }), isAdmin && isCellActive && /*#__PURE__*/React.createElement("div", {
      className: "attendance-status-selector show monthly-selector",
      onClick: e => e.stopPropagation()
    }, STATUS_KEYS.map(k => /*#__PURE__*/React.createElement("button", {
      key: k,
      type: "button",
      className: `status-option ${status === k ? 'active' : ''}`,
      onClick: () => handleMonthlyStatus(s.id, d, k)
    }, STATUS_MAP[k].label))));
  }))))))));
}

/* ===================================================================
   LearningPage
   =================================================================== */
function LearningPage() {
  const {
    state,
    showToast,
    addAssignment,
    updateAssignmentSubmissions,
    deleteAssignment
  } = useApp();
  const {
    students,
    assignments,
    isAdmin
  } = state;
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
  const handlePointerDown = studentId => {
    if (!activeActivity) return;
    timerRef.current = setTimeout(() => {
      if (activeActivity.submissions?.[studentId]) updateAssignmentSubmissions(activeActivity.id, {
        [studentId]: null
      });
      timerRef.current = null;
    }, 500);
  };
  const handlePointerUp = studentId => {
    if (!activeActivity) return;
    if (timerRef.current) {
      clearTimeout(timerRef.current);
      timerRef.current = null;
      if (!activeActivity.submissions?.[studentId]) updateAssignmentSubmissions(activeActivity.id, {
        [studentId]: {
          timestamp: new Date().toISOString()
        }
      });
    }
  };
  const handlePointerLeave = () => {
    if (timerRef.current) {
      clearTimeout(timerRef.current);
      timerRef.current = null;
    }
  };

  // 키보드(Enter/Space)로도 체크·해제할 수 있게
  const toggleByKey = studentId => {
    if (!activeActivity) return;
    updateAssignmentSubmissions(activeActivity.id, {
      [studentId]: activeActivity.submissions?.[studentId] ? null : {
        timestamp: new Date().toISOString()
      }
    });
  };
  const closeForm = () => {
    setTempActivityContent('');
    setTempActivityType('과제');
    setIsEditingActivity(false);
  };
  const saveActivity = () => {
    if (tempActivityContent.trim() !== '') {
      addAssignment({
        id: generateUUID(),
        title: tempActivityContent,
        subject: '활동',
        type: tempActivityType,
        dueDate: toISODate(new Date()),
        createdAt: new Date().toISOString(),
        submissions: {},
        total: students.length
      });
      showToast('새로운 활동이 추가되었습니다.', 'success');
      setSelectedAssignment(null);
      setCurrentPage(1);
    }
    closeForm();
  };
  const sortedAssignments = [...assignments].reverse();
  const totalPages = Math.ceil(sortedAssignments.length / itemsPerPage) || 1;
  const startIndex = (currentPage - 1) * itemsPerPage;
  const displayAssignments = sortedAssignments.slice(startIndex, startIndex + itemsPerPage);
  const doneCount = activeActivity ? students.filter(s => activeActivity.submissions?.[s.id]).length : 0;
  const donePct = students.length ? Math.round(doneCount / students.length * 100) : 0;
  return /*#__PURE__*/React.createElement("div", {
    className: "lrn"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wk-top"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wk-title"
  }, /*#__PURE__*/React.createElement("h1", {
    className: "dash-headline"
  }, activeActivity ? /*#__PURE__*/React.createElement(React.Fragment, null, activeActivity.title) : '오늘 할 활동을 정해요'), activeActivity && /*#__PURE__*/React.createElement("span", {
    className: `lrn-type ${activeActivity.type === '수행평가' ? 'is-eval' : ''}`
  }, activeActivity.type || '수행평가'), selectedAssignment && /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "chip-btn",
    onClick: () => setSelectedAssignment(null)
  }, "\uCD5C\uC2E0 \uD65C\uB3D9\uC73C\uB85C")), isAdmin && !isEditingActivity && /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-primary",
    onClick: () => {
      closeForm();
      setIsEditingActivity(true);
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 18
  }), " \uD65C\uB3D9 \uCD94\uAC00")), isAdmin && isEditingActivity && /*#__PURE__*/React.createElement("form", {
    className: "lrn-form",
    onSubmit: e => {
      e.preventDefault();
      saveActivity();
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "mode-switch",
    role: "group",
    "aria-label": "\uC885\uB958"
  }, ['과제', '수행평가'].map(t => /*#__PURE__*/React.createElement("button", {
    key: t,
    type: "button",
    "aria-pressed": tempActivityType === t,
    onClick: () => setTempActivityType(t)
  }, t))), /*#__PURE__*/React.createElement("input", {
    type: "text",
    className: "form-input",
    "aria-label": "\uD65C\uB3D9 \uB0B4\uC6A9",
    value: tempActivityContent,
    onChange: e => setTempActivityContent(e.target.value),
    placeholder: "\uC608: \uC218\uD559 \uC775\uD798\uCC45 34\uCABD",
    autoFocus: true
  }), /*#__PURE__*/React.createElement("div", {
    className: "lrn-form-actions"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-secondary btn-sm",
    onClick: closeForm
  }, "\uCDE8\uC18C"), /*#__PURE__*/React.createElement("button", {
    type: "submit",
    className: "btn btn-primary btn-sm"
  }, "\uC800\uC7A5"))), /*#__PURE__*/React.createElement("section", {
    className: "wcard lrn-check"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wcard-head"
  }, /*#__PURE__*/React.createElement("h2", {
    className: "wcard-title"
  }, "\uD65C\uB3D9 \uC644\uB8CC \uCCB4\uD06C"), activeActivity && /*#__PURE__*/React.createElement("span", {
    className: "wcard-meta"
  }, /*#__PURE__*/React.createElement("strong", {
    className: "num lrn-done"
  }, doneCount), " / ", /*#__PURE__*/React.createElement("span", {
    className: "num"
  }, students.length), "\uBA85 \uC644\uB8CC")), /*#__PURE__*/React.createElement("div", {
    className: "wcard-body"
  }, activeActivity ? /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    className: "lrn-track",
    role: "progressbar",
    "aria-label": "\uC644\uB8CC\uC728",
    "aria-valuenow": donePct,
    "aria-valuemin": 0,
    "aria-valuemax": 100
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      width: `${donePct}%`
    },
    className: donePct === 100 ? 'is-done' : ''
  })), /*#__PURE__*/React.createElement("p", {
    className: "lrn-hint"
  }, "\uC774\uB984\uC744 \uB204\uB974\uBA74 \uC644\uB8CC, \uAE38\uAC8C \uB204\uB974\uBA74 \uCDE8\uC18C\uB3FC\uC694."), /*#__PURE__*/React.createElement("ul", {
    className: "lrn-grid"
  }, students.map(student => {
    const isCompleted = !!activeActivity.submissions?.[student.id];
    return /*#__PURE__*/React.createElement("li", {
      key: student.id
    }, /*#__PURE__*/React.createElement("button", {
      type: "button",
      "aria-pressed": isCompleted,
      className: `lrn-student ${isCompleted ? 'is-done' : ''}`,
      onPointerDown: () => handlePointerDown(student.id),
      onPointerUp: () => handlePointerUp(student.id),
      onPointerLeave: handlePointerLeave,
      onContextMenu: e => e.preventDefault(),
      onKeyDown: e => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          toggleByKey(student.id);
        }
      }
    }, /*#__PURE__*/React.createElement("span", {
      className: "att-box"
    }), /*#__PURE__*/React.createElement("span", {
      className: "att-no num"
    }, student.number), /*#__PURE__*/React.createElement("span", {
      className: "lrn-name"
    }, student.name)));
  }))) : /*#__PURE__*/React.createElement("div", {
    className: "wcard-empty"
  }, isAdmin ? '오른쪽 위 “활동 추가”로 새 활동을 만들어 주세요.' : '아직 등록된 활동이 없어요.'))), /*#__PURE__*/React.createElement("section", {
    className: "wcard"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wcard-head"
  }, /*#__PURE__*/React.createElement("h2", {
    className: "wcard-title"
  }, "\uACFC\uC81C\xB7\uC218\uD589\uD3C9\uAC00 \uD604\uD669"), /*#__PURE__*/React.createElement("span", {
    className: "wcard-meta"
  }, "\uB20C\uB7EC\uC11C \uC81C\uCD9C \uBA85\uB2E8 \uBCF4\uAE30")), /*#__PURE__*/React.createElement("div", {
    className: "wcard-body"
  }, displayAssignments.length > 0 ? /*#__PURE__*/React.createElement("ul", {
    className: "lrn-list"
  }, displayAssignments.map(a => {
    const total = a.total || students.length;
    const submittedCount = typeof a.submitted === 'number' ? a.submitted : Object.keys(a.submissions || {}).length;
    const pct = total ? Math.round(submittedCount / total * 100) : 0;
    const remaining = total - submittedCount;
    const isOpen = selectedAssignment?.id === a.id;
    const done = students.filter(s => a.submissions?.[s.id]);
    const notDone = students.filter(s => !a.submissions?.[s.id]);
    return /*#__PURE__*/React.createElement("li", {
      key: a.id,
      className: `lrn-item ${isOpen ? 'is-open' : ''}`
    }, /*#__PURE__*/React.createElement("div", {
      className: "lrn-row"
    }, /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: "lrn-row-main",
      "aria-expanded": isOpen,
      onClick: () => setSelectedAssignment(isOpen ? null : a.id)
    }, /*#__PURE__*/React.createElement("span", {
      className: `lrn-type ${a.type === '수행평가' ? 'is-eval' : ''}`
    }, a.type), /*#__PURE__*/React.createElement("span", {
      className: "lrn-row-text"
    }, /*#__PURE__*/React.createElement("span", {
      className: "lrn-row-title"
    }, a.title), /*#__PURE__*/React.createElement("span", {
      className: "lrn-row-meta"
    }, /*#__PURE__*/React.createElement("span", {
      className: "num"
    }, a.dueDate), remaining > 0 ? /*#__PURE__*/React.createElement("span", {
      className: "lrn-left"
    }, "\uBBF8\uC81C\uCD9C ", remaining, "\uBA85") : /*#__PURE__*/React.createElement("span", {
      className: "lrn-all"
    }, "\uBAA8\uB450 \uC81C\uCD9C"))), /*#__PURE__*/React.createElement("span", {
      className: "lrn-row-bar"
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        width: `${pct}%`
      },
      className: pct === 100 ? 'is-done' : ''
    })), /*#__PURE__*/React.createElement("span", {
      className: "lrn-pct num"
    }, pct, /*#__PURE__*/React.createElement("small", null, "%"))), isAdmin && /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: "icon-ghost",
      "aria-label": `${a.title} 삭제`,
      onClick: () => {
        if (confirm('이 항목을 정말 삭제하시겠습니까?')) {
          deleteAssignment(a.id);
          if (isOpen) setSelectedAssignment(null);
          showToast('삭제되었습니다.', 'success');
        }
      }
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "trash",
      size: 18
    }))), isOpen && /*#__PURE__*/React.createElement("div", {
      className: "lrn-who"
    }, /*#__PURE__*/React.createElement("div", {
      className: "lrn-who-group"
    }, /*#__PURE__*/React.createElement("span", {
      className: "lrn-who-label is-done"
    }, "\uC81C\uCD9C ", /*#__PURE__*/React.createElement("span", {
      className: "num"
    }, done.length)), /*#__PURE__*/React.createElement("div", {
      className: "lrn-chips"
    }, done.length ? done.map(s => /*#__PURE__*/React.createElement("span", {
      key: s.id,
      className: "lrn-chip is-done"
    }, /*#__PURE__*/React.createElement("span", {
      className: "num"
    }, s.number), s.name)) : /*#__PURE__*/React.createElement("span", {
      className: "lrn-none"
    }, "\uC544\uC9C1 \uC5C6\uC5B4\uC694"))), /*#__PURE__*/React.createElement("div", {
      className: "lrn-who-group"
    }, /*#__PURE__*/React.createElement("span", {
      className: "lrn-who-label is-left"
    }, "\uBBF8\uC81C\uCD9C ", /*#__PURE__*/React.createElement("span", {
      className: "num"
    }, notDone.length)), /*#__PURE__*/React.createElement("div", {
      className: "lrn-chips"
    }, notDone.length ? notDone.map(s => /*#__PURE__*/React.createElement("span", {
      key: s.id,
      className: "lrn-chip is-left"
    }, /*#__PURE__*/React.createElement("span", {
      className: "num"
    }, s.number), s.name)) : /*#__PURE__*/React.createElement("span", {
      className: "lrn-none"
    }, "\uBAA8\uB450 \uB0C8\uC5B4\uC694")))));
  })) : /*#__PURE__*/React.createElement("div", {
    className: "wcard-empty"
  }, "\uB4F1\uB85D\uB41C \uACFC\uC81C\uAC00 \uC5C6\uC5B4\uC694. \uD65C\uB3D9\uC744 \uCD94\uAC00\uD558\uBA74 \uC5EC\uAE30\uC5D0 \uC313\uC5EC\uC694."), totalPages > 1 && /*#__PURE__*/React.createElement("div", {
    className: "lrn-pager"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "round-btn round-btn--sm",
    "aria-label": "\uC774\uC804 \uCABD",
    disabled: currentPage === 1,
    onClick: () => setCurrentPage(p => Math.max(1, p - 1))
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "chevron",
    size: 18
  })), /*#__PURE__*/React.createElement("span", {
    className: "num"
  }, currentPage, " / ", totalPages), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "round-btn round-btn--sm wk-next",
    "aria-label": "\uB2E4\uC74C \uCABD",
    disabled: currentPage === totalPages,
    onClick: () => setCurrentPage(p => Math.min(totalPages, p + 1))
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "chevron",
    size: 18
  }))))));
}

/* ===================================================================
   CalendarPage
   =================================================================== */
function CalendarPage() {
  const {
    state
  } = useApp();
  const {
    isAdmin,
    settings
  } = state;
  const calendars = [settings?.calendarId1, settings?.calendarId2, settings?.calendarId3].filter(Boolean);
  const getCombinedCalendarSrc = cals => {
    if (cals.length === 0) return '';
    let baseUrl = 'https://calendar.google.com/calendar/embed?ctz=Asia%2FSeoul&showPrint=0&showCalendars=0&showTz=0&mode=MONTH';
    const colors = ['%23039BE5', '%23D50000', '%230B8043'];
    cals.forEach((input, index) => {
      const srcMatch = input.match(/src="([^"]+)"/);
      let extracted = srcMatch ? srcMatch[1] : input.trim();
      let calId = '';
      if (extracted.includes('calendar.google.com') && extracted.includes('src=')) {
        const m = extracted.match(/src=([^&]+)/);
        if (m) calId = decodeURIComponent(m[1]);
      } else calId = extracted;
      if (calId) baseUrl += `&src=${encodeURIComponent(calId)}&color=${colors[index % colors.length]}`;
    });
    return baseUrl;
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "calendar-page",
    style: {
      height: '100%'
    }
  }, calendars.length > 0 ? /*#__PURE__*/React.createElement("div", {
    className: "card google-calendar-embed",
    style: {
      height: '100%',
      padding: '0'
    }
  }, /*#__PURE__*/React.createElement("iframe", {
    src: getCombinedCalendarSrc(calendars),
    width: "100%",
    height: "800",
    frameBorder: "0",
    title: "Google Calendar Combined",
    style: {
      display: 'block',
      border: 'none'
    }
  })) : /*#__PURE__*/React.createElement("div", {
    className: "card google-calendar-embed",
    style: {
      padding: '3rem',
      textAlign: 'center',
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      minHeight: '800px'
    }
  }, /*#__PURE__*/React.createElement("h3", null, "\uD83D\uDCC5 \uAD6C\uAE00 \uCE98\uB9B0\uB354\uAC00 \uC5F0\uB3D9\uB418\uC9C0 \uC54A\uC558\uC2B5\uB2C8\uB2E4."), isAdmin ? /*#__PURE__*/React.createElement("p", {
    style: {
      color: 'var(--text-secondary)',
      marginTop: '0.5rem'
    }
  }, "\uC6B0\uCE21 \uC0C1\uB2E8\uC758 \u2699\uFE0F \uC124\uC815 \uBA54\uB274\uC5D0\uC11C \uAD6C\uAE00 \uCE98\uB9B0\uB354 ID\uB97C \uC785\uB825\uD574\uC8FC\uC138\uC694.") : /*#__PURE__*/React.createElement("p", {
    style: {
      color: 'var(--text-secondary)',
      marginTop: '0.5rem'
    }
  }, "\uAD00\uB9AC\uC790\uAC00 \uCE98\uB9B0\uB354\uB97C \uC5F0\uB3D9\uD558\uBA74 \uC774 \uACF3\uC5D0 \uD45C\uC2DC\uB429\uB2C8\uB2E4.")));
}

/* ===================================================================
   ToolsPage
   =================================================================== */
function ToolsPage() {
  const {
    state,
    showToast,
    addPoll,
    deletePoll,
    votePoll
  } = useApp();
  const {
    students,
    polls,
    isAdmin
  } = state;
  const [pickerResult, setPickerResult] = useState('?');
  const [pickerHistory, setPickerHistory] = useState([]);
  const [isSpinning, setIsSpinning] = useState(false);
  const [pickerCount, setPickerCount] = useState(1);
  const spinRef = useRef(null);
  const handlePick = useCallback(() => {
    if (isSpinning) return;
    const available = students.filter(s => !pickerHistory.find(h => h.id === s.id));
    if (available.length === 0) {
      showToast('모든 학생이 이미 뽑혔습니다!', 'warning');
      return;
    }
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
  const resetPicker = () => {
    setPickerHistory([]);
    setPickerResult('?');
  };
  const [timerMode, setTimerMode] = useState('timer');
  const [timerSeconds, setTimerSeconds] = useState(300);
  const [timerRunning, setTimerRunning] = useState(false);
  const [timerElapsed, setTimerElapsed] = useState(0);
  const [isAlarm, setIsAlarm] = useState(false);
  const [customMin, setCustomMin] = useState('');
  const [customSec, setCustomSec] = useState('');
  const timerRef = useRef(null);
  const presets = [{
    label: '1분',
    sec: 60
  }, {
    label: '3분',
    sec: 180
  }, {
    label: '5분',
    sec: 300
  }, {
    label: '10분',
    sec: 600
  }, {
    label: '15분',
    sec: 900
  }];
  const formatTimer = sec => `${String(Math.floor(sec / 60)).padStart(2, '0')}:${String(sec % 60).padStart(2, '0')}`;
  const formatStopwatch = ms => {
    const totalSec = Math.floor(ms / 1000);
    return `${String(Math.floor(totalSec / 60)).padStart(2, '0')}:${String(totalSec % 60).padStart(2, '0')}.${String(Math.floor(ms % 1000 / 10)).padStart(2, '0')}`;
  };
  const startTimer = () => {
    if (timerRunning) return;
    setIsAlarm(false);
    setTimerRunning(true);
    if (timerMode === 'timer') {
      timerRef.current = setInterval(() => {
        setTimerSeconds(prev => {
          if (prev <= 1) {
            clearInterval(timerRef.current);
            setTimerRunning(false);
            setIsAlarm(true);
            showToast('⏰ 시간 종료!', 'warning');
            return 0;
          }
          return prev - 1;
        });
      }, 1000);
    } else {
      const startTime = Date.now() - timerElapsed;
      timerRef.current = setInterval(() => setTimerElapsed(Date.now() - startTime), 10);
    }
  };
  const pauseTimer = () => {
    clearInterval(timerRef.current);
    setTimerRunning(false);
  };
  const resetTimer = () => {
    clearInterval(timerRef.current);
    setTimerRunning(false);
    setTimerSeconds(300);
    setTimerElapsed(0);
    setIsAlarm(false);
  };
  const applyCustomTime = () => {
    const m = parseInt(customMin) || 0,
      s = parseInt(customSec) || 0;
    if (m === 0 && s === 0) {
      showToast('시간을 입력해주세요', 'error');
      return;
    }
    setTimerSeconds(m * 60 + s);
    showToast(`${m}분 ${s}초로 설정되었습니다`, 'success');
  };
  useEffect(() => () => {
    clearInterval(timerRef.current);
    clearInterval(spinRef.current);
  }, []);
  const [groupCount, setGroupCount] = useState(4);
  const [groups, setGroups] = useState([]);
  const makeGroups = () => {
    const shuffled = shuffleArray(students);
    const result = Array.from({
      length: groupCount
    }, () => []);
    shuffled.forEach((s, i) => result[i % groupCount].push(s));
    setGroups(result);
    showToast(`${groupCount}개 팀이 편성되었습니다!`, 'success');
  };
  const [newPollQ, setNewPollQ] = useState('');
  const [newPollOpts, setNewPollOpts] = useState(['', '']);
  const getVotedPolls = () => {
    try {
      return JSON.parse(sessionStorage.getItem('votedPolls') || '{}');
    } catch {
      return {};
    }
  };
  const markPollVoted = pollId => {
    const voted = getVotedPolls();
    voted[pollId] = true;
    sessionStorage.setItem('votedPolls', JSON.stringify(voted));
  };
  const hasVoted = pollId => !!getVotedPolls()[pollId];
  const handleAddPoll = () => {
    if (!newPollQ.trim() || newPollOpts.filter(o => o.trim()).length < 2) {
      showToast('질문과 최소 2개 항목을 입력하세요', 'error');
      return;
    }
    const opts = newPollOpts.filter(o => o.trim());
    addPoll({
      id: generateUUID(),
      question: newPollQ,
      options: opts,
      votes: opts.map(() => 0),
      isActive: true
    });
    setNewPollQ('');
    setNewPollOpts(['', '']);
    showToast('투표가 생성되었습니다!', 'success');
  };
  const handleVote = (pollId, optIdx) => {
    if (hasVoted(pollId)) {
      showToast('이미 투표하셨습니다!', 'warning');
      return;
    }
    votePoll(pollId, optIdx);
    markPollVoted(pollId);
    showToast('투표가 완료되었습니다!', 'success');
  };
  const handleDeletePoll = pollId => {
    deletePoll(pollId);
    showToast('투표가 삭제되었습니다.', 'success');
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "tools-page"
  }, /*#__PURE__*/React.createElement("div", {
    className: "tools-grid"
  }, /*#__PURE__*/React.createElement("div", {
    className: `tool-card ${!isAdmin ? 'disabled-tool' : ''}`
  }, /*#__PURE__*/React.createElement("div", {
    className: "tool-card-header"
  }, /*#__PURE__*/React.createElement("span", {
    className: "tool-emoji"
  }, "\uD83C\uDFB2"), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "tool-title"
  }, "\uB79C\uB364 \uBF51\uAE30"), /*#__PURE__*/React.createElement("div", {
    className: "tool-desc"
  }, "\uBC1C\uD45C\uC790\uB97C \uB79C\uB364\uC73C\uB85C \uC120\uD0DD\uD574\uC694"))), /*#__PURE__*/React.createElement("div", {
    className: "picker-container"
  }, /*#__PURE__*/React.createElement("div", {
    className: "roulette-wheel"
  }, /*#__PURE__*/React.createElement("div", {
    className: "roulette-pointer"
  }), /*#__PURE__*/React.createElement("div", {
    className: `roulette-circle ${isSpinning ? 'spinning' : pickerResult !== '?' ? 'reveal' : ''}`
  }, /*#__PURE__*/React.createElement("div", {
    className: "roulette-result"
  }, pickerResult))), isAdmin && /*#__PURE__*/React.createElement("div", {
    className: "picker-count"
  }, /*#__PURE__*/React.createElement("label", null, "\uBF51\uC744 \uC778\uC6D0:"), /*#__PURE__*/React.createElement("input", {
    type: "number",
    min: "1",
    max: "28",
    value: pickerCount,
    onChange: e => setPickerCount(Math.max(1, parseInt(e.target.value) || 1))
  }), /*#__PURE__*/React.createElement("span", null, "\uBA85")), isAdmin && /*#__PURE__*/React.createElement("div", {
    style: {
      display: 'flex',
      gap: '0.5rem',
      justifyContent: 'center'
    }
  }, /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary btn-lg",
    onClick: handlePick,
    disabled: isSpinning
  }, isSpinning ? '🎰 뽑는 중...' : '🎲 뽑기!'), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-secondary",
    onClick: resetPicker
  }, "\uCD08\uAE30\uD654")), pickerHistory.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "picker-history"
  }, pickerHistory.map((s, i) => /*#__PURE__*/React.createElement("span", {
    key: i,
    className: "picked-badge"
  }, s.name))))), /*#__PURE__*/React.createElement("div", {
    className: `tool-card ${!isAdmin ? 'disabled-tool' : ''}`
  }, /*#__PURE__*/React.createElement("div", {
    className: "tool-card-header"
  }, /*#__PURE__*/React.createElement("span", {
    className: "tool-emoji"
  }, "\u23F1\uFE0F"), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "tool-title"
  }, "\uD0C0\uC774\uBA38 / \uC2A4\uD1B1\uC6CC\uCE58"), /*#__PURE__*/React.createElement("div", {
    className: "tool-desc"
  }, "\uD65C\uB3D9 \uC2DC\uAC04\uC744 \uAD00\uB9AC\uD574\uC694"))), /*#__PURE__*/React.createElement("div", {
    className: "timer-container"
  }, /*#__PURE__*/React.createElement("div", {
    className: "timer-mode-toggle"
  }, /*#__PURE__*/React.createElement("button", {
    className: `timer-mode-btn ${timerMode === 'timer' ? 'active' : ''}`,
    onClick: () => {
      resetTimer();
      setTimerMode('timer');
    }
  }, "\u23F3 \uD0C0\uC774\uBA38"), /*#__PURE__*/React.createElement("button", {
    className: `timer-mode-btn ${timerMode === 'stopwatch' ? 'active' : ''}`,
    onClick: () => {
      resetTimer();
      setTimerMode('stopwatch');
    }
  }, "\u23F1\uFE0F \uC2A4\uD1B1\uC6CC\uCE58")), /*#__PURE__*/React.createElement("div", {
    className: `timer-display ${timerRunning ? 'running' : ''} ${isAlarm ? 'alarm' : ''}`
  }, timerMode === 'timer' ? formatTimer(timerSeconds) : formatStopwatch(timerElapsed)), timerMode === 'timer' && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    className: "timer-presets"
  }, presets.map(p => /*#__PURE__*/React.createElement("button", {
    key: p.sec,
    className: `timer-preset ${timerSeconds === p.sec && !timerRunning ? 'active' : ''}`,
    onClick: () => {
      if (!timerRunning) setTimerSeconds(p.sec);
    }
  }, p.label))), !timerRunning && /*#__PURE__*/React.createElement("div", {
    className: "timer-custom-input"
  }, /*#__PURE__*/React.createElement("input", {
    type: "number",
    className: "custom-time-input",
    placeholder: "\uBD84",
    min: "0",
    max: "99",
    value: customMin,
    onChange: e => setCustomMin(e.target.value)
  }), /*#__PURE__*/React.createElement("span", {
    className: "time-separator"
  }, ":"), /*#__PURE__*/React.createElement("input", {
    type: "number",
    className: "custom-time-input",
    placeholder: "\uCD08",
    min: "0",
    max: "59",
    value: customSec,
    onChange: e => setCustomSec(e.target.value)
  }), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-sm btn-primary",
    onClick: applyCustomTime
  }, "\uC801\uC6A9"))), /*#__PURE__*/React.createElement("div", {
    className: "timer-controls"
  }, !timerRunning ? /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary",
    onClick: startTimer
  }, "\u25B6 \uC2DC\uC791") : /*#__PURE__*/React.createElement("button", {
    className: "btn btn-secondary",
    onClick: pauseTimer
  }, "\u23F8 \uC77C\uC2DC\uC815\uC9C0"), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-danger",
    onClick: resetTimer,
    style: {
      background: 'var(--danger)',
      color: 'white'
    }
  }, "\u21BA \uCD08\uAE30\uD654")))), /*#__PURE__*/React.createElement("div", {
    className: `tool-card ${!isAdmin ? 'disabled-tool' : ''}`
  }, /*#__PURE__*/React.createElement("div", {
    className: "tool-card-header"
  }, /*#__PURE__*/React.createElement("span", {
    className: "tool-emoji"
  }, "\uD83D\uDC65"), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "tool-title"
  }, "\uD300 \uD3B8\uC131"), /*#__PURE__*/React.createElement("div", {
    className: "tool-desc"
  }, "\uB79C\uB364\uC73C\uB85C \uD300\uC744 \uB098\uB220\uC694"))), /*#__PURE__*/React.createElement("div", {
    className: "group-maker"
  }, isAdmin && /*#__PURE__*/React.createElement("div", {
    className: "group-settings"
  }, /*#__PURE__*/React.createElement("div", {
    className: "group-setting"
  }, /*#__PURE__*/React.createElement("label", null, "\uD300 \uC218:"), /*#__PURE__*/React.createElement("input", {
    type: "number",
    min: "2",
    max: "14",
    value: groupCount,
    onChange: e => setGroupCount(Math.max(2, parseInt(e.target.value) || 2))
  }), /*#__PURE__*/React.createElement("span", null, "\uAC1C")), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary",
    onClick: makeGroups
  }, "\uD83D\uDD00 \uD300 \uD3B8\uC131\uD558\uAE30")), groups.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "groups-result"
  }, groups.map((g, i) => /*#__PURE__*/React.createElement("div", {
    key: i,
    className: "group-card",
    style: {
      animationDelay: `${i * 0.1}s`
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "group-card-title"
  }, i + 1, "\uD300"), /*#__PURE__*/React.createElement("div", {
    className: "group-members"
  }, g.map(s => /*#__PURE__*/React.createElement("div", {
    key: s.id,
    className: "group-member"
  }, /*#__PURE__*/React.createElement("span", null, s.gender === 'M' ? '👦' : '👧'), /*#__PURE__*/React.createElement("span", null, s.number, ". ", s.name))))))))), /*#__PURE__*/React.createElement("div", {
    className: "tool-card"
  }, /*#__PURE__*/React.createElement("div", {
    className: "tool-card-header"
  }, /*#__PURE__*/React.createElement("span", {
    className: "tool-emoji"
  }, "\uD83D\uDCCA"), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "tool-title"
  }, "\uD22C\uD45C / \uC124\uBB38"), /*#__PURE__*/React.createElement("div", {
    className: "tool-desc"
  }, "\uD559\uAE09 \uC758\uACAC\uC744 \uBAA8\uC544\uC694"))), /*#__PURE__*/React.createElement("div", {
    className: "poll-container"
  }, isAdmin && /*#__PURE__*/React.createElement("div", {
    className: "poll-form"
  }, /*#__PURE__*/React.createElement("input", {
    className: "form-input",
    placeholder: "\uD22C\uD45C \uC9C8\uBB38\uC744 \uC785\uB825\uD558\uC138\uC694",
    value: newPollQ,
    onChange: e => setNewPollQ(e.target.value),
    style: {
      marginBottom: '0.5rem'
    }
  }), /*#__PURE__*/React.createElement("div", {
    className: "poll-options-list"
  }, newPollOpts.map((opt, i) => /*#__PURE__*/React.createElement("div", {
    key: i,
    className: "poll-option-input"
  }, /*#__PURE__*/React.createElement("input", {
    className: "form-input",
    placeholder: `항목 ${i + 1}`,
    value: opt,
    onChange: e => {
      const o = [...newPollOpts];
      o[i] = e.target.value;
      setNewPollOpts(o);
    }
  }), newPollOpts.length > 2 && /*#__PURE__*/React.createElement("button", {
    className: "poll-remove",
    onClick: () => setNewPollOpts(newPollOpts.filter((_, j) => j !== i))
  }, "\xD7")))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: 'flex',
      gap: '0.5rem'
    }
  }, /*#__PURE__*/React.createElement("button", {
    className: "btn btn-secondary btn-sm",
    onClick: () => setNewPollOpts([...newPollOpts, ''])
  }, "+ \uD56D\uBAA9 \uCD94\uAC00"), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary btn-sm",
    onClick: handleAddPoll
  }, "\uD22C\uD45C \uC0DD\uC131"))), polls.map(p => {
    const totalVotes = p.votes.reduce((a, b) => a + b, 0);
    const voted = hasVoted(p.id);
    return /*#__PURE__*/React.createElement("div", {
      key: p.id,
      className: "poll-active"
    }, /*#__PURE__*/React.createElement("div", {
      className: "poll-header-row"
    }, /*#__PURE__*/React.createElement("h4", {
      className: "poll-question"
    }, p.question), isAdmin && /*#__PURE__*/React.createElement("button", {
      className: "poll-delete-btn",
      onClick: () => handleDeletePoll(p.id)
    }, "\uD83D\uDDD1\uFE0F")), voted && /*#__PURE__*/React.createElement("div", {
      className: "poll-voted-badge"
    }, "\u2705 \uD22C\uD45C \uC644\uB8CC"), /*#__PURE__*/React.createElement("div", {
      className: "poll-results"
    }, p.options.map((opt, i) => {
      const pct = totalVotes > 0 ? Math.round(p.votes[i] / totalVotes * 100) : 0;
      return /*#__PURE__*/React.createElement("div", {
        key: i,
        className: `poll-result-item ${voted ? 'voted' : ''}`,
        onClick: () => !voted && handleVote(p.id, i),
        style: {
          cursor: voted ? 'default' : 'pointer'
        }
      }, /*#__PURE__*/React.createElement("span", {
        className: "poll-result-label"
      }, opt), /*#__PURE__*/React.createElement("div", {
        className: "poll-result-bar"
      }, /*#__PURE__*/React.createElement("div", {
        className: "poll-result-fill",
        style: {
          width: `${pct}%`
        }
      }, pct > 10 ? `${pct}%` : '')), /*#__PURE__*/React.createElement("span", {
        className: "poll-result-count"
      }, p.votes[i], "\uD45C"));
    })), /*#__PURE__*/React.createElement("div", {
      className: "poll-total"
    }, "\uCD1D ", totalVotes, "\uD45C"));
  })))));
}

/* ===================================================================
   WeeklyPlanPage
   =================================================================== */
const WEEK_DAYS = ['월', '화', '수', '목', '금'];
const WEEK_PERIODS = [1, 2, 3, 4, 5, 6];
function WeeklyPlanPage() {
  const {
    state,
    showToast,
    saveWeeklyPlan
  } = useApp();
  const {
    weeklyPlans,
    isAdmin
  } = state;
  const [currentDate, setCurrentDate] = useState(new Date());
  const [editing, setEditing] = useState(null);
  const [editSubject, setEditSubject] = useState('');
  const [editContent, setEditContent] = useState('');
  const weekStart = useMemo(() => getWeekStart(currentDate), [currentDate]);
  const weekKey = toISODate(weekStart);
  const dayDates = useMemo(() => WEEK_DAYS.map((d, i) => {
    const date = new Date(weekStart);
    date.setDate(date.getDate() + i);
    return {
      day: d,
      dateStr: `${date.getMonth() + 1}/${date.getDate()}`,
      isoDate: toISODate(date)
    };
  }), [weekStart]);
  const weekEndDate = new Date(weekStart);
  weekEndDate.setDate(weekEndDate.getDate() + 4);
  const currentWeekPlans = weeklyPlans[weekKey] || {};
  const todayStr = toISODate(new Date());
  const isCurrentWeek = dayDates.some(dd => dd.isoDate === todayStr);
  const goToPrevWeek = () => {
    const d = new Date(currentDate);
    d.setDate(d.getDate() - 7);
    setCurrentDate(d);
    setEditing(null);
  };
  const goToNextWeek = () => {
    const d = new Date(currentDate);
    d.setDate(d.getDate() + 7);
    setCurrentDate(d);
    setEditing(null);
  };
  const goToToday = () => {
    setCurrentDate(new Date());
    setEditing(null);
  };
  const handleEdit = (day, period, subject, content) => {
    if (!isAdmin) return;
    setEditing({
      day,
      period
    });
    setEditSubject(subject || '');
    setEditContent(content || '');
  };
  const handleSave = useCallback(() => {
    saveWeeklyPlan(weekKey, editing.day, editing.period, editSubject, editContent);
    setEditing(null);
    showToast('주간 일정이 저장되었습니다.', 'success');
  }, [saveWeeklyPlan, weekKey, editing, editSubject, editContent, showToast]);
  const sameMonth = weekStart.getMonth() === weekEndDate.getMonth();
  const rangeLabel = `${weekStart.getMonth() + 1}월 ${weekStart.getDate()}일 – ${sameMonth ? '' : `${weekEndDate.getMonth() + 1}월 `}${weekEndDate.getDate()}일`;
  const renderRow = p => /*#__PURE__*/React.createElement("tr", {
    key: p
  }, /*#__PURE__*/React.createElement("th", {
    scope: "row",
    className: "wk-period"
  }, /*#__PURE__*/React.createElement("span", {
    className: "wk-period-no"
  }, p, "\uAD50\uC2DC"), /*#__PURE__*/React.createElement("span", {
    className: "wk-period-time num"
  }, minToClock(PERIOD_SLOTS[p][0]))), dayDates.map(dd => {
    const plan = currentWeekPlans[dd.day]?.[p] || {};
    const isEditing = editing?.day === dd.day && editing?.period === p;
    const isToday = dd.isoDate === todayStr;
    const isEmpty = !plan.subject && !plan.content;
    return /*#__PURE__*/React.createElement("td", {
      key: dd.day,
      className: `wk-cell ${isToday ? 'is-today' : ''} ${isEditing ? 'is-editing' : ''}`
    }, isEditing ? /*#__PURE__*/React.createElement("div", {
      className: "wk-form"
    }, /*#__PURE__*/React.createElement("input", {
      className: "form-input",
      "aria-label": `${dd.day}요일 ${p}교시 과목`,
      placeholder: "\uACFC\uBAA9 (\uC608: \uAD6D\uC5B4)",
      value: editSubject,
      onChange: e => setEditSubject(e.target.value),
      autoFocus: true
    }), /*#__PURE__*/React.createElement("input", {
      className: "form-input",
      "aria-label": `${dd.day}요일 ${p}교시 학습 내용`,
      placeholder: "\uD559\uC2B5 \uB0B4\uC6A9",
      value: editContent,
      onChange: e => setEditContent(e.target.value),
      onKeyDown: e => e.key === 'Enter' && handleSave()
    }), /*#__PURE__*/React.createElement("div", {
      className: "wk-form-actions"
    }, /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: "btn btn-sm btn-secondary",
      onClick: () => setEditing(null)
    }, "\uCDE8\uC18C"), /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: "btn btn-sm btn-primary",
      onClick: handleSave
    }, "\uC800\uC7A5"))) : isAdmin ? /*#__PURE__*/React.createElement("button", {
      type: "button",
      "data-tone": subjectTone(plan.subject),
      className: `wk-plan ${isEmpty ? 'is-empty' : ''}`,
      "aria-label": `${dd.day}요일 ${p}교시 ${isEmpty ? '입력' : '수정'}`,
      onClick: () => handleEdit(dd.day, p, plan.subject, plan.content)
    }, isEmpty ? /*#__PURE__*/React.createElement("span", {
      className: "wk-add"
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "plus",
      size: 14
    }), "\uC785\uB825") : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("span", {
      className: "wk-subject"
    }, plan.subject), plan.content && /*#__PURE__*/React.createElement("span", {
      className: "wk-content"
    }, plan.content))) : /*#__PURE__*/React.createElement("div", {
      "data-tone": subjectTone(plan.subject),
      className: `wk-plan ${isEmpty ? 'is-empty' : ''}`
    }, plan.subject && /*#__PURE__*/React.createElement("span", {
      className: "wk-subject"
    }, plan.subject), plan.content && /*#__PURE__*/React.createElement("span", {
      className: "wk-content"
    }, plan.content)));
  }));
  return /*#__PURE__*/React.createElement("div", {
    className: "wk"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wk-top"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wk-title"
  }, /*#__PURE__*/React.createElement("h1", {
    className: "dash-headline"
  }, rangeLabel), isCurrentWeek && /*#__PURE__*/React.createElement("span", {
    className: "wk-chip"
  }, "\uC774\uBC88 \uC8FC")), /*#__PURE__*/React.createElement("div", {
    className: "wk-nav"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "round-btn",
    "aria-label": "\uC774\uC804 \uC8FC",
    onClick: goToPrevWeek
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "chevron",
    size: 18
  })), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "btn btn-sm btn-secondary",
    onClick: goToToday,
    disabled: isCurrentWeek
  }, "\uC624\uB298"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "round-btn wk-next",
    "aria-label": "\uB2E4\uC74C \uC8FC",
    onClick: goToNextWeek
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "chevron",
    size: 18
  })))), isAdmin && /*#__PURE__*/React.createElement("p", {
    className: "dash-edit-hint"
  }, "\uCE78\uC744 \uB20C\uB7EC \uACFC\uBAA9\uACFC \uD559\uC2B5 \uB0B4\uC6A9\uC744 \uC801\uACE0, Enter\uB85C \uC800\uC7A5\uD558\uC138\uC694."), /*#__PURE__*/React.createElement("div", {
    className: "wk-card"
  }, /*#__PURE__*/React.createElement("table", {
    className: "wk-table"
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", {
    scope: "col",
    className: "wk-corner"
  }, "\uAD50\uC2DC"), dayDates.map(dd => /*#__PURE__*/React.createElement("th", {
    scope: "col",
    key: dd.day,
    className: dd.isoDate === todayStr ? 'is-today' : ''
  }, /*#__PURE__*/React.createElement("span", {
    className: "wk-day"
  }, dd.day), /*#__PURE__*/React.createElement("span", {
    className: "wk-date num"
  }, dd.dateStr))))), /*#__PURE__*/React.createElement("tbody", null, WEEK_PERIODS.map(renderRow)))));
}

/* ===================================================================
   LinksPage
   =================================================================== */
const LINK_EMOJIS = [['🔗', '링크'], ['🌐', '웹사이트'], ['🏫', '학교'], ['📚', '책·자료'], ['📝', '과제·문서'], ['📊', '통계·표'], ['📅', '일정'], ['💬', '소통·메신저'], ['🎬', '영상'], ['🎵', '음악'], ['🎨', '미술'], ['🧮', '수학'], ['🔬', '과학'], ['🌏', '지도·사회'], ['🎮', '게임']];
function LinksPage() {
  const {
    state,
    showToast,
    addLink,
    updateLink,
    deleteLink
  } = useApp();
  const {
    quickLinks,
    isAdmin
  } = state;
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [formData, setFormData] = useState({
    title: '',
    url: '',
    icon: '🔗',
    desc: ''
  });
  const handleEdit = link => {
    setEditingId(link.id);
    setFormData(link);
    setShowForm(true);
  };
  const handleDelete = id => {
    if (confirm('정말 삭제하시겠습니까?')) {
      deleteLink(id);
      showToast('삭제되었습니다.', 'success');
    }
  };
  const handleSave = () => {
    if (!formData.title || !formData.url) {
      showToast('이름과 주소를 모두 입력해주세요.', 'error');
      return;
    }
    const url = formData.url.startsWith('http') ? formData.url : `https://${formData.url}`;
    if (editingId) {
      updateLink({
        ...formData,
        url,
        id: editingId
      });
      showToast('수정되었습니다.', 'success');
    } else {
      addLink({
        ...formData,
        url,
        id: generateUUID()
      });
      showToast('추가되었습니다.', 'success');
    }
    setShowForm(false);
    setEditingId(null);
    setFormData({
      title: '',
      url: '',
      icon: '🔗',
      desc: ''
    });
  };
  const handleCancel = () => {
    setShowForm(false);
    setEditingId(null);
    setFormData({
      title: '',
      url: '',
      icon: '🔗',
      desc: ''
    });
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "links-page"
  }, /*#__PURE__*/React.createElement("div", {
    className: "links-header"
  }, /*#__PURE__*/React.createElement("h2", {
    className: "section-title"
  }, "\uD83D\uDD17 \uC990\uACA8\uCC3E\uAE30 & \uBC14\uB85C\uAC00\uAE30"), isAdmin && /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary",
    onClick: () => setShowForm(true),
    disabled: showForm
  }, "+ \uB9C1\uD06C \uCD94\uAC00")), isAdmin && showForm && /*#__PURE__*/React.createElement("div", {
    className: "card links-form"
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      marginBottom: '1rem',
      fontSize: '1.125rem'
    }
  }, editingId ? '링크 수정' : '새 링크 추가'), /*#__PURE__*/React.createElement("div", {
    className: "form-group"
  }, /*#__PURE__*/React.createElement("label", {
    id: "link-icon-label"
  }, "\uC544\uC774\uCF58 \uACE0\uB974\uAE30"), /*#__PURE__*/React.createElement("div", {
    className: "emoji-picker",
    role: "group",
    "aria-labelledby": "link-icon-label"
  }, LINK_EMOJIS.map(([e, name]) => /*#__PURE__*/React.createElement("button", {
    key: e,
    type: "button",
    className: "emoji-opt",
    "aria-pressed": formData.icon === e,
    "aria-label": name,
    title: name,
    onClick: () => setFormData({
      ...formData,
      icon: e
    })
  }, e)))), /*#__PURE__*/React.createElement("div", {
    className: "form-row"
  }, /*#__PURE__*/React.createElement("div", {
    className: "form-group",
    style: {
      flex: 1
    }
  }, /*#__PURE__*/React.createElement("label", null, "\uC9C1\uC811 \uC785\uB825"), /*#__PURE__*/React.createElement("input", {
    className: "form-input",
    value: formData.icon,
    onChange: e => setFormData({
      ...formData,
      icon: e.target.value
    }),
    maxLength: "2"
  })), /*#__PURE__*/React.createElement("div", {
    className: "form-group",
    style: {
      flex: 4
    }
  }, /*#__PURE__*/React.createElement("label", null, "\uC774\uB984"), /*#__PURE__*/React.createElement("input", {
    className: "form-input",
    placeholder: "\uC608: \uAD6C\uAE00 \uBC88\uC5ED",
    value: formData.title,
    onChange: e => setFormData({
      ...formData,
      title: e.target.value
    })
  }))), /*#__PURE__*/React.createElement("div", {
    className: "form-group"
  }, /*#__PURE__*/React.createElement("label", null, "\uC8FC\uC18C (URL)"), /*#__PURE__*/React.createElement("input", {
    className: "form-input",
    placeholder: "https://...",
    value: formData.url,
    onChange: e => setFormData({
      ...formData,
      url: e.target.value
    })
  })), /*#__PURE__*/React.createElement("div", {
    className: "form-group"
  }, /*#__PURE__*/React.createElement("label", null, "\uC124\uBA85 (\uC120\uD0DD)"), /*#__PURE__*/React.createElement("input", {
    className: "form-input",
    placeholder: "\uC5B4\uB5A4 \uC0AC\uC774\uD2B8\uC778\uC9C0 \uC124\uBA85...",
    value: formData.desc,
    onChange: e => setFormData({
      ...formData,
      desc: e.target.value
    })
  })), /*#__PURE__*/React.createElement("div", {
    className: "form-actions"
  }, /*#__PURE__*/React.createElement("button", {
    className: "btn btn-secondary",
    onClick: handleCancel
  }, "\uCDE8\uC18C"), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary",
    onClick: handleSave
  }, "\uC800\uC7A5"))), /*#__PURE__*/React.createElement("div", {
    className: "links-grid"
  }, quickLinks.map(link => /*#__PURE__*/React.createElement("div", {
    key: link.id,
    className: "link-card-wrapper"
  }, /*#__PURE__*/React.createElement("a", {
    href: link.url,
    target: "_blank",
    rel: "noopener noreferrer",
    className: "link-card card"
  }, /*#__PURE__*/React.createElement("div", {
    className: "link-icon"
  }, link.icon), /*#__PURE__*/React.createElement("div", {
    className: "link-content"
  }, /*#__PURE__*/React.createElement("h3", {
    className: "link-title"
  }, link.title), link.desc && /*#__PURE__*/React.createElement("p", {
    className: "link-desc"
  }, link.desc))), isAdmin && /*#__PURE__*/React.createElement("div", {
    className: "link-actions"
  }, /*#__PURE__*/React.createElement("button", {
    className: "btn-icon link-action-btn",
    onClick: () => handleEdit(link)
  }, "\u270F\uFE0F"), /*#__PURE__*/React.createElement("button", {
    className: "btn-icon link-action-btn",
    onClick: () => handleDelete(link.id)
  }, "\uD83D\uDDD1\uFE0F")))), quickLinks.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "empty-state",
    style: {
      gridColumn: '1 / -1',
      textAlign: 'center',
      padding: '3rem',
      color: 'var(--text-secondary)'
    }
  }, "\uB4F1\uB85D\uB41C \uC990\uACA8\uCC3E\uAE30\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.")));
}

/* ===================================================================
   Root
   =================================================================== */
const ROUTE_PAGES = {
  '/': DashboardPage,
  '/students': StudentsPage,
  '/weekly': WeeklyPlanPage,
  '/attendance': AttendancePage,
  '/learning': LearningPage,
  '/calendar': CalendarPage,
  '/tools': ToolsPage,
  '/links': LinksPage
};
function AppShell() {
  const pathname = useHashRoute();
  const Page = ROUTE_PAGES[pathname] || DashboardPage;
  return /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(Layout, {
    pathname: pathname
  }, /*#__PURE__*/React.createElement(Page, null)), /*#__PURE__*/React.createElement(Toast, null));
}
function Root() {
  return /*#__PURE__*/React.createElement(AppProvider, null, /*#__PURE__*/React.createElement(AppShell, null));
}
ReactDOM.createRoot(document.getElementById('root')).render(/*#__PURE__*/React.createElement(Root, null));
