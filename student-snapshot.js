/* 학생 화면(QR)에 보여줄 내용만 골라 담는 함수 — 브라우저(app.js)와 교실 서버(server.js)가 함께 씀
   전화번호·알레르기·생일·출결·학생 기록·비밀번호·API 키는 절대 넣지 않음 (아래에 적은 항목만 골라 담음) */
(function (root) {
  var CLASS_LABEL = { school: '하남중앙초등학교', klass: '4학년 2반', teacher: '김대붕 선생님' };

  function pad(n) { return String(n).padStart(2, '0'); }
  function isoDate(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function weekStart(d) { var x = new Date(d); var day = x.getDay(); x.setDate(x.getDate() - (day === 0 ? 6 : day - 1)); x.setHours(0, 0, 0, 0); return x; }

  // 학습 관리의 가장 최근 '활동' (선생님 화면의 활동 완료 체크와 같은 기준)
  function currentActivity(db) {
    return (db.assignments || []).filter(function (a) { return a.subject === '활동'; })
      .sort(function (a, b) { return new Date(b.createdAt || 0) - new Date(a.createdAt || 0); })[0] || null;
  }

  // opts.withActivity: 교실 서버(같은 학교망)에서만 — 활동 체크용으로 반 명단(번호·이름)과 완료한 학생 id를 함께 담음
  function build(db, now, opts) {
    db = db || {};
    var students = db.students || [], weeklyPlans = db.weeklyPlans || {}, ddays = db.ddays || [];
    var announcements = db.announcements || [], quickLinks = db.quickLinks || [], settings = db.settings || {};
    var ws = weekStart(now || new Date());
    var plans = {};
    [-7, 0, 7, 14].forEach(function (off) {
      var d = new Date(ws); d.setDate(d.getDate() + off);
      var k = isoDate(d), wk = weeklyPlans[k];
      if (!wk) return;
      plans[k] = {};
      Object.keys(wk).forEach(function (day) {
        var periods = wk[day];
        if (!periods || typeof periods !== 'object') return;
        plans[k][day] = {};
        Object.keys(periods).forEach(function (p) {
          var v = periods[p];
          if (v && (v.subject || v.content)) plans[k][day][p] = { subject: v.subject || '', content: v.content || '' };
        });
      });
    });
    var ranked = students.slice().sort(function (a, b) { return (b.points || 0) - (a.points || 0); }).slice(0, 5);
    var extra = {};
    if (opts && opts.withActivity) {
      var pack = function (a) { return { id: String(a.id), title: a.title, type: a.type || '과제', date: (a.createdAt || a.dueDate || '').slice(0, 10), done: Object.keys(a.submissions || {}).filter(function (k) { return a.submissions[k]; }) }; };
      // 지난 활동까지 모두 (최근 것부터, 최대 60개). 맨 앞이 '오늘의 활동'
      var all = (db.assignments || []).filter(function (a) { return a.subject === '활동'; })
        .sort(function (a, b) { return new Date(b.createdAt || 0) - new Date(a.createdAt || 0); }).slice(0, 60).map(pack);
      extra.roster = students.map(function (s) { return { id: String(s.id), number: s.number, name: s.name }; });
      extra.activities = all;
      extra.activity = all[0] || null;
    }
    return Object.assign(extra, {
      v: 1,
      label: CLASS_LABEL,
      weeklyPlans: plans,
      ddays: ddays.map(function (e) { return { name: e.name, date: e.date, emoji: e.emoji || '' }; }),
      announcements: announcements.map(function (n) { return { id: n.id, title: n.title, content: n.content || '', type: n.type || 'general', date: n.date }; }),
      thermometer: { total: students.reduce(function (s, x) { return s + (x.points || 0); }, 0), goal: settings.thermometerGoal || 1500, reward: settings.thermometerReward || '' },
      ranking: ranked.map(function (s) { return { name: s.name, points: s.points || 0 }; }),
      // 투표/설문: 진행 중인 것만 (질문·항목·표 수. 누가 냈는지는 원래 기록되지 않음)
      polls: (db.polls || []).filter(function (p) { return p && p.isActive !== false && Array.isArray(p.options); }).map(function (p) {
        return { id: String(p.id), question: p.question, options: p.options.slice(), votes: p.options.map(function (_, i) { return (p.votes && p.votes[i]) || 0; }) };
      }),
      links: quickLinks.filter(function (l) { return /^https?:\/\//i.test(l.url || ''); }).map(function (l) { return { id: l.id, title: l.title, url: l.url, icon: l.icon || '', desc: l.desc || '' }; }),
    });
  }

  root.StudentSnapshot = { build: build, currentActivity: currentActivity, CLASS_LABEL: CLASS_LABEL };
})(typeof window !== 'undefined' ? window : globalThis);
