/**
 * ============================================================================
 * 속담 온(ON) - GitHub Pages 전용 코어 애플리케이션 스크립트 (app.js)
 * ============================================================================
 * 
 * [주요 기능]
 * 1. 학생용 5문항 객관식 속담 퀴즈 (100점 만점 채점 및 상세 오답 노트)
 * 2. 학생 식별번호(학번) 및 이름 입력 폼과 진행 타이머
 * 3. 브라우저 로컬 데이터베이스(LocalStorage) 기반 성적 자동 영구 저장
 * 4. 교사용 대시보드 (통계 KPI, 문항별 정답률, 학생별 검색, 오답 상세 모달)
 * 5. 엑셀 및 구글 시트에서 바로 열 수 있는 UTF-8 BOM CSV 성적표 다운로드 기능
 * 6. 교사 인증 암호(기본: teacher2026!) 및 가상 학생 10명 데이터 생성 기능
 * 7. Web Audio API 무의존성 사운드 & 캔버스 폭죽(Confetti) 효과
 */

// 1. 고등학교 1학년 필수 속담 5문항 마스터 데이터
const QUIZ_MASTER_DATA = [
  {
    id: 1,
    category: "상황 맥락 추론",
    context: "민수는 내일이 기말고사인데 한 달 동안 공부를 미루다 밤을 새워 벼락치기를 하려고 한다. 그러자 친구 지훈이가 다음과 같이 조언했다.\n\n\"지금 그렇게 잠도 안 자고 암기만 해봤자 시험 시간에 머리만 멍해지고 뒤통수 맞기 십상이야. 근본적인 공부 습관을 바꾸지 않고 당장 급한 불만 끄려는 건 오히려 해가 될 텐데...\"",
    prompt: "위 대화에서 지훈이의 생각과 가장 잘 어울리는 속담은?",
    options: [
      "밑 빠진 독에 물 붓기",
      "언 발에 오줌 누기",
      "우물을 파도 한 우물을 파라",
      "뱁새가 황새 따라가다 가랑이 찢어진다"
    ],
    answerIndex: 1, // 2번: 언 발에 오줌 누기
    explanation: "'언 발에 오줌 누기'는 당장에는 효력이 있을지 모르나 결국 더 나쁜 결과를 초래하는 임시방편을 뜻합니다.",
    tip: "💡 연계 한자성어: 동족방뇨(凍足放尿), 하석상대(下石上臺), 임시변통(臨時變通)"
  },
  {
    id: 2,
    category: "비유적 의미 파악",
    context: "소설 지문 속 한 인물은 최신형 고급 승용차를 몰고 다니며 사람들에게 화려한 외양을 과시하지만, 실제로는 월세조차 내지 못해 끼니를 거르는 형편이었다. 이를 바라보는 마을 사람들의 평가는 냉담했다.",
    prompt: "위 인물의 처지를 꼬집기에 가장 적절한 속담은?",
    options: [
      "보기 좋은 떡이 먹기도 좋다",
      "빛 좋은 개살구",
      "그림의 떡",
      "호박에 줄 긋는다고 수박 되랴"
    ],
    answerIndex: 1, // 2번: 빛 좋은 개살구
    explanation: "'빛 좋은 개살구'는 겉모양은 그럴듯하고 화려해 보이지만 실속이 전혀 없음을 비유합니다.",
    tip: "💡 연계 한자성어: 양두구육(羊頭狗肉), 유명무실(有名無實)"
  },
  {
    id: 3,
    category: "문맥적 빈칸 추론",
    context: "\"하루 10분씩 낭비하는 시간은 얼핏 별것 아닌 것처럼 느껴집니다. 하지만 이 사소한 시간 낭비가 1년, 3년 쌓이면 결국 인생의 큰 손실로 이어집니다. 옛 속담에 (       )라는 말이 있듯, 작은 습관이 오랜 시간 누적되어 미치는 파급력을 결코 간과해서는 안 됩니다.\"",
    prompt: "다음 강연의 빈칸에 들어갈 속담으로 가장 알맞은 것은?",
    options: [
      "티끌 모아 태산",
      "가랑잎이 솔잎더러 바스락거린다고 한다",
      "가랑비에 옷 젖는 줄 모른다",
      "바늘 도둑이 소도둑 된다"
    ],
    answerIndex: 2, // 3번: 가랑비에 옷 젖는 줄 모른다
    explanation: "'가랑비에 옷 젖는 줄 모른다'는 사소한 일이나 지출이라도 거듭되면 무시하지 못할 정도로 커져 손해를 입게 됨을 비유합니다.",
    tip: "💡 오답 주의: '티끌 모아 태산'은 작은 것을 모아 큰 것을 이룬다는 긍정적 측면이며, 본 문맥은 사소한 손실/습관의 누적을 경계하므로 '가랑비에 옷 젖는 줄 모른다'가 정답입니다."
  },
  {
    id: 4,
    category: "인물의 태도 비판",
    context: "조별 과제 발표를 준비할 때, 본인의 분량인 자료 조사를 하나도 해오지 않은 수현이가 오히려 5분 늦게 도착한 영호에게 \"너 때문에 우리 팀 발표가 망가지게 생겼다\"며 삿대질을 하고 화를 냈다.",
    prompt: "이 상황에서 수현이의 뻔뻔한 태도를 꼬집기에 가장 적절한 속담은?",
    options: [
      "닭 쫓던 개 지붕 쳐다본다",
      "도둑이 제 발 저린다",
      "가는 말이 고와야 오는 말이 곱다",
      "똥 묻은 개가 겨 묻은 개 나무란다"
    ],
    answerIndex: 3, // 4번: 똥 묻은 개가 겨 묻은 개 나무란다
    explanation: "'똥 묻은 개가 겨 묻은 개 나무란다'는 자기는 더 큰 흉과 결함이 있으면서 남의 사소한 잘못을 들추어 흉본다는 뜻입니다.",
    tip: "💡 연계 한자성어: 적반하장(賊反荷杖), 하로진태(河魯陳太)"
  },
  {
    id: 5,
    category: "인물의 성품 파악",
    context: "전교 1등을 놓치지 않는 정우는 친구들이 어려운 문제를 물어보러 올 때마다 늘 친절하고 겸손한 태도로 가르쳐 주며, 결코 자신의 뛰어난 실력을 뽐내거나 으스대지 않는다.",
    prompt: "정우의 성숙한 인품과 태도를 가장 잘 드러내는 속담은?",
    options: [
      "벼는 익을수록 고개를 숙인다",
      "빈 수레가 요란하다",
      "얌전한 고양이가 부뚜막에 먼저 올라간다",
      "백지장도 맞들면 낫다"
    ],
    answerIndex: 0, // 1번: 벼는 익을수록 고개를 숙인다
    explanation: "'벼는 익을수록 고개를 숙인다'는 교양과 학식, 인격이 깊고 뛰어난 사람일수록 교만하지 않고 더욱 겸손해진다는 뜻입니다.",
    tip: "💡 국어 내신 포인트: '빈 수레가 요란하다'와 반의적(대비적) 의미를 지닌 대표 속담으로 자주 출제됩니다."
  }
];

// 2. 브라우저 로컬 저장소 관리 (Database Engine)
const Storage = {
  DB_KEY: 'proverb_quiz_records_gh_v1',
  TEACHER_KEY: 'teacher2026!',

  getRecords() {
    try {
      const raw = localStorage.getItem(this.DB_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch (e) {
      console.warn("Storage read error", e);
      return [];
    }
  },

  saveRecord(record) {
    const list = this.getRecords();
    list.unshift(record); // 최신 제출 순
    try {
      localStorage.setItem(this.DB_KEY, JSON.stringify(list));
    } catch (e) {
      console.warn("Storage save error", e);
    }
  },

  clearRecords() {
    localStorage.removeItem(this.DB_KEY);
  },

  seedMockStudents() {
    const mockList = [
      { id: "10101", name: "김민준", answers: [1, 1, 2, 3, 0], secs: 65 },  // 100점
      { id: "10102", name: "이서연", answers: [1, 1, 2, 2, 0], secs: 82 },  // 80점
      { id: "10103", name: "박도윤", answers: [1, 1, 0, 3, 0], secs: 110 }, // 80점
      { id: "10104", name: "정하은", answers: [0, 1, 2, 3, 0], secs: 74 },  // 80점
      { id: "10105", name: "최예준", answers: [1, 1, 2, 3, 1], secs: 90 },  // 80점
      { id: "10106", name: "강지우", answers: [1, 0, 2, 1, 0], secs: 135 }, // 60점
      { id: "10107", name: "윤서준", answers: [1, 1, 2, 3, 0], secs: 58 },  // 100점
      { id: "10108", name: "장수아", answers: [2, 1, 1, 3, 0], secs: 145 }, // 60점
      { id: "10109", name: "임현우", answers: [1, 1, 2, 0, 0], secs: 98 },  // 80점
      { id: "10110", name: "한지민", answers: [0, 0, 2, 3, 2], secs: 120 }  // 40점
    ];

    const records = [];
    const now = new Date();

    mockList.forEach((s, idx) => {
      const submitTime = new Date(now.getTime() - (mockList.length - idx) * 8 * 60000);
      let correctCount = 0;
      const details = [];

      QUIZ_MASTER_DATA.forEach((q, qIdx) => {
        const choice = s.answers[qIdx];
        const isCorrect = choice === q.answerIndex;
        if (isCorrect) correctCount++;
        details.push({
          questionId: q.id,
          choice: choice + 1,
          correctAnswer: q.answerIndex + 1,
          isCorrect: isCorrect
        });
      });

      records.push({
        timestamp: formatDate(submitTime),
        studentId: s.id,
        studentName: s.name,
        score: correctCount * 20,
        correctCount: correctCount,
        durationText: formatDuration(s.secs),
        elapsedSeconds: s.secs,
        questionDetails: details
      });
    });

    localStorage.setItem(this.DB_KEY, JSON.stringify(records));
    return records.length;
  },

  exportToCSV() {
    const records = this.getRecords();
    if (records.length === 0) {
      alert("다운로드할 학생 제출 데이터가 없습니다.");
      return;
    }

    const headers = [
      "제출일시", "학번/식별번호", "성명", "총점", "정답수", "소요시간",
      "1번선택", "1번정오", "2번선택", "2번정오", "3번선택", "3번정오", "4번선택", "4번정오", "5번선택", "5번정오"
    ];

    const rows = [headers];

    records.forEach(r => {
      const row = [
        `"${r.timestamp}"`,
        `"${r.studentId}"`,
        `"${r.studentName}"`,
        r.score,
        r.correctCount,
        `"${r.durationText}"`
      ];

      (r.questionDetails || []).forEach(qd => {
        row.push(qd.choice);
        row.push(qd.isCorrect ? "O" : "X");
      });

      rows.push(row);
    });

    // 엑셀에서 한글 깨짐을 방지하기 위한 UTF-8 BOM(\uFEFF) 추가
    const csvContent = "\uFEFF" + rows.map(e => e.join(",")).join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `속담퀴즈_학급성적표_${UtilitiesDateStr()}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
};

// 3. 상태 관리 객체 (State)
const state = {
  studentId: '',
  studentName: '',
  currentIndex: 0,
  userAnswers: [null, null, null, null, null],
  elapsedSeconds: 0,
  timerInterval: null,
  startTime: null,
  isSubmitted: false,
  soundEnabled: true,
  reviewFilter: 'all',

  // 교사용 모드 상태
  isTeacherAuth: false
};

// 4. 효과음 관리자 (Web Audio API)
class SoundManager {
  constructor() { this.ctx = null; }

  init() {
    if (!this.ctx) {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (AudioCtx) this.ctx = new AudioCtx();
    }
    if (this.ctx && this.ctx.state === 'suspended') {
      this.ctx.resume();
    }
  }

  playTone(freq, type = 'sine', duration = 0.15, gainVal = 0.08) {
    if (!state.soundEnabled) return;
    this.init();
    if (!this.ctx) return;
    try {
      const osc = this.ctx.createOscillator();
      const gain = this.ctx.createGain();
      osc.type = type;
      osc.frequency.setValueAtTime(freq, this.ctx.currentTime);
      gain.gain.setValueAtTime(gainVal, this.ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.0001, this.ctx.currentTime + duration);
      osc.connect(gain);
      gain.connect(this.ctx.destination);
      osc.start();
      osc.stop(this.ctx.currentTime + duration);
    } catch (e) {}
  }

  playClick() { this.playTone(600, 'triangle', 0.05, 0.05); }
  playSelect() { this.playTone(520, 'sine', 0.08, 0.06); }
  playNext() {
    if (!state.soundEnabled) return;
    this.init();
    if (!this.ctx) return;
    try {
      const now = this.ctx.currentTime;
      const osc = this.ctx.createOscillator();
      const gain = this.ctx.createGain();
      osc.frequency.setValueAtTime(440, now);
      osc.frequency.exponentialRampToValueAtTime(880, now + 0.12);
      gain.gain.setValueAtTime(0.06, now);
      gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.12);
      osc.connect(gain);
      gain.connect(this.ctx.destination);
      osc.start(now);
      osc.stop(now + 0.12);
    } catch(e) {}
  }

  playFanfare() {
    if (!state.soundEnabled) return;
    this.init();
    if (!this.ctx) return;
    const notes = [523.25, 659.25, 783.99, 1046.50];
    notes.forEach((freq, idx) => {
      setTimeout(() => {
        this.playTone(freq, 'triangle', 0.35, 0.1);
      }, idx * 120);
    });
  }
}

const sound = new SoundManager();

// 5. 폭죽 캔버스 파티클 엔진 (Confetti)
class ConfettiEngine {
  constructor(canvasId) {
    this.canvas = document.getElementById(canvasId);
    this.ctx = this.canvas ? this.canvas.getContext('2d') : null;
    this.particles = [];
    this.animId = null;
    this.resize();
    window.addEventListener('resize', () => this.resize());
  }

  resize() {
    if (!this.canvas) return;
    this.canvas.width = window.innerWidth;
    this.canvas.height = window.innerHeight;
  }

  burst(count = 100) {
    if (!this.canvas || !this.ctx) return;
    this.particles = [];
    const colors = ['#FF385C', '#E00B41', '#008489', '#FFB400', '#222222', '#717171', '#484848'];

    for (let i = 0; i < count; i++) {
      this.particles.push({
        x: this.canvas.width * 0.5 + (Math.random() - 0.5) * 200,
        y: this.canvas.height * 0.4 + (Math.random() - 0.5) * 100,
        vx: (Math.random() - 0.5) * 16,
        vy: (Math.random() - 1.2) * 16,
        size: Math.random() * 8 + 4,
        color: colors[Math.floor(Math.random() * colors.length)],
        rotation: Math.random() * 360,
        rotSpeed: (Math.random() - 0.5) * 12,
        alpha: 1,
        gravity: 0.35,
        decay: Math.random() * 0.008 + 0.006
      });
    }

    if (!this.animId) this.render();
  }

  render() {
    if (!this.ctx) return;
    this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);

    for (let i = this.particles.length - 1; i >= 0; i--) {
      const p = this.particles[i];
      p.x += p.vx;
      p.y += p.vy;
      p.vy += p.gravity;
      p.rotation += p.rotSpeed;
      p.alpha -= p.decay;

      if (p.alpha <= 0) {
        this.particles.splice(i, 1);
        continue;
      }

      this.ctx.save();
      this.ctx.translate(p.x, p.y);
      this.ctx.rotate((p.rotation * Math.PI) / 180);
      this.ctx.globalAlpha = p.alpha;
      this.ctx.fillStyle = p.color;
      this.ctx.fillRect(-p.size / 2, -p.size / 2, p.size, p.size);
      this.ctx.restore();
    }

    if (this.particles.length > 0) {
      this.animId = requestAnimationFrame(() => this.render());
    } else {
      this.animId = null;
    }
  }
}

let confetti = null;

// 6. DOM 요소 및 초기화
document.addEventListener('DOMContentLoaded', () => {
  confetti = new ConfettiEngine('confetti-canvas');
  bindEvents();
});

function bindEvents() {
  // Brand Logo Click -> Intro Screen
  document.getElementById('header-brand-btn').addEventListener('click', () => {
    sound.playClick();
    resetToIntro();
  });

  // Sound Toggle
  document.getElementById('sound-toggle-btn').addEventListener('click', () => {
    state.soundEnabled = !state.soundEnabled;
    document.getElementById('sound-icon-on').classList.toggle('hidden', !state.soundEnabled);
    document.getElementById('sound-icon-off').classList.toggle('hidden', state.soundEnabled);
    showToast(state.soundEnabled ? "🔊 효과음이 켜졌습니다." : "🔇 효과음이 꺼졌습니다.");
    if (state.soundEnabled) sound.playClick();
  });

  // Start Quiz
  document.getElementById('start-quiz-btn').addEventListener('click', () => {
    sound.playClick();
    handleStartQuiz();
  });

  // Quiz Navigation
  document.getElementById('prev-question-btn').addEventListener('click', () => {
    if (state.currentIndex > 0) {
      sound.playClick();
      navigateToQuestion(state.currentIndex - 1);
    }
  });

  document.getElementById('next-question-btn').addEventListener('click', () => {
    if (state.currentIndex === QUIZ_MASTER_DATA.length - 1) {
      handleSubmitQuiz();
    } else {
      sound.playNext();
      navigateToQuestion(state.currentIndex + 1);
    }
  });

  // Result Actions
  document.getElementById('retry-quiz-btn').addEventListener('click', () => {
    sound.playClick();
    resetToIntro();
  });

  document.getElementById('share-score-btn').addEventListener('click', () => {
    sound.playClick();
    copyResultsToClipboard();
  });

  // Review Filters
  document.getElementById('filter-all-btn').addEventListener('click', () => {
    sound.playClick();
    setReviewFilter('all');
  });

  document.getElementById('filter-wrong-btn').addEventListener('click', () => {
    sound.playClick();
    setReviewFilter('wrong');
  });

  // Teacher Mode & Auth Modal
  document.getElementById('teacher-mode-btn').addEventListener('click', () => {
    sound.playClick();
    if (state.isTeacherAuth) {
      openTeacherDashboard();
    } else {
      openTeacherAuthModal();
    }
  });

  document.getElementById('close-auth-modal-btn').addEventListener('click', closeTeacherAuthModal);
  document.getElementById('cancel-auth-btn').addEventListener('click', closeTeacherAuthModal);
  document.getElementById('submit-auth-btn').addEventListener('click', handleTeacherAuthSubmit);

  document.getElementById('teacher-passcode-input').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') handleTeacherAuthSubmit();
  });

  // Teacher Dashboard Controls
  document.getElementById('download-csv-btn').addEventListener('click', () => {
    sound.playClick();
    Storage.exportToCSV();
  });

  document.getElementById('generate-mock-btn').addEventListener('click', () => {
    sound.playClick();
    handleGenerateMockData();
  });

  document.getElementById('clear-records-btn').addEventListener('click', () => {
    if (confirm("저장된 모든 학생 제출 기록을 초기화하시겠습니까?")) {
      sound.playClick();
      Storage.clearRecords();
      renderTeacherDashboard();
      showToast("기록이 초기화되었습니다.");
    }
  });

  document.getElementById('exit-teacher-btn').addEventListener('click', () => {
    sound.playClick();
    resetToIntro();
  });

  document.getElementById('teacher-search-input').addEventListener('input', (e) => {
    renderStudentTable(e.target.value);
  });

  // Student Detail Modal
  document.getElementById('close-student-detail-btn').addEventListener('click', closeStudentDetailModal);
  document.getElementById('close-student-detail-btn-bottom').addEventListener('click', closeStudentDetailModal);

  // Keyboard navigation during quiz
  window.addEventListener('keydown', (e) => {
    if (!document.getElementById('quiz-screen').classList.contains('active')) return;

    if (['1', '2', '3', '4'].includes(e.key)) {
      selectOption(parseInt(e.key, 10) - 1);
    } else if (e.key === 'ArrowRight' && !document.getElementById('next-question-btn').disabled) {
      document.getElementById('next-question-btn').click();
    } else if (e.key === 'ArrowLeft' && !document.getElementById('prev-question-btn').disabled) {
      document.getElementById('prev-question-btn').click();
    }
  });
}

// 7. 화면 전환
function switchScreen(targetScreenId) {
  ['intro-screen', 'quiz-screen', 'result-screen', 'teacher-screen'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove('active');
  });

  const target = document.getElementById(targetScreenId);
  if (target) {
    target.classList.add('active');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
}

// 8. 퀴즈 진행 로직
function handleStartQuiz() {
  const idInput = document.getElementById('student-id-input');
  const nameInput = document.getElementById('student-name-input');
  const errorEl = document.getElementById('student-input-error');

  const sId = (idInput.value || '').trim();
  const sName = (nameInput.value || '').trim();

  if (!sId || !sName) {
    errorEl.classList.remove('hidden');
    idInput.focus();
    return;
  }
  errorEl.classList.add('hidden');

  state.studentId = sId;
  state.studentName = sName;
  state.currentIndex = 0;
  state.userAnswers = [null, null, null, null, null];
  state.isSubmitted = false;
  state.elapsedSeconds = 0;

  document.getElementById('active-student-display').textContent = `${sId} ${sName} 학생 응시 중`;

  state.startTime = Date.now();
  if (state.timerInterval) clearInterval(state.timerInterval);
  state.timerInterval = setInterval(updateTimerDisplay, 1000);
  updateTimerDisplay();

  renderNavPills();
  renderQuestion(0);
  switchScreen('quiz-screen');
}

function updateTimerDisplay() {
  state.elapsedSeconds = Math.floor((Date.now() - state.startTime) / 1000);
  const mins = String(Math.floor(state.elapsedSeconds / 60)).padStart(2, '0');
  const secs = String(state.elapsedSeconds % 60).padStart(2, '0');
  document.getElementById('quiz-timer').textContent = `${mins}:${secs}`;
}

function renderNavPills() {
  const navContainer = document.getElementById('question-navigator');
  navContainer.innerHTML = '';

  QUIZ_MASTER_DATA.forEach((q, idx) => {
    const pill = document.createElement('button');
    pill.className = `nav-pill-btn ${idx === state.currentIndex ? 'active' : ''} ${state.userAnswers[idx] !== null ? 'answered' : ''}`;
    pill.textContent = `Q${idx + 1}`;
    pill.title = `${idx + 1}번 문항으로 이동`;
    pill.addEventListener('click', () => {
      sound.playClick();
      navigateToQuestion(idx);
    });
    navContainer.appendChild(pill);
  });
}

function navigateToQuestion(index) {
  state.currentIndex = index;
  renderQuestion(index);
  renderNavPills();
}

function renderQuestion(index) {
  const q = QUIZ_MASTER_DATA[index];
  const total = QUIZ_MASTER_DATA.length;

  document.getElementById('current-q-num').textContent = String(index + 1).padStart(2, '0');
  const pct = ((index + 1) / total) * 100;
  document.getElementById('progress-fill').style.width = `${pct}%`;

  document.getElementById('question-category').textContent = q.category;
  document.getElementById('question-context').innerHTML = q.context.replace(/\n\n/g, '<br><br>');
  document.getElementById('question-prompt').textContent = q.prompt;

  const optionsContainer = document.getElementById('options-container');
  optionsContainer.innerHTML = '';
  const currentSelection = state.userAnswers[index];

  q.options.forEach((optText, optIdx) => {
    const btn = document.createElement('button');
    btn.className = `option-btn ${currentSelection === optIdx ? 'selected' : ''}`;
    btn.setAttribute('role', 'radio');
    btn.setAttribute('aria-checked', currentSelection === optIdx ? 'true' : 'false');
    btn.innerHTML = `
      <span class="option-num-badge">${optIdx + 1}</span>
      <span class="option-text">${optText}</span>
    `;

    btn.addEventListener('click', () => selectOption(optIdx));
    optionsContainer.appendChild(btn);
  });

  document.getElementById('prev-question-btn').disabled = index === 0;

  const nextBtnText = document.getElementById('next-btn-text');
  if (index === total - 1) {
    nextBtnText.textContent = '최종 제출하기 💯';
  } else {
    nextBtnText.textContent = '다음 문제';
  }

  const hintEl = document.getElementById('option-hint-text');
  if (currentSelection === null) {
    hintEl.textContent = '보기를 선택해 주세요 (키보드 1~4)';
  } else {
    hintEl.textContent = `선택 완료: ${currentSelection + 1}번 보기`;
  }
}

function selectOption(optIndex) {
  sound.playSelect();
  state.userAnswers[state.currentIndex] = optIndex;

  const allBtns = document.querySelectorAll('#options-container .option-btn');
  allBtns.forEach((btn, idx) => {
    btn.classList.toggle('selected', idx === optIndex);
    btn.setAttribute('aria-checked', idx === optIndex ? 'true' : 'false');
  });

  document.getElementById('option-hint-text').textContent = `선택 완료: ${optIndex + 1}번 보기`;
  renderNavPills();
}

// 9. 학생 답안 채점 및 저장
function handleSubmitQuiz() {
  const unanswered = [];
  state.userAnswers.forEach((ans, idx) => {
    if (ans === null) unanswered.push(idx + 1);
  });

  if (unanswered.length > 0) {
    sound.playTone(300, 'square', 0.2);
    const ok = confirm(`아직 풀지 않은 문항(${unanswered.join(', ')}번)이 있습니다.\n그대로 최종 제출하시겠습니까?`);
    if (!ok) {
      navigateToQuestion(unanswered[0] - 1);
      return;
    }
  }

  if (state.timerInterval) clearInterval(state.timerInterval);
  state.isSubmitted = true;
  sound.playClick();

  // 채점 계산
  let correctCount = 0;
  const questionDetails = [];
  const gradedResults = QUIZ_MASTER_DATA.map((q, idx) => {
    const userChoice = state.userAnswers[idx];
    const isCorrect = userChoice === q.answerIndex;
    if (isCorrect) correctCount++;

    questionDetails.push({
      questionId: q.id,
      choice: userChoice !== null ? userChoice + 1 : "미선택",
      correctAnswer: q.answerIndex + 1,
      isCorrect: isCorrect
    });

    return {
      id: q.id,
      category: q.category,
      context: q.context,
      prompt: q.prompt,
      userChoice: userChoice,
      userChoiceText: userChoice !== null ? q.options[userChoice] : "미선택",
      correctChoice: q.answerIndex,
      correctChoiceText: q.options[q.answerIndex],
      isCorrect: isCorrect,
      explanation: q.explanation,
      tip: q.tip
    };
  });

  const totalScore = correctCount * 20;
  const durationText = formatDuration(state.elapsedSeconds);

  // 로컬 저장소에 영구 저장 (GitHub Pages 환경)
  Storage.saveRecord({
    timestamp: formatDate(new Date()),
    studentId: state.studentId,
    studentName: state.studentName,
    score: totalScore,
    correctCount: correctCount,
    durationText: durationText,
    elapsedSeconds: state.elapsedSeconds,
    questionDetails: questionDetails
  });

  // 결과 화면 세팅
  switchScreen('result-screen');
  document.getElementById('result-student-info').textContent = `${state.studentId} ${state.studentName} 학생의 성적표`;

  const tierTitleEl = document.getElementById('result-tier-title');
  const tierDescEl = document.getElementById('result-tier-desc');

  if (totalScore === 100) {
    tierTitleEl.textContent = "👑 만점! 완벽한 속담 마스터";
    tierDescEl.textContent = "고1 수준의 모든 속담 문맥을 완벽히 마스터했습니다. 탁월한 국어 어휘력입니다!";
  } else if (totalScore >= 80) {
    tierTitleEl.textContent = "✨ 우수한 실력! 1등급 유망주";
    tierDescEl.textContent = "대부분의 속담 맥락을 정확하게 파악했습니다. 아쉽게 놓친 문항의 오답 노트를 복습해 보세요.";
  } else if (totalScore >= 60) {
    tierTitleEl.textContent = "🌱 꾸준한 발전이 기대되는 실력";
    tierDescEl.textContent = "기본적인 속담은 잘 이해하고 있으나, 유의 속담과의 구별 학습을 보완해 보세요.";
  } else {
    tierTitleEl.textContent = "📖 오답 복습이 필요한 단계";
    tierDescEl.textContent = "아래 상세 오답 노트를 통해 속담의 유래와 한자성어 연계 개념을 꼼꼼히 정리해 보세요.";
  }

  document.getElementById('stat-correct-count').textContent = `${correctCount} / 5`;
  document.getElementById('stat-accuracy').textContent = `${Math.round((correctCount / 5) * 100)}%`;
  document.getElementById('stat-duration').textContent = durationText;

  document.getElementById('wrong-count-badge').textContent = 5 - correctCount;

  animateScoreGauge(totalScore);

  if (totalScore >= 80) {
    setTimeout(() => {
      sound.playFanfare();
      confetti.burst(120);
    }, 400);
  }

  renderReviewList('all', gradedResults);
}

function animateScoreGauge(targetScore) {
  const duration = 1200;
  const startTime = performance.now();
  const circumference = 326.7;
  const targetOffset = circumference - (circumference * targetScore) / 100;

  const fillEl = document.getElementById('circle-score-fill');
  fillEl.style.strokeDashoffset = targetOffset;

  if (targetScore >= 80) {
    fillEl.style.stroke = '#FF385C';
  } else if (targetScore >= 60) {
    fillEl.style.stroke = '#F59E0B';
  } else {
    fillEl.style.stroke = '#E00B41';
  }

  function step(now) {
    const progress = Math.min((now - startTime) / duration, 1);
    const easeOutQuad = 1 - (1 - progress) * (1 - progress);
    const currentVal = Math.round(targetScore * easeOutQuad);
    document.getElementById('score-counter').textContent = currentVal;

    if (progress < 1) requestAnimationFrame(step);
  }

  requestAnimationFrame(step);
}

// 10. 오답 노트 렌더링
let currentGradedResultsCache = [];

function setReviewFilter(mode) {
  state.reviewFilter = mode;
  document.getElementById('filter-all-btn').classList.toggle('active', mode === 'all');
  document.getElementById('filter-wrong-btn').classList.toggle('active', mode === 'wrong');
  renderReviewList(mode, currentGradedResultsCache);
}

function renderReviewList(mode, gradedResults) {
  if (gradedResults) currentGradedResultsCache = gradedResults;
  const container = document.getElementById('review-list');
  container.innerHTML = '';

  const items = currentGradedResultsCache.filter(r => mode === 'all' || (mode === 'wrong' && !r.isCorrect));

  if (items.length === 0) {
    container.innerHTML = `
      <div class="empty-filter-state">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
          <polyline points="22 4 12 14.01 9 11.01"></polyline>
        </svg>
        <h4>틀린 문제가 없습니다!</h4>
        <p>5문항 모두 완벽하게 정답을 맞혔습니다. 축하합니다!</p>
      </div>
    `;
    return;
  }

  items.forEach(item => {
    const card = document.createElement('div');
    card.className = `review-item-card ${item.isCorrect ? 'is-correct' : 'is-wrong'}`;

    card.innerHTML = `
      <div class="item-top-row">
        <span class="item-num-tag">QUESTION ${String(item.id).padStart(2, '0')} · ${item.category}</span>
        <span class="status-badge ${item.isCorrect ? 'correct' : 'wrong'}">
          ${item.isCorrect ? '✓ 정답 (20점)' : '✕ 오답 (0점)'}
        </span>
      </div>

      <div class="item-context-quote">
        ${item.context.replace(/\n\n/g, '<br>')}
      </div>

      <h4 class="item-prompt">${item.prompt}</h4>

      <div class="answers-comparison">
        <div class="answer-box user-choice ${!item.isCorrect ? 'wrong-choice' : ''}">
          <span class="ans-tag">내가 고른 답</span>
          <span class="ans-text">${item.userChoice !== null ? `${item.userChoice + 1}. ${item.userChoiceText}` : '미선택'}</span>
        </div>
        <div class="answer-box correct-choice">
          <span class="ans-tag">실제 정답</span>
          <span class="ans-text">${item.correctChoice + 1}. ${item.correctChoiceText}</span>
        </div>
      </div>

      <div class="item-explanation-box">
        <div class="exp-title">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="12" cy="12" r="10"></circle>
            <line x1="12" y1="16" x2="12" y2="12"></line>
            <line x1="12" y1="8" x2="12.01" y2="8"></line>
          </svg>
          속담 해설 & 핵심 의미
        </div>
        <p class="exp-desc">${item.explanation}</p>
        <div class="exp-tip">${item.tip}</div>
      </div>
    `;

    container.appendChild(card);
  });
}

// 11. 교사 인증 및 대시보드 로직
function openTeacherAuthModal() {
  document.getElementById('teacher-auth-modal').classList.remove('hidden');
  document.getElementById('teacher-auth-error').classList.add('hidden');
  document.getElementById('teacher-passcode-input').value = '';
  document.getElementById('teacher-passcode-input').focus();
}

function closeTeacherAuthModal() {
  document.getElementById('teacher-auth-modal').classList.add('hidden');
}

function handleTeacherAuthSubmit() {
  const passcode = document.getElementById('teacher-passcode-input').value.trim();
  const errorEl = document.getElementById('teacher-auth-error');

  if (passcode === Storage.TEACHER_KEY) {
    state.isTeacherAuth = true;
    closeTeacherAuthModal();
    openTeacherDashboard();
  } else {
    errorEl.textContent = '인증 암호가 올바르지 않습니다. (기본: teacher2026!)';
    errorEl.classList.remove('hidden');
  }
}

function openTeacherDashboard() {
  switchScreen('teacher-screen');
  renderTeacherDashboard();
}

function renderTeacherDashboard() {
  const records = Storage.getRecords();
  const totalCount = records.length;

  let totalScoreSum = 0;
  let maxScore = 0;
  let minScore = totalCount > 0 ? 100 : 0;
  const qCorrectCounts = [0, 0, 0, 0, 0];

  records.forEach(r => {
    totalScoreSum += r.score;
    if (r.score > maxScore) maxScore = r.score;
    if (r.score < minScore) minScore = r.score;

    (r.questionDetails || []).forEach((qd, idx) => {
      if (qd.isCorrect) qCorrectCounts[idx]++;
    });
  });

  const avgScore = totalCount > 0 ? (Math.round((totalScoreSum / totalCount) * 10) / 10).toFixed(1) : "0.0";

  document.getElementById('kpi-total-students').textContent = totalCount;
  document.getElementById('kpi-avg-score').textContent = avgScore;
  document.getElementById('kpi-max-score').textContent = maxScore;
  document.getElementById('kpi-min-score').textContent = minScore;

  // 문항별 정답률 바 렌더링
  const accuracyContainer = document.getElementById('question-accuracy-container');
  accuracyContainer.innerHTML = '';

  if (totalCount === 0) {
    accuracyContainer.innerHTML = '<p class="text-muted">아직 응시한 학생 데이터가 없습니다.</p>';
  } else {
    QUIZ_MASTER_DATA.forEach((q, idx) => {
      const cCount = qCorrectCounts[idx];
      const rate = Math.round((cCount / totalCount) * 100);

      let rateClass = 'rate-mid';
      let barClass = 'bar-mid';
      if (rate >= 80) { rateClass = 'rate-high'; barClass = 'bar-high'; }
      else if (rate < 60) { rateClass = 'rate-low'; barClass = 'bar-low'; }

      const row = document.createElement('div');
      row.className = 'accuracy-item-row';
      row.innerHTML = `
        <div class="accuracy-meta-row">
          <div class="q-title-wrap">
            <span class="q-num-chip">Q${q.id}</span>
            <span class="q-prompt-short">${q.prompt}</span>
          </div>
          <span class="q-rate-val ${rateClass}">${rate}% (${cCount}/${totalCount}명 정답)</span>
        </div>
        <div class="accuracy-bar-track">
          <div class="accuracy-bar-fill ${barClass}" style="width: ${rate}%;"></div>
        </div>
      `;
      accuracyContainer.appendChild(row);
    });
  }

  // 테이블 렌더링
  renderStudentTable();
}

function renderStudentTable(keyword = '') {
  const tbody = document.getElementById('student-table-body');
  const countLabel = document.getElementById('table-count-label');
  tbody.innerHTML = '';

  const records = Storage.getRecords();
  const filtered = records.filter(s => {
    if (!keyword) return true;
    const kw = keyword.toLowerCase();
    return s.studentId.toLowerCase().includes(kw) || s.studentName.toLowerCase().includes(kw);
  });

  countLabel.textContent = `총 ${filtered.length}건`;

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="11" class="text-center text-muted py-8">일치하는 학생 데이터가 없습니다.</td></tr>`;
    return;
  }

  filtered.forEach(s => {
    const tr = document.createElement('tr');

    let qBadgesHtml = '';
    (s.questionDetails || []).forEach(qd => {
      qBadgesHtml += `<td class="text-center"><span class="ox-badge ${qd.isCorrect ? 'badge-o' : 'badge-x'}">${qd.isCorrect ? 'O' : 'X'}</span></td>`;
    });

    tr.innerHTML = `
      <td>${s.timestamp}</td>
      <td><strong>${s.studentId}</strong></td>
      <td>${s.studentName}</td>
      <td><strong class="text-rausch">${s.score}점</strong></td>
      <td>${s.durationText}</td>
      ${qBadgesHtml}
      <td class="text-center">
        <button class="view-detail-btn" type="button">오답 보기</button>
      </td>
    `;

    tr.querySelector('.view-detail-btn').addEventListener('click', () => {
      sound.playClick();
      openStudentDetailModal(s);
    });

    tbody.appendChild(tr);
  });
}

function openStudentDetailModal(student) {
  document.getElementById('modal-student-name').textContent = `${student.studentId} ${student.studentName} 학생`;
  document.getElementById('modal-student-meta').textContent = `제출 일시: ${student.timestamp} | 소요 시간: ${student.durationText}`;
  document.getElementById('modal-student-score').textContent = student.score;

  const container = document.getElementById('modal-question-answers-list');
  container.innerHTML = '';

  (student.questionDetails || []).forEach(qd => {
    const qMaster = QUIZ_MASTER_DATA.find(q => q.id === qd.questionId) || {};
    const div = document.createElement('div');
    div.className = `modal-answer-item ${qd.isCorrect ? 'correct' : 'wrong'}`;

    div.innerHTML = `
      <div class="item-q-header">
        <span>문항 ${qd.questionId}번 (${qd.isCorrect ? '✓ 정답 20점' : '✕ 오답 0점'})</span>
      </div>
      <div class="item-q-prompt">${qMaster.prompt || ''}</div>
      <div class="item-choices-comparison">
        <div><strong>학생 선택:</strong> ${qd.choice}번 보기</div>
        <div><strong>실제 정답:</strong> ${qd.correctAnswer}번 보기</div>
      </div>
    `;

    container.appendChild(div);
  });

  document.getElementById('student-detail-modal').classList.remove('hidden');
}

function closeStudentDetailModal() {
  document.getElementById('student-detail-modal').classList.add('hidden');
}

function handleGenerateMockData() {
  const count = Storage.seedMockStudents();
  renderTeacherDashboard();
  showToast(`✓ 가상 학생 ${count}명의 데이터가 생성되었습니다!`);
}

// 12. 공통 유틸리티
function resetToIntro() {
  if (state.timerInterval) clearInterval(state.timerInterval);
  state.currentIndex = 0;
  state.userAnswers = [null, null, null, null, null];
  state.isSubmitted = false;
  document.getElementById('progress-fill').style.width = '0%';
  switchScreen('intro-screen');
}

function copyResultsToClipboard() {
  const studentInfo = `${state.studentId} ${state.studentName}`;
  const score = document.getElementById('score-counter').textContent;
  const duration = document.getElementById('stat-duration').textContent;
  const accuracy = document.getElementById('stat-accuracy').textContent;

  const text = `[속담 온(ON) - 고1 국어 어휘력 모의평가]\n` +
    `학생: ${studentInfo}\n` +
    `최종 점수: ${score}점 / 100점 (정답률 ${accuracy})\n` +
    `소요 시간: ${duration}\n` +
    `평가 완료 및 성적 저장 완료`;

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(() => {
      showToast("📋 성적표가 클립보드에 복사되었습니다!");
    }).catch(() => fallbackCopy(text));
  } else {
    fallbackCopy(text);
  }
}

function fallbackCopy(text) {
  const ta = document.createElement('textarea');
  ta.value = text;
  document.body.appendChild(ta);
  ta.select();
  try {
    document.execCommand('copy');
    showToast("📋 성적표가 클립보드에 복사되었습니다!");
  } catch(e) {
    showToast("복사에 실패했습니다.");
  }
  document.body.removeChild(ta);
}

function formatDuration(seconds) {
  const mins = Math.floor(seconds / 60);
  const secs = seconds % 60;
  return mins > 0 ? `${mins}분 ${secs}초` : `${secs}초`;
}

function formatDate(d) {
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function UtilitiesDateStr() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
}

function showToast(message) {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.classList.remove('hidden');
  setTimeout(() => {
    toast.classList.add('hidden');
  }, 2400);
}
