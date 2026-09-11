// 수능 고전문학 <보기> 키워드 매핑 (c) 2026 KIMYERANG.
/**************************************************************
 * <보기> 키워드-원문 시어 매핑 훈련 웹앱 (Google Apps Script)
 * - 학생: 공개된 문제만 조회, 클릭-클릭으로 매핑 연습, 즉시 채점/해설
 * - 교사: 평문(<보기>, 원문)만 입력 → AI가 매핑 쌍 초안 생성 → 검수/보정 후 저장
 * - 문제는 이 스크립트가 바인딩된 스프레드시트의 'Problems' 시트에 저장(=구글 드라이브에 저장)
 * - 교사가 개별 문제 단위로 공개/비공개를 켜고 꺼서 학생에게 보이는 문제를 통제
 *
 * 배포: Code.gs + Shared.html + Admin.html + Student.html → 웹 앱(실행:나, 액세스:모든 사용자)
 **************************************************************/

/* ===================== 설정 =====================
 * 관리자 비밀번호: [프로젝트 설정 → 스크립트 속성] → ADMIN_PASSWORD
 * Gemini API 키   : [프로젝트 설정 → 스크립트 속성] → GEMINI_API_KEY
 * (둘 다 코드에 두지 않는다 — 사본에 딸려가지 않게 하기 위함, training-register와 동일 관례)
 */
const SHEET_NAME = 'Problems';
const HEADERS = ['id', 'title', 'guideText', 'originalText', 'pairsJson', 'published', 'createdAt', 'mcJson'];
const GEMINI_MODEL = 'gemini-3.6-flash';

const COL = { ID: 1, TITLE: 2, GUIDE: 3, ORIGINAL: 4, PAIRS: 5, PUBLISHED: 6, CREATED: 7, MC: 8 };

/* ===================== 라우팅 ===================== */
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const fileName = (p.mode === 'admin') ? 'Admin' : 'Student';
  const title = (p.mode === 'admin') ? '키워드 매핑 · 문제 관리' : '수능 고전문학 <보기> 키워드 매핑';
  return HtmlService.createTemplateFromFile(fileName).evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Admin.html/Student.html에서 <?!= include('Shared') ?>로 공통 CSS/JS를 끼워 넣는다 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ===================== 관리자 인증 =====================
 * training-register와 동일 관례: 비밀번호는 코드가 아니라 스크립트 속성에 둔다.
 */
function getAdminPassword_() {
  const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!pw) {
    throw new Error('관리자 비밀번호가 설정되지 않았습니다. '
      + '[프로젝트 설정 → 스크립트 속성]에서 ADMIN_PASSWORD를 추가하세요.');
  }
  return pw;
}

function requireAdmin_(params) {
  if (!params || String(params.password || '') !== getAdminPassword_()) {
    throw new Error('비밀번호가 올바르지 않습니다.');
  }
}

function verifyAdmin(pw) {
  return String(pw || '') === getAdminPassword_();
}

/* ===================== 시트 ===================== */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('이 스크립트는 스프레드시트에 연결되어 있어야 합니다. '
      + '스프레드시트에서 [확장 프로그램 → Apps Script]로 열어 배포하세요.');
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    seedDefaultProblem_(sheet);
  }
  return sheet;
}

/** 처음 시트를 만들 때 예시 문제 1개를 비공개 초안으로 넣어준다(빈 화면 방지 + 사용법 예시) */
function seedDefaultProblem_(sheet) {
  const guideText = '작가는 당대 정치적 갈등으로 인해 억울하게 유배를 간 상황에 놓여 있다. 그는 절망적인 처지 속에서도 임금에 대한 변함없는 충성심을 드러내며 자신의 결백을 호소하고 있다.';
  const originalText = '천상 백옥경 십이루 어듸매오\n오색운 깊은 곳에 자청전이 가렷으니\n천맥이 아득하니 지척인들 이내 알며\n오늘도 다 새거다 이불 안고 일어 앉아\n창을 열고 바라보니 강호 삼월 다 지나가고\n적객의 이 셜운 듸 어느덧 다 지나거다\n매화나 보내고져 임 계신 데 바라보니\n산은 어이 높은지고 물은 어이 넓은지고';
  const pairs = [
    { g: '억울하게 유배를 간 상황', o: '적객', exp: '화자 자신을 귀양 간 사람인 적객으로 표현하여 유배지에서의 처지를 나타냅니다.' },
    { g: '변함없는 충성심', o: '매화', exp: '매화는 시련 속에서도 꺾이지 않는 절개·충절을 상징하는 소재입니다.' }
  ];
  const mc = {
    question: '<보기>를 바탕으로 윗글을 감상한 내용으로 적절하지 않은 것은?',
    options: [
      '화자는 유배지에서도 임금에 대한 충심을 버리지 않고 있다.',
      "'적객'은 화자가 자신의 처지를 가리키는 표현이다.",
      "'매화'는 임금에 대한 화자의 변함없는 절개를 상징한다.",
      '화자는 자신의 결백이 이미 밝혀져 홀가분한 심정을 드러내고 있다.',
      '화자는 임과 멀리 떨어진 상황을 산과 물의 높고 넓음에 빗대어 표현하고 있다.'
    ],
    answerIndex: 3,
    explanation: '화자는 자신의 결백을 호소하는 처지일 뿐 결백이 밝혀진 것이 아니므로, 홀가분한 심정이라는 감상은 적절하지 않습니다.'
  };
  sheet.appendRow([Utilities.getUuid(), "조위, '만분가' (예시)", guideText, originalText, JSON.stringify(pairs), false, new Date(), JSON.stringify(mc)]);
}

function readAllRows_() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .map(function (r, i) { return { row: i + 2, values: r }; })
    .filter(function (x) { return x.values[COL.ID - 1]; }); // id 빈 행(중간 공백) 제외
}

function rowToProblem_(values) {
  let pairs = [];
  try { pairs = JSON.parse(values[COL.PAIRS - 1] || '[]'); } catch (e) { pairs = []; }
  let mc = null;
  try { mc = values[COL.MC - 1] ? JSON.parse(values[COL.MC - 1]) : null; } catch (e) { mc = null; }
  return {
    id: values[COL.ID - 1],
    title: values[COL.TITLE - 1],
    guideText: values[COL.GUIDE - 1],
    originalText: values[COL.ORIGINAL - 1],
    pairs: pairs,
    mc: mc,
    published: values[COL.PUBLISHED - 1] === true || values[COL.PUBLISHED - 1] === 'TRUE',
    createdAt: values[COL.CREATED - 1]
  };
}

/* ===================== 학생용: 공개된 문제만 ===================== */
function getPublishedProblems() {
  return readAllRows_()
    .map(function (x) { return rowToProblem_(x.values); })
    .filter(function (p) { return p.published; })
    .sort(function (a, b) { return new Date(a.createdAt) - new Date(b.createdAt); });
}

/* ===================== 교사용: 전체 문제 ===================== */
function getAllProblems(params) {
  requireAdmin_(params);
  return readAllRows_()
    .map(function (x) { return rowToProblem_(x.values); })
    .sort(function (a, b) { return new Date(b.createdAt) - new Date(a.createdAt); });
}

/* ===================== 교사용: 저장/수정 ===================== */
function saveProblem(params) {
  requireAdmin_(params);
  const data = params.data || {};
  if (!data.title || !data.guideText || !data.originalText) {
    throw new Error('작품명, <보기>, 원문을 모두 입력하세요.');
  }
  const pairs = Array.isArray(data.pairs) ? data.pairs : [];
  if (!pairs.length) {
    throw new Error('정답 쌍이 하나도 없습니다. AI 생성 또는 수동 추가로 최소 1개 이상 만드세요.');
  }
  for (const p of pairs) {
    if (data.guideText.indexOf(p.g) === -1 || data.originalText.indexOf(p.o) === -1) {
      throw new Error('"' + p.g + '" / "' + p.o + '" 쌍의 문구가 본문에 정확히 존재하지 않습니다.');
    }
  }
  const pairsJson = JSON.stringify(pairs.map(function (p) {
    return { g: String(p.g), o: String(p.o), exp: String(p.exp || '') };
  }));

  let mcJson = '';
  if (data.mc) {
    const mc = data.mc;
    const opts = Array.isArray(mc.options) ? mc.options.map(String) : [];
    const idx = Number(mc.answerIndex);
    if (!mc.question || opts.length !== 5 || opts.some(function (o) { return !o.trim(); })
      || !(idx >= 0 && idx <= 4) || !mc.explanation) {
      throw new Error('2단계 감상 문제는 질문·선택지 5개·정답·해설을 모두 채우거나, 비워서 2단계 없이 저장하세요.');
    }
    mcJson = JSON.stringify({ question: String(mc.question), options: opts, answerIndex: idx, explanation: String(mc.explanation) });
  }

  const sheet = getSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const rows = readAllRows_();
    const existing = data.id ? rows.find(function (x) { return x.values[COL.ID - 1] === data.id; }) : null;
    if (existing) {
      sheet.getRange(existing.row, COL.TITLE).setValue(data.title);
      sheet.getRange(existing.row, COL.GUIDE).setValue(data.guideText);
      sheet.getRange(existing.row, COL.ORIGINAL).setValue(data.originalText);
      sheet.getRange(existing.row, COL.PAIRS).setValue(pairsJson);
      sheet.getRange(existing.row, COL.MC).setValue(mcJson);
      return { id: data.id };
    }
    const id = Utilities.getUuid();
    sheet.appendRow([id, data.title, data.guideText, data.originalText, pairsJson, false, new Date(), mcJson]);
    return { id: id };
  } finally {
    lock.releaseLock();
  }
}

function deleteProblem(params) {
  requireAdmin_(params);
  const rows = readAllRows_();
  const target = rows.find(function (x) { return x.values[COL.ID - 1] === params.id; });
  if (!target) throw new Error('문제를 찾을 수 없습니다.');
  getSheet_().deleteRow(target.row);
  return { ok: true };
}

function togglePublish(params) {
  requireAdmin_(params);
  const rows = readAllRows_();
  const target = rows.find(function (x) { return x.values[COL.ID - 1] === params.id; });
  if (!target) throw new Error('문제를 찾을 수 없습니다.');
  getSheet_().getRange(target.row, COL.PUBLISHED).setValue(!!params.published);
  return { ok: true };
}

/* ===================== 교사용: AI 매핑 초안 생성 =====================
 * 서버(Code.gs)에서 Gemini를 호출한다 — Admin.html은 학생도 열어볼 수 있는
 * 웹앱 클라이언트 코드이므로, 거기에 API 키를 두면 그대로 노출된다.
 */
function generatePairsWithAI(params) {
  requireAdmin_(params);
  const guideText = String(params.guideText || '').trim();
  const originalText = String(params.originalText || '').trim();
  if (!guideText || !originalText) throw new Error('<보기>와 원문을 먼저 입력하세요.');

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('Gemini API 키가 설정되지 않았습니다. '
      + '[프로젝트 설정 → 스크립트 속성]에서 GEMINI_API_KEY를 추가하거나, 아래에서 수동으로 쌍을 추가하세요.');
  }

  const prompt = '당신은 수능 국어 고전문학 출제 전문가입니다.\n'
    + '아래 <보기>와 원문(고전시가/고전산문)을 읽고, <보기>에서 설명하는 내용과 원문 속 시어가 서로 대응하는 '
    + '핵심어-시어 쌍을 3~6개 찾아주세요.\n\n'
    + '[<보기>]\n' + guideText + '\n\n'
    + '[원문]\n' + originalText + '\n\n'
    + '[중요 규칙]\n'
    + '1. guideKeyword는 <보기> 원문에서, originalKeyword는 원문에서 한 글자도 바꾸지 말고 그대로 복사한 부분 문자열이어야 합니다.\n'
    + '2. 각 구절은 너무 길지 않게(핵심 어절·구 단위)로 끊어주세요.\n'
    + '3. explanation은 왜 두 구절이 대응하는지 수능 국어 학습자가 이해할 수 있게 1~2문장으로 설명하세요.\n'
    + '4. 이미 사용한 원문 구절을 다른 쌍에 중복 사용하지 마세요.';

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            guideKeyword: { type: 'STRING' },
            originalKeyword: { type: 'STRING' },
            explanation: { type: 'STRING' }
          },
          required: ['guideKeyword', 'originalKeyword', 'explanation']
        }
      }
    }
  };

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL
    + ':generateContent?key=' + encodeURIComponent(apiKey);
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('AI 호출 실패(' + resp.getResponseCode() + '): ' + resp.getContentText().slice(0, 300));
  }

  const body = JSON.parse(resp.getContentText());
  const text = body.candidates && body.candidates[0] && body.candidates[0].content
    && body.candidates[0].content.parts && body.candidates[0].content.parts[0]
    && body.candidates[0].content.parts[0].text;
  if (!text) throw new Error('AI 응답이 비어 있습니다.');

  const raw = parseAiJson_(text);
  return raw.map(function (p) {
    return { g: String(p.guideKeyword || '').trim(), o: String(p.originalKeyword || '').trim(), exp: String(p.explanation || '').trim() };
  }).filter(function (p) { return p.g && p.o; });
}

/* ===================== 교사용: AI 2단계 감상 문제 생성 =====================
 * 매칭(1단계)을 마친 학생이 그 힌트로 원문 전체를 실제로 이해했는지 확인하는
 * 수능형 5지선다("적절하지 않은 것은?") 문제를 만든다. 확정된 매칭 쌍이 있으면
 * 함께 넘겨 AI가 그 맥락 위에서 문제를 만들게 한다.
 */
function generateMcQuestionWithAI(params) {
  requireAdmin_(params);
  const guideText = String(params.guideText || '').trim();
  const originalText = String(params.originalText || '').trim();
  const pairs = Array.isArray(params.pairs) ? params.pairs : [];
  if (!guideText || !originalText) throw new Error('<보기>와 원문을 먼저 입력하세요.');

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('Gemini API 키가 설정되지 않았습니다. '
      + '[프로젝트 설정 → 스크립트 속성]에서 GEMINI_API_KEY를 추가하거나, 아래에서 수동으로 입력하세요.');
  }

  const pairsDesc = pairs.map(function (p) { return '- ' + p.g + ' ↔ ' + p.o + (p.exp ? ' (' + p.exp + ')' : ''); }).join('\n') || '(없음)';

  const prompt = '당신은 수능 국어 고전문학 출제 전문가입니다.\n'
    + '아래 <보기>와 원문, 그리고 이미 확인된 핵심어 대응 쌍을 참고하여, '
    + '"<보기>를 바탕으로 윗글을 감상한 내용으로 적절하지 않은 것은?" 형태의 5지선다 문제를 하나 만드세요.\n\n'
    + '[<보기>]\n' + guideText + '\n\n'
    + '[원문]\n' + originalText + '\n\n'
    + '[확인된 대응 쌍]\n' + pairsDesc + '\n\n'
    + '[중요 규칙]\n'
    + '1. 선택지는 정확히 5개를 만드세요.\n'
    + '2. 4개는 <보기>와 원문 내용에 부합하는 올바른 감상, 1개는 <보기>나 원문 내용과 어긋나는 틀린 감상으로 만드세요.\n'
    + '3. 틀린 선택지도 매력적인 오답이 되도록, 원문을 제대로 이해해야만 구별할 수 있게 만드세요.\n'
    + '4. answerIndex는 틀린 선택지의 0부터 시작하는 인덱스입니다.\n'
    + '5. explanation은 그 선택지가 왜 틀렸는지 수능 국어 학습자가 이해할 수 있게 1~2문장으로 설명하세요.';

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          question: { type: 'STRING' },
          options: { type: 'ARRAY', items: { type: 'STRING' } },
          answerIndex: { type: 'INTEGER' },
          explanation: { type: 'STRING' }
        },
        required: ['question', 'options', 'answerIndex', 'explanation']
      }
    }
  };

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL
    + ':generateContent?key=' + encodeURIComponent(apiKey);
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('AI 호출 실패(' + resp.getResponseCode() + '): ' + resp.getContentText().slice(0, 300));
  }

  const body = JSON.parse(resp.getContentText());
  const text = body.candidates && body.candidates[0] && body.candidates[0].content
    && body.candidates[0].content.parts && body.candidates[0].content.parts[0]
    && body.candidates[0].content.parts[0].text;
  if (!text) throw new Error('AI 응답이 비어 있습니다.');

  const raw = parseAiJson_(text);
  const options = Array.isArray(raw.options) ? raw.options.map(function (o) { return String(o).trim(); }) : [];
  if (options.length !== 5) throw new Error('AI가 선택지 5개를 만들지 못했습니다. 다시 시도해 주세요.');
  const answerIndex = Number(raw.answerIndex);
  if (!(answerIndex >= 0 && answerIndex <= 4)) throw new Error('AI 응답의 정답 인덱스가 올바르지 않습니다. 다시 시도해 주세요.');

  return {
    question: String(raw.question || '').trim(),
    options: options,
    answerIndex: answerIndex,
    explanation: String(raw.explanation || '').trim()
  };
}

/** 마크다운 코드펜스가 섞여 와도 견디는 JSON 파싱(hwp_gemini_generator/gemini_client.py와 동일 관례) */
function parseAiJson_(text) {
  let t = String(text).trim();
  if (t.indexOf('```json') === 0) t = t.slice(7);
  else if (t.indexOf('```') === 0) t = t.slice(3);
  if (t.slice(-3) === '```') t = t.slice(0, -3);
  t = t.trim();
  try { return JSON.parse(t); }
  catch (e) { throw new Error('AI 응답 JSON 파싱 실패: ' + t.slice(0, 300)); }
}

/* ===================== 자가진단(수동 실행용) =====================
 * Apps Script 편집기 함수 목록에서 이 함수를 선택해 실행하면
 * parseAiJson_의 코드펜스 처리 로직이 깨지지 않았는지 로그로 확인할 수 있다.
 */
function 자가진단_AI파싱() {
  let ok = 0;
  [
    ['[{"guideKeyword":"a","originalKeyword":"b","explanation":"c"}]', 1],
    ['```json\n[{"guideKeyword":"a","originalKeyword":"b","explanation":"c"}]\n```', 1],
    ['```\n[]\n```', 0]
  ].forEach(function (c) {
    const parsed = parseAiJson_(c[0]);
    const pass = Array.isArray(parsed) && parsed.length === c[1];
    Logger.log((pass ? 'PASS ' : 'FAIL ') + c[0].slice(0, 20) + '...');
    if (pass) ok++;
  });
  Logger.log(ok + '/3 통과');
}
