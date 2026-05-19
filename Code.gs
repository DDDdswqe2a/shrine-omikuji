// ============================================================
// 신사 악귀의 속삭임 — Apps Script Backend
// ============================================================
// 설정: 아래 두 값만 수정하면 됩니다
const CONFIG = {
  SPREADSHEET_ID : 'YOUR_SPREADSHEET_ID_HERE',  // 구글 시트 ID
  ADMIN_KEY      : 'shrine2026admin',            // 관리자 비밀번호 (변경 권장)
  SHEET_ANALYTICS: '방문자 분석',
  SHEET_REGISTER : '이벤트 응모',
};

// ─── GET 핸들러 ────────────────────────────────────────────
function doGet(e) {
  const p = e.parameter;

  // 방문자 추적 (인증 불필요)
  if (p.action === 'track') {
    try {
      const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      const sheet = getOrCreate(ss, CONFIG.SHEET_ANALYTICS,
        ['타임스탬프','이벤트','세부정보','운세결과','기기','레퍼러','체류시간(초)']);
      sheet.appendRow([
        new Date(), p.event||'', p.detail||'', p.fortune||'',
        p.device||'', p.ref||'direct', p.dwell||''
      ]);
    } catch(err) { Logger.log('track error: '+err); }
    return ContentService.createTextOutput('OK');
  }

  // 관리자 인증
  if (p.key !== CONFIG.ADMIN_KEY) {
    return json({ error: 'Unauthorized' });
  }

  if (p.action === 'stats')         return json(getStats());
  if (p.action === 'registrations') return json(getRegistrations());
  if (p.action === 'daily')         return json(getDailyStats());

  return json({ error: 'unknown action' });
}

// ─── POST 핸들러 (이벤트 응모) ─────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.parameter.data || '{}');
    if (data.action === 'register') {
      const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      const sheet = getOrCreate(ss, CONFIG.SHEET_REGISTER,
        ['응모시간','이름','연락처','메일','운세결과','운세내용','마케팅동의','기기']);
      sheet.appendRow([
        new Date(), data.name||'', data.contact||'', data.email||'',
        data.fortune||'', data.fortuneText||'',
        data.marketing ? '동의' : '미동의', data.device||''
      ]);
    }
  } catch(err) { Logger.log('register error: '+err); }
  return json({ success: true });
}

// ─── 통계 요약 ──────────────────────────────────────────────
function getStats() {
  const ss     = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const aSheet = ss.getSheetByName(CONFIG.SHEET_ANALYTICS);
  const rSheet = ss.getSheetByName(CONFIG.SHEET_REGISTER);

  let pageviews=0, draws=0, shares=0, events=0, dwellTotal=0, dwellCount=0;

  if (aSheet && aSheet.getLastRow() > 1) {
    aSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[1]==='pageview')     pageviews++;
      if (r[1]==='fortune_draw') draws++;
      if (r[1]==='share_click')  shares++;
      if (r[1]==='page_exit' && r[6]) { dwellTotal += Number(r[6])||0; dwellCount++; }
    });
  }

  const registrations = rSheet && rSheet.getLastRow()>1 ? rSheet.getLastRow()-1 : 0;
  const avgDwell      = dwellCount>0 ? Math.round(dwellTotal/dwellCount) : 0;

  return {
    pageviews, draws, shares, registrations, avgDwell,
    sheetsUrl: 'https://docs.google.com/spreadsheets/d/'+CONFIG.SPREADSHEET_ID
  };
}

// ─── 응모자 목록 (최신순) ───────────────────────────────────
function getRegistrations() {
  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_REGISTER);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  return sheet.getDataRange().getValues().slice(1)
    .map((r, i) => ({
      no     : i+1,
      name   : r[1], contact: r[2], email: r[3],
      fortune: r[4], marketing: r[6],
      time   : r[0] ? Utilities.formatDate(new Date(r[0]),'Asia/Seoul','MM/dd HH:mm') : ''
    }))
    .reverse();
}

// ─── 일별 통계 (최근 14일) ──────────────────────────────────
function getDailyStats() {
  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_ANALYTICS);
  if (!sheet || sheet.getLastRow() <= 1)
    return { dates:[], pageviews:[], draws:[], shares:[] };

  const byDate = {};
  sheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0]) return;
    const d = Utilities.formatDate(new Date(r[0]),'Asia/Seoul','MM/dd');
    if (!byDate[d]) byDate[d] = { pageviews:0, draws:0, shares:0 };
    if (r[1]==='pageview')     byDate[d].pageviews++;
    if (r[1]==='fortune_draw') byDate[d].draws++;
    if (r[1]==='share_click')  byDate[d].shares++;
  });

  const dates = Object.keys(byDate).slice(-14);
  return {
    dates,
    pageviews: dates.map(d => byDate[d].pageviews),
    draws    : dates.map(d => byDate[d].draws),
    shares   : dates.map(d => byDate[d].shares),
  };
}

// ─── 유틸 ───────────────────────────────────────────────────
function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreate(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length)
      .setBackground('#1a0a08').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ─── 최초 설정 (Apps Script 편집기에서 1회 실행) ────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  getOrCreate(ss, CONFIG.SHEET_ANALYTICS,
    ['타임스탬프','이벤트','세부정보','운세결과','기기','레퍼러','체류시간(초)']);
  getOrCreate(ss, CONFIG.SHEET_REGISTER,
    ['응모시간','이름','연락처','메일','운세결과','운세내용','마케팅동의','기기']);
  Logger.log('✅ 시트 설정 완료!');
}
