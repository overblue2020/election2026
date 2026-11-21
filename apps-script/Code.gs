/**
 * 決戰 2026 預測積分賽：收件程式
 * 貼到 Apps Script 專案，部署成「網頁應用程式」。
 * 每份預測會新增一列到「預測名單」工作表。
 */

// 收件試算表的 ID（網址 /d/ 與 /edit 之間那段）
const SPREADSHEET_ID = '1VvcyPCfXGYOOrgpLO6jR7gXryaDYoL3lkKvBbndQ4lc';
const SHEET_NAME = '預測名單';
const TEST_SHEET_NAME = '測試名單'; // test.html 送出的資料寫這裡，不計入正式比賽
const DEADLINE = new Date('2026-11-21T23:59:00+08:00');
const INVITE_SHEET = '邀請碼';
// 邀請碼只填在 Apps Script 裡，不要提交到 GitHub。格式：['0001', 'XXXXXXXX'], ...
// 首次執行時會建立「邀請碼」工作表；之後新增或停用邀請碼，直接改那張工作表即可。
const INVITE_CODES = [];

const COUNTIES = [
  '基隆市', '桃園市', '新竹市', '新竹縣', '苗栗縣', '台中市', '彰化縣', '南投縣', '宜蘭縣', '花蓮縣',
  '雲林縣', '嘉義市', '嘉義縣', '台南市', '屏東縣', '台東縣', '澎湖縣', '金門縣', '連江縣'
];

const HEADER = [
  '送出時間', '暱稱', '台北市', '台北票差', '新北市', '高雄市',
  ...COUNTIES,
  '加分1 佀廣洋', '加分2 得票數', '加分3 綠營席次', '加分3 泛藍席次', '邀請編號'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return reply({ ok: false, error: 'busy' });
  }
  try {
    if (new Date() > DEADLINE) return reply({ ok: false, error: 'closed' });

    const d = JSON.parse(e.postData.contents);
    const test = d.test === true;
    // 邀請碼改為選填：有帶就核對並記錄，沒帶也可送出
    let invite = null;
    if (d.key) {
      invite = findInvite(d.key);
      if (!invite) return reply({ ok: false, error: 'badkey' });
      if (invite.used) return reply({ ok: false, error: 'used' });
    }
    const nickname = String(d.nickname || '').trim().slice(0, 20);
    const counties = d.counties || {};
    const complete =
      nickname && d.taipei && Number(d.margin) > 0 && d.ntpc && d.khh &&
      COUNTIES.every(c => counties[c]) &&
      (d.q1 === '是' || d.q1 === '否') && Number(d.q2) >= 0 &&
      Number(d.q3) >= 0 && Number(d.q3) <= 22;
    if (!complete) return reply({ ok: false, error: 'invalid' });

    const sheet = getSheet(test);
    const last = sheet.getLastRow();
    if (last > 1) {
      const names = sheet.getRange(2, 2, last - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase());
      if (names.indexOf(nickname.toLowerCase()) !== -1) return reply({ ok: false, error: 'duplicate' });
    }

    sheet.appendRow([
      new Date(),
      "'" + nickname, // 前置單引號，避免暱稱被當成公式
      d.taipei, Number(d.margin), d.ntpc, d.khh,
      ...COUNTIES.map(c => counties[c]),
      d.q1 === '是' ? '當選' : '落選',
      Number(d.q2),
      Number(d.q3),
      22 - Number(d.q3),
      invite ? invite.label : ''
    ]);
    if (invite) invite.sheet.getRange(invite.row, 3, 1, 3).setValues([['已使用', new Date(), nickname]]);
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.check !== undefined) {
    const inv = findInvite(p.check);
    return reply({ ok: true, valid: !!inv, used: !!(inv && inv.used), open: new Date() <= DEADLINE });
  }
  if (p.stats) return reply(stats(p.test === '1'));
  return reply({ ok: true });
}

function getInviteSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(INVITE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(INVITE_SHEET);
    sh.appendRow(['編號', '邀請碼', '狀態', '送出時間', '暱稱']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    if (INVITE_CODES.length) {
      sh.getRange(2, 1, INVITE_CODES.length, 2).setNumberFormat('@').setValues(INVITE_CODES);
    }
  }
  return sh;
}

// 依邀請碼找到那一列；狀態欄填「停用」即可讓該碼失效
function findInvite(code) {
  code = String(code || '').trim().toUpperCase();
  if (!code) return null;
  const sh = getInviteSheet();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const rows = sh.getRange(2, 1, last - 1, 3).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim().toUpperCase() === code) {
      const status = String(rows[i][2]).trim();
      if (status === '停用') return null;
      return { sheet: sh, row: i + 2, label: String(rows[i][0]), used: status === '已使用' };
    }
  }
  return null;
}

// 匯總統計（不含暱稱），供網頁「戰況總覽」使用
function stats(test) {
  const sheet = getSheet(test);
  const last = sheet.getLastRow();
  const raceCols = { '台北市': 2, '新北市': 4, '高雄市': 5 };
  COUNTIES.forEach((c, i) => { raceCols[c] = 6 + i; });
  const out = { ok: true, n: 0, races: {}, margin: { avg: 0, median: 0 }, q1: { yes: 0, no: 0 }, q2: { avg: 0 }, q3: { avg: 0 } };
  Object.keys(raceCols).forEach(r => { out.races[r] = {}; });
  if (last < 2) return out;
  const rows = sheet.getRange(2, 1, last - 1, HEADER.length).getValues().filter(r => String(r[1]).trim());
  out.n = rows.length;
  if (!rows.length) return out;
  rows.forEach(r => {
    Object.keys(raceCols).forEach(name => {
      const v = String(r[raceCols[name]]);
      out.races[name][v] = (out.races[name][v] || 0) + 1;
    });
    if (r[25] === '當選') out.q1.yes++; else out.q1.no++;
  });
  const nums = col => rows.map(r => Number(r[col])).filter(x => !isNaN(x));
  const avg = a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0;
  const m = nums(3).sort((a, b) => a - b);
  out.margin.avg = Math.round(avg(m));
  out.margin.median = m.length % 2 ? m[(m.length - 1) / 2] : Math.round((m[m.length / 2 - 1] + m[m.length / 2]) / 2);
  out.q2.avg = Math.round(avg(nums(26)));
  out.q3.avg = Math.round(avg(nums(27)) * 10) / 10;
  return out;
}

function getSheet(test) {
  const name = test ? TEST_SHEET_NAME : SHEET_NAME;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() > 0 && sheet.getRange(1, HEADER.length).getValue() !== HEADER[HEADER.length - 1]) {
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
    sheet.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
  }
  return sheet;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 依陣營替候選人欄位上色（條件式格式，之後新增的資料也會自動套用）
// 在 Apps Script 編輯器選這個函式按「執行」即可；重複執行會覆蓋為最新設定。
const GREEN_CANDIDATES = ['沈伯洋', '蘇巧慧', '賴瑞隆', '童子瑋', '黃世杰', '莊競程', '鄭朝方', '陳品安', '何欣純',
  '陳素月', '温世政', '林國漳', '劉建國', '王美惠', '蔡易餘', '陳亭妃', '周春米', '陳瑩', '吳淑瑾'];

function colorCamps() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const green = { bg: '#c8e6c9', fg: '#1b5e20' };
  const blue = { bg: '#cfe0fc', fg: '#0d3c8c' };
  const pattern = '^(' + GREEN_CANDIDATES.join('|') + ')$';
  [SHEET_NAME, TEST_SHEET_NAME].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const rows = Math.max(sh.getMaxRows(), 1000);
    // C 台北、E 新北、F 高雄、G–Y 其餘 19 縣市（D 是台北票差，不上色）
    const cand = [sh.getRange(2, 3, rows - 1, 1), sh.getRange(2, 5, rows - 1, 21)];
    const isGreen = 'REGEXMATCH(TO_TEXT(C2),"' + pattern + '")';
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=' + isGreen)
        .setBackground(green.bg).setFontColor(green.fg).setRanges(cand).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(C2<>"",NOT(' + isGreen + '))')
        .setBackground(blue.bg).setFontColor(blue.fg).setRanges(cand).build(),
      // 加分3：綠營席次、泛藍席次
      SpreadsheetApp.newConditionalFormatRule().whenCellNotEmpty()
        .setBackground(green.bg).setFontColor(green.fg).setRanges([sh.getRange(2, 28, rows - 1, 1)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenCellNotEmpty()
        .setBackground(blue.bg).setFontColor(blue.fg).setRanges([sh.getRange(2, 29, rows - 1, 1)]).build()
    ];
    sh.setConditionalFormatRules(rules);
  });
}
