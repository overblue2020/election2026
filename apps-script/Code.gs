/**
 * 決戰 2026 預測積分賽：收件程式
 * 貼到 Google 試算表的「擴充功能 > Apps Script」，部署成「網頁應用程式」。
 * 每份預測會新增一列到「預測名單」工作表。
 */

const SHEET_NAME = '預測名單';
const DEADLINE = new Date('2026-11-27T23:59:00+08:00');

const COUNTIES = [
  '基隆市', '桃園市', '新竹市', '新竹縣', '苗栗縣', '台中市', '彰化縣', '南投縣', '宜蘭縣', '花蓮縣',
  '雲林縣', '嘉義市', '嘉義縣', '台南市', '屏東縣', '台東縣', '澎湖縣', '金門縣', '連江縣'
];

const HEADER = [
  '送出時間', '暱稱', '台北市', '台北票差', '新北市', '高雄市',
  ...COUNTIES,
  '加分1 佀廣洋', '加分2 得票數', '加分3 綠營席次', '加分3 泛藍席次'
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
    const nickname = String(d.nickname || '').trim().slice(0, 20);
    const counties = d.counties || {};
    const complete =
      nickname && d.taipei && Number(d.margin) > 0 && d.ntpc && d.khh &&
      COUNTIES.every(c => counties[c]) &&
      (d.q1 === '是' || d.q1 === '否') && Number(d.q2) >= 0 &&
      Number(d.q3) >= 0 && Number(d.q3) <= 22;
    if (!complete) return reply({ ok: false, error: 'invalid' });

    const sheet = getSheet();
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
      22 - Number(d.q3)
    ]);
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return reply({ ok: true, open: new Date() <= DEADLINE });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
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
