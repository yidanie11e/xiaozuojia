/**
 * 小作家幫手 · 匿名使用數據接收端（Google Apps Script）
 *
 * 用法：
 * 1. 開一個 Google Sheet（例如「小作家使用數據」）
 * 2. 擴充功能 → Apps Script → 把本檔內容貼進 Code.gs
 * 3. 部署 → 新增部署作業 → 類型「網頁應用程式」
 *    執行身分：我    誰可以存取：任何人
 * 4. 複製 /exec 網址，貼到 index.html 的 STATS_ENDPOINT
 *
 * 只會收到：時間、模組、成功與否、耗時、輸入字數、輸出字數、隨機裝置代號
 * 不會收到學生輸入的文字或 AI 的回覆內容。
 */

const SHEET_NAME = 'usage';
const HEADERS = ['時間', '模組', '狀態', '耗時(ms)', '輸入字數', '輸出字數', '裝置代號', '收到時間'];

const MODULE_NAMES = {
  word: '詞語百寶箱',
  sentence: '句子小診所',
  expression: '五感魔法棒'
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = JSON.parse(e.postData.contents);
    getSheet().appendRow([
      new Date(data.ts),
      MODULE_NAMES[data.event] || data.event,
      data.ok ? '成功' : '失敗',
      Number(data.ms) || 0,
      Number(data.inLen) || 0,
      Number(data.outLen) || 0,
      data.sid || '',
      new Date()
    ]);
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err);
  } finally {
    lock.releaseLock();
  }
}

// 方便用瀏覽器直接打開 /exec 確認部署成功
function doGet() {
  return ContentService.createTextOutput('stats endpoint ok');
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
