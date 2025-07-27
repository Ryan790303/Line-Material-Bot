/**
 * @file Logger.gs
 * @description 負責記錄日誌到 Google Sheet
 */

/**
 * @description 將除錯訊息寫入到試算表的 'Log' 分頁中。
 * @param {string} message - 要記錄的事件訊息。
 * @param {Object|string} [data=''] - (可選) 要記錄的相關資料，物件會自動轉為 JSON 字串。
 */
function log(message, data = '') {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    // --- 修改處 START ---
    // 從 CONFIG 讀取分頁名稱，如果找不到就使用預設值 'Log'
    const logSheetName = CONFIG.SHEET_NAME_LOG || 'Log';
    let logSheet = ss.getSheetByName(logSheetName);

    // 如果分頁不存在，就用設定好的名稱自動建立一個
    if (!logSheet) {
      logSheet = ss.insertSheet(logSheetName);
      logSheet.appendRow(['時間', '事件', '詳細資料']);
    }
    // --- 修改處 END ---
    
    const timestamp = new Date();
    const dataString = (data && typeof data === 'object') ? JSON.stringify(data, null, 2) : data;
    
    logSheet.appendRow([timestamp, message, dataString]);

  } catch (e) {
    // 如果連寫入日誌都失敗，就在原始的執行紀錄中印出錯誤
    console.error("寫入日誌時發生錯誤: " + e.message);
  }
}
