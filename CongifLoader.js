/**
 * @description 讀取 'Config' 分頁的所有設定，並將其轉換成一個鍵值對物件。
 * 這個函式會在專案啟動時自動執行一次，並將結果存放在全域常數 CONFIG 中。
 * @returns {Object} 包含所有設定的物件。例如：{ MSG_HELP: '你好...', MSG_CANCEL_CONFIRM: '好的...' }
 */
function loadConfig() {
  try {
    // 從指令碼屬性中，取得我們設定好的 SPREADSHEET_ID
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error("指令碼屬性中找不到 'SPREADSHEET_ID'。");
    }
    
    // 使用取得的 ID，明確地打開指定的試算表檔案
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('Config');
    
    if (!sheet) {
      throw new Error("在試算表中找不到名為 'Config' 的分頁。");
    }
    
    // 取得 Config 分頁中所有有資料的儲存格範圍，並獲取其值
    const data = sheet.getDataRange().getValues();
    
    const config = {};
    
    // 從第二行開始迴圈 (i = 1)，跳過第一行的表頭
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const key = row[0];    // A欄 的參數鑰匙
      const value = row[1];  // B欄 的參數內容
      
      // 確保 key 不是空的，才加入到 config 物件中
      if (key) {
        config[key] = value;
      }
    }
    
    console.log('Config loaded successfully!');
    return config;

  } catch (e) {
    // 如果發生任何錯誤，在日誌中詳細記錄，方便除錯
    console.error("讀取 Config 設定時發生嚴重錯誤: " + e.message);
    console.error("錯誤詳情: " + e.stack);
    // 回傳一個空的物件，避免整個專案因讀取失敗而崩潰
    return {};
  }
}

/**
 * @description 一個全域常數，用來存放從 Config 分頁讀取出的所有設定。
 * 整個專案中，我們都會透過 CONFIG 這個常數來取得設定值。
 */
const CONFIG = loadConfig();