// =================================================================
// SECTION: LINE API Communication Utilities
// 說明：所有與 LINE API 直接溝通的底層工具函式。
// =================================================================

/**
 * @description 取得包含 LINE Access Token 的請求標頭。
 * @param {string} accessToken - Channel Access Token。
 * @returns {Object} 包含 Authorization 標頭的物件。
 */
function getLineHeaders(accessToken) {
  return {
    'Authorization': 'Bearer ' + accessToken
  };
}

/**
 * @description 呼叫 LINE Reply Message API 來回覆訊息 (含錯誤日誌)。
 */
function replyMessage(replyToken, messages, accessToken) {
  try {
    const url = 'https://api.line.me/v2/bot/message/reply';
    const payload = { 'replyToken': replyToken, 'messages': messages };
    const options = {
      'method': 'post', 'contentType': 'application/json',
      'headers': getLineHeaders(accessToken), 'payload': JSON.stringify(payload),
      'muteHttpExceptions': true, // 關鍵：設定為 true 才能手動處理錯誤
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    // --- 新增的日誌紀錄 ---
    if (responseCode !== 200) {
      log(`[${responseCode}] Reply API 呼叫失敗`, response.getContentText());
    }
    // ---
    
  } catch (e) {
    log("Reply Message 函式本身發生錯誤", e.message);
  }
}

/**
 * @description 呼叫 LINE Push Message API 來主動推播訊息 (含錯誤日誌)。
 */
function pushMessage(userId, messages, accessToken) {
  try {
    const url = 'https://api.line.me/v2/bot/message/push';
    const payload = { 'to': userId, 'messages': messages };
    const options = {
      'method': 'post', 'contentType': 'application/json',
      'headers': getLineHeaders(accessToken), 'payload': JSON.stringify(payload),
      'muteHttpExceptions': true, // 關鍵
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    // --- 新增的日誌紀錄 ---
    if (responseCode !== 200) {
      log(`[${responseCode}] Push API 呼叫失敗`, response.getContentText());
    }
    // ---

  } catch (e) {
    log("Push Message 函式本身發生錯誤", e.message);
  }
}

/**
 * @description 取得使用者的 LINE 個人資料。優先從本地 'Users' 分頁快取查詢，若無資料才呼叫 LINE API。
 * @param {string} userId - 使用者的 LINE User ID。
 * @param {string} accessToken - Channel Access Token。
 * @returns {string} - 一個包含 displayName 的 JSON 字串。例如：'{"displayName":"使用者A"}'
 */
function getUserProfile(userId, accessToken) {
  // --- 修改處 START ---
  // 從 Config 讀取預設名稱
  const defaultName = CONFIG.DEFAULT_UNKNOWN_USER || '未知使用者';
  // --- 修改處 END ---

  try {
    const usersMap = getUsersMap();
    if (usersMap.has(userId)) {
      const displayName = usersMap.get(userId);
      return JSON.stringify({ displayName: displayName });
    }

    log('新使用者，開始呼叫 LINE Profile API...', { userId: userId });
    const url = `https://api.line.me/v2/bot/profile/${userId}`;
    const options = { 'method': 'get', 'headers': getLineHeaders(accessToken), 'muteHttpExceptions': true };
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) {
      const profile = JSON.parse(responseText);
      const displayName = profile.displayName;

      const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      if (spreadsheetId) {
          const ss = SpreadsheetApp.openById(spreadsheetId);
          const sheetName = CONFIG.SHEET_NAME_USERS || 'Users';
          const sheet = ss.getSheetByName(sheetName);
          if (sheet) {
            sheet.appendRow([userId, displayName]);
            const cacheKey = CONFIG.CACHE_KEY_USERS || 'users_map';
            CacheService.getScriptCache().remove(cacheKey);
            log(`新使用者 ${displayName} 已成功寫入資料庫。`);
          }
      } else {
        log("寫入新使用者失敗：找不到 SPREADSHEET_ID。");
      }
      
      return responseText;
    } else {
      log(`[${responseCode}] LINE Profile API 呼叫失敗`, responseText);
      // --- 修改處 START ---
      return JSON.stringify({ displayName: defaultName });
      // --- 修改處 END ---
    }
  } catch (e) {
    console.error("Get User Profile 失敗: " + e.message);
    log("Get User Profile 函式本身發生錯誤", e.message);
    // --- 修改處 START ---
    return JSON.stringify({ displayName: defaultName });
    // --- 修改處 END ---
  }
}


// =================================================================
// SECTION: Core Data & Logic Utilities
// 說明：所有與資料庫讀寫、商業邏輯計算相關的工具函式。
// =================================================================

/**
 * @description 從 CONFIG 全域變數中，安全地取得訊息文案、替換變數並修正換行符號。
 * @param {string} key - Config 中的參數鑰匙。
 * @param {Object} replacements - 一個包含要替換的變數物件。
 * @returns {string} - 回傳處理完成的文字。
 */
function getConfigMessage(key, replacements = {}) {
  let message = CONFIG[key] || `[錯誤:找不到設定 ${key}]`;
  for (const placeholder in replacements) {
    message = message.replace(new RegExp(`{${placeholder}}`, 'g'), replacements[placeholder]);
  }
  return message.replace(/\\n/g, '\n');
}


/**
 * @description (智慧版) 將 Google Drive 的分享網址或檔案ID，轉換成可以直接顯示的圖片網址。
 */
function convertGoogleDriveFileIdToDirectUrl(urlOrId) {
  // --- 修改處 START ---
  // 從 Config 讀取預設圖片網址，若找不到則使用後備網址
  const defaultUrl = CONFIG.DEFAULT_IMAGE_URL || 'https://via.placeholder.com/500x300.png?text=No+Image';

  if (!urlOrId || typeof urlOrId !== 'string' || urlOrId.trim() === '') {
    return defaultUrl;
  }
  // --- 修改處 END ---
  
  let fileId = urlOrId;
  if (urlOrId.includes('drive.google.com/file/d/')) {
    fileId = urlOrId.split('/d/')[1].split('/')[0];
  } else if (urlOrId.includes('drive.google.com/open?id=')) {
    fileId = urlOrId.split('id=')[1].split('&')[0];
  }
  if (fileId && fileId.length > 20 && !fileId.includes('http')) {
    return `https://drive.google.com/uc?export=view&id=${fileId}`;
  }
  if (urlOrId.startsWith('http')) {
    return urlOrId;
  }

  // --- 修改處 START ---
  // 如果傳入的 ID 或 URL 格式不符，也回傳預設圖片
  return defaultUrl;
  // --- 修改處 END ---
}


/**
 * @description 讀取資料庫並計算所有物料的即時庫存與詳細資訊 (已加入快取機制)。
 * @returns {Map<string, Object>} - 一個以組合鍵為 key，物料物件為 value 的 Map。
 */
function getInventoryMap() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = CONFIG.CACHE_KEY_INVENTORY || 'inventory_map';

    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      log('從快取讀取庫存資料...');
      return new Map(JSON.parse(cachedData));
    }

    log('快取未命中，從試算表讀取庫存資料...');
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error("指令碼屬性中找不到 'SPREADSHEET_ID'。");
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = CONFIG.SHEET_NAME_RECORDS || '出入庫記錄';
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      log(`找不到庫存記錄分頁 '${sheetName}'。`);
      return new Map();
    }
    
    const allData = sheet.getDataRange().getValues();
    const materialsMap = new Map();

    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      if (row.length > 8 && row[8] === CONFIG.STATUS_VALID) {
        const compositeKey = `${row[0]}${row[1]}`;
        if (!materialsMap.has(compositeKey)) {
          materialsMap.set(compositeKey, {
            分類: row[0], 序號: row[1], 品名: row[2], 型號: row[3], 規格: row[4],
            單位: row[5], 庫存: 0, 照片: (row.length > 12 ? row[12] : '')
          });
        }
        materialsMap.get(compositeKey).庫存 += Number(row[6] || 0);
      }
    }

    // --- 修改處 START ---
    // 從 Config 讀取快取時間，若找不到則預設為 300 秒
    const expiration = Number(CONFIG.CACHE_EXPIRATION_INVENTORY) || 300;
    cache.put(cacheKey, JSON.stringify(Array.from(materialsMap.entries())), expiration);
    // --- 修改處 END ---

    return materialsMap;
  } catch (e) {
    console.error("讀取庫存資料時發生錯誤: " + e.message);
    log("讀取庫存資料時發生錯誤", e.message);
    return new Map();
  }
}


/**
 * @description (模糊搜尋版) 根據查詢詞，找出「品名」包含查詢詞的物料。
 */
function searchMaterials(query) {
  const materialsMap = getInventoryMap();
  const normalizedQuery = query.toLowerCase().replace(/\s/g, '');
  if (!normalizedQuery) return [];

  const results = [];
  for (const material of materialsMap.values()) {
    const normalizedName = material.品名.toLowerCase().replace(/\s/g, '');
    if (normalizedName.includes(normalizedQuery)) {
      results.push(material);
    }
  }
  return results;
}


/**
 * @description 使用「分類-序號」組合鍵，精準查詢單一物料。
 */
function searchMaterialByCompositeKey(compositeKey) {
  const materialsMap = getInventoryMap();
  return materialsMap.get(compositeKey.toUpperCase()) || null;
}

/**
 * @description 取得特定使用者的最新操作紀錄(含原始列號)，並按時間排序。筆數由 Config 控制。
 * @param {string} userName - 使用者的 LINE 顯示名稱。
 * @returns {Array<Object>} - 回傳包含紀錄物件的陣列。
 */
function getUserRecords(userName) {
  try {
    // --- 修改處 START ---
    // 從 Config 讀取要擷取的筆數，若找不到則預設為 5
    const limit = Number(CONFIG.RECORDS_FETCH_LIMIT) || 5;
    // --- 修改處 END ---

    const sheetName = CONFIG.SHEET_NAME_RECORDS || '出入庫記錄';
    const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')).getSheetByName(sheetName);

    if (!sheet) {
      log('在試算表中找不到指定的庫存記錄分頁', { name: sheetName });
      return [];
    }

    const allData = sheet.getDataRange().getValues();
    const userRecords = [];
    
    for (let i = 1; i < allData.length; i++) {
      if (allData[i].length > 10 && allData[i][10] === userName) { 
        userRecords.push({
          index: i + 1,
          data: allData[i]
        });
      }
    }

    userRecords.sort((a, b) => {
      const timeA = a.data.length > 11 ? new Date(a.data[11]) : 0;
      const timeB = b.data.length > 11 ? new Date(b.data[11]) : 0;
      return timeB - timeA;
    });
    
    return userRecords.slice(0, limit);
  } catch (e) {
    console.error("讀取個人紀錄時發生錯誤: " + e.message);
    return [];
  }
}

/**
 * @description 產生一個新的、不重複的物料序號 (會忽略作廢紀錄)。
 * @param {string} category - 新物料的分類代號，例如 "T01"。
 * @returns {string} - 回傳一個三位數的「文字」序號，例如 "003"。
 */
function generateNewSerial(category) {
  try {
    const sheetName = CONFIG.SHEET_NAME_RECORDS || '出入庫記錄';
    const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')).getSheetByName(sheetName);
    
    if (!sheet) {
      log('在試算表中找不到指定的庫存記錄分頁', { name: sheetName });
      return "ERROR";
    }

    const allData = sheet.getDataRange().getValues();
    let maxSerial = 0;

    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      // --- 修改處 START ---
      // 檢查A欄(分類)是否相符，並且I欄(狀態)不是作廢狀態
      if (row.length > 8 && row[0] === category && row[8] !== CONFIG.STATUS_VOID) {
        const currentSerial = parseInt(row[1], 10);
        if (currentSerial > maxSerial) {
          maxSerial = currentSerial;
        }
      }
      // --- 修改處 END ---
    }
    
    const newSerial = (maxSerial + 1).toString().padStart(3, '0');
    return newSerial;

  } catch (e) {
    console.error("產生新序號時發生錯誤: " + e.message);
    return "ERROR";
  }
}


/**
 * @description 將一筆新的交易紀錄寫入資料庫，並強制序號為文字格式。
 * @param {Object} recordData - 包含交易資訊的物件。
 * @param {string} userName - 操作者的 LINE 名稱。
 */
function addTransactionRecord(recordData, userName) {
  try {
    // --- 修改處 START ---
    // 從 CONFIG 讀取分頁名稱
    const sheetName = CONFIG.SHEET_NAME_RECORDS || '出入庫記錄';
    const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')).getSheetByName(sheetName);
    // --- 修改處 END ---
    
    if (!sheet) {
      log('寫入新紀錄失敗：找不到指定的庫存記錄分頁', { name: sheetName });
      return; // 如果找不到分頁，直接結束函式
    }

    // 準備要寫入的一整列資料
    const newRow = [
      recordData.分類,
      // 在序號前加上一個單引號，強制 Google Sheets 將其視為純文字
      `'${recordData.序號}`,
      recordData.品名,
      recordData.型號 || '', // 確保跳過時寫入空白
      recordData.規格 || '', // 確保跳過時寫入空白
      recordData.單位,
      Number(recordData.數量),
      recordData.類型, // '新增', '入庫', '出庫'
      CONFIG.STATUS_VALID, // I欄: 狀態
      '', // J欄: 修改原因 (初始為空)
      userName, // K欄: 來源名稱
      new Date(), // L欄: 時間
      recordData.照片 || '' // M欄: 照片
    ];
    sheet.appendRow(newRow);
    
    // 寫入成功後，立刻清除快取
    clearInventoryCache();

  } catch (e) {
    console.error("寫入新紀錄時發生錯誤: " + e.message);
    log("寫入新紀錄時發生錯誤", e.message);
  }
}


/**
 * @description 作廢一筆指定列的紀錄，並清除快取。
 */
function voidRecordByRowIndex(rowIndex, reason, userName) {
  try {
    // --- 修改處 START ---
    // 從 CONFIG 讀取分頁名稱
    const sheetName = CONFIG.SHEET_NAME_RECORDS || '出入庫記錄';
    const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')).getSheetByName(sheetName);
    // --- 修改處 END ---
    
    if (!sheet) {
      log(`作廢紀錄失敗：找不到指定的庫存記錄分頁`, { name: sheetName, row: rowIndex });
      return; // 如果找不到分頁，直接結束函式
    }

    sheet.getRange(rowIndex, 9).setValue(CONFIG.STATUS_VOID); // I欄: 狀態
    sheet.getRange(rowIndex, 10).setValue(reason); // J欄: 修改原因
    sheet.getRange(rowIndex, 11).setValue(userName); // K欄: 來源名稱
    sheet.getRange(rowIndex, 12).setValue(new Date()); // L欄: 時間
    
    // --- 失效機制 ---
    // 寫入成功後，立刻清除快取
    clearInventoryCache();

  } catch (e) {
    console.error(`作廢第 ${rowIndex} 列紀錄時發生錯誤: ` + e.message);
    log(`作廢第 ${rowIndex} 列紀錄時發生錯誤`, e.message);
  }
}


// =================================================================
// SECTION: Result Formatting Utilities
// 說明：所有將後端資料轉換為 LINE 訊息格式的工具函式。
// =================================================================

/**
 * @description 主要的結果格式化函式 (修正排序邏輯)。
 * @param {Array<Object>} results - 搜尋結果陣列。
 * @param {string} context - 當前情境 ('query', 'inbound', 'outbound')。
 * @returns {Object} - 一個 LINE 訊息物件。
 */
function formatSearchResults(results, context = 'query') {
  if (!results || results.length === 0) {
    return { 'type': 'text', 'text': getConfigMessage('MSG_QUERY_NOT_FOUND', { dummy: '' }) };
  }

  if (results.length > 1) {
    results.sort((a, b) => {
      // --- 修改處 START ---
      // 將排序的主要關鍵字從 '類型' 改為正確的 '分類'
      const categoryA = a.分類 || ''; 
      const categoryB = b.分類 || '';
      // --- 修改處 END ---
      const serialA = a.序號 || '';
      const serialB = b.序號 || '';

      // 步驟 1: 先比較分類
      const categoryCompare = categoryA.localeCompare(categoryB);
      if (categoryCompare !== 0) {
        return categoryCompare;
      }
      
      // 步驟 2: 如果分類相同，再比較序號
      return serialA.localeCompare(serialB);
    });
  }

  if (results.length === 1) {
    return createSingleResultFlex(results[0], context);
  } else if (results.length > 1 && results.length <= 12) {
    return createCarouselFlex(results, context);
  } else { // 結果 > 12 筆
    let replyText = getConfigMessage('INFO_TOO_MANY_RESULTS_HEADER', { count: results.length });
    results.forEach(item => {
      replyText += getConfigMessage('TEMPLATE_ALL_INVENTORY_ITEM', {
        id: `${item.分類}${item.序號}`,
        name: item.品名,
        model: item.型號 || '-',
        spec: item.規格 || '-',
        stock: item.庫存,
        unit: item.單位
      });
    });
    return { 'type': 'text', 'text': replyText.trim() };
  }
}

/**
 * @description [輔助] 建立單張物料資訊的 Flex Message 卡片 (含「入庫」「出庫」按鈕)。
 * @param {Object} item - 單一物料的資訊物件。
 * @param {string} context - (保留備用) 當前情境。
 * @returns {Object} - Flex Message 物件。
 */
function createSingleResultFlex(item, context = 'query') {
  const compositeKey = `${item.分類}${item.序號}`;
  const bubble = {
    "type": "bubble",
    "hero": {
      "type": "image", "url": convertGoogleDriveFileIdToDirectUrl(item.照片), "size": "full",
      "aspectRatio": "20:13", "aspectMode": "fit", "backgroundColor": "#EEEEEE"
    },
    "body": {
      "type": "box", "layout": "vertical", "spacing": "md", "contents": [
        { "type": "text", "text": item.品名, "weight": "bold", "size": "xl", "wrap": true },
        {
          "type": "box",
          "layout": "vertical",
          "margin": "lg",
          "spacing": "sm",
          "contents": [
            { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [
                { "type": "text", "text": "【庫存】", "color": "#aaaaaa", "size": "sm", "flex": 2 },
                { "type": "text", "text": `${String(item.庫存)} ${item.單位}`, "wrap": true, "color": "#666666", "size": "md", "flex": 5, "weight": "bold" }
            ]},
            { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [
                { "type": "text", "text": "【序號】", "color": "#aaaaaa", "size": "sm", "flex": 2 },
                { "type": "text", "text": compositeKey, "wrap": true, "color": "#666666", "size": "sm", "flex": 5 }
            ]},
            { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [
                { "type": "text", "text": "【型號】", "color": "#aaaaaa", "size": "sm", "flex": 2 },
                { "type": "text", "text": String(item.型號 || '-'), "wrap": true, "color": "#666666", "size": "sm", "flex": 5 }
            ]},
            { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [
                { "type": "text", "text": "【規格】", "color": "#aaaaaa", "size": "sm", "flex": 2 },
                { "type": "text", "text": String(item.規格 || '-'), "wrap": true, "color": "#666666", "size": "sm", "flex": 5 }
            ]}
          ]
        }
      ]
    },
    // --- 程式碼修改處 START ---
    // 無論何種情境，footer 都固定為「入庫」和「出庫」兩個按鈕
    "footer": {
      "type": "box",
      "layout": "horizontal", // 水平排列
      "spacing": "sm",
      "contents": [
        {
          "type": "button",
          "style": "primary",
          "color": "#4CAF50", // 綠色代表增加
          "height": "sm",
          "action": {
            "type": "postback",
            "label": "入庫",
            "data": `stock_select&action=inbound&key=${compositeKey}`
          }
        },
        {
          "type": "button",
          "style": "primary",
          "color": "#F44336", // 紅色代表減少
          "height": "sm",
          "action": {
            "type": "postback",
            "label": "出庫",
            "data": `stock_select&action=outbound&key=${compositeKey}`
          }
        }
      ]
    }
    // --- 程式碼修改處 END ---
  };

  return {
    "type": "flex",
    "altText": `查詢結果：${item.品名}`,
    "contents": bubble
  };
}

/**
 * @description [輔助] 建立輪播 Flex Message。
 */
function createCarouselFlex(items, context = 'query') {
  const bubbles = items.map(item => createSingleResultFlex(item, context).contents);
  return { "type": "flex", "altText": `找到了 ${items.length} 筆相關結果`, "contents": { "type": "carousel", "contents": bubbles } };
}

/**
 * @description 計算並回傳所有物料的即時庫存。
 * @returns {Array<Object>} - 一個包含所有物料及其庫存的陣列。
 */
function getAllInventory() {
  const allMaterials = Array.from(getInventoryMap().values());
  return allMaterials;
}

/**
 * @description 將個人經手紀錄格式化為 Flex Message Carousel，並附帶情境式修改/刪除按鈕。
 * @param {Array<Object>} records - 來自 getUserRecords 的紀錄物件陣列 [{index: rowNum, data: [...rowData]}...]
 * @returns {Object} - Flex Message Carousel 物件或 Text Message 物件。
 */
function formatUserRecords(records) {
  if (!records || records.length === 0) {
    return { 'type': 'text', 'text': getConfigMessage('INFO_NO_RECORDS', { dummy: '' }).trim() };
  }

  const materialsMap = getInventoryMap();

  const bubbles = records.map(recordInfo => {
    const record = recordInfo.data;
    const rowIndex = recordInfo.index;

    const category = record[0];
    const serial = record[1];
    const compositeKey = `${category}${serial}`;
    const material = materialsMap.get(compositeKey);
    const photoUrl = material ? material.照片 : '';

    const recordType = String(record[7]);
    const quantity = String(Math.abs(Number(record[6])));
    const unit = String(record[5]);
    const productName = String(record[2]);
    const model = String(record[3] || '-');
    const spec = String(record[4] || '-');
    const timestamp = Utilities.formatDate(new Date(record[11]), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
    const status = record[8];

    let editButtonLabel = '';
    let editButtonData = '';
    if (recordType === '新增') {
      editButtonLabel = '修改整筆資料';
      editButtonData = `edit_start&type=new&row=${rowIndex}`;
    } else if (recordType === '入庫' || recordType === '出庫') {
      editButtonLabel = '修改數量/類型';
      editButtonData = `edit_start&type=stock&row=${rowIndex}`;
    }

    const bubble = {
      "type": "bubble",
      "hero": {
        "type": "image", "url": convertGoogleDriveFileIdToDirectUrl(photoUrl),
        "size": "full", "aspectRatio": "20:13", "aspectMode": "fit", "backgroundColor": "#EEEEEE"
      },
      "body": {
        "type": "box", "layout": "vertical", "spacing": "md",
        "contents": [
          ...(status === CONFIG.STATUS_VOID ? [{ "type": "text", "text": "⚠️ 這是一筆已作廢的舊紀錄", "color": "#FF5555", "size": "sm", "weight": "bold", "margin": "md", "wrap": true }] : []),
          { "type": "text", "text": productName, "weight": "bold", "size": "lg", "wrap": true },
          {
            "type": "box", "layout": "vertical", "margin": "md", "spacing": "sm",
            "contents": [
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [
                  { "type": "text", "text": "型號", "color": "#aaaaaa", "size": "sm", "flex": 2 },
                  { "type": "text", "text": model, "wrap": true, "color": "#666666", "size": "sm", "flex": 5 }
              ]},
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [
                  { "type": "text", "text": "規格", "color": "#aaaaaa", "size": "sm", "flex": 2 },
                  { "type": "text", "text": spec, "wrap": true, "color": "#666666", "size": "sm", "flex": 5 }
              ]},
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [
                  { "type": "text", "text": "類型", "color": "#aaaaaa", "size": "sm", "flex": 2 },
                  { "type": "text", "text": recordType, "wrap": true, "color": "#666666", "size": "sm", "flex": 5, "weight": "bold" }
              ]},
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [
                  { "type": "text", "text": "數量", "color": "#aaaaaa", "size": "sm", "flex": 2 },
                  { "type": "text", "text": `${quantity} ${unit}`, "wrap": true, "color": "#666666", "size": "sm", "flex": 5 }
              ]},
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [
                  { "type": "text", "text": "時間", "color": "#aaaaaa", "size": "sm", "flex": 2 },
                  { "type": "text", "text": timestamp, "wrap": true, "color": "#666666", "size": "sm", "flex": 5 }
              ]}
            ]
          }
        ]
      }
    };
    
    // --- 修改處 START ---
    // 只有當紀錄的狀態不是「作廢」時，才為它加上操作按鈕
    if (status !== CONFIG.STATUS_VOID) {
      const buttons = [];
      // 如果這筆紀錄可以修改，就加入修改按鈕
      if (editButtonLabel) {
        buttons.push({
          "type": "button", "style": "primary", "color": "#5E81AC", "height": "sm",
          "action": { "type": "postback", "label": editButtonLabel, "data": editButtonData }
        });
      }
      // 加入刪除按鈕
      buttons.push({
        "type": "button", "style": "primary", "color": "#ff0000ff", "height": "sm",
        "action": { "type": "postback", "label": "刪除", "data": `delete_record&row=${rowIndex}` }
      });
      
      bubble.footer = {
        "type": "box",
        "layout": "horizontal", // 改為水平排列
        "spacing": "sm",
        "contents": buttons
      };
    }
    // --- 修改處 END ---
    return bubble;
  });

  return {
    "type": "flex",
    "altText": `這是您最近的 ${records.length} 筆經手紀錄`,
    "contents": { "type": "carousel", "contents": bubbles }
  };
}

function doesMaterialExist(materialData) {
  const materialsMap = getInventoryMap(); // 取得所有現存物料的 Map
  const { 品名, 型號, 規格 } = materialData;

  // 遍歷所有現存物料
  for (const existingMaterial of materialsMap.values()) {
    if (existingMaterial.品名 === 品名 &&
        existingMaterial.型號 === 型號 &&
        existingMaterial.規格 === 規格) {
      return true; // 找到完全相符的項目
    }
  }
  return false; // 沒找到任何相符的項目
}

/**
 * @description [全新] 清除庫存快取。在任何寫入操作後都應呼叫此函式。
 */
function clearInventoryCache() {
  try {
    const cache = CacheService.getScriptCache();
    
    // --- 修改處 START ---
    // 從 CONFIG 讀取快取鑰匙，確保與 getInventoryMap 使用的鑰匙一致
    const cacheKey = CONFIG.CACHE_KEY_INVENTORY || 'inventory_map';
    cache.remove(cacheKey); // 移除指定的快取項目
    // --- 修改處 END ---
    
    log('庫存快取已清除。');
  } catch (e) {
    log('清除快取時發生錯誤', e.message);
  }
}

/**
 * @description (For Edit Flow) 通用的儲存格更新工具，用來更新特定列的指定欄位，不會清除快取或動到其他欄位。
 * @param {number} rowIndex - 要修改的紀錄在試算表中的列號。
 * @param {Object} updateObject - 一個包含要更新欄位號碼和新值的物件。例如：{10: '新的修改原因'}
 */
function updateRecordCells(rowIndex, updateObject) {
  try {
    const sheetName = CONFIG.SHEET_NAME_RECORDS || '出入庫記錄';
    const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')).getSheetByName(sheetName);
    
    if (!sheet) {
      log('更新儲存格失敗：找不到指定的庫存記錄分頁', { name: sheetName, row: rowIndex });
      return;
    }

    for (const colIndex in updateObject) {
      sheet.getRange(rowIndex, Number(colIndex)).setValue(updateObject[colIndex]);
    }
    log(`第 ${rowIndex} 列的儲存格已成功更新。`, updateObject);
  } catch (e) {
    console.error(`更新第 ${rowIndex} 列儲存格時發生錯誤: ` + e.message);
    log(`更新第 ${rowIndex} 列儲存格時發生錯誤`, { row: rowIndex, error: e.message });
  }
}

/**
 * @description 讀取 'Users' 分頁的所有使用者資料，轉換為 Map 並加入快取。
 * @returns {Map<string, string>} - 一個以 userId 為 key，displayName 為 value 的 Map。
 */
function getUsersMap() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = CONFIG.CACHE_KEY_USERS || 'users_map';

    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      return new Map(JSON.parse(cachedData));
    }

    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error("指令碼屬性中找不到 'SPREADSHEET_ID'。");
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = CONFIG.SHEET_NAME_USERS || 'Users';
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      log(`找不到使用者分頁 '${sheetName}'，將建立新分頁。`);
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['userId', 'displayName']);
      return new Map();
    }

    const allData = sheet.getDataRange().getValues();
    const usersMap = new Map();

    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const userId = row[0];
      const displayName = row[1];
      if (userId) {
        usersMap.set(userId, displayName);
      }
    }

    // --- 修改處 START ---
    // 從 Config 讀取快取時間，若找不到則預設為 3600 秒
    const expiration = Number(CONFIG.CACHE_EXPIRATION_USERS) || 3600;
    cache.put(cacheKey, JSON.stringify(Array.from(usersMap.entries())), expiration);
    // --- 修改處 END ---

    return usersMap;
  } catch (e) {
    console.error("讀取 Users 資料時發生錯誤: " + e.message);
    log("讀取 Users 資料時發生錯誤", e.message);
    return new Map();
  }
}