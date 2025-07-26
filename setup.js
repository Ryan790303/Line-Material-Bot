/**
 * @description 【完整流程】建立一個全新的圖文選單、上傳圖片、設為預設，並將 ID 存起來。
 */
function createAndSetDefaultRichMenu() {
  try {
    const imageFileId = PropertiesService.getScriptProperties().getProperty('RICH_MENU_IMAGE_ID_MAIN');
    if (!imageFileId || imageFileId === 'none') {
      throw new Error("指令碼屬性中找不到 'RICH_MENU_IMAGE_ID_MAIN'，請先設定Google Drive圖片檔案的ID。");
    }

    const richMenuJson = {
      "size": { "width": 2500, "height": 1686 }, "selected": true, "name": "MaterialBot-MainMenu-v2", "chatBarText": "點此開啟功能選單",
      "areas": [
        { "bounds": { "x": 0,    "y": 0,    "width": 625, "height": 843 }, "action": { "type": "postback", "data": "action=add" }},
        { "bounds": { "x": 625,  "y": 0,    "width": 625, "height": 843 }, "action": { "type": "postback", "data": "action=inbound" }},
        { "bounds": { "x": 1250, "y": 0,    "width": 625, "height": 843 }, "action": { "type": "postback", "data": "action=outbound" }},
        { "bounds": { "x": 1875, "y": 0,    "width": 625, "height": 843 }, "action": { "type": "postback", "data": "action=edit" }},
        { "bounds": { "x": 0,    "y": 843,  "width": 625, "height": 843 }, "action": { "type": "postback", "data": "action=query" }},
        { "bounds": { "x": 1250,  "y": 843,  "width": 625, "height": 843 }, "action": { "type": "postback", "data": "action=help" }},
        { "bounds": { "x": 1875, "y": 843,  "width": 625, "height": 843 }, "action": { "type": "postback", "data": "action=cancel" }}
      ]
    };

    const richMenuId = createRichMenu(richMenuJson);
    console.log("成功建立圖文選單，新ID:", richMenuId);
    uploadRichMenuImage(richMenuId, imageFileId);
    console.log("成功上傳圖片到圖文選單。");
    setDefaultRichMenu(richMenuId);
    console.log("成功將【新的】圖文選單設為預設。");
    PropertiesService.getScriptProperties().setProperty('RICH_MENU_ID_MAIN', richMenuId);
  } catch (e) {
    console.error("設定圖文選單時發生錯誤:", e);
  }
}

/**
 * @description 【主要使用】連結一個「已經存在」的圖文選單，將它設為所有使用者的預設選單。
 */
function linkDefaultRichMenu() {
  try {
    const richMenuId = PropertiesService.getScriptProperties().getProperty('RICH_MENU_ID_MAIN');
    
    // --- 修正處 START ---
    // 檢查 ID 是否為 null, undefined, 或是我們的預設空值 'none'
    if (!richMenuId || richMenuId === 'none') {
      throw new Error("指令碼屬性中找不到有效的 'RICH_MENU_ID_MAIN'。請先執行一次 createAndSetDefaultRichMenu 來建立選單。");
    }
    // --- 修正處 END ---

    setDefaultRichMenu(richMenuId);
    console.log(`成功將 ID 為 ${richMenuId} 的圖文選單設為預設。`);
  } catch (e) {
    console.error("連結預設圖文選單時發生錯誤:", e);
  }
}

/**
 * @description 【可選工具】解除所有使用者的預設圖文選單。
 */
function unlinkDefaultRichMenu() {
  try {
    const accessToken = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    const url = 'https://api.line.me/v2/bot/user/all/richmenu';
    const options = { 'method': 'delete', 'headers': { 'Authorization': 'Bearer ' + accessToken }};
    UrlFetchApp.fetch(url, options);
    console.log("已成功解除所有使用者的預設圖文選單。");
  } catch(e) {
    console.error("解除預設圖文選單時發生錯誤:", e);
  }
}

/**
 * @description 【管理工具】查詢並列出目前這個 LINE 官方帳號底下，所有已經建立的圖文選單。
 */
function getAllRichMenus() {
  try {
    const accessToken = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    const url = 'https://api.line.me/v2/bot/richmenu/list';
    const options = {
      'method': 'get',
      'headers': { 'Authorization': 'Bearer ' + accessToken }
    };
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    console.log(`查詢到 ${data.richmenus.length} 個圖文選單:`);
    console.log(JSON.stringify(data.richmenus, null, 2)); // 格式化輸出，方便閱讀

  } catch (e) {
    console.error("查詢所有圖文選單時發生錯誤:", e);
  }
}


/**
 * @description 【管理工具-危險】刪除這個 LINE 官方帳號底下「所有」的圖文選單。
 */
function deleteAllRichMenus() {
  console.warn("⚠️警告：即將開始刪除此帳號底下所有的 Rich Menu，此操作無法復原。");
  try {
    // 步驟 1: 先取得所有 Rich Menu 的列表
    const accessToken = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    const listUrl = 'https://api.line.me/v2/bot/richmenu/list';
    const listOptions = {
      'method': 'get',
      'headers': { 'Authorization': 'Bearer ' + accessToken }
    };
    const listResponse = UrlFetchApp.fetch(listUrl, listOptions);
    const listData = JSON.parse(listResponse.getContentText());

    if (listData.richmenus.length === 0) {
      console.log("沒有找到任何已建立的圖文選單，無需刪除。");
      return;
    }

    console.log(`找到 ${listData.richmenus.length} 個圖文選單，開始逐一刪除...`);

    // 步驟 2: 遍歷列表，逐一呼叫刪除 API
    listData.richmenus.forEach(menu => {
      console.log(`正在刪除 ID: ${menu.richMenuId}, 名稱: ${menu.name}`);
      deleteRichMenuById(menu.richMenuId); // 呼叫我們放在 Utilities.js 的工具函式
    });

    console.log("✅ 所有圖文選單已成功刪除。");

  } catch (e) {
    console.error("刪除所有圖文選單時發生錯誤:", e);
  }
}