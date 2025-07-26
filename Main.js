/**
 * @description 主函式，所有 LINE 事件的入口。
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) { return ContentService.createTextOutput(JSON.stringify({'status': 'ok'})).setMimeType(ContentService.MimeType.JSON); }
    const contents = JSON.parse(e.postData.contents);
    if (contents.events.length === 0) { return ContentService.createTextOutput(JSON.stringify({ 'status': 'ok' })).setMimeType(ContentService.MimeType.JSON); }

    const event = contents.events[0];
    const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    
    // 將所有事件統一交給主處理器
    mainHandler(event, token);

  } catch (error) {
    console.error(`doPost 發生嚴重錯誤: ${error.message}\n${error.stack}`);
    log('doPost 發生嚴重錯誤', { message: error.message, stack: error.stack });
  } finally {
    return ContentService.createTextOutput(JSON.stringify({ 'status': 'ok' })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * @description 主處理器，負責所有邏輯的路由。
 */
function mainHandler(event, token) {
  const userProperties = PropertiesService.getUserProperties();
  const state = userProperties.getProperty('state');
  const eventType = event.type;
  
  if (eventType === 'postback' && event.postback.data.startsWith('action=')) {
    userProperties.deleteAllProperties();
    startFlow(event, token);
    return;
  }
  
  let flowType = null;
  if (state) {
    flowType = state.split('_')[0];
  } else if (eventType === 'postback') {
    const postbackPrefix = event.postback.data.split('_')[0];

    // --- 修改處 START (1/2) ---
    // 在合法的流程清單中，加入 'delete'
    if (['query', 'add', 'stock', 'edit', 'delete'].includes(postbackPrefix)) {
      flowType = postbackPrefix;
    }
    // --- 修改處 END (1/2) ---
  }

  switch (flowType) {
    case 'query':
      handleQueryFlow(event, token);
      break;
    case 'add':
      handleAddFlow(event, token);
      break;
    case 'stock':
      handleStockFlow(event, token);
      break;
    case 'edit':
      handleEditFlow(event, token);
      break;
    // --- 修改處 START (2/2) ---
    // 新增 'delete' 流程的轉接規則
    case 'delete':
      handleDeleteFlow(event, token); // 我們將在下一步建立這個函式
      break;
    // --- 修改處 END (2/2) ---
  }
}

/**
 * @description 根據主指令啟動對應的流程。
 */
function startFlow(event, token) {
  const action = event.postback.data.split('=')[1];
  const userId = event.source.userId;
  const replyToken = event.replyToken;
  const userProperties = PropertiesService.getUserProperties();
  let replyMessages = [];
  
  switch (action) {
    case 'query':
      const queryTypes = CONFIG.QR_QUERY_TYPES.split(',');
      const queryButtons = queryTypes.map(type => ({ type: 'action', action: { type: 'postback', label: type, data: `query_type=${encodeURIComponent(type)}` } }));
      replyMessages.push({ 'type': 'text', 'text': getConfigMessage('PROMPT_QUERY_TYPE'), 'quickReply': { 'items': queryButtons } });
      break;
    case 'add':
      const categoriesRaw = CONFIG.QR_CATEGORIES;
      if (!categoriesRaw) {
        replyMessages.push({ 'type': 'text', 'text': '系統錯誤：找不到分類設定。' });
        break;
      }
      userProperties.setProperties({
        'state': 'add_awaiting_category',
        'temp_data': JSON.stringify({})
      });
      const categories = categoriesRaw.split(',');
      const categoryButtons = categories.map(cat => ({ type: 'action', action: { type: 'postback', label: cat, data: `add_category=${cat.split(':')[0]}` } }));
      replyMessages.push({ 'type': 'text', 'text': getConfigMessage('PROMPT_ADD_CATEGORY'), 'quickReply': { 'items': categoryButtons } });
      break;

    case 'inbound':
    case 'outbound': {
      userProperties.setProperty('state', `stock_${action}_awaiting_search_type`);
      
      const searchTypes = CONFIG.QR_STOCK_SEARCH_TYPES.split(',');
      const searchTypeButtons = searchTypes.map(type => {
        const searchMethod = type === '用品名查詢' ? 'by_name' : 'by_serial';
        return { type: 'action', action: { type: 'postback', label: type, data: `stock_search_type=${searchMethod}` }};
      });
      
      // --- 修改處 START ---
      // 將「取消」按鈕的文字，也改成從 Config 讀取
      const cancelLabel = CONFIG.LABEL_CANCEL || '取消';
      searchTypeButtons.push({ type: 'action', action: { type: 'postback', label: cancelLabel, data: 'action=cancel' }});
      // --- 修改處 END ---

      const actionText = action === 'inbound' ? '入庫' : '出庫';
      replyMessages.push({ 'type': 'text', 'text': getConfigMessage('PROMPT_STOCK_SEARCH', { action: actionText }), 'quickReply': { 'items': searchTypeButtons } });
      break;
    }

    case 'edit': {
      const userName = JSON.parse(getUserProfile(userId, token)).displayName;
      const records = getUserRecords(userName);
      replyMessages.push(formatUserRecords(records));
      break;
    }

    case 'help':
      replyMessages.push({ 'type': 'text', 'text': getConfigMessage('MSG_HELP') });
      break;
    case 'cancel':
      replyMessages.push({ 'type': 'text', 'text': getConfigMessage('MSG_CANCEL_CONFIRM') });
      break;
    default:
      if (action) {
        replyMessages.push({ 'type': 'text', 'text': getConfigMessage('INFO_WIP', { action: action }) });
      }
      break;
  }
  if (replyMessages.length > 0) {
    replyMessage(replyToken, replyMessages, token);
  }
}

/**
 * @description 處理所有「查詢」相關的對話流程 (移除長列表後的取消按鈕)
 */
function handleQueryFlow(event, token) {
  const userProperties = PropertiesService.getUserProperties();
  const state = userProperties.getProperty('state');
  const userId = event.source.userId;
  const replyToken = event.replyToken;
  const data = event.type === 'postback' ? event.postback.data : event.message.text;

  if (!state && event.type === 'postback') {
    const type = decodeURIComponent(data.split('=')[1]);
    let nextState = null;
    let replyMessages = [];

    if (type === '用品名查詢') {
      nextState = 'query_awaiting_name';
      replyMessages.push({'type': 'text', 'text': getConfigMessage('PROMPT_QUERY_BY_NAME')});
    } else if (type === '用序號查詢') {
      nextState = 'query_awaiting_serial';
      replyMessages.push({'type': 'text', 'text': getConfigMessage('PROMPT_QUERY_BY_SERIAL')});
    } else if (type === '查詢所有庫存') {
      const allInventory = getAllInventory();
      replyMessages.push(formatSearchResults(allInventory, 'query'));
    } else if (type === '查我的經手紀錄') {
      const userName = JSON.parse(getUserProfile(userId, token)).displayName;
      const records = getUserRecords(userName);
      replyMessages.push(formatUserRecords(records));
    }
    
    if (nextState) userProperties.setProperty('state', nextState);
    if (replyMessages.length > 0) replyMessage(replyToken, replyMessages, token);

  } else if (state && event.type === 'message') {
    let replyMessageObject;
    if (state === 'query_awaiting_name') {
      const searchResults = searchMaterials(data);
      replyMessageObject = formatSearchResults(searchResults, 'query');
      
      // --- 修改處 START ---
      // 移除原有的 if (searchResults.length > 12) 判斷區塊。
      // 現在無論結果有幾筆，查詢完畢後都直接清除狀態，結束這次對話。
      userProperties.deleteAllProperties();
      // --- 修改處 END ---
    } else if (state === 'query_awaiting_serial') {
      userProperties.deleteAllProperties();
      const result = searchMaterialByCompositeKey(data);
      replyMessageObject = result ? createSingleResultFlex(result, 'query') : { 'type': 'text', 'text': getConfigMessage('MSG_QUERY_NOT_FOUND') };
    }
    replyMessage(replyToken, [replyMessageObject], token);
  }
}

/**
 * @description (最終修正版) 處理所有「修改」相關的對話流程。
 */
function handleEditFlow(event, token) {
  const userProperties = PropertiesService.getUserProperties();
  const state = userProperties.getProperty('state');
  const eventType = event.type;
  const userId = event.source.userId;
  const replyToken = event.replyToken;
  const data = eventType === 'postback' ? event.postback.data : (event.message ? event.message.text : '');
  
  let tempData = JSON.parse(userProperties.getProperty('temp_data') || '{}');
  let replyMessages = [];
  let nextState = null;

  // --- 內部輔助函式區 ---
  function sendNewItemEditMenu(leadingText = '') {
    const confirmText = getConfigMessage('PROMPT_NEW_ITEM_CHOICE', {
      leadingText: leadingText, name: tempData.newData.品名, model: tempData.newData.型號 || '-',
      spec: tempData.newData.規格 || '-', unit: tempData.newData.單位, quantity: tempData.newData.數量
    });
    const fieldButtons = [
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_NAME || '修改品名', data: 'edit_field=品名' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_MODEL || '修改型號', data: 'edit_field=型號' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_SPEC || '修改規格', data: 'edit_field=規格' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_UNIT || '修改單位', data: 'edit_field=單位' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_QUANTITY || '修改數量', data: 'edit_field=數量' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_FINISH_EDIT || '✅ 完成修改，儲存', data: 'edit_field=finish' }}
    ];
    replyMessages.push({ type: 'text', text: confirmText, quickReply: { items: fieldButtons } });
  }
  
  function sendStockEditMenu(leadingText = '') {
    const promptText = getConfigMessage('PROMPT_EDIT_STOCK_CHOICE', {
      leadingText: leadingText, name: tempData.newData.品名,
      type: tempData.newData.類型, quantity: tempData.newData.數量
    });
    const choiceButtons = [
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_QUANTITY || '修改數量', data: 'edit_stock_choice=quantity' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_TYPE || '修改類型', data: 'edit_stock_choice=type' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_FINISH_EDIT || '✅ 完成修改，儲存', data: 'edit_stock_choice=finish' }}
    ];
    replyMessages.push({ type: 'text', text: promptText, quickReply: { items: choiceButtons }});
  }
  
  function finalizeEdit() {
    const userName = JSON.parse(getUserProfile(userId, token)).displayName;
    const finalData = tempData.newData;
    const originalRecord = tempData.originalRecord;
    const rowIndex = tempData.rowIndex;

    if (finalData.類型 === '出庫') {
      const compositeKey = `${finalData.分類}${finalData.序號}`;
      const currentStock = (searchMaterialByCompositeKey(compositeKey) || {}).庫存 || 0;
      const originalEffect = Number(originalRecord[6]);
      const stockAfterVoid = currentStock - originalEffect;
      
      if (stockAfterVoid < finalData.數量) {
        replyMessages.push({ type: 'text', text: getConfigMessage('MSG_STOCK_INSUFFICIENT', {name: finalData.品名, currentStock: stockAfterVoid, unit: finalData.單位}) });
        nextState = 'edit_stock_awaiting_choice';
        return;
      }
    }

    let reason = '使用者修改(未變更內容)';
    if (originalRecord[7] === '新增') {
      const original = {品名:originalRecord[2], 型號:originalRecord[3], 規格:originalRecord[4], 單位:originalRecord[5], 數量:originalRecord[6]};
      const changedFields = Object.keys(original).filter(key => String(original[key] || '') !== String(finalData[key] || ''));
      if (changedFields.length > 0) reason = `因【${changedFields.join('、')}】錯誤修改`;
    } else {
      const changedFields = [];
      const originalQty = Math.abs(Number(originalRecord[6]));
      const originalType = originalRecord[7];
      if (originalQty !== finalData.數量) changedFields.push('數量');
      if (originalType !== finalData.類型) changedFields.push('類型');
      if (changedFields.length > 0) reason = `因【${changedFields.join('、')}】錯誤修改`;
    }

    const recordToSave = { ...finalData };
    if (recordToSave.類型 === '出庫') {
      recordToSave.數量 = -Math.abs(recordToSave.數量);
    } else {
      recordToSave.數量 = Math.abs(recordToSave.數量);
    }

    voidRecordByRowIndex(rowIndex, `由 ${userName} 修改`, userName);
    updateRecordCells(rowIndex, { 10: reason });
    addTransactionRecord(recordToSave, userName);
    
    // --- 修正處 1 START ---
    // 明確使用 MSG_EDIT_SUCCESS_MODIFY 作為成功訊息
    replyMessages.push({ type: 'text', text: getConfigMessage('MSG_EDIT_SUCCESS_MODIFY') });
    // --- 修正處 1 END ---
    userProperties.deleteAllProperties();
  }

  // --- 流程起點 (唯讀模式) ---
  if (!state && eventType === 'postback' && data.startsWith('edit_start')) {
    const params = data.split('&').reduce((acc, part) => {
      const [key, value] = part.split('=');
      if (value !== undefined) { acc[key] = decodeURIComponent(value); }
      return acc;
    }, {});
    
    const editType = params.type;
    const rowIndex = Number(params.row);
    const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')).getSheetByName(CONFIG.SHEET_NAME_RECORDS || '出入庫記錄');
    const recordData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    tempData = {
      originalRecord: recordData,
      rowIndex: rowIndex,
      newData: {
        分類: recordData[0], 序號: recordData[1], 品名: recordData[2], 型號: recordData[3],
        規格: recordData[4], 單位: recordData[5], 數量: Math.abs(recordData[6]), 類型: recordData[7], 照片: recordData[12]
      }
    };

    if (editType === 'new') {
      nextState = 'edit_new_awaiting_choice';
    } else if (editType === 'stock') {
      nextState = 'edit_stock_awaiting_choice';
    }
  }
  // --- 狀態機 ---
  else if (state) {
    // --- 「修改出入庫」的循環選單 ---
    if (state === 'edit_stock_awaiting_choice') {
      const choice = data.split('=')[1];
      if (choice === 'finish') {
        finalizeEdit();
      } else if (choice === 'quantity') {
        nextState = 'edit_stock_awaiting_quantity';
        replyMessages.push({ type: 'text', text: getConfigMessage('PROMPT_EDIT_NEW_VALUE', {field: '數量'}) });
      } else if (choice === 'type') {
        nextState = 'edit_stock_awaiting_type';
        const typeButtons = [
          { type: 'action', action: { type: 'postback', label: '入庫', data: 'edit_type=入庫' }},
          { type: 'action', action: { type: 'postback', label: '出庫', data: 'edit_type=出庫' }}
        ];
        replyMessages.push({ type: 'text', text: '請選擇新的紀錄類型：', quickReply: { items: typeButtons } });
      }
    }
    else if (state === 'edit_stock_awaiting_quantity') {
      if (isNaN(data) || Number(data) < 0) {
        replyMessages.push({ 'type': 'text', text: getConfigMessage('ERROR_INVALID_QUANTITY') });
        nextState = state;
      } else {
        tempData.newData.數量 = Number(data);
        nextState = 'edit_stock_awaiting_choice';
      }
    }
    else if (state === 'edit_stock_awaiting_type') {
      tempData.newData.類型 = data.split('=')[1];
      nextState = 'edit_stock_awaiting_choice';
    }

    // --- 「修改新增物料」的循環選單 ---
    else if (state === 'edit_new_awaiting_choice') {
      const choice = data.split('=')[1];
      if (choice === 'finish') {
        finalizeEdit();
      } else if (choice === '單位') {
        // --- 修正處 2 START ---
        nextState = 'edit_new_awaiting_unit_choice';
        const units = (CONFIG.QR_UNITS || '').split(',').filter(u => u.trim() !== '手動輸入' && u.trim() !== '');
        const unitButtons = units.map(unit => ({ type: 'action', action: { type: 'postback', label: unit, data: `edit_unit=${encodeURIComponent(unit)}` } }));
        // 將「手動輸入」按鈕固定加上去
        unitButtons.push({ type: 'action', action: { type: 'postback', label: '手動輸入', data: `edit_unit=${encodeURIComponent('手動輸入')}` } });
        // --- 修正處 2 END ---
        
        // --- 修正處 3 START ---
        // 使用更通用的提問
        replyMessages.push({ type: 'text', text: getConfigMessage('PROMPT_EDIT_SELECT_FIELD', {field: '單位'}), quickReply: { items: unitButtons } });
        // --- 修正處 3 END ---
      } else {
        tempData.fieldToEdit = choice;
        nextState = 'edit_new_awaiting_new_value';
        replyMessages.push({ type: 'text', text: getConfigMessage('PROMPT_EDIT_NEW_VALUE', {field: choice}) });
      }
    }
    else if (state === 'edit_new_awaiting_unit_choice') {
      const value = decodeURIComponent(data.split('=')[1]);
      if (value === '手動輸入') {
        nextState = 'edit_new_awaiting_manual_unit';
        replyMessages.push({ type: 'text', text: getConfigMessage('PROMPT_MANUAL_UNIT') });
      } else {
        tempData.newData['單位'] = value;
        nextState = 'edit_new_awaiting_choice';
      }
    }
    else if (state === 'edit_new_awaiting_manual_unit') {
      tempData.newData['單位'] = data;
      nextState = 'edit_new_awaiting_choice';
    }
    else if (state === 'edit_new_awaiting_new_value') {
      const fieldToEdit = tempData.fieldToEdit;
      tempData.newData[fieldToEdit] = data;
      delete tempData.fieldToEdit;
      nextState = 'edit_new_awaiting_choice';
    }
  }

  // --- 根據下一個狀態，決定是否要發送主選單訊息 ---
  if (nextState === 'edit_new_awaiting_choice') {
    const leadingText = tempData.fieldToEdit ? `「${tempData.fieldToEdit}」已更新。\n\n` : '';
    sendNewItemEditMenu(leadingText);
  } else if (nextState === 'edit_stock_awaiting_choice') {
    sendStockEditMenu('資料已更新。\n');
  }
  
  if (nextState) {
    userProperties.setProperty('state', nextState);
    userProperties.setProperty('temp_data', JSON.stringify(tempData));
  }
  if (replyMessages.length > 0) {
    replyMessage(replyToken, replyMessages, token);
  }
}

/**
 * @description 處理所有「入庫/出庫」相關的對話流程
 */
function handleStockFlow(event, token) {
  const userProperties = PropertiesService.getUserProperties();
  const state = userProperties.getProperty('state');
  const userId = event.source.userId;
  const eventType = event.type;
  const replyToken = event.replyToken;
  const data = eventType === 'postback' ? event.postback.data : event.message.text;
  
  let tempData = JSON.parse(userProperties.getProperty('temp_data') || '{}');
  let replyMessages = [];
  let nextState = null;

  // 從 state 中解析出當前的核心動作 (inbound/outbound)
  const action = state ? state.split('_')[1] : null;

  // --- 處理 Postback 指令 ---
  if (eventType === 'postback') {
    const params = data.split('&').reduce((acc, part) => { const [k, v] = part.split('='); acc[k] = decodeURIComponent(v || ''); return acc; }, {});

    if (params.stock_search_type) { // 步驟 2: 使用者選擇了搜尋方式
      nextState = `stock_${action}_awaiting_${params.stock_search_type}_search`;
      const promptKey = params.stock_search_type === 'by_name' ? 'PROMPT_QUERY_BY_NAME' : 'PROMPT_QUERY_BY_SERIAL';
      replyMessages.push({ 'type': 'text', 'text': getConfigMessage(promptKey, { dummy: '' }) });
    } 
    // --- 修正處 START ---
    // 步驟 4: 使用者從搜尋結果卡片上選擇了物料 (捷徑功能)
    else if ('stock_select' in params) {
      const selectedKey = params.key;
      const selectedAction = params.action;
      const selectedMaterial = searchMaterialByCompositeKey(selectedKey);
      if (selectedMaterial) {
        tempData.selectedItem = selectedMaterial;
        tempData.action = selectedAction; // 從按鈕指令中確認操作類型
        nextState = `stock_${selectedAction}_awaiting_quantity`;
        const actionText = selectedAction === 'inbound' ? '入庫' : '出庫';
        replyMessages.push({ 'type': 'text', 'text': getConfigMessage('PROMPT_STOCK_QUANTITY', { name: selectedMaterial.品名, action: actionText }) });
      } else {
        replyMessages.push({ 'type': 'text', 'text': getConfigMessage('MSG_QUERY_NOT_FOUND', { dummy: '' }) });
        userProperties.deleteAllProperties();
      }
    } 
    // --- 修正處 END ---
    else if (params.stock_confirm) { // 步驟 7: 使用者最終確認入庫/出庫
      if (params.stock_confirm === '確認') {
        const item = tempData.selectedItem;
        const record = {
          分類: item.分類, 序號: item.序號, 品名: item.品名, 型號: item.型號,
          規格: item.規格, 單位: item.單位, 照片: item.照片,
          數量: tempData.action === 'inbound' ? tempData.quantity : -tempData.quantity,
          類型: tempData.action === 'inbound' ? '入庫' : '出庫',
        };
        const userName = JSON.parse(getUserProfile(userId, token)).displayName;
        addTransactionRecord(record, userName);
        
        const newStock = searchMaterialByCompositeKey(`${item.分類}${item.序號}`).庫存;
        const successIcon = tempData.action === 'inbound' ? '✅' : '➡️';
        replyMessages.push({ 'type': 'text', 'text': getConfigMessage('MSG_STOCK_SUCCESS', { icon: successIcon, action: record.類型, name: item.品名, newStock: newStock, unit: item.單位 }) });
      } else {
        replyMessages.push({ 'type': 'text', 'text': getConfigMessage('MSG_CANCEL_CONFIRM', { dummy: '' }) });
      }
      userProperties.deleteAllProperties();
    }
  } 
  // --- 處理文字輸入 ---
  else if (eventType === 'message' && state) {
    // 步驟 3: 使用者輸入了搜尋關鍵字
    if (state.endsWith('_awaiting_by_name_search') || state.endsWith('_awaiting_by_serial_search')) {
      const searchResults = state.includes('_by_name_') ? searchMaterials(data) : [searchMaterialByCompositeKey(data)].filter(Boolean);
      const message = formatSearchResults(searchResults, action); 
      replyMessages.push(message);
      userProperties.deleteProperty('state'); // 結束搜尋狀態，等待卡片按鈕的回應
    } 
    // 步驟 5: 使用者輸入了數量
    else if (state.endsWith('_awaiting_quantity')) {
      if (isNaN(data) || data.trim() === '' || Number(data) <= 0) {
        replyMessages.push({ 'type': 'text', 'text': getConfigMessage('ERROR_INVALID_QUANTITY', { dummy: '' }) });
        nextState = state; // 狀態不變，讓使用者重輸
      } else {
        const quantity = Number(data);
        const item = tempData.selectedItem;
        if (action === 'outbound' && item.庫存 < quantity) {
          replyMessages.push({ 'type': 'text', 'text': getConfigMessage('MSG_STOCK_INSUFFICIENT', { name: item.品名, currentStock: item.庫存, unit: item.單位 }) });
          userProperties.deleteAllProperties();
        } else {
          tempData.quantity = quantity;
          nextState = `stock_${action}_awaiting_confirmation`;
          const actionText = action === 'inbound' ? '入庫' : '出庫';
          const confirmText = getConfigMessage('PROMPT_STOCK_CONFIRM_PROMPT', { action: actionText, name: item.品名, quantity: quantity, unit: item.單位 });
          const confirmButtons = [
            { type: 'action', action: { type: 'postback', label: `確認${actionText}`, data: 'stock_confirm=確認' }},
            { type: 'action', action: { type: 'postback', label: '取消', data: 'stock_confirm=取消' }}
          ];
          replyMessages.push({ 'type': 'text', 'text': confirmText, 'quickReply': { 'items': confirmButtons } });
        }
      }
    }
  }

  if (nextState) {
    userProperties.setProperty('state', nextState);
    userProperties.setProperty('temp_data', JSON.stringify(tempData));
  }
  if (replyMessages.length > 0) {
    replyMessage(replyToken, replyMessages, token);
  }
}

/**
 * @description (最終修正版) 處理所有「修改」相關的對話流程。
 */
function handleEditFlow(event, token) {
  const userProperties = PropertiesService.getUserProperties();
  const state = userProperties.getProperty('state');
  const eventType = event.type;
  const userId = event.source.userId;
  const replyToken = event.replyToken;
  const data = eventType === 'postback' ? event.postback.data : (event.message ? event.message.text : '');
  
  let tempData = JSON.parse(userProperties.getProperty('temp_data') || '{}');
  let replyMessages = [];
  let nextState = null;

  // --- 內部輔助函式區 ---
  function sendNewItemEditMenu(leadingText = '') {
    const confirmText = getConfigMessage('PROMPT_NEW_ITEM_CHOICE', {
      leadingText: leadingText, name: tempData.newData.品名, model: tempData.newData.型號 || '-',
      spec: tempData.newData.規格 || '-', unit: tempData.newData.單位, quantity: tempData.newData.數量
    });
    const fieldButtons = [
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_NAME || '修改品名', data: 'edit_field=品名' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_MODEL || '修改型號', data: 'edit_field=型號' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_SPEC || '修改規格', data: 'edit_field=規格' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_UNIT || '修改單位', data: 'edit_field=單位' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_QUANTITY || '修改數量', data: 'edit_field=數量' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_FINISH_EDIT || '✅ 完成修改，儲存', data: 'edit_field=finish' }}
    ];
    replyMessages.push({ type: 'text', text: confirmText, quickReply: { items: fieldButtons } });
  }
  
  function sendStockEditMenu(leadingText = '') {
    const promptText = getConfigMessage('PROMPT_EDIT_STOCK_CHOICE', {
      leadingText: leadingText, name: tempData.newData.品名,
      type: tempData.newData.類型, quantity: tempData.newData.數量
    });
    const choiceButtons = [
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_QUANTITY || '修改數量', data: 'edit_stock_choice=quantity' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_EDIT_TYPE || '修改類型', data: 'edit_stock_choice=type' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_FINISH_EDIT || '✅ 完成修改，儲存', data: 'edit_stock_choice=finish' }}
    ];
    replyMessages.push({ type: 'text', text: promptText, quickReply: { items: choiceButtons }});
  }
  
  function finalizeEdit() {
    const userName = JSON.parse(getUserProfile(userId, token)).displayName;
    const finalData = tempData.newData;
    const originalRecord = tempData.originalRecord;
    const rowIndex = tempData.rowIndex;

    if (finalData.類型 === '出庫') {
      const compositeKey = `${finalData.分類}${finalData.序號}`;
      const currentStock = (searchMaterialByCompositeKey(compositeKey) || {}).庫存 || 0;
      const originalEffect = Number(originalRecord[6]);
      const stockAfterVoid = currentStock - originalEffect;
      
      if (stockAfterVoid < finalData.數量) {
        replyMessages.push({ type: 'text', text: getConfigMessage('MSG_STOCK_INSUFFICIENT', {name: finalData.品名, currentStock: stockAfterVoid, unit: finalData.單位}) });
        nextState = 'edit_stock_awaiting_choice';
        return;
      }
    }

    let reason = '使用者修改(未變更內容)';
    if (originalRecord[7] === '新增') {
      const original = {品名:originalRecord[2], 型號:originalRecord[3], 規格:originalRecord[4], 單位:originalRecord[5], 數量:originalRecord[6]};
      const changedFields = Object.keys(original).filter(key => String(original[key] || '') !== String(finalData[key] || ''));
      if (changedFields.length > 0) reason = `因【${changedFields.join('、')}】錯誤修改`;
    } else {
      const changedFields = [];
      const originalQty = Math.abs(Number(originalRecord[6]));
      const originalType = originalRecord[7];
      if (originalQty !== finalData.數量) changedFields.push('數量');
      if (originalType !== finalData.類型) changedFields.push('類型');
      if (changedFields.length > 0) reason = `因【${changedFields.join('、')}】錯誤修改`;
    }

    const recordToSave = { ...finalData };
    if (recordToSave.類型 === '出庫') {
      recordToSave.數量 = -Math.abs(recordToSave.數量);
    } else {
      recordToSave.數量 = Math.abs(recordToSave.數量);
    }

    voidRecordByRowIndex(rowIndex, `由 ${userName} 修改`, userName);
    updateRecordCells(rowIndex, { 10: reason });
    addTransactionRecord(recordToSave, userName);
    
    // --- 修正處 1 START ---
    // 明確使用 MSG_EDIT_SUCCESS_MODIFY 作為成功訊息
    replyMessages.push({ type: 'text', text: getConfigMessage('MSG_EDIT_SUCCESS_MODIFY') });
    // --- 修正處 1 END ---
    userProperties.deleteAllProperties();
  }

  // --- 流程起點 (唯讀模式) ---
  if (!state && eventType === 'postback' && data.startsWith('edit_start')) {
    const params = data.split('&').reduce((acc, part) => {
      const [key, value] = part.split('=');
      if (value !== undefined) { acc[key] = decodeURIComponent(value); }
      return acc;
    }, {});
    
    const editType = params.type;
    const rowIndex = Number(params.row);
    const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')).getSheetByName(CONFIG.SHEET_NAME_RECORDS || '出入庫記錄');
    const recordData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    tempData = {
      originalRecord: recordData,
      rowIndex: rowIndex,
      newData: {
        分類: recordData[0], 序號: recordData[1], 品名: recordData[2], 型號: recordData[3],
        規格: recordData[4], 單位: recordData[5], 數量: Math.abs(recordData[6]), 類型: recordData[7], 照片: recordData[12]
      }
    };

    if (editType === 'new') {
      nextState = 'edit_new_awaiting_choice';
    } else if (editType === 'stock') {
      nextState = 'edit_stock_awaiting_choice';
    }
  }
  // --- 狀態機 ---
  else if (state) {
    // --- 「修改出入庫」的循環選單 ---
    if (state === 'edit_stock_awaiting_choice') {
      const choice = data.split('=')[1];
      if (choice === 'finish') {
        finalizeEdit();
      } else if (choice === 'quantity') {
        nextState = 'edit_stock_awaiting_quantity';
        replyMessages.push({ type: 'text', text: getConfigMessage('PROMPT_EDIT_NEW_VALUE', {field: '數量'}) });
      } else if (choice === 'type') {
        nextState = 'edit_stock_awaiting_type';
        const typeButtons = [
          { type: 'action', action: { type: 'postback', label: '入庫', data: 'edit_type=入庫' }},
          { type: 'action', action: { type: 'postback', label: '出庫', data: 'edit_type=出庫' }}
        ];
        replyMessages.push({ type: 'text', text: '請選擇新的紀錄類型：', quickReply: { items: typeButtons } });
      }
    }
    else if (state === 'edit_stock_awaiting_quantity') {
      if (isNaN(data) || Number(data) < 0) {
        replyMessages.push({ 'type': 'text', text: getConfigMessage('ERROR_INVALID_QUANTITY') });
        nextState = state;
      } else {
        tempData.newData.數量 = Number(data);
        nextState = 'edit_stock_awaiting_choice';
      }
    }
    else if (state === 'edit_stock_awaiting_type') {
      tempData.newData.類型 = data.split('=')[1];
      nextState = 'edit_stock_awaiting_choice';
    }

    // --- 「修改新增物料」的循環選單 ---
    else if (state === 'edit_new_awaiting_choice') {
      const choice = data.split('=')[1];
      if (choice === 'finish') {
        finalizeEdit();
      } else if (choice === '單位') {
        // --- 修正處 2 START ---
        nextState = 'edit_new_awaiting_unit_choice';
        const units = (CONFIG.QR_UNITS || '').split(',').filter(u => u.trim() !== '手動輸入' && u.trim() !== '');
        const unitButtons = units.map(unit => ({ type: 'action', action: { type: 'postback', label: unit, data: `edit_unit=${encodeURIComponent(unit)}` } }));
        // 將「手動輸入」按鈕固定加上去
        unitButtons.push({ type: 'action', action: { type: 'postback', label: '手動輸入', data: `edit_unit=${encodeURIComponent('手動輸入')}` } });
        // --- 修正處 2 END ---
        
        // --- 修正處 3 START ---
        // 使用更通用的提問
        replyMessages.push({ type: 'text', text: getConfigMessage('PROMPT_EDIT_SELECT_FIELD', {field: '單位'}), quickReply: { items: unitButtons } });
        // --- 修正處 3 END ---
      } else {
        tempData.fieldToEdit = choice;
        nextState = 'edit_new_awaiting_new_value';
        replyMessages.push({ type: 'text', text: getConfigMessage('PROMPT_EDIT_NEW_VALUE', {field: choice}) });
      }
    }
    else if (state === 'edit_new_awaiting_unit_choice') {
      const value = decodeURIComponent(data.split('=')[1]);
      if (value === '手動輸入') {
        nextState = 'edit_new_awaiting_manual_unit';
        replyMessages.push({ type: 'text', text: getConfigMessage('PROMPT_MANUAL_UNIT') });
      } else {
        tempData.newData['單位'] = value;
        nextState = 'edit_new_awaiting_choice';
      }
    }
    else if (state === 'edit_new_awaiting_manual_unit') {
      tempData.newData['單位'] = data;
      nextState = 'edit_new_awaiting_choice';
    }
    else if (state === 'edit_new_awaiting_new_value') {
      const fieldToEdit = tempData.fieldToEdit;
      tempData.newData[fieldToEdit] = data;
      delete tempData.fieldToEdit;
      nextState = 'edit_new_awaiting_choice';
    }
  }

  // --- 根據下一個狀態，決定是否要發送主選單訊息 ---
  if (nextState === 'edit_new_awaiting_choice') {
    const leadingText = tempData.fieldToEdit ? `「${tempData.fieldToEdit}」已更新。\n\n` : '';
    sendNewItemEditMenu(leadingText);
  } else if (nextState === 'edit_stock_awaiting_choice') {
    sendStockEditMenu('資料已更新。\n');
  }
  
  if (nextState) {
    userProperties.setProperty('state', nextState);
    userProperties.setProperty('temp_data', JSON.stringify(tempData));
  }
  if (replyMessages.length > 0) {
    replyMessage(replyToken, replyMessages, token);
  }
}

/**
 * @description (全新) 處理所有「刪除」相關的對話流程。
 */
function handleDeleteFlow(event, token) {
  const userProperties = PropertiesService.getUserProperties();
  const state = userProperties.getProperty('state');
  const eventType = event.type;
  const userId = event.source.userId;
  const replyToken = event.replyToken;
  const data = eventType === 'postback' ? event.postback.data : '';
  
  let tempData = JSON.parse(userProperties.getProperty('temp_data') || '{}');
  let replyMessages = [];
  let nextState = null;

  // 步驟 1: 收到刪除請求，發送二次確認
  if (!state && eventType === 'postback' && data.startsWith('delete_record')) {
    const params = data.split('&').reduce((acc, part) => {
      const [key, value] = part.split('=');
      if (value !== undefined) { acc[key] = decodeURIComponent(value); }
      return acc;
    }, {});
    
    const rowIndex = Number(params.row);
    
    const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')).getSheetByName(CONFIG.SHEET_NAME_RECORDS || '出入庫記錄');
    const recordData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const recordName = recordData[2];
    const recordType = recordData[7];

    tempData.rowIndex = rowIndex;
    nextState = 'delete_awaiting_confirmation';
    
    // --- 修改處 START ---
    const confirmText = getConfigMessage('PROMPT_DELETE_CONFIRM', { name: recordName, type: recordType });
    const confirmButtons = [
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_CONFIRM_DELETE || '⚠️ 確認刪除', data: 'delete_confirm=yes' }},
      { type: 'action', action: { type: 'postback', label: CONFIG.LABEL_CANCEL || '取消', data: 'delete_confirm=no' }}
    ];
    replyMessages.push({ type: 'text', text: confirmText, quickReply: { items: confirmButtons }});
    // --- 修改處 END ---
  } 
  // 步驟 2: 處理使用者的確認結果
  else if (state === 'delete_awaiting_confirmation') {
    const choice = data.split('=')[1];
    
    if (choice === 'yes') {
      const rowIndex = tempData.rowIndex;
      const userName = JSON.parse(getUserProfile(userId, token)).displayName;
      // --- 修改處 START ---
      const reason = CONFIG.DEFAULT_DELETE_REASON || '資料錯誤';
      voidRecordByRowIndex(rowIndex, reason, userName);
      replyMessages.push({ type: 'text', text: getConfigMessage('MSG_DELETE_SUCCESS') });
      // --- 修改處 END ---
    } else {
      replyMessages.push({ type: 'text', text: getConfigMessage('MSG_CANCEL_CONFIRM') });
    }
    
    userProperties.deleteAllProperties();
  }

  if (nextState) {
    userProperties.setProperty('state', nextState);
    userProperties.setProperty('temp_data', JSON.stringify(tempData));
  }
  if (replyMessages.length > 0) {
    replyMessage(replyToken, replyMessages, token);
  }
}