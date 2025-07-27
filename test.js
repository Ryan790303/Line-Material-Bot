// --- 這是測試專用的函式，之後可以刪除 ---
function testSearch() {
  // 你可以換成任何你想測試的關鍵字
  const query = "螺絲"; 
  
  const results = searchMaterials(query);
  
  // JSON.stringify(..., null, 2) 是一種讓物件格式化輸出的技巧，方便在日誌中閱讀
  console.log('搜尋結果:', JSON.stringify(results, null, 2));
}
