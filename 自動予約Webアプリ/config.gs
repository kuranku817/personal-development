function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("設定"); 
  
  // シートが見つからない場合のエラーハンドリング
  if (!sheet) {
    throw new Error("シート「設定」が見つかりません。シート名を「設定」に変更してください。");
  }

  const data = sheet.getRange("A1:B5").getValues();
  
  // スプレッドシートの時刻データを安全に文字列(HH:mm)に変換する関数
  const formatTime = (val) => {
    if (val instanceof Date) {
      return Utilities.formatDate(val, "JST", "HH:mm");
    }
    return String(val); // すでに文字列の場合
  };

  const config = {
    displayStart: formatTime(data[0][1]), // B1
    displayEnd: formatTime(data[1][1]),   // B2
    buffer: Number(data[2][1]),           // B3
    slotDuration: Number(data[3][1]),     // B4
    maxContinuous: Number(data[4][1])      // B5
  };

  console.log("読み込み成功:", config);
  return config;
}