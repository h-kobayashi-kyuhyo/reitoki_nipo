const SS_ID = PropertiesService.getScriptProperties().getProperty('Spsheet_ID');

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('冷凍機械 運転日報')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 初期データ取得（マスタと担当者）
function getInitialData() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    
    const masterSheet = ss.getSheetByName('マスタ');
    if(!masterSheet) return {master:[], staff:[], error:"マスタシートなし"};
    const masterData = masterSheet.getDataRange().getDisplayValues();
    if (masterData.length > 0) masterData.shift();

    const staffSheet = ss.getSheetByName('担当者マスタ');
    const staffData = staffSheet ? staffSheet.getDataRange().getDisplayValues() : [];
    if (staffData.length > 0) staffData.shift();
    
    return { master: masterData, staff: staffData };
  } catch (e) {
    return { master: [], staff: [], error: e.message };
  }
}

// ログイン判定
function checkLogin(id, password) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('事業所マスタ');
  const data = sheet.getDataRange().getDisplayValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id && String(data[i][2]) === password) {
      return { success: true, siteName: data[i][1] };
    }
  }
  return { success: false };
}

// 事業所リスト取得
function getSiteList() {
 try{ 
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('事業所マスタ');
  const data = sheet.getDataRange().getDisplayValues();
  data.shift();
  console.log("事業所リスト取得数:"+data.length);
  return data.map(row => ({ id:String(row[0]),name:String(row[1])}));
}catch(e){
  console.error("getSiteListエラー:"+e.message);
  return[];
  }
}

// データ保存
function saveData(results) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('データ保存');
    const now = new Date();
    
    results.forEach(res => {
      sheet.appendRow([
        res.date,      // A: 点検日
        res.site,      // B: 事業所
        res.staff,     // C: 担当者
        res.weather,   // D: 天候
        res.building,  // E: 棟（マスタ2列目）
        res.area,      // F: エリア項目（マスタ3列目）
        "'"+res.no,        // G: 設備名
        res.item,      // H: 項目
        res.value,     // I: 値
        now            // J: 書込時間
      ]);
    });
    return "送信が完了しました";
  } catch (e) {
    return "保存エラー: " + e.message;
  }
}
