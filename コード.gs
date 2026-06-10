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
  return data.map(row => ({ id:String(row[0]),name:String(row[1])}));
}catch(e){
  console.error("getSiteListエラー:"+e.message);
  return[];
  }
}

// ==========================================
// 追加: 既存データを取得する処理（エリア別・フォーム反映用）
// ==========================================
function getDailyData(dateStr, siteName, building, area) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('データ保存');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const result = [];
    
    for (let i = 1; i < data.length; i++) { // 1行目はヘッダー想定
      const row = data[i];
      if(!row[0]) continue;
      
      let dStr = "";
      if (row[0] instanceof Date) {
         dStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
         dStr = String(row[0]).replace(/\//g, '-');
      }
      
      // 日付、事業所、棟、エリアがすべて一致する行を抽出
      if (dStr === dateStr && row[1] === siteName && row[4] === building && row[5] === area) {
        const eqName = String(row[6]).replace(/^'/, ''); // 先頭のシングルクォートを除外
        result.push({
          no: eqName,
          item: String(row[7]),
          value: String(row[8])
        });
      }
    }
    return result;
  } catch(e) {
    console.error(e);
    return [];
  }
}

// ==========================================
// 追加: 既存データを取得する処理（棟全体・未入力バッジ計算用）
// ==========================================
function getBuildingData(dateStr, siteName, building) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('データ保存');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const result = [];
    
    for (let i = 1; i < data.length; i++) { // 1行目はヘッダー想定
      const row = data[i];
      if(!row[0]) continue;
      
      let dStr = "";
      if (row[0] instanceof Date) {
         dStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
         dStr = String(row[0]).replace(/\//g, '-');
      }
      
      // エリアは問わず、棟までの条件が一致する行を抽出
      if (dStr === dateStr && row[1] === siteName && row[4] === building) {
        const eqName = String(row[6]).replace(/^'/, '');
        result.push({
          area: String(row[5]),
          no: eqName,
          item: String(row[7]),
          value: String(row[8])
        });
      }
    }
    return result;
  } catch(e) {
    console.error(e);
    return [];
  }
}

// ==========================================
// 変更: データ保存（上書き処理対応）
// ==========================================
function saveData(results) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('データ保存');
    const now = new Date();
    
    // 既存データをすべて取得
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    
    // 既存データの検索用マップを作成（行番号を記録）
    const rowMap = new Map();
    for (let i = 1; i < data.length; i++) { // 1行目はヘッダー想定
      const row = data[i];
      if(!row[0]) continue;
      
      let dateStr = "";
      if (row[0] instanceof Date) {
         dateStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
         dateStr = String(row[0]).replace(/\//g, '-');
      }
      
      const eqName = String(row[6]).replace(/^'/, ''); // シングルクォートを除外
      
      // ユニークな検索キーを作成（日付|事業所|棟|エリア|設備名|項目名）
      const key = `${dateStr}|${row[1]}|${row[4]}|${row[5]}|${eqName}|${row[7]}`;
      rowMap.set(key, i + 1); // getRangeで使うため1-indexedの行番号にする
    }

    const newRows = [];
    
    results.forEach(res => {
      // 今回画面から送信されたデータのキー
      const key = `${res.date}|${res.site}|${res.building}|${res.area}|${res.no}|${res.item}`;
      
      if (rowMap.has(key)) {
        // 【上書き処理】すでに同じキーのデータが存在する場合は、その行をまるごと更新
        const rowNum = rowMap.get(key);
        sheet.getRange(rowNum, 1, 1, 10).setValues([[
          res.date,      // A: 点検日
          res.site,      // B: 事業所
          res.staff,     // C: 担当者
          res.weather,   // D: 天候
          res.building,  // E: 棟
          res.area,      // F: エリア
          "'" + res.no,  // G: 設備名
          res.item,      // H: 項目
          res.value,     // I: 値
          now            // J: 書込時間
        ]]);
      } else {
        // 【新規追加処理】データが存在しない場合は、新規追加リストに入れる
        newRows.push([
          res.date,
          res.site,
          res.staff,
          res.weather,
          res.building,
          res.area,
          "'" + res.no,
          res.item,
          res.value,
          now
        ]);
      }
    });

    // 新規追加行があれば、まとめてスプレッドシートの末尾に追加
    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }
    
    return "報告データの保存（上書き）が完了しました！";
  } catch (e) {
    return "保存エラー: " + e.message;
  }
}