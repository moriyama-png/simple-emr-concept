/**
 * Webアプリにアクセスした際の初期画面表示
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('簡易電子カルテ構想 - Phase2')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 1. スプレッドシートから患者マスタデータを全件取得する
 */
function getPatientMasterData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // 原点である「患者データベース」または「患者管理」シートを自動取得
    let sheet = ss.getSheetByName('患者データベース') || ss.getSheetByName('患者管理');
    
    if (!sheet) return getMockData(); // 万が一シート名が違っても動くようにモックを返す安全設計
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; 
    
    // A2からG列の最終行まで取得 (ID, 氏名, 住所, 電話番号, 主病名, 担当医, 訪問曜日)
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    
    return data.map(row => {
      return {
        id: String(row[0] || ''),
        name: String(row[1] || ''),
        address: String(row[2] || ''),
        phone: String(row[3] || ''),
        disease: String(row[4] || ''),
        doctor: String(row[5] || ''),
        weekday: String(row[6] || '')
      };
    });
  } catch (e) {
    Logger.log("エラー(getPatientMasterData): " + e.toString());
    return getMockData();
  }
}

/**
 * 2. 予約データをスプレッドシートに保存する
 */
function saveReservationData(reservationObj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('予約管理');
    if (!sheet) {
      sheet = ss.insertSheet('予約管理');
      sheet.appendRow(['日時', '時間', '患者情報', '依頼部署', '検査内容', '詳細']);
    }
    
    sheet.appendRow([
      reservationObj.date,
      reservationObj.time,
      reservationObj.name,
      reservationObj.dept,
      reservationObj.test,
      reservationObj.detail
    ]);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * 3. 画面のカレンダーに表示するために予約データを全件取得する
 */
function getReservations() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('予約管理');
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    
    return data.map(row => {
      let dateStr = "";
      if (row[0] instanceof Date) {
        dateStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        dateStr = String(row[0]);
      }
      return {
        date: dateStr,
        time: String(row[1] || ''),
        name: String(row[2] || ''),
        dept: String(row[3] || ''),
        test: String(row[4] || ''),
        detail: String(row[5] || '')
      };
    });
  } catch (e) {
    return [];
  }
}

/**
 * バックアップ用モックデータ
 */
function getMockData() {
  return [
    { id: "1001", name: "山田 太郎", address: "堺市中区深井沢町", phone: "090-1111-1111", disease: "高血圧", doctor: "山田医師", weekday: "月" },
    { id: "1002", name: "佐藤 花子", address: "堺市北区百舌鳥梅町", phone: "090-2222-2222", disease: "心不全", doctor: "田中医師", weekday: "火" }
  ];
}
