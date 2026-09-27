/**
 * =========================================================================
 * パナデンヨシマチ 配布ナビサポーター - Google Apps Script バックエンドAPI
 * =========================================================================
 * 
 * 【概要】
 * スマートフォンの配布アプリ（PWA）とGoogleスプレッドシート（2025年カレンダー等）
 * をリアルタイムで双方向連携させるAPIプログラムです。
 * 
 * 【使い方】
 * 1. Googleスプレッドシートを開き、メニューの「拡張機能」>「Apps Script」を開きます。
 * 2. 既存のコードをすべて消去し、このファイルの内容を丸ごと貼り付けます。
 * 3. 右上の「デプロイ」>「新しいデプロイ」をクリックします。
 *    - 種類の選択:「ウェブアプリ」
 *    - 説明:「配布サポーターAPI」
 *    - 次のユーザーとして実行:「自分」
 *    - アクセスできるユーザー:「全員 (Anyone)」※重要
 * 4. 発行された「ウェブアプリのURL」をコピーし、スマホアプリの設定画面に入力します。
 */

// 対象のシート名（エクセルファイル内のシート名に合わせています）
const TARGET_SHEET_NAME = "シート";

/**
 * 顧客一覧取得 API (GETリクエスト)
 * 配布予定（E列 = 1）のお客さんデータを抽出し、JSONでアプリへ返します。
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    
    if (!sheet) {
      return createJsonResponse({ error: "シート「" + TARGET_SHEET_NAME + "」が見つかりません。" }, 404);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 8) {
      return createJsonResponse({ customers: [], message: "データがありません。" });
    }

    // 7行目がヘッダー、8行目からデータ
    // A列からT列までのデータを一括取得
    const range = sheet.getRange(8, 1, lastRow - 7, 20);
    const values = range.getValues();

    const customers = [];
    let countTotal = 0;
    let countDone = 0;

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const rowNumber = i + 8; // 実際のスプレッドシート行番号

      const doneFlag = row[0];     // A列: 配布済み (1 or 0)
      const planFlag = row[4];     // E列: 配布予定 (1なら対象)
      const custId   = row[5];     // F列: 顧客番号
      const custName = row[6];     // G列: 顧客名
      const custKana = row[7];     // H列: フリガナ
      const zipCode  = row[8];     // I列: 郵便番号
      const address  = row[9];     // J列: 住所
      const phone    = row[10];    // K列: 電話番号
      const rank     = row[11];    // L列: ランク (A, B, C, D, E)
      const rfm      = row[12];    // M列: RFM値
      const rVal     = row[13];    // N列: R(最新購買日)
      const fVal     = row[14];    // O列: F(購買頻度)
      const mVal     = row[15];    // P列: M(購買金額)
      const gpsRaw   = String(row[16] || "").trim(); // Q列: GPS座標 (緯度, 経度)
      const updated  = row[19];    // T列: 更新日時

      // 配布予定が 1 の顧客、または氏名が入力されている行を対象とする
      if (custName && (planFlag === 1 || planFlag === "1" || planFlag === "")) {
        countTotal++;
        const isDone = (doneFlag === 1 || doneFlag === "1");
        if (isDone) countDone++;

        // GPS座標のパース
        let lat = null;
        let lng = null;
        if (gpsRaw && gpsRaw.indexOf(",") > -1) {
          const parts = gpsRaw.split(",");
          lat = parseFloat(parts[0].trim());
          lng = parseFloat(parts[1].trim());
        }

        customers.push({
          row: rowNumber,
          id: custId || rowNumber,
          name: String(custName).trim(),
          kana: String(custKana || "").trim(),
          zip: String(zipCode || "").trim(),
          addr: String(address || "").trim(),
          phone: String(phone || "").trim(),
          rank: String(rank || "C").trim(),
          rfm: String(rfm || "").trim(),
          r: formatDisplayDate(rVal),
          f: fVal || 0,
          m: mVal || 0,
          coords: (lat && lng) ? { lat: lat, lng: lng } : null,
          done: isDone ? 1 : 0,
          timestamp: formatDisplayTime(updated)
        });
      }
    }

    return createJsonResponse({
      success: true,
      timestamp: new Date().toISOString(),
      summary: {
        total: countTotal,
        done: countDone,
        remain: countTotal - countDone
      },
      customers: customers
    });

  } catch (error) {
    return createJsonResponse({ error: error.toString() }, 500);
  }
}

/**
 * 配布結果の書き込み API (POSTリクエスト)
 * スマホからの「手渡し」「投函」「不在」結果をスプレッドシートに書き込みます。
 */
function doPost(e) {
  try {
    let params;
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else {
      params = e.parameter;
    }

    const rowNumber = parseInt(params.row, 10);
    const status = parseInt(params.status, 10); // 1 = 完了, 0 = 未完了
    const type = params.type || "hand";         // hand (手渡し), post (投函), absent (不在)

    if (!rowNumber || isNaN(rowNumber)) {
      return createJsonResponse({ error: "有効な行番号 (row) が指定されていません。" }, 400);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    
    if (!sheet) {
      return createJsonResponse({ error: "シートが見つかりません。" }, 404);
    }

    const now = new Date();
    const formattedTimestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

    if (status === 1) {
      // 配布完了（手渡し or 投函）
      // A列に 1 を書き込み
      sheet.getRange(rowNumber, 1).setValue(1);
      // T列に 更新日時 を書き込み
      sheet.getRange(rowNumber, 20).setValue(formattedTimestamp);
      // R列（処理/備考）があれば種別を書き込み
      sheet.getRange(rowNumber, 18).setValue(type === "hand" ? "手渡し" : "ポスト投函");
    } else if (type === "absent") {
      // 不在・後日
      sheet.getRange(rowNumber, 1).setValue(0);
      sheet.getRange(rowNumber, 18).setValue("不在 (" + Utilities.formatDate(now, "Asia/Tokyo", "HH:mm") + ")");
      sheet.getRange(rowNumber, 20).setValue(formattedTimestamp);
    } else {
      // リセット（未訪問に戻す）
      sheet.getRange(rowNumber, 1).setValue(0);
      sheet.getRange(rowNumber, 18).setValue("");
      sheet.getRange(rowNumber, 20).setValue("");
    }

    return createJsonResponse({
      success: true,
      row: rowNumber,
      status: status,
      type: type,
      timestamp: formattedTimestamp,
      message: "スプレッドシートの行 " + rowNumber + " を更新しました。"
    });

  } catch (error) {
    return createJsonResponse({ error: error.toString() }, 500);
  }
}

/**
 * JSONレスポンスの生成ヘルパー
 */
function createJsonResponse(data, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function formatDisplayDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Tokyo", "yyyy/MM/dd");
  }
  return String(val);
}

function formatDisplayTime(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Tokyo", "HH:mm");
  }
  return String(val);
}
