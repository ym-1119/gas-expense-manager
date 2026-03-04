// ============================================================
//  💰 My支出管理アプリ - Code.gs
//  既存シート構成:
//    「支出データ」シート … 支出記録
//    「設定」シート      … 月予算 / 通知ON/OFF
// ============================aa================================

const DATA_SHEET   = "支出データ";
const CONFIG_SHEET = "設定";

// カテゴリ一覧
const CATEGORIES = [
  "🍔 食費"
];

// ============================================================
//  Webアプリ エントリーポイント
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("💰 My支出管理アプリ")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ============================================================
//  シート取得 / 初期化
// ============================================================
function getDataSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET);
}

function getConfigSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
}

// ============================================================
//  設定シート 読み書き
//  既存シート: A列=項目名, B列=値  例) "月予算" | 30000
// ============================================================
function getConfig() {
  const sheet   = getConfigSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { budget: 30000, notify: "ON" };

  const data   = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const config = { budget: 30000, notify: "ON" };
  data.forEach(row => {
    if (row[0] === "月予算") config.budget = Number(String(row[1]).replace(/,/g, ""));
    if (row[0] === "通知")   config.notify = row[1];
  });
  return config;
}

function saveConfig(budget, notify) {
  try {
    const sheet   = getConfigSheet();
    const lastRow = sheet.getLastRow();
    const data    = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    let budgetRow = -1, notifyRow = -1;
    data.forEach((row, i) => {
      if (row[0] === "月予算") budgetRow = i + 2;
      if (row[0] === "通知")   notifyRow = i + 2;
    });

    if (budgetRow > 0) sheet.getRange(budgetRow, 2).setValue(Number(budget));
    else               sheet.appendRow(["月予算", Number(budget)]);

    if (notifyRow > 0) sheet.getRange(notifyRow, 2).setValue(notify);
    else               sheet.appendRow(["通知", notify]);

    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ============================================================
//  カテゴリ取得
// ============================================================
function getCategories() { return CATEGORIES; }

// ============================================================
//  支出を保存（既存コードの saveExpence に合わせた命名）
// ============================================================
function saveExpence(data) {
  try {
    const sheet = getDataSheet();
    const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

    sheet.appendRow([
      new Date(data.date),
      data.category,
      Number(data.amount),
      data.memo || "",
      now
    ]);

    // 金額セルに通貨フォーマット
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 3).setNumberFormat("¥#,##0");
    sheet.getRange(lastRow, 1).setNumberFormat("yyyy/MM/dd");

    // 予算チェック → 通知
    const notifResult = checkBudget(data.date);
    return { success: true, message: "記録しました！", notification: notifResult };
  } catch (e) {
    return { sccess: false, message: e.toString() };
  }
}

// ============================================================
//  支出一覧取得
// ============================================================
function getExpenses(yearMonth) {
  if (!yearMonth) {
    yearMonth = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM");
  }

  try {
    const sheet = getDataSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    // A列からE列まで全データを取得
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    Logger.log("row[0]の型: " + (data[0][0] instanceof Date) + " 値: " + data[0][0]);

    const filteredRows = data.filter(row => {
      let dateVal = row[0];
      if (!dateVal) return false;

      let dateString = "";
      
      // 日付が Date オブジェクトの場合
      if (dateVal instanceof Date) {
        dateString = Utilities.formatDate(dateVal, "Asia/Tokyo", "yyyy-MM");
        Logger.log("dateString: " + dateString + " === " + yearMonth + " : " + (dateString === yearMonth));
      } else {
        // 文字列（2026/02/27など）で入っている場合
        let s = String(dateVal);
        // yyyy/MM/dd を yyyy-MM に変換
        let parts = s.split(/[\/\-]/); 
        if (parts.length >= 2) {
          dateString = parts[0] + "-" + parts[1].padStart(2, '0');
        }
      }

      return dateString === yearMonth;
    });

    // 表示用に整形して返す
    return filteredRows.map(row => {
      let d = row[0];
      let displayDate = (d instanceof Date) 
        ? Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd") 
        : String(d);

      return {
        date: displayDate,
        category: row[1],
        amount: Number(row[2]) || 0,
        memo: row[3] || "",
        createdAt: row[4] || ""
      };
    });

  } catch (e) {
    console.error("エラー詳細: " + e.message);
    return [];
  }
}

function getExpensesWrapped(params) {
  // params 自体が無い、または ym プロパティが無い場合の徹底ガード
  let targetYm = "";
  
  if (params && params.ym) {
    targetYm = params.ym;
  } else {
    // 最終手段：サーバー側の現在時刻から年月を作る
    targetYm = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM");
  }

  Logger.log("最終的に使用する年月: " + targetYm);
  const result = getExpenses(targetYm);
  Logger.log("返却件数: " + result.length);
  return JSON.stringify(result);
}

function getMonthlySummaryWrapped(params) {
  const result = getMonthlySummary(params.ym);
  return JSON.stringify(result);
}

// ============================================================
//  月次サマリー
// ============================================================
function getMonthlySummary(yearMonth) {
  try {
    const expenses   = JSON.parse(getExpensesWrapped({ ym: yearMonth }));
    const total      = expenses.reduce((s, e) => s + e.amount, 0);
    const byCategory = {};
    expenses.forEach(e => {
      byCategory[e.category] = (byCategory[e.category] || 0) + e.amount;
    });
    const categoryList = Object.entries(byCategory)
      .map(([category, amount]) => ({ category, amount}))
      .sort((a, b) => b.amount - a.amount);
    return { total, categoryList, count: expenses.length };
  } catch (e) { return { total: 0, categoryList: [], count: 0 }; }
}

// ============================================================
//  支出削除（行番号で削除）
// ============================================================
function deleteExpense(rowIndex) {
  try {
    // rowIndex は実際のシート行番号ではなく getExpenses の順番なので
    // 日付+カテゴリ+金額で一致行を探して削除
    const sheet   = getDataSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false };
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

    for (let i = 0; i < data.length; i++) {
      const d = Utilities.formatDate(new Date(data[i][0]), "Asia/Tokyo", "yyyy/MM/dd");
      if (d === rowIndex.date && data[i][1] === rowIndex.category && Number(data[i][2]) === Number(rowIndex.amount)) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, message: "対象が見つかりません" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ============================================================
//  🔔 予算超過チェック
// ============================================================
function checkBudget(dateStr) {
  try {
    const config = getConfig();

    // 通知OFFなら何もしない
    if (config.notify !== "ON") return { notified: false, reason: "通知OFF" };

    const targetDate = dateStr ? new Date(dateStr) : new Date();
    const yearMonth  = Utilities.formatDate(targetDate, "Asia/Tokyo", "yyyy-MM");

    // 今月の合計支出
    const expenses = getExpenses(yearMonth);
    const total    = expenses.reduce((s, e) => s + e.amount, 0);
    const budget   = config.budget;

    if (total <= budget) return { notified: false, reason: "予算内", total, budget };

    // 超過確定
    const over       = total - budget;
    const overRate   = Math.round((total / budget) * 100);
    const monthLabel = yearMonth.replace("-", "年") + "月";

    const lineMsg = 
      `⚠️ 月予算超過アラート\n` +
      `────────────────\n` + 
      `📅 ${monthLabel}\n` +
      `💰 月予算: ¥${budget.toLocaleString()}\n` +
      `💸 今月支出: ¥${total.toLocaleString()} (${overRate}%)\n` +
      `🚨 超過額: ¥${over.toLocaleString()}\n` +
      `────────────────\n` +
      `My支出管理アプリより`;
    
    const lineResult  = sendLineNotify(lineMsg);
    const gmailResult = sendGmailNotify(monthLabel, budget, total, over, overRate);

    return { notified: true, total, budget, over, line: lineResult, gmail: gmailResult };
  } catch (e) { return { notified: false, error: e.toString() }; }
}

// ============================================================
//  LINE Messaging API
// ============================================================

// ▼ ここを書き換えてください
const LINE_TOKEN   = "CN/xGPyORczufPcKHDtHLLgg4r12MkowE16GMIN54fY7CtaRD70NoLLZBd8KhjuBEjoHnEr+ky1wYDCtjYot1ahYGijKooJFpb7YGLyFL5wQo2MTfR/jy52FfG5GpGxRwfN6860njFvz+tqLpZxrOAdB04t89/1O/w1cDnyilFU=";
const LINE_USER_ID = "U3a020b9410a55db96d6ce339efcc0419";

function sendLineNotify(message) {
  try {
    if (!LINE_TOKEN) {
      return { success: false, reason: "LINEトークン未設定" };
    }
    const res = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method : "post",
      headers: {
        "Content-Type" : "application/json",
        "Authorization": "Bearer " + LINE_TOKEN
      },
      payload: JSON.stringify({
        to      : LINE_USER_ID,
        messages: [{ type: "text", text: message }]
      }),
      muteHttpExceptions: true
    });
    return { success: res.getResponseCode() === 200, code: res.getResponseCode() };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ============================================================
//  Gmail 通知（設定シートの通知先 or 実行アカウント）
// ============================================================
function sendGmailNotify(monthLabel, budget, total, over, overRate) {
  try {
    const toEmail = Session.getEffectiveUser().getEmail();
    const subject = `⚠️【My支出管理】月予算を超過しました（${monthLabel}）`;
    const html = `
<!DOCTYPE html>
<html lang="ja">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="font-family:-apple-system,sans-serif;background:#F0F4FF;margin:0;padding:16px;">
  <div style="max-width:480px;margin:0 auto;background:#fff;border-radius:20px;overflow:hidden;box-shadow:0 8px 32px rgba(0,0,0,.12);">
    <div style="background:linear-gradient(135deg,#FF6B6B,#ee5a24);padding:28px 24px;text-align:center;">
      <div style="font-size:48px;">⚠️</div>
      <h1 style="color:#fff;font-size:22px;font-weight:800;margin:8px 0 4px;">月予算超過アラート</h1>
      <p style="color:rgba(255,255,255,.85);font-size:14px;margin:0;">${monthLabel}</p>
    </div>
    <div style="display:flex;border-bottom:1px solid #eee;">
      <div style="flex:1;padding:18px;text-align:center;border-right:1px solid #eee;">
        <div style="font-size:11px;color:#a0aec0;margin-bottom:4px;">月予算</div>
        <div style="font-size:20px;font-weight:800;color:#4A90D9;">¥${budget.toLocaleString()}</div>
      </div>
      <div style="flex:1;padding:18px;text-align:center;border-right:1px solid #eee;">
        <div style="font-size:11px;color:#a0aec0;margin-bottom:4px;">今月支出</div>
        <div style="font-size:20px;font-weight:800;color:#e53e3e;">¥${total.toLocaleString()}</div>
      </div>
      <div style="flex:1;padding:18px;text-align:center;">
        <div style="font-size:11px;color:#a0aec0;margin-bottom:4px;">超過額</div>
        <div style="font-size:20px;font-weight:800;color:#e53e3e;">¥${over.toLocaleString()}</div>
      </div>
    </div>
    <div style="padding:20px 24px;">
      <div style="display:flex;justify-content:space-between;font-size:13px;color:#718096;margin-bottom:8px;">
        <span>予算消化率</span><span style="font-weight:700;color:#e53e3e;">${overRate}%</span>
      </div>
      <div style="background:#FED7D7;border-radius:99px;height:12px;">
        <div style="width:100%;background:linear-gradient(90deg,#FC8181,#e53e3e);height:12px;border-radius:99px;"></div>
      </div>
    </div>
    <div style="margin:0 24px 20px;padding:16px;background:#FFF5F5;border-radius:12px;border-left:4px solid #FC8181;">
      <p style="color:#c53030;font-size:14px;line-height:1.7;margin:0;">
        今月の支出が月予算を <strong>¥${over.toLocaleString()}（${overRate - 100}%オーバー）</strong> 超えました。
        アプリで内訳を確認しましょう。
      </p>
    </div>
    <div style="padding:14px 24px;background:#F7FAFC;text-align:center;border-top:1px solid #eee;">
      <p style="color:#a0aec0;font-size:12px;margin:0;">💰 My支出管理アプリ by Google Apps Script</p>
    </div>
  </div>
</body>
</html>`;
    GmailApp.sendEmail(toEmail, subject, "", { htmlBody: html });
    return { success: true, to: toEmail };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ============================================================
//  ✅ 通知テスト（スクリプトエディタから実行）
// ============================================================
function testNotification() {
  const lineResult = sendLineNotify(
    "🔔 テスト通知\nMy支出管理アプリのLINE通知が\n正常に動作しています！"
  );
  Logger.log("LINE: " + JSON.stringify(lineResult));

  const email = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(email, "✅ My支出管理アプリ - 通知テスト", "",
    { htmlBody: "<div style='font-family:sans-serif;padding:20px'><h2 style='color:#4A90D9'>✅ Gmail通知テスト成功</h2><p>正常に設定されています！</p></div>" }
  );
  Logger.log("Gmail送信先: " + email);
}

// ============================================================
//  ✅ 削除テスト（スクリプトエディタから実行）
// ============================================================
function testDeleteLastMonth() {
  deleteMonthlyData("2026-02");
}














