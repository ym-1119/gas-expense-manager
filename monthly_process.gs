// ============================================================
//  📅 月末処理（毎月1日の朝8時にトリガーで自動実行）
//  - 前月の合計を「月次サマリー」シートに記録
//  - 前月データをGoogleドライブにバックアップ
//  - LINE & Gmail で月次レポート通知
// ============================================================
function runMonthlyProcess() {
  // 前月を計算
  const now       = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const yearMonth = Utilities.formatDate(lastMonth, "Asia/Tokyo", "yyyy-MM");
  const monthLabel = yearMonth.replace("-", "年") + "月";

  Logger.log("月末処理開始: " + monthLabel);

  // 1. 前月サマリーをシートに保存
  const summary = saveMonthlySummaryToSheet(yearMonth, monthLabel);

  // 2. 前月データをGoogleドライブにバックアップ
  const backup = backupMonthlyData(yearMonth, monthLabel);

  // 3. LINE & Gmail で月次レポート通知
  sendMonthlyReport(monthLabel, summary, backup);
  deleteMonthlyData(yearMonth);

  Logger.log("月末処理完了");
}

// ============================================================
//  月次サマリーをシートに保存
// ============================================================
function saveMonthlySummaryToSheet(yearMonth, monthLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 「月次サマリー」シートがなければ作成
  let sheet = ss.getSheetByName("月次サマリー");
  if (!sheet) {
    sheet = ss.insertSheet("月次サマリー");
    sheet.appendRow(["年月", "合計金額", "件数", "予算", "過不足", "記録日時"]);
    const h = sheet.getRange(1, 1, 1, 6);
    h.setBackground("#2ECC71").setFontColor("#fff")
     .setFontWeight("bold").setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    [100, 120, 80, 120, 120, 160].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }

  // 前月データ集計
  const expenses = getExpenses(yearMonth);
  const total    = expenses.reduce((s, e) => s + e.amount, 0);
  const count    = expenses.length;
  const config   = getConfig();
  const budget   = config.budget;
  const diff     = budget - total; // プラスなら黒字、マイナスなら赤字
  const now      = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

  // 既存の同じ年月の行を全て削除してから追加（Date型対応）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    // 後ろから削除（行番号がずれないように）
    for (let i = data.length - 1; i >= 0; i--) {
      const cellVal = data[i][0] instanceof Date
        ? Utilities.formatDate(data[i][0], "Asia/Tokyo", "yyyy-MM")
        : String(data[i][0]);
      if (cellVal === yearMonth) sheet.deleteRow(i + 2);
    }
  }

  // 末尾に追加
  const rowData = [yearMonth, total, count, budget, diff, now];
  sheet.appendRow(rowData);
  const targetRow = sheet.getLastRow();

  // 金額フォーマット
  sheet.getRange(targetRow, 2).setNumberFormat("¥#,##0");
  sheet.getRange(targetRow, 4).setNumberFormat("¥#,##0");
  sheet.getRange(targetRow, 5).setNumberFormat("¥#,##0");

  // 過不足がマイナスなら赤くハイライト
  if (diff < 0) {
    sheet.getRange(targetRow, 5).setBackground("#FFF5F5").setFontColor("#E53E3E");
  } else {
    sheet.getRange(targetRow, 5).setBackground("#F0FFF4").setFontColor("#276749");
  }

  Logger.log("月次サマリー保存完了: " + monthLabel + " 合計¥" + total);
  return { total, count, budget, diff };
}

// ============================================================
//  前月データをGoogleドライブにバックアップ
// ============================================================
function backupMonthlyData(yearMonth, monthLabel) {
  try {
    const expenses = getExpenses(yearMonth);
    if (expenses.length === 0) {
      Logger.log("バックアップ対象データなし: " + monthLabel);
      return { success: true, message: "データなし" };
    }

    // バックアップ用スプレッドシートを新規作成
    const fileName = `【バックアップ】支出データ_${yearMonth}`;
    const newSS    = SpreadsheetApp.create(fileName);
    const sheet    = newSS.getActiveSheet();
    sheet.setName(yearMonth);

    // ヘッダー
    sheet.appendRow(["日付", "カテゴリ", "金額", "メモ", "登録日時"]);
    const h = sheet.getRange(1, 1, 1, 5);
    h.setBackground("#4A90D9").setFontColor("#fff")
     .setFontWeight("bold").setHorizontalAlignment("center");
    sheet.setFrozenRows(1);

    // データ書き込み
    expenses.forEach(e => {
      sheet.appendRow([e.date, e.category, e.amount, e.memo, e.createdAt]);
    });

    // 金額フォーマット
    if (expenses.length > 0) {
      sheet.getRange(2, 3, expenses.length, 1).setNumberFormat("¥#,##0");
    }

    // 列幅調整
    [100, 150, 100, 200, 160].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

    // 「支出管理バックアップ」フォルダを取得 or 作成
    const file         = DriveApp.getFileById(newSS.getId());
    const rootFolder   = DriveApp.getRootFolder();
    const backupIt     = DriveApp.getFoldersByName("支出管理バックアップ");
    const backupFolder = backupIt.hasNext()
      ? backupIt.next()
      : rootFolder.createFolder("支出管理バックアップ");

    // 年フォルダ（例: 2026）を取得 or 作成
    const year       = yearMonth.split("-")[0];
    const yearIt     = backupFolder.getFoldersByName(year);
    const yearFolder = yearIt.hasNext()
      ? yearIt.next()
      : backupFolder.createFolder(year);

    // ファイルを年フォルダに移動
    yearFolder.addFile(file);
    rootFolder.removeFile(file);

    const url = newSS.getUrl();
    Logger.log("バックアップ完了: " + fileName + " " + url);
    return { success: true, fileName, url, count: expenses.length };
  } catch (e) {
    Logger.log("バックアップエラー: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ============================================================
//  月次レポート通知（LINE & Gmail）
// ============================================================
function sendMonthlyReport(monthLabel, summary, backup) {
  const { total, count, budget, diff } = summary;
  const overBudget = diff < 0;
  const diffAbs    = Math.abs(diff);
  const usedRate   = budget > 0 ? Math.round(total / budget * 100) : 0;

  // LINE メッセージ
  const lineMsg =
    `📊 ${monthLabel} 月次レポート\n` +
    `────────────────\n` +
    `💸 合計支出: ¥${total.toLocaleString()}\n` +
    `📋 件数: ${count}件\n` +
    `💰 月予算: ¥${budget.toLocaleString()}\n` +
    `${overBudget ? "🚨" : "✅"} 予算対比: ${usedRate}%\n` +
    `${overBudget ? "⚠️ 超過" : "💚 残り"}: ¥${diffAbs.toLocaleString()}\n` +
    `────────────────\n` +
    (backup.success && backup.url ? `💾 バックアップ保存済み\n` : "") +
    `My支出管理アプリより`;

  sendLineNotify(lineMsg);

  // Gmail HTML
  const color     = overBudget ? "#E53E3E" : "#2ECC71";
  const bgColor   = overBudget ? "#FFF5F5" : "#F0FFF4";
  const emoji     = overBudget ? "⚠️" : "✅";
  const diffLabel = overBudget ? "超過額" : "残額";
  const toEmail   = Session.getActiveUser().getEmail();
  const subject   = `【My支出管理】${monthLabel} 月次レポート`;

  const html = `<!DOCTYPE html>
<html lang="ja">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="font-family:-apple-system,sans-serif;background:#F0F4FF;margin:0;padding:16px;">
  <div style="max-width:480px;margin:0 auto;background:#fff;border-radius:20px;overflow:hidden;box-shadow:0 8px 32px rgba(0,0,0,.12);">

    <!-- ヘッダー -->
    <div style="background:linear-gradient(135deg,#4A90D9,#667eea);padding:28px 24px;text-align:center;">

      <h1 style="color:#fff;font-size:22px;font-weight:800;margin:8px 0 4px;">月次レポート</h1>
      <p style="color:rgba(255,255,255,.85);font-size:14px;margin:0;">${monthLabel}</p>
    </div>

    <!-- 数字サマリー -->
    <div style="display:flex;border-bottom:1px solid #eee;">
      <div style="flex:1;padding:16px;text-align:center;border-right:1px solid #eee;">
        <div style="font-size:11px;color:#a0aec0;margin-bottom:4px;">合計支出</div>
        <div style="font-size:20px;font-weight:800;color:#4A90D9;">¥${total.toLocaleString()}</div>
      </div>
      <div style="flex:1;padding:16px;text-align:center;border-right:1px solid #eee;">
        <div style="font-size:11px;color:#a0aec0;margin-bottom:4px;">件数</div>
        <div style="font-size:20px;font-weight:800;color:#4A90D9;">${count}件</div>
      </div>
      <div style="flex:1;padding:16px;text-align:center;">
        <div style="font-size:11px;color:#a0aec0;margin-bottom:4px;">月予算</div>
        <div style="font-size:20px;font-weight:800;color:#4A90D9;">¥${budget.toLocaleString()}</div>
      </div>
    </div>

    <!-- 予算対比 -->
    <div style="padding:20px 24px;">
      <div style="display:flex;justify-content:space-between;font-size:13px;color:#718096;margin-bottom:8px;">
        <span>予算消化率</span>
        <span style="font-weight:700;color:${color};">${usedRate}%</span>
      </div>
      <div style="background:#eee;border-radius:99px;height:12px;overflow:hidden;">
        <div style="width:${Math.min(usedRate,100)}%;background:${color};height:12px;border-radius:99px;transition:width .8s;"></div>
      </div>
    </div>

    <!-- 過不足 -->
    <div style="margin:0 24px 20px;padding:16px;background:${bgColor};border-radius:12px;border-left:4px solid ${color};">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <span style="font-size:14px;font-weight:700;color:${color};">${emoji} ${diffLabel}</span>
        <span style="font-size:22px;font-weight:800;color:${color};">¥${diffAbs.toLocaleString()}</span>
      </div>
    </div>

    <!-- バックアップ情報 -->
    ${backup.success && backup.url ? `
    <div style="margin:0 24px 20px;padding:14px;background:#F7FAFC;border-radius:12px;border:1px solid #E2E8F0;">
      <div style="font-size:13px;font-weight:700;color:#4A90D9;margin-bottom:6px;">バックアップ完了</div>
      <div style="font-size:12px;color:#718096;">${backup.fileName}（${backup.count}件）</div>
      <a href="${backup.url}" style="font-size:12px;color:#4A90D9;">Googleドライブで確認 →</a>
    </div>` : ""}

    <!-- フッター -->
    <div style="padding:14px 24px;background:#F7FAFC;text-align:center;border-top:1px solid #eee;">
      <p style="color:#a0aec0;font-size:12px;margin:0;">My支出管理アプリ by Google Apps Script</p>
    </div>
  </div>
</body>
</html>`;

  try {
    GmailApp.sendEmail(toEmail, subject, "", { htmlBody: html });
    Logger.log("月次レポートメール送信完了: " + toEmail);
  } catch (e) {
    Logger.log("月次レポートメール送信エラー: " + e.toString());
  }
}

// ============================================================
//  🔧 トリガー設定（一度だけ実行してセットアップ）
//  GASエディタから setupMonthlyTrigger() を手動実行してください
// ============================================================
function setupMonthlyTrigger() {
  // 既存の月末トリガーを削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "runMonthlyProcess") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 毎月1日の朝8時に実行
  ScriptApp.newTrigger("runMonthlyProcess")
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  Logger.log("✅ 月次処理トリガーを設定しました（毎月1日 朝8時）");
}

function deleteMonthlyData(yearMonth) {
  const sheet   = getDataSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (!data[i][0]) continue;
    const ym = Utilities.formatDate(new Date(data[i][0]), "Asia/Tokyo", "yyyy-MM");
    if (ym === yearMonth) sheet.deleteRow(i + 2);
  }
  Logger.log("前月データ削除完了: " + yearMonth);
}

// ============================================================
//  🧪 月末処理テスト（手動実行用）
// ============================================================
function testMonthlyProcess() {
  // テスト用に現在月を使う
  const now       = new Date();
  const yearMonth = Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM");
  const monthLabel = yearMonth.replace("-", "年") + "月";

  Logger.log("テスト実行: " + monthLabel);
  const summary = saveMonthlySummaryToSheet(yearMonth, monthLabel);
  const backup  = backupMonthlyData(yearMonth, monthLabel);
  sendMonthlyReport(monthLabel, summary, backup);
  Logger.log("テスト完了");
}
