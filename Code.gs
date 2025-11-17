/**
 * Google スプレッドシート To Do リスト with リマインドメール
 * 
 * 機能:
 * - To Doリストの管理
 * - 締切日が近づいたら自動でリマインドメールを送信
 * - 毎日自動チェック
 */

// ========================================
// 設定項目
// ========================================

// リマインドを送信する日数（締切日の何日前/後）
const REMINDER_DAYS = [3, 1, 0, -1]; // 3日前、1日前、当日、1日後

// スプレッドシートの列番号
const COLUMNS = {
  TASK_NAME: 1,      // A列: タスク名
  ASSIGNEE: 2,       // B列: 担当者メールアドレス
  DUE_DATE: 3,       // C列: 締切日
  STATUS: 4,         // D列: ステータス
  NOTE: 5            // E列: 備考
};

// ========================================
// 初期設定関数
// ========================================

/**
 * スプレッドシートの初期設定を行う
 * 最初に一度だけ実行してください
 */
function setupSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // シート名を設定
  sheet.setName('ToDoリスト');
  
  // ヘッダー行を作成
  const headers = ['タスク名', '担当者メールアドレス', '締切日', 'ステータス', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーのフォーマット
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 300);  // タスク名
  sheet.setColumnWidth(2, 200);  // 担当者メールアドレス
  sheet.setColumnWidth(3, 120);  // 締切日
  sheet.setColumnWidth(4, 100);  // ステータス
  sheet.setColumnWidth(5, 250);  // 備考
  
  // 締切日の列に日付フォーマットを設定
  sheet.getRange(2, COLUMNS.DUE_DATE, 1000, 1).setNumberFormat('yyyy/mm/dd');
  
  // サンプルデータを追加
  const sampleData = [
    ['プロジェクト提案書作成', 'your-email@example.com', new Date(new Date().getTime() + 2 * 24 * 60 * 60 * 1000), '未着手', '重要案件'],
    ['クライアントミーティング準備', 'your-email@example.com', new Date(new Date().getTime() + 5 * 24 * 60 * 60 * 1000), '進行中', '資料準備必要'],
    ['月次レポート提出', 'your-email@example.com', new Date(new Date().getTime() + 7 * 24 * 60 * 60 * 1000), '未着手', ''],
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // データ行のフォーマット
  const dataRange = sheet.getRange(2, 1, 1000, headers.length);
  dataRange.setBorder(true, true, true, true, true, true);
  
  // ステータス列にドロップダウンを設定
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未着手', '進行中', '完了', '保留'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, COLUMNS.STATUS, 1000, 1).setDataValidation(statusRule);
  
  // 行の高さを調整
  sheet.setRowHeight(1, 40);
  
  // シートを固定
  sheet.setFrozenRows(1);
  
  Logger.log('スプレッドシートの初期設定が完了しました');
  SpreadsheetApp.getUi().alert('設定完了', 'To Doリストの初期設定が完了しました！\n\nサンプルデータのメールアドレスを自分のメールアドレスに変更してください。', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========================================
// メインのリマインド処理
// ========================================

/**
 * リマインドメールを送信するメイン関数
 * 毎日自動実行されます
 */
function sendReminders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // 今日の日付（時刻を00:00:00にリセット）
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  let reminderCount = 0;
  
  // ヘッダー行をスキップして、各行をチェック
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const taskName = row[COLUMNS.TASK_NAME - 1];
    const assignee = row[COLUMNS.ASSIGNEE - 1];
    const dueDate = row[COLUMNS.DUE_DATE - 1];
    const status = row[COLUMNS.STATUS - 1];
    const note = row[COLUMNS.NOTE - 1];
    
    // データが空の行やステータスが「完了」の場合はスキップ
    if (!taskName || !assignee || !dueDate || status === '完了') {
      continue;
    }
    
    // 締切日を日付型に変換し、時刻を00:00:00にリセット
    const dueDateObj = new Date(dueDate);
    dueDateObj.setHours(0, 0, 0, 0);
    
    // 締切日までの日数を計算
    const daysUntilDue = Math.floor((dueDateObj - today) / (1000 * 60 * 60 * 24));
    
    // リマインド対象の日数かチェック
    if (REMINDER_DAYS.includes(daysUntilDue)) {
      sendReminderEmail(taskName, assignee, dueDateObj, daysUntilDue, status, note);
      reminderCount++;
    }
  }
  
  Logger.log(`リマインドメールを${reminderCount}件送信しました`);
}

/**
 * リマインドメールを送信する
 * 
 * @param {string} taskName - タスク名
 * @param {string} assignee - 担当者メールアドレス
 * @param {Date} dueDate - 締切日
 * @param {number} daysUntilDue - 締切日までの日数
 * @param {string} status - ステータス
 * @param {string} note - 備考
 */
function sendReminderEmail(taskName, assignee, dueDate, daysUntilDue, status, note) {
  // メールの件名を日数に応じて変更
  let subject = '';
  let urgency = '';
  
  if (daysUntilDue > 1) {
    subject = `【リマインド】タスク締切まであと${daysUntilDue}日`;
    urgency = `締切まであと${daysUntilDue}日です。`;
  } else if (daysUntilDue === 1) {
    subject = '【リマインド】タスク締切は明日です';
    urgency = '締切は明日です！';
  } else if (daysUntilDue === 0) {
    subject = '【重要】タスク締切は本日です';
    urgency = '本日が締切日です！';
  } else {
    subject = '【緊急】タスクの締切を過ぎています';
    urgency = `締切を${Math.abs(daysUntilDue)}日過ぎています！`;
  }
  
  // メール本文を作成
  const formattedDueDate = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy年MM月dd日');
  
  const body = `
${urgency}

━━━━━━━━━━━━━━━━━━━━━━━━
【タスク情報】
━━━━━━━━━━━━━━━━━━━━━━━━

タスク名: ${taskName}
締切日: ${formattedDueDate}
現在のステータス: ${status}
${note ? '備考: ' + note : ''}

━━━━━━━━━━━━━━━━━━━━━━━━

スプレッドシートで確認・更新:
${SpreadsheetApp.getActiveSpreadsheet().getUrl()}

※ このメールは自動送信されています
※ タスクが完了したら、スプレッドシートのステータスを「完了」に変更してください
`;
  
  try {
    // メールを送信
    MailApp.sendEmail({
      to: assignee,
      subject: subject,
      body: body
    });
    
    Logger.log(`メール送信成功: ${assignee} - ${taskName}`);
  } catch (error) {
    Logger.log(`メール送信エラー: ${assignee} - ${taskName} - ${error}`);
  }
}

// ========================================
// トリガー設定関数
// ========================================

/**
 * 毎日自動実行するトリガーを作成
 * 初回のみ実行してください
 */
function createDailyTrigger() {
  // 既存のトリガーを削除（重複防止）
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 新しいトリガーを作成（毎日午前8時に実行）
  ScriptApp.newTrigger('sendReminders')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
  
  Logger.log('毎日のトリガーを作成しました（午前8時実行）');
  SpreadsheetApp.getUi().alert('トリガー設定完了', '毎日午前8時にリマインドメールを自動送信するトリガーを設定しました！', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * トリガーを削除
 */
function deleteDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
      deletedCount++;
    }
  }
  
  Logger.log(`${deletedCount}個のトリガーを削除しました`);
  SpreadsheetApp.getUi().alert('トリガー削除完了', `${deletedCount}個のトリガーを削除しました`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========================================
// テスト・ユーティリティ関数
// ========================================

/**
 * テスト用: 今すぐリマインドをチェック
 */
function testReminders() {
  Logger.log('=== リマインドメールのテスト実行 ===');
  sendReminders();
  SpreadsheetApp.getUi().alert('テスト完了', 'リマインドメールのテスト送信が完了しました。実行ログを確認してください。', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 現在設定されているトリガーを表示
 */
function showTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let message = '現在のトリガー設定:\n\n';
  
  if (triggers.length === 0) {
    message += 'トリガーは設定されていません';
  } else {
    for (let i = 0; i < triggers.length; i++) {
      const trigger = triggers[i];
      message += `${i + 1}. 関数: ${trigger.getHandlerFunction()}\n`;
      message += `   種類: ${trigger.getEventType()}\n\n`;
    }
  }
  
  SpreadsheetApp.getUi().alert('トリガー情報', message, SpreadsheetApp.getUi().ButtonSet.OK);
}
