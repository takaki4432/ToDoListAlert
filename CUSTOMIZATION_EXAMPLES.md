# カスタマイズ実例集

このドキュメントでは、実際のカスタマイズ例を紹介します。

## 目次
1. [リマインドタイミングのカスタマイズ](#1-リマインドタイミングのカスタマイズ)
2. [実行時刻の変更](#2-実行時刻の変更)
3. [メール内容のカスタマイズ](#3-メール内容のカスタマイズ)
4. [優先度機能の追加](#4-優先度機能の追加)
5. [複数担当者への送信](#5-複数担当者への送信)

---

## 1. リマインドタイミングのカスタマイズ

### 例1: 1週間前にもリマインドを追加

**変更箇所**: `Code.gs` の設定部分

```javascript
// 変更前
const REMINDER_DAYS = [3, 1, 0, -1];

// 変更後
const REMINDER_DAYS = [7, 3, 1, 0, -1]; // 7日前を追加
```

### 例2: 当日のみリマインド

```javascript
const REMINDER_DAYS = [0]; // 当日のみ
```

### 例3: 締切後のリマインドを削除

```javascript
const REMINDER_DAYS = [3, 1, 0]; // 締切後のリマインドなし
```

---

## 2. 実行時刻の変更

### 例1: 朝7時に実行

**変更箇所**: `Code.gs` の `createDailyTrigger` 関数

```javascript
// 変更前
ScriptApp.newTrigger('sendReminders')
  .timeBased()
  .atHour(8)  // 午前8時
  .everyDays(1)
  .create();

// 変更後
ScriptApp.newTrigger('sendReminders')
  .timeBased()
  .atHour(7)  // 午前7時
  .everyDays(1)
  .create();
```

### 例2: 夜8時に実行

```javascript
.atHour(20)  // 午後8時（24時間表記）
```

### 例3: 1日2回実行（朝と夕方）

```javascript
function createDailyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 朝8時のトリガー
  ScriptApp.newTrigger('sendReminders')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
  
  // 夕方18時のトリガー
  ScriptApp.newTrigger('sendReminders')
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .create();
  
  Logger.log('朝8時と夕方18時のトリガーを作成しました');
  SpreadsheetApp.getUi().alert('トリガー設定完了', '朝8時と夕方18時にリマインドメールを送信します', SpreadsheetApp.getUi().ButtonSet.OK);
}
```

---

## 3. メール内容のカスタマイズ

### 例1: シンプルなメール

**変更箇所**: `Code.gs` の `sendReminderEmail` 関数

```javascript
// 変更前
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

// 変更後（シンプル版）
const body = `
${urgency}

【${taskName}】
締切: ${formattedDueDate}
状態: ${status}

詳細: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}
`;
```

### 例2: 絵文字を追加

```javascript
const body = `
⏰ ${urgency}

📋 タスク名: ${taskName}
📅 締切日: ${formattedDueDate}
📊 ステータス: ${status}
${note ? '📝 備考: ' + note : ''}

🔗 スプレッドシート:
${SpreadsheetApp.getActiveSpreadsheet().getUrl()}

✅ 完了したら、ステータスを「完了」に更新してください！
`;
```

### 例3: HTML形式のメール

```javascript
function sendReminderEmail(taskName, assignee, dueDate, daysUntilDue, status, note) {
  // 件名設定（変更なし）
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
  
  const formattedDueDate = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy年MM月dd日');
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  // HTML形式のメール本文
  const htmlBody = `
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 2px solid #4285F4; border-radius: 10px;">
          <h2 style="color: #4285F4; border-bottom: 2px solid #4285F4; padding-bottom: 10px;">
            📌 タスクリマインド
          </h2>
          
          <div style="background-color: #FFF3CD; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <strong style="color: #856404; font-size: 16px;">${urgency}</strong>
          </div>
          
          <table style="width: 100%; border-collapse: collapse;">
            <tr style="background-color: #F8F9FA;">
              <td style="padding: 10px; border: 1px solid #DEE2E6; width: 30%;"><strong>タスク名</strong></td>
              <td style="padding: 10px; border: 1px solid #DEE2E6;">${taskName}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #DEE2E6;"><strong>締切日</strong></td>
              <td style="padding: 10px; border: 1px solid #DEE2E6;">${formattedDueDate}</td>
            </tr>
            <tr style="background-color: #F8F9FA;">
              <td style="padding: 10px; border: 1px solid #DEE2E6;"><strong>ステータス</strong></td>
              <td style="padding: 10px; border: 1px solid #DEE2E6;">${status}</td>
            </tr>
            ${note ? `
            <tr>
              <td style="padding: 10px; border: 1px solid #DEE2E6;"><strong>備考</strong></td>
              <td style="padding: 10px; border: 1px solid #DEE2E6;">${note}</td>
            </tr>
            ` : ''}
          </table>
          
          <div style="text-align: center; margin-top: 30px;">
            <a href="${spreadsheetUrl}" 
               style="display: inline-block; padding: 12px 30px; background-color: #4285F4; color: white; text-decoration: none; border-radius: 5px; font-weight: bold;">
              📊 スプレッドシートで確認
            </a>
          </div>
          
          <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #DEE2E6; color: #6C757D; font-size: 12px;">
            <p>※ このメールは自動送信されています</p>
            <p>※ タスクが完了したら、スプレッドシートのステータスを「完了」に更新してください</p>
          </div>
        </div>
      </body>
    </html>
  `;
  
  try {
    // HTML形式でメールを送信
    MailApp.sendEmail({
      to: assignee,
      subject: subject,
      body: urgency, // プレーンテキスト版（フォールバック）
      htmlBody: htmlBody
    });
    
    Logger.log(`メール送信成功: ${assignee} - ${taskName}`);
  } catch (error) {
    Logger.log(`メール送信エラー: ${assignee} - ${taskName} - ${error}`);
  }
}
```

---

## 4. 優先度機能の追加

### ステップ1: 列の追加

`Code.gs` の `COLUMNS` オブジェクトに優先度列を追加：

```javascript
const COLUMNS = {
  TASK_NAME: 1,      // A列: タスク名
  ASSIGNEE: 2,       // B列: 担当者メールアドレス
  DUE_DATE: 3,       // C列: 締切日
  STATUS: 4,         // D列: ステータス
  PRIORITY: 5,       // E列: 優先度（新規追加）
  NOTE: 6            // F列: 備考
};
```

### ステップ2: setupSheet関数の修正

```javascript
function setupSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  sheet.setName('ToDoリスト');
  
  // ヘッダー行（優先度を追加）
  const headers = ['タスク名', '担当者メールアドレス', '締切日', 'ステータス', '優先度', '備考'];
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
  sheet.setColumnWidth(5, 80);   // 優先度（新規）
  sheet.setColumnWidth(6, 250);  // 備考
  
  // 締切日の列に日付フォーマットを設定
  sheet.getRange(2, COLUMNS.DUE_DATE, 1000, 1).setNumberFormat('yyyy/mm/dd');
  
  // サンプルデータ（優先度を追加）
  const sampleData = [
    ['プロジェクト提案書作成', 'your-email@example.com', new Date(new Date().getTime() + 2 * 24 * 60 * 60 * 1000), '未着手', '高', '重要案件'],
    ['クライアントミーティング準備', 'your-email@example.com', new Date(new Date().getTime() + 5 * 24 * 60 * 60 * 1000), '進行中', '中', '資料準備必要'],
    ['月次レポート提出', 'your-email@example.com', new Date(new Date().getTime() + 7 * 24 * 60 * 60 * 1000), '未着手', '低', ''],
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // データ行のフォーマット
  const dataRange = sheet.getRange(2, 1, 1000, headers.length);
  dataRange.setBorder(true, true, true, true, true, true);
  
  // ステータス列にドロップダウン
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未着手', '進行中', '完了', '保留'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, COLUMNS.STATUS, 1000, 1).setDataValidation(statusRule);
  
  // 優先度列にドロップダウン（新規）
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['高', '中', '低'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, COLUMNS.PRIORITY, 1000, 1).setDataValidation(priorityRule);
  
  // 優先度による色分け
  applyPriorityFormatting();
  
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);
  
  Logger.log('スプレッドシートの初期設定が完了しました');
  SpreadsheetApp.getUi().alert('設定完了', 'To Doリストの初期設定が完了しました！', SpreadsheetApp.getUi().ButtonSet.OK);
}

// 優先度による色分け関数
function applyPriorityFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  for (let i = 1; i < values.length; i++) {
    const priority = values[i][COLUMNS.PRIORITY - 1];
    const rowRange = sheet.getRange(i + 1, 1, 1, values[i].length);
    
    if (priority === '高') {
      rowRange.setBackground('#FFEBEE'); // 薄い赤
    } else if (priority === '中') {
      rowRange.setBackground('#FFF9E6'); // 薄い黄
    } else if (priority === '低') {
      rowRange.setBackground('#E8F5E9'); // 薄い緑
    }
  }
}
```

### ステップ3: メール送信時に優先度を含める

```javascript
function sendReminders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  let reminderCount = 0;
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const taskName = row[COLUMNS.TASK_NAME - 1];
    const assignee = row[COLUMNS.ASSIGNEE - 1];
    const dueDate = row[COLUMNS.DUE_DATE - 1];
    const status = row[COLUMNS.STATUS - 1];
    const priority = row[COLUMNS.PRIORITY - 1]; // 優先度を取得
    const note = row[COLUMNS.NOTE - 1];
    
    if (!taskName || !assignee || !dueDate || status === '完了') {
      continue;
    }
    
    const dueDateObj = new Date(dueDate);
    dueDateObj.setHours(0, 0, 0, 0);
    
    const daysUntilDue = Math.floor((dueDateObj - today) / (1000 * 60 * 60 * 24));
    
    if (REMINDER_DAYS.includes(daysUntilDue)) {
      sendReminderEmail(taskName, assignee, dueDateObj, daysUntilDue, status, priority, note); // 優先度を追加
      reminderCount++;
    }
  }
  
  Logger.log(`リマインドメールを${reminderCount}件送信しました`);
}

// メール送信関数（優先度パラメータを追加）
function sendReminderEmail(taskName, assignee, dueDate, daysUntilDue, status, priority, note) {
  // 件名に優先度を含める
  let subject = '';
  let urgency = '';
  const priorityEmoji = priority === '高' ? '🔴' : priority === '中' ? '🟡' : '🟢';
  
  if (daysUntilDue > 1) {
    subject = `${priorityEmoji}【リマインド】タスク締切まであと${daysUntilDue}日`;
    urgency = `締切まであと${daysUntilDue}日です。`;
  } else if (daysUntilDue === 1) {
    subject = `${priorityEmoji}【リマインド】タスク締切は明日です`;
    urgency = '締切は明日です！';
  } else if (daysUntilDue === 0) {
    subject = `${priorityEmoji}【重要】タスク締切は本日です`;
    urgency = '本日が締切日です！';
  } else {
    subject = `${priorityEmoji}【緊急】タスクの締切を過ぎています`;
    urgency = `締切を${Math.abs(daysUntilDue)}日過ぎています！`;
  }
  
  const formattedDueDate = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy年MM月dd日');
  
  const body = `
${urgency}

━━━━━━━━━━━━━━━━━━━━━━━━
【タスク情報】
━━━━━━━━━━━━━━━━━━━━━━━━

タスク名: ${taskName}
優先度: ${priorityEmoji} ${priority}
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
```

---

## 5. 複数担当者への送信

### メール送信関数の修正

```javascript
function sendReminders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  let reminderCount = 0;
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const taskName = row[COLUMNS.TASK_NAME - 1];
    const assignees = row[COLUMNS.ASSIGNEE - 1]; // 複数のメールアドレスを取得
    const dueDate = row[COLUMNS.DUE_DATE - 1];
    const status = row[COLUMNS.STATUS - 1];
    const note = row[COLUMNS.NOTE - 1];
    
    if (!taskName || !assignees || !dueDate || status === '完了') {
      continue;
    }
    
    const dueDateObj = new Date(dueDate);
    dueDateObj.setHours(0, 0, 0, 0);
    
    const daysUntilDue = Math.floor((dueDateObj - today) / (1000 * 60 * 60 * 24));
    
    if (REMINDER_DAYS.includes(daysUntilDue)) {
      // カンマ区切りで複数のメールアドレスに分割
      const assigneeList = assignees.split(',').map(email => email.trim());
      
      // 各担当者にメールを送信
      assigneeList.forEach(assignee => {
        if (assignee) { // 空でないことを確認
          sendReminderEmail(taskName, assignee, dueDateObj, daysUntilDue, status, note);
          reminderCount++;
        }
      });
    }
  }
  
  Logger.log(`リマインドメールを${reminderCount}件送信しました`);
}
```

**使い方**: 担当者列に `tanaka@example.com, suzuki@example.com, yamada@example.com` のようにカンマ区切りでメールアドレスを入力

---

## まとめ

これらのカスタマイズ例を参考に、自分のニーズに合わせてスクリプトを改変してください。

変更後は必ず：
1. コードを保存
2. 該当する関数を再実行
3. テスト実行で動作確認

を行うことをお勧めします。
