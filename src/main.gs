/**
 * ======================================================================
 * mt. inn 修繕稟議システム v2.0 (Gemini 3.0 Pro版)
 * ======================================================================
 * 目的: 現場からの修繕報告メールをAIが解析し、Chatでのワンタップ承認フローを実現
 * 
 * 主な機能:
 * - メール受信トリガー（subject:修繕依頼）
 * - Gemini 3.0 Proで画像・本文解析（Google検索ツール有効）
 * - 稟議書Docs自動生成
 * - Chat通知（Cards V2ボタン付き）
 * - Webアプリでボタンアクション処理
 * ======================================================================
 */

// ====== 設定 ======
const CONFIG = {
  GEMINI_API_KEY: '***REDACTED***', // 新APIキー（更新）
  REPAIR_SYSTEM_SHEET_ID: '1ZAUzoCIIy3h6TNiVnYB7_hjWY-Id9oZ_iX1z88M2yNI',
  SHEET_NAME: '修繕ログ',
  FOLDER_ID: '1Qz-HYebqH-vfd8-cYD-xoLsOdEL7PEg5',
  TEMPLATE_DOC_ID: '1iazbzvlh-VQ046dVgRXyO2BEEWbnGIVHbBTeejeGSjk',
  WEBHOOK_URL: 'https://chat.googleapis.com/v1/spaces/AAQAmERWyO4/messages?key=***REDACTED***&token=***REDACTED***',
  SCRIPT_WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbxNdoh_xEOgTzxG-uflNUxm0ELkv28_7sw7s1541ueeQ47f_qdQpahxg2m7ddlFAvvdTQ/exec', // 最新デプロイURL
  // Geminiモデル設定（2025年12月最新構成）
  GEMINI_MODEL_PRIMARY: 'gemini-2.5-flash',          // 安定版を優先（3.0は404のため一時停止）
  GEMINI_MODEL_FALLBACK: 'gemini-1.5-pro-latest',    // 代替の安定モデル
  EMAIL_SUBJECT: '修繕依頼',
  ENABLE_ERROR_NOTIFICATION: true // エラー通知をChatに送るか（false=ログのみ、true=Chatにも通知）
};

// ====== 列定義（42列） ======
const COL = {
  REPAIR_ID: 0,            // A: 修繕ID (R-yyyy-MM-dd-xxx)
  RECEIVED_DATETIME: 1,     // B: 受付日時
  REPORTER_NAME: 2,         // C: 報告者名 (Gmail差出人)
  AREA: 3,                  // D: エリア (AI出力1)
  LOCATION_DETAIL: 4,       // E: 場所詳細 (AI出力2)
  PHOTO1: 5,                // F: 写真1 (Drive URL / HYPERLINK関数)
  PHOTO2: 6,                // G: 写真2 (Drive URL / HYPERLINK関数)
  PHOTO3: 7,                // H: 写真3 (Drive URL / HYPERLINK関数)
  ORIGINAL_TEXT: 8,         // I: 原文 (メール本文)
  AI_FORMATTED: 9,          // J: AI整形文 (AI出力3)
  PROBLEM_SUMMARY: 10,      // K: 問題要約 (AI出力4)
  CAUSE_ANALYSIS: 11,       // L: 原因分析 (AI出力5)
  PRIORITY_RANK: 12,        // M: 重要度 (AI出力6)
  RANK_REASON: 13,          // N: ランク理由 (AI出力7)
  RECOMMENDED_TYPE: 14,     // O: 推奨対応 (AI出力8)
  WORK_SUMMARY: 15,         // P: 作業内容要約 (AI出力9)
  RECOMMENDED_STEPS: 16,    // Q: 作業手順 (AI出力10)
  MATERIALS_LIST: 17,       // R: 必要部材 (AI出力11 - URL付き)
  ESTIMATED_TIME: 18,       // S: 作業時間 (AI出力12)
  AI_COST_MIN: 19,          // T: 費用下限 (AI出力13)
  AI_COST_MAX: 20,          // U: 費用上限 (AI出力14)
  CONTRACTOR_CATEGORY: 21,  // V: 業者カテゴリ (AI出力15)
  CONTRACTOR_AREA: 22,      // W: 業者エリア (AI出力16)
  SEARCH_KEYWORDS: 23,      // X: 検索KW (AI出力17)
  POSTPONE_RISK: 24,        // Y: 先送りリスク (AI出力18)
  ESTIMATE1: 25,            // Z: 見積1 (AI出力19 - HYPERLINK)
  ESTIMATE2: 26,            // AA: 見積2 (AI出力20 - HYPERLINK)
  ESTIMATE3: 27,            // AB: 見積3 (AI出力21 - HYPERLINK)
  SELECTED_CONTRACTOR: 28,  // AC: 選定業者 (AI出力22)
  RINGI_INITIATOR: 29,      // AD: 稟議起案者 (=報告者名)
  RINGI_REQUIRED: 30,       // AE: 稟議要否 (自動判定)
  RINGI_REASON: 31,         // AF: 稟議理由 (AI出力23)
  STATUS: 32,               // AG: ステータス (初期値: "下書き")
  ACTUAL_OWNER: 33,         // AH: 実務担当者
  ACTUAL_COST: 34,          // AI: 実際費用
  RINGI_ID: 35,             // AJ: 稟議ID (Docs URL)
  RINGI_STATUS: 36,         // AK: 稟議ステータス
  APPROVAL_STATUS: 37,      // AL: 承認ステータス
  APPROVAL_TYPE: 38,        // AM: 承認区分
  APPROVAL_REASON: 39,      // AN: 理由
  COMPLETION_DATE: 40,      // AO: 完了日
  NOTES: 41                 // AP: 備考
};

// ====== ユーティリティ ======
function getSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.REPAIR_SYSTEM_SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    // ヘッダー行を設定
    const headers = [
      '修繕ID', '受付日時', '報告者名', 'エリア', '場所詳細', '写真1', '写真2', '写真3',
      '原文（現場入力）', 'AI整形文', '問題要約', '原因分析', '重要度ランク（A/B/C）', 'ランク理由',
      '推奨対応タイプ', '作業内容要約', '推奨作業手順', '必要部材リスト', '想定作業時間（分）',
      'AI概算費用下限（円）', 'AI概算費用上限（円）', '想定業者カテゴリ', '想定業者エリア', '業者検索キーワード',
      '先送りリスク', '見積1', '見積2', '見積3', '選定業者', '稟議起案者', '稟議要否', '稟議理由',
      'ステータス', '実務担当者', '実際費用（円）', '稟議ID', '稟議ステータス', '承認ステータス',
      '承認区分', '理由', '完了日', '備考'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function getTodayTokyo() {
  const now = new Date();
  const tokyoTime = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  return tokyoTime;
}

function generateRepairId() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, 'Asia/Tokyo', 'HHmm');
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  const seq = lastRow > 1 ? String(lastRow).padStart(3, '0') : '001';
  return `R-${dateStr}-${seq}`;
}

// ====== メール受信処理 ======
function processRepairEmails() {
  try {
    Logger.log('【メール処理開始】修繕依頼メールを検索します');
    
    // 今日以降のメールのみ処理（過去分は除外）
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
    
    // 検索条件：件名完全一致、今日以降
    const query = `subject:"${CONFIG.EMAIL_SUBJECT}" after:${todayStr}`;
    Logger.log(`【メール処理】検索クエリ: ${query}`);
    
    const threads = GmailApp.search(query, 0, 20);
    Logger.log(`【メール処理】見つかったスレッド数: ${threads.length}`);
    
    if (threads.length === 0) {
      Logger.log('【メール処理】処理対象のメールがありません');
      return '処理対象のメールがありません';
    }
    
    const processedEmailIds = getProcessedEmailIds();
    Logger.log(`【メール処理】処理済みメールID数: ${processedEmailIds.size}`);
    
    let processedCount = 0;
    let skippedCount = 0;
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      Logger.log(`【メール処理】スレッド内のメッセージ数: ${messages.length}`);
      
      for (const message of messages) {
        const messageId = message.getId();
        const subject = message.getSubject();
        const messageDate = message.getDate();
        
        // 件名が完全一致するか確認
        if (subject !== CONFIG.EMAIL_SUBJECT) {
          Logger.log(`【メール処理】件名が一致しないためスキップ: "${subject}"`);
          skippedCount++;
          continue;
        }
        
        // 今日以降のメールのみ処理
        if (messageDate < today) {
          Logger.log(`【メール処理】過去のメールのためスキップ: 件名="${subject}", 日時=${messageDate}`);
          skippedCount++;
          continue;
        }
        
        // 処理済みチェック
        if (processedEmailIds.has(messageId)) {
          Logger.log(`【メール処理】処理済みのためスキップ: 件名="${subject}", ID=${messageId}`);
          skippedCount++;
          continue;
        }
        
        Logger.log(`【メール処理】新規メールを処理: 件名="${subject}", ID=${messageId}, 日時=${messageDate}`);
        
        try {
          processRepairEmail(message, messageId);
          processedCount++;
        } catch (error) {
          Logger.log(`【メール処理】エラー: ${error.toString()}`);
          throw error;
        }
      }
    }
    
    Logger.log(`【メール処理】処理完了: 新規処理=${processedCount}件, スキップ=${skippedCount}件`);
    return `処理完了: 新規処理=${processedCount}件, スキップ=${skippedCount}件`;
  } catch (error) {
    const errorMessage = `メール処理エラー: ${error.toString()}`;
    Logger.log(errorMessage);
    Logger.log(`【メール処理】スタック: ${error.stack}`);
    sendErrorNotification(errorMessage);
    return `エラー: ${error.toString()}`;
  }
}

// ====== 処理済みメールIDを取得 ======
function getProcessedEmailIds() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return new Set(); // ヘッダー行のみの場合は空のSetを返す
  }
  
  const processedIds = new Set();
  
  // NOTES列（AP列、42列目）から処理済みメールIDを抽出
  const notesRange = sheet.getRange(2, COL.NOTES + 1, lastRow - 1, 1);
  const notesValues = notesRange.getValues();
  
  for (let i = 0; i < notesValues.length; i++) {
    const notes = notesValues[i][0];
    if (notes && typeof notes === 'string') {
      // 「メールID:xxx」の形式からIDを抽出
      const match = notes.match(/メールID[：:]\s*([a-zA-Z0-9]+)/);
      if (match && match[1]) {
        processedIds.add(match[1]);
      }
    }
  }
  
  return processedIds;
}

// ====== 手動実行用関数（デバッグ・テスト用） ======
function testProcessRepairEmails() {
  Logger.log('【手動実行】メール処理を開始します');
  processRepairEmails();
  Logger.log('【手動実行】メール処理が完了しました');
}

// ====== メール検索テスト関数（デバッグ用） ======
function testEmailSearch() {
  try {
    // パターン1: 元の検索条件
    const query1 = `subject:"${CONFIG.EMAIL_SUBJECT}" is:unread newer_than:1d`;
    Logger.log(`【テスト1】検索クエリ: ${query1}`);
    const threads1 = GmailApp.search(query1, 0, 10);
    Logger.log(`【テスト1】見つかったスレッド数: ${threads1.length}`);
    
    // パターン2: 未読条件を外す
    const query2 = `subject:"${CONFIG.EMAIL_SUBJECT}" newer_than:1d`;
    Logger.log(`【テスト2】検索クエリ: ${query2}`);
    const threads2 = GmailApp.search(query2, 0, 10);
    Logger.log(`【テスト2】見つかったスレッド数: ${threads2.length}`);
    
    // パターン3: 件名を部分一致に
    const query3 = `subject:修繕 newer_than:1d`;
    Logger.log(`【テスト3】検索クエリ: ${query3}`);
    const threads3 = GmailApp.search(query3, 0, 10);
    Logger.log(`【テスト3】見つかったスレッド数: ${threads3.length}`);
    
    // パターン4: 件名条件を外す（最近のメールを確認）
    const query4 = `newer_than:1d`;
    Logger.log(`【テスト4】検索クエリ: ${query4} (最近24時間のメール)`);
    const threads4 = GmailApp.search(query4, 0, 5);
    Logger.log(`【テスト4】見つかったスレッド数: ${threads4.length}`);
    
    // パターン4の結果を詳しく表示
    for (let i = 0; i < Math.min(threads4.length, 5); i++) {
      const thread = threads4[i];
      const messages = thread.getMessages();
      for (const message of messages) {
        const subject = message.getSubject();
        const date = message.getDate();
        const isUnread = message.isUnread();
        Logger.log(`【テスト4】メール${i + 1}: 件名="${subject}", 未読=${isUnread}, 日時=${date}`);
      }
    }
    
    // 見つかったメールの詳細を表示
    const allThreads = threads2.length > 0 ? threads2 : threads3.length > 0 ? threads3 : threads4;
    if (allThreads.length > 0) {
      Logger.log(`\n【詳細】見つかったメールの詳細:`);
      for (const thread of allThreads) {
        const messages = thread.getMessages();
        for (const message of messages) {
          Logger.log(`  件名: "${message.getSubject()}"`);
          Logger.log(`  未読: ${message.isUnread()}`);
          Logger.log(`  日時: ${message.getDate()}`);
          Logger.log(`  差出人: ${message.getFrom()}`);
          Logger.log(`  ---`);
        }
      }
    }
    
    return `テスト完了。パターン1: ${threads1.length}, パターン2: ${threads2.length}, パターン3: ${threads3.length}, パターン4: ${threads4.length}`;
  } catch (error) {
    Logger.log(`【テスト】エラー: ${error.toString()}`);
    Logger.log(`【テスト】スタック: ${error.stack}`);
    return `エラー: ${error.toString()}`;
  }
}

// ====== 強制メール処理関数（件名部分一致、処理済みチェックあり） ======
function forceProcessRepairEmails() {
  try {
    Logger.log('【強制実行】メール処理を開始します（件名部分一致、処理済みチェックあり）');
    
    // 検索条件を緩和：件名に「修繕」を含む、24時間以内
    const query = `subject:修繕 newer_than:1d`;
    Logger.log(`【強制実行】検索クエリ: ${query}`);
    
    const threads = GmailApp.search(query, 0, 10);
    Logger.log(`【強制実行】見つかったスレッド数: ${threads.length}`);
    
    if (threads.length === 0) {
      Logger.log('【強制実行】処理対象のメールがありません');
      return '処理対象のメールがありません';
    }
    
    // 処理済みメールIDのリストを取得
    const processedEmailIds = getProcessedEmailIds();
    Logger.log(`【強制実行】処理済みメールID数: ${processedEmailIds.size}`);
    
    let processedCount = 0;
    let skippedCount = 0;
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      Logger.log(`【強制実行】スレッド内のメッセージ数: ${messages.length}`);
      
      for (const message of messages) {
        const messageId = message.getId();
        const subject = message.getSubject();
        const date = message.getDate();
        
        // システムで処理済みかどうかをチェック
        if (processedEmailIds.has(messageId)) {
          Logger.log(`【強制実行】処理済みメールをスキップ: 件名="${subject}", ID=${messageId}`);
          skippedCount++;
          continue;
        }
        
        // 件名に「修繕依頼」が含まれているか確認
        if (subject.includes('修繕依頼') || subject === '修繕依頼') {
          Logger.log(`【強制実行】新規メールを処理: 件名="${subject}", ID=${messageId}, 日時=${date}`);
          processRepairEmail(message, messageId);
          processedCount++;
        } else {
          Logger.log(`【強制実行】件名が一致しないためスキップ: "${subject}"`);
          skippedCount++;
        }
      }
    }
    
    Logger.log(`【強制実行】処理完了: 新規処理=${processedCount}件, スキップ=${skippedCount}件`);
    return `処理完了: 新規処理=${processedCount}件, スキップ=${skippedCount}件`;
  } catch (error) {
    const errorMessage = `強制メール処理エラー: ${error.toString()}`;
    Logger.log(`【強制実行】エラー: ${error.toString()}`);
    Logger.log(`【強制実行】スタック: ${error.stack}`);
    // エラー通知は重複防止機能付きで送信
    sendErrorNotification(errorMessage);
    return `エラー: ${error.toString()}`;
  }
}

// ====== 過去30分以内のメールでChat通知未送信のものを処理 ======
function processRecentEmailsWithoutChat() {
  try {
    Logger.log('【最近のメール処理】過去30分以内でChat通知未送信のメールを処理します');
    
    // 検索条件：件名完全一致、過去30分以内
    const query = `subject:"${CONFIG.EMAIL_SUBJECT}" newer_than:30m`;
    Logger.log(`【最近のメール処理】検索クエリ: ${query}`);
    
    const threads = GmailApp.search(query, 0, 20);
    Logger.log(`【最近のメール処理】見つかったスレッド数: ${threads.length}`);
    
    if (threads.length === 0) {
      Logger.log('【最近のメール処理】処理対象のメールがありません');
      return '処理対象のメールがありません';
    }
    
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    
    // スプレッドシートからメールIDとChat通知送信状況を確認
    const emailStatusMap = new Map(); // messageId -> { hasChat: boolean, rowNum: number }
    
    if (lastRow > 1) {
      const notesRange = sheet.getRange(2, COL.NOTES + 1, lastRow - 1, 1);
      const ringiIdRange = sheet.getRange(2, COL.RINGI_ID + 1, lastRow - 1, 1);
      const notesValues = notesRange.getValues();
      const ringiIdValues = ringiIdRange.getValues();
      
      for (let i = 0; i < notesValues.length; i++) {
        const notes = notesValues[i][0] || '';
        const ringiId = ringiIdValues[i][0] || '';
        const rowNum = i + 2;
        
        // メールIDを抽出
        if (notes && typeof notes === 'string') {
          const emailIdMatch = notes.match(/メールID:([a-zA-Z0-9]+)/);
          if (emailIdMatch) {
            const messageId = emailIdMatch[1];
            // RINGI_ID（Docs URL）が存在すればChat通知も送られていると判断
            const hasChat = ringiId && ringiId.trim() !== '' && ringiId.startsWith('http');
            emailStatusMap.set(messageId, { hasChat, rowNum });
          }
        }
      }
    }
    
    let processedCount = 0;
    let skippedCount = 0;
    let chatSentCount = 0;
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      Logger.log(`【最近のメール処理】スレッド内のメッセージ数: ${messages.length}`);
      
      for (const message of messages) {
        const messageId = message.getId();
        const subject = message.getSubject();
        const date = message.getDate();
        
        // 件名が完全一致するか確認
        if (subject !== CONFIG.EMAIL_SUBJECT) {
          Logger.log(`【最近のメール処理】件名が一致しないためスキップ: "${subject}"`);
          skippedCount++;
          continue;
        }
        
        // Chat通知が送られているかチェック
        const status = emailStatusMap.get(messageId);
        if (status && status.hasChat) {
          Logger.log(`【最近のメール処理】Chat通知済みのためスキップ: 件名="${subject}", ID=${messageId}, 行${status.rowNum}`);
          chatSentCount++;
          skippedCount++;
          continue;
        }
        
        // Chat通知が送られていないメールを処理
        Logger.log(`【最近のメール処理】Chat通知未送信のメールを処理: 件名="${subject}", ID=${messageId}, 日時=${date}`);
        
        // 処理開始時に即座にメールIDを記録（重複処理を防ぐ）
        const tempRowNum = sheet.getLastRow() + 1;
        const tempNotesCell = sheet.getRange(tempRowNum, COL.NOTES + 1);
        tempNotesCell.setValue(`メールID:${messageId}`);
        Logger.log(`【最近のメール処理】処理開始時にメールIDを記録しました（行${tempRowNum}）`);
        
        try {
          processRepairEmail(message, messageId);
          processedCount++;
        } catch (error) {
          Logger.log(`【最近のメール処理】エラーが発生しましたが、メールIDは既に記録済み: ${error.toString()}`);
          throw error;
        }
      }
    }
    
    Logger.log(`【最近のメール処理】処理完了: 新規処理=${processedCount}件, Chat通知済み=${chatSentCount}件, スキップ=${skippedCount}件`);
    return `処理完了: 新規処理=${processedCount}件, Chat通知済み=${chatSentCount}件, スキップ=${skippedCount}件`;
  } catch (error) {
    const errorMessage = `最近のメール処理エラー: ${error.toString()}`;
    Logger.log(errorMessage);
    Logger.log(`【最近のメール処理】スタック: ${error.stack}`);
    sendErrorNotification(errorMessage);
    return `エラー: ${error.toString()}`;
  }
}

// ====== 強制メール処理（処理済みチェックを無視して全件再処理） ======
// 過去に失敗して「メールID」だけ記録されてしまったものも含めて再送したい場合に使用
function forceProcessAllEmailsIgnoreProcessed() {
  try {
    Logger.log('【強制実行(全件)】処理済みチェックを無視して再処理を開始します');
    
    // 件名完全一致・直近7日を対象（必要に応じて調整可）
    const query = `subject:"${CONFIG.EMAIL_SUBJECT}" newer_than:7d`;
    Logger.log(`【強制実行(全件)】検索クエリ: ${query}`);
    
    const threads = GmailApp.search(query, 0, 20); // 上限20件まで（必要に応じて調整）
    Logger.log(`【強制実行(全件)】見つかったスレッド数: ${threads.length}`);
    
    if (threads.length === 0) {
      Logger.log('【強制実行(全件)】処理対象のメールがありません');
      return '処理対象のメールがありません';
    }
    
    let processedCount = 0;
    let skippedCount = 0;
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      Logger.log(`【強制実行(全件)】スレッド内のメッセージ数: ${messages.length}`);
      
      for (const message of messages) {
        const messageId = message.getId();
        const subject = message.getSubject();
        const date = message.getDate();
        
        if (subject !== CONFIG.EMAIL_SUBJECT) {
          Logger.log(`【強制実行(全件)】件名が一致しないためスキップ: "${subject}"`);
          skippedCount++;
          continue;
        }
        
        Logger.log(`【強制実行(全件)】メールを再処理: 件名="${subject}", ID=${messageId}, 日時=${date}`);
        processRepairEmail(message, messageId);
        processedCount++;
      }
    }
    
    Logger.log(`【強制実行(全件)】処理完了: 再処理=${processedCount}件, スキップ=${skippedCount}件`);
    return `処理完了: 再処理=${processedCount}件, スキップ=${skippedCount}件`;
  } catch (error) {
    const errorMessage = `強制メール処理(全件)エラー: ${error.toString()}`;
    Logger.log(errorMessage);
    Logger.log(`【強制実行(全件)】スタック: ${error.stack}`);
    sendErrorNotification(errorMessage);
    return `エラー: ${error.toString()}`;
  }
}

// ====== 直近の行に対してChat下書きを再送する（AI再解析なし） ======
// 引数: count (省略時5) 直近の行から指定件数分、Docs URLがある行に対して下書きChatを再送
function forceResendRecentChats(count) {
  try {
    const resendCount = count || 5;
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('【再送】データ行がありません');
      return 'データ行がありません';
    }
    
    let sent = 0;
    for (let r = lastRow; r >= 2 && sent < resendCount; r--) {
      const repairId = sheet.getRange(r, COL.REPAIR_ID + 1).getValue();
      const docUrl = sheet.getRange(r, COL.RINGI_ID + 1).getValue();
      const area = sheet.getRange(r, COL.AREA + 1).getValue() || '';
      const location = sheet.getRange(r, COL.LOCATION_DETAIL + 1).getValue() || '';
      const title = `${area} ${location}`.trim() || '修繕案件';
      
      if (!repairId || !docUrl) {
        Logger.log(`【再送】Docs URLまたは修繕IDが空のためスキップ: 行${r}`);
        continue;
      }
      
      sendDraftNotification(repairId, docUrl, r, title);
      Logger.log(`【再送】Chat下書きを再送しました: 行${r}, 修繕ID=${repairId}`);
      sent++;
    }
    
    const msg = `Chat下書きの再送完了: ${sent}件`;
    Logger.log(`【再送】${msg}`);
    return msg;
  } catch (error) {
    const errorMessage = `Chat再送エラー: ${error.toString()}`;
    Logger.log(errorMessage);
    sendErrorNotification(errorMessage);
    return errorMessage;
  }
}

function processRepairEmail(message, messageId) {
  try {
    Logger.log(`【メール処理開始】メールID: ${messageId}`);
    
    const subject = message.getSubject();
    const body = message.getPlainBody();
    const from = message.getFrom();
    const date = message.getDate();
    
    Logger.log(`【メール処理】件名: ${subject}, 差出人: ${from}, 日時: ${date}`);
    
    // 報告者名を抽出
    const reporterName = extractReporterNameFromBody(body) || extractNameFromEmail(from) || '不明';
    Logger.log(`【メール処理】報告者名: ${reporterName}`);
    
    // 画像を取得
    const attachments = message.getAttachments();
    const images = [];
    for (const attachment of attachments) {
      if (attachment.getContentType().startsWith('image/')) {
        images.push({
          mimeType: attachment.getContentType(),
          data: Utilities.base64Encode(attachment.getBytes())
        });
        Logger.log(`【メール処理】画像添付: ${attachment.getName()}, タイプ: ${attachment.getContentType()}`);
      }
    }
    
    // Gemini AIで解析
    Logger.log('【メール処理】Gemini AIで解析を開始します');
    const aiResult = analyzeWithGemini(body, images, subject);
    Logger.log(`【メール処理】AI解析完了: ${aiResult.substring(0, 100)}...`);
    
    // AI結果をパース
    const repairId = generateRepairId();
    const rowData = parseAIResult(aiResult, repairId, reporterName, body, []);
    
    // スプレッドシートに書き込み
    const sheet = getSheet();
    const rowNum = sheet.getLastRow() + 1;
    Logger.log(`【メール処理】スプレッドシートに書き込み: 行${rowNum}`);
    writeRowToSheet(sheet, rowNum, rowData);
    
    // 画像をBlobに変換
    const imageBlobs = [];
    for (const attachment of attachments) {
      if (attachment.getContentType().startsWith('image/')) {
        imageBlobs.push(attachment.copyBlob());
      }
    }
    
    // Docsを生成
    Logger.log('【メール処理】Docsを生成します');
    const docUrl = createOrUpdateRingiDoc(rowData, repairId, imageBlobs);
    Logger.log(`【メール処理】Docs生成完了: ${docUrl}`);
    
    // スプレッドシートにDocs URLを記録
    sheet.getRange(rowNum, COL.RINGI_ID + 1).setValue(docUrl);
    
    // メールIDをNOTES列に記録
    const notesCell = sheet.getRange(rowNum, COL.NOTES + 1);
    const existingNotes = notesCell.getValue() || '';
    notesCell.setValue(existingNotes ? `${existingNotes}\nメールID:${messageId}` : `メールID:${messageId}`);
    
    // スプレッドシートへの変更を保存
    SpreadsheetApp.flush();
    
    // Chat通知を送信（下書き通知）
    const area = rowData[COL.AREA] || '';
    const location = rowData[COL.LOCATION_DETAIL] || '';
    const title = `${area} ${location}`.trim() || '修繕案件';
    Logger.log('【メール処理】Chat通知を送信します');
    sendDraftNotification(repairId, docUrl, rowNum, title);
    
    Logger.log(`【メール処理完了】修繕ID: ${repairId}, 行${rowNum}`);
  } catch (error) {
    Logger.log(`【メール処理エラー】メールID: ${messageId}, エラー: ${error.toString()}`);
    Logger.log(`【メール処理エラー】スタック: ${error.stack}`);
    throw error;
  }
}

// ====== Gemini AI解析 ======
function analyzeWithGemini(body, images, subject, retryNote) {
  try {
  const prompt = buildAnalysisPrompt(body, subject, retryNote);
    const models = [CONFIG.GEMINI_MODEL_PRIMARY, CONFIG.GEMINI_MODEL_FALLBACK];
    let lastError = null;
    
    for (const model of models) {
      try {
        Logger.log(`【Gemini】モデル試行: ${model}`);
        const rawResult = callGeminiAPI(model, prompt, images);
        const sanitizedResult = sanitizeAIResponse(rawResult);
        if (model === CONFIG.GEMINI_MODEL_FALLBACK && lastError) {
          Logger.log(`【Gemini】フォールバック成功（前回エラー: ${lastError.toString()})`);
        }
        return sanitizedResult;
      } catch (error) {
        Logger.log(`【Gemini】モデル失敗: ${model} -> ${error.toString()}`);
        lastError = error;
        // 次のモデルでリトライ
      }
    }
    
    // 全モデル失敗
    throw lastError || new Error('Gemini解析で不明なエラーが発生しました');
  } catch (error) {
    throw new Error(`Gemini解析エラー: ${error.toString()}`);
  }
}

// ====== AI応答のサニタイズ処理 ======
function sanitizeAIResponse(aiResult) {
  if (!aiResult || typeof aiResult !== 'string') {
    Logger.log('【警告】AI応答が空です');
    return '';
  }
  
  Logger.log('【デバッグ】AI応答（生）: ' + aiResult.substring(0, 200));
  
  // 1. Markdownコードブロックを除去
  let cleanText = aiResult.replace(/```[\s\S]*?```/g, "").replace(/```/g, "").trim();
  
  // 2. 改行を全て除去（1行にする）
  cleanText = cleanText.replace(/\r?\n/g, "").replace(/\r/g, "");
  
  // 3. Markdown表形式のヘッダー行（|---|など）を除去
  cleanText = cleanText.replace(/\|[\s-:]+\|/g, "");
  
  // 4. パイプを含む行を抽出（挨拶文の除去）
  const lines = cleanText.split(/[|]/);
  if (lines.length < 2) {
    // パイプがない場合は、元のテキストからパイプを含む部分を探す
    const pipeMatch = cleanText.match(/[^|]*\|[^|]*/);
    if (pipeMatch) {
      cleanText = pipeMatch[0];
    }
  }
  
  // 5. 先頭や末尾のパイプ削除
  cleanText = cleanText.replace(/^\|+/, "").replace(/\|+$/, "").trim();
  
  // 6. 挨拶や説明文を除去（「承知しました」「以下が...」など）
  cleanText = cleanText.replace(/^(承知|了解|以下|出力|結果|データ|回答|分析)[:：\s]*/i, "");
  cleanText = cleanText.replace(/^[^|]*?(?=\|)/, "");
  
  // 7. 連続するパイプを1つに統一
  cleanText = cleanText.replace(/\|+/g, "|");
  
  // 8. 前後の空白を除去
  cleanText = cleanText.trim();
  
  Logger.log('【デバッグ】AI応答（クリーニング後）: ' + cleanText.substring(0, 200));
  
  return cleanText;
}

// ====== URL検証（実在チェック） ======
function validateUrl(url) {
  if (!url || typeof url !== 'string' || url.trim() === '') return false;
  if (!url.startsWith('http')) return false; // http/httpsのみ許可
  if (url.includes('example.com') || url.includes('dummy') || url.includes('placeholder')) return false;
  
  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true, // エラーでも止まらないようにする
      followRedirects: true,     // リダイレクトを追う
      validateHttpsCertificates: false, // 証明書エラーを許容
      timeout: 3 // タイムアウトを3秒に短縮（全体タイムアウト対策）
    });
    const code = resp.getResponseCode();
    // 200〜399番台を許可（リダイレクトも許容）
    const isValid = code >= 200 && code < 400;
    if (!isValid) {
      Logger.log(`【URL検証】無効なステータスコード: ${url} -> ${code}`);
    }
    return isValid;
  } catch (e) {
    Logger.log(`【URL検証】エラーまたはタイムアウト: ${url} / ${e.toString()}`);
    return false;
  }
}

function sanitizeEstimateRecord(est, label) {
  if (!est) return { name: '', price: '', url: '', detail: '' };
  let name = est.name || '';
  const price = est.price || '';
  let url = est.url || '';
  let detail = '';
  
  // URLが空の場合は警告のみ（業者名と価格は保持）。URLなしで返す。
  if (!url || url.trim() === '') {
    Logger.log(`【警告】${label}: URLが空です（業者名と価格は保持）`);
    detail = name ? `${name}${price ? ` (${price}円)` : ''} - URL取得不可（AIが実在URLを返さず）` : 'URL取得不可（AIが実在URLを返さず）';
    return { name, price, url: '', detail };
  }
  
  // URL検証を実施
  const ok = validateUrl(url);
  if (!ok) {
    Logger.log(`【警告】${label}: URLが無効または実在確認NG（URLを削除、業者名と価格は保持） -> ${url}`);
    detail = name ? `${name}${price ? ` (${price}円)` : ''} - URL無効（実在確認NG）` : 'URL無効（実在確認NG）';
    url = ''; // 無効なURLは空文字に
  } else {
    // 有効なURLの場合
    if (!name) {
      // URLだけでも通すため、名前が空ならプレースホルダ名を付与
      name = '見積候補';
    }
    detail = name ? `${name}${price ? ` (${price}円)` : ''} - ${url}` : url;
  }
  
  return { name, price, url, detail };
}

function buildAnalysisPrompt(body, subject, retryNote) {
  const retryText = retryNote || '';
  return `あなたはホテル「mt. inn」（福島県二本松市岳温泉）の設備管理事務代行AIです。

現場から届いた写真とメールを基に、以下の23項目を調査・作成してください。

【場所指定と検索範囲】
- 福島県二本松市岳温泉 周辺の業者・通販を最優先に探してください。
- 近い順に拡大: 二本松市 → 福島市 → 郡山市 → 福島県内 → （次に）仙台・首都圏を同時に探索 → 全国EC（Amazon/楽天/MonotaRO等）。地域を広げた場合は理由を記載。
- 地域を広げた場合は、なぜその地域になったか理由を簡潔に書いてください。

【最重要：URLの実在性とハルシネーション完全禁止】
**絶対に守るべきルール：**
1. **業者名もURLも生成・推測・創作禁止。必ずGoogle検索ツールで実在確認できた情報のみ使用。**
2. **URLは「公式サイト」または「信頼できるポータルの該当ページ」（例: 公式サイト、Amazon/楽天/MonotaRO等）に限定。短縮URL・プレースホルダ・example.com等は禁止。**
3. **見積情報（見積A/B/C）は、URLが実在確認できた場合のみ「品名#金額#URL」形式で出力。URLが無い/不明/無効なら項目自体を空欄にする（出力しない）。**
4. **金額が不明なら捏造せず「見積もり要問い合わせ」と記載し、必ず実在URLを添付。**
5. **実在URLが1件も取れない場合、その見積は空欄のまま。捏造や推測URLは絶対に出力しない。**

【重要：Google検索の強制と実在URL取得】
1. **部材特定:** 写真から必要部材を特定し、Google検索ツールで購入可能な実在URL（httpsのみ）を取得。
2. **業者特定:** 地域優先順位に従いGoogle検索ツールで実在URLを取得（短縮URL・プレースホルダ禁止）。
3. **見積:** 実在URLが確認できたものだけ出力。URLが無い/不明/無効な場合は空欄。金額不明は「見積もり要問い合わせ」＋実在URL。
4. **徹底検索:** 実在URLが見つかるまで再検索。見つからない場合は捏造せず空欄。
5. **ハルシネーション完全禁止:** 曖昧・無効・推測URLは即破棄し、再検索して実在URLのみ出力。
${retryText ? `【再検索指示】${retryText}` : ''}

【出力フォーマット】
・以下の23項目をパイプライン「|」のみで区切り、**改行なしの1行**で出力してください。
・Markdownの表形式（|---|など）や、冒頭の挨拶（「承知しました」等）は**一切出力しない**でください。

1.エリア | 2.場所詳細 | 3.AI整形文 | 4.問題要約 | 5.原因分析 | 6.重要度(A/B/C) | 7.ランク理由 | 8.推奨対応 | 9.作業内容要約 | 10.作業手順 | 11.必要部材(URL付) | 12.作業時間 | 13.費用下限 | 14.費用上限 | 15.業者カテゴリ | 16.業者エリア | 17.検索KW | 18.先送りリスク | 19.見積A(品名#金#URL) | 20.見積B(品名#金#URL) | 21.見積C(品名#金#URL) | 22.選定業者 | 23.稟議理由

入力内容:
件名: ${subject}
本文:
${body}

必ずGoogle検索ツールを使用して、実在するURL（Amazon, MonotaRO, 楽天、メーカーサイト等）を取得してください。`;
}

function callGeminiAPI(model, prompt, images) {
  // 404回避のため 3.0系も v1beta で呼び出す
  const apiVersion = 'v1beta';
  const url = `https://generativelanguage.googleapis.com/${apiVersion}/models/${model}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
  
  const parts = [];
  
  // テキスト部分
  parts.push({ text: prompt });
  
  // 画像部分
  for (const img of images) {
    parts.push({
      inlineData: {
        mimeType: img.mimeType,
        data: img.data
      }
    });
  }
  
  const payload = {
    contents: [{
      parts: parts
    }],
    generationConfig: {
      temperature: 0.7,
      topK: 40,
      topP: 0.95,
      maxOutputTokens: 8192
    }
  };

  // v1beta（2.5系など）のみ Google検索ツールを付与。v1（3.0系）は付けると400になるため除外。
  if (apiVersion === 'v1beta') {
    payload.tools = [{
      googleSearch: {}
    }];
  }
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode !== 200) {
    const errorText = response.getContentText();
    throw new Error(`Gemini API error (${responseCode}): ${errorText}`);
  }
  
  const result = JSON.parse(response.getContentText());
  
  if (!result.candidates || !result.candidates[0] || !result.candidates[0].content) {
    throw new Error('Gemini API: 無効なレスポンス');
  }
  
  const candidate = result.candidates[0];
  const content = candidate.content;
  
  // Google検索ツールのfunctionCallが使用された場合の処理
  // Geminiは自動でfunctionCallを実行して結果を返すため、最終的なテキストレスポンスを取得
  if (content.parts) {
    let textResult = '';
    for (const part of content.parts) {
      if (part.text) {
        textResult += part.text;
      }
    }
    if (textResult) {
      return textResult;
    }
  }
  
  // フォールバック: テキストを直接取得
  if (content.parts && content.parts[0] && content.parts[0].text) {
    return content.parts[0].text;
  }
  
  throw new Error('Gemini API: テキストレスポンスが見つかりません');
}

// ====== AI結果のパース ======
function parseAIResult(aiText, repairId, reporterName, originalText, imageUrls) {
  if (!aiText || typeof aiText !== 'string') {
    Logger.log('【警告】AI応答が空です');
    aiText = '';
  }
  
  // 1. クリーニング: Markdown記号、改行、挨拶文を削除
  let cleanText = aiText.replace(/```.*?```/gs, "").replace(/\r?\n/g, "").trim();
  
  // 2. 先頭や末尾のパイプ削除
  if (cleanText.startsWith("|")) cleanText = cleanText.substring(1);
  if (cleanText.endsWith("|")) cleanText = cleanText.slice(0, -1);
  
  // 3. Markdown表形式のヘッダー行を除去
  cleanText = cleanText.replace(/\|[\s-:]+\|/g, "");
  
  // 4. 分割
  let parts = cleanText.split("|").map(s => s.trim());
  
  // 5. エラーハンドリング: 項目数が足りない場合
  if (parts.length < 23) {
    Logger.log(`【警告】AI回答の項目数が不足しています（${parts.length}/23）。補完します。`);
    Logger.log('【デバッグ】parts: ' + JSON.stringify(parts.slice(0, 5)));
    while (parts.length < 23) parts.push("");
  }
  
  // 6. 23項目を超える場合は最初の23項目のみを使用
  if (parts.length > 23) {
    Logger.log(`【警告】AI回答の項目数が超過しています（${parts.length}/23）。最初の23項目のみを使用します。`);
    parts = parts.slice(0, 23);
  }
  
  Logger.log(`【デバッグ】パース結果: ${parts.length}項目`);
  Logger.log(`【デバッグ】最初の3項目: ${parts[0] || '(空)'} | ${parts[1] || '(空)'} | ${parts[2] || '(空)'}`);
  
  // 報告者名は引数から直接使用
  const reporter = reporterName || '不明';
  const defaults = {
    area: '',
    location: '',
    formatted: originalText.substring(0, 500),
    problem: '',
    cause: '',
    rank: 'C',
    rankReason: '',
    type: 'SELF_SIMPLE',
    workSummary: '',
    steps: '',
    materials: '',
    time: '30',
    costMin: '0',
    costMax: '0',
    category: '',
    contractorArea: '',
    keywords: '',
    risk: '',
    selectedContractor: '',
    ringiReason: ''
  };
  
  const row = new Array(42).fill('');
  
  // A列〜AF列の正確なマッピング（0〜31列）
  // A: 修繕ID
  row[COL.REPAIR_ID] = repairId;
  // B: 受付日時
  row[COL.RECEIVED_DATETIME] = getTodayTokyo();
  // C: 報告者名
  row[COL.REPORTER_NAME] = reporter;
  // D: エリア (AI出力1 = parts[0])
  row[COL.AREA] = parts[0] || defaults.area;
  // E: 場所詳細 (AI出力2 = parts[1])
  row[COL.LOCATION_DETAIL] = parts[1] || defaults.location;
  // F: 写真1 (Drive URL)
  row[COL.PHOTO1] = imageUrls[0] || '';
  // G: 写真2 (Drive URL)
  row[COL.PHOTO2] = imageUrls[1] || '';
  // H: 写真3 (Drive URL)
  row[COL.PHOTO3] = imageUrls[2] || '';
  // I: 原文（全文を保存、最大50000文字まで）
  row[COL.ORIGINAL_TEXT] = originalText.length > 50000 ? originalText.substring(0, 50000) : originalText;
  // J: AI整形文 (AI出力3 = parts[2])
  row[COL.AI_FORMATTED] = parts[2] || defaults.formatted;
  // K: 問題要約 (AI出力4 = parts[3])
  row[COL.PROBLEM_SUMMARY] = parts[3] || defaults.problem;
  // L: 原因分析 (AI出力5 = parts[4])
  row[COL.CAUSE_ANALYSIS] = parts[4] || defaults.cause;
  // M: 重要度ランク (AI出力6 = parts[5])
  row[COL.PRIORITY_RANK] = parts[5] || defaults.rank;
  // N: ランク理由 (AI出力7 = parts[6])
  row[COL.RANK_REASON] = parts[6] || defaults.rankReason;
  // O: 推奨対応タイプ (AI出力8 = parts[7])
  row[COL.RECOMMENDED_TYPE] = parts[7] || defaults.type;
  // P: 作業内容要約 (AI出力9 = parts[8])
  row[COL.WORK_SUMMARY] = parts[8] || defaults.workSummary;
  // Q: 推奨作業手順 (AI出力10 = parts[9])
  row[COL.RECOMMENDED_STEPS] = parts[9] || defaults.steps;
  // R: 必要部材リスト (AI出力11 = parts[10])
  row[COL.MATERIALS_LIST] = parts[10] || defaults.materials;
  // S: 想定作業時間 (AI出力12 = parts[11])
  row[COL.ESTIMATED_TIME] = parts[11] || defaults.time;
  // T: AI概算費用下限 (AI出力13 = parts[12])
  row[COL.AI_COST_MIN] = parts[12] || defaults.costMin;
  // U: AI概算費用上限 (AI出力14 = parts[13])
  row[COL.AI_COST_MAX] = parts[13] || defaults.costMax;
  // V: 想定業者カテゴリ (AI出力15 = parts[14])
  row[COL.CONTRACTOR_CATEGORY] = parts[14] || defaults.category;
  // W: 想定業者エリア (AI出力16 = parts[15])
  row[COL.CONTRACTOR_AREA] = parts[15] || defaults.contractorArea;
  // X: 業者検索キーワード (AI出力17 = parts[16])
  row[COL.SEARCH_KEYWORDS] = parts[16] || defaults.keywords;
  // Y: 先送りリスク (AI出力18 = parts[17])
  row[COL.POSTPONE_RISK] = parts[17] || defaults.risk;
  
  // 見積A/B/Cをパース（品名#金額#URL形式）
  const parseEstimate = (estimateStr) => {
    if (!estimateStr || estimateStr === 'URLなし' || estimateStr.trim() === '') {
      return { name: '', price: '', url: '' };
    }
    const items = estimateStr.split('#');
    return {
      name: items[0] || '',
      price: items[1] || '',
      url: items[2] || ''
    };
  };
  
  // 見積A/B/Cをパースし、URL実在チェックを実施（無効・欠落時は理由をdetailに入れる）
  const estimateA = sanitizeEstimateRecord(parseEstimate(parts[18]), '見積A');
  const estimateB = sanitizeEstimateRecord(parseEstimate(parts[19]), '見積B');
  const estimateC = sanitizeEstimateRecord(parseEstimate(parts[20]), '見積C');

  // 見積もりはURLが実在確認できたもののみ出力（URLがない場合は空欄）
  // URLが無効/欠落のときは空欄にする（ハルシネーション防止）
  // Z: 見積1
  row[COL.ESTIMATE1] = estimateA.url || ''; // URLがない場合は空欄
  // AA: 見積2
  row[COL.ESTIMATE2] = estimateB.url || ''; // URLがない場合は空欄
  // AB: 見積3
  row[COL.ESTIMATE3] = estimateC.url || ''; // URLがない場合は空欄
  
  // AC: 選定業者 (AI出力22 = parts[21])
  row[COL.SELECTED_CONTRACTOR] = parts[21] || defaults.selectedContractor;
  
  // AD: 稟議起案者（報告者と同じ）
  row[COL.RINGI_INITIATOR] = reporter;
  
  // AE: 稟議要否（費用が発生するなら「要」）
  const costMax = parseInt(parts[13] || '0', 10);
  row[COL.RINGI_REQUIRED] = costMax > 0 ? '要' : '不要';
  
  // AF: 稟議理由 (AI出力23 = parts[22])
  row[COL.RINGI_REASON] = parts[22] || defaults.ringiReason;
  
  // 見積情報をrowDataに保存（Docs生成用）
  row['_estimateA'] = estimateA;
  row['_estimateB'] = estimateB;
  row['_estimateC'] = estimateC;
  
  // ステータス（Phase 1: 下書き）
  row[COL.STATUS] = '受付';
  
  return row;
}

// ====== スプレッドシートへの書き込み（HYPERLINK形式対応） ======
function writeRowToSheet(sheet, rowNum, rowData) {
  // 最初にSTATUS列（AG列）の入力規則を削除（確実に実行）
  try {
    const statusCell = sheet.getRange(rowNum, COL.STATUS + 1);
    statusCell.clearDataValidations();
    Logger.log(`【STATUS列】関数開始時に入力規則を削除しました（行${rowNum}、列${COL.STATUS + 1}）`);
  } catch (clearError) {
    Logger.log(`【警告】STATUS列の入力規則削除エラー（無視して続行）: ${clearError.toString()}`);
  }
  
  const values = [];
  const formulaMap = {}; // 列番号をキーに数式を保存
  
  for (let i = 0; i < 42; i++) {
    const value = rowData[i] || '';
    
    // 見積1-3列（Z, AA, AB列）はURLがある場合HYPERLINK形式に
    if ((i === COL.ESTIMATE1 || i === COL.ESTIMATE2 || i === COL.ESTIMATE3) && value && value.startsWith('http')) {
      const estimate = i === COL.ESTIMATE1 ? rowData['_estimateA'] : 
                      i === COL.ESTIMATE2 ? rowData['_estimateB'] : rowData['_estimateC'];
      const displayText = (estimate.name || estimate.price || 'リンク').replace(/"/g, '""'); // ダブルクォートをエスケープ
      const escapedUrl = value.replace(/"/g, '""');
      formulaMap[i] = `=HYPERLINK("${escapedUrl}", "${displayText}")`;
      values.push(''); // 値は空にして、後で数式を設定
    }
    // 必要部材リスト（R列）にURLがある場合もHYPERLINK形式に
    else if (i === COL.MATERIALS_LIST && value && value.includes('http')) {
      // 必要部材リスト内のURLを抽出してHYPERLINK化
      const urlMatch = value.match(/https?:\/\/[^\s#]+/);
      if (urlMatch) {
        const url = urlMatch[0];
        const materialName = value.split('#')[0] || '部材';
        const escapedUrl = url.replace(/"/g, '""');
        const displayText = materialName.replace(/"/g, '""');
        formulaMap[i] = `=HYPERLINK("${escapedUrl}", "${displayText}")`;
        values.push('');
      } else {
        values.push(value);
      }
    }
    // 写真列（F, G, H列）はURLがある場合HYPERLINK形式に
    else if ((i === COL.PHOTO1 || i === COL.PHOTO2 || i === COL.PHOTO3) && value && value.startsWith('http')) {
      const photoNum = i === COL.PHOTO1 ? '1' : i === COL.PHOTO2 ? '2' : '3';
      const escapedUrl = value.replace(/"/g, '""');
      formulaMap[i] = `=HYPERLINK("${escapedUrl}", "写真${photoNum}")`;
      values.push('');
    }
    // 稟議ID列（AJ列）はURLがある場合HYPERLINK形式に
    else if (i === COL.RINGI_ID && value && value.startsWith('http')) {
      const escapedUrl = value.replace(/"/g, '""');
      formulaMap[i] = `=HYPERLINK("${escapedUrl}", "稟議書")`;
      values.push('');
    }
    else {
      values.push(value);
    }
  }
  
  // STATUS列（AG列）の入力規則を事前に削除（一括書き込み前に実行）
  try {
    const statusCell = sheet.getRange(rowNum, COL.STATUS + 1);
    statusCell.clearDataValidations();
    Logger.log(`【STATUS列】入力規則を削除しました（行${rowNum}、列${COL.STATUS + 1}）`);
  } catch (clearError) {
    Logger.log(`【警告】STATUS列の入力規則削除エラー（無視して続行）: ${clearError.toString()}`);
  }
  
  // 値を書き込み（入力規則エラーは無視）
  try {
    sheet.getRange(rowNum, 1, 1, 42).setValues([values]);
    Logger.log(`【スプレッドシート書き込み】42列の値を書き込みました（行${rowNum}）`);
  } catch (error) {
    // 入力規則エラーなどは無視して、可能な限り書き込む
    Logger.log(`【警告】一括書き込みでエラー（個別書き込みに切り替え）: ${error.toString()}`);
    
    // 個別に書き込む（エラーが発生した列はスキップ）
    for (let i = 0; i < 42; i++) {
      try {
        const cell = sheet.getRange(rowNum, i + 1);
        
        // STATUS列（AG列、COL.STATUS = 32）の場合は入力規則を削除してから値を設定
        if (i === COL.STATUS) {
          try {
            cell.clearDataValidations();
          } catch (clearError) {
            Logger.log(`【警告】STATUS列の入力規則削除エラー（無視）: ${clearError.toString()}`);
          }
        }
        
        if (formulaMap[i]) {
          // HYPERLINK数式がある場合は数式を書き込み
          cell.setFormula(formulaMap[i]);
        } else {
          // 通常の値を書き込み
          cell.setValue(values[i]);
        }
      } catch (colError) {
        // STATUS列の場合は、入力規則エラーを無視して再試行
        if (i === COL.STATUS) {
          try {
            const cell = sheet.getRange(rowNum, i + 1);
            cell.clearDataValidations();
            cell.setValue(values[i]);
            Logger.log(`【警告】STATUS列の書き込みを再試行（入力規則削除後）: 成功`);
          } catch (retryError) {
            Logger.log(`【警告】列${i + 1}（${String.fromCharCode(65 + i)}列）の書き込みをスキップ: ${retryError.toString()}`);
          }
        } else {
          Logger.log(`【警告】列${i + 1}（${String.fromCharCode(65 + i)}列）の書き込みをスキップ: ${colError.toString()}`);
        }
      }
    }
  }
  
  // HYPERLINK数式を書き込み（まだ書き込まれていない場合）
  for (const colIndex in formulaMap) {
    try {
      const cell = sheet.getRange(rowNum, parseInt(colIndex) + 1);
      // 既に数式が設定されているかチェック
      const currentFormula = cell.getFormula();
      if (!currentFormula || !currentFormula.startsWith('=HYPERLINK')) {
        cell.setFormula(formulaMap[colIndex]);
      }
    } catch (formulaError) {
      Logger.log(`【警告】列${parseInt(colIndex) + 1}のHYPERLINK数式書き込みをスキップ: ${formulaError.toString()}`);
    }
  }
}

// ====== 報告者名をメール本文の一番上から抽出 ======
function extractReporterNameFromBody(body) {
  if (!body || typeof body !== 'string') {
    return '不明';
  }
  
  // メール本文の最初の行（改行前）を取得
  const firstLine = body.split(/\r?\n/)[0].trim();
  
  // 空行や特殊文字を除去
  let reporterName = firstLine
    .replace(/^[:\-=\s]+/, '') // 先頭の記号を除去
    .replace(/[:\-=\s]+$/, '') // 末尾の記号を除去
    .trim();
  
  // 報告者名が取得できた場合
  if (reporterName && reporterName.length > 0 && reporterName.length < 50) {
    Logger.log(`【報告者名抽出】メール本文から抽出: "${reporterName}"`);
    return reporterName;
  }
  
  // 取得できない場合は「不明」を返す
  Logger.log(`【報告者名抽出】抽出失敗、デフォルト値「不明」を使用`);
  return '不明';
}

function extractNameFromEmail(email) {
  const match = email.match(/^(.+?)\s*<.+>$/);
  if (match) return match[1];
  const atIndex = email.indexOf('@');
  if (atIndex > 0) return email.substring(0, atIndex);
  return email;
}

// ====== Docs生成 ======
function createOrUpdateRingiDoc(rowData, repairId, imageBlobs) {
  try {
    const templateDoc = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const newDoc = templateDoc.makeCopy(`${repairId}_稟議書`, folder);
    const doc = DocumentApp.openById(newDoc.getId());
    const body = doc.getBody();
    
    // テンプレート置換
    const replacements = {
      '{{修繕ID}}': rowData[COL.REPAIR_ID] || '',
      '{{受付日時}}': rowData[COL.RECEIVED_DATETIME] || '',
      '{{報告者名}}': rowData[COL.REPORTER_NAME] || '',
      '{{エリア}}': rowData[COL.AREA] || '',
      '{{場所詳細}}': rowData[COL.LOCATION_DETAIL] || '',
      '{{AI整形文}}': rowData[COL.AI_FORMATTED] || '',
      '{{問題要約}}': rowData[COL.PROBLEM_SUMMARY] || '',
      '{{原因分析}}': rowData[COL.CAUSE_ANALYSIS] || '',
      '{{重要度ランク}}': rowData[COL.PRIORITY_RANK] || '',
      '{{ランク理由}}': rowData[COL.RANK_REASON] || '',
      '{{推奨対応タイプ}}': rowData[COL.RECOMMENDED_TYPE] || '',
      '{{作業内容要約}}': rowData[COL.WORK_SUMMARY] || '',
      '{{推奨作業手順}}': rowData[COL.RECOMMENDED_STEPS] || '',
      '{{必要部材リスト}}': rowData[COL.MATERIALS_LIST] || '',
      '{{想定作業時間}}': rowData[COL.ESTIMATED_TIME] || '',
      '{{AI概算費用下限}}': rowData[COL.AI_COST_MIN] || '',
      '{{AI概算費用上限}}': rowData[COL.AI_COST_MAX] || '',
      '{{想定業者カテゴリ}}': rowData[COL.CONTRACTOR_CATEGORY] || '',
      '{{想定業者エリア}}': rowData[COL.CONTRACTOR_AREA] || '',
      '{{業者検索キーワード}}': rowData[COL.SEARCH_KEYWORDS] || '',
      '{{先送りリスク}}': rowData[COL.POSTPONE_RISK] || ''
    };
    
    // 見積もり情報の置換
    const estimateA = rowData['_estimateA'] || {};
    const estimateB = rowData['_estimateB'] || {};
    const estimateC = rowData['_estimateC'] || {};
    
    const formatEstimate = (est) => {
      // detailフィールドがあればそれを使用、なければ従来の形式
      if (est.detail) {
        return est.detail;
      }
      if (!est.url) {
        return est.name ? `${est.name}${est.price ? ` (${est.price}円)` : ''} - URLなし` : 'URLなし';
      }
      return `${est.name || '部材'} (${est.price || '価格未定'}円) - ${est.url}`;
    };
    
    replacements['{{見積1}}'] = formatEstimate(estimateA);
    replacements['{{見積2}}'] = formatEstimate(estimateB);
    replacements['{{見積3}}'] = formatEstimate(estimateC);
    
    // テキスト置換
    for (const [key, value] of Object.entries(replacements)) {
      body.replaceText(key, value);
    }
    
    // 画像を埋め込み（{{写真1}}などのプレースホルダーを探して置換）
    const photoPlaceholders = [
      { pattern: '{{写真1}}', index: 0 },
      { pattern: '{{写真2}}', index: 1 },
      { pattern: '{{写真3}}', index: 2 },
      { pattern: '{{写真エリア}}', index: 0 }
    ];
    
    for (const placeholder of photoPlaceholders) {
      const elements = body.findText(placeholder.pattern);
      if (elements && imageBlobs && imageBlobs[placeholder.index]) {
        const element = elements.getElement();
        const parent = element.getParent();
        const index = parent.getChildIndex(element);
        
        // プレースホルダーを削除
        parent.removeChild(element);
        
        // 画像を挿入
        const imageBlob = imageBlobs[placeholder.index];
        const insertedImage = parent.insertInlineImage(index, imageBlob);
        insertedImage.setWidth(400); // 幅400pxにリサイズ
      }
    }
    
    // 画像が残っている場合は文書の末尾に追加
    if (imageBlobs && imageBlobs.length > 0) {
      for (let i = 0; i < imageBlobs.length; i++) {
        const imageBlob = imageBlobs[i];
        const insertedImage = body.appendImage(imageBlob);
        insertedImage.setWidth(400);
        body.appendParagraph(''); // 改行
      }
    }
    
    doc.saveAndClose();
    return newDoc.getUrl();
  } catch (error) {
    Logger.log(`Docs生成エラー: ${error.toString()}`);
    throw error;
  }
}

// ====== Chat通知（Cards V2） ======
// Phase 1: 修繕報告（下書き）通知
function sendDraftNotification(repairId, docUrl, rowNum, title) {
  const applyUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=apply&row=${rowNum}`;
  
  // 見出しにIDとタイトルを表示
  const headerTitle = title ? `${repairId} - ${title}` : repairId;
  
  const card = {
    cardsV2: {
      cardId: 'draft_notification',
      card: {
        header: {
          title: '修繕報告（下書き）が作成されました',
          subtitle: headerTitle
        },
        sections: [{
          widgets: [
            {
              textParagraph: {
                text: `修繕ID: <b>${repairId}</b><br/>修繕報告書（下書き）を確認して、正式申請してください。`
              }
            },
            {
              buttonList: {
                buttons: [
                  {
                    text: '📝 Docs確認・修正',
                    onClick: {
                      openLink: {
                        url: docUrl
                      }
                    },
                    icon: {
                      knownIcon: 'DESCRIPTION'
                    }
                  },
                  {
                    text: '✅ 正式申請する',
                    onClick: {
                      openLink: {
                        url: applyUrl
                      }
                    },
                    icon: {
                      knownIcon: 'CHECK_CIRCLE'
                    }
                  }
                ]
              }
            }
          ]
        }]
      }
    }
  };
  
  sendChatMessage(card);
}

function sendApprovalRequest(repairId, docUrl, rowNum, approverType) {
  const approverName = approverType === 'GM' ? 'GM（村松さん）' : '代表（鈴木安太郎さん）';
  const approveUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=approve&row=${rowNum}&type=${approverType}`;
  const rejectUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=reject&row=${rowNum}&type=${approverType}`;
  const commentUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=comment&row=${rowNum}&type=${approverType}`;
  
  // デバッグログ
  Logger.log(`【承認リクエスト送信】修繕ID: ${repairId}, 行: ${rowNum}, 承認者: ${approverType}`);
  Logger.log(`【承認リクエスト送信】承認URL: ${approveUrl}`);
  Logger.log(`【承認リクエスト送信】否決URL: ${rejectUrl}`);
  Logger.log(`【承認リクエスト送信】コメントURL: ${commentUrl}`);
  
  const card = {
    cardsV2: {
      cardId: 'approval_request',
      card: {
        header: {
          title: '修繕稟議の承認依頼',
          subtitle: repairId
        },
        sections: [{
          widgets: [
            {
              textParagraph: {
                text: `修繕ID: <b>${repairId}</b><br/>承認をお願いします。`
              }
            },
            {
              buttonList: {
                buttons: [
                  {
                    text: '承認する',
                    onClick: {
                      openLink: {
                        url: approveUrl
                      }
                    }
                  },
                  {
                    text: '否決する',
                    onClick: {
                      openLink: {
                        url: rejectUrl
                      }
                    }
                  },
                  {
                    text: 'コメント',
                    onClick: {
                      openLink: {
                        url: commentUrl
                      }
                    }
                  }
                ]
              }
            }
          ]
        }]
      }
    }
  };
  
  sendChatMessage(card);
}

// 承認/否決完了時のChat通知
function sendApprovalCompleteNotification(repairId, docUrl, result) {
  const resultText = result === '承認' ? '承認' : result === 'GM承認' ? 'GM承認' : '否決';
  const resultColor = result === '否決' ? '#f44336' : '#4caf50';
  
  const card = {
    cardsV2: {
      cardId: 'approval_complete',
      card: {
        header: {
          title: `修繕稟議が${resultText}されました`,
          subtitle: repairId
        },
        sections: [{
          widgets: [
            {
              textParagraph: {
                text: `修繕ID: <b>${repairId}</b><br/>${resultText}が完了しました。`
              }
            },
            {
              buttonList: {
                buttons: [
                  {
                    text: '📝 Docs確認',
                    onClick: {
                      openLink: {
                        url: docUrl
                      }
                    },
                    icon: {
                      knownIcon: 'DESCRIPTION'
                    }
                  }
                ]
              }
            }
          ]
        }]
      }
    }
  };
  
  sendChatMessage(card);
}

function sendChatMessage(card) {
  try {
    Logger.log('【Chat送信】Chat通知を送信します');
    
    if (!CONFIG.WEBHOOK_URL) {
      Logger.log('【Chat送信エラー】WEBHOOK_URLが設定されていません');
      return;
    }
    
    const payload = JSON.stringify(card);
    Logger.log(`【Chat送信】ペイロード: ${payload.substring(0, 200)}...`);
    
    const response = UrlFetchApp.fetch(CONFIG.WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      payload: payload,
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode >= 200 && responseCode < 300) {
      Logger.log(`【Chat送信成功】レスポンスコード: ${responseCode}`);
    } else {
      Logger.log(`【Chat送信エラー】レスポンスコード: ${responseCode}, レスポンス: ${responseText}`);
    }
  } catch (error) {
    Logger.log(`【Chat送信エラー】例外: ${error.toString()}`);
    Logger.log(`【Chat送信エラー】スタック: ${error.stack}`);
  }
}

// ====== エラー通知（Chat通知は無効化） ======
function sendErrorNotification(message) {
  // エラーログのみ記録（Chat通知は送信しない）
  Logger.log(`【エラー】${message}`);
  Logger.log(`【エラー通知】Chat通知は無効です（ログのみ記録）`);
  
  // Chat通知は完全に無効化（以下のコードは実行されません）
  // 将来的に有効化する場合は、CONFIG.ENABLE_ERROR_NOTIFICATIONをtrueに変更し、
  // 以下のコメントを解除してください。
  
  /*
  try {
    // エラーメッセージのハッシュを作成（重複チェック用）
    const errorHash = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      message,
      Utilities.Charset.UTF_8
    );
    const errorHashStr = errorHash.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
    
    // プロパティストアから最後の通知時刻を取得
    const properties = PropertiesService.getScriptProperties();
    const lastNotificationKey = `error_notification_${errorHashStr}`;
    const lastNotificationTime = properties.getProperty(lastNotificationKey);
    const now = new Date().getTime();
    
    // 同じエラーを1時間以内に再通知しない
    const NOTIFICATION_INTERVAL_MS = 60 * 60 * 1000; // 1時間
    if (lastNotificationTime) {
      const timeSinceLastNotification = now - parseInt(lastNotificationTime, 10);
      if (timeSinceLastNotification < NOTIFICATION_INTERVAL_MS) {
        Logger.log(`【エラー通知】重複を防止: ${message.substring(0, 50)}... (${Math.floor(timeSinceLastNotification / 1000 / 60)}分前にも通知済み)`);
        return; // 通知をスキップ
      }
    }
    
    // エラー通知を送信
    const card = {
      cardsV2: {
        cardId: 'error_notification',
        card: {
          header: {
            title: 'エラー通知',
            subtitle: '修繕システム'
          },
          sections: [{
            widgets: [{
              textParagraph: {
                text: message
              }
            }]
          }]
        }
      }
    };
    
    sendChatMessage(card);
    
    // 通知時刻を記録
    properties.setProperty(lastNotificationKey, now.toString());
    
    Logger.log(`【エラー通知】Chatに送信しました: ${message.substring(0, 50)}...`);
  } catch (error) {
    // エラー通知自体が失敗した場合はログのみ
    Logger.log(`【エラー通知】送信失敗: ${error.toString()}`);
    Logger.log(`【エラー通知】元のメッセージ: ${message}`);
  }
  */
}

// ====== Webアプリ（doPost：コメント送信） ======
function doPost(e) {
  try {
    // === PWAからの修繕報告（JSON POST）===
    // CORSプリフライト回避のため text/plain でも受付可能
    if (e.postData && e.postData.contents) {
      try {
        const jsonData = JSON.parse(e.postData.contents);
        if (jsonData.action === 'repair_report') {
          Logger.log('【PWA】修繕報告を受信しました');
          const result = processRepairFromPWA(jsonData);
          return ContentService.createTextOutput(JSON.stringify(result))
            .setMimeType(ContentService.MimeType.JSON);
        }
      } catch (jsonError) {
        // JSONパースに失敗した場合は既存のコメント処理にフォールスルー
        Logger.log(`【PWA】JSONパース失敗（コメント処理にフォールスルー）: ${jsonError.toString()}`);
      }
    }

    // === 既存のコメント処理 ===
    // POSTデータからパラメータを取得
    let action = '';
    let row = 0;
    let type = '';
    let comment = '';

    // e.parameterから取得を試行
    if (e.parameter) {
      action = e.parameter.action || '';
      row = parseInt(e.parameter.row || '0', 10);
      type = e.parameter.type || '';
      comment = e.parameter.comment || '';
    }

    // e.postDataからも取得を試行（念のため）
    if (!comment && e.postData && e.postData.contents) {
      try {
        const postData = JSON.parse(e.postData.contents);
        if (postData.comment) {
          comment = postData.comment;
        }
      } catch (parseError) {
        // JSONパースに失敗した場合は、URLエンコードされた形式を試行
        const contents = e.postData.contents || '';
        const commentMatch = contents.match(/comment=([^&]*)/);
        if (commentMatch) {
          comment = decodeURIComponent(commentMatch[1].replace(/\+/g, ' '));
        }
      }
    }

    if (!row || row < 2) {
      return HtmlService.createHtmlOutput('<html><body><h1>エラー: 無効な行番号</h1></body></html>');
    }

    if (action === 'comment') {
      const sheet = getSheet();
      // e.parameterにcommentを設定（handleComment関数で使用）
      if (!e.parameter) {
        e.parameter = {};
      }
      e.parameter.comment = comment;
      return handleComment(sheet, row, type, e);
    }

    return HtmlService.createHtmlOutput('<html><body><h1>エラー: 無効なアクション</h1></body></html>');
  } catch (error) {
    Logger.log(`doPostエラー: ${error.toString()}`);
    // PWA JSONリクエストの場合はJSONでエラーを返す
    if (e.postData && e.postData.type === 'application/json') {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
    return HtmlService.createHtmlOutput(`<html><body><h1>エラー</h1><p>${error.toString()}</p></body></html>`);
  }
}

// ====== PWAからの修繕報告処理 ======
function processRepairFromPWA(data) {
  try {
    Logger.log('【PWA処理開始】修繕報告をPWAから処理します');

    const reporterName = data.reporter || '不明';
    const description = data.description || '';
    const base64Images = data.images || [];

    Logger.log(`【PWA処理】報告者: ${reporterName}, 画像数: ${base64Images.length}`);

    // base64画像をBlobに変換してDriveにアップロード
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const imageUrls = [];
    const imageBlobs = [];
    const geminiImages = [];

    for (let i = 0; i < Math.min(base64Images.length, 3); i++) {
      try {
        const imgData = base64Images[i];
        // data:image/jpeg;base64,xxxxx の形式からbase64部分を取得
        const base64Match = imgData.match(/^data:(image\/\w+);base64,(.+)$/);
        if (!base64Match) {
          Logger.log(`【PWA処理】画像${i + 1}: 無効なbase64形式`);
          continue;
        }

        const mimeType = base64Match[1];
        const base64Content = base64Match[2];
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Content), mimeType, `repair_${Date.now()}_${i + 1}.jpg`);

        // Driveにアップロード
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileUrl = file.getUrl();
        imageUrls.push(fileUrl);
        imageBlobs.push(blob);

        // Gemini用の画像データ
        geminiImages.push({
          mimeType: mimeType,
          data: base64Content
        });

        Logger.log(`【PWA処理】画像${i + 1}をDriveにアップロード: ${fileUrl}`);
      } catch (imgError) {
        Logger.log(`【PWA処理】画像${i + 1}のアップロードエラー: ${imgError.toString()}`);
      }
    }

    // Gemini AIで解析
    Logger.log('【PWA処理】Gemini AIで解析を開始します');
    const aiResult = analyzeWithGemini(description, geminiImages, '修繕依頼');
    Logger.log(`【PWA処理】AI解析完了: ${aiResult.substring(0, 100)}...`);

    // AI結果をパース
    const repairId = generateRepairId();
    const rowData = parseAIResult(aiResult, repairId, reporterName, description, imageUrls);

    // スプレッドシートに書き込み
    const sheet = getSheet();
    const rowNum = sheet.getLastRow() + 1;
    Logger.log(`【PWA処理】スプレッドシートに書き込み: 行${rowNum}`);
    writeRowToSheet(sheet, rowNum, rowData);

    // Docsを生成
    Logger.log('【PWA処理】稟議書Docsを生成します');
    const docUrl = createOrUpdateRingiDoc(rowData, repairId, imageBlobs);
    Logger.log(`【PWA処理】Docs生成完了: ${docUrl}`);

    // スプレッドシートにDocs URLを記録
    sheet.getRange(rowNum, COL.RINGI_ID + 1).setValue(docUrl);

    // 備考にPWAからの報告であることを記録
    const notesCell = sheet.getRange(rowNum, COL.NOTES + 1);
    notesCell.setValue(`PWA報告 ${getTodayTokyo()}`);

    SpreadsheetApp.flush();

    // Chat通知を送信（下書き通知）
    const area = rowData[COL.AREA] || '';
    const location = rowData[COL.LOCATION_DETAIL] || '';
    const title = `${area} ${location}`.trim() || '修繕案件';
    Logger.log('【PWA処理】Chat通知を送信します');
    sendDraftNotification(repairId, docUrl, rowNum, title);

    Logger.log(`【PWA処理完了】修繕ID: ${repairId}, 行${rowNum}`);

    return {
      success: true,
      repairId: repairId,
      message: `修繕報告を受け付けました（${repairId}）`
    };
  } catch (error) {
    Logger.log(`【PWA処理エラー】${error.toString()}`);
    Logger.log(`【PWA処理エラー】スタック: ${error.stack}`);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ====== Webアプリ（doGet） ======
function doGet(e) {
  try {
    Logger.log('=== doGet関数開始 ===');
    Logger.log(`パラメータ: ${e ? JSON.stringify(e.parameter) : 'なし'}`);
    Logger.log(`クエリ文字列: ${e && e.queryString ? e.queryString : 'なし'}`);
    Logger.log(`eオブジェクト: ${e ? '存在' : 'なし'}`);
    Logger.log(`e.parameter: ${e && e.parameter ? JSON.stringify(e.parameter) : 'なし'}`);
    
    // パラメータを取得（安全に）
    let action = '';
    let row = 0;
    let type = '';
    let comment = '';
    
    try {
      if (e && e.parameter) {
        action = String(e.parameter.action || '').trim();
        row = parseInt(String(e.parameter.row || '0'), 10);
        type = String(e.parameter.type || '').trim();
        comment = String(e.parameter.comment || '').trim();
      } else if (e && e.queryString) {
        // クエリ文字列から直接パース（フォールバック）
        const params = new URLSearchParams(e.queryString);
        action = params.get('action') || '';
        row = parseInt(params.get('row') || '0', 10);
        type = params.get('type') || '';
        comment = params.get('comment') || '';
        Logger.log(`【フォールバック】クエリ文字列から取得: action="${action}", row=${row}, type="${type}"`);
      }
    } catch (paramError) {
      Logger.log(`パラメータ取得エラー: ${paramError.toString()}`);
      Logger.log(`エラースタック: ${paramError.stack}`);
    }
    
    Logger.log(`取得したパラメータ: action="${action}", row=${row}, type="${type}", comment=${comment ? 'あり' : 'なし'}`);
    
    // 正式申請の場合は、承認リクエストを送信して成功画面を返す
    if (action === 'apply') {
      Logger.log('【正式申請】正式申請アクションを検知しました');
      Logger.log('【正式申請】承認リクエストを送信します。');
      
      // 行番号のチェック
      if (!row || row < 2) {
        Logger.log(`エラー: 無効な行番号 (${row})`);
        const errorHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>エラー</title></head><body style="font-family:sans-serif;padding:20px;"><h1>エラー: 無効な行番号</h1><p>行番号が指定されていません。</p></body></html>';
        return HtmlService.createHtmlOutput(errorHtml);
      }
      
      try {
        const sheet = getSheet();
        const repairId = sheet.getRange(row, COL.REPAIR_ID + 1).getValue();
        const docUrl = getDocUrlFromRow(sheet, row);
        
        // 金額を確認（AI概算費用上限または実際費用）
        const aiCostMax = sheet.getRange(row, COL.AI_COST_MAX + 1).getValue();
        const actualCost = sheet.getRange(row, COL.ACTUAL_COST + 1).getValue();
        const cost = actualCost || aiCostMax || 0;
        
        // 金額に応じて承認者を決定（閾値は50000円と仮定、必要に応じて調整）
        const GM_APPROVAL_THRESHOLD = 50000;
        const approverType = cost >= GM_APPROVAL_THRESHOLD ? 'GM' : '代表';
        
        Logger.log(`【正式申請】修繕ID: ${repairId}, 金額: ${cost}円, 承認者: ${approverType}`);
        Logger.log(`【正式申請】Docs URL: ${docUrl}`);
        
        // 承認リクエストを送信
        Logger.log('【正式申請】sendApprovalRequest関数を呼び出します');
        sendApprovalRequest(repairId, docUrl, row, approverType);
        Logger.log('【正式申請】sendApprovalRequest関数の呼び出し完了');
        
        // ステータスを「承認待ち」に更新
        try {
          const statusCell = sheet.getRange(row, COL.STATUS + 1);
          statusCell.clearDataValidations();
          statusCell.setValue('承認待ち');
          sheet.getRange(row, COL.RINGI_STATUS + 1).setValue('承認依頼中');
          SpreadsheetApp.flush();
        } catch (e) {
          Logger.log(`【正式申請】ステータス更新エラー（無視）: ${e.toString()}`);
        }
        
        Logger.log(`【正式申請】承認リクエスト送信完了: ${approverType}`);
        const html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>完了</title></head><body style="font-family:sans-serif;text-align:center;padding:50px;"><h1 style="color:#4caf50;">✓ 完了</h1><p>正式申請を受け付けました。</p><p>承認リクエストを送信しました。</p></body></html>';
        return HtmlService.createHtmlOutput(html);
      } catch (error) {
        Logger.log(`【正式申請】エラー: ${error.toString()}`);
        const errorHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>エラー</title></head><body style="font-family:sans-serif;padding:20px;"><h1>エラー</h1><p>' + error.toString() + '</p></body></html>';
        return HtmlService.createHtmlOutput(errorHtml);
      }
    }
    
    // 行番号のチェック
    if (!row || row < 2) {
      Logger.log(`エラー: 無効な行番号 (${row})`);
      const errorHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>エラー</title></head><body style="font-family:sans-serif;padding:20px;"><h1>エラー: 無効な行番号</h1><p>行番号が指定されていません。</p><p>行番号: ' + row + '</p></body></html>';
      return HtmlService.createHtmlOutput(errorHtml);
    }
    
    Logger.log(`行番号OK: ${row}`);
    
    // スプレッドシートを取得
    let sheet;
    try {
      sheet = getSheet();
      Logger.log('スプレッドシート取得成功');
    } catch (sheetError) {
      Logger.log(`スプレッドシート取得エラー: ${sheetError.toString()}`);
      const errorHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>エラー</title></head><body style="font-family:sans-serif;padding:20px;"><h1>エラー: スプレッドシート取得失敗</h1><p>' + sheetError.toString() + '</p></body></html>';
      return HtmlService.createHtmlOutput(errorHtml);
    }
    
    // アクションに応じて処理
    Logger.log(`アクション処理開始: "${action}"`);
    Logger.log(`行番号: ${row}, タイプ: "${type}"`);
    
    // アクションが空の場合はエラー
    if (!action || action === '') {
      Logger.log(`【エラー】アクションが空です`);
      const errorHtml = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="UTF-8">
          <title>エラー</title>
          <style>
            body { font-family: sans-serif; padding: 20px; background: #fff3cd; }
            h1 { color: #856404; }
            pre { background: #f5f5f5; padding: 10px; border-radius: 4px; }
          </style>
        </head>
        <body>
          <h1>エラー: アクションが指定されていません</h1>
          <p>URLにactionパラメータが含まれていません。</p>
          <p><strong>現在のパラメータ:</strong></p>
          <pre>${JSON.stringify(e ? e.parameter : {}, null, 2)}</pre>
          <p><strong>クエリ文字列:</strong> ${e && e.queryString ? e.queryString : 'なし'}</p>
          <p><strong>正しいURL形式:</strong></p>
          <ul>
            <li>承認: ${CONFIG.SCRIPT_WEB_APP_URL}?action=approve&row=2&type=GM</li>
            <li>否決: ${CONFIG.SCRIPT_WEB_APP_URL}?action=reject&row=2&type=GM</li>
            <li>コメント: ${CONFIG.SCRIPT_WEB_APP_URL}?action=comment&row=2&type=GM</li>
          </ul>
        </body>
        </html>
      `;
      return HtmlService.createHtmlOutput(errorHtml);
    }
    
    let result;
    try {
      switch (action) {
        case 'approve':
          Logger.log(`【承認処理開始】行${row}, タイプ: ${type}`);
          result = handleApprove(sheet, row, type);
          Logger.log('【承認処理完了】');
          return result;
          
        case 'reject':
          Logger.log(`【否決処理開始】行${row}, タイプ: ${type}`);
          result = handleReject(sheet, row, type);
          Logger.log('【否決処理完了】');
          return result;
          
        case 'comment':
          Logger.log(`【コメント処理開始】行${row}, タイプ: ${type}`);
          result = handleComment(sheet, row, type, e);
          Logger.log('【コメント処理完了】');
          return result;
          
        default:
          Logger.log(`【エラー】無効なアクション "${action}"`);
          const errorHtml = `
            <!DOCTYPE html>
            <html>
            <head>
              <meta charset="UTF-8">
              <title>エラー</title>
              <style>
                body { font-family: sans-serif; padding: 20px; background: #fff3cd; }
                h1 { color: #856404; }
                pre { background: #f5f5f5; padding: 10px; border-radius: 4px; }
              </style>
            </head>
            <body>
              <h1>エラー: 無効なアクション</h1>
              <p>アクション "<strong>${action}</strong>" は認識されませんでした。</p>
              <p><strong>利用可能なアクション:</strong> approve, reject, comment</p>
              <p><strong>現在のパラメータ:</strong></p>
              <pre>${JSON.stringify(e ? e.parameter : {}, null, 2)}</pre>
            </body>
            </html>
          `;
          return HtmlService.createHtmlOutput(errorHtml);
      }
    } catch (actionError) {
      Logger.log(`アクション処理エラー: ${actionError.toString()}`);
      Logger.log(`スタック: ${actionError.stack}`);
      const errorHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>エラー</title></head><body style="font-family:sans-serif;padding:20px;"><h1>処理エラー</h1><p>' + actionError.toString() + '</p></body></html>';
      return HtmlService.createHtmlOutput(errorHtml);
    }
    
  } catch (error) {
    Logger.log(`=== doGet全体エラー ===`);
    Logger.log(`エラー: ${error.toString()}`);
    Logger.log(`スタック: ${error.stack || 'なし'}`);
    
    // エラーメッセージを詳細に表示
    const errorHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>エラー</title><style>body{font-family:sans-serif;padding:20px;}pre{background:#f5f5f5;padding:10px;border-radius:4px;overflow-x:auto;}</style></head><body><h1>エラーが発生しました</h1><p><strong>エラーメッセージ:</strong></p><pre>' + error.toString() + '</pre><p><strong>スタックトレース:</strong></p><pre>' + (error.stack || 'スタック情報なし') + '</pre></body></html>';
    return HtmlService.createHtmlOutput(errorHtml);
  }
}

function handleApply(sheet, row) {
  // この関数は呼ばれないはず（doGetで先に処理される）が、念のため完全に無効化
  Logger.log('【正式申請】handleApply関数が呼ばれましたが、処理を完全にスキップします');
  // 最もシンプルなHTMLを直接返す
  try {
    const html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>完了</title></head><body style="font-family:sans-serif;text-align:center;padding:50px;"><h1 style="color:#4caf50;">✓ 完了</h1><p>正式申請を受け付けました。</p></body></html>';
    return HtmlService.createHtmlOutput(html);
  } catch (error) {
    Logger.log(`【正式申請】handleApply HTML生成エラー: ${error.toString()}`);
    return ContentService.createTextOutput('正式申請を受け付けました。');
  }
}

function handleApprove(sheet, row, type) {
  try {
    Logger.log(`【承認処理開始】行${row}, タイプ: ${type}`);
    const repairId = sheet.getRange(row, COL.REPAIR_ID + 1).getValue();
    const now = getTodayTokyo();
    const approverName = type === 'GM' ? 'GM（村松さん）' : '代表（鈴木安太郎さん）';
    
    Logger.log(`【承認処理】修繕ID: ${repairId}, 承認者: ${approverName}`);
    
    // 承認ステータス更新
    try {
      sheet.getRange(row, COL.APPROVAL_STATUS + 1).setValue('承認');
      Logger.log(`【承認処理】APPROVAL_STATUS更新完了`);
    } catch (e) {
      Logger.log(`【承認処理】APPROVAL_STATUS更新エラー: ${e.toString()}`);
    }
    
    try {
      sheet.getRange(row, COL.APPROVAL_TYPE + 1).setValue(type === 'GM' ? 'GM承認' : '代表承認');
      Logger.log(`【承認処理】APPROVAL_TYPE更新完了`);
    } catch (e) {
      Logger.log(`【承認処理】APPROVAL_TYPE更新エラー: ${e.toString()}`);
    }
    
    try {
      sheet.getRange(row, COL.APPROVAL_REASON + 1).setValue(`${now} ${approverName}承認`);
      Logger.log(`【承認処理】APPROVAL_REASON更新完了`);
    } catch (e) {
      Logger.log(`【承認処理】APPROVAL_REASON更新エラー: ${e.toString()}`);
    }
    
    // ログ記録
    try {
      const logText = `${now} ${approverName}承認`;
      appendApprovalLog(sheet, row, logText);
      Logger.log(`【承認処理】ログ記録完了`);
    } catch (e) {
      Logger.log(`【承認処理】ログ記録エラー: ${e.toString()}`);
    }
    
    // ステータスを「対応中」に更新、稟議ステータスを「承認完了」に更新（GM承認・代表承認どちらでも）
    try {
      const statusCell = sheet.getRange(row, COL.STATUS + 1);
      statusCell.clearDataValidations(); // 入力規則を削除
      statusCell.setValue('対応中');
      // 稟議ステータスも更新
      sheet.getRange(row, COL.RINGI_STATUS + 1).setValue('承認完了');
      Logger.log(`【承認処理】STATUS更新完了: 対応中, RINGI_STATUS: 承認完了`);
    } catch (validationError) {
      Logger.log(`【警告】ステータス更新で入力規則エラー（無視）: ${validationError.toString()}`);
      // 再試行（入力規則削除を試行）
      try {
        const statusCell = sheet.getRange(row, COL.STATUS + 1);
        statusCell.clearDataValidations();
        statusCell.setValue('対応中');
        sheet.getRange(row, COL.RINGI_STATUS + 1).setValue('承認完了');
        Logger.log(`【承認処理】STATUS再試行成功`);
      } catch (retryError) {
        Logger.log(`【警告】ステータス更新の再試行も失敗（無視）: ${retryError.toString()}`);
      }
    }
    
    // スプレッドシートへの変更を確実に保存
    try {
      SpreadsheetApp.flush();
      Logger.log(`【承認処理】スプレッドシートへの変更を保存しました`);
    } catch (flushError) {
      Logger.log(`【承認処理】flushエラー（無視）: ${flushError.toString()}`);
    }
    
    // 承認完了時のChat通知（GM承認・代表承認どちらでも）
    try {
      const docUrl = getDocUrlFromRow(sheet, row);
      const notificationType = type === 'GM' ? 'GM承認' : '承認';
      sendApprovalCompleteNotification(repairId, docUrl, notificationType);
      Logger.log(`【承認処理】承認完了通知送信完了: ${notificationType}`);
    } catch (e) {
      Logger.log(`【承認処理】承認完了通知送信エラー: ${e.toString()}`);
    }
    
    Logger.log(`【承認処理完了】`);
    
    // デバッグ: スプレッドシートの値を確認
    try {
      const debugInfo = {
        repairId: repairId,
        approvalStatus: sheet.getRange(row, COL.APPROVAL_STATUS + 1).getValue(),
        approvalType: sheet.getRange(row, COL.APPROVAL_TYPE + 1).getValue(),
        approvalReason: sheet.getRange(row, COL.APPROVAL_REASON + 1).getValue(),
        ringiStatus: sheet.getRange(row, COL.RINGI_STATUS + 1).getValue(),
        status: sheet.getRange(row, COL.STATUS + 1).getValue()
      };
      Logger.log(`【承認処理デバッグ】スプレッドシートの値: ${JSON.stringify(debugInfo)}`);
      
      // 成功画面にデバッグ情報を含める
      const debugHtml = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="UTF-8">
          <title>承認完了</title>
          <style>
            body { font-family: sans-serif; padding: 20px; }
            .success { color: #4caf50; font-size: 2rem; margin-bottom: 1rem; }
            .debug { background: #f5f5f5; padding: 1rem; border-radius: 4px; margin-top: 1rem; font-size: 0.9rem; }
            .debug h3 { margin-top: 0; }
            .debug pre { margin: 0; white-space: pre-wrap; }
          </style>
        </head>
        <body>
          <div class="success">✓</div>
          <h1>承認が完了しました</h1>
          <div class="debug">
            <h3>デバッグ情報:</h3>
            <pre>${JSON.stringify(debugInfo, null, 2)}</pre>
          </div>
        </body>
        </html>
      `;
      return HtmlService.createHtmlOutput(debugHtml);
    } catch (debugError) {
      Logger.log(`【承認処理】デバッグ情報取得エラー: ${debugError.toString()}`);
      return createSuccessHtml('承認が完了しました。');
    }
  } catch (error) {
    Logger.log(`【承認エラー】全体エラー: ${error.toString()}`);
    Logger.log(`【承認エラー】スタック: ${error.stack}`);
    return createErrorHtml(`承認エラー: ${error.toString()}`);
  }
}

function handleReject(sheet, row, type) {
  try {
    Logger.log(`【否決処理開始】行${row}, タイプ: ${type}`);
    const repairId = sheet.getRange(row, COL.REPAIR_ID + 1).getValue();
    const now = getTodayTokyo();
    const approverName = type === 'GM' ? 'GM（村松さん）' : '代表（鈴木安太郎さん）';
    
    Logger.log(`【否決処理】修繕ID: ${repairId}, 否決者: ${approverName}`);
    
    // 否決ステータス更新（入力規則エラーは無視）
    try {
      sheet.getRange(row, COL.APPROVAL_STATUS + 1).setValue('否決');
      Logger.log(`【否決処理】APPROVAL_STATUS更新完了`);
    } catch (e) {
      Logger.log(`【否決処理】APPROVAL_STATUS更新エラー: ${e.toString()}`);
    }
    
    try {
      sheet.getRange(row, COL.APPROVAL_TYPE + 1).setValue(type === 'GM' ? 'GM否決' : '代表否決');
      Logger.log(`【否決処理】APPROVAL_TYPE更新完了`);
    } catch (e) {
      Logger.log(`【否決処理】APPROVAL_TYPE更新エラー: ${e.toString()}`);
    }
    
    try {
      sheet.getRange(row, COL.APPROVAL_REASON + 1).setValue(`${now} ${approverName}否決`);
      Logger.log(`【否決処理】APPROVAL_REASON更新完了`);
    } catch (e) {
      Logger.log(`【否決処理】APPROVAL_REASON更新エラー: ${e.toString()}`);
    }
    
    try {
      sheet.getRange(row, COL.RINGI_STATUS + 1).setValue('否決');
      Logger.log(`【否決処理】RINGI_STATUS更新完了`);
    } catch (e) {
      Logger.log(`【否決処理】RINGI_STATUS更新エラー: ${e.toString()}`);
    }
    
    // STATUS列の入力規則を削除してから値を設定
    try {
      const statusCell = sheet.getRange(row, COL.STATUS + 1);
      statusCell.clearDataValidations(); // 入力規則を削除
      statusCell.setValue('クローズ');
      Logger.log(`【否決処理】STATUS更新完了: クローズ`);
    } catch (validationError) {
      Logger.log(`【警告】ステータス更新で入力規則エラー（無視）: ${validationError.toString()}`);
      // 再試行（入力規則削除を試行）
      try {
        const statusCell = sheet.getRange(row, COL.STATUS + 1);
        statusCell.clearDataValidations();
        statusCell.setValue('クローズ');
        Logger.log(`【否決処理】STATUS再試行成功`);
      } catch (retryError) {
        Logger.log(`【警告】ステータス更新の再試行も失敗（無視）: ${retryError.toString()}`);
      }
    }
    
    // ログ記録
    try {
      const logText = `${now} ${approverName}否決`;
      appendApprovalLog(sheet, row, logText);
      Logger.log(`【否決処理】ログ記録完了`);
    } catch (e) {
      Logger.log(`【否決処理】ログ記録エラー: ${e.toString()}`);
    }
    
    // スプレッドシートへの変更を確実に保存
    try {
      SpreadsheetApp.flush();
      Logger.log(`【否決処理】スプレッドシートへの変更を保存しました`);
    } catch (flushError) {
      Logger.log(`【否決処理】flushエラー（無視）: ${flushError.toString()}`);
    }
    
    // 否決時のChat通知
    try {
      const docUrl = getDocUrlFromRow(sheet, row);
      sendApprovalCompleteNotification(repairId, docUrl, '否決');
      Logger.log(`【否決処理】否決完了通知送信完了`);
    } catch (e) {
      Logger.log(`【否決処理】否決完了通知送信エラー: ${e.toString()}`);
    }
    
    Logger.log(`【否決処理完了】`);
    
    // デバッグ: スプレッドシートの値を確認
    try {
      const debugInfo = {
        repairId: repairId,
        approvalStatus: sheet.getRange(row, COL.APPROVAL_STATUS + 1).getValue(),
        approvalType: sheet.getRange(row, COL.APPROVAL_TYPE + 1).getValue(),
        approvalReason: sheet.getRange(row, COL.APPROVAL_REASON + 1).getValue(),
        ringiStatus: sheet.getRange(row, COL.RINGI_STATUS + 1).getValue(),
        status: sheet.getRange(row, COL.STATUS + 1).getValue()
      };
      Logger.log(`【否決処理デバッグ】スプレッドシートの値: ${JSON.stringify(debugInfo)}`);
      
      // 成功画面にデバッグ情報を含める
      const debugHtml = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="UTF-8">
          <title>否決完了</title>
          <style>
            body { font-family: sans-serif; padding: 20px; }
            .success { color: #f44336; font-size: 2rem; margin-bottom: 1rem; }
            .debug { background: #f5f5f5; padding: 1rem; border-radius: 4px; margin-top: 1rem; font-size: 0.9rem; }
            .debug h3 { margin-top: 0; }
            .debug pre { margin: 0; white-space: pre-wrap; }
          </style>
        </head>
        <body>
          <div class="success">✗</div>
          <h1>否決が完了しました</h1>
          <div class="debug">
            <h3>デバッグ情報:</h3>
            <pre>${JSON.stringify(debugInfo, null, 2)}</pre>
          </div>
        </body>
        </html>
      `;
      return HtmlService.createHtmlOutput(debugHtml);
    } catch (debugError) {
      Logger.log(`【否決処理】デバッグ情報取得エラー: ${debugError.toString()}`);
      return createSuccessHtml('否決が完了しました。');
    }
  } catch (error) {
    Logger.log(`【否決エラー】全体エラー: ${error.toString()}`);
    Logger.log(`【否決エラー】スタック: ${error.stack}`);
    return createErrorHtml(`否決エラー: ${error.toString()}`);
  }
}

function handleComment(sheet, row, type, e) {
  try {
    Logger.log(`【コメント処理開始】行${row}, タイプ: ${type}`);
    Logger.log(`【コメント処理】パラメータ: ${JSON.stringify(e ? e.parameter : {})}`);
    
    const repairId = sheet.getRange(row, COL.REPAIR_ID + 1).getValue();
    const approverName = type === 'GM' ? 'GM（村松さん）' : '代表（鈴木安太郎さん）';
    const docUrl = getDocUrlFromRow(sheet, row);
    
    Logger.log(`【コメント処理】修繕ID: ${repairId}, 承認者: ${approverName}`);
    
    // GET/POSTリクエスト（コメント送信）の場合
    const comment = (e && e.parameter && e.parameter.comment) ? String(e.parameter.comment).trim() : '';
    
    Logger.log(`【コメント処理】コメント有無: ${comment ? 'あり' : 'なし'}`);
    
    if (comment && comment.length > 0) {
      Logger.log(`【コメント処理】コメントを受信: ${comment.substring(0, 50)}...`);
      const now = getTodayTokyo();
      
      try {
        // コメントをAPPROVAL_REASON列に記録（理由として）
        sheet.getRange(row, COL.APPROVAL_REASON + 1).setValue(comment);
        Logger.log(`【コメント処理】APPROVAL_REASON更新完了: ${comment.substring(0, 30)}...`);
      } catch (err) {
        Logger.log(`【コメント処理】APPROVAL_REASON更新エラー: ${err.toString()}`);
      }
      
      try {
        // コメントをNOTES列にも記録（ログとして）
        const logText = `${now} ${approverName}コメント: ${comment}`;
        appendApprovalLog(sheet, row, logText);
        Logger.log(`【コメント処理】ログ記録完了`);
      } catch (err) {
        Logger.log(`【コメント処理】ログ記録エラー: ${err.toString()}`);
      }
      
      // スプレッドシートへの変更を確実に保存
      try {
        SpreadsheetApp.flush();
        Logger.log(`【コメント処理】スプレッドシートへの変更を保存しました`);
      } catch (flushError) {
        Logger.log(`【コメント処理】flushエラー（無視）: ${flushError.toString()}`);
      }
      
      Logger.log(`【コメント処理完了】`);
      
      // デバッグ: スプレッドシートの値を確認
      try {
        const debugInfo = {
          repairId: repairId,
          approvalReason: sheet.getRange(row, COL.APPROVAL_REASON + 1).getValue(),
          notes: sheet.getRange(row, COL.NOTES + 1).getValue()
        };
        Logger.log(`【コメント処理デバッグ】スプレッドシートの値: ${JSON.stringify(debugInfo)}`);
        
        // 成功画面にデバッグ情報を含める
        const debugHtml = `
          <!DOCTYPE html>
          <html>
          <head>
            <meta charset="UTF-8">
            <title>コメント記録完了</title>
            <style>
              body { font-family: sans-serif; padding: 20px; }
              .success { color: #4caf50; font-size: 2rem; margin-bottom: 1rem; }
              .debug { background: #f5f5f5; padding: 1rem; border-radius: 4px; margin-top: 1rem; font-size: 0.9rem; }
              .debug h3 { margin-top: 0; }
              .debug pre { margin: 0; white-space: pre-wrap; }
            </style>
          </head>
          <body>
            <div class="success">✓</div>
            <h1>コメントを記録しました</h1>
            <div class="debug">
              <h3>デバッグ情報:</h3>
              <pre>${JSON.stringify(debugInfo, null, 2)}</pre>
            </div>
          </body>
          </html>
        `;
        return HtmlService.createHtmlOutput(debugHtml);
      } catch (debugError) {
        Logger.log(`【コメント処理】デバッグ情報取得エラー: ${debugError.toString()}`);
        return createSuccessHtml('コメントを記録しました。');
      }
    }
    
    // GETリクエスト（コメント入力画面表示）の場合
    Logger.log(`【コメント処理】コメント入力画面を表示`);
    const commentUrl = CONFIG.SCRIPT_WEB_APP_URL + '?action=comment&row=' + row + '&type=' + encodeURIComponent(type);
    Logger.log(`【コメント処理】コメントURL: ${commentUrl}`);
    Logger.log(`【コメント処理】WebアプリURL: ${CONFIG.SCRIPT_WEB_APP_URL}`);
    Logger.log(`【コメント処理】行番号: ${row}, タイプ: ${type}`);
    
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>コメント入力</title>
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            max-width: 600px;
            margin: 50px auto;
            padding: 20px;
            background: #f5f5f5;
          }
          .container {
            background: white;
            padding: 2rem;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          }
          h1 {
            color: #333;
            margin-top: 0;
          }
          .info {
            background: #e3f2fd;
            padding: 1rem;
            border-radius: 4px;
            margin-bottom: 1.5rem;
          }
          .info p {
            margin: 0.5rem 0;
          }
          label {
            display: block;
            margin-bottom: 0.5rem;
            color: #666;
            font-weight: 500;
          }
          textarea {
            width: 100%;
            min-height: 150px;
            padding: 0.75rem;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-family: inherit;
            font-size: 14px;
            resize: vertical;
            box-sizing: border-box;
          }
          .buttons {
            margin-top: 1.5rem;
            display: flex;
            gap: 1rem;
          }
          button {
            flex: 1;
            padding: 0.75rem 1.5rem;
            border: none;
            border-radius: 4px;
            font-size: 16px;
            cursor: pointer;
            transition: background 0.2s;
          }
          .btn-submit {
            background: #4caf50;
            color: white;
          }
          .btn-submit:hover {
            background: #45a049;
          }
          .btn-cancel {
            background: #f44336;
            color: white;
          }
          .btn-cancel:hover {
            background: #da190b;
          }
          .btn-docs {
            background: #2196f3;
            color: white;
          }
          .btn-docs:hover {
            background: #0b7dda;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h1>コメント入力</h1>
          <div class="info">
            <p><strong>修繕ID:</strong> ${repairId}</p>
            <p><strong>承認者:</strong> ${approverName}</p>
          </div>
          <form id="commentForm" onsubmit="return submitComment(event)">
            <label for="comment">コメント:</label>
            <textarea id="comment" name="comment" placeholder="コメントを入力してください..." required></textarea>
            <div class="buttons">
              <button type="submit" class="btn-submit">送信</button>
              <button type="button" class="btn-cancel" onclick="window.close()">キャンセル</button>
              ${docUrl ? `<a href="${docUrl}" target="_blank" style="text-decoration: none;"><button type="button" class="btn-docs">Docsを開く</button></a>` : ''}
            </div>
          </form>
          <script>
            function submitComment(event) {
              event.preventDefault();
              const comment = document.getElementById('comment').value;
              if (!comment.trim()) {
                alert('コメントを入力してください。');
                return false;
              }
              
              // GETパラメータで送信（URLエンコード）
              const encodedComment = encodeURIComponent(comment);
              const baseUrl = '${commentUrl}';
              const url = baseUrl + '&comment=' + encodedComment;
              console.log('【コメント送信】URL:', url);
              window.location.href = url;
              return false;
            }
          </script>
        </div>
      </body>
      </html>
    `;
    
    Logger.log(`【コメント処理】HTML生成完了`);
    return HtmlService.createHtmlOutput(html);
    
  } catch (error) {
    Logger.log(`【コメントエラー】全体エラー: ${error.toString()}`);
    Logger.log(`【コメントエラー】スタック: ${error.stack}`);
    return createErrorHtml(`コメントエラー: ${error.toString()}`);
  }
}

function getDocUrlFromRow(sheet, row) {
  // 備考列からDocs URLを取得（またはフォルダから検索）
  const repairId = sheet.getRange(row, COL.REPAIR_ID + 1).getValue();
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const files = folder.getFilesByName(`${repairId}_稟議書`);
  if (files.hasNext()) {
    return files.next().getUrl();
  }
  return '';
}

function appendApprovalLog(sheet, row, logText) {
  const notesCell = sheet.getRange(row, COL.NOTES + 1);
  const existing = notesCell.getValue() || '';
  notesCell.setValue(existing ? `${existing}\n${logText}` : logText);
}

function createSuccessHtml(message) {
  try {
    const safeMessage = message || '処理が完了しました。';
    return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>完了</title>
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
          display: flex;
          justify-content: center;
          align-items: center;
          min-height: 100vh;
          margin: 0;
          background: #f5f5f5;
        }
        .container {
          background: white;
          padding: 2rem;
          border-radius: 8px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          text-align: center;
          max-width: 400px;
        }
        .success {
          color: #4caf50;
          font-size: 3rem;
          margin-bottom: 1rem;
        }
        h1 {
          color: #333;
          margin: 0 0 1rem 0;
        }
        p {
          color: #666;
          margin: 0;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="success">✓</div>
        <h1>完了</h1>
        <p>${safeMessage}</p>
      </div>
    </body>
    </html>
  `);
  } catch (error) {
    Logger.log(`【警告】createSuccessHtmlエラー（フォールバック）: ${error.toString()}`);
    // フォールバック: シンプルなHTMLを直接返す
    return HtmlService.createHtmlOutput(`<html><body><h1>✓ 完了</h1><p>${message || '処理が完了しました。'}</p></body></html>`);
  }
}

function createErrorHtml(message) {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>エラー</title>
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
          display: flex;
          justify-content: center;
          align-items: center;
          min-height: 100vh;
          margin: 0;
          background: #f5f5f5;
        }
        .container {
          background: white;
          padding: 2rem;
          border-radius: 8px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          text-align: center;
          max-width: 400px;
        }
        .error {
          color: #f44336;
          font-size: 3rem;
          margin-bottom: 1rem;
        }
        h1 {
          color: #333;
          margin: 0 0 1rem 0;
        }
        p {
          color: #666;
          margin: 0;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="error">✗</div>
        <h1>エラー</h1>
        <p>${message}</p>
      </div>
    </body>
    </html>
  `);
}

// ====== トリガー設定 ======
function setupTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log(`【トリガー設定】既存のトリガー数: ${triggers.length}`);
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processRepairEmails') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`【トリガー設定】既存のトリガーを削除: ${trigger.getUniqueId()}`);
    }
  }
  
  // 新しいトリガーを作成（5分ごと）
  const newTrigger = ScriptApp.newTrigger('processRepairEmails')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  Logger.log(`【トリガー設定】新しいトリガーを作成しました: ${newTrigger.getUniqueId()}`);
  Logger.log('トリガーを設定しました（5分ごとに実行）');
}

// ====== トリガー確認関数（デバッグ用） ======
function checkTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log(`【トリガー確認】登録されているトリガー数: ${triggers.length}`);
  
  for (const trigger of triggers) {
    Logger.log(`【トリガー確認】関数: ${trigger.getHandlerFunction()}, タイプ: ${trigger.getEventType()}, ID: ${trigger.getUniqueId()}`);
  }
  
  return triggers.length;
}

// ====== 承認/否決/コメント処理のテスト関数 ======
function testApprove(row) {
  try {
    Logger.log(`【テスト】承認処理をテストします（行${row}）`);
    const sheet = getSheet();
    const type = 'GM';
    const result = handleApprove(sheet, row, type);
    Logger.log(`【テスト】承認処理完了`);
    return '承認処理が完了しました。ログを確認してください。';
  } catch (error) {
    Logger.log(`【テスト】承認処理エラー: ${error.toString()}`);
    return `エラー: ${error.toString()}`;
  }
}

function testReject(row) {
  try {
    Logger.log(`【テスト】否決処理をテストします（行${row}）`);
    const sheet = getSheet();
    const type = 'GM';
    const result = handleReject(sheet, row, type);
    Logger.log(`【テスト】否決処理完了`);
    return '否決処理が完了しました。ログを確認してください。';
  } catch (error) {
    Logger.log(`【テスト】否決処理エラー: ${error.toString()}`);
    return `エラー: ${error.toString()}`;
  }
}

function testComment(row) {
  try {
    Logger.log(`【テスト】コメント処理をテストします（行${row}）`);
    const sheet = getSheet();
    const type = 'GM';
    const e = {
      parameter: {
        action: 'comment',
        row: row.toString(),
        type: type
      }
    };
    const result = handleComment(sheet, row, type, e);
    Logger.log(`【テスト】コメント処理完了`);
    return 'コメント処理が完了しました。ログを確認してください。';
  } catch (error) {
    Logger.log(`【テスト】コメント処理エラー: ${error.toString()}`);
    return `エラー: ${error.toString()}`;
  }
}

// ====== WebアプリURL確認関数 ======
function checkWebAppUrl() {
  Logger.log(`【WebアプリURL確認】設定値: ${CONFIG.SCRIPT_WEB_APP_URL}`);
  return CONFIG.SCRIPT_WEB_APP_URL;
}

// ====== ボタンURL生成テスト関数 ======
function testButtonUrls(rowNum, approverType) {
  const approveUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=approve&row=${rowNum}&type=${approverType}`;
  const rejectUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=reject&row=${rowNum}&type=${approverType}`;
  const commentUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=comment&row=${rowNum}&type=${approverType}`;
  
  Logger.log(`【ボタンURLテスト】`);
  Logger.log(`承認URL: ${approveUrl}`);
  Logger.log(`否決URL: ${rejectUrl}`);
  Logger.log(`コメントURL: ${commentUrl}`);
  
  return {
    approve: approveUrl,
    reject: rejectUrl,
    comment: commentUrl
  };
}

// ====== 承認・否決・コメントの完全診断関数 ======
function diagnoseApprovalFlowComplete(rowNum) {
  try {
    const sheet = getSheet();
    const repairId = sheet.getRange(rowNum, COL.REPAIR_ID + 1).getValue();
    const docUrl = getDocUrlFromRow(sheet, rowNum);
    
    // 現在のステータスを確認
    const currentStatus = sheet.getRange(rowNum, COL.STATUS + 1).getValue();
    const approvalStatus = sheet.getRange(rowNum, COL.APPROVAL_STATUS + 1).getValue();
    const approvalType = sheet.getRange(rowNum, COL.APPROVAL_TYPE + 1).getValue();
    const approvalReason = sheet.getRange(rowNum, COL.APPROVAL_REASON + 1).getValue();
    const ringiStatus = sheet.getRange(rowNum, COL.RINGI_STATUS + 1).getValue();
    
    // URLを生成（GMと代表の両方）
    const approveUrlGM = `${CONFIG.SCRIPT_WEB_APP_URL}?action=approve&row=${rowNum}&type=GM`;
    const rejectUrlGM = `${CONFIG.SCRIPT_WEB_APP_URL}?action=reject&row=${rowNum}&type=GM`;
    const commentUrlGM = `${CONFIG.SCRIPT_WEB_APP_URL}?action=comment&row=${rowNum}&type=GM`;
    const approveUrlRep = `${CONFIG.SCRIPT_WEB_APP_URL}?action=approve&row=${rowNum}&type=代表`;
    const rejectUrlRep = `${CONFIG.SCRIPT_WEB_APP_URL}?action=reject&row=${rowNum}&type=代表`;
    const commentUrlRep = `${CONFIG.SCRIPT_WEB_APP_URL}?action=comment&row=${rowNum}&type=代表`;
    
    Logger.log(`【完全診断】修繕ID: ${repairId}`);
    Logger.log(`【完全診断】現在のステータス: ${currentStatus}`);
    Logger.log(`【完全診断】承認ステータス: ${approvalStatus}`);
    Logger.log(`【完全診断】承認区分: ${approvalType}`);
    Logger.log(`【完全診断】理由: ${approvalReason}`);
    Logger.log(`【完全診断】稟議ステータス: ${ringiStatus}`);
    Logger.log(`【完全診断】WebアプリURL: ${CONFIG.SCRIPT_WEB_APP_URL}`);
    
    // HTML診断ページを生成
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>承認フロー完全診断</title>
        <style>
          body { font-family: sans-serif; padding: 20px; max-width: 1000px; margin: 0 auto; }
          h1 { color: #333; }
          .section { background: #f5f5f5; padding: 15px; margin: 15px 0; border-radius: 4px; }
          .section h2 { margin-top: 0; color: #666; }
          .url { background: #e3f2fd; padding: 10px; border-radius: 4px; margin: 10px 0; word-break: break-all; font-size: 0.9rem; }
          .url a { color: #1976d2; text-decoration: none; }
          .url a:hover { text-decoration: underline; }
          .status { padding: 5px 10px; border-radius: 4px; display: inline-block; margin: 5px 0; }
          .status.current { background: #fff3cd; color: #856404; }
          .status.approval { background: #d1ecf1; color: #0c5460; }
          .button { display: inline-block; padding: 10px 20px; margin: 5px; background: #4caf50; color: white; text-decoration: none; border-radius: 4px; }
          .button:hover { background: #45a049; }
          .button.reject { background: #f44336; }
          .button.reject:hover { background: #da190b; }
          .button.comment { background: #2196F3; }
          .button.comment:hover { background: #0b7dda; }
          pre { background: #fff; padding: 10px; border-radius: 4px; overflow-x: auto; font-size: 0.85rem; }
          .warning { background: #fff3cd; padding: 15px; border-radius: 4px; margin: 15px 0; border-left: 4px solid #ffc107; }
        </style>
      </head>
      <body>
        <h1>承認フロー完全診断</h1>
        
        <div class="warning">
          <strong>⚠️ 重要:</strong> このページでボタンをクリックして、実際の動作を確認してください。
        </div>
        
        <div class="section">
          <h2>現在の状態</h2>
          <p><strong>修繕ID:</strong> ${repairId}</p>
          <p><strong>行番号:</strong> ${rowNum}</p>
          <p><strong>ステータス:</strong> <span class="status current">${currentStatus || '(未設定)'}</span></p>
          <p><strong>承認ステータス:</strong> <span class="status approval">${approvalStatus || '(未設定)'}</span></p>
          <p><strong>承認区分:</strong> ${approvalType || '(未設定)'}</p>
          <p><strong>理由:</strong> ${approvalReason || '(未設定)'}</p>
          <p><strong>稟議ステータス:</strong> ${ringiStatus || '(未設定)'}</p>
        </div>
        
        <div class="section">
          <h2>GM承認用テストボタン</h2>
          <p>以下のボタンをクリックして動作を確認してください：</p>
          <a href="${approveUrlGM}" class="button" target="_blank">承認する（GM）</a>
          <a href="${rejectUrlGM}" class="button reject" target="_blank">否決する（GM）</a>
          <a href="${commentUrlGM}" class="button comment" target="_blank">コメント（GM）</a>
        </div>
        
        <div class="section">
          <h2>代表承認用テストボタン</h2>
          <p>以下のボタンをクリックして動作を確認してください：</p>
          <a href="${approveUrlRep}" class="button" target="_blank">承認する（代表）</a>
          <a href="${rejectUrlRep}" class="button reject" target="_blank">否決する（代表）</a>
          <a href="${commentUrlRep}" class="button comment" target="_blank">コメント（代表）</a>
        </div>
        
        <div class="section">
          <h2>生成されたURL（GM）</h2>
          <div class="url">
            <strong>承認URL:</strong><br>
            <a href="${approveUrlGM}" target="_blank">${approveUrlGM}</a>
          </div>
          <div class="url">
            <strong>否決URL:</strong><br>
            <a href="${rejectUrlGM}" target="_blank">${rejectUrlGM}</a>
          </div>
          <div class="url">
            <strong>コメントURL:</strong><br>
            <a href="${commentUrlGM}" target="_blank">${commentUrlGM}</a>
          </div>
        </div>
        
        <div class="section">
          <h2>設定値</h2>
          <pre>SCRIPT_WEB_APP_URL: ${CONFIG.SCRIPT_WEB_APP_URL}</pre>
        </div>
      </body>
      </html>
    `;
    
    return HtmlService.createHtmlOutput(html);
  } catch (error) {
    Logger.log(`【完全診断エラー】: ${error.toString()}`);
    return HtmlService.createHtmlOutput(`<html><body><h1>診断エラー</h1><p>${error.toString()}</p></body></html>`);
  }
}

// ====== 承認・否決・コメントの診断関数 ======
function diagnoseApprovalFlow(rowNum) {
  try {
    const sheet = getSheet();
    const repairId = sheet.getRange(rowNum, COL.REPAIR_ID + 1).getValue();
    const docUrl = getDocUrlFromRow(sheet, rowNum);
    
    // 現在のステータスを確認
    const currentStatus = sheet.getRange(rowNum, COL.STATUS + 1).getValue();
    const approvalStatus = sheet.getRange(rowNum, COL.APPROVAL_STATUS + 1).getValue();
    const approvalType = sheet.getRange(rowNum, COL.APPROVAL_TYPE + 1).getValue();
    const approvalReason = sheet.getRange(rowNum, COL.APPROVAL_REASON + 1).getValue();
    const ringiStatus = sheet.getRange(rowNum, COL.RINGI_STATUS + 1).getValue();
    
    // URLを生成
    const approveUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=approve&row=${rowNum}&type=GM`;
    const rejectUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=reject&row=${rowNum}&type=GM`;
    const commentUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=comment&row=${rowNum}&type=GM`;
    
    Logger.log(`【診断】修繕ID: ${repairId}`);
    Logger.log(`【診断】現在のステータス: ${currentStatus}`);
    Logger.log(`【診断】承認ステータス: ${approvalStatus}`);
    Logger.log(`【診断】承認区分: ${approvalType}`);
    Logger.log(`【診断】理由: ${approvalReason}`);
    Logger.log(`【診断】稟議ステータス: ${ringiStatus}`);
    Logger.log(`【診断】承認URL: ${approveUrl}`);
    Logger.log(`【診断】否決URL: ${rejectUrl}`);
    Logger.log(`【診断】コメントURL: ${commentUrl}`);
    
    // HTML診断ページを生成
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>承認フロー診断</title>
        <style>
          body { font-family: sans-serif; padding: 20px; max-width: 800px; margin: 0 auto; }
          h1 { color: #333; }
          .section { background: #f5f5f5; padding: 15px; margin: 15px 0; border-radius: 4px; }
          .section h2 { margin-top: 0; color: #666; }
          .url { background: #e3f2fd; padding: 10px; border-radius: 4px; margin: 10px 0; word-break: break-all; }
          .url a { color: #1976d2; text-decoration: none; }
          .url a:hover { text-decoration: underline; }
          .status { padding: 5px 10px; border-radius: 4px; display: inline-block; margin: 5px 0; }
          .status.current { background: #fff3cd; color: #856404; }
          .status.approval { background: #d1ecf1; color: #0c5460; }
          .button { display: inline-block; padding: 10px 20px; margin: 5px; background: #4caf50; color: white; text-decoration: none; border-radius: 4px; }
          .button:hover { background: #45a049; }
          .button.reject { background: #f44336; }
          .button.reject:hover { background: #da190b; }
          .button.comment { background: #2196F3; }
          .button.comment:hover { background: #0b7dda; }
          pre { background: #fff; padding: 10px; border-radius: 4px; overflow-x: auto; }
        </style>
      </head>
      <body>
        <h1>承認フロー診断</h1>
        
        <div class="section">
          <h2>現在の状態</h2>
          <p><strong>修繕ID:</strong> ${repairId}</p>
          <p><strong>行番号:</strong> ${rowNum}</p>
          <p><strong>ステータス:</strong> <span class="status current">${currentStatus || '(未設定)'}</span></p>
          <p><strong>承認ステータス:</strong> <span class="status approval">${approvalStatus || '(未設定)'}</span></p>
          <p><strong>承認区分:</strong> ${approvalType || '(未設定)'}</p>
          <p><strong>理由:</strong> ${approvalReason || '(未設定)'}</p>
          <p><strong>稟議ステータス:</strong> ${ringiStatus || '(未設定)'}</p>
        </div>
        
        <div class="section">
          <h2>テストボタン</h2>
          <p>以下のボタンをクリックして動作を確認してください：</p>
          <a href="${approveUrl}" class="button" target="_blank">承認する</a>
          <a href="${rejectUrl}" class="button reject" target="_blank">否決する</a>
          <a href="${commentUrl}" class="button comment" target="_blank">コメント</a>
        </div>
        
        <div class="section">
          <h2>生成されたURL</h2>
          <div class="url">
            <strong>承認URL:</strong><br>
            <a href="${approveUrl}" target="_blank">${approveUrl}</a>
          </div>
          <div class="url">
            <strong>否決URL:</strong><br>
            <a href="${rejectUrl}" target="_blank">${rejectUrl}</a>
          </div>
          <div class="url">
            <strong>コメントURL:</strong><br>
            <a href="${commentUrl}" target="_blank">${commentUrl}</a>
          </div>
        </div>
        
        <div class="section">
          <h2>設定値</h2>
          <pre>SCRIPT_WEB_APP_URL: ${CONFIG.SCRIPT_WEB_APP_URL}</pre>
        </div>
      </body>
      </html>
    `;
    
    return HtmlService.createHtmlOutput(html);
  } catch (error) {
    Logger.log(`【診断エラー】: ${error.toString()}`);
    return HtmlService.createHtmlOutput(`<html><body><h1>診断エラー</h1><p>${error.toString()}</p></body></html>`);
  }
}

// ====== 最小限の承認テスト関数 ======
function testApproveMinimal(row) {
  try {
    const sheet = getSheet();
    const type = 'GM';
    
    // 最小限の処理：スプレッドシートに直接書き込む
    sheet.getRange(row, COL.APPROVAL_STATUS + 1).setValue('承認');
    sheet.getRange(row, COL.APPROVAL_TYPE + 1).setValue('GM承認');
    sheet.getRange(row, COL.APPROVAL_REASON + 1).setValue('テスト承認');
    sheet.getRange(row, COL.RINGI_STATUS + 1).setValue('承認完了');
    sheet.getRange(row, COL.STATUS + 1).clearDataValidations();
    sheet.getRange(row, COL.STATUS + 1).setValue('対応中');
    
    SpreadsheetApp.flush();
    
    Logger.log(`【最小限テスト】承認処理完了（行${row}）`);
    return '最小限の承認処理が完了しました。スプレッドシートを確認してください。';
  } catch (error) {
    Logger.log(`【最小限テスト】エラー: ${error.toString()}`);
    return `エラー: ${error.toString()}`;
  }
}
