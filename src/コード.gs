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
  // APIキー
  GEMINI_API_KEY: '***REDACTED***',
  
  // スプレッドシート
  REPAIR_SYSTEM_SHEET_ID: '1ZAUzoCIIy3h6TNiVnYB7_hjWY-Id9oZ_iX1z88M2yNI',
  SHEET_NAME: '修繕ログ',
  
  // ドライブフォルダ
  FOLDER_ID: '1Qz-HYebqH-vfd8-cYD-xoLsOdEL7PEg5',
  
  // テンプレートDocs
  TEMPLATE_DOC_ID: '1iazbzvlh-VQ046dVgRXyO2BEEWbnGIVHbBTeejeGSjk',
  
  // Chat Webhook
  WEBHOOK_URL: 'https://chat.googleapis.com/v1/spaces/AAQAmERWyO4/messages?key=***REDACTED***&token=***REDACTED***',
  
  // WebアプリURL（デプロイ後に更新）
  SCRIPT_WEB_APP_URL: 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec',
  
  // Geminiモデル
  GEMINI_MODEL_PRIMARY: 'gemini-3.0-pro-preview',
  GEMINI_MODEL_FALLBACK: 'gemini-2.0-flash-exp',
  
  // メール検索条件
  EMAIL_SUBJECT: '修繕依頼',
  EMAIL_SEARCH_MINUTES: 35
};

// ====== 列定義（42列） ======
const COL = {
  REPAIR_ID: 0,           // A: 修繕ID
  RECEIVED_DATETIME: 1,   // B: 受付日時
  REPORTER_NAME: 2,       // C: 報告者名
  AREA: 3,                // D: エリア
  LOCATION_DETAIL: 4,     // E: 場所詳細
  PHOTO1: 5,              // F: 写真1
  PHOTO2: 6,              // G: 写真2
  PHOTO3: 7,              // H: 写真3
  ORIGINAL_TEXT: 8,       // I: 原文（現場入力）
  AI_FORMATTED: 9,        // J: AI整形文
  PROBLEM_SUMMARY: 10,    // K: 問題要約
  CAUSE_ANALYSIS: 11,     // L: 原因分析
  PRIORITY_RANK: 12,      // M: 重要度ランク（A/B/C）
  RANK_REASON: 13,        // N: ランク理由
  RECOMMENDED_TYPE: 14,   // O: 推奨対応タイプ
  WORK_SUMMARY: 15,       // P: 作業内容要約
  RECOMMENDED_STEPS: 16,  // Q: 推奨作業手順
  MATERIALS_LIST: 17,     // R: 必要部材リスト
  ESTIMATED_TIME: 18,     // S: 想定作業時間（分）
  AI_COST_MIN: 19,        // T: AI概算費用下限（円）
  AI_COST_MAX: 20,        // U: AI概算費用上限（円）
  CONTRACTOR_CATEGORY: 21, // V: 想定業者カテゴリ
  CONTRACTOR_AREA: 22,     // W: 想定業者エリア
  SEARCH_KEYWORDS: 23,     // X: 業者検索キーワード
  POSTPONE_RISK: 24,      // Y: 先送りリスク
  ESTIMATE1: 25,          // Z: 見積1
  ESTIMATE2: 26,          // AA: 見積2
  ESTIMATE3: 27,          // AB: 見積3
  SELECTED_CONTRACTOR: 28, // AC: 選定業者
  RINGI_INITIATOR: 29,    // AD: 稟議起案者
  RINGI_REQUIRED: 30,     // AE: 稟議要否
  RINGI_REASON: 31,       // AF: 稟議理由
  STATUS: 32,             // AG: ステータス
  ACTUAL_OWNER: 33,       // AH: 実務担当者
  ACTUAL_COST: 34,        // AI: 実際費用（円）
  RINGI_ID: 35,           // AJ: 稟議ID
  RINGI_STATUS: 36,       // AK: 稟議ステータス
  APPROVAL_STATUS: 37,    // AL: 承認ステータス
  APPROVAL_TYPE: 38,      // AM: 承認区分
  APPROVAL_REASON: 39,    // AN: 理由
  COMPLETION_DATE: 40,    // AO: 完了日
  NOTES: 41               // AP: 備考
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
    const query = `subject:"${CONFIG.EMAIL_SUBJECT}" is:unread newer_than:${CONFIG.EMAIL_SEARCH_MINUTES}m`;
    const threads = GmailApp.search(query, 0, 10);
    
    if (threads.length === 0) {
      Logger.log('処理対象のメールがありません');
      return;
    }
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const message of messages) {
        if (message.isUnread()) {
          processRepairEmail(message);
          message.markRead();
        }
      }
    }
  } catch (error) {
    Logger.log(`メール処理エラー: ${error.toString()}`);
    sendErrorNotification(`メール処理エラー: ${error.toString()}`);
  }
}

function processRepairEmail(message) {
  try {
    const subject = message.getSubject();
    const body = message.getPlainBody();
    const from = message.getFrom();
    const date = message.getDate();
    
    // 画像を取得
    const attachments = message.getAttachments();
    const images = [];
    for (const att of attachments) {
      if (att.getContentType().indexOf('image/') === 0) {
        images.push({
          mimeType: att.getContentType(),
          data: Utilities.base64Encode(att.getBytes())
        });
      }
    }
    
    // AI解析
    const aiResult = analyzeWithGemini(body, images, subject);
    
    // スプレッドシートに記録
    const repairId = generateRepairId();
    const rowData = parseAIResult(aiResult, repairId, from, body);
    const sheet = getSheet();
    sheet.appendRow(rowData);
    const rowNum = sheet.getLastRow();
    
    // Docs生成
    const docUrl = createOrUpdateRingiDoc(rowData, repairId);
    
    // 下書き通知（Chat Cards V2）
    sendDraftNotification(repairId, docUrl, rowNum);
    
    Logger.log(`修繕ID ${repairId} の処理が完了しました`);
  } catch (error) {
    Logger.log(`メール処理エラー: ${error.toString()}`);
    sendErrorNotification(`メール処理エラー: ${error.toString()}`);
  }
}

// ====== Gemini AI解析 ======
function analyzeWithGemini(body, images, subject) {
  try {
    const prompt = buildAnalysisPrompt(body, subject);
    const model = CONFIG.GEMINI_MODEL_PRIMARY;
    
    let result;
    try {
      result = callGeminiAPI(model, prompt, images);
    } catch (error) {
      Logger.log(`Primary model failed, using fallback: ${error.toString()}`);
      result = callGeminiAPI(CONFIG.GEMINI_MODEL_FALLBACK, prompt, images);
    }
    
    return result;
  } catch (error) {
    throw new Error(`Gemini解析エラー: ${error.toString()}`);
  }
}

function buildAnalysisPrompt(body, subject) {
  return `あなたは熟練した設備管理責任者です。修繕報告メールを分析し、以下の形式で出力してください。

【重要】Google検索ツール（googleSearch）を使用して、部材や業者の実在するURLを必ず見つけてください。適当な回答（ハルシネーション）は禁止です。

出力形式（パイプライン区切り）:
報告者名|エリア|場所詳細|AI整形文|問題要約|原因分析|重要度ランク（A/B/C）|ランク理由|推奨対応タイプ|作業内容要約|推奨作業手順|必要部材リスト（URL付き）|想定作業時間（分）|AI概算費用下限（円）|AI概算費用上限（円）|想定業者カテゴリ|想定業者エリア|業者検索キーワード|先送りリスク|部材URL1|部材URL2|部材URL3|業者URL1|業者URL2|業者URL3

入力内容:
件名: ${subject}
本文:
${body}

各項目の説明:
- 報告者名: メール送信者から推測
- エリア: 大浴場、客室フロア、ロビーなど
- 場所詳細: 具体的な場所（例: 305号室 窓側カーテンレール）
- AI整形文: 原文を整理した文章
- 問題要約: 簡潔な問題説明
- 原因分析: 技術的な原因
- 重要度ランク: A（緊急）、B（重要）、C（通常）
- ランク理由: ランク付けの理由
- 推奨対応タイプ: SELF_SIMPLE, SELF_COMPLEX, EXTERNAL など
- 作業内容要約: 簡潔な作業説明
- 推奨作業手順: 番号付き手順
- 必要部材リスト: 部材名とURL（検索ツール使用必須）
- 想定作業時間: 分数
- AI概算費用: 下限と上限（円）
- 想定業者カテゴリ: 内装、配管、電気など
- 想定業者エリア: 地域名
- 業者検索キーワード: 検索用キーワード
- 先送りリスク: 放置した場合のリスク説明
- 部材URL1-3: 実在する部材のURL（検索ツール使用）
- 業者URL1-3: 実在する業者のURL（検索ツール使用）

必ずGoogle検索ツールを使用して、実在するURLを取得してください。`;
}

function callGeminiAPI(model, prompt, images) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
  
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
    tools: [{
      googleSearch: {}
    }],
    generationConfig: {
      temperature: 0.7,
      topK: 40,
      topP: 0.95,
      maxOutputTokens: 8192
    }
  };
  
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
  if (content.parts) {
    for (const part of content.parts) {
      if (part.text) {
        return part.text;
      }
      // functionCallの場合は、最終的なテキストレスポンスを待つ
      // （通常、Geminiは自動でfunctionCallを実行して結果を返す）
    }
  }
  
  // フォールバック: テキストを直接取得
  if (content.parts && content.parts[0] && content.parts[0].text) {
    return content.parts[0].text;
  }
  
  throw new Error('Gemini API: テキストレスポンスが見つかりません');
}

// ====== AI結果のパース ======
function parseAIResult(aiText, repairId, reporterEmail, originalText) {
  const parts = aiText.split('|').map(s => s.trim());
  
  // デフォルト値
  const defaults = {
    reporter: extractNameFromEmail(reporterEmail),
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
    materialUrl1: '',
    materialUrl2: '',
    materialUrl3: '',
    contractorUrl1: '',
    contractorUrl2: '',
    contractorUrl3: ''
  };
  
  const row = new Array(42).fill('');
  
  // 基本情報
  row[COL.REPAIR_ID] = repairId;
  row[COL.RECEIVED_DATETIME] = getTodayTokyo();
  row[COL.REPORTER_NAME] = parts[0] || defaults.reporter;
  row[COL.AREA] = parts[1] || defaults.area;
  row[COL.LOCATION_DETAIL] = parts[2] || defaults.location;
  row[COL.ORIGINAL_TEXT] = originalText.substring(0, 1000);
  row[COL.AI_FORMATTED] = parts[3] || defaults.formatted;
  row[COL.PROBLEM_SUMMARY] = parts[4] || defaults.problem;
  row[COL.CAUSE_ANALYSIS] = parts[5] || defaults.cause;
  row[COL.PRIORITY_RANK] = parts[6] || defaults.rank;
  row[COL.RANK_REASON] = parts[7] || defaults.rankReason;
  row[COL.RECOMMENDED_TYPE] = parts[8] || defaults.type;
  row[COL.WORK_SUMMARY] = parts[9] || defaults.workSummary;
  row[COL.RECOMMENDED_STEPS] = parts[10] || defaults.steps;
  row[COL.MATERIALS_LIST] = parts[11] || defaults.materials;
  row[COL.ESTIMATED_TIME] = parts[12] || defaults.time;
  row[COL.AI_COST_MIN] = parts[13] || defaults.costMin;
  row[COL.AI_COST_MAX] = parts[14] || defaults.costMax;
  row[COL.CONTRACTOR_CATEGORY] = parts[15] || defaults.category;
  row[COL.CONTRACTOR_AREA] = parts[16] || defaults.contractorArea;
  row[COL.SEARCH_KEYWORDS] = parts[17] || defaults.keywords;
  row[COL.POSTPONE_RISK] = parts[18] || defaults.risk;
  
  // URL（見積1-3に部材URL、業者URLを設定）
  const materialUrls = [parts[19] || '', parts[20] || '', parts[21] || ''].filter(u => u);
  const contractorUrls = [parts[22] || '', parts[23] || '', parts[24] || ''].filter(u => u);
  
  if (materialUrls.length > 0) {
    row[COL.ESTIMATE1] = materialUrls[0];
  }
  if (materialUrls.length > 1) {
    row[COL.ESTIMATE2] = materialUrls[1];
  }
  if (contractorUrls.length > 0 && materialUrls.length < 3) {
    row[COL.ESTIMATE3] = contractorUrls[0];
  }
  
  // ステータス
  row[COL.STATUS] = '下書き';
  row[COL.RINGI_REQUIRED] = '要';
  
  return row;
}

function extractNameFromEmail(email) {
  const match = email.match(/^(.+?)\s*<.+>$/);
  if (match) return match[1];
  const atIndex = email.indexOf('@');
  if (atIndex > 0) return email.substring(0, atIndex);
  return email;
}

// ====== Docs生成 ======
function createOrUpdateRingiDoc(rowData, repairId) {
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
    
    for (const [key, value] of Object.entries(replacements)) {
      body.replaceText(key, value);
    }
    
    doc.saveAndClose();
    return newDoc.getUrl();
  } catch (error) {
    Logger.log(`Docs生成エラー: ${error.toString()}`);
    throw error;
  }
}

// ====== Chat通知（Cards V2） ======
function sendDraftNotification(repairId, docUrl, rowNum) {
  const applyUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=apply&row=${rowNum}`;
  
  const card = {
    cardsV2: {
      cardId: 'draft_notification',
      card: {
        header: {
          title: '修繕報告が作成されました',
          subtitle: repairId
        },
        sections: [{
          widgets: [
            {
              textParagraph: {
                text: `修繕ID: <b>${repairId}</b><br/>稟議書を確認して、正式申請してください。`
              }
            },
            {
              buttonList: {
                buttons: [
                  {
                    text: 'Docs確認・修正',
                    onClick: {
                      openLink: {
                        url: docUrl
                      }
                    }
                  },
                  {
                    text: '正式申請する',
                    onClick: {
                      openLink: {
                        url: applyUrl
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

function sendApprovalRequest(repairId, docUrl, rowNum, approverType) {
  const approverName = approverType === 'GM' ? 'GM（村松さん）' : '代表（鈴木安太郎さん）';
  const approveUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=approve&row=${rowNum}&type=${approverType}`;
  const rejectUrl = `${CONFIG.SCRIPT_WEB_APP_URL}?action=reject&row=${rowNum}&type=${approverType}`;
  
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
                        url: docUrl
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

function sendChatMessage(card) {
  try {
    const payload = JSON.stringify(card);
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: payload,
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(CONFIG.WEBHOOK_URL, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log(`Chat送信エラー: ${responseCode} - ${response.getContentText()}`);
    }
  } catch (error) {
    Logger.log(`Chat送信エラー: ${error.toString()}`);
  }
}

function sendErrorNotification(message) {
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
}

// ====== Webアプリ（doGet） ======
function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const row = parseInt(e.parameter.row || '0');
    const type = e.parameter.type || '';
    
    if (!row || row < 2) {
      return HtmlService.createHtmlOutput('<html><body><h1>エラー: 無効な行番号</h1></body></html>');
    }
    
    const sheet = getSheet();
    
    switch (action) {
      case 'apply':
        return handleApply(sheet, row);
      case 'approve':
        return handleApprove(sheet, row, type);
      case 'reject':
        return handleReject(sheet, row, type);
      default:
        return HtmlService.createHtmlOutput('<html><body><h1>エラー: 無効なアクション</h1></body></html>');
    }
  } catch (error) {
    Logger.log(`doGetエラー: ${error.toString()}`);
    return HtmlService.createHtmlOutput(`<html><body><h1>エラー</h1><p>${error.toString()}</p></body></html>`);
  }
}

function handleApply(sheet, row) {
  try {
    // ステータスを「承認依頼中」に更新
    sheet.getRange(row, COL.STATUS + 1).setValue('承認依頼中');
    
    const repairId = sheet.getRange(row, COL.REPAIR_ID + 1).getValue();
    const docUrl = getDocUrlFromRow(sheet, row);
    
    // GMへ承認依頼
    sendApprovalRequest(repairId, docUrl, row, 'GM');
    
    return createSuccessHtml('正式申請が完了しました。GMへ承認依頼を送信しました。');
  } catch (error) {
    Logger.log(`正式申請エラー: ${error.toString()}`);
    return createErrorHtml(`正式申請エラー: ${error.toString()}`);
  }
}

function handleApprove(sheet, row, type) {
  try {
    const repairId = sheet.getRange(row, COL.REPAIR_ID + 1).getValue();
    const now = getTodayTokyo();
    
    // 承認ステータス更新
    sheet.getRange(row, COL.APPROVAL_STATUS + 1).setValue('承認');
    sheet.getRange(row, COL.APPROVAL_TYPE + 1).setValue(type === 'GM' ? 'GM承認' : '代表承認');
    sheet.getRange(row, COL.APPROVAL_REASON + 1).setValue(`${now} ${type === 'GM' ? 'GM' : '代表'}承認`);
    
    // ログ記録
    const logText = `${now} ${type === 'GM' ? 'GM' : '代表'}承認（承認）`;
    appendApprovalLog(sheet, row, logText);
    
    // 代表承認が必要な場合は代表へ送信
    if (type === 'GM') {
      const docUrl = getDocUrlFromRow(sheet, row);
      sendApprovalRequest(repairId, docUrl, row, 'REPRESENTATIVE');
    } else {
      // 代表承認完了
      sheet.getRange(row, COL.STATUS + 1).setValue('承認済');
    }
    
    return createSuccessHtml('承認が完了しました。');
  } catch (error) {
    Logger.log(`承認エラー: ${error.toString()}`);
    return createErrorHtml(`承認エラー: ${error.toString()}`);
  }
}

function handleReject(sheet, row, type) {
  try {
    const repairId = sheet.getRange(row, COL.REPAIR_ID + 1).getValue();
    const now = getTodayTokyo();
    
    // 否決ステータス更新
    sheet.getRange(row, COL.APPROVAL_STATUS + 1).setValue('否決');
    sheet.getRange(row, COL.APPROVAL_TYPE + 1).setValue(type === 'GM' ? 'GM否決' : '代表否決');
    sheet.getRange(row, COL.STATUS + 1).setValue('否決');
    
    // ログ記録
    const logText = `${now} ${type === 'GM' ? 'GM' : '代表'}否決`;
    appendApprovalLog(sheet, row, logText);
    
    return createSuccessHtml('否決が完了しました。');
  } catch (error) {
    Logger.log(`否決エラー: ${error.toString()}`);
    return createErrorHtml(`否決エラー: ${error.toString()}`);
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
        <p>${message}</p>
      </div>
    </body>
    </html>
  `);
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
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processRepairEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // 新しいトリガーを作成（5分ごと）
  ScriptApp.newTrigger('processRepairEmails')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  Logger.log('トリガーを設定しました');
}

