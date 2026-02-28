/**
 * app.js - メインアプリケーションロジック
 * フォーム送信・API通信・オフライン対応
 */
(() => {
  // GAS Web App URL（デプロイ後に更新してください）
  const GAS_API_URL = 'https://script.google.com/macros/s/AKfycbwSaq3Nt11DNMcwITRzUWR2aEzRiARa-davYZSEUZASt3_Jsy0WyMkHT1dy4ZtpQa0zkQ/exec';

  // IndexedDB設定（オフラインキュー用）
  const DB_NAME = 'mtinn-repair';
  const STORE_NAME = 'queue';

  // 初期化
  document.addEventListener('DOMContentLoaded', () => {
    CameraModule.init();
    VoiceModule.init();
    initForm();
    registerServiceWorker();
    processOfflineQueue();
  });

  // Service Worker登録
  function registerServiceWorker() {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('./sw.js').catch((err) => {
        console.warn('SW registration failed:', err);
      });
    }
  }

  // フォーム初期化
  function initForm() {
    const form = document.getElementById('repairForm');
    const reporterSelect = document.getElementById('reporter');
    const reporterCustom = document.getElementById('reporterCustom');
    const newReportBtn = document.getElementById('newReportBtn');
    const retryBtn = document.getElementById('retryBtn');

    // 「その他」選択時にテキスト入力を表示
    reporterSelect.addEventListener('change', () => {
      if (reporterSelect.value === 'その他') {
        reporterCustom.classList.remove('hidden');
        reporterCustom.required = true;
        reporterCustom.focus();
      } else {
        reporterCustom.classList.add('hidden');
        reporterCustom.required = false;
        reporterCustom.value = '';
      }
    });

    // 保存された報告者名を復元
    const savedReporter = localStorage.getItem('lastReporter');
    if (savedReporter) {
      const option = Array.from(reporterSelect.options).find((o) => o.value === savedReporter);
      if (option) {
        reporterSelect.value = savedReporter;
      }
    }

    // フォーム送信
    form.addEventListener('submit', handleSubmit);

    // 新規報告ボタン
    newReportBtn.addEventListener('click', resetForm);

    // リトライボタン
    retryBtn.addEventListener('click', () => {
      showSection('form');
    });
  }

  // フォーム送信処理
  async function handleSubmit(event) {
    event.preventDefault();

    const reporterSelect = document.getElementById('reporter');
    const reporterCustom = document.getElementById('reporterCustom');
    const description = document.getElementById('description').value.trim();

    // バリデーション
    let reporter = reporterSelect.value;
    if (reporter === 'その他') {
      reporter = reporterCustom.value.trim();
      if (!reporter) {
        alert('報告者名を入力してください');
        reporterCustom.focus();
        return;
      }
    }
    if (!reporter) {
      alert('報告者を選択してください');
      return;
    }

    const photos = CameraModule.getPhotos();
    if (!description && photos.length === 0) {
      alert('写真または修繕内容の説明を入力してください');
      return;
    }

    // 報告者名を保存
    localStorage.setItem('lastReporter', reporterSelect.value === 'その他' ? 'その他' : reporter);

    // 音声入力を停止
    VoiceModule.stop();

    // 送信データ作成
    const payload = {
      action: 'repair_report',
      reporter: reporter,
      description: description || '（写真のみ）',
      images: photos
    };

    showSection('loading');

    // オンラインなら直接送信、オフラインならキューに保存
    if (navigator.onLine) {
      await sendReport(payload);
    } else {
      await saveToQueue(payload);
      showSection('offline');
    }
  }

  // API送信
  // GAS Web Appは302リダイレクトでCORSヘッダーを返さないため、
  // mode: 'no-cors' を使用。レスポンスは読めないが、リクエストは届く。
  async function sendReport(payload) {
    try {
      const response = await fetch(GAS_API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify(payload),
        mode: 'no-cors',
        redirect: 'follow'
      });

      // no-corsモードではopaqueレスポンス（status=0）が返る
      // リクエスト自体はGASに届いて処理されるため、送信成功とする
      document.getElementById('successMessage').textContent =
        '修繕報告を送信しました。AI解析・稟議書生成が自動で実行されます。';
      showSection('success');
    } catch (error) {
      console.error('送信エラー:', error);

      // ネットワークエラーの場合はキューに保存
      if (!navigator.onLine || error.name === 'TypeError') {
        await saveToQueue(payload);
        showSection('offline');
        return;
      }

      document.getElementById('errorMessage').textContent = error.message;
      showSection('error');
    }
  }

  // セクション表示切替
  function showSection(name) {
    const form = document.getElementById('repairForm');
    const loading = document.getElementById('loading');
    const success = document.getElementById('success');
    const error = document.getElementById('error');
    const offline = document.getElementById('offlineNotice');

    form.classList.toggle('hidden', name !== 'form');
    loading.classList.toggle('hidden', name !== 'loading');
    success.classList.toggle('hidden', name !== 'success');
    error.classList.toggle('hidden', name !== 'error');
    offline.classList.toggle('hidden', name !== 'offline');
  }

  // フォームリセット
  function resetForm() {
    document.getElementById('description').value = '';
    CameraModule.clear();
    showSection('form');
  }

  // === IndexedDB オフラインキュー ===

  function openDB() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = (e) => {
        e.target.result.createObjectStore(STORE_NAME, { keyPath: 'id', autoIncrement: true });
      };
      request.onsuccess = (e) => resolve(e.target.result);
      request.onerror = (e) => reject(e.target.error);
    });
  }

  async function saveToQueue(payload) {
    try {
      const db = await openDB();
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).add({
        payload: payload,
        timestamp: new Date().toISOString()
      });
      await new Promise((resolve, reject) => {
        tx.oncomplete = resolve;
        tx.onerror = reject;
      });
    } catch (e) {
      console.error('キュー保存エラー:', e);
    }
  }

  async function processOfflineQueue() {
    if (!navigator.onLine) return;

    try {
      const db = await openDB();
      const tx = db.transaction(STORE_NAME, 'readonly');
      const store = tx.objectStore(STORE_NAME);
      const request = store.getAll();

      request.onsuccess = async () => {
        const items = request.result;
        if (items.length === 0) return;

        for (const item of items) {
          try {
            await sendReport(item.payload);
            // 送信成功したらキューから削除
            const deleteTx = db.transaction(STORE_NAME, 'readwrite');
            deleteTx.objectStore(STORE_NAME).delete(item.id);
          } catch (e) {
            console.warn('キュー再送信エラー:', e);
            break; // 失敗したら残りは次回に
          }
        }
      };
    } catch (e) {
      console.error('キュー処理エラー:', e);
    }
  }

  // オンライン復帰時にキューを処理
  window.addEventListener('online', () => {
    processOfflineQueue();
  });
})();
