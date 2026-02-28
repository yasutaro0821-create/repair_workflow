/**
 * voice.js - 音声入力モジュール（Web Speech API）
 * iOS Safari / Chrome対応
 */
const VoiceModule = (() => {
  let recognition = null;
  let isListening = false;

  function init() {
    const voiceBtn = document.getElementById('voiceBtn');
    const voiceStatus = document.getElementById('voiceStatus');

    // Web Speech API対応チェック
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
      voiceBtn.style.display = 'none';
      return;
    }

    recognition = new SpeechRecognition();
    recognition.lang = 'ja-JP';
    recognition.continuous = true;
    recognition.interimResults = true;

    recognition.onstart = () => {
      isListening = true;
      voiceBtn.classList.add('listening');
      voiceStatus.classList.remove('hidden');
      voiceStatus.textContent = '音声認識中...';
    };

    recognition.onresult = (event) => {
      const textarea = document.getElementById('description');
      let interimTranscript = '';
      let finalTranscript = '';

      for (let i = event.resultIndex; i < event.results.length; i++) {
        const transcript = event.results[i][0].transcript;
        if (event.results[i].isFinal) {
          finalTranscript += transcript;
        } else {
          interimTranscript += transcript;
        }
      }

      if (finalTranscript) {
        // 確定テキストをテキストエリアに追加
        const current = textarea.value;
        textarea.value = current + (current ? '\n' : '') + finalTranscript;
      }

      if (interimTranscript) {
        voiceStatus.textContent = `認識中: ${interimTranscript}`;
      }
    };

    recognition.onerror = (event) => {
      if (event.error === 'no-speech') {
        voiceStatus.textContent = '音声が検出されませんでした';
      } else if (event.error === 'not-allowed') {
        voiceStatus.textContent = 'マイクの使用が許可されていません';
      } else {
        voiceStatus.textContent = `エラー: ${event.error}`;
      }
      stopListening();
    };

    recognition.onend = () => {
      if (isListening) {
        // 自動停止の場合は再開を試みる
        try {
          recognition.start();
        } catch (e) {
          stopListening();
        }
      }
    };

    voiceBtn.addEventListener('click', toggleListening);
  }

  function toggleListening() {
    if (isListening) {
      stopListening();
    } else {
      startListening();
    }
  }

  function startListening() {
    if (!recognition) return;
    try {
      recognition.start();
    } catch (e) {
      // 既に開始されている場合
    }
  }

  function stopListening() {
    isListening = false;
    if (recognition) {
      try {
        recognition.stop();
      } catch (e) {
        // 既に停止されている場合
      }
    }
    const voiceBtn = document.getElementById('voiceBtn');
    const voiceStatus = document.getElementById('voiceStatus');
    voiceBtn.classList.remove('listening');
    setTimeout(() => {
      voiceStatus.classList.add('hidden');
    }, 2000);
  }

  function stop() {
    stopListening();
  }

  return { init, stop };
})();
