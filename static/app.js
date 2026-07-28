const loginPanel = document.getElementById("login-panel");
const appPanel = document.getElementById("app-panel");
const loginForm = document.getElementById("login-form");
const roleSelect = document.getElementById("role");
const employeeLoginFields = document.getElementById("employee-login-fields");
const sessionTitle = document.getElementById("session-title");
const sessionSubtitle = document.getElementById("session-subtitle");
const sectionSelect = document.getElementById("section-select");
const startButton = document.getElementById("start-button");
const activeSection = document.getElementById("active-section");
const chat = document.getElementById("chat");
const messageInput = document.getElementById("message-input");
const bubbleTemplate = document.getElementById("bubble-template");
const voiceButton = document.getElementById("voice-button");
const speakButton = document.getElementById("speak-button");
const endPracticeButton = document.getElementById("end-practice-button");
const micIndicator = document.getElementById("mic-indicator");
const voiceStatus = document.getElementById("voice-status");
const practiceTab = document.getElementById("practice-tab");
const adminTab = document.getElementById("admin-tab");
const modeSwitch = document.querySelector(".mode-switch");
const practiceView = document.getElementById("practice-view");
const adminView = document.getElementById("admin-view");
const rulesForm = document.getElementById("rules-form");
const authForm = document.getElementById("auth-form");
const employeesForm = document.getElementById("employees-form");
const adminSections = document.getElementById("admin-sections");
const saveAdminButton = document.getElementById("save-admin-button");
const saveAuthButton = document.getElementById("save-auth-button");
const addSectionButton = document.getElementById("add-section-button");
const importExcelInput = document.getElementById("excel-import-input");
const exportExcelButton = document.getElementById("export-excel-button");
const importExcelButton = document.getElementById("import-excel-button");
const exportReportsButton = document.getElementById("export-reports-button");
const filterReportsButton = document.getElementById("filter-reports-button");
const reportsFilters = document.getElementById("reports-filters");
const reportsSummary = document.getElementById("reports-summary");
const reportsList = document.getElementById("reports-list");
const addEmployeeButton = document.getElementById("add-employee-button");
const adminFeedback = document.getElementById("admin-feedback");
const adminLogoutButton = document.getElementById("admin-logout-button");
const adminNavButtons = [...document.querySelectorAll("[data-admin-nav]")];
const adminPanels = [...document.querySelectorAll("[data-admin-panel]")];
const employeeSearch = document.getElementById("employee-search");
const employeeCount = document.getElementById("employee-count");
const showPasswords = document.getElementById("show-passwords");
const saveAdminButtonBottom = document.getElementById("save-admin-button-bottom");

let trainingActive = false;
let lastCoachMessage = "";
let speechReady = false;
let mediaReady = false;
let isListening = false;
let mediaRecorder = null;
let mediaStream = null;
let recordedChunks = [];
let discardRecording = false;
let responsePending = false;
let analyserNode = null;
let audioContext = null;
let audioSourceNode = null;
let silenceMonitorId = null;
let autoListenTimeoutId = null;
let microphoneRecoveryTimeoutId = null;
let heardSpeech = false;
let speechStartedAt = 0;
let lastSpeechAt = 0;
let recordingStartedAt = 0;
let speakingThreshold = 0.032;
let noiseCalibrationUntil = 0;
let noiseSampleSum = 0;
let noiseSampleCount = 0;
let keywordRecognition = null;
let finishedByKeyword = false;
let currentConfig = null;
let currentSessionInfo = null;
let currentRole = null;
let currentAuth = null;
let currentReports = [];
let currentAdminSection = "accounts";
let isDirty = false;
let toastTimer = null;
let reportSort = { key: "created_at", dir: "desc" };
let currentAdminRev = null;
let awaitingCoachReply = false;
let waitingForUserReply = false;
let reportFilters = {
  employee_id: "",
  section_id: "",
  date_from: "",
  date_to: "",
};

const AUTO_SUBMIT_SILENCE_MS = 5000;
const AUTO_SUBMIT_MIN_SPEECH_MS = 1800;
const AUTO_SUBMIT_IDLE_HINT_MS = 20000;
const AUTO_SUBMIT_MAX_RECORDING_MS = 60000;
// 收音前先量這麼久的環境底噪,動態決定「算說話」的門檻(吵店自動墊高、安靜店維持靈敏)
const NOISE_CALIBRATION_MS = 300;
// 有效語音門檻 = 底噪 RMS × 此倍數,夾在地板與上限之間
// 倍數放低(1.3)+ 上限保護,避免安靜環境底噪極低時把門檻壓太低、害正常停頓被當靜音
const SPEAKING_NOISE_MULTIPLIER = 1.3;
const SPEAKING_FLOOR_RMS = 0.032;
const SPEAKING_CEILING_RMS = 0.075;
// 說過話後總時長到此即可送出(僅作異常兜底,平時靠 5 秒靜音判斷)
const AUTO_SUBMIT_SOFT_CAP_MS = 20000;
// 夥伴說出這些口令其中之一,立即送出評分,不再等靜音(用瀏覽器即時辨識偵測,僅作觸發用)
const FINISH_KEYWORDS = [
  "回答完成",
  "回答完畢",
  "回答結束",
  "我說完了",
  "我講完了",
  "說完了",
  "講完了",
  "結束回答",
  "回答好了",
];
// 記住上次登入的員工編號(不記密碼),門市共用平板可少打一次
const LAST_EMPLOYEE_ID_KEY = "diybc_training_last_employee_id";

const RULE_FIELDS = [
  { key: "max_attempts_before_answer", label: "幾次後公布答案", type: "number", group: "score", help: "員工答錯幾次後，教練直接公布標準答案。" },
  { key: "question_suffix", label: "每題結尾句", group: "talk", help: "每題唸完後接的話，例如「請開始回答」。" },
  { key: "retry_prompt", label: "未滿分重答提示", group: "talk", help: "答得不夠好時，請員工再試一次的提示語。" },
  { key: "answer_reveal_prompt", label: "公布答案前提示", group: "talk", help: "準備公布標準答案前說的話。" },
  { key: "reference_answer_intro", label: "標準答案前綴", group: "talk", help: "唸標準答案前的引言，例如「參考答案是」。" },
  { key: "pass_feedback", label: "滿分評語", group: "talk", help: "答得很好時的稱讚語。" },
  { key: "retry_feedback", label: "未滿分評語", group: "talk", help: "答得不夠好時的回饋語。" },
  { key: "pass_message", label: "過關提示", group: "talk", help: "通過一題後的提示。" },
  { key: "end_phrase", label: "結束口令", group: "talk", help: "員工說出這句話就結束本單元。" },
  { key: "summary_intro_if_empty", label: "未作答總結鼓勵", type: "textarea", group: "encourage", help: "員工幾乎沒作答時，結尾給的鼓勵。" },
  { key: "summary_encouragement", label: "結訓鼓勵語", type: "textarea", group: "encourage", help: "整個單元結束後的總結鼓勵。" },
  { key: "scoring_instruction", label: "AI 評分指令", type: "textarea", group: "advanced", help: "進階：直接調整 AI 的評分標準與風格，沒把握可先不動。" },
  { key: "assistant_role", label: "AI 角色設定", type: "textarea", group: "advanced", help: "進階：設定 AI 教練的個性與口吻。" },
];
const RULE_GROUPS = [
  { id: "score", title: "評分設定", advanced: false },
  { id: "talk", title: "提示與話術", advanced: false },
  { id: "encourage", title: "鼓勵語", advanced: false },
  { id: "advanced", title: "進階 AI 設定（沒把握可不動）", advanced: true },
];
const AUTH_FIELDS = [
  { key: "admin_password", label: "管理員密碼" },
];

function updateSessionBanner(title, subtitle, roleLabel = null) {
  sessionTitle.textContent = title;
  sessionSubtitle.textContent = subtitle;
}

// 練習進度顯示：第 x／N 題(＋上一答得分)
function updateProgressDisplay(questionNo, totalQuestions, lastScore = null) {
  if (!questionNo || !totalQuestions) return;
  const scorePart = (lastScore === null || typeof lastScore === "undefined")
    ? ""
    : `｜上一答 ${lastScore}/10 分`;
  sessionSubtitle.textContent = `第 ${questionNo}／${totalQuestions} 題${scorePart}`;
}

function setVoiceStatus(text) {
  voiceStatus.textContent = text;
}

function setMode(mode) {
  const practice = mode === "practice";
  practiceView.classList.toggle("hidden", !practice);
  adminView.classList.toggle("hidden", practice);
  if (practiceTab) practiceTab.classList.toggle("active", practice);
  if (adminTab) adminTab.classList.toggle("active", !practice);
}

function setAdminSection(section) {
  currentAdminSection = section;
  adminNavButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.adminNav === section);
  });
  adminPanels.forEach((panel) => {
    panel.classList.toggle("hidden", panel.dataset.adminPanel !== section);
  });
}

function updateRoleUi() {
  const isAdmin = currentRole === "admin";
  if (appPanel) {
    appPanel.classList.toggle("role-admin", isAdmin);
  }
  if (modeSwitch) {
    modeSwitch.classList.add("hidden");
  }
  if (isAdmin) {
    setMode("admin");
    setAdminSection(currentAdminSection);
    return;
  }
  setMode("practice");
}

function clearAutoListenTimer() {
  if (autoListenTimeoutId) {
    window.clearTimeout(autoListenTimeoutId);
    autoListenTimeoutId = null;
  }
}

function clearMicrophoneRecoveryTimer() {
  if (microphoneRecoveryTimeoutId) {
    window.clearTimeout(microphoneRecoveryTimeoutId);
    microphoneRecoveryTimeoutId = null;
  }
}

function speak(text, onDone = null) {
  if (!("speechSynthesis" in window) || !text) {
    if (typeof onDone === "function") onDone();
    return;
  }
  window.speechSynthesis.cancel();
  const utterance = new SpeechSynthesisUtterance(text);
  utterance.lang = "zh-TW";
  utterance.rate = 1.0;
  if (typeof onDone === "function") {
    utterance.onend = () => onDone();
    utterance.onerror = () => onDone();
  }
  window.speechSynthesis.speak(utterance);
}

function summarizeSpeechLabels(rawText, limit = 2) {
  const labels = rawText
    .split("、")
    .map((item) => item.trim())
    .filter(Boolean);
  if (!labels.length) return "";
  if (labels.length <= limit) return labels.join("、");
  return `${labels.slice(0, limit).join("、")}等重點`;
}

function buildCoachSpeechText(message, { autoListen = false, meta = null } = {}) {
  const lines = message
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const questionLine = lines.find((line) => /^Q\d+：/.test(line));

  // 後端有回傳結構化欄位時，優先用它組語音：
  // 讓員工「聽得到」分數與 AI 教練的具體指導，不用一直盯螢幕。
  if (meta && !meta.done) {
    if (meta.skipped) {
      return questionLine ? `已跳過，下一題。${questionLine}` : "已跳過這題。";
    }
    if (meta.passed) {
      return questionLine ? `這題通過。下一題。${questionLine}` : "這題通過。";
    }
    if (typeof meta.score === "number") {
      const spokenParts = [`這次 ${meta.score} 分。`];
      const coaching = (meta.coaching || "").trim().slice(0, 180);
      if (coaching) {
        spokenParts.push(coaching);
      }
      if (meta.revealed) {
        spokenParts.push("為使練習流暢，參考答案只顯示文字、不會唸出來。請看螢幕，照這個方向再說一次；真的卡住也可以說「跳過這題」。");
      }
      if (!coaching) {
        spokenParts.push("請再回答一次。");
      }
      return spokenParts.join("");
    }
  }

  const passLine = lines.find((line) => line.includes("這題通過"));
  const missingLine = lines.find((line) => line.startsWith("請補上：") || line.startsWith("應該補上："));
  const profanityLine = lines.find((line) => line.includes("不能帶髒話"));
  const expressionLine = lines.find((line) => line.includes("完整句子") || line.includes("更自然"));
  const revealLine = lines.find((line) => line.startsWith("參考答案："));

  if (questionLine && passLine) {
    return `這題通過。下一題。${questionLine}`;
  }

  if (autoListen) {
    const spokenParts = [];
    if (missingLine) {
      const labels = summarizeSpeechLabels(missingLine.replace(/^請補上：|^應該補上：/, "").replace(/。$/, ""));
      if (labels) spokenParts.push(`請補上${labels}。`);
    }
    if (profanityLine) {
      spokenParts.push("回答不能帶髒話。");
    }
    if (expressionLine) {
      spokenParts.push("請用完整句子，再回答一次。");
    }
    if (revealLine) {
      spokenParts.push("參考答案已顯示在畫面上，請照這個方向回答。");
    }
    if (questionLine) {
      spokenParts.push(`題目是，${questionLine}`);
    } else {
      spokenParts.push("請直接再回答一次。");
    }
    return spokenParts.join("");
  }

  if (questionLine) {
    return questionLine;
  }

  return lines[0] || message;
}

function refreshActionState() {
  // 主按鈕＝「說完了，送出」：只有正在收音時可按（按了立即送出，不用等靜音）。
  // 結束練習改由 endPracticeButton 負責，避免員工講完話誤按大按鈕整場結束。
  if (!trainingActive) {
    voiceButton.textContent = "說完了，送出";
    voiceButton.disabled = true;
  } else if (isListening) {
    voiceButton.textContent = "說完了，送出";
    voiceButton.disabled = false;
  } else if (responsePending) {
    voiceButton.textContent = "教練評分中…";
    voiceButton.disabled = true;
  } else {
    voiceButton.textContent = "等教練說完，會自動收音";
    voiceButton.disabled = true;
  }
  startButton.disabled = responsePending || isListening;
  speakButton.disabled = responsePending;
  if (endPracticeButton) {
    endPracticeButton.disabled = responsePending || !trainingActive;
  }
}

function stopListeningUi() {
  isListening = false;
  voiceButton.classList.remove("listening");
  if (micIndicator) micIndicator.classList.remove("recording");
  refreshActionState();
}

function stopSilenceMonitor() {
  stopKeywordListener();
  if (silenceMonitorId) {
    window.clearInterval(silenceMonitorId);
    silenceMonitorId = null;
  }
  if (audioSourceNode) {
    audioSourceNode.disconnect();
    audioSourceNode = null;
  }
  if (analyserNode) {
    analyserNode.disconnect();
    analyserNode = null;
  }
  if (audioContext) {
    audioContext.close().catch(() => {});
    audioContext = null;
  }
  heardSpeech = false;
  speechStartedAt = 0;
  lastSpeechAt = 0;
  recordingStartedAt = 0;
  speakingThreshold = SPEAKING_FLOOR_RMS;
  noiseCalibrationUntil = 0;
  noiseSampleSum = 0;
  noiseSampleCount = 0;
}

function releaseRecordingResources(stopStream = true) {
  stopSilenceMonitor();
  if (stopStream && mediaStream) {
    mediaStream.getTracks().forEach((track) => track.stop());
    mediaStream = null;
  }
  mediaRecorder = null;
  recordedChunks = [];
}

function scheduleMicrophoneRecovery(reason = "麥克風剛剛被中斷，正在重新連線。") {
  if (!trainingActive || responsePending || document.hidden || !waitingForUserReply) return;
  if (microphoneRecoveryTimeoutId) return;

  setVoiceStatus(reason);
  microphoneRecoveryTimeoutId = window.setTimeout(async () => {
    microphoneRecoveryTimeoutId = null;
    if (!trainingActive || responsePending || isListening || document.hidden || !waitingForUserReply) return;
    try {
      await ensureMediaStream();
      beginAutoListenAfterCoach();
    } catch (error) {
      mediaStream = null;
      setVoiceStatus(`麥克風尚未恢復：${error.message || "請再等一下或重新整理頁面。"}`);
    }
  }, 800);
}

function handleMicrophoneInterrupted(reason = "麥克風被系統中斷，正在恢復。") {
  const wasListening = isListening;
  clearAutoListenTimer();
  clearMicrophoneRecoveryTimer();
  mediaStream = null;
  if (wasListening) {
    stopRecording(true);
  } else {
    releaseRecordingResources(false);
    stopListeningUi();
  }
  if (trainingActive && waitingForUserReply) {
    scheduleMicrophoneRecovery(reason);
  }
}

function attachStreamWatchers(stream) {
  stream.getAudioTracks().forEach((track) => {
    track.onended = () => handleMicrophoneInterrupted("麥克風連線已中斷，正在重新連線。");
    track.onmute = () => {
      if (isListening) {
        setVoiceStatus("麥克風暫時被系統接管，正在等待恢復。");
      }
    };
    track.onunmute = () => {
      if (trainingActive && waitingForUserReply && !responsePending && !isListening) {
        scheduleMicrophoneRecovery("麥克風已恢復，正在重新開始收音。");
      }
    };
  });
}

async function ensureMediaStream() {
  const activeTrack = mediaStream?.getAudioTracks?.().find((track) => track.readyState === "live");
  if (activeTrack) {
    return mediaStream;
  }

  mediaStream = await navigator.mediaDevices.getUserMedia({ audio: true });
  attachStreamWatchers(mediaStream);
  return mediaStream;
}

function getRecordingMimeType() {
  if (!window.MediaRecorder || typeof window.MediaRecorder.isTypeSupported !== "function") {
    return "";
  }
  const candidates = [
    "audio/webm;codecs=opus",
    "audio/mp4",
    "audio/webm",
    "audio/ogg;codecs=opus",
  ];
  return candidates.find((candidate) => window.MediaRecorder.isTypeSupported(candidate)) || "";
}

function startSilenceMonitor(stream) {
  stopSilenceMonitor();
  const AudioContextClass = window.AudioContext || window.webkitAudioContext;
  if (!AudioContextClass) return;

  audioContext = new AudioContextClass();
  analyserNode = audioContext.createAnalyser();
  analyserNode.fftSize = 2048;
  audioSourceNode = audioContext.createMediaStreamSource(stream);
  audioSourceNode.connect(analyserNode);

  const buffer = new Uint8Array(analyserNode.fftSize);
  recordingStartedAt = Date.now();
  heardSpeech = false;
  speechStartedAt = 0;
  lastSpeechAt = 0;
  speakingThreshold = SPEAKING_FLOOR_RMS;
  noiseCalibrationUntil = recordingStartedAt + NOISE_CALIBRATION_MS;
  noiseSampleSum = 0;
  noiseSampleCount = 0;

  silenceMonitorId = window.setInterval(() => {
    if (!isListening || !analyserNode) return;
    analyserNode.getByteTimeDomainData(buffer);

    let energy = 0;
    for (const sample of buffer) {
      const normalized = (sample - 128) / 128;
      energy += normalized * normalized;
    }
    const rms = Math.sqrt(energy / buffer.length);
    const now = Date.now();

    // 開頭先量環境底噪,動態決定「算說話」的門檻:吵店自動墊高、安靜店維持靈敏
    if (now < noiseCalibrationUntil) {
      noiseSampleSum += rms;
      noiseSampleCount += 1;
      setVoiceStatus("正在校準環境音,請稍等半秒再開口。");
      return;
    }
    if (noiseSampleCount > 0) {
      const noiseAvg = noiseSampleSum / noiseSampleCount;
      // 夾在地板(0.032)與上限(0.075)之間:安靜環境吃地板、不過度靈敏;吵店才墊高但不離譜
      const dynamic = noiseAvg * SPEAKING_NOISE_MULTIPLIER;
      speakingThreshold = Math.min(SPEAKING_CEILING_RMS, Math.max(SPEAKING_FLOOR_RMS, dynamic));
      noiseSampleCount = 0; // 只算一次
    }

    const speakingNow = rms >= speakingThreshold;

    if (speakingNow) {
      heardSpeech = true;
      if (!speechStartedAt) speechStartedAt = now;
      lastSpeechAt = now;
      setVoiceStatus("教練正在聽你回答,可以慢慢講;說完停一下、按「說完了，送出」,或說「回答完成」。");
      return;
    }

    // 一般情況:偵測到有效語音後,靜音超過門檻就送
    if (heardSpeech && now - lastSpeechAt > AUTO_SUBMIT_SILENCE_MS && now - speechStartedAt > AUTO_SUBMIT_MIN_SPEECH_MS) {
      stopRecording(false);
      return;
    }

    // 保底(安全版):只有「已經靜音一段時間」+ 總時長到軟上限才送
    // 加上靜音條件後,夥伴只要還在持續講話就絕不會被 20 秒硬切,只兜底處理講完卻卡住的情況
    const softCapSilenceMs = AUTO_SUBMIT_SILENCE_MS / 2; // 2.5 秒
    if (
      heardSpeech &&
      now - speechStartedAt > AUTO_SUBMIT_MIN_SPEECH_MS &&
      now - recordingStartedAt > AUTO_SUBMIT_SOFT_CAP_MS &&
      now - lastSpeechAt > softCapSilenceMs
    ) {
      stopRecording(false);
      return;
    }

    if (!heardSpeech && now - recordingStartedAt > AUTO_SUBMIT_IDLE_HINT_MS) {
      setVoiceStatus("教練正在等你回答,請直接開口。");
    }

    if (now - recordingStartedAt > AUTO_SUBMIT_MAX_RECORDING_MS) {
      stopRecording(false);
    }
  }, 180);
}

function normalizeKeyword(text) {
  return (text || "").replace(/[\s，。、！？!?～~.,]/g, "");
}

function transcriptHasFinishKeyword(text) {
  const cleaned = normalizeKeyword(text);
  return FINISH_KEYWORDS.some((kw) => cleaned.includes(kw));
}

function startKeywordListener() {
  finishedByKeyword = false;
  const SpeechRecognitionClass = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognitionClass) return; // 不支援就靜默降級,維持原本靜音偵測

  try {
    keywordRecognition = new SpeechRecognitionClass();
    keywordRecognition.lang = "zh-TW";
    keywordRecognition.continuous = true;
    keywordRecognition.interimResults = true;
    keywordRecognition.onresult = (event) => {
      if (!isListening) return;
      for (let i = event.resultIndex; i < event.results.length; i++) {
        const transcript = event.results[i][0]?.transcript || "";
        if (transcriptHasFinishKeyword(transcript)) {
          finishedByKeyword = true;
          setVoiceStatus("已聽到「回答完成」,正在送出…");
          stopRecording(false);
          return;
        }
      }
    };
    keywordRecognition.onerror = () => {}; // 辨識失敗不影響主流程
    keywordRecognition.onend = () => {};
    keywordRecognition.start();
  } catch (_err) {
    keywordRecognition = null; // 任何例外都回退到純靜音偵測
  }
}

function stopKeywordListener() {
  if (keywordRecognition) {
    try {
      keywordRecognition.onresult = null;
      keywordRecognition.stop();
    } catch (_err) {
      // ignore
    }
    keywordRecognition = null;
  }
}

async function startRecording() {
  if (!mediaReady || responsePending || isListening || !trainingActive) return;
  try {
    if ("speechSynthesis" in window) {
      window.speechSynthesis.cancel();
    }
    const stream = await ensureMediaStream();
    const mimeType = getRecordingMimeType();
    const options = mimeType ? { mimeType } : undefined;
    mediaRecorder = new MediaRecorder(stream, options);
    recordedChunks = [];
    discardRecording = false;

    mediaRecorder.ondataavailable = (event) => {
      if (event.data && event.data.size > 0) {
        recordedChunks.push(event.data);
      }
    };

    mediaRecorder.onerror = (event) => {
      releaseRecordingResources(false);
      stopListeningUi();
      setVoiceStatus(`錄音失敗：${event.error?.message || "請再試一次。"}`);
    };

    mediaRecorder.onstop = async () => {
      const blobType = mediaRecorder?.mimeType || mimeType || recordedChunks[0]?.type || "audio/webm";
      const audioBlob = new Blob(recordedChunks, { type: blobType });
      const shouldStopStream = discardRecording && !trainingActive;
      releaseRecordingResources(shouldStopStream);
      stopListeningUi();
      if (discardRecording) {
        discardRecording = false;
        setVoiceStatus("已取消這次錄音。");
        return;
      }
      if (!audioBlob.size) {
        setVoiceStatus("沒有收到清楚的語音內容，請再說一次。");
        return;
      }
      await submitAudio(audioBlob, blobType);
    };

    mediaRecorder.start();
    startSilenceMonitor(stream);
    startKeywordListener();
    isListening = true;
    voiceButton.classList.add("listening");
    if (micIndicator) micIndicator.classList.add("recording");
    setVoiceStatus("● 收音中…可以慢慢講,說完按「說完了，送出」、停一下,或說「回答完成」。");
    refreshActionState();
  } catch (error) {
    releaseRecordingResources(false);
    stopListeningUi();
    setVoiceStatus(`無法開始錄音：${error.message || "請確認麥克風權限後再試一次。"}`);
  }
}

function stopRecording(discard = false) {
  if (!mediaRecorder || mediaRecorder.state === "inactive") {
    if (discard) {
      releaseRecordingResources(true);
      stopListeningUi();
    }
    return;
  }
  stopSilenceMonitor();
  discardRecording = discard;
  setVoiceStatus(discard ? "已取消這次錄音。" : "已收到你的回答，正在送給教練。");
  mediaRecorder.stop();
}

function toggleExitButtons(_showPracticeExit = false, showAdminExit = false) {
  adminLogoutButton.classList.toggle("hidden", !showAdminExit);
}

function prefillLastEmployeeId() {
  try {
    const savedId = window.localStorage.getItem(LAST_EMPLOYEE_ID_KEY) || "";
    const employeeIdInput = document.getElementById("employee-id");
    if (savedId && employeeIdInput && !employeeIdInput.value) {
      employeeIdInput.value = savedId;
    }
  } catch (_err) {
    // localStorage 不可用時直接略過,不影響登入
  }
}

function rememberEmployeeId(employeeId) {
  try {
    if (employeeId) {
      window.localStorage.setItem(LAST_EMPLOYEE_ID_KEY, employeeId);
    }
  } catch (_err) {
    // ignore
  }
}

async function goToLoginPage() {
  if (isDirty && currentRole === "admin" && !window.confirm("有尚未儲存的變更，確定要離開後台嗎？變更會遺失。")) {
    return;
  }
  clearDirty();
  try {
    await api("/api/logout", {});
  } catch (_err) {
    // Even if logout fails, still return to the login view locally.
  }

  if ("speechSynthesis" in window) {
    window.speechSynthesis.cancel();
  }

  currentRole = null;
  currentConfig = null;
  currentSessionInfo = null;
  currentAuth = null;
  currentReports = [];
  trainingActive = false;
  awaitingCoachReply = false;
  waitingForUserReply = false;
  clearAutoListenTimer();
  clearMicrophoneRecoveryTimer();
  stopRecording(true);
  stopListeningUi();
  toggleExitButtons(false, false);
  adminFeedback.textContent = "";
  adminFeedback.className = "status-text";
  chat.innerHTML = "";
  messageInput.value = "";
  sectionSelect.innerHTML = "";
  reportsList.innerHTML = "";
  reportsSummary.textContent = "尚未載入報告。";
  loginForm.reset();
  toggleLoginFields();
  prefillLastEmployeeId();
  appPanel.classList.add("hidden");
  loginPanel.classList.remove("hidden");
  updateSessionBanner("準備開始今天的口語訓練", "選好單元後就能直接開口。", "員工練習");
  setVoiceStatus(mediaReady ? "教練出題後會自動開始聽你回答。" : "這台裝置目前不支援站內錄音。");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function resetPracticeHome(summaryMessage = "") {
  if (currentRole === "admin") {
    return;
  }
  trainingActive = false;
  awaitingCoachReply = false;
  waitingForUserReply = false;
  clearAutoListenTimer();
  clearMicrophoneRecoveryTimer();
  stopRecording(true);
  stopListeningUi();
  toggleExitButtons(false, false);
  activeSection.textContent = "請先選擇單元。";
  messageInput.value = "";
  chat.innerHTML = "";
  const homeMessage = summaryMessage
    ? `本輪練習已結束，已返回首頁。\n\n${summaryMessage}\n\n請重新選擇練習單元並按下「開始本單元」。`
    : "請選擇一個訓練單元，然後按下「開始本單元」。";
  if (currentRole === "admin") {
    updateSessionBanner("管理與練習都已準備完成", "你可以切換後台管理，或回到練習模式開始測試。", "管理員");
  } else {
    const name = currentSessionInfo?.employee_name || "夥伴";
    const employeeId = currentSessionInfo?.employee_id ? `（${currentSessionInfo.employee_id}）` : "";
    updateSessionBanner(
      `${name}${employeeId}，準備開始今天的口語訓練`,
      "選好單元後按下開始本單元。",
      "員工練習",
    );
  }
  addBubble("coach", homeMessage);
  setVoiceStatus(mediaReady ? "教練出題後會自動開始聽你回答。" : "這台裝置目前不支援站內錄音。");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function addBubble(role, text) {
  const node = bubbleTemplate.content.firstElementChild.cloneNode(true);
  node.classList.add(role);
  node.querySelector("p").textContent = text;
  chat.appendChild(node);
  chat.scrollTop = chat.scrollHeight;
  if (role === "coach") lastCoachMessage = text;
  refreshActionState();
  return node;
}

function showError(target, message) {
  const existing = target.querySelector(".error");
  if (existing) existing.remove();
  const el = document.createElement("p");
  el.className = "error";
  el.textContent = message;
  target.appendChild(el);
}

function removePendingCoachBubble() {
  const pendingBubble = chat.querySelector(".bubble.pending");
  if (pendingBubble) pendingBubble.remove();
}

function addPendingCoachBubble(text = "教練思考中…") {
  removePendingCoachBubble();
  const node = bubbleTemplate.content.firstElementChild.cloneNode(true);
  node.classList.add("coach", "pending");
  node.querySelector("p").textContent = text;
  chat.appendChild(node);
  chat.scrollTop = chat.scrollHeight;
  return node;
}

function setResponsePending(pending) {
  responsePending = pending;
  awaitingCoachReply = pending;
  refreshActionState();
}

function showAdminFeedback(message, kind = "ok") {
  adminFeedback.textContent = message;
  adminFeedback.className = kind === "error" ? "status-text status-error" : "status-text status-ok";
  toggleExitButtons(false, kind === "ok");
}

function beginAutoListenAfterCoach() {
  clearAutoListenTimer();
  if (!trainingActive || !mediaReady || responsePending || awaitingCoachReply || !waitingForUserReply) return;
  autoListenTimeoutId = window.setTimeout(() => {
    autoListenTimeoutId = null;
    if (!trainingActive || responsePending || awaitingCoachReply || !waitingForUserReply) return;
    if ("speechSynthesis" in window && (window.speechSynthesis.speaking || window.speechSynthesis.pending)) {
      beginAutoListenAfterCoach();
      return;
    }
    startRecording();
  }, 1200);
}

function deliverCoachMessage(message, { autoListen = false, meta = null } = {}) {
  const spokenText = buildCoachSpeechText(message, { autoListen, meta });
  waitingForUserReply = autoListen;
  if (speechReady) {
    if (autoListen) {
      setVoiceStatus("請先聽教練說明，接著會開始收音。");
      speak(spokenText, () => beginAutoListenAfterCoach());
      return;
    }
    speak(spokenText);
    return;
  }

  if (autoListen) {
    setVoiceStatus("教練已出題，請直接回答。");
    beginAutoListenAfterCoach();
  }
}

async function api(path, payload, method = "POST") {
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), 15000);
  try {
    const response = await fetch(path, {
      method,
      headers: { "Content-Type": "application/json" },
      body: method === "GET" ? undefined : JSON.stringify(payload),
      signal: controller.signal,
    });
    const data = await response.json();
    if (!response.ok || data.ok === false) {
      throw new Error(data.error || "發生未預期錯誤");
    }
    return data;
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("目前連線比較慢，教練暫時沒有回應，請再試一次。");
    }
    throw error;
  } finally {
    window.clearTimeout(timeout);
  }
}

function createField(labelText, inputEl, full = false) {
  const wrapper = document.createElement("label");
  wrapper.className = full ? "field full" : "field";
  const label = document.createElement("span");
  label.textContent = labelText;
  wrapper.append(label, inputEl);
  return wrapper;
}

function createButton(text, className, handler) {
  const button = document.createElement("button");
  button.type = "button";
  button.textContent = text;
  button.className = className;
  button.addEventListener("click", handler);
  return button;
}

function createInput(value = "", rows = 2) {
  const textarea = document.createElement("textarea");
  textarea.rows = rows;
  textarea.value = value;
  return textarea;
}

function createFieldWithHelp(labelText, inputEl, helpText, full = false) {
  const wrapper = document.createElement("label");
  wrapper.className = full ? "field full" : "field";
  const label = document.createElement("span");
  label.textContent = labelText;
  wrapper.append(label);
  if (helpText) {
    const help = document.createElement("span");
    help.className = "field-help";
    help.textContent = helpText;
    wrapper.append(help);
  }
  wrapper.append(inputEl);
  return wrapper;
}

function confirmDelete(message) {
  return window.confirm(message);
}

function showToast(message, kind = "ok") {
  const toast = document.getElementById("toast");
  if (!toast) return;
  toast.textContent = message;
  toast.className = "toast " + (kind === "error" ? "toast-error" : "toast-ok");
  if (toastTimer) window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => toast.classList.add("hidden"), kind === "error" ? 5000 : 3000);
}

function markDirty() {
  if (isDirty) return;
  isDirty = true;
  const bar = document.querySelector(".admin-save-bar");
  if (bar) bar.classList.add("dirty");
  const hint = document.getElementById("admin-save-hint");
  if (hint) hint.textContent = "● 有尚未儲存的變更";
}

function clearDirty() {
  isDirty = false;
  const bar = document.querySelector(".admin-save-bar");
  if (bar) bar.classList.remove("dirty");
  const hint = document.getElementById("admin-save-hint");
  if (hint) hint.textContent = "";
}

function refreshEmployeeCount() {
  if (!employeeCount) return;
  const total = employeesForm.querySelectorAll(".point-row").length;
  employeeCount.textContent = `目前 ${total} 位員工`;
}

function applyEmployeeFilter() {
  const q = (employeeSearch && employeeSearch.value ? employeeSearch.value : "").trim().toLowerCase();
  employeesForm.querySelectorAll(".point-row").forEach((row) => {
    const id = row.querySelector('[data-employee-field="employee_id"]').value.toLowerCase();
    const name = row.querySelector('[data-employee-field="employee_name"]').value.toLowerCase();
    const match = !q || id.includes(q) || name.includes(q);
    row.classList.toggle("hidden", !match);
  });
}

function applyPasswordVisibility() {
  const show = !!(showPasswords && showPasswords.checked);
  employeesForm.querySelectorAll('[data-employee-field="password"]').forEach((inp) => {
    inp.type = show ? "text" : "password";
  });
}

async function saveAllSettings() {
  try {
    const result = await api("/api/admin/save-all", {
      auth: collectAuthData(),
      content: collectAdminData(),
      base_rev: currentAdminRev,
    });
    currentAuth = result.auth;
    currentConfig = result.content;
    if (typeof result.rev !== "undefined") currentAdminRev = result.rev;
    renderAdmin();
    renderReportFilters();
    await loadConfig();
    await loadReports();
    setMode("admin");
    setAdminSection(currentAdminSection);
    clearDirty();
    showToast("已儲存：密碼、員工、題庫與規則都更新了 ✓", "ok");
  } catch (err) {
    // 衝突（別人剛存過）時，保留目前未存的修改，提醒先重新整理
    showToast(err && err.message ? err.message : "儲存失敗，請再試一次", "error");
  }
}

function renderRulesForm() {
  if (!currentConfig) return;
  rulesForm.innerHTML = "";
  RULE_GROUPS.forEach((group) => {
    const wrap = document.createElement("div");
    wrap.className = "rule-group" + (group.advanced ? " collapsed" : "");

    const head = document.createElement("button");
    head.type = "button";
    head.className = "rule-group-head";
    head.innerHTML = `<span>${group.title}</span><span class="chevron">▾</span>`;
    head.addEventListener("click", () => wrap.classList.toggle("collapsed"));

    const body = document.createElement("div");
    body.className = "rule-group-body rules-grid";

    RULE_FIELDS.filter((f) => f.group === group.id).forEach((field) => {
      let input;
      if (field.type === "textarea") {
        const rows = field.key === "scoring_instruction" ? 12 : 3;
        input = createInput(currentConfig.rules[field.key] || "", rows);
      } else {
        input = document.createElement("input");
        input.type = field.type || "text";
        input.value = currentConfig.rules[field.key] ?? "";
      }
      input.dataset.ruleKey = field.key;
      body.append(createFieldWithHelp(field.label, input, field.help, field.type === "textarea"));
    });

    wrap.append(head, body);
    rulesForm.append(wrap);
  });
}

function renderAuthForm() {
  if (!currentAuth) return;
  authForm.innerHTML = "";
  AUTH_FIELDS.forEach((field) => {
    const input = document.createElement("input");
    input.type = "text";
    input.value = currentAuth[field.key] || "";
    input.dataset.authKey = field.key;
    authForm.append(createField(field.label, input));
  });
}

function toggleLoginFields() {
  employeeLoginFields.classList.toggle("hidden", roleSelect.value === "admin");
}

function addEmployeeEditor(employee = { employee_id: "", employee_name: "", password: "" }) {
  const row = document.createElement("div");
  row.className = "point-row";

  const idInput = document.createElement("input");
  idInput.placeholder = "員工編號";
  idInput.value = employee.employee_id || "";
  idInput.dataset.employeeField = "employee_id";
  idInput.addEventListener("input", applyEmployeeFilter);

  const nameInput = document.createElement("input");
  nameInput.placeholder = "員工姓名";
  nameInput.value = employee.employee_name || "";
  nameInput.dataset.employeeField = "employee_name";
  nameInput.addEventListener("input", applyEmployeeFilter);

  const passwordInput = document.createElement("input");
  passwordInput.placeholder = "個人密碼";
  passwordInput.value = employee.password || "";
  passwordInput.dataset.employeeField = "password";
  passwordInput.type = (showPasswords && showPasswords.checked) ? "text" : "password";

  const removeButton = createButton("刪除員工", "ghost-button mini-button", () => {
    if (confirmDelete(`確定要刪除員工「${nameInput.value || idInput.value || "這位"}」嗎？`)) {
      row.remove();
      markDirty();
      refreshEmployeeCount();
    }
  });
  row.append(idInput, nameInput, passwordInput, removeButton);
  employeesForm.appendChild(row);
}

function renderEmployeesForm() {
  if (!currentAuth) return;
  employeesForm.innerHTML = "";
  (currentAuth.employees || []).forEach((employee) => addEmployeeEditor(employee));
  refreshEmployeeCount();
  applyEmployeeFilter();
  applyPasswordVisibility();
}

function addPointEditor(pointsWrap, point = { label: "", keywords: [] }) {
  const row = document.createElement("div");
  row.className = "point-row";

  const labelInput = document.createElement("input");
  labelInput.placeholder = "重點名稱";
  labelInput.value = point.label || "";
  labelInput.dataset.pointField = "label";

  const chipWrap = document.createElement("div");
  chipWrap.className = "chip-input";
  const hidden = document.createElement("input");
  hidden.type = "hidden";
  hidden.dataset.pointField = "keywords";
  hidden.value = Array.isArray(point.keywords) ? point.keywords.join(", ") : (point.keywords || "");
  const chips = document.createElement("div");
  chips.className = "chips";
  const entry = document.createElement("input");
  entry.className = "chip-entry";
  entry.placeholder = "輸入關鍵詞後按 Enter";

  function syncHidden() {
    hidden.value = [...chips.querySelectorAll(".chip")].map((c) => c.dataset.value).join(", ");
    markDirty();
  }
  function addChip(text) {
    const v = String(text).trim();
    if (!v) return;
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.dataset.value = v;
    const t = document.createElement("span");
    t.textContent = v;
    const x = document.createElement("button");
    x.type = "button";
    x.className = "chip-x";
    x.textContent = "×";
    x.addEventListener("click", () => { chip.remove(); syncHidden(); });
    chip.append(t, x);
    chips.appendChild(chip);
  }
  (hidden.value ? hidden.value.split(",") : []).map((s) => s.trim()).filter(Boolean).forEach(addChip);
  entry.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === ",") {
      e.preventDefault();
      addChip(entry.value);
      entry.value = "";
      syncHidden();
    }
  });
  entry.addEventListener("blur", () => {
    if (entry.value.trim()) { addChip(entry.value); entry.value = ""; syncHidden(); }
  });

  chipWrap.append(chips, entry, hidden);

  const removeButton = createButton("刪除重點", "ghost-button mini-button", () => {
    if (confirmDelete("確定刪除這個重點？")) { row.remove(); markDirty(); }
  });

  row.append(labelInput, chipWrap, removeButton);
  pointsWrap.appendChild(row);
}

function addQuestionEditor(questionsWrap, question = null) {
  const questionIndex = questionsWrap.children.length + 1;
  const editor = document.createElement("article");
  editor.className = "question-editor collapsed";

  const header = document.createElement("div");
  header.className = "question-row collapsible-head";
  const title = document.createElement("strong");
  function setTitle() {
    const idx = [...questionsWrap.children].indexOf(editor) + 1 || questionIndex;
    const p = (editor.querySelector('[data-question-field="prompt"]')?.value || "").trim();
    title.textContent = `題目 ${idx}` + (p ? `：${p.slice(0, 18)}${p.length > 18 ? "…" : ""}` : "（空白）");
  }
  const headRight = document.createElement("div");
  headRight.className = "head-actions";
  const chevron = document.createElement("span");
  chevron.className = "chevron";
  chevron.textContent = "▸";
  const removeButton = createButton("刪除此題", "ghost-button mini-button", (e) => {
    e.stopPropagation();
    if (confirmDelete("確定刪除這一題？")) {
      editor.remove();
      renumberQuestions(questionsWrap);
      markDirty();
    }
  });
  headRight.append(removeButton, chevron);
  header.append(title, headRight);
  header.addEventListener("click", (e) => {
    if (e.target.closest("button")) return;
    editor.classList.toggle("collapsed");
  });

  const fields = document.createElement("div");
  fields.className = "question-fields collapsible-body";

  const promptInput = createInput(question?.prompt || "", 2);
  promptInput.dataset.questionField = "prompt";
  promptInput.addEventListener("input", setTitle);
  fields.append(createField("題目內容", promptInput, true));

  const answerInput = createInput(question?.answer || "", 3);
  answerInput.dataset.questionField = "answer";
  fields.append(createField("標準答案", answerInput, true));

  const pointsList = document.createElement("div");
  pointsList.className = "points-list";
  pointsList.dataset.role = "points-list";
  (question?.required_points || [{ label: "", keywords: [] }]).forEach((point) => addPointEditor(pointsList, point));
  const addPointButton = createButton("新增重點", "secondary-button mini-button", () => addPointEditor(pointsList));
  fields.append(pointsList, addPointButton);

  editor.append(header, fields);
  questionsWrap.appendChild(editor);
  setTitle();
}

function renumberQuestions(questionsWrap) {
  [...questionsWrap.children].forEach((editor, index) => {
    const title = editor.querySelector("strong");
    if (!title) return;
    const p = (editor.querySelector('[data-question-field="prompt"]')?.value || "").trim();
    title.textContent = `題目 ${index + 1}` + (p ? `：${p.slice(0, 18)}${p.length > 18 ? "…" : ""}` : "（空白）");
  });
}

function addSectionEditor(section = null) {
  const editor = document.createElement("section");
  editor.className = "section-editor";

  const header = document.createElement("div");
  header.className = "admin-sections-header collapsible-head";
  const title = document.createElement("strong");
  title.textContent = section?.title || "新單元";
  const headRight = document.createElement("div");
  headRight.className = "head-actions";
  const chevron = document.createElement("span");
  chevron.className = "chevron";
  chevron.textContent = "▾";
  const removeButton = createButton("刪除此單元", "ghost-button mini-button", (e) => {
    e.stopPropagation();
    if (confirmDelete(`確定刪除單元「${title.textContent}」？此單元的題目會一起刪除。`)) {
      editor.remove();
      markDirty();
    }
  });
  headRight.append(removeButton, chevron);
  header.append(title, headRight);
  header.addEventListener("click", (e) => {
    if (e.target.closest("button")) return;
    editor.classList.toggle("collapsed");
  });

  const fields = document.createElement("div");
  fields.className = "section-fields collapsible-body";

  const idInput = document.createElement("input");
  idInput.type = "hidden";
  idInput.dataset.sectionField = "id";
  idInput.value = section?.id || `section_${Math.random().toString(36).slice(2, 8)}`;

  const titleInput = document.createElement("input");
  titleInput.value = section?.title || "";
  titleInput.placeholder = "例如：顧客服務應對練習";
  titleInput.dataset.sectionField = "title";
  titleInput.addEventListener("input", () => {
    title.textContent = titleInput.value.trim() || "新單元";
  });
  fields.append(idInput, createField("單元名稱", titleInput));

  const questionsWrap = document.createElement("div");
  questionsWrap.className = "questions-list";
  (section?.questions || []).forEach((question) => addQuestionEditor(questionsWrap, question));

  const addQuestionButton = createButton("新增題目", "secondary-button mini-button", () => {
    addQuestionEditor(questionsWrap);
    markDirty();
  });

  fields.append(questionsWrap, addQuestionButton);
  editor.append(header, fields);
  adminSections.appendChild(editor);
}

function renderAdmin() {
  if (!currentConfig || !currentAuth) return;
  renderAuthForm();
  renderEmployeesForm();
  renderRulesForm();
  adminSections.innerHTML = "";
  currentConfig.sections.forEach((section) => addSectionEditor(section));
}

function renderReportFilters() {
  if (!currentAuth || !currentConfig) return;
  reportsFilters.innerHTML = "";

  const employeeSelect = document.createElement("select");
  employeeSelect.dataset.reportFilter = "employee_id";
  const defaultEmployee = document.createElement("option");
  defaultEmployee.value = "";
  defaultEmployee.textContent = "全部員工";
  employeeSelect.appendChild(defaultEmployee);
  (currentAuth?.employees || []).forEach((employee) => {
    const option = document.createElement("option");
    option.value = employee.employee_id;
    option.textContent = `${employee.employee_name} (${employee.employee_id})`;
    if (reportFilters.employee_id === employee.employee_id) option.selected = true;
    employeeSelect.appendChild(option);
  });

  const sectionSelectEl = document.createElement("select");
  sectionSelectEl.dataset.reportFilter = "section_id";
  const defaultSection = document.createElement("option");
  defaultSection.value = "";
  defaultSection.textContent = "全部單元";
  sectionSelectEl.appendChild(defaultSection);
  (currentConfig?.sections || []).forEach((section) => {
    const option = document.createElement("option");
    option.value = section.id;
    option.textContent = section.title;
    if (reportFilters.section_id === section.id) option.selected = true;
    sectionSelectEl.appendChild(option);
  });

  const fromInput = document.createElement("input");
  fromInput.type = "date";
  fromInput.value = reportFilters.date_from;
  fromInput.dataset.reportFilter = "date_from";

  const toInput = document.createElement("input");
  toInput.type = "date";
  toInput.value = reportFilters.date_to;
  toInput.dataset.reportFilter = "date_to";

  reportsFilters.append(
    createField("員工", employeeSelect),
    createField("單元", sectionSelectEl),
    createField("開始日期", fromInput),
    createField("結束日期", toInput),
  );
}

function renderReports() {
  reportsList.innerHTML = "";
  reportsSummary.textContent = `共 ${currentReports.length} 筆練習紀錄`;
  if (!currentReports.length) {
    const empty = document.createElement("p");
    empty.className = "status-text";
    empty.textContent = "目前還沒有員工完成練習。";
    reportsList.appendChild(empty);
    return;
  }
  const rows = currentReports.slice().sort((a, b) => {
    let av, bv;
    if (reportSort.key === "average_score") {
      av = Number(a.average_score) || 0; bv = Number(b.average_score) || 0;
    } else if (reportSort.key === "employee_name") {
      av = a.employee_name || ""; bv = b.employee_name || "";
    } else {
      av = a.created_at || ""; bv = b.created_at || "";
    }
    if (av < bv) return reportSort.dir === "asc" ? -1 : 1;
    if (av > bv) return reportSort.dir === "asc" ? 1 : -1;
    return 0;
  });

  const cols = [
    { key: "employee_name", label: "員工", sortable: true },
    { key: "employee_id", label: "員編", sortable: false },
    { key: "section_title", label: "單元", sortable: false },
    { key: "average_score", label: "平均分", sortable: true },
    { key: "question_count", label: "題數", sortable: false },
    { key: "created_at", label: "時間", sortable: true },
  ];

  const table = document.createElement("table");
  table.className = "report-table";
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  cols.forEach((c) => {
    const th = document.createElement("th");
    if (!c.sortable) {
      th.textContent = c.label;
    } else {
      th.className = "sortable";
      const arrow = reportSort.key === c.key ? (reportSort.dir === "asc" ? " ▲" : " ▼") : "";
      th.textContent = c.label + arrow;
      th.addEventListener("click", () => {
        if (reportSort.key === c.key) {
          reportSort.dir = reportSort.dir === "asc" ? "desc" : "asc";
        } else {
          reportSort.key = c.key;
          reportSort.dir = c.key === "employee_name" ? "asc" : "desc";
        }
        renderReports();
      });
    }
    trh.appendChild(th);
  });
  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.forEach((report) => {
    const tr = document.createElement("tr");
    const cells = [
      report.employee_name || "未記名",
      report.employee_id || "-",
      report.section_title || "-",
      report.average_score,
      report.question_count,
      report.created_at,
    ];
    cells.forEach((val, i) => {
      const td = document.createElement("td");
      td.textContent = val;
      if (i === 3) td.className = "score-cell";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  const wrap = document.createElement("div");
  wrap.className = "table-scroll";
  wrap.appendChild(table);
  reportsList.appendChild(wrap);
}

function collectAuthData() {
  const payload = {};
  authForm.querySelectorAll("[data-auth-key]").forEach((input) => {
    payload[input.dataset.authKey] = input.value.trim();
  });
  payload.employees = [...employeesForm.querySelectorAll(".point-row")].map((row) => ({
    employee_id: row.querySelector('[data-employee-field="employee_id"]').value.trim(),
    employee_name: row.querySelector('[data-employee-field="employee_name"]').value.trim(),
    password: row.querySelector('[data-employee-field="password"]').value.trim(),
  }));
  return payload;
}

function collectAdminData() {
  const rules = {};
  rulesForm.querySelectorAll("[data-rule-key]").forEach((input) => {
    rules[input.dataset.ruleKey] = input.value;
  });

  const sections = [...adminSections.querySelectorAll(".section-editor")].map((sectionEl, sectionIndex) => {
    const sectionFields = sectionEl.querySelectorAll("[data-section-field]");
    const id = sectionFields[0].value.trim() || `section_${sectionIndex + 1}`;
    const title = sectionFields[1].value.trim();

    const questions = [...sectionEl.querySelectorAll(".question-editor")].map((questionEl) => {
      const questionFields = questionEl.querySelectorAll("[data-question-field]");
      const points = [...questionEl.querySelectorAll(".point-row")].map((pointEl) => {
        const label = pointEl.querySelector('[data-point-field="label"]').value.trim();
        const keywords = pointEl
          .querySelector('[data-point-field="keywords"]')
          .value
          .split(",")
          .map((item) => item.trim())
          .filter(Boolean);
        return { label, keywords };
      });

      return {
        prompt: questionFields[0].value.trim(),
        answer: questionFields[1].value.trim(),
        required_points: points,
      };
    });

    return { id, title, questions };
  });

  return { rules, sections };
}

async function loadConfig() {
  const response = await fetch("/api/config");
  const data = await response.json();
  currentRole = data.role;
  currentSessionInfo = data;
  updateRoleUi();
  const name = data.employee_name || "夥伴";
  const employeeId = data.employee_id ? `（${data.employee_id}）` : "";
  const isAdmin = data.role === "admin";
  if (isAdmin) {
    updateSessionBanner("後台管理", "登入後可直接管理題庫、帳號與規則。", "管理員");
  } else {
    updateSessionBanner(
      `${name}${employeeId}，歡迎回來`,
      "選好單元後按下開始本單元。",
      "員工練習",
    );
  }
  sectionSelect.innerHTML = "";
  for (const section of data.sections) {
    const option = document.createElement("option");
    option.value = section.id;
    option.textContent = section.title;
    sectionSelect.appendChild(option);
  }
  messageInput.value = "";
}

async function loadAdminContent() {
  const data = await api("/api/admin/content", null, "GET");
  currentConfig = data.content;
  if (typeof data.rev !== "undefined") currentAdminRev = data.rev;
}

async function loadAdminAuth() {
  const data = await api("/api/admin/auth", null, "GET");
  currentAuth = data.auth;
  if (typeof data.rev !== "undefined") currentAdminRev = data.rev;
}

async function loadReports() {
  const params = new URLSearchParams(reportFilters);
  const data = await api(`/api/admin/reports?${params.toString()}`, null, "GET");
  currentReports = data.reports;
  renderReports();
}

async function loadAdminView() {
  await Promise.all([loadAdminAuth(), loadAdminContent()]);
  renderAdmin();
  renderReportFilters();
  await loadReports();
  setAdminSection(currentAdminSection);
  clearDirty();
}

function setupVoice() {
  speechReady = "speechSynthesis" in window;
  mediaReady = Boolean(window.MediaRecorder && navigator.mediaDevices?.getUserMedia);
  if (!mediaReady) {
    setVoiceStatus("這台裝置目前不支援站內錄音，請改用新版 Chrome 或 Safari。");
  } else if (!speechReady) {
    setVoiceStatus("可錄音作答，但這台裝置不支援自動朗讀教練回覆。");
  } else {
    setVoiceStatus("教練出題後會自動開始聽你回答。");
  }
  refreshActionState();
}

function buildAudioFilename(blobType) {
  if (blobType.includes("mp4")) return "reply.m4a";
  if (blobType.includes("ogg")) return "reply.ogg";
  return "reply.webm";
}

async function submitAudio(audioBlob, blobType) {
  if (!trainingActive) return;

  waitingForUserReply = false;
  setResponsePending(true);
  let restartListeningAfterError = false;
  removePendingCoachBubble();
  addPendingCoachBubble("已收到你的回答,教練思考中…");
  setVoiceStatus("已送出,教練思考中,請稍候…");

  try {
    const formData = new FormData();
    formData.append("audio", audioBlob, buildAudioFilename(blobType));

    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), 45000);
    let data;
    try {
      const response = await fetch("/api/respond-audio", {
        method: "POST",
        body: formData,
        signal: controller.signal,
      });
      data = await response.json();
      if (!response.ok || data.ok === false) {
        throw new Error(data.error || "語音送出失敗");
      }
    } catch (error) {
      if (error.name === "AbortError") {
        throw new Error("教練這次聽得比較久，請稍候再試一次。");
      }
      throw error;
    } finally {
      window.clearTimeout(timeout);
    }

    removePendingCoachBubble();
    let shownTranscript = data.transcript || "已收到你的語音回答。";
    if (finishedByKeyword && data.transcript) {
      // 把結尾口令詞從顯示文字中去除,讓對話泡泡乾淨
      shownTranscript = FINISH_KEYWORDS.reduce(
        (acc, kw) => acc.replace(new RegExp(kw + "[。.!！~～\\s]*$"), ""),
        data.transcript,
      ).trim() || data.transcript;
    }
    finishedByKeyword = false;
    addBubble("user", shownTranscript);
    addBubble("coach", data.message);
    if (data.done) {
      await loadConfig();
      resetPracticeHome(data.message);
      return;
    }
    updateProgressDisplay(data.question_no, data.total_questions, data.score);
    awaitingCoachReply = false;
    deliverCoachMessage(data.message, { autoListen: true, meta: data });
  } catch (err) {
    removePendingCoachBubble();
    addBubble("coach", `這次沒有成功收到完整語音。\n原因：${err.message}\n請重新錄一次。`);
    setVoiceStatus("這次錄音沒有成功送出，教練會繼續等待你的下一次回答。");
    showError(appPanel, err.message);
    waitingForUserReply = true;
    restartListeningAfterError = trainingActive && mediaReady;
  } finally {
    setResponsePending(false);
    if (restartListeningAfterError) {
      beginAutoListenAfterCoach();
    }
  }
}

async function submitMessage(rawMessage) {
  if (!trainingActive) {
    showError(appPanel, "請先開始一個練習單元。");
    return;
  }

  if (responsePending) {
    setVoiceStatus("教練正在回應上一句，請稍等一下。");
    return;
  }

  const message = rawMessage.trim();
  if (!message) return;

  waitingForUserReply = false;
  addBubble("user", message);
  messageInput.value = "";
  setResponsePending(true);
  let restartListeningAfterError = false;
  removePendingCoachBubble();
  addPendingCoachBubble("教練思考中…");

  try {
    const result = await api("/api/respond", { message });
    removePendingCoachBubble();
    addBubble("coach", result.message);
    if (result.done) {
      await loadConfig();
      resetPracticeHome(result.message);
      return;
    }
    updateProgressDisplay(result.question_no, result.total_questions, result.score);
    awaitingCoachReply = false;
    deliverCoachMessage(result.message, { autoListen: true, meta: result });
  } catch (err) {
    removePendingCoachBubble();
    addBubble("coach", `目前暫時沒有收到教練回應。\n原因：${err.message}\n請再說一次。`);
    setVoiceStatus("這次沒有成功取得教練回應，可再試一次。");
    showError(appPanel, err.message);
    waitingForUserReply = true;
    restartListeningAfterError = trainingActive && mediaReady;
  } finally {
    setResponsePending(false);
    if (restartListeningAfterError) {
      beginAutoListenAfterCoach();
    }
  }
}

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const error = loginPanel.querySelector(".error");
  if (error) error.remove();

  try {
    const password = new FormData(loginForm).get("password");
    const role = new FormData(loginForm).get("role");
    const employee_id = new FormData(loginForm).get("employee_id");
    const login = await api("/api/login", { password, role, employee_id });
    currentRole = login.role;
    if (currentRole === "employee") {
      rememberEmployeeId(login.employee_id);
    }
    loginPanel.classList.add("hidden");
    appPanel.classList.remove("hidden");
    toggleExitButtons(false, false);
    await loadConfig();
    if (currentRole === "admin") {
      currentAdminSection = "accounts";
      await loadAdminView();
      setMode("admin");
      showAdminFeedback("管理員已登入，可直接在這裡進行設定。");
      return;
    }
    resetPracticeHome();
    if (speechReady) speak("登入成功。請選好單元，然後按下開始本單元。");
  } catch (err) {
    showError(loginPanel, err.message);
  }
});

practiceTab.addEventListener("click", () => setMode("practice"));
adminTab.addEventListener("click", async () => {
  if (currentRole !== "admin") return;
  await loadAdminView();
  setMode("admin");
});

adminNavButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setAdminSection(button.dataset.adminNav);
  });
});

startButton.addEventListener("click", async () => {
  try {
    const result = await api("/api/start", { section_id: sectionSelect.value });
    trainingActive = true;
    clearAutoListenTimer();
    chat.innerHTML = "";
    activeSection.textContent = result.title;
    updateSessionBanner(
      `目前練習：${result.title}`,
      "教練說完後會自動開始聽你回答。",
      currentRole === "admin" ? "管理員測試" : "員工練習",
    );
    updateProgressDisplay(result.question_no, result.total_questions);
    addBubble("coach", result.message);
    deliverCoachMessage(result.message, { autoListen: true });
  } catch (err) {
    showError(appPanel, err.message);
  }
});

async function stopPracticeSession() {
  if (!trainingActive || responsePending) return;

  clearAutoListenTimer();
  stopRecording(true);
  removePendingCoachBubble();
  setResponsePending(true);

  try {
    const result = await api("/api/respond", { message: currentSessionInfo?.rules?.end_phrase || "練習結束" });
    removePendingCoachBubble();
    addBubble("coach", result.message);
    await loadConfig();
    resetPracticeHome(result.message);
  } catch (err) {
    removePendingCoachBubble();
    addBubble("coach", `目前無法結束這輪練習。\n原因：${err.message}`);
    setVoiceStatus("目前無法結束練習，請稍後再試一次。");
    showError(appPanel, err.message);
  } finally {
    setResponsePending(false);
  }
}

// 大按鈕＝提前送出這次回答(不用等 5 秒靜音);還沒開口時按會先提醒。
voiceButton.addEventListener("click", () => {
  if (!isListening) return;
  if (!heardSpeech) {
    setVoiceStatus("還沒聽到你的聲音，請先開口回答，說完再按「說完了，送出」。");
    return;
  }
  stopRecording(false);
});

// 結束練習改為獨立按鈕＋確認,避免誤觸直接斷練
if (endPracticeButton) {
  endPracticeButton.addEventListener("click", async () => {
    if (!trainingActive || responsePending) return;
    if (!window.confirm("確定要結束本輪練習嗎？結束後會直接產生總結。")) return;
    await stopPracticeSession();
  });
}

speakButton.addEventListener("click", async () => {
  await goToLoginPage();
});

addSectionButton.addEventListener("click", () => {
  addSectionEditor();
  markDirty();
  const secs = adminSections.querySelectorAll(".section-editor");
  const last = secs[secs.length - 1];
  if (last) last.scrollIntoView({ behavior: "smooth", block: "center" });
});
addEmployeeButton.addEventListener("click", () => {
  addEmployeeEditor();
  refreshEmployeeCount();
  applyPasswordVisibility();
  markDirty();
  const rows = employeesForm.querySelectorAll(".point-row");
  const last = rows[rows.length - 1];
  if (last) {
    last.scrollIntoView({ behavior: "smooth", block: "center" });
    const firstInput = last.querySelector("input");
    if (firstInput) firstInput.focus();
  }
});
roleSelect.addEventListener("change", toggleLoginFields);


adminLogoutButton.addEventListener("click", async () => {
  await goToLoginPage();
});

exportExcelButton.addEventListener("click", () => {
  window.location.href = "/api/admin/export.xlsx";
});

exportReportsButton.addEventListener("click", () => {
  const params = new URLSearchParams(reportFilters);
  window.location.href = `/api/admin/reports.xlsx?${params.toString()}`;
});

filterReportsButton.addEventListener("click", async () => {
  reportsFilters.querySelectorAll("[data-report-filter]").forEach((input) => {
    reportFilters[input.dataset.reportFilter] = input.value;
  });
  await loadReports();
});

importExcelButton.addEventListener("click", () => {
  importExcelInput.click();
});

importExcelInput.addEventListener("change", async () => {
  const [file] = importExcelInput.files;
  if (!file) return;
  try {
    const arrayBuffer = await file.arrayBuffer();
    const response = await fetch("/api/admin/import", {
      method: "POST",
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      body: arrayBuffer,
    });
    const data = await response.json();
    if (!response.ok || data.ok === false) {
      throw new Error(data.error || "匯入失敗");
    }
    currentConfig = data.content;
    renderAdmin();
    renderReportFilters();
    await loadConfig();
    await loadReports();
    showError(adminView, "Excel 匯入成功，新的題庫與規則已生效。");
  } catch (err) {
    showError(adminView, err.message);
  } finally {
    importExcelInput.value = "";
  }
});

saveAdminButton.addEventListener("click", saveAllSettings);
if (saveAdminButtonBottom) saveAdminButtonBottom.addEventListener("click", saveAllSettings);

document.addEventListener("visibilitychange", () => {
  if (!document.hidden && trainingActive && waitingForUserReply && !responsePending && !isListening) {
    scheduleMicrophoneRecovery("已回到練習畫面，正在確認麥克風。");
  }
});

window.addEventListener("focus", () => {
  if (trainingActive && waitingForUserReply && !responsePending && !isListening) {
    scheduleMicrophoneRecovery("已回到練習畫面，正在確認麥克風。");
  }
});

window.addEventListener("pageshow", () => {
  if (trainingActive && waitingForUserReply && !responsePending && !isListening) {
    scheduleMicrophoneRecovery("已回到練習畫面，正在確認麥克風。");
  }
});

setupVoice();
toggleLoginFields();
prefillLastEmployeeId();
refreshActionState();

// ---- 後台友善化：搜尋、密碼顯示、未存提醒 ----
if (employeeSearch) employeeSearch.addEventListener("input", applyEmployeeFilter);
if (showPasswords) showPasswords.addEventListener("change", applyPasswordVisibility);

if (adminView) {
  adminView.addEventListener("input", (event) => {
    const t = event.target;
    if (t.id === "employee-search" || (t.closest && t.closest("#reports-filters"))) return;
    markDirty();
  });
  adminView.addEventListener("change", (event) => {
    const t = event.target;
    if (t.id === "show-passwords" || (t.closest && t.closest("#reports-filters"))) return;
    markDirty();
  });
}

window.addEventListener("beforeunload", (event) => {
  if (isDirty && currentRole === "admin") {
    event.preventDefault();
    event.returnValue = "";
  }
});
