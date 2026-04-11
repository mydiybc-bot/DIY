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
let currentConfig = null;
let currentSessionInfo = null;
let currentRole = null;
let currentAuth = null;
let currentReports = [];
let currentAdminSection = "accounts";
let awaitingCoachReply = false;
let waitingForUserReply = false;
let reportFilters = {
  employee_id: "",
  section_id: "",
  date_from: "",
  date_to: "",
};

const AUTO_SUBMIT_SILENCE_MS = 2200;
const AUTO_SUBMIT_MIN_SPEECH_MS = 1800;
const AUTO_SUBMIT_IDLE_HINT_MS = 12000;
const AUTO_SUBMIT_MAX_RECORDING_MS = 45000;

const RULE_FIELDS = [
  { key: "assistant_role", label: "AI 角色設定", type: "textarea" },
  { key: "question_suffix", label: "每題結尾句" },
  { key: "retry_prompt", label: "未滿分重答提示" },
  { key: "answer_reveal_prompt", label: "公布答案後提示" },
  { key: "reference_answer_intro", label: "標準答案前綴" },
  { key: "pass_feedback", label: "滿分評語" },
  { key: "retry_feedback", label: "未滿分評語" },
  { key: "pass_message", label: "過關提示" },
  { key: "summary_intro_if_empty", label: "未作答總結鼓勵", type: "textarea" },
  { key: "summary_encouragement", label: "結訓鼓勵語", type: "textarea" },
  { key: "max_attempts_before_answer", label: "幾次後公布答案", type: "number" },
  { key: "end_phrase", label: "結束口令" },
];
const AUTH_FIELDS = [
  { key: "admin_password", label: "管理員密碼" },
];

function updateSessionBanner(title, subtitle, roleLabel = null) {
  sessionTitle.textContent = title;
  sessionSubtitle.textContent = subtitle;
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
  utterance.rate = 1.25;
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

function buildCoachSpeechText(message, { autoListen = false } = {}) {
  const lines = message
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const questionLine = lines.find((line) => /^Q\d+：/.test(line));
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
  voiceButton.disabled = responsePending || !trainingActive;
  startButton.disabled = responsePending || isListening;
  speakButton.disabled = responsePending;
}

function stopListeningUi() {
  isListening = false;
  voiceButton.classList.remove("listening");
  voiceButton.textContent = "停止練習";
  refreshActionState();
}

function stopSilenceMonitor() {
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
    const speakingNow = rms >= 0.032;

    if (speakingNow) {
      heardSpeech = true;
      if (!speechStartedAt) speechStartedAt = now;
      lastSpeechAt = now;
      setVoiceStatus("教練正在聽你回答，停頓約 2 秒後才會自動送出。");
      return;
    }

    if (heardSpeech && now - lastSpeechAt > AUTO_SUBMIT_SILENCE_MS && now - speechStartedAt > AUTO_SUBMIT_MIN_SPEECH_MS) {
      stopRecording(false);
      return;
    }

    if (!heardSpeech && now - recordingStartedAt > AUTO_SUBMIT_IDLE_HINT_MS) {
      setVoiceStatus("教練正在等你回答，請直接開口。");
    }

    if (now - recordingStartedAt > AUTO_SUBMIT_MAX_RECORDING_MS) {
      stopRecording(false);
    }
  }, 180);
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
    isListening = true;
    voiceButton.classList.add("listening");
    voiceButton.textContent = "停止練習";
    setVoiceStatus("教練正在聽你回答，停頓約 2 秒後才會自動送出。");
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

async function goToLoginPage() {
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
  appPanel.classList.add("hidden");
  loginPanel.classList.remove("hidden");
  updateSessionBanner("準備開始今天的口語訓練", "選好單元後就能直接開口。", "員工練習");
  setVoiceStatus(mediaReady ? "教練出題後會自動開始聽你回答。" : "這台裝置目前不支援开內錄音。");
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
    ? `本輾練習已結束，已返回首頁。\n\n${summaryMessage}\n\n請重新選擇練習單元並按下「開始本單元」。`
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

function deliverCoachMessage(message, { autoListen = false } = {}) {
  const spokenText = buildCoachSpeechText(message, { autoListen });
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

function renderRulesForm() {
  if (!currentConfig) return;
  rulesForm.innerHTML = "";
  RULE_FIELDS.forEach((field) => {
    let input;
    if (field.type === "textarea") {
      input = createInput(currentConfig.rules[field.key] || "", 3);
    } else {
      input = document.createElement("input");
      input.type = field.type || "text";
      input.value = currentConfig.rules[field.key] ?? "";
    }
    input.dataset.ruleKey = field.key;
    rulesForm.append(createField(field.label, input, field.type === "textarea"));
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

  const nameInput = document.createElement("input");
  nameInput.placeholder = "員工姓名";
  nameInput.value = employee.employee_name || "";
  nameInput.dataset.employeeField = "employee_name";

  const passwordInput = document.createElement("input");
  passwordInput.placeholder = "個人密碼";
  passwordInput.value = employee.password || "";
  passwordInput.dataset.employeeField = "password";

  const removeButton = createButton("刪除員工", "ghost-button mini-button", () => row.remove());
  row.append(idInput, nameInput, passwordInput, removeButton);
  employeesForm.appendChild(row);
}

function renderEmployeesForm() {
  if (!currentAuth) return;
  employeesForm.innerHTML = "";
  (currentAuth.employees || []).forEach((employee) => addEmployeeEditor(employee));
}

function addPointEditor(pointsWrap, point = { label: "", keywords: [] }) {
  const row = document.createElement("div");
  row.className = "point-row";

  const labelInput = document.createElement("input");
  labelInput.placeholder = "重點名稱";
  labelInput.value = point.label || "";
  labelInput.dataset.pointField = "label";

  const keywordsInput = document.createElement("input");
  keywordsInput.placeholder = "關鍵詞，用逗號分隔";
  keywordsInput.value = Array.isArray(point.keywords) ? point.keywords.join(", ") : "";
  keywordsInput.dataset.pointField = "keywords";

  const removeButton = createButton("刪除重點", "ghost-button mini-button", () => row.remove());

  row.append(labelInput, keywordsInput, removeButton);
  pointsWrap.appendChild(row);
}

function addQuestionEditor(questionsWrap, question = null) {
  const questionIndex = questionsWrap.children.length + 1;
  const editor = document.createElement("article");
  editor.className = "question-editor";

  const header = document.createElement("div");
  header.className = "question-row";
  const title = document.createElement("strong");
  title.textContent = `題目 ${questionIndex}`;
  const removeButton = createButton("刪除此題", "ghost-button mini-button", () => {
    editor.remove();
    renumberQuestions(questionsWrap);
  });
  header.append(title, removeButton);

  const fields = document.createElement("div");
  fields.className = "question-fields";

  const promptInput = createInput(question?.prompt || "", 2);
  promptInput.dataset.questionField = "prompt";
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
}

function renumberQuestions(questionsWrap) {
  [...questionsWrap.children].forEach((editor, index) => {
    const title = editor.querySelector("strong");
    title.textContent = `題目 ${index + 1}`;
  });
}

function addSectionEditor(section = null) {
  const editor = document.createElement("section");
  editor.className = "section-editor";

  const header = document.createElement("div");
  header.className = "admin-sections-header";
  const title = document.createElement("strong");
  title.textContent = section?.title || "新單元";
  const removeButton = createButton("刪除此單元", "ghost-button mini-button", () => editor.remove());
  header.append(title, removeButton);

  const fields = document.createElement("div");
  fields.className = "section-fields";
  const idInput = document.createElement("input");
  idInput.value = section?.id || "";
  idInput.placeholder = "section_id";
  idInput.dataset.sectionField = "id";
  const titleInput = document.createElement("input");
  titleInput.value = section?.title || "";
  titleInput.placeholder = "單元名稱";
  titleInput.dataset.sectionField = "title";
  titleInput.addEventListener("input", () => {
    title.textContent = titleInput.value.trim() || "新單元";
  });
  fields.append(createField("單元 ID", idInput), createField("單元名稱", titleInput));

  const questionsWrap = document.createElement("div");
  questionsWrap.className = "questions-list";
  (section?.questions || []).forEach((question) => addQuestionEditor(questionsWrap, question));

  const addQuestionButton = createButton("新增題目", "secondary-button mini-button", () => addQuestionEditor(questionsWrap));

  editor.append(header, fields, questionsWrap, addQuestionButton);
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
  currentReports
    .slice()
    .reverse()
    .forEach((report) => {
      const card = document.createElement("article");
      card.className = "section-editor";
      const title = document.createElement("strong");
      title.textContent = `${report.employee_name || "未記名"} (${report.employee_id || "-"}) | ${report.section_title} | 平均 ${report.average_score} 分`;
      const meta = document.createElement("p");
      meta.className = "status-text";
      meta.textContent = `${report.created_at} | 題數 ${report.question_count} | ${report.scores.join("；") || "尚未作答"}`;
      card.append(title, meta);
      reportsList.appendChild(card);
    });
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
}

async function loadAdminAuth() {
  const data = await api("/api/admin/auth", null, "GET");
  currentAuth = data.auth;
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
}

function setupVoice() {
  speechReady = "speechSynthesis" in window;
  mediaReady = Boolean(window.MediaRecorder && navigator.mediaDevices?.getUserMedia);
  if (!mediaReady) {
    setVoiceStatus("這台裝置目前不支援站內錄音，請改用新版 Chrome 或 Safari。");
    voiceButton.textContent = "停止練習";
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
    addBubble("user", data.transcript || "已收到你的語音回答。");
    addBubble("coach", data.message);
    if (data.done) {
      await loadConfig();
      resetPracticeHome(data.message);
      return;
    }
    awaitingCoachReply = false;
    deliverCoachMessage(data.message, { autoListen: true });
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

  try {
    const result = await api("/api/respond", { message });
    removePendingCoachBubble();
    addBubble("coach", result.message);
    if (result.done) {
      await loadConfig();
      resetPracticeHome(result.message);
      return;
    }
    awaitingCoachReply = false;
    deliverCoachMessage(result.message, { autoListen: true });
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
    addBubble("coach", `目前無法結束這輾練習。\n原因：${err.message}`);
    setVoiceStatus("目前無法結束練習，請稍後再試一次。");
    showError(appPanel, err.message);
  } finally {
    setResponsePending(false);
  }
}

voiceButton.addEventListener("click", stopPracticeSession);

speakButton.addEventListener("click", async () => {
  await goToLoginPage();
});

addSectionButton.addEventListener("click", () => addSectionEditor());
addEmployeeButton.addEventListener("click", () => addEmployeeEditor());
roleSelect.addEventListener("change", toggleLoginFields);

saveAuthButton.addEventListener("click", async () => {
  try {
    const result = await api("/api/admin/auth", collectAuthData());
    currentAuth = result.auth;
    renderAuthForm();
    renderEmployeesForm();
    renderReportFilters();
    showAdminFeedback("帳號與員工設定已儲存。");
  } catch (err) {
    showAdminFeedback(err.message, "error");
  }
});

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

saveAdminButton.addEventListener("click", async () => {
  try {
    const authResult = await api("/api/admin/auth", collectAuthData());
    currentAuth = authResult.auth;

    const contentResult = await api("/api/admin/content", collectAdminData());
    currentConfig = contentResult.content;
    renderAdmin();
    renderReportFilters();
    await loadConfig();
    await loadReports();
    setMode("admin");
    showAdminFeedback("後台全部設定已儲存，包含員工帳號、密碼、題庫與規則。");
  } catch (err) {
    showAdminFeedback(err.message, "error");
  }
});

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
refreshActionState();
