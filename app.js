const APP_CONFIG = window.APP_CONFIG || {};
const DEFAULT_API_URL = APP_CONFIG.apiUrl;
const DEFAULT_STAFF_NAME = APP_CONFIG.staffName;
const API_PLACEHOLDER = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
const STORAGE_KEY = "smesScanner.apiUrl";
const STAFF_STORAGE_KEY = "smesScanner.staffName";
const REQUEST_TIMEOUT_MS = 15000;
const STATUS_CACHE_MS = 60000;
const SCAN_COOLDOWN_MS = 8000;
const AUTO_RESUME_MS = 1800;
const BUNDLED_API_URL = normalizeApiUrl(DEFAULT_API_URL);
const BUNDLED_STAFF_NAME = normalizeStaffName(DEFAULT_STAFF_NAME);
const DEFAULT_TITLE = document.title;

let apiUrl = resolveInitialApiUrl();
let staffName = resolveInitialStaffName();
let html5QrCode = null;
let scannerRunning = false;
let liveScanSession = false;
let resumeTimer = null;
let lastHandledCode = "";
let lastHandledAt = 0;
let backendReady = false;
let lastStatusCheckAt = 0;
let audioContext = null;
let lastStatusPayload = null;
let pendingCheckIn = null;

const refs = {
  card: document.getElementById("scannerCard"),
  systemMeta: document.getElementById("systemMeta"),
  systemNamePill: document.getElementById("systemNamePill"),
  eventDetailsPill: document.getElementById("eventDetailsPill"),
  staffNamePill: document.getElementById("staffNamePill"),
  startBtn: document.getElementById("startBtn"),
  stopBtn: document.getElementById("stopBtn"),
  launchPanel: document.getElementById("launchPanel"),
  launchLinkInput: document.getElementById("launchLinkInput"),
  copyLaunchLinkBtn: document.getElementById("copyLaunchLinkBtn"),
  shareLaunchLinkBtn: document.getElementById("shareLaunchLinkBtn"),
  shareMessage: document.getElementById("shareMessage"),
  toggleManualBtn: document.getElementById("toggleManualBtn"),
  manualForm: document.getElementById("manualForm"),
  manualCode: document.getElementById("manualCode"),
  manualSubmitBtn: document.getElementById("manualSubmitBtn"),
  qrImageInput: document.getElementById("qrImageInput"),
  imageTrigger: document.querySelector('[for="qrImageInput"]'),
  message: document.getElementById("message"),
  result: document.getElementById("result"),
  confirmationPanel: document.getElementById("confirmationPanel"),
  confirmationHint: document.getElementById("confirmationHint"),
  confirmCheckInBtn: document.getElementById("confirmCheckInBtn"),
  cancelCheckInBtn: document.getElementById("cancelCheckInBtn"),
  statusBadge: document.getElementById("statusBadge"),
  statusHint: document.getElementById("statusHint"),
  systemStatus: document.getElementById("systemStatus"),
  systemToggleBtn: document.getElementById("systemToggleBtn"),
  systemPanel: document.getElementById("systemPanel"),
  configForm: document.getElementById("configForm"),
  apiUrlInput: document.getElementById("apiUrlInput"),
  staffNameInput: document.getElementById("staffNameInput"),
  testApiBtn: document.getElementById("testApiBtn"),
  clearApiBtn: document.getElementById("clearApiBtn"),
  configMessage: document.getElementById("configMessage")
};

function safeStorageGet(key) {
  try {
    return window.localStorage.getItem(key);
  } catch (_) {
    return "";
  }
}

function safeStorageSet(key, value) {
  try {
    window.localStorage.setItem(key, value);
  } catch (_) {}
}

function safeStorageRemove(key) {
  try {
    window.localStorage.removeItem(key);
  } catch (_) {}
}

function normalizeApiUrl(value) {
  const normalized = String(value || "").trim();

  if (!normalized || normalized.includes(API_PLACEHOLDER)) {
    return "";
  }

  return normalized;
}

function normalizeStaffName(value) {
  return String(value || "").trim();
}

function resolveInitialApiUrl() {
  const params = new URLSearchParams(window.location.search);
  const queryValue = normalizeApiUrl(params.get("apiUrl"));

  if (queryValue) {
    safeStorageSet(STORAGE_KEY, queryValue);
    return queryValue;
  }

  return normalizeApiUrl(safeStorageGet(STORAGE_KEY) || BUNDLED_API_URL);
}

function resolveInitialStaffName() {
  const params = new URLSearchParams(window.location.search);
  const queryValue = normalizeStaffName(params.get("staffName"));

  if (queryValue) {
    safeStorageSet(STAFF_STORAGE_KEY, queryValue);
    return queryValue;
  }

  return normalizeStaffName(safeStorageGet(STAFF_STORAGE_KEY) || BUNDLED_STAFF_NAME);
}

function hasUsableApiUrl() {
  return Boolean(apiUrl);
}

function isBackendStatusFresh() {
  return backendReady && Date.now() - lastStatusCheckAt < STATUS_CACHE_MS;
}

function setSystemPanelOpen(open) {
  refs.systemPanel.hidden = !open;
  refs.systemToggleBtn.setAttribute("aria-expanded", String(open));
}

function setSystemStatus(text, tone = "idle") {
  refs.systemStatus.className = `status-system ${tone}`;
  refs.systemStatus.textContent = text;
}

function setConfigMessage(text = "", type = "success") {
  refs.configMessage.className = text ? `${type}box` : "";
  refs.configMessage.textContent = text;
}

function setShareMessage(text = "", type = "success") {
  refs.shareMessage.className = text ? `${type}box` : "";
  refs.shareMessage.textContent = text;
}

function updateStaffNamePill() {
  if (!staffName) {
    refs.staffNamePill.hidden = true;
    refs.staffNamePill.textContent = "";
    return;
  }

  refs.staffNamePill.hidden = false;
  refs.staffNamePill.textContent = `ผู้สแกน: ${staffName}`;
}

function setSystemMeta(payload) {
  lastStatusPayload = payload || null;

  if (!payload) {
    refs.systemMeta.hidden = true;
    refs.systemNamePill.textContent = "ยังไม่เชื่อมต่อระบบ";
    refs.eventDetailsPill.textContent = "รอข้อมูลอีเวนต์";
    document.title = DEFAULT_TITLE;
    updateStaffNamePill();
    return;
  }

  const systemName = payload.eventName || payload.apiName || "ระบบเช็กอิน";
  const eventDetails = [payload.eventDateText, payload.eventTimeText].filter(Boolean).join(" | ");

  refs.systemMeta.hidden = false;
  refs.systemNamePill.textContent = systemName;
  refs.eventDetailsPill.textContent = eventDetails || payload.message || "พร้อมใช้งาน";
  document.title = `${systemName} | QR Check-in`;
  updateStaffNamePill();
}

function buildLaunchLink() {
  if (!hasUsableApiUrl()) {
    return "";
  }

  const baseUrl = (lastStatusPayload && lastStatusPayload.scannerBaseUrl) || window.location.href;
  const launchUrl = new URL(baseUrl, window.location.href);
  launchUrl.search = "";
  launchUrl.hash = "";
  launchUrl.searchParams.set("apiUrl", apiUrl);

  if (staffName) {
    launchUrl.searchParams.set("staffName", staffName);
  }

  return launchUrl.toString();
}

function updateLaunchPanel() {
  const launchLink = buildLaunchLink();
  const hasLink = Boolean(launchLink);

  refs.launchPanel.hidden = !hasLink;
  refs.launchLinkInput.value = launchLink;
  refs.copyLaunchLinkBtn.disabled = !hasLink;
  refs.shareLaunchLinkBtn.disabled = !hasLink;

  if (!hasLink) {
    setShareMessage("");
  }

  updateStaffNamePill();
}

function clearPendingCheckIn() {
  pendingCheckIn = null;
  refs.confirmationPanel.hidden = true;
  refs.confirmationHint.textContent = "ยืนยันชื่อผู้เข้าร่วมบนหน้าจอก่อนบันทึกเช็กอิน";
  refs.confirmCheckInBtn.disabled = false;
  refs.cancelCheckInBtn.disabled = false;
}

function showConfirmationPanel(text) {
  refs.confirmationPanel.hidden = false;
  refs.confirmationHint.textContent = text || "ยืนยันชื่อผู้เข้าร่วมบนหน้าจอก่อนบันทึกเช็กอิน";
  refs.confirmCheckInBtn.disabled = false;
  refs.cancelCheckInBtn.disabled = false;
}

function focusConfig(message) {
  setSystemPanelOpen(true);
  setSystemStatus("ยังไม่ได้ตั้งค่าระบบ", "error");
  setConfigMessage(message, "error");
  refs.apiUrlInput.focus();
}

function applyApiUrl(value, options = {}) {
  const { persist = true } = options;
  apiUrl = normalizeApiUrl(value);
  backendReady = false;
  lastStatusCheckAt = 0;
  refs.apiUrlInput.value = apiUrl;

  if (apiUrl && persist) {
    safeStorageSet(STORAGE_KEY, apiUrl);
  } else if (!apiUrl) {
    safeStorageRemove(STORAGE_KEY);
  }

  if (!apiUrl) {
    setSystemStatus("ยังไม่ได้ตั้งค่าระบบ", "error");
    setUiState("error", "ใส่ URL ของ Apps Script ก่อนใช้งาน");
  } else {
    setSystemStatus("ยังไม่ทดสอบระบบ", "warning");
    setUiState("idle", "พร้อมทดสอบระบบหรือเริ่มสแกน");
  }

  setSystemMeta(null);
  setShareMessage("");
  clearPendingCheckIn();
  updateLaunchPanel();
}

function applyStaffName(value, options = {}) {
  const { persist = true } = options;
  staffName = normalizeStaffName(value);
  refs.staffNameInput.value = staffName;

  if (staffName && persist) {
    safeStorageSet(STAFF_STORAGE_KEY, staffName);
  } else if (!staffName) {
    safeStorageRemove(STAFF_STORAGE_KEY);
  }

  setShareMessage("");
  updateLaunchPanel();
}

function setUiState(state, hintOverride = "") {
  const states = {
    idle: {
      label: "พร้อมเริ่ม",
      hint: "กดเริ่มสแกนเพื่อเปิดกล้อง"
    },
    scanning: {
      label: "กำลังสแกน",
      hint: "ยกกล้องไปที่ QR ให้อยู่ในกรอบ"
    },
    processing: {
      label: "กำลังตรวจสอบ",
      hint: "รอสักครู่ ระบบกำลังตรวจสอบข้อมูล"
    },
    success: {
      label: "สำเร็จ",
      hint: "พร้อมสำหรับรายการถัดไป"
    },
    warning: {
      label: "ต้องตรวจสอบ",
      hint: "ตรวจสอบผลลัพธ์ก่อนดำเนินการต่อ"
    },
    error: {
      label: "มีปัญหา",
      hint: "ตรวจสอบข้อความและลองอีกครั้ง"
    }
  };

  const current = states[state] || states.idle;
  const hint = hintOverride || current.hint;
  const missingConfig = !hasUsableApiUrl();
  const disabledForProcessing = state === "processing";
  const disableFallback = disabledForProcessing || missingConfig;
  const canStart = !missingConfig && state !== "scanning" && state !== "processing";

  refs.card.dataset.state = state;
  refs.statusBadge.className = `status-badge ${state}`;
  refs.statusBadge.textContent = current.label;
  refs.statusHint.textContent = hint;

  refs.startBtn.disabled = !canStart;
  refs.stopBtn.disabled = state !== "scanning";
  refs.toggleManualBtn.disabled = disableFallback;
  refs.manualCode.disabled = disableFallback;
  refs.manualSubmitBtn.disabled = disableFallback;
  refs.qrImageInput.disabled = disableFallback;
  refs.confirmCheckInBtn.disabled = disabledForProcessing || !pendingCheckIn;
  refs.cancelCheckInBtn.disabled = disabledForProcessing || !pendingCheckIn;

  if (refs.imageTrigger) {
    refs.imageTrigger.classList.toggle("is-disabled", disableFallback);
  }
}

function setMessage(text, type = "success") {
  refs.message.className = `${type}box`;
  refs.message.textContent = text;
}

function clearMessage() {
  refs.message.className = "";
  refs.message.textContent = "";
}

function setResult(data, tone = "success") {
  if (!data) {
    refs.result.innerHTML = "";
    return;
  }

  const items = [];
  const outcome = data.message || (tone === "success" ? "เช็กอินสำเร็จ" : "ผลการตรวจสอบ");

  if (data.fullName) {
    items.push(`<div class="result-item"><strong>ชื่อ:</strong> ${escapeHtml(data.fullName)}</div>`);
  }

  if (data.organization) {
    items.push(`<div class="result-item"><strong>องค์กร:</strong> ${escapeHtml(data.organization)}</div>`);
  }

  if (data.registrationId) {
    items.push(`<div class="result-item"><strong>Registration ID:</strong> ${escapeHtml(data.registrationId)}</div>`);
  }

  if (data.status) {
    items.push(`<div class="result-item"><strong>สถานะ:</strong> ${escapeHtml(data.status)}</div>`);
  }

  if (data.checkedInAt) {
    items.push(`<div class="result-item"><strong>เวลา:</strong> ${escapeHtml(formatDateTime(data.checkedInAt))}</div>`);
  }

  refs.result.innerHTML = `
    <div class="result ${tone}">
      <div class="result-title">${escapeHtml(outcome)}</div>
      ${items.join("")}
    </div>
  `;
}

function formatDateTime(value) {
  const date = new Date(value);

  if (Number.isNaN(date.getTime())) {
    return String(value || "-");
  }

  return date.toLocaleString("th-TH");
}

async function fetchJson(url, options = {}) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

  try {
    const response = await fetch(url, {
      redirect: "follow",
      cache: "no-store",
      ...options,
      signal: controller.signal
    });

    const text = await response.text();

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    try {
      return JSON.parse(text);
    } catch (_) {
      throw new Error("INVALID_JSON");
    }
  } catch (err) {
    if (err && err.name === "AbortError") {
      throw new Error("REQUEST_TIMEOUT");
    }

    throw err;
  } finally {
    clearTimeout(timer);
  }
}

function buildStatusUrl() {
  const separator = apiUrl.includes("?") ? "&" : "?";
  return `${apiUrl}${separator}action=status`;
}

async function testApiConnection(options = {}) {
  const { silent = false } = options;

  if (!hasUsableApiUrl()) {
    focusConfig("กรุณาวาง Apps Script Web App URL ก่อน");
    return false;
  }

  setSystemStatus("กำลังเชื่อมต่อระบบ", "warning");

  if (!silent) {
    setConfigMessage("กำลังทดสอบการเชื่อมต่อ...", "warning");
  }

  try {
    const result = await fetchJson(buildStatusUrl(), { method: "GET" });

    if (!result || result.ok !== true) {
      throw new Error("STATUS_FAILED");
    }

    backendReady = true;
    lastStatusCheckAt = Date.now();
    setSystemStatus("เชื่อมต่อระบบแล้ว", "success");
    setSystemMeta(result);
    updateLaunchPanel();

    if (!silent) {
      setConfigMessage("เชื่อมต่อสำเร็จ พร้อมใช้งาน", "success");
    }

    if (refs.card.dataset.state === "error" || !refs.card.dataset.state) {
      setUiState("idle");
    }

    return true;
  } catch (_) {
    backendReady = false;
    lastStatusCheckAt = 0;
    setSystemStatus("เชื่อมต่อระบบไม่สำเร็จ", "error");
    setSystemMeta(null);

    if (!silent) {
      setConfigMessage(
        "เชื่อมต่อไม่สำเร็จ ตรวจสอบ URL, สิทธิ์การเข้าถึง และการ Deploy ของ Apps Script",
        "error"
      );
    }

    return false;
  }
}

async function ensureBackendReady(options = {}) {
  const { force = false, silent = true } = options;

  if (!hasUsableApiUrl()) {
    focusConfig("กรุณาเชื่อมต่อ Apps Script ก่อนใช้งาน");
    return false;
  }

  if (!force && isBackendStatusFresh()) {
    return true;
  }

  const ready = await testApiConnection({ silent });

  if (!ready) {
    setUiState("error", "ยังไม่สามารถเชื่อมต่อระบบเช็กอินได้");

    if (!silent) {
      setSystemPanelOpen(true);
    }
  }

  return ready;
}

function requireConfigured() {
  if (hasUsableApiUrl()) {
    return true;
  }

  focusConfig("กรุณาวาง Apps Script Web App URL ก่อนใช้งาน");
  setUiState("error", "ตั้งค่าระบบก่อนเริ่มใช้งาน");
  return false;
}

function clearResumeTimer() {
  if (resumeTimer) {
    clearTimeout(resumeTimer);
    resumeTimer = null;
  }
}

function hideManualForm() {
  refs.manualForm.hidden = true;
  refs.toggleManualBtn.setAttribute("aria-expanded", "false");
}

async function toggleManualForm() {
  if (!requireConfigured()) {
    return;
  }

  clearResumeTimer();
  clearPendingCheckIn();
  liveScanSession = false;

  if (scannerRunning) {
    await stopScanner();
  }

  const shouldShow = refs.manualForm.hidden;
  refs.manualForm.hidden = !shouldShow;
  refs.toggleManualBtn.setAttribute("aria-expanded", String(shouldShow));

  if (shouldShow) {
    refs.manualCode.focus();
  }
}

function getScannerConfig() {
  const readerWidth = document.getElementById("reader").clientWidth || 320;
  const edge = Math.max(220, Math.min(300, Math.floor(readerWidth * 0.72)));
  const config = {
    fps: 12,
    qrbox: { width: edge, height: edge },
    aspectRatio: 1
  };

  if (window.Html5QrcodeSupportedFormats) {
    config.formatsToSupport = [Html5QrcodeSupportedFormats.QR_CODE];
  }

  return config;
}

function getQrScanner() {
  if (!html5QrCode) {
    html5QrCode = new Html5Qrcode("reader");
  }

  return html5QrCode;
}

function pickPreferredCamera(cameras) {
  const keywords = ["back", "rear", "environment"];

  return cameras.find(camera => {
    const label = String(camera.label || "").toLowerCase();
    return keywords.some(keyword => label.includes(keyword));
  }) || cameras[0];
}

function getAudioContext() {
  const AudioContextClass = window.AudioContext || window.webkitAudioContext;

  if (!AudioContextClass) {
    return null;
  }

  if (!audioContext) {
    audioContext = new AudioContextClass();
  }

  if (audioContext.state === "suspended") {
    audioContext.resume().catch(() => {});
  }

  return audioContext;
}

function playTone(frequency, durationMs) {
  const context = getAudioContext();

  if (!context) {
    return;
  }

  const oscillator = context.createOscillator();
  const gain = context.createGain();
  const now = context.currentTime;

  oscillator.type = "sine";
  oscillator.frequency.value = frequency;

  gain.gain.setValueAtTime(0.0001, now);
  gain.gain.exponentialRampToValueAtTime(0.08, now + 0.01);
  gain.gain.exponentialRampToValueAtTime(0.0001, now + durationMs / 1000);

  oscillator.connect(gain);
  gain.connect(context.destination);

  oscillator.start(now);
  oscillator.stop(now + durationMs / 1000 + 0.02);
}

function triggerFeedback(tone) {
  const vibrationPatterns = {
    success: [45],
    warning: [25, 35, 25],
    error: [70, 40, 70]
  };

  if (navigator.vibrate && vibrationPatterns[tone]) {
    navigator.vibrate(vibrationPatterns[tone]);
  }

  if (tone === "success") {
    playTone(880, 90);
  } else if (tone === "warning") {
    playTone(660, 130);
  } else if (tone === "error") {
    playTone(240, 160);
  }
}

async function requestCameraAndStart() {
  if (!requireConfigured()) {
    return;
  }

  const backendOk = await ensureBackendReady({ silent: false });
  if (!backendOk) {
    setMessage("ยังไม่สามารถเชื่อมต่อระบบเช็กอินได้", "error");
    return;
  }

  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    setUiState("error", "อุปกรณ์นี้ไม่รองรับการเปิดกล้องผ่านเบราว์เซอร์");
    setMessage("เบราว์เซอร์นี้ไม่รองรับการใช้งานกล้อง", "error");
    return;
  }

  clearResumeTimer();
  clearMessage();
  setResult(null);
  clearPendingCheckIn();
  hideManualForm();
  liveScanSession = true;
  lastHandledCode = "";
  lastHandledAt = 0;
  setUiState("processing", "ถ้ามี popup ขอสิทธิ์กล้อง ให้กดอนุญาต แล้วภาพจะขึ้นในกรอบนี้");
  setMessage("หน้าต่างกล้องที่ลอยขึ้นมาจาก Chrome เป็นการขอสิทธิ์ ไม่ใช่ตัวสแกนของหน้าเว็บ", "warning");

  try {
    await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: "environment" } },
      audio: false
    }).then(stream => {
      stream.getTracks().forEach(track => track.stop());
    });

    await startScanner();
  } catch (err) {
    liveScanSession = false;
    setUiState("error", "อนุญาตการใช้กล้องแล้วลองอีกครั้ง");
    setMessage("เปิดกล้องไม่สำเร็จ: " + (err.message || err), "error");
  }
}

async function startScanner() {
  if (scannerRunning) {
    return;
  }

  clearResumeTimer();
  clearMessage();
  setResult(null);

  try {
    const scanner = getQrScanner();
    setUiState("scanning");

    await scanner.start(
      { facingMode: "environment" },
      getScannerConfig(),
      onScanSuccess
    );

    scannerRunning = true;
  } catch (_) {
    try {
      const cameras = await Html5Qrcode.getCameras();
      if (!cameras.length) {
        liveScanSession = false;
        setUiState("error", "ไม่พบกล้องบนอุปกรณ์นี้");
        setMessage("ไม่พบกล้องบนอุปกรณ์นี้", "error");
        return;
      }

      const scanner = getQrScanner();
      const preferredCamera = pickPreferredCamera(cameras);

      await scanner.start(
        preferredCamera.id,
        getScannerConfig(),
        onScanSuccess
      );

      scannerRunning = true;
      setUiState("scanning");
    } catch (err) {
      liveScanSession = false;
      setUiState("error", "เปิดกล้องไม่สำเร็จ ลองใช้รูปภาพหรือกรอกรหัสแทน");
      setMessage("เปิดกล้องไม่สำเร็จ: " + (err.message || err), "error");
    }
  }
}

async function stopScanner(options = {}) {
  const {
    resetSession = true,
    nextState = "idle",
    hint = "กดเริ่มสแกนเมื่อพร้อม"
  } = options;

  clearResumeTimer();

  if (html5QrCode && scannerRunning) {
    try {
      await html5QrCode.stop();
    } catch (_) {}

    try {
      await html5QrCode.clear();
    } catch (_) {}
  }

  scannerRunning = false;

  if (resetSession) {
    liveScanSession = false;
  }

  document.getElementById("reader").innerHTML = "";
  html5QrCode = null;

  if (nextState) {
    setUiState(nextState, hint);
  }
}

function getResultTone(result) {
  if (result.success) {
    return "success";
  }

  if (result.registrationId && result.checkedInAt) {
    return "warning";
  }

  return "error";
}

function getResultMessage(result, tone) {
  if (tone === "success") {
    return `${result.message}: ${result.fullName} (${result.registrationId})`;
  }

  if (tone === "warning") {
    return `${result.fullName} เช็กอินแล้วเมื่อ ${formatDateTime(result.checkedInAt)}`;
  }

  return result.message || "เช็กอินไม่สำเร็จ";
}

function queueAutoResume(tone) {
  if (!liveScanSession || tone === "error") {
    return;
  }

  clearResumeTimer();
  setUiState(
    tone === "success" ? "success" : "warning",
    "พร้อมสแกนคนถัดไปในอีกครู่"
  );

  resumeTimer = setTimeout(() => {
    startScanner().catch(err => {
      liveScanSession = false;
      setUiState("error", "กลับเข้าสู่โหมดสแกนไม่สำเร็จ");
      setMessage("เริ่มสแกนใหม่ไม่สำเร็จ: " + (err.message || err), "error");
    });
  }, AUTO_RESUME_MS);
}

async function onScanSuccess(decodedText) {
  const code = normalizeCode(decodedText);
  if (!code || refs.card.dataset.state === "processing") {
    return;
  }

  const now = Date.now();
  if (lastHandledCode === code && now - lastHandledAt < SCAN_COOLDOWN_MS) {
    return;
  }

  lastHandledCode = code;
  lastHandledAt = now;

  setUiState("processing");
  await stopScanner({ resetSession: false, nextState: null });
  await prepareCheckIn(code, { source: "live" });
}

function normalizeCode(value) {
  return String(value || "").trim();
}

async function postJson(payload) {
  return fetchJson(apiUrl, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload)
  });
}

async function lookupRegistration(registrationId) {
  return postJson({
    action: "lookup",
    registrationId
  });
}

function isCheckedInStatus(result) {
  const status = String((result && result.status) || "").trim().toLowerCase();
  return status === "checked-in" || status === "checked in";
}

function buildLookupOutcome(result) {
  if (!result || result.success !== true) {
    return {
      tone: "error",
      message: (result && result.message) || "ไม่พบข้อมูลผู้ลงทะเบียน",
      result: result || { success: false, message: "ไม่พบข้อมูลผู้ลงทะเบียน" }
    };
  }

  if (isCheckedInStatus(result)) {
    return {
      tone: "warning",
      message: result.checkedInAt
        ? `${result.fullName} เช็กอินแล้วเมื่อ ${formatDateTime(result.checkedInAt)}`
        : "ผู้เข้าร่วมท่านนี้เช็กอินแล้ว",
      result: {
        ...result,
        message: "ผู้เข้าร่วมท่านนี้เช็กอินแล้ว"
      }
    };
  }

  return {
    tone: "success",
    message: `พบข้อมูล ${result.fullName || result.registrationId} กรุณาตรวจสอบชื่อก่อนกดยืนยันเช็กอิน`,
    result: {
      ...result,
      message: "พบข้อมูลผู้ลงทะเบียน"
    }
  };
}

function getConfirmationHint(source) {
  if (source === "live") {
    return "ตรวจสอบชื่อบนหน้าจอ แล้วกดยืนยันเช็กอิน หรือยกเลิกเพื่อกลับไปสแกนต่อ";
  }

  return "ตรวจสอบชื่อบนหน้าจอ แล้วกดยืนยันเช็กอินเมื่อพร้อม";
}

async function prepareCheckIn(registrationId, options = {}) {
  const { source = "manual" } = options;
  const code = normalizeCode(registrationId);

  if (!code) {
    setUiState("error", "กรุณากรอก Registration ID");
    setMessage("กรุณากรอก Registration ID", "error");
    return null;
  }

  const ready = await ensureBackendReady({ silent: true });
  if (!ready) {
    liveScanSession = false;
    clearPendingCheckIn();
    setUiState("error", "ยังไม่สามารถเชื่อมต่อระบบเช็กอินได้");
    setMessage("เชื่อมต่อระบบไม่สำเร็จ กรุณาตรวจสอบการตั้งค่าระบบ", "error");
    setSystemPanelOpen(true);
    return null;
  }

  clearPendingCheckIn();
  clearMessage();
  setResult(null);
  setUiState("processing", "กำลังค้นหาข้อมูลผู้ลงทะเบียน");

  try {
    const lookupResult = await lookupRegistration(code);
    const outcome = buildLookupOutcome(lookupResult);

    backendReady = true;
    lastStatusCheckAt = Date.now();
    setSystemStatus("เชื่อมต่อระบบแล้ว", "success");
    setMessage(outcome.message, outcome.tone);
    setResult(outcome.result, outcome.tone);

    if (outcome.tone !== "success") {
      liveScanSession = false;
      setUiState(
        outcome.tone,
        outcome.tone === "warning" ? "รายการนี้ไม่สามารถเช็กอินซ้ำได้" : "ตรวจสอบข้อมูลแล้วลองใหม่อีกครั้ง"
      );
      triggerFeedback(outcome.tone);
      return outcome;
    }

    pendingCheckIn = {
      registrationId: code,
      source,
      lookupResult
    };

    showConfirmationPanel(getConfirmationHint(source));
    setUiState("warning", "ตรวจสอบชื่อบนหน้าจอ แล้วกดยืนยันเช็กอิน");
    return outcome;
  } catch (_) {
    backendReady = false;
    lastStatusCheckAt = 0;
    liveScanSession = false;
    clearPendingCheckIn();
    setSystemStatus("เชื่อมต่อระบบมีปัญหา", "error");
    setUiState("error", "ยังไม่สามารถติดต่อระบบเช็กอินได้");
    setMessage("เชื่อมต่อระบบไม่สำเร็จ กรุณาตรวจสอบอินเทอร์เน็ตหรือ URL ของระบบ", "error");
    triggerFeedback("error");
    return null;
  }
}

async function checkIn(registrationId, options = {}) {
  const { source = "manual" } = options;
  const code = normalizeCode(registrationId);

  if (!code) {
    setUiState("error", "กรุณากรอก Registration ID");
    setMessage("กรุณากรอก Registration ID", "error");
    return null;
  }

  const ready = await ensureBackendReady({ silent: true });
  if (!ready) {
    liveScanSession = false;
    setUiState("error", "ยังไม่สามารถเชื่อมต่อระบบเช็กอินได้");
    setMessage("เชื่อมต่อระบบไม่สำเร็จ กรุณาตรวจสอบการตั้งค่าระบบ", "error");
    setSystemPanelOpen(true);
    return null;
  }

  clearMessage();
  setUiState("processing");
  refs.confirmationHint.textContent = "กำลังบันทึกเช็กอิน...";

  try {
    const payload = {
      action: "checkin",
      registrationId: code
    };

    if (staffName) {
      payload.staffName = staffName;
    }

    const result = await postJson(payload);

    const tone = getResultTone(result);
    backendReady = true;
    lastStatusCheckAt = Date.now();
    setSystemStatus("เชื่อมต่อระบบแล้ว", "success");
    clearPendingCheckIn();
    setMessage(getResultMessage(result, tone), tone);
    setResult(result, tone);
    triggerFeedback(tone);

    if (source === "live") {
      queueAutoResume(tone);
      if (tone === "error") {
        liveScanSession = false;
        setUiState("error", "ตรวจสอบผลลัพธ์แล้วเริ่มสแกนใหม่ได้");
      }
    } else {
      liveScanSession = false;
      setUiState(
        tone,
        tone === "success" ? "ดำเนินการรายการถัดไปได้ทันที" : "ตรวจสอบผลลัพธ์บนหน้าจอ"
      );
    }

    return { result, tone };
  } catch (_) {
    backendReady = false;
    lastStatusCheckAt = 0;
    setSystemStatus("เชื่อมต่อระบบมีปัญหา", "error");
    setUiState("error", "ยังไม่สามารถติดต่อระบบเช็กอินได้");
    setMessage("เชื่อมต่อระบบไม่สำเร็จ กรุณาตรวจสอบอินเทอร์เน็ตหรือ URL ของระบบ", "error");
    showConfirmationPanel("ยังบันทึกเช็กอินไม่สำเร็จ กดยืนยันอีกครั้งหรือตรวจสอบระบบ");
    triggerFeedback("error");
    return null;
  }
}

async function handleConfirmCheckIn() {
  if (!pendingCheckIn) {
    return;
  }

  await checkIn(pendingCheckIn.registrationId, { source: pendingCheckIn.source });
}

async function handleCancelCheckIn() {
  const currentPending = pendingCheckIn;
  clearPendingCheckIn();
  clearMessage();
  setResult(null);

  if (currentPending && currentPending.source === "live" && liveScanSession) {
    setUiState("processing", "กำลังกลับเข้าสู่โหมดสแกน");

    try {
      await startScanner();
    } catch (err) {
      liveScanSession = false;
      setUiState("error", "กลับเข้าสู่โหมดสแกนไม่สำเร็จ");
      setMessage("เริ่มสแกนใหม่ไม่สำเร็จ: " + (err.message || err), "error");
    }

    return;
  }

  liveScanSession = false;
  setUiState("idle", "ยกเลิกรายการแล้ว พร้อมเริ่มใหม่");
}

async function copyText(value) {
  if (navigator.clipboard && window.isSecureContext) {
    await navigator.clipboard.writeText(value);
    return;
  }

  refs.launchLinkInput.focus();
  refs.launchLinkInput.select();
  refs.launchLinkInput.setSelectionRange(0, refs.launchLinkInput.value.length);

  if (!document.execCommand("copy")) {
    throw new Error("COPY_FAILED");
  }
}

async function handleCopyLaunchLink() {
  const launchLink = buildLaunchLink();

  if (!launchLink) {
    setShareMessage("ตั้งค่า Apps Script ให้เรียบร้อยก่อน ระบบจึงจะสร้างลิงก์พร้อมใช้งานได้", "warning");
    return;
  }

  try {
    await copyText(launchLink);
    setShareMessage("คัดลอกลิงก์พร้อมใช้งานแล้ว", "success");
  } catch (_) {
    setShareMessage("คัดลอกลิงก์ไม่สำเร็จ ลองกดค้างที่ช่องลิงก์แล้วคัดลอกเอง", "error");
  }
}

async function handleShareLaunchLink() {
  const launchLink = buildLaunchLink();

  if (!launchLink) {
    setShareMessage("ตั้งค่า Apps Script ให้เรียบร้อยก่อน ระบบจึงจะสร้างลิงก์พร้อมใช้งานได้", "warning");
    return;
  }

  if (navigator.share) {
    try {
      await navigator.share({
        title: document.title,
        text: "ลิงก์ระบบเช็กอินพร้อมใช้งาน",
        url: launchLink
      });
      setShareMessage("เปิดหน้าต่างแชร์ลิงก์แล้ว", "success");
      return;
    } catch (err) {
      if (err && err.name === "AbortError") {
        return;
      }
    }
  }

  await handleCopyLaunchLink();
}

async function handleManualSubmit(event) {
  event.preventDefault();

  if (!requireConfigured()) {
    return;
  }

  if (scannerRunning) {
    await stopScanner();
  }

  const outcome = await prepareCheckIn(refs.manualCode.value, { source: "manual" });
  if (outcome && outcome.tone !== "error") {
    refs.manualCode.value = "";
  }
}

async function handleImageSelection(event) {
  const file = event.target.files[0];
  refs.qrImageInput.value = "";

  if (!file) {
    return;
  }

  if (!requireConfigured()) {
    return;
  }

  if (scannerRunning) {
    await stopScanner();
  }

  hideManualForm();
  clearResumeTimer();
  clearMessage();
  setResult(null);
  clearPendingCheckIn();
  liveScanSession = false;
  setUiState("processing", "กำลังอ่าน QR จากรูปภาพ");

  try {
    const scanner = getQrScanner();
    const decodedText = await scanner.scanFile(file, false);
    await prepareCheckIn(decodedText, { source: "image" });
  } catch (_) {
    setUiState("error", "ลองเลือกรูปใหม่หรือกรอกรหัสแทน");
    setMessage("อ่าน QR จากรูปไม่สำเร็จ กรุณาเลือกรูปที่เห็น QR ชัดเจน", "error");
    triggerFeedback("error");
  } finally {
    document.getElementById("reader").innerHTML = "";
    html5QrCode = null;
  }
}

async function handleConfigSubmit(event) {
  event.preventDefault();

  applyApiUrl(refs.apiUrlInput.value);
  applyStaffName(refs.staffNameInput.value);

  if (!hasUsableApiUrl()) {
    focusConfig("กรุณาวาง Apps Script Web App URL ก่อน");
    return;
  }

  const ok = await testApiConnection({ silent: false });
  if (ok) {
    clearMessage();
    setUiState("idle");
    setSystemPanelOpen(false);
  }
}

async function handleRetestClick() {
  applyApiUrl(refs.apiUrlInput.value);
  applyStaffName(refs.staffNameInput.value);

  if (!hasUsableApiUrl()) {
    focusConfig("กรุณาวาง Apps Script Web App URL ก่อน");
    return;
  }

  await testApiConnection({ silent: false });
}

async function handleClearApiClick() {
  if (scannerRunning) {
    await stopScanner();
  }

  applyApiUrl("", { persist: false });
  applyStaffName("", { persist: false });
  clearMessage();
  setResult(null);
  liveScanSession = false;
  setConfigMessage("ล้างค่าแล้ว วาง URL ใหม่เพื่อเริ่มใช้งาน", "warning");
  setSystemPanelOpen(true);
}

function toggleSystemPanel() {
  setSystemPanelOpen(refs.systemPanel.hidden);

  if (!refs.systemPanel.hidden) {
    refs.apiUrlInput.focus();
  }
}

function escapeHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

async function initializeApp() {
  refs.apiUrlInput.value = apiUrl;
  refs.staffNameInput.value = staffName;
  updateLaunchPanel();

  if (hasUsableApiUrl()) {
    setSystemStatus("กำลังตรวจสอบระบบ", "warning");
    setSystemPanelOpen(false);
    setUiState("idle");
    let ok = await testApiConnection({ silent: true });

    if (!ok && BUNDLED_API_URL && apiUrl !== BUNDLED_API_URL) {
      applyApiUrl(BUNDLED_API_URL);
      setConfigMessage("ตรวจพบค่า URL เดิมที่ใช้ไม่ได้ ระบบจึงสลับกลับไปใช้ค่า default ล่าสุด", "warning");
      ok = await testApiConnection({ silent: true });
    }

    if (!ok) {
      setUiState("error", "ยังไม่สามารถเชื่อมต่อระบบเช็กอินได้");
      setConfigMessage(
        "เชื่อมต่อไม่สำเร็จ ตรวจสอบ URL, สิทธิ์การเข้าถึง และการ Deploy ของ Apps Script",
        "error"
      );
      setSystemPanelOpen(true);
    }
  } else {
    setSystemMeta(null);
    setSystemPanelOpen(true);
    setUiState("error", "ใส่ URL ของ Apps Script ก่อนใช้งาน");
    setMessage("ยังไม่ได้ตั้งค่า API URL สำหรับระบบเช็กอิน", "error");
    setConfigMessage("วาง Apps Script Web App URL แล้วกดบันทึกและทดสอบ", "warning");
  }
}

refs.startBtn.addEventListener("click", requestCameraAndStart);
refs.stopBtn.addEventListener("click", () => stopScanner({
  hint: "สแกนถูกหยุดแล้ว กดเริ่มใหม่เมื่อต้องการ"
}));
refs.confirmCheckInBtn.addEventListener("click", handleConfirmCheckIn);
refs.cancelCheckInBtn.addEventListener("click", handleCancelCheckIn);
refs.copyLaunchLinkBtn.addEventListener("click", handleCopyLaunchLink);
refs.shareLaunchLinkBtn.addEventListener("click", handleShareLaunchLink);
refs.toggleManualBtn.addEventListener("click", toggleManualForm);
refs.manualForm.addEventListener("submit", handleManualSubmit);
refs.qrImageInput.addEventListener("change", handleImageSelection);
refs.systemToggleBtn.addEventListener("click", toggleSystemPanel);
refs.configForm.addEventListener("submit", handleConfigSubmit);
refs.testApiBtn.addEventListener("click", handleRetestClick);
refs.clearApiBtn.addEventListener("click", handleClearApiClick);

initializeApp();
