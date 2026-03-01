"use strict";

const CONFIG_KEY = "unetiScheduleConfigV1";
const DEFAULT_CONFIG_URL = chrome.runtime.getURL("config/schedule-default.json");
const DEFAULT_COMPANY = "TranDangKhoaTechnology";
const DEFAULT_LOGO_PATH = "assets/image/logo.png";
const SCHEDULE_TAB_HOST = "sinhvien.uneti.edu.vn";
const SCHEDULE_TAB_PATHS = new Set([
  "/lich-theo-tuan.html",
  "/lich-theo-tien-do.html"
]);
const GRADE_TAB_PATH = "/ket-qua-hoc-tap.html";
const SCHEDULE_HOST_PERMISSION = `https://${SCHEDULE_TAB_HOST}/*`;
const CORE_CONTENT_SCRIPT_FILE = "content/content-script.js";
const VENDOR_SCRIPT_FILES = [
  "vendor/exceljs.min.js",
  "vendor/html2canvas.min.js",
  "vendor/jspdf.umd.min.js"
];
const SLOT_TIME = {
  morning: "06:00-11:30",
  afternoon: "13:00-18:30"
};
const WEEKDAY_LABEL = {
  2: "Thứ 2",
  3: "Thứ 3",
  4: "Thứ 4",
  5: "Thứ 5",
  6: "Thứ 6",
  7: "Thứ 7",
  8: "Chủ nhật"
};
const HHMM_REGEX = /^([01]\d|2[0-3]):[0-5]\d$/;
const GRADE_EXPORT_FORMATS = new Set(["xlsx", "csv", "json"]);

let currentConfig = null;
let gradeExportReady = false;

const el = {
  brandLogo: document.getElementById("brandLogo"),
  brandText: document.getElementById("brandText"),
  manualForm: document.getElementById("manualForm"),
  manualEditId: document.getElementById("manualEditId"),
  manualSubject: document.getElementById("manualSubject"),
  manualClass: document.getElementById("manualClass"),
  manualRoom: document.getElementById("manualRoom"),
  manualTeacher: document.getElementById("manualTeacher"),
  manualMode: document.getElementById("manualMode"),
  manualWeekdayWrap: document.getElementById("manualWeekdayWrap"),
  manualWeekday: document.getElementById("manualWeekday"),
  manualDateWrap: document.getElementById("manualDateWrap"),
  manualDate: document.getElementById("manualDate"),
  manualSlot: document.getElementById("manualSlot"),
  manualEnabled: document.getElementById("manualEnabled"),
  btnManualReset: document.getElementById("btnManualReset"),
  btnManualImport: document.getElementById("btnManualImport"),
  btnManualExport: document.getElementById("btnManualExport"),
  manualImportFileInput: document.getElementById("manualImportFileInput"),
  manualStatus: document.getElementById("manualStatus"),
  manualTableBody: document.getElementById("manualTableBody"),
  brandingLinkInput: document.getElementById("brandingLinkInput"),
  btnBrandingSave: document.getElementById("btnBrandingSave"),
  brandingStatus: document.getElementById("brandingStatus"),
  exportFormat: document.getElementById("exportFormat"),
  exportRange: document.getElementById("exportRange"),
  exportMonthWrap: document.getElementById("exportMonthWrap"),
  exportMonth: document.getElementById("exportMonth"),
  exportMonthLayoutWrap: document.getElementById("exportMonthLayoutWrap"),
  exportMonthLayout: document.getElementById("exportMonthLayout"),
  exportScheduleType: document.getElementById("exportScheduleType"),
  btnExportNow: document.getElementById("btnExportNow"),
  btnOpenExportPanel: document.getElementById("btnOpenExportPanel"),
  btnRefreshExportContext: document.getElementById("btnRefreshExportContext"),
  exportStatus: document.getElementById("exportStatus"),
  gradeFormat: document.getElementById("gradeFormat"),
  gradeContext: document.getElementById("gradeContext"),
  gradeHint: document.getElementById("gradeHint"),
  btnGradeExportNow: document.getElementById("btnGradeExportNow"),
  btnRefreshGradeContext: document.getElementById("btnRefreshGradeContext"),
  gradeStatus: document.getElementById("gradeStatus"),
  configInput: document.getElementById("configInput"),
  btnSaveConfig: document.getElementById("btnSaveConfig"),
  btnSyncFromJson: document.getElementById("btnSyncFromJson"),
  btnLoadDefault: document.getElementById("btnLoadDefault"),
  btnImport: document.getElementById("btnImport"),
  btnExport: document.getElementById("btnExport"),
  importFileInput: document.getElementById("importFileInput"),
  configStatus: document.getElementById("configStatus")
};

const text = (v) => String(v || "").replace(/\s+/g, " ").trim();
const esc = (v) =>
  String(v || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");

function parseDDMMYYYY(value) {
  const match = text(value).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return null;

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  const date = new Date(year, month - 1, day);

  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null;
  }

  date.setHours(0, 0, 0, 0);
  return date;
}

function toDateInputValue(ddmmyyyy) {
  const date = parseDDMMYYYY(ddmmyyyy);
  if (!date) return "";

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function fromDateInputValue(isoDate) {
  const match = text(isoDate).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return "";
  return `${match[3]}/${match[2]}/${match[1]}`;
}

function setStatus(target, message, type) {
  target.textContent = message || "";
  target.className = "status";
  if (type) target.classList.add(type);
}

function isValidUrl(value) {
  const raw = text(value);
  if (!raw) return true;

  try {
    const parsed = new URL(raw);
    return parsed.protocol === "http:" || parsed.protocol === "https:";
  } catch (_error) {
    return false;
  }
}

function toggleExportRangeControls() {
  const isMonth = el.exportRange.value === "month";
  el.exportMonthWrap.style.display = isMonth ? "" : "none";
  el.exportMonthLayoutWrap.style.display = isMonth ? "" : "none";
}

function setExportBusy(busy) {
  el.btnExportNow.disabled = busy;
  el.btnOpenExportPanel.disabled = busy;
  el.btnRefreshExportContext.disabled = busy;
}

async function getActiveTab() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  return tabs && tabs.length ? tabs[0] : null;
}

function wait(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

function parseTabUrl(tab) {
  try {
    return new URL(String((tab || {}).url || ""));
  } catch (_error) {
    return null;
  }
}

function isScheduleTab(tab) {
  if (!tab || typeof tab.id !== "number") return false;
  const parsed = parseTabUrl(tab);
  if (!parsed) return false;

  return (
    parsed.protocol === "https:" &&
    parsed.hostname.toLowerCase() === SCHEDULE_TAB_HOST &&
    SCHEDULE_TAB_PATHS.has(parsed.pathname.toLowerCase())
  );
}

function isGradeTab(tab) {
  if (!tab || typeof tab.id !== "number") return false;
  const parsed = parseTabUrl(tab);
  if (!parsed) return false;

  return (
    parsed.protocol === "https:" &&
    parsed.hostname.toLowerCase() === SCHEDULE_TAB_HOST &&
    parsed.pathname.toLowerCase() === GRADE_TAB_PATH
  );
}

function isUnetiHostTab(tab) {
  if (!tab || typeof tab.id !== "number") return false;
  const parsed = parseTabUrl(tab);
  if (!parsed) return false;

  return parsed.protocol === "https:" && parsed.hostname.toLowerCase() === SCHEDULE_TAB_HOST;
}

function isReceivingEndMissing(error) {
  const message = String(error && error.message ? error.message : error).toLowerCase();
  return message.includes("receiving end does not exist");
}

function getErrorMessage(error, fallback) {
  if (error && typeof error.message === "string" && error.message.trim()) {
    return error.message.trim();
  }
  const raw = String(error || "").trim();
  return raw || fallback;
}

async function sendMessageRaw(tabId, payload) {
  return new Promise((resolve, reject) => {
    try {
      chrome.tabs.sendMessage(tabId, payload, (response) => {
        const runtimeError = chrome.runtime.lastError;
        if (runtimeError) {
          reject(new Error(runtimeError.message || "Không gửi được message tới tab."));
          return;
        }
        resolve(response);
      });
    } catch (error) {
      reject(error);
    }
  });
}

async function probeContentScriptReady(tabId, retries, delayMs, probeTypes) {
  const totalRetries = Number.isInteger(retries) && retries > 0 ? retries : 8;
  const waitMs = Number.isInteger(delayMs) && delayMs > 0 ? delayMs : 250;
  const probes = Array.isArray(probeTypes) && probeTypes.length
    ? probeTypes
    : ["tdk_export_ping", "tdk_export_get_context"];

  for (let i = 0; i < totalRetries; i += 1) {
    for (const messageType of probes) {
      try {
        const response = await sendMessageRaw(tabId, { type: messageType });
        if (response && response.ok === true) {
          return;
        }
      } catch (_error) {
        // ignore and continue probing
      }
    }

    await wait(waitMs);
  }

  throw new Error("Không nhận được phản hồi từ content script.");
}

function hasScriptingExecuteScript() {
  return Boolean(
    chrome &&
      chrome.scripting &&
      typeof chrome.scripting.executeScript === "function"
  );
}

function hasTabsExecuteScriptFallback() {
  return Boolean(chrome && chrome.tabs && typeof chrome.tabs.executeScript === "function");
}

function hasInjectApiSupport() {
  return hasScriptingExecuteScript() || hasTabsExecuteScriptFallback();
}

function hasManifestPermission(permissionName) {
  try {
    const manifest = chrome && chrome.runtime && typeof chrome.runtime.getManifest === "function"
      ? chrome.runtime.getManifest()
      : null;
    const permissions = manifest && Array.isArray(manifest.permissions) ? manifest.permissions : [];
    return permissions.includes(permissionName);
  } catch (_error) {
    return false;
  }
}

function hasPermissionsApi() {
  return Boolean(
    chrome &&
      chrome.permissions &&
      typeof chrome.permissions.contains === "function" &&
      typeof chrome.permissions.request === "function"
  );
}

async function permissionsContains(origins) {
  return new Promise((resolve, reject) => {
    try {
      chrome.permissions.contains(
        {
          origins
        },
        (hasPermission) => {
          const runtimeError = chrome.runtime.lastError;
          if (runtimeError) {
            reject(new Error(runtimeError.message || "Không kiểm tra được quyền truy cập site."));
            return;
          }
          resolve(Boolean(hasPermission));
        }
      );
    } catch (error) {
      reject(error);
    }
  });
}

async function permissionsRequest(origins) {
  return new Promise((resolve, reject) => {
    try {
      chrome.permissions.request(
        {
          origins
        },
        (granted) => {
          const runtimeError = chrome.runtime.lastError;
          if (runtimeError) {
            reject(new Error(runtimeError.message || "Không yêu cầu được quyền truy cập site."));
            return;
          }
          resolve(Boolean(granted));
        }
      );
    } catch (error) {
      reject(error);
    }
  });
}

async function ensureScheduleHostAccess(options) {
  const opts = options && typeof options === "object" ? options : {};
  if (!hasPermissionsApi()) return;

  let hasAccess = true;
  try {
    hasAccess = await permissionsContains([SCHEDULE_HOST_PERMISSION]);
  } catch (_error) {
    return;
  }

  if (hasAccess) return;

  const allowPrompt = opts.allowPermissionPrompt !== false;
  if (!allowPrompt) {
    throw new Error("Extension chưa có Site access cho sinhvien.uneti.edu.vn.");
  }

  if (typeof opts.onRequestingPermission === "function") {
    opts.onRequestingPermission();
  }

  let granted = false;
  try {
    granted = await permissionsRequest([SCHEDULE_HOST_PERMISSION]);
  } catch (error) {
    const message = getErrorMessage(error, "Không thể yêu cầu quyền Site access.");
    if (message.toLowerCase().includes("user gesture")) {
      throw new Error("Trình duyệt yêu cầu thao tác người dùng để cấp Site access. Hãy bấm lại nút thao tác.");
    }
    throw new Error(`Không thể yêu cầu quyền Site access: ${message}`);
  }

  if (!granted) {
    throw new Error("Bạn chưa cấp Site access cho sinhvien.uneti.edu.vn.");
  }
}

async function runScriptingExecuteScript(tabId, files) {
  return new Promise((resolve, reject) => {
    try {
      chrome.scripting.executeScript(
        {
          target: { tabId },
          files
        },
        () => {
          const runtimeError = chrome.runtime.lastError;
          if (runtimeError) {
            reject(new Error(runtimeError.message || "Không thể inject script bằng scripting API."));
            return;
          }
          resolve();
        }
      );
    } catch (error) {
      reject(error);
    }
  });
}

async function runTabsExecuteScript(tabId, file) {
  return new Promise((resolve, reject) => {
    try {
      chrome.tabs.executeScript(
        tabId,
        {
          file,
          runAt: "document_idle"
        },
        () => {
          const runtimeError = chrome.runtime.lastError;
          if (runtimeError) {
            reject(new Error(runtimeError.message || "Không thể inject script bằng tabs.executeScript."));
            return;
          }
          resolve();
        }
      );
    } catch (error) {
      reject(error);
    }
  });
}

async function injectFilesToTab(tabId, files) {
  const fileList = Array.isArray(files)
    ? files.filter((file) => typeof file === "string" && file.trim())
    : [];

  if (!fileList.length) return;

  if (hasScriptingExecuteScript()) {
    await runScriptingExecuteScript(tabId, fileList);
    return;
  }

  if (hasTabsExecuteScriptFallback()) {
    for (const file of fileList) {
      await runTabsExecuteScript(tabId, file);
    }
    return;
  }

  throw new Error("Trình duyệt hiện tại không hỗ trợ API inject content script.");
}

function normalizeInjectError(error) {
  const message = getErrorMessage(error, "Không thể inject vào tab hiện tại.");
  const lower = message.toLowerCase();

  if (
    lower.includes("cannot access contents of the page") ||
    lower.includes("missing host permission") ||
    lower.includes("permission denied") ||
    lower.includes("forbidden")
  ) {
    return new Error(`Thiếu quyền truy cập tab/site: ${message}`);
  }

  if (lower.includes("no tab with id") || lower.includes("tab not found")) {
    return new Error(`Tab không còn tồn tại: ${message}`);
  }

  if (lower.includes("does not support") || lower.includes("not supported")) {
    return new Error(`Runtime không hỗ trợ inject script: ${message}`);
  }

  return new Error(`Không thể inject vào tab hiện tại: ${message}`);
}

async function reloadTabAndWait(tabId, timeoutMs) {
  const waitTimeout = Number.isInteger(timeoutMs) && timeoutMs > 0 ? timeoutMs : 22000;

  await new Promise((resolve, reject) => {
    let done = false;
    let timer = null;

    const cleanup = () => {
      if (timer) clearTimeout(timer);
      chrome.tabs.onUpdated.removeListener(handleUpdated);
    };

    const finish = (fn) => {
      if (done) return;
      done = true;
      cleanup();
      fn();
    };

    const handleUpdated = (updatedTabId, changeInfo) => {
      if (updatedTabId !== tabId) return;
      if (changeInfo && changeInfo.status === "complete") {
        finish(resolve);
      }
    };

    chrome.tabs.onUpdated.addListener(handleUpdated);
    timer = setTimeout(() => {
      finish(() => reject(new Error("Tab tải lại quá thời gian chờ.")));
    }, waitTimeout);

    chrome.tabs.reload(tabId, { bypassCache: false }, () => {
      const runtimeError = chrome.runtime.lastError;
      if (runtimeError) {
        finish(() => reject(new Error(runtimeError.message || "Không thể tải lại tab.")));
      }
    });
  });
}

async function ensureContentScriptReady(tab, options) {
  const opts = options && typeof options === "object" ? options : {};
  const validateTab = typeof opts.validateTab === "function" ? opts.validateTab : isUnetiHostTab;
  if (!validateTab(tab)) {
    throw new Error(opts.invalidTabMessage || "Tab hiện tại không thuộc cổng sinh viên UNETI.");
  }

  const tabId = tab.id;
  const probeTypes = Array.isArray(opts.probeTypes) && opts.probeTypes.length
    ? opts.probeTypes
    : ["tdk_export_ping", "tdk_export_get_context"];

  await ensureScheduleHostAccess({
    allowPermissionPrompt: opts.allowPermissionPrompt !== false,
    onRequestingPermission: opts.onRequestingPermission
  });

  try {
    await probeContentScriptReady(tabId, 1, 1, probeTypes);
    return;
  } catch (_error) {
    // continue to fallback path
  }

  if (typeof opts.onConnecting === "function") {
    opts.onConnecting();
  }

  const canInject = hasInjectApiSupport();
  if (!canInject) {
    if (typeof opts.onReloading === "function") {
      opts.onReloading();
    }

    try {
      await reloadTabAndWait(tabId, 22000);
      await probeContentScriptReady(tabId, 24, 250, probeTypes);
      return;
    } catch (error) {
      const permissionHint = hasManifestPermission("scripting")
        ? "Runtime hiện tại không hỗ trợ inject script."
        : "Extension hiện chưa nạp quyền scripting (cần reload extension trong trang quản lý tiện ích).";
      throw new Error(
        `${permissionHint} Sau khi tải lại tab vẫn chưa kết nối được content script: ${getErrorMessage(
          error,
          "Không nhận được phản hồi từ content script."
        )}`
      );
    }
  }

  if (typeof opts.onInjecting === "function") {
    opts.onInjecting();
  }

  try {
    // Inject core listener first so popup can reconnect even if vendor libs fail to load.
    await injectFilesToTab(tabId, [CORE_CONTENT_SCRIPT_FILE]);
  } catch (error) {
    throw normalizeInjectError(error);
  }

  try {
    await injectFilesToTab(tabId, VENDOR_SCRIPT_FILES);
  } catch (_error) {
    // Keep connection alive; export handlers will surface missing library errors when needed.
  }

  await wait(180);

  try {
    await probeContentScriptReady(tabId, 12, 220, probeTypes);
  } catch (error) {
    throw new Error(
      `Inject xong nhưng content script chưa phản hồi: ${getErrorMessage(
        error,
        "Không nhận được phản hồi từ content script."
      )}`
    );
  }
}

async function sendMessageToTab(tab, payload, options) {
  if (!tab || typeof tab.id !== "number") {
    throw new Error("Không lấy được tab hiện tại.");
  }

  try {
    return await sendMessageRaw(tab.id, payload);
  } catch (error) {
    if (!isReceivingEndMissing(error)) {
      throw error;
    }

    await ensureContentScriptReady(tab, {
      allowPermissionPrompt: !(options && options.allowPermissionPrompt === false),
      validateTab: options && typeof options.validateTab === "function" ? options.validateTab : undefined,
      invalidTabMessage: options && typeof options.invalidTabMessage === "string" ? options.invalidTabMessage : undefined,
      probeTypes: options && Array.isArray(options.probeTypes) ? options.probeTypes : undefined,
      onConnecting: () => {
        if (options && typeof options.onConnecting === "function") {
          options.onConnecting();
        }
      },
      onInjecting: () => {
        if (options && typeof options.onInjecting === "function") {
          options.onInjecting();
        }
      },
      onReloading: () => {
        if (options && typeof options.onReloading === "function") {
          options.onReloading();
        }
      },
      onRequestingPermission: () => {
        if (options && typeof options.onRequestingPermission === "function") {
          options.onRequestingPermission();
        }
      }
    });

    try {
      return await sendMessageRaw(tab.id, payload);
    } catch (retryError) {
      if (isReceivingEndMissing(retryError)) {
        throw new Error("Đã tự kết nối nhưng tab vẫn chưa nhận content script.");
      }
      throw retryError;
    }
  }
}

async function sendGradeMessageToTab(tab, payload, options) {
  const opts = options && typeof options === "object" ? options : {};
  return sendMessageToTab(tab, payload, {
    ...opts,
    validateTab: isGradeTab,
    invalidTabMessage: "Tab hiện tại không phải trang điểm UNETI.",
    probeTypes: ["tdk_grade_ping", "tdk_grade_get_context"]
  });
}

function fillExportContext(context) {
  el.exportScheduleType.value = context && context.scheduleTypeLabel ? context.scheduleTypeLabel : "--";

  const monthFromContext = context && context.currentMonth ? context.currentMonth : "";
  if (monthFromContext) {
    el.exportMonth.value = monthFromContext;
    return;
  }

  if (!el.exportMonth.value) {
    const now = new Date();
    el.exportMonth.value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  }
}

function buildExportOptionsFromForm() {
  return {
    format: el.exportFormat.value,
    rangeMode: el.exportRange.value,
    month: el.exportMonth.value,
    monthLayout: el.exportMonthLayout.value
  };
}

async function refreshExportContext(showSuccessMessage, options) {
  const opts = options && typeof options === "object" ? options : {};
  setStatus(el.exportStatus, "", "");

  try {
    const tab = await getActiveTab();
    if (!isScheduleTab(tab)) {
      fillExportContext(null);
      setStatus(el.exportStatus, "Tab hiện tại không phải trang lịch UNETI.", "error");
      return;
    }

    const response = await sendMessageToTab(
      tab,
      { type: "tdk_export_get_context" },
      {
        allowPermissionPrompt: showSuccessMessage === true,
        onConnecting: () => {
          setStatus(el.exportStatus, "Đang kết nối content script với tab lịch...", "info");
        },
        onInjecting: () => {
          setStatus(el.exportStatus, "Đang inject content script vào tab lịch...", "info");
        },
        onReloading: () => {
          setStatus(el.exportStatus, "Đang tải lại tab lịch để kết nối content script...", "info");
        },
        onRequestingPermission: () => {
          setStatus(el.exportStatus, "Đang yêu cầu Site access cho sinhvien.uneti.edu.vn...", "info");
        }
      }
    );
    if (!response || response.ok !== true) {
      throw new Error(response && response.error ? response.error : "Không lấy được ngữ cảnh xuất.");
    }

    fillExportContext(response.context);
    if (showSuccessMessage) {
      setStatus(el.exportStatus, "Đã làm mới ngữ cảnh tab lịch.", "info");
    }
  } catch (error) {
    fillExportContext(null);
    if (opts.silentOnFail) {
      setStatus(el.exportStatus, "Chưa lấy được ngữ cảnh tab lịch. Bạn có thể bấm \"Làm mới ngữ cảnh\".", "info");
    } else {
      setStatus(el.exportStatus, `Không lấy được ngữ cảnh: ${error.message}`, "error");
    }
  }
}

async function handleOpenExportPanel() {
  setExportBusy(true);
  setStatus(el.exportStatus, "", "");

  try {
    const tab = await getActiveTab();
    if (!isScheduleTab(tab)) {
      setStatus(el.exportStatus, "Tab hiện tại không phải trang lịch UNETI.", "error");
      return;
    }

    const response = await sendMessageToTab(
      tab,
      { type: "tdk_export_open_panel" },
      {
        allowPermissionPrompt: true,
        onConnecting: () => {
          setStatus(el.exportStatus, "Đang kết nối content script với tab lịch...", "info");
        },
        onInjecting: () => {
          setStatus(el.exportStatus, "Đang inject content script vào tab lịch...", "info");
        },
        onReloading: () => {
          setStatus(el.exportStatus, "Đang tải lại tab lịch để kết nối content script...", "info");
        },
        onRequestingPermission: () => {
          setStatus(el.exportStatus, "Đang yêu cầu Site access cho sinhvien.uneti.edu.vn...", "info");
        }
      }
    );
    if (!response || response.ok !== true) {
      throw new Error(response && response.error ? response.error : "Không mở được panel xuất trên trang.");
    }

    setStatus(el.exportStatus, "Đã mở panel xuất trên trang lịch.", "success");
  } catch (error) {
    setStatus(el.exportStatus, `Mở panel thất bại: ${error.message}`, "error");
  } finally {
    setExportBusy(false);
  }
}

async function handleExportNow() {
  setExportBusy(true);
  setStatus(el.exportStatus, "", "");

  try {
    const tab = await getActiveTab();
    if (!isScheduleTab(tab)) {
      setStatus(el.exportStatus, "Tab hiện tại không phải trang lịch UNETI.", "error");
      return;
    }

    const options = buildExportOptionsFromForm();
    const response = await sendMessageToTab(
      tab,
      {
        type: "tdk_export_execute",
        options
      },
      {
        allowPermissionPrompt: true,
        onConnecting: () => {
          setStatus(el.exportStatus, "Đang kết nối content script với tab lịch...", "info");
        },
        onInjecting: () => {
          setStatus(el.exportStatus, "Đang inject content script vào tab lịch...", "info");
        },
        onReloading: () => {
          setStatus(el.exportStatus, "Đang tải lại tab lịch để kết nối content script...", "info");
        },
        onRequestingPermission: () => {
          setStatus(el.exportStatus, "Đang yêu cầu Site access cho sinhvien.uneti.edu.vn...", "info");
        }
      }
    );

    if (!response || response.ok !== true) {
      throw new Error(response && response.error ? response.error : "Xuất lịch thất bại.");
    }

    const total = response.result && Number.isFinite(response.result.totalEvents)
      ? response.result.totalEvents
      : 0;

    setStatus(el.exportStatus, `Xuất lịch thành công (${total} sự kiện).`, "success");
  } catch (error) {
    setStatus(el.exportStatus, `Xuất lịch thất bại: ${error.message}`, "error");
  } finally {
    setExportBusy(false);
  }
}

function setGradeAvailability(ready, contextText, hintText) {
  gradeExportReady = ready === true;
  if (el.btnGradeExportNow) {
    el.btnGradeExportNow.disabled = !gradeExportReady;
  }
  if (el.gradeFormat) {
    el.gradeFormat.disabled = !gradeExportReady;
  }
  if (el.gradeContext) {
    el.gradeContext.value = contextText || "--";
  }
  if (el.gradeHint) {
    el.gradeHint.textContent = hintText || "Mở trang Kết quả học tập để xuất; sửa điểm trực tiếp sẽ tự cập nhật GPA. Giữ Ctrl/Cmd + click để chọn nhiều môn, rồi bấm Tính TB thường kỳ.";
  }
}

function setGradeBusy(busy) {
  if (el.btnRefreshGradeContext) {
    el.btnRefreshGradeContext.disabled = busy;
  }
  if (el.btnGradeExportNow) {
    el.btnGradeExportNow.disabled = busy || !gradeExportReady;
  }
  if (el.gradeFormat) {
    el.gradeFormat.disabled = busy || !gradeExportReady;
  }
}

function buildGradeContextText(context) {
  if (!context || context.isGradePage !== true) {
    return "--";
  }
  if (context.hasTable !== true) {
    return "Trang điểm chưa có bảng dữ liệu";
  }
  const totalRows = Number.isFinite(context.totalRows) ? context.totalRows : 0;
  return `Trang điểm hợp lệ (${totalRows} dòng)`;
}

function buildGradeExportOptionsFromForm() {
  const rawFormat = text(el.gradeFormat && el.gradeFormat.value).toLowerCase();
  const format = GRADE_EXPORT_FORMATS.has(rawFormat) ? rawFormat : "xlsx";
  return { format };
}

async function refreshGradeContext(showSuccessMessage, options) {
  const opts = options && typeof options === "object" ? options : {};
  setStatus(el.gradeStatus, "", "");

  try {
    const tab = await getActiveTab();
    if (!isGradeTab(tab)) {
      setGradeAvailability(false, "--", "Mở trang Kết quả học tập để xuất; sửa điểm trực tiếp sẽ tự cập nhật GPA. Giữ Ctrl/Cmd + click để chọn nhiều môn, rồi bấm Tính TB thường kỳ.");
      if (opts.silentOnFail) {
        setStatus(el.gradeStatus, "Mở trang Kết quả học tập để bật xuất điểm.", "info");
      } else {
        setStatus(el.gradeStatus, "Tab hiện tại không phải trang điểm UNETI.", "error");
      }
      return;
    }

    const response = await sendGradeMessageToTab(
      tab,
      { type: "tdk_grade_get_context" },
      {
        allowPermissionPrompt: showSuccessMessage === true,
        onConnecting: () => {
          setStatus(el.gradeStatus, "Đang kết nối content script với tab điểm...", "info");
        },
        onInjecting: () => {
          setStatus(el.gradeStatus, "Đang inject content script vào tab điểm...", "info");
        },
        onReloading: () => {
          setStatus(el.gradeStatus, "Đang tải lại tab điểm để kết nối content script...", "info");
        },
        onRequestingPermission: () => {
          setStatus(el.gradeStatus, "Đang yêu cầu Site access cho sinhvien.uneti.edu.vn...", "info");
        }
      }
    );

    if (!response || response.ok !== true) {
      throw new Error(response && response.error ? response.error : "Không lấy được ngữ cảnh trang điểm.");
    }

    const context = response.context || {};
    const isReady = context.isGradePage === true && context.hasTable === true;
    setGradeAvailability(
      isReady,
      buildGradeContextText(context),
      isReady
        ? "Sẵn sàng xuất bảng điểm. Trên website bạn có thể sửa trực tiếp điểm để GPA cập nhật ngay; nút Tính TB thường kỳ chỉ áp dụng cho các môn đã chọn."
        : "Trang điểm đã mở nhưng chưa tìm thấy bảng #xemDiem_aaa."
    );

    if (showSuccessMessage) {
      setStatus(
        el.gradeStatus,
        isReady ? "Đã làm mới ngữ cảnh tab điểm." : "Đã làm mới, nhưng chưa thấy bảng điểm để xuất.",
        "info"
      );
    }
  } catch (error) {
    setGradeAvailability(false, "--", "Mở trang Kết quả học tập để xuất; sửa điểm trực tiếp sẽ tự cập nhật GPA. Giữ Ctrl/Cmd + click để chọn nhiều môn, rồi bấm Tính TB thường kỳ.");
    if (opts.silentOnFail) {
      setStatus(el.gradeStatus, "Chưa lấy được ngữ cảnh tab điểm. Bạn có thể bấm \"Làm mới ngữ cảnh điểm\".", "info");
    } else {
      setStatus(el.gradeStatus, `Không lấy được ngữ cảnh điểm: ${error.message}`, "error");
    }
  }
}

async function handleGradeExportNow() {
  setGradeBusy(true);
  setStatus(el.gradeStatus, "", "");

  try {
    const tab = await getActiveTab();
    if (!isGradeTab(tab)) {
      setGradeAvailability(false, "--", "Mở trang Kết quả học tập để xuất; sửa điểm trực tiếp sẽ tự cập nhật GPA. Giữ Ctrl/Cmd + click để chọn nhiều môn, rồi bấm Tính TB thường kỳ.");
      setStatus(el.gradeStatus, "Tab hiện tại không phải trang điểm UNETI.", "error");
      return;
    }

    const options = buildGradeExportOptionsFromForm();
    const response = await sendGradeMessageToTab(
      tab,
      {
        type: "tdk_grade_export_execute",
        options
      },
      {
        allowPermissionPrompt: true,
        onConnecting: () => {
          setStatus(el.gradeStatus, "Đang kết nối content script với tab điểm...", "info");
        },
        onInjecting: () => {
          setStatus(el.gradeStatus, "Đang inject content script vào tab điểm...", "info");
        },
        onReloading: () => {
          setStatus(el.gradeStatus, "Đang tải lại tab điểm để kết nối content script...", "info");
        },
        onRequestingPermission: () => {
          setStatus(el.gradeStatus, "Đang yêu cầu Site access cho sinhvien.uneti.edu.vn...", "info");
        }
      }
    );

    if (!response || response.ok !== true) {
      throw new Error(response && response.error ? response.error : "Xuất điểm thất bại.");
    }

    const total = response.result && Number.isFinite(response.result.totalRecords)
      ? response.result.totalRecords
      : 0;
    setStatus(el.gradeStatus, `Xuất điểm thành công (${total} môn).`, "success");
  } catch (error) {
    setStatus(el.gradeStatus, `Xuất điểm thất bại: ${error.message}`, "error");
  } finally {
    setGradeBusy(false);
  }
}

function normalizeRecord(record, index) {
  const src = record && typeof record === "object" ? record : {};

  const slot = text(src.slot).toLowerCase() === "afternoon" || text(src.timeRange) === SLOT_TIME.afternoon
    ? "afternoon"
    : "morning";

  const mode = src.mode === "date" ? "date" : "weekly";
  const weekdayNumber = Number(src.weekday);

  return {
    id: text(src.id) || `mp-${index + 1}`,
    enabled: src.enabled !== false,
    mode,
    weekday: Number.isInteger(weekdayNumber) && weekdayNumber >= 2 && weekdayNumber <= 8 ? weekdayNumber : 2,
    date: parseDDMMYYYY(src.date) ? text(src.date) : "",
    slot,
    timeRange: slot === "afternoon" ? SLOT_TIME.afternoon : SLOT_TIME.morning,
    subject: text(src.subject),
    class: text(src.class),
    room: text(src.room),
    teacher: text(src.teacher)
  };
}

function normalizeConfig(raw) {
  const source = raw && typeof raw === "object" ? raw : {};
  const maps = source.maps && typeof source.maps === "object" ? source.maps : {};
  const branding = source.branding && typeof source.branding === "object" ? source.branding : {};

  const normalized = {
    version: 2,
    maps: {
      normal: maps.normal && typeof maps.normal === "object" ? maps.normal : {},
      practical: maps.practical && typeof maps.practical === "object" ? maps.practical : {}
    },
    rules: Array.isArray(source.rules) ? source.rules : [],
    manualPracticalSchedules: (Array.isArray(source.manualPracticalSchedules)
      ? source.manualPracticalSchedules
      : []
    ).map(normalizeRecord),
    branding: {
      companyName: text(branding.companyName) || DEFAULT_COMPANY,
      logoPath: text(branding.logoPath) || DEFAULT_LOGO_PATH,
      linkUrl: text(branding.linkUrl)
    }
  };

  if (!isValidUrl(normalized.branding.linkUrl)) {
    normalized.branding.linkUrl = "";
  }

  return normalized;
}

function validateMap(mapValue, mapName, allowEmpty) {
  const errors = [];

  if (!mapValue || typeof mapValue !== "object" || Array.isArray(mapValue)) {
    errors.push(`maps.${mapName} phải là object.`);
    return errors;
  }

  const keys = Object.keys(mapValue);
  if (!allowEmpty && keys.length === 0) {
    errors.push(`maps.${mapName} không được rỗng.`);
    return errors;
  }

  for (const key of keys) {
    const item = mapValue[key];
    const start = item && typeof item.start === "string" ? text(item.start) : "";
    const end = item && typeof item.end === "string" ? text(item.end) : "";

    if (!start || !end) {
      errors.push(`maps.${mapName}.${key} phải có start/end.`);
      continue;
    }

    if (!HHMM_REGEX.test(start) || !HHMM_REGEX.test(end)) {
      errors.push(`maps.${mapName}.${key} phải theo HH:mm.`);
    }
  }

  return errors;
}

function validateRules(rules) {
  const errors = [];
  if (!Array.isArray(rules)) return ["rules phải là mảng."];

  rules.forEach((rule, index) => {
    if (!rule || typeof rule !== "object" || Array.isArray(rule)) {
      errors.push(`rules[${index}] phải là object.`);
      return;
    }

    if (rule.enabled === false) return;

    for (const field of ["subject", "class", "room", "teacher"]) {
      if (!text(rule[field])) {
        errors.push(`rules[${index}].${field} là bắt buộc khi rule bật.`);
      }
    }
  });

  return errors;
}

function validateManualRecords(records) {
  const errors = [];
  if (!Array.isArray(records)) return ["manualPracticalSchedules phải là mảng."];

  records.forEach((record, index) => {
    if (!record || typeof record !== "object" || Array.isArray(record)) {
      errors.push(`manualPracticalSchedules[${index}] phải là object.`);
      return;
    }

    if (record.enabled === false) return;

    if (!text(record.subject) || !text(record.class) || !text(record.room) || !text(record.teacher)) {
      errors.push(`manualPracticalSchedules[${index}] thiếu Môn/Lớp/Phòng/GV.`);
    }

    if (record.mode === "date") {
      if (!parseDDMMYYYY(record.date)) {
        errors.push(`manualPracticalSchedules[${index}].date không hợp lệ.`);
      }
    } else {
      const weekdayNumber = Number(record.weekday);
      if (!(Number.isInteger(weekdayNumber) && weekdayNumber >= 2 && weekdayNumber <= 8)) {
        errors.push(`manualPracticalSchedules[${index}].weekday không hợp lệ.`);
      }
    }

    if (!(record.timeRange === SLOT_TIME.morning || record.timeRange === SLOT_TIME.afternoon)) {
      errors.push(`manualPracticalSchedules[${index}].timeRange chỉ nhận 06:00-11:30 hoặc 13:00-18:30.`);
    }
  });

  return errors;
}

function validateBranding(branding) {
  if (!branding || typeof branding !== "object" || Array.isArray(branding)) {
    return ["branding phải là object."];
  }

  const errors = [];

  if (!text(branding.companyName)) errors.push("branding.companyName là bắt buộc.");
  if (!text(branding.logoPath)) errors.push("branding.logoPath là bắt buộc.");
  if (!isValidUrl(branding.linkUrl)) errors.push("branding.linkUrl không hợp lệ.");

  return errors;
}

function extractManualRecordsFromImport(payload) {
  if (Array.isArray(payload)) {
    return payload;
  }

  if (payload && typeof payload === "object" && Array.isArray(payload.manualPracticalSchedules)) {
    return payload.manualPracticalSchedules;
  }

  throw new Error("File import phải là mảng hoặc object có manualPracticalSchedules.");
}

function normalizeImportedManualRecords(records) {
  const normalizedRecords = (Array.isArray(records) ? records : []).map(normalizeRecord);
  const usedIds = new Set();

  return normalizedRecords.map((record, index) => {
    const baseId = text(record.id) || `mp-import-${index + 1}`;
    let candidate = baseId;
    let suffix = 2;

    while (usedIds.has(candidate)) {
      candidate = `${baseId}-${suffix}`;
      suffix += 1;
    }

    usedIds.add(candidate);
    return {
      ...record,
      id: candidate
    };
  });
}

function validateConfig(config) {
  if (!config || typeof config !== "object" || Array.isArray(config)) {
    return ["JSON gốc phải là object."];
  }

  const errors = [];
  errors.push(...validateMap(config.maps && config.maps.normal, "normal", false));
  errors.push(...validateMap(config.maps && config.maps.practical, "practical", true));
  errors.push(...validateRules(config.rules));
  errors.push(...validateManualRecords(config.manualPracticalSchedules));
  errors.push(...validateBranding(config.branding));

  return errors;
}

async function loadDefaultConfig() {
  const response = await fetch(DEFAULT_CONFIG_URL);
  if (!response.ok) {
    throw new Error("Không thể tải file cấu hình mặc định.");
  }

  const data = await response.json();
  return normalizeConfig(data);
}

async function loadStoredConfig() {
  const data = await chrome.storage.local.get(CONFIG_KEY);
  if (data[CONFIG_KEY]) {
    return normalizeConfig(data[CONFIG_KEY]);
  }
  return loadDefaultConfig();
}

async function persistCurrentConfig() {
  currentConfig = normalizeConfig(currentConfig);
  await chrome.storage.local.set({ [CONFIG_KEY]: currentConfig });
}

function renderBrand() {
  const branding = currentConfig.branding || {};
  const company = text(branding.companyName) || DEFAULT_COMPANY;
  const logoPath = text(branding.logoPath) || DEFAULT_LOGO_PATH;

  el.brandLogo.src = chrome.runtime.getURL(logoPath);
  el.brandText.textContent = `Bản quyền ${company}`;
  el.brandingLinkInput.value = text(branding.linkUrl);
}

function toggleManualMode() {
  const isDateMode = el.manualMode.value === "date";
  el.manualWeekdayWrap.style.display = isDateMode ? "none" : "";
  el.manualDateWrap.style.display = isDateMode ? "" : "none";
}

function resetManualForm() {
  el.manualEditId.value = "";
  el.manualSubject.value = "";
  el.manualClass.value = "";
  el.manualRoom.value = "";
  el.manualTeacher.value = "";
  el.manualMode.value = "weekly";
  el.manualWeekday.value = "2";
  el.manualDate.value = "";
  el.manualSlot.value = "morning";
  el.manualEnabled.checked = true;
  toggleManualMode();
  setStatus(el.manualStatus, "", "");
}

function renderManualTable() {
  const records = Array.isArray(currentConfig.manualPracticalSchedules) ? currentConfig.manualPracticalSchedules : [];

  if (records.length === 0) {
    el.manualTableBody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:#64748b">Chưa có lịch thực hành bổ sung.</td></tr>';
    return;
  }

  el.manualTableBody.innerHTML = records
    .map((record, index) => {
      const modeText = record.mode === "date"
        ? `Theo ngày (${record.date || "--"})`
        : `Hàng tuần (${WEEKDAY_LABEL[record.weekday] || "--"})`;

      return `
      <tr>
        <td>${index + 1}</td>
        <td>${esc(record.subject)}</td>
        <td>${esc(record.class)}</td>
        <td>${esc(record.room)}</td>
        <td>${esc(record.teacher)}</td>
        <td>${esc(modeText)}</td>
        <td>${esc(record.timeRange)}</td>
        <td>
          <span class="row-actions">
            <label><input type="checkbox" class="manualEnabledToggle" data-id="${esc(record.id)}" ${record.enabled ? "checked" : ""}> Bật</label>
            <button type="button" class="manualEdit" data-id="${esc(record.id)}">Sửa</button>
            <button type="button" class="manualDelete" data-id="${esc(record.id)}">Xóa</button>
          </span>
        </td>
      </tr>
    `;
    })
    .join("");
}

function syncEditorFromCurrentConfig() {
  el.configInput.value = JSON.stringify(currentConfig, null, 2);
}

function renderAll() {
  renderBrand();
  renderManualTable();
  syncEditorFromCurrentConfig();
}

function buildManualRecordFromForm() {
  const mode = el.manualMode.value === "date" ? "date" : "weekly";
  const slot = el.manualSlot.value === "afternoon" ? "afternoon" : "morning";

  return {
    id: text(el.manualEditId.value) || `mp-${Date.now()}-${Math.floor(Math.random() * 100000)}`,
    enabled: el.manualEnabled.checked,
    mode,
    weekday: Number(el.manualWeekday.value) || 2,
    date: mode === "date" ? fromDateInputValue(el.manualDate.value) : "",
    slot,
    timeRange: slot === "afternoon" ? SLOT_TIME.afternoon : SLOT_TIME.morning,
    subject: text(el.manualSubject.value),
    class: text(el.manualClass.value),
    room: text(el.manualRoom.value),
    teacher: text(el.manualTeacher.value)
  };
}

function validateManualFormRecord(record) {
  if (!record.subject || !record.class || !record.room || !record.teacher) {
    return "Vui lòng nhập đầy đủ Môn, Lớp, Phòng, Giảng viên.";
  }

  if (record.mode === "date" && !parseDDMMYYYY(record.date)) {
    return "Ngày cụ thể không hợp lệ.";
  }

  if (record.mode === "weekly" && !(Number.isInteger(record.weekday) && record.weekday >= 2 && record.weekday <= 8)) {
    return "Thứ không hợp lệ.";
  }

  if (!(record.timeRange === SLOT_TIME.morning || record.timeRange === SLOT_TIME.afternoon)) {
    return "Chỉ cho phép 06:00-11:30 hoặc 13:00-18:30.";
  }

  return "";
}

function fillFormForEdit(record) {
  el.manualEditId.value = record.id;
  el.manualSubject.value = record.subject;
  el.manualClass.value = record.class;
  el.manualRoom.value = record.room;
  el.manualTeacher.value = record.teacher;
  el.manualMode.value = record.mode;
  el.manualWeekday.value = String(record.weekday || 2);
  el.manualDate.value = toDateInputValue(record.date);
  el.manualSlot.value = record.slot;
  el.manualEnabled.checked = record.enabled !== false;

  toggleManualMode();
  setStatus(el.manualStatus, "Đang chỉnh sửa bản ghi.", "info");
}

function buildExportStamp() {
  return new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
}

function downloadJsonFile(fileName, payload) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);

  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();

  URL.revokeObjectURL(url);
}

async function handleManualSave(event) {
  event.preventDefault();

  const record = buildManualRecordFromForm();
  const error = validateManualFormRecord(record);

  if (error) {
    setStatus(el.manualStatus, error, "error");
    return;
  }

  const records = currentConfig.manualPracticalSchedules || [];
  const index = records.findIndex((item) => item.id === record.id);

  if (index >= 0) {
    records[index] = record;
  } else {
    records.push(record);
  }

  currentConfig.manualPracticalSchedules = records;
  await persistCurrentConfig();

  renderManualTable();
  syncEditorFromCurrentConfig();
  setStatus(el.manualStatus, index >= 0 ? "Đã cập nhật bản ghi." : "Đã thêm bản ghi.", "success");
  resetManualForm();
}

function handleManualImportClick() {
  el.manualImportFileInput.click();
}

async function handleManualImportFile(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  try {
    const payload = JSON.parse(await file.text());
    const sourceRecords = extractManualRecordsFromImport(payload);
    const normalizedRecords = normalizeImportedManualRecords(sourceRecords);
    const errors = validateManualRecords(normalizedRecords);

    if (errors.length) {
      setStatus(el.manualStatus, `Import lỗi: ${errors.join(" ")}`, "error");
      return;
    }

    const accepted = window.confirm("Import sẽ ghi đè toàn bộ lịch thực hành bổ sung hiện tại. Tiếp tục?");
    if (!accepted) {
      setStatus(el.manualStatus, "Đã hủy import lịch thực hành bổ sung.", "info");
      return;
    }

    currentConfig.manualPracticalSchedules = normalizedRecords;
    await persistCurrentConfig();

    renderManualTable();
    syncEditorFromCurrentConfig();
    resetManualForm();
    setStatus(el.manualStatus, `Đã import ${normalizedRecords.length} lịch thực hành bổ sung.`, "success");
  } catch (error) {
    setStatus(el.manualStatus, `Import thất bại: ${error.message}`, "error");
  } finally {
    event.target.value = "";
  }
}

function handleManualExport() {
  try {
    const records = normalizeImportedManualRecords(currentConfig.manualPracticalSchedules || []);
    const payload = {
      version: 1,
      exportedAt: new Date().toISOString(),
      manualPracticalSchedules: records
    };

    const stamp = buildExportStamp();
    downloadJsonFile(`uneti-time-mapper-manual-practical-${stamp}.json`, payload);
    setStatus(el.manualStatus, `Đã export ${records.length} lịch thực hành bổ sung.`, "success");
  } catch (error) {
    setStatus(el.manualStatus, `Export thất bại: ${error.message}`, "error");
  }
}

async function handleTableClick(event) {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  const editButton = target.closest(".manualEdit");
  if (editButton) {
    const id = text(editButton.getAttribute("data-id"));
    const record = (currentConfig.manualPracticalSchedules || []).find((item) => item.id === id);
    if (record) fillFormForEdit(record);
    return;
  }

  const deleteButton = target.closest(".manualDelete");
  if (deleteButton) {
    const id = text(deleteButton.getAttribute("data-id"));
    const accepted = window.confirm("Bạn có chắc muốn xóa lịch thực hành bổ sung này?");
    if (!accepted) return;

    currentConfig.manualPracticalSchedules = (currentConfig.manualPracticalSchedules || []).filter((item) => item.id !== id);
    await persistCurrentConfig();

    renderManualTable();
    syncEditorFromCurrentConfig();
    setStatus(el.manualStatus, "Đã xóa bản ghi.", "success");
    resetManualForm();
  }
}

async function handleTableChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) || !target.classList.contains("manualEnabledToggle")) return;

  const id = text(target.getAttribute("data-id"));
  const records = currentConfig.manualPracticalSchedules || [];
  const index = records.findIndex((item) => item.id === id);
  if (index < 0) return;

  records[index].enabled = target.checked;
  currentConfig.manualPracticalSchedules = records;
  await persistCurrentConfig();

  renderManualTable();
  syncEditorFromCurrentConfig();
}

async function handleBrandingSave() {
  const link = text(el.brandingLinkInput.value);

  if (!isValidUrl(link)) {
    setStatus(el.brandingStatus, "Link thương hiệu không hợp lệ.", "error");
    return;
  }

  currentConfig.branding.linkUrl = link;
  await persistCurrentConfig();

  renderBrand();
  syncEditorFromCurrentConfig();
  setStatus(el.brandingStatus, "Đã cập nhật branding.", "success");
}

function parseEditorConfig() {
  const raw = el.configInput.value.trim();
  if (!raw) throw new Error("JSON đang trống.");
  return JSON.parse(raw);
}

async function handleSaveConfig() {
  try {
    const parsed = normalizeConfig(parseEditorConfig());
    const errors = validateConfig(parsed);

    if (errors.length) {
      setStatus(el.configStatus, errors.join(" "), "error");
      return;
    }

    currentConfig = parsed;
    await persistCurrentConfig();
    renderAll();

    setStatus(el.configStatus, "Đã lưu cấu hình thành công.", "success");
  } catch (error) {
    setStatus(el.configStatus, `Lưu thất bại: ${error.message}`, "error");
  }
}

function handleSyncFromJson() {
  try {
    const parsed = normalizeConfig(parseEditorConfig());
    const errors = validateConfig(parsed);

    if (errors.length) {
      setStatus(el.configStatus, errors.join(" "), "error");
      return;
    }

    currentConfig = parsed;
    renderAll();
    setStatus(el.configStatus, "Đã đồng bộ từ JSON (chưa lưu).", "info");
  } catch (error) {
    setStatus(el.configStatus, `Đồng bộ thất bại: ${error.message}`, "error");
  }
}

async function handleLoadDefault() {
  try {
    currentConfig = await loadDefaultConfig();
    renderAll();
    setStatus(el.configStatus, "Đã nạp cấu hình mặc định (chưa lưu).", "info");
  } catch (error) {
    setStatus(el.configStatus, `Không tải được mặc định: ${error.message}`, "error");
  }
}

function handleImportClick() {
  el.importFileInput.click();
}

function handleImportFile(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function () {
    try {
      const parsed = normalizeConfig(JSON.parse(String(reader.result || "")));
      const errors = validateConfig(parsed);

      if (errors.length) {
        setStatus(el.configStatus, `Import lỗi: ${errors.join(" ")}`, "error");
        return;
      }

      currentConfig = parsed;
      renderAll();
      setStatus(el.configStatus, "Đã import JSON (chưa lưu).", "info");
    } catch (error) {
      setStatus(el.configStatus, `Import thất bại: ${error.message}`, "error");
    }
  };

  reader.onerror = function () {
    setStatus(el.configStatus, "Không đọc được file import.", "error");
  };

  reader.readAsText(file, "utf-8");
}

function handleExport() {
  try {
    let exportConfig;

    try {
      const parsed = normalizeConfig(parseEditorConfig());
      const errors = validateConfig(parsed);
      if (errors.length) throw new Error(errors.join(" "));
      exportConfig = parsed;
    } catch (_error) {
      exportConfig = currentConfig;
    }

    const stamp = buildExportStamp();
    downloadJsonFile(`uneti-time-mapper-config-${stamp}.json`, exportConfig);
    setStatus(el.configStatus, "Đã export JSON cấu hình.", "success");
  } catch (error) {
    setStatus(el.configStatus, `Export thất bại: ${error.message}`, "error");
  }
}

function bindEvents() {
  el.manualMode.addEventListener("change", toggleManualMode);
  el.manualForm.addEventListener("submit", (event) => {
    void handleManualSave(event);
  });
  el.btnManualReset.addEventListener("click", resetManualForm);
  el.btnManualImport.addEventListener("click", handleManualImportClick);
  el.manualImportFileInput.addEventListener("change", (event) => {
    void handleManualImportFile(event);
  });
  el.btnManualExport.addEventListener("click", handleManualExport);

  el.manualTableBody.addEventListener("click", (event) => {
    void handleTableClick(event);
  });
  el.manualTableBody.addEventListener("change", (event) => {
    void handleTableChange(event);
  });

  el.btnBrandingSave.addEventListener("click", () => {
    void handleBrandingSave();
  });

  el.exportRange.addEventListener("change", toggleExportRangeControls);
  el.btnRefreshExportContext.addEventListener("click", () => {
    void refreshExportContext(true);
  });
  el.btnOpenExportPanel.addEventListener("click", () => {
    void handleOpenExportPanel();
  });
  el.btnExportNow.addEventListener("click", () => {
    void handleExportNow();
  });
  el.btnRefreshGradeContext.addEventListener("click", () => {
    void refreshGradeContext(true);
  });
  el.btnGradeExportNow.addEventListener("click", () => {
    void handleGradeExportNow();
  });

  el.btnSaveConfig.addEventListener("click", () => {
    void handleSaveConfig();
  });
  el.btnSyncFromJson.addEventListener("click", handleSyncFromJson);
  el.btnLoadDefault.addEventListener("click", () => {
    void handleLoadDefault();
  });

  el.btnImport.addEventListener("click", handleImportClick);
  el.importFileInput.addEventListener("change", handleImportFile);
  el.btnExport.addEventListener("click", handleExport);
}

async function init() {
  currentConfig = await loadStoredConfig();
  bindEvents();
  toggleManualMode();
  toggleExportRangeControls();
  resetManualForm();
  renderAll();
  fillExportContext(null);
  setGradeAvailability(false, "--", "Mở trang Kết quả học tập để xuất; sửa điểm trực tiếp sẽ tự cập nhật GPA. Giữ Ctrl/Cmd + click để chọn nhiều môn, rồi bấm Tính TB thường kỳ.");
  await refreshExportContext(false, { silentOnFail: true });
  await refreshGradeContext(false, { silentOnFail: true });
  setStatus(el.configStatus, "Đã tải cấu hình hiện tại.", "info");
}

void init();
