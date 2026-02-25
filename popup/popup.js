"use strict";

const CONFIG_KEY = "unetiScheduleConfigV1";
const DEFAULT_CONFIG_URL = chrome.runtime.getURL("config/schedule-default.json");
const DEFAULT_COMPANY = "TranDangKhoaTechnology";
const DEFAULT_LOGO_PATH = "assets/image/logo.png";
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

let currentConfig = null;

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
  resetManualForm();
  renderAll();
  setStatus(el.configStatus, "Đã tải cấu hình hiện tại.", "info");
}

void init();
