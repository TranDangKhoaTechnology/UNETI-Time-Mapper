(function () {
  "use strict";

  const CONFIG_KEY = "unetiScheduleConfigV1";
  const DEFAULT_CONFIG_URL = chrome.runtime.getURL("config/schedule-default.json");
  const DEFAULT_COMPANY = "TranDangKhoaTechnology";
  const DEFAULT_LOGO_PATH = "assets/image/logo.png";

  const TIET_REGEX = /(?:tiet)\s*:?\s*(\d+)\s*-\s*(\d+)/;
  const REAL_TIME_REGEX = /(?:\b\d{1,2}h\d{2}\b|\b\d{1,2}:\d{2}\b)/i;

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

  let debounceTimer = null;
  let convertLock = false;
  let modalConfig = null;
  let refreshHooksBound = false;
  let jqueryAjaxHookBound = false;
  let observerStarted = false;

  const text = (value) => String(value || "").replace(/\s+/g, " ").trim();
  const norm = (value) =>
    String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  const esc = (value) =>
    String(value || "")
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

  function fromDateInput(value) {
    const match = text(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return "";
    return `${match[3]}/${match[2]}/${match[1]}`;
  }

  function toDateInput(value) {
    const date = parseDDMMYYYY(value);
    if (!date) return "";

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  function jsDayToWeekday(jsDay) {
    return jsDay === 0 ? 8 : jsDay + 1;
  }

  function isWeeklyPage() {
    return location.pathname.toLowerCase().includes("/lich-theo-tuan.html");
  }

  function isProgressPage() {
    return location.pathname.toLowerCase().includes("/lich-theo-tien-do.html");
  }

  function isSchedulePage() {
    return isWeeklyPage() || isProgressPage();
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
    const source = record && typeof record === "object" ? record : {};

    const slot =
      text(source.slot).toLowerCase() === "afternoon" || text(source.timeRange) === SLOT_TIME.afternoon
        ? "afternoon"
        : "morning";

    const mode = source.mode === "date" ? "date" : "weekly";
    const weekdayNumber = Number(source.weekday);

    return {
      id: text(source.id) || `mp-${index + 1}`,
      enabled: source.enabled !== false,
      mode,
      weekday: Number.isInteger(weekdayNumber) && weekdayNumber >= 2 && weekdayNumber <= 8 ? weekdayNumber : 2,
      date: parseDDMMYYYY(source.date) ? text(source.date) : "",
      slot,
      timeRange: slot === "afternoon" ? SLOT_TIME.afternoon : SLOT_TIME.morning,
      subject: text(source.subject),
      class: text(source.class),
      room: text(source.room),
      teacher: text(source.teacher)
    };
  }

  function normalizeConfig(raw) {
    const source = raw && typeof raw === "object" ? raw : {};
    const maps = source.maps && typeof source.maps === "object" ? source.maps : {};
    const branding = source.branding && typeof source.branding === "object" ? source.branding : {};

    const config = {
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

    if (!isValidUrl(config.branding.linkUrl)) {
      config.branding.linkUrl = "";
    }

    return config;
  }

  async function loadDefaultConfig() {
    const response = await fetch(DEFAULT_CONFIG_URL);
    if (!response.ok) {
      throw new Error("Không thể tải cấu hình mặc định.");
    }

    const data = await response.json();
    return normalizeConfig(data);
  }

  async function loadConfig() {
    const data = await chrome.storage.local.get(CONFIG_KEY);
    if (data[CONFIG_KEY]) {
      return normalizeConfig(data[CONFIG_KEY]);
    }
    return loadDefaultConfig();
  }

  async function saveConfig(config) {
    const normalized = normalizeConfig(config);
    await chrome.storage.local.set({ [CONFIG_KEY]: normalized });
    return normalized;
  }

  function matchRule(rule, cardInfo) {
    for (const field of ["subject", "class", "room", "teacher"]) {
      const expected = text(rule && rule[field]);
      if (!expected) return false;
      if (!norm(cardInfo[field]).includes(norm(expected))) return false;
    }
    return true;
  }

  function chooseMap(config, cardInfo) {
    const rules = Array.isArray(config.rules) ? config.rules : [];

    for (const rule of rules) {
      if (rule && rule.enabled !== false && matchRule(rule, cardInfo)) {
        return "practical";
      }
    }

    return "normal";
  }

  function extractCardInfo(cardElement) {
    const info = {
      subject: "",
      class: "",
      room: "",
      teacher: "",
      periodLine: null
    };

    const subjectElement = cardElement.querySelector("b a, b");
    if (subjectElement) {
      info.subject = text(subjectElement.textContent);
    }

    for (const paragraph of Array.from(cardElement.querySelectorAll("p"))) {
      const rowText = text(paragraph.textContent);
      const rowNorm = norm(rowText);
      if (!rowText) continue;

      if (rowNorm.includes("tiet")) {
        info.periodLine = paragraph;
        continue;
      }

      if (rowNorm.includes("phong")) {
        const index = rowText.indexOf(":");
        info.room = index >= 0 ? text(rowText.slice(index + 1)) : rowText;
        continue;
      }

      if (rowNorm.includes("gv") || rowNorm.includes("giang vien")) {
        const index = rowText.indexOf(":");
        info.teacher = index >= 0 ? text(rowText.slice(index + 1)) : rowText;
        continue;
      }

      if (!info.class) {
        info.class = rowText;
      }
    }

    return info;
  }

  function applyPeriodConversion(config) {
    if (convertLock) return;

    convertLock = true;
    try {
      const cards = Array.from(document.querySelectorAll(".content"));

      for (const card of cards) {
        const cardInfo = extractCardInfo(card);
        if (!cardInfo.periodLine) continue;

        const line = text(cardInfo.periodLine.textContent);
        if (!line || REAL_TIME_REGEX.test(line)) continue;

        const match = norm(line).match(TIET_REGEX);
        if (!match) continue;

        const startPeriod = Number(match[1]);
        const endPeriod = Number(match[2]);
        if (!Number.isFinite(startPeriod) || !Number.isFinite(endPeriod) || startPeriod > endPeriod) continue;

        const mapName = chooseMap(config, cardInfo);
        const map = config.maps && config.maps[mapName] ? config.maps[mapName] : {};

        const start = map[String(startPeriod)] && text(map[String(startPeriod)].start);
        const end = map[String(endPeriod)] && text(map[String(endPeriod)].end);

        if (!start || !end) {
          console.warn("[UNETI] Thiếu map tiết", {
            startPeriod,
            endPeriod,
            mapName,
            subject: cardInfo.subject
          });
          continue;
        }

        cardInfo.periodLine.textContent = `Giờ: ${start} - ${end}`;
      }
    } finally {
      convertLock = false;
    }
  }

  function getWeekRange() {
    const dateInput = document.querySelector("#dateNgayXemLich");
    const baseDate = parseDDMMYYYY((dateInput || {}).value) || new Date();

    baseDate.setHours(0, 0, 0, 0);

    const offsetToMonday = baseDate.getDay() === 0 ? 6 : baseDate.getDay() - 1;
    const monday = new Date(baseDate);
    monday.setDate(baseDate.getDate() - offsetToMonday);
    monday.setHours(0, 0, 0, 0);

    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);
    sunday.setHours(23, 59, 59, 999);

    return { monday, sunday };
  }

  function findMorningAfternoonRows(tableElement) {
    const rows = tableElement.tBodies && tableElement.tBodies[0]
      ? Array.from(tableElement.tBodies[0].rows)
      : [];

    let morningRow = null;
    let afternoonRow = null;

    for (const row of rows) {
      const headerText = norm(row.cells && row.cells[0] ? row.cells[0].textContent : "");

      if (!morningRow && headerText.includes("sang")) {
        morningRow = row;
      }

      if (!afternoonRow && headerText.includes("chieu")) {
        afternoonRow = row;
      }
    }

    return {
      morning: morningRow || rows[0] || null,
      afternoon: afternoonRow || rows[1] || null
    };
  }

  function isValidManualRecord(record) {
    if (!record || record.enabled === false) return false;

    if (!text(record.subject) || !text(record.class) || !text(record.room) || !text(record.teacher)) {
      return false;
    }

    if (!(record.slot === "morning" || record.slot === "afternoon")) {
      return false;
    }

    if (!(record.timeRange === SLOT_TIME.morning || record.timeRange === SLOT_TIME.afternoon)) {
      return false;
    }

    if (record.mode === "date") {
      return Boolean(parseDDMMYYYY(record.date));
    }

    const weekdayNumber = Number(record.weekday);
    return Number.isInteger(weekdayNumber) && weekdayNumber >= 2 && weekdayNumber <= 8;
  }

  function getScheduleTable() {
    const candidates = [
      "#viewLichTheoTuan table",
      "#viewLichTheoTienDo table",
      ".table-responsive table.fl-table",
      ".table-responsive table",
      "table.fl-table.table.table-bordered",
      "table.table.table-bordered"
    ];

    for (const selector of candidates) {
      const table = document.querySelector(selector);
      if (!table) continue;
      if (!table.tHead || !table.tBodies || !table.tBodies.length) continue;
      if (!table.tBodies[0].rows || !table.tBodies[0].rows.length) continue;
      return table;
    }

    return null;
  }

  function parseHeaderColumns(tableElement) {
    let headerCells = [];
    if (tableElement.tHead && tableElement.tHead.rows && tableElement.tHead.rows[0]) {
      headerCells = Array.from(tableElement.tHead.rows[0].cells);
    } else if (tableElement.tBodies && tableElement.tBodies[0] && tableElement.tBodies[0].rows[0]) {
      // Một số layout không dùng thead, hàng đầu tbody chính là header ngày/thứ.
      headerCells = Array.from(tableElement.tBodies[0].rows[0].cells);
    }

    const columns = [];
    for (let i = 1; i < headerCells.length; i += 1) {
      const cellText = text(headerCells[i].textContent);
      const cellNorm = norm(cellText);

      let weekday = null;
      let dateText = "";

      if (cellNorm.includes("chu nhat")) {
        weekday = 8;
      } else {
        const wdMatch = cellNorm.match(/thu\s*([2-7])/);
        if (wdMatch) {
          weekday = Number(wdMatch[1]);
        }
      }

      const dateMatch = cellText.match(/(\d{2}\/\d{2}\/\d{4})/);
      if (dateMatch) {
        dateText = dateMatch[1];
        if (weekday == null) {
          const dateObj = parseDDMMYYYY(dateText);
          if (dateObj) {
            weekday = jsDayToWeekday(dateObj.getDay());
          }
        }
      }

      columns.push({
        colIndex: i,
        weekday,
        dateText
      });
    }

    // Fallback: nếu không có nhãn thứ rõ ràng nhưng có 7 cột lịch thì mặc định Thứ 2 -> CN.
    if (columns.length === 7 && columns.every((c) => c.weekday == null)) {
      columns.forEach((c, idx) => {
        c.weekday = idx === 6 ? 8 : idx + 2;
      });
    }

    return columns;
  }

  function getColumnIndexForRecord(record, columns) {
    if (!columns.length) return -1;

    if (record.mode === "date") {
      const exactDate = parseDDMMYYYY(record.date);
      if (!exactDate) return -1;

      const exactText = text(record.date);
      const byDate = columns.find((c) => c.dateText === exactText);
      if (byDate) return byDate.colIndex;

      // Fallback khi header không hiển thị ngày: dùng weekday nếu ngày đó đang trong tuần hiện tại.
      const { monday, sunday } = getWeekRange();
      if (exactDate >= monday && exactDate <= sunday) {
        const wd = jsDayToWeekday(exactDate.getDay());
        const byWeekday = columns.find((c) => c.weekday === wd);
        if (byWeekday) return byWeekday.colIndex;
      }

      return -1;
    }

    const wd = Number(record.weekday);
    const byWeekday = columns.find((c) => c.weekday === wd);
    if (byWeekday) return byWeekday.colIndex;

    return -1;
  }

  function renderManualCards(config) {
    const table = getScheduleTable();
    if (!table) {
      console.warn("[UNETI] Không tìm thấy bảng lịch để chèn lịch thực hành bổ sung.");
      return;
    }

    for (const node of Array.from(table.querySelectorAll(".tdk-manual-card"))) {
      node.remove();
    }

    const rows = findMorningAfternoonRows(table);
    const columns = parseHeaderColumns(table);
    let renderedCount = 0;

    const records = Array.isArray(config.manualPracticalSchedules) ? config.manualPracticalSchedules : [];
    for (const record of records) {
      if (!isValidManualRecord(record)) continue;

      const targetRow = record.slot === "afternoon" ? rows.afternoon : rows.morning;
      if (!targetRow) continue;

      const columnIndex = getColumnIndexForRecord(record, columns);
      if (columnIndex < 1) continue;

      // Dùng row.cells để hỗ trợ cả th/td, tránh lệch chỉ số ở layout mới.
      const targetCell = targetRow.cells && targetRow.cells[columnIndex] ? targetRow.cells[columnIndex] : null;
      if (!targetCell) continue;

      const card = document.createElement("div");
      card.className = "content color-lichhoc text-start tdk-manual-card";
      card.style.backgroundColor = "#71cb35";
      card.style.borderColor = "#c9d0db";
      card.style.textAlign = "left";
      card.style.marginTop = "6px";

      card.innerHTML =
        `<b>${esc(record.subject)}</b>` +
        `<p>${esc(record.class)}</p>` +
        `<p>Giờ: ${esc(record.timeRange)}</p>` +
        `<p>Phòng: ${esc(record.room)}</p>` +
        `<p>GV: ${esc(record.teacher)}</p>` +
        "<p><i>TH bổ sung</i></p>";

      targetCell.appendChild(card);
      renderedCount += 1;
    }

    if (records.length > 0) {
      console.debug("[UNETI] Render lịch thực hành bổ sung", {
        totalRecords: records.length,
        renderedCount,
        columnsDetected: columns.length
      });
    }
  }

  function ensureStyle() {
    if (document.getElementById("tdk-style")) return;

    const style = document.createElement("style");
    style.id = "tdk-style";
    style.textContent =
      "#tdk-modal-wrap{position:fixed;inset:0;background:rgba(15,23,42,.45);display:none;z-index:2147483640}" +
      "#tdk-modal{width:min(980px,95vw);max-height:90vh;overflow:auto;margin:4vh auto;background:#fff;border-radius:10px;padding:14px}" +
      ".tdk-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}" +
      ".tdk-grid label{display:flex;flex-direction:column;font-size:12px}" +
      ".tdk-grid input,.tdk-grid select{height:34px;border:1px solid #cbd5e1;border-radius:8px;padding:0 8px}" +
      ".tdk-full{grid-column:1/-1}" +
      ".tdk-act{display:flex;gap:8px;align-items:center;margin-top:8px}" +
      ".tdk-act button,.tdk-btn{border:1px solid #cbd5e1;border-radius:8px;background:#fff;padding:6px 10px;cursor:pointer}" +
      ".tdk-primary{background:#2563eb;border-color:#2563eb;color:#fff}" +
      "#tdk-list{width:100%;border-collapse:collapse;font-size:12px;margin-top:10px}" +
      "#tdk-list th,#tdk-list td{border:1px solid #e2e8f0;padding:6px}" +
      "#tdk-brand-footer{margin-top:10px;font-size:12px;display:flex;justify-content:flex-end;gap:6px;align-items:center}" +
      "#tdk-brand-footer img{width:18px;height:18px;object-fit:contain}";

    document.head.appendChild(style);
  }

  function modalElements() {
    return {
      wrap: document.getElementById("tdk-modal-wrap"),
      form: document.getElementById("tdk-form"),
      status: document.getElementById("tdk-status"),
      editId: document.getElementById("tdk-edit-id"),
      subject: document.getElementById("tdk-subject"),
      class: document.getElementById("tdk-class"),
      room: document.getElementById("tdk-room"),
      teacher: document.getElementById("tdk-teacher"),
      mode: document.getElementById("tdk-mode"),
      weekdayWrap: document.getElementById("tdk-weekday-wrap"),
      weekday: document.getElementById("tdk-weekday"),
      dateWrap: document.getElementById("tdk-date-wrap"),
      date: document.getElementById("tdk-date"),
      slot: document.getElementById("tdk-slot"),
      enabled: document.getElementById("tdk-enabled"),
      tbody: document.getElementById("tdk-tbody")
    };
  }

  function setModalStatus(message, type) {
    const modal = modalElements();
    if (!modal.status) return;

    modal.status.textContent = message || "";
    modal.status.style.color =
      type === "error"
        ? "#b91c1c"
        : type === "success"
          ? "#065f46"
          : "#334155";
  }

  function toggleModalMode() {
    const modal = modalElements();
    if (!modal.mode) return;

    const isDateMode = modal.mode.value === "date";
    modal.weekdayWrap.style.display = isDateMode ? "none" : "";
    modal.dateWrap.style.display = isDateMode ? "" : "none";
  }

  function resetModalForm() {
    const modal = modalElements();
    if (!modal.form) return;

    modal.editId.value = "";
    modal.subject.value = "";
    modal.class.value = "";
    modal.room.value = "";
    modal.teacher.value = "";
    modal.mode.value = "weekly";
    modal.weekday.value = "2";
    modal.date.value = "";
    modal.slot.value = "morning";
    modal.enabled.checked = true;

    toggleModalMode();
    setModalStatus("", "");
  }

  function renderModalList(config) {
    const modal = modalElements();
    if (!modal.tbody) return;

    const records = Array.isArray(config.manualPracticalSchedules) ? config.manualPracticalSchedules : [];

    if (records.length === 0) {
      modal.tbody.innerHTML =
        '<tr><td colspan="8" style="text-align:center;color:#64748b">Chưa có lịch thực hành bổ sung.</td></tr>';
      return;
    }

    modal.tbody.innerHTML = records
      .map((record, index) => {
        const modeText =
          record.mode === "date"
            ? `Theo ngày (${record.date || "--"})`
            : `Hàng tuần (${WEEKDAY_LABEL[record.weekday] || "--"})`;

        return (
          `<tr>` +
            `<td>${index + 1}</td>` +
            `<td>${esc(record.subject)}</td>` +
            `<td>${esc(record.class)}</td>` +
            `<td>${esc(record.room)}</td>` +
            `<td>${esc(record.teacher)}</td>` +
            `<td>${esc(modeText)}</td>` +
            `<td>${esc(record.timeRange)}</td>` +
            `<td>` +
              `<label><input type="checkbox" class="tdk-enabled-toggle" data-id="${esc(record.id)}" ${record.enabled ? "checked" : ""}> Bật</label> ` +
              `<button class="tdk-btn tdk-edit" data-id="${esc(record.id)}">Sửa</button> ` +
              `<button class="tdk-btn tdk-del" data-id="${esc(record.id)}">Xóa</button>` +
            `</td>` +
          `</tr>`
        );
      })
      .join("");
  }

  async function persistModalConfig() {
    modalConfig = await saveConfig(modalConfig);
    renderModalList(modalConfig);
    scheduleRun(50);
  }

  function fillModalFormForEdit(record) {
    const modal = modalElements();

    modal.editId.value = record.id;
    modal.subject.value = record.subject;
    modal.class.value = record.class;
    modal.room.value = record.room;
    modal.teacher.value = record.teacher;
    modal.mode.value = record.mode;
    modal.weekday.value = String(record.weekday || 2);
    modal.date.value = toDateInput(record.date);
    modal.slot.value = record.slot;
    modal.enabled.checked = record.enabled !== false;

    toggleModalMode();
    setModalStatus("Đang chỉnh sửa bản ghi.", "success");
  }

  function buildRecordFromModalForm() {
    const modal = modalElements();

    const mode = modal.mode.value === "date" ? "date" : "weekly";
    const slot = modal.slot.value === "afternoon" ? "afternoon" : "morning";

    return {
      id: text(modal.editId.value) || `mp-${Date.now()}-${Math.floor(Math.random() * 100000)}`,
      enabled: modal.enabled.checked,
      mode,
      weekday: Number(modal.weekday.value) || 2,
      date: mode === "date" ? fromDateInput(modal.date.value) : "",
      slot,
      timeRange: slot === "afternoon" ? SLOT_TIME.afternoon : SLOT_TIME.morning,
      subject: text(modal.subject.value),
      class: text(modal.class.value),
      room: text(modal.room.value),
      teacher: text(modal.teacher.value)
    };
  }

  function validateModalRecord(record) {
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

  function ensureModal() {
    if (!isSchedulePage() || document.getElementById("tdk-modal-wrap")) return;

    ensureStyle();

    const wrap = document.createElement("div");
    wrap.id = "tdk-modal-wrap";
    wrap.innerHTML =
      `<div id="tdk-modal">` +
        `<div style="display:flex;justify-content:space-between;align-items:center">` +
          `<h3 style="margin:0">Quản lý lịch thực hành bổ sung</h3>` +
          `<button id="tdk-close" class="tdk-btn">Đóng</button>` +
        `</div>` +
        `<form id="tdk-form">` +
          `<input id="tdk-edit-id" type="hidden">` +
          `<div class="tdk-grid">` +
            `<label>Môn học<input id="tdk-subject" type="text"></label>` +
            `<label>Lớp<input id="tdk-class" type="text"></label>` +
            `<label>Phòng<input id="tdk-room" type="text"></label>` +
            `<label>Giảng viên<input id="tdk-teacher" type="text"></label>` +
            `<label>Kiểu áp dụng<select id="tdk-mode"><option value="weekly">Hàng tuần theo thứ</option><option value="date">Theo ngày cụ thể</option></select></label>` +
            `<label id="tdk-weekday-wrap">Thứ<select id="tdk-weekday"><option value="2">Thứ 2</option><option value="3">Thứ 3</option><option value="4">Thứ 4</option><option value="5">Thứ 5</option><option value="6">Thứ 6</option><option value="7">Thứ 7</option><option value="8">Chủ nhật</option></select></label>` +
            `<label id="tdk-date-wrap" style="display:none">Ngày cụ thể<input id="tdk-date" type="date"></label>` +
            `<label>Khung giờ thực hành<select id="tdk-slot"><option value="morning">06:00 - 11:30 (Sáng)</option><option value="afternoon">13:00 - 18:30 (Chiều)</option></select></label>` +
            `<label class="tdk-full" style="flex-direction:row;gap:8px;align-items:center"><input id="tdk-enabled" type="checkbox" checked>Bật bản ghi này</label>` +
          `</div>` +
          `<div class="tdk-act">` +
            `<button class="tdk-primary" type="submit">Lưu bản ghi</button>` +
            `<button id="tdk-reset" type="button">Hủy chỉnh sửa</button>` +
            `<span id="tdk-status"></span>` +
          `</div>` +
        `</form>` +
        `<table id="tdk-list">` +
          `<thead><tr><th>#</th><th>Môn</th><th>Lớp</th><th>Phòng</th><th>GV</th><th>Kiểu áp dụng</th><th>Khung giờ</th><th>Thao tác</th></tr></thead>` +
          `<tbody id="tdk-tbody"></tbody>` +
        `</table>` +
      `</div>`;

    document.body.appendChild(wrap);

    const modal = modalElements();
    modal.mode.addEventListener("change", toggleModalMode);

    document.getElementById("tdk-close").addEventListener("click", () => {
      wrap.style.display = "none";
    });

    document.getElementById("tdk-reset").addEventListener("click", resetModalForm);

    wrap.addEventListener("click", (event) => {
      if (event.target === wrap) {
        wrap.style.display = "none";
      }
    });

    modal.form.addEventListener("submit", async (event) => {
      event.preventDefault();

      if (!modalConfig) {
        modalConfig = await loadConfig();
      }

      const record = buildRecordFromModalForm();
      const error = validateModalRecord(record);
      if (error) {
        setModalStatus(error, "error");
        return;
      }

      const records = modalConfig.manualPracticalSchedules || [];
      const index = records.findIndex((item) => item.id === record.id);

      if (index >= 0) {
        records[index] = record;
      } else {
        records.push(record);
      }

      modalConfig.manualPracticalSchedules = records;
      await persistModalConfig();
      setModalStatus(index >= 0 ? "Đã cập nhật bản ghi." : "Đã thêm bản ghi.", "success");
      resetModalForm();
    });

    modal.tbody.addEventListener("click", async (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement) || !modalConfig) return;

      const editButton = target.closest(".tdk-edit");
      if (editButton) {
        const id = text(editButton.getAttribute("data-id"));
        const record = (modalConfig.manualPracticalSchedules || []).find((item) => item.id === id);
        if (record) {
          fillModalFormForEdit(record);
        }
        return;
      }

      const deleteButton = target.closest(".tdk-del");
      if (deleteButton) {
        const accepted = window.confirm("Bạn có chắc muốn xóa lịch thực hành bổ sung này?");
        if (!accepted) return;

        const id = text(deleteButton.getAttribute("data-id"));
        modalConfig.manualPracticalSchedules = (modalConfig.manualPracticalSchedules || []).filter((item) => item.id !== id);
        await persistModalConfig();

        resetModalForm();
        setModalStatus("Đã xóa bản ghi.", "success");
      }
    });

    modal.tbody.addEventListener("change", async (event) => {
      const target = event.target;
      if (!(target instanceof HTMLInputElement) || !target.classList.contains("tdk-enabled-toggle") || !modalConfig) {
        return;
      }

      const id = text(target.getAttribute("data-id"));
      const records = modalConfig.manualPracticalSchedules || [];
      const index = records.findIndex((item) => item.id === id);
      if (index < 0) return;

      records[index].enabled = target.checked;
      modalConfig.manualPracticalSchedules = records;
      await persistModalConfig();
    });
  }

  async function openModal() {
    ensureModal();

    const modal = modalElements();
    if (!modal.wrap) return;

    modalConfig = await loadConfig();
    renderModalList(modalConfig);
    resetModalForm();

    modal.wrap.style.display = "block";
  }

  function ensureManageButton() {
    if (!isSchedulePage()) return;

    const actions = document.querySelector(".portlet .actions");
    if (!actions || document.getElementById("tdk-open-btn")) return;

    const button = document.createElement("a");
    button.id = "tdk-open-btn";
    button.href = "javascript:;";
    button.className = "btn btn-action";
    button.innerHTML = '<i class="fa fa-plus" aria-hidden="true"></i> Thêm lịch TH';
    button.addEventListener("click", () => {
      void openModal();
    });

    actions.appendChild(button);
  }

  function ensureBrandFooter(config) {
    if (!isSchedulePage()) return;

    const host = document.querySelector("#viewLichTheoTuan, #viewLichTheoTienDo");
    if (!host || !host.parentElement) return;

    let footer = document.getElementById("tdk-brand-footer");
    if (!footer) {
      footer = document.createElement("div");
      footer.id = "tdk-brand-footer";
      host.parentElement.insertBefore(footer, host.nextSibling);
    }

    const branding = config.branding || {};
    const company = text(branding.companyName) || DEFAULT_COMPANY;
    const logoUrl = chrome.runtime.getURL(text(branding.logoPath) || DEFAULT_LOGO_PATH);

    const content =
      `<img src="${esc(logoUrl)}" alt="logo" onerror="this.style.display='none';">` +
      `<span>Bản quyền ${esc(company)}</span>`;

    const link = text(branding.linkUrl);
    if (isValidUrl(link) && link) {
      footer.innerHTML = `<a href="${esc(link)}" target="_blank" rel="noopener noreferrer">${content}</a>`;
    } else {
      footer.innerHTML = content;
    }
  }

  function isExtensionManagedElement(node) {
    if (!(node instanceof Element)) return false;

    if (
      node.id === "tdk-modal-wrap" ||
      node.id === "tdk-brand-footer" ||
      node.id === "tdk-open-btn" ||
      node.classList.contains("tdk-manual-card")
    ) {
      return true;
    }

    return Boolean(node.closest("#tdk-modal-wrap, #tdk-brand-footer, #tdk-open-btn, .tdk-manual-card"));
  }

  function isMutationFromExtension(mutation) {
    const changedNodes = [...mutation.addedNodes, ...mutation.removedNodes];
    if (!changedNodes.length) return false;

    return changedNodes.every((node) => {
      if (node instanceof Element) {
        return isExtensionManagedElement(node);
      }

      const parent = node.parentElement || (mutation.target instanceof Element ? mutation.target : null);
      return parent ? isExtensionManagedElement(parent) : false;
    });
  }

  function isScheduleControlValueChange(mutation) {
    if (mutation.type !== "attributes" || mutation.attributeName !== "value") return false;

    const target = mutation.target;
    if (!(target instanceof Element)) return false;
    if (isExtensionManagedElement(target)) return false;

    return target.matches("#dateNgayXemLich, #firstDateOffWeek, #firstDatePrevOffWeek, #firstDateNextOffWeek");
  }

  function bindScheduleRefreshHooks() {
    if (!isSchedulePage() || refreshHooksBound) return;
    refreshHooksBound = true;

    document.addEventListener(
      "click",
      (event) => {
        const target = event.target;
        if (!(target instanceof Element)) return;

        if (target.closest("#btn_Tiep, #btn_TroVe, #btn_HienTai")) {
          scheduleRun(360);
        }
      },
      true
    );

    document.addEventListener(
      "change",
      (event) => {
        const target = event.target;
        if (!(target instanceof Element)) return;

        if (target.matches("#dateNgayXemLich, input[name='rdoLoaiLich']")) {
          scheduleRun(260);
        }
      },
      true
    );
  }

  function bindAjaxRefreshHook() {
    if (!isSchedulePage() || jqueryAjaxHookBound) return;

    const jq = window.jQuery || window.$;
    if (typeof jq !== "function") return;

    const docRef = jq(document);
    if (!docRef || typeof docRef.ajaxComplete !== "function") return;

    jqueryAjaxHookBound = true;
    docRef.ajaxComplete((_event, _xhr, settings) => {
      const url = String((settings && settings.url) || "").toLowerCase();
      if (!url) return;

      if (
        url.includes("/sinhvien/getdanhsachlichtheotuan") ||
        url.includes("/sinhvien/getdanhsachlichtheotiendo") ||
        url.includes("/sinhvien/getdscamthi")
      ) {
        scheduleRun(40);
      }
    });
  }

  async function run() {
    try {
      const config = await loadConfig();

      applyPeriodConversion(config);

      if (isSchedulePage()) {
        bindScheduleRefreshHooks();
        bindAjaxRefreshHook();
        ensureManageButton();
        ensureModal();
        ensureBrandFooter(config);
        renderManualCards(config);
      }
    } catch (error) {
      console.error("[UNETI] Không thể cập nhật extension.", error);
    }
  }

  function scheduleRun(delay = 150) {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
      void run();
    }, delay);
  }

  function startObserver() {
    if (observerStarted) return;

    const root = document.body || document.documentElement;
    if (!root) return;
    observerStarted = true;

    const observer = new MutationObserver((mutations) => {
      for (const mutation of mutations) {
        if (isScheduleControlValueChange(mutation)) {
          scheduleRun(90);
          return;
        }

        if (mutation.type !== "childList") continue;
        if (!mutation.addedNodes.length && !mutation.removedNodes.length) continue;
        if (isMutationFromExtension(mutation)) continue;

        if (mutation.addedNodes.length > 0 || mutation.removedNodes.length > 0) {
          scheduleRun(120);
          return;
        }
      }
    });

    observer.observe(root, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["value"]
    });
  }

  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === "local" && changes[CONFIG_KEY]) {
      scheduleRun(60);
    }
  });

  if (document.readyState === "loading") {
    document.addEventListener(
      "DOMContentLoaded",
      () => {
        startObserver();
        scheduleRun(0);
      },
      { once: true }
    );
  } else {
    startObserver();
    scheduleRun(0);
  }
})();
