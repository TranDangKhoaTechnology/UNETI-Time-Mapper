(function () {
  "use strict";

  const BOOTSTRAP_FLAG = "__tdkUnetiTimeMapperBootstrapV1";
  if (globalThis[BOOTSTRAP_FLAG]) {
    return;
  }
  globalThis[BOOTSTRAP_FLAG] = true;

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

  const EXPORT_FORMATS = ["xlsx", "csv", "json", "html", "png", "pdf"];
  const EXPORT_RANGE_MODES = ["week", "month"];
  const EXPORT_MONTH_LAYOUTS = ["timetable", "list"];
  const SESSION_SORT_HINT = {
    sang: 1,
    morning: 1,
    chieu: 2,
    afternoon: 2,
    toi: 3,
    evening: 3
  };
  const CATEGORY_SORT_HINT = {
    study: 1,
    exam: 2,
    unknown: 3
  };

  let debounceTimer = null;
  let convertLock = false;
  let modalConfig = null;
  let exportInProgress = false;
  let refreshHooksBound = false;
  let jqueryAjaxHookBound = false;
  let observerStarted = false;
  let gradeEditorState = null;
  let gradeInlineHandlersBound = false;

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

  function toDDMMYYYY(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
    const day = String(date.getDate()).padStart(2, "0");
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const year = String(date.getFullYear());
    return `${day}/${month}/${year}`;
  }

  function toMonthInputValue(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
    const month = String(date.getMonth() + 1).padStart(2, "0");
    return `${date.getFullYear()}-${month}`;
  }

  function parseMonthInput(monthValue) {
    const match = text(monthValue).match(/^(\d{4})-(\d{2})$/);
    if (!match) return null;

    const year = Number(match[1]);
    const month = Number(match[2]);
    if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) {
      return null;
    }

    const firstDay = new Date(year, month - 1, 1);
    firstDay.setHours(0, 0, 0, 0);
    return firstDay;
  }

  function getMondayOfDate(date) {
    const base = new Date(date);
    base.setHours(0, 0, 0, 0);

    const offsetToMonday = base.getDay() === 0 ? 6 : base.getDay() - 1;
    base.setDate(base.getDate() - offsetToMonday);
    base.setHours(0, 0, 0, 0);
    return base;
  }

  function getSundayOfMonday(monday) {
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);
    sunday.setHours(23, 59, 59, 999);
    return sunday;
  }

  function getMonthRange(monthValue) {
    const firstDay = parseMonthInput(monthValue);
    if (!firstDay) return null;

    const lastDay = new Date(firstDay.getFullYear(), firstDay.getMonth() + 1, 0);
    lastDay.setHours(23, 59, 59, 999);
    return { firstDay, lastDay };
  }

  function getWeekMondaysForMonth(monthValue) {
    const monthRange = getMonthRange(monthValue);
    if (!monthRange) return [];

    const startMonday = getMondayOfDate(monthRange.firstDay);
    const endSunday = getSundayOfMonday(getMondayOfDate(monthRange.lastDay));
    const mondays = [];

    const pointer = new Date(startMonday);
    while (pointer <= endSunday) {
      mondays.push(new Date(pointer));
      pointer.setDate(pointer.getDate() + 7);
    }

    return mondays;
  }

  function getCurrentScheduleTypeValue() {
    const checked = document.querySelector("input[name='rdoLoaiLich']:checked");
    const value = Number(checked && checked.value);
    if (value === 1 || value === 2) return value;
    return 0;
  }

  function scheduleTypeLabel(scheduleType) {
    if (scheduleType === 1) return "Lịch học";
    if (scheduleType === 2) return "Lịch thi";
    return "Tất cả";
  }

  function normalizeTimeToken(token) {
    const clean = String(token || "").trim().toLowerCase().replace("h", ":");
    const match = clean.match(/^(\d{1,2}):(\d{2})$/);
    if (!match) return "";

    const hour = Number(match[1]);
    const minute = Number(match[2]);
    if (!Number.isInteger(hour) || hour < 0 || hour > 23 || !Number.isInteger(minute) || minute < 0 || minute > 59) {
      return "";
    }

    return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
  }

  function parseTimeRangeText(value) {
    const raw = text(value);
    if (!raw) {
      return {
        timeRange: "",
        startTime: "",
        endTime: ""
      };
    }

    const match = raw.match(/(\d{1,2}(?:h|:)\d{2})\s*[-–]\s*(\d{1,2}(?:h|:)\d{2})/i);
    if (!match) {
      return {
        timeRange: raw,
        startTime: "",
        endTime: ""
      };
    }

    const startTime = normalizeTimeToken(match[1]);
    const endTime = normalizeTimeToken(match[2]);
    const timeRange = startTime && endTime ? `${startTime}-${endTime}` : raw;

    return { timeRange, startTime, endTime };
  }

  function sessionSortValue(session) {
    const n = norm(session);
    for (const key of Object.keys(SESSION_SORT_HINT)) {
      if (n.includes(key)) {
        return SESSION_SORT_HINT[key];
      }
    }
    return 99;
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

  function isGradePage() {
    return location.pathname.toLowerCase().includes("/ket-qua-hoc-tap.html");
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

  function getScheduleTable(root = document) {
    const candidates = [
      "#viewLichTheoTuan table",
      "#viewLichTheoTienDo table",
      ".table-responsive table.fl-table",
      ".table-responsive table",
      "table.fl-table.table.table-bordered",
      "table.table.table-bordered"
    ];

    for (const selector of candidates) {
      const table = root.querySelector(selector);
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
        dateText,
        headerText: cellText
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

  function inferCategoryFromClassName(className, fallbackScheduleType) {
    const classNorm = norm(className);
    if (classNorm.includes("lichthi")) return "exam";
    if (classNorm.includes("lichhoc")) return "study";

    if (fallbackScheduleType === 2) return "exam";
    if (fallbackScheduleType === 1) return "study";
    return "unknown";
  }

  function parseCardData(cardElement) {
    const subject = text((cardElement.querySelector("b a, b") || {}).textContent);
    let className = "";
    let room = "";
    let teacher = "";
    let timeRange = "";
    let startTime = "";
    let endTime = "";

    for (const paragraph of Array.from(cardElement.querySelectorAll("p"))) {
      const line = text(paragraph.textContent);
      const lineNorm = norm(line);
      if (!line) continue;

      if (lineNorm.includes("gio") || lineNorm.includes("tiet")) {
        const parsedTime = parseTimeRangeText(line);
        timeRange = parsedTime.timeRange || timeRange;
        startTime = parsedTime.startTime || startTime;
        endTime = parsedTime.endTime || endTime;
        continue;
      }

      if (lineNorm.includes("phong")) {
        const idx = line.indexOf(":");
        room = idx >= 0 ? text(line.slice(idx + 1)) : line;
        continue;
      }

      if (lineNorm.includes("gv") || lineNorm.includes("giang vien")) {
        const idx = line.indexOf(":");
        teacher = idx >= 0 ? text(line.slice(idx + 1)) : line;
        continue;
      }

      if (!className) {
        className = line;
      }
    }

    if (!timeRange) {
      const parsedFromCard = parseTimeRangeText(text(cardElement.textContent));
      timeRange = parsedFromCard.timeRange;
      startTime = parsedFromCard.startTime;
      endTime = parsedFromCard.endTime;
    }

    return {
      subject,
      className,
      room,
      teacher,
      timeRange,
      startTime,
      endTime
    };
  }

  function parseScheduleEventsFromTable(tableElement, options = {}) {
    if (!tableElement) return [];

    const columns = parseHeaderColumns(tableElement);
    const bodyRows = tableElement.tBodies && tableElement.tBodies[0] ? Array.from(tableElement.tBodies[0].rows) : [];
    const fallbackScheduleType = options.scheduleType;
    const weekStart = options.weekStart instanceof Date ? new Date(options.weekStart) : null;
    const events = [];

    for (const row of bodyRows) {
      const session = text(row.cells && row.cells[0] ? row.cells[0].textContent : "") || "Không rõ ca";

      for (const column of columns) {
        const cell = row.cells && row.cells[column.colIndex] ? row.cells[column.colIndex] : null;
        if (!cell) continue;

        const cards = Array.from(cell.querySelectorAll(".content"));
        if (!cards.length) continue;

        for (const card of cards) {
          const cardData = parseCardData(card);
          if (!cardData.subject && !cardData.className && !cardData.room && !cardData.teacher) continue;

          let dateText = text(column.dateText);
          let weekday = column.weekday;

          if (!dateText && weekStart && Number.isInteger(weekday) && weekday >= 2 && weekday <= 8) {
            const date = new Date(weekStart);
            const offset = weekday === 8 ? 6 : weekday - 2;
            date.setDate(weekStart.getDate() + offset);
            dateText = toDDMMYYYY(date);
          }

          if (!weekday && dateText) {
            const dateObj = parseDDMMYYYY(dateText);
            if (dateObj) {
              weekday = jsDayToWeekday(dateObj.getDay());
            }
          }

          const source =
            card.classList.contains("tdk-manual-card") || norm(card.textContent).includes("th bo sung")
              ? "manual"
              : "system";

          events.push({
            date: dateText,
            weekday: Number.isInteger(weekday) ? weekday : null,
            session,
            timeRange: cardData.timeRange,
            startTime: cardData.startTime,
            endTime: cardData.endTime,
            subject: cardData.subject,
            class: cardData.className,
            room: cardData.room,
            teacher: cardData.teacher,
            source,
            category: inferCategoryFromClassName(card.className, fallbackScheduleType)
          });
        }
      }
    }

    return events;
  }

  function eventDedupKey(event) {
    return [
      text(event.date),
      norm(event.session),
      text(event.timeRange),
      norm(event.subject),
      norm(event.class),
      norm(event.room),
      norm(event.teacher),
      event.source || "system"
    ].join("|");
  }

  function dedupeScheduleEvents(events) {
    const byKey = new Map();
    for (const event of events) {
      if (!event || !text(event.subject)) continue;
      byKey.set(eventDedupKey(event), event);
    }
    return Array.from(byKey.values());
  }

  function parseEventDate(event) {
    return parseDDMMYYYY(event && event.date);
  }

  function compareScheduleEvents(a, b) {
    const aDate = parseEventDate(a);
    const bDate = parseEventDate(b);
    const aTime = aDate ? aDate.getTime() : Number.MAX_SAFE_INTEGER;
    const bTime = bDate ? bDate.getTime() : Number.MAX_SAFE_INTEGER;
    if (aTime !== bTime) return aTime - bTime;

    const sessionDiff = sessionSortValue(a.session) - sessionSortValue(b.session);
    if (sessionDiff !== 0) return sessionDiff;

    const startDiff = text(a.startTime).localeCompare(text(b.startTime));
    if (startDiff !== 0) return startDiff;

    const categoryDiff = (CATEGORY_SORT_HINT[a.category] || 9) - (CATEGORY_SORT_HINT[b.category] || 9);
    if (categoryDiff !== 0) return categoryDiff;

    return text(a.subject).localeCompare(text(b.subject));
  }

  function sortScheduleEvents(events) {
    return [...events].sort(compareScheduleEvents);
  }

  function isDateBetween(dateObj, startDate, endDate) {
    if (!(dateObj instanceof Date) || Number.isNaN(dateObj.getTime())) return false;
    return dateObj >= startDate && dateObj <= endDate;
  }

  function sessionFromSlot(slot) {
    return slot === "afternoon" ? "Chiều" : "Sáng";
  }

  function buildManualEventsForRange(records, rangeStart, rangeEnd) {
    const result = [];
    const sourceRecords = Array.isArray(records) ? records : [];

    for (const record of sourceRecords) {
      if (!isValidManualRecord(record)) continue;

      if (record.mode === "date") {
        const dateObj = parseDDMMYYYY(record.date);
        if (!isDateBetween(dateObj, rangeStart, rangeEnd)) continue;
        const parsedTime = parseTimeRangeText(record.timeRange);

        result.push({
          date: toDDMMYYYY(dateObj),
          weekday: jsDayToWeekday(dateObj.getDay()),
          session: sessionFromSlot(record.slot),
          timeRange: parsedTime.timeRange || record.timeRange,
          startTime: parsedTime.startTime,
          endTime: parsedTime.endTime,
          subject: text(record.subject),
          class: text(record.class),
          room: text(record.room),
          teacher: text(record.teacher),
          source: "manual",
          category: "study"
        });
        continue;
      }

      const expectedWeekday = Number(record.weekday);
      const parsedTime = parseTimeRangeText(record.timeRange);
      for (let date = new Date(rangeStart); date <= rangeEnd; date.setDate(date.getDate() + 1)) {
        const dateCopy = new Date(date);
        dateCopy.setHours(0, 0, 0, 0);
        const weekday = jsDayToWeekday(dateCopy.getDay());
        if (weekday !== expectedWeekday) continue;

        result.push({
          date: toDDMMYYYY(dateCopy),
          weekday,
          session: sessionFromSlot(record.slot),
          timeRange: parsedTime.timeRange || record.timeRange,
          startTime: parsedTime.startTime,
          endTime: parsedTime.endTime,
          subject: text(record.subject),
          class: text(record.class),
          room: text(record.room),
          teacher: text(record.teacher),
          source: "manual",
          category: "study"
        });
      }
    }

    return result;
  }

  function dateLabelWithWeekday(dateObj) {
    const weekday = jsDayToWeekday(dateObj.getDay());
    const weekdayText = WEEKDAY_LABEL[weekday] || "Không rõ";
    return `${weekdayText} (${toDDMMYYYY(dateObj)})`;
  }

  function buildTimetableWeeks(events, weekMondays) {
    const sourceEvents = Array.isArray(events) ? events : [];
    const mondays = Array.isArray(weekMondays) ? weekMondays : [];
    const weeks = [];

    for (const monday of mondays) {
      const sunday = getSundayOfMonday(monday);
      const weekEvents = sourceEvents.filter((event) => {
        const dateObj = parseEventDate(event);
        return isDateBetween(dateObj, monday, sunday);
      });

      const sessions = Array.from(new Set(weekEvents.map((event) => text(event.session)).filter(Boolean)));
      sessions.sort((a, b) => sessionSortValue(a) - sessionSortValue(b) || a.localeCompare(b));

      const dayColumns = [];
      for (let i = 0; i < 7; i += 1) {
        const dateObj = new Date(monday);
        dateObj.setDate(monday.getDate() + i);
        dayColumns.push({
          date: toDDMMYYYY(dateObj),
          label: dateLabelWithWeekday(dateObj)
        });
      }

      const byDayAndSession = new Map();
      for (const event of weekEvents) {
        const key = `${event.date}__${text(event.session)}`;
        if (!byDayAndSession.has(key)) {
          byDayAndSession.set(key, []);
        }
        byDayAndSession.get(key).push(event);
      }

      for (const [key, list] of byDayAndSession.entries()) {
        byDayAndSession.set(key, sortScheduleEvents(list));
      }

      weeks.push({
        monday: toDDMMYYYY(monday),
        sunday: toDDMMYYYY(sunday),
        title: `Tuần ${toDDMMYYYY(monday)} - ${toDDMMYYYY(sunday)}`,
        sessions,
        dayColumns,
        byDayAndSession
      });
    }

    return weeks;
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
      ".tdk-danger{background:#b91c1c;border-color:#b91c1c;color:#fff}" +
      "#tdk-list{width:100%;border-collapse:collapse;font-size:12px;margin-top:10px}" +
      "#tdk-list th,#tdk-list td{border:1px solid #e2e8f0;padding:6px}" +
      "#tdk-brand-footer{margin-top:10px;font-size:12px;display:flex;justify-content:flex-end;gap:6px;align-items:center}" +
      "#tdk-brand-footer img{width:18px;height:18px;object-fit:contain}" +
      "#tdk-export-modal-wrap{position:fixed;inset:0;background:rgba(15,23,42,.45);display:none;z-index:2147483641}" +
      "#tdk-export-modal{width:min(980px,95vw);max-height:90vh;overflow:auto;margin:4vh auto;background:#fff;border-radius:10px;padding:14px}" +
      ".tdk-export-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}" +
      ".tdk-export-grid label{display:flex;flex-direction:column;font-size:12px}" +
      ".tdk-export-grid input,.tdk-export-grid select{height:34px;border:1px solid #cbd5e1;border-radius:8px;padding:0 8px}" +
      ".tdk-export-status{font-size:12px;min-height:18px}" +
      ".tdk-export-note{font-size:12px;color:#475569;margin:0}" +
      ".tdk-render-root{position:fixed;left:-99999px;top:0;z-index:-1;background:#fff;padding:16px;width:1400px}" +
      ".tdk-export-sheet{font-family:Arial,sans-serif;color:#0f172a}" +
      ".tdk-export-sheet h2{margin:0 0 8px;font-size:18px}" +
      ".tdk-export-sheet h3{margin:14px 0 8px;font-size:14px}" +
      ".tdk-export-meta{font-size:12px;color:#334155;margin-bottom:8px}" +
      ".tdk-export-table{width:100%;border-collapse:collapse;font-size:12px}" +
      ".tdk-export-table th,.tdk-export-table td{border:1px solid #cbd5e1;padding:6px;vertical-align:top}" +
      ".tdk-export-table th{background:#e2e8f0;text-align:center}" +
      ".tdk-export-event{border-radius:6px;padding:4px 6px;margin:4px 0;background:#f8fafc;border:1px solid #e2e8f0}" +
      ".tdk-export-event.exam{background:#eef7ff;border-color:#bfdbfe}" +
      ".tdk-export-event.manual{background:#f0fdf4;border-color:#86efac}" +
      ".tdk-export-cell-empty{color:#94a3b8;text-align:center}";

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

  function exportModalElements() {
    return {
      wrap: document.getElementById("tdk-export-modal-wrap"),
      form: document.getElementById("tdk-export-form"),
      close: document.getElementById("tdk-export-close"),
      format: document.getElementById("tdk-export-format"),
      rangeMode: document.getElementById("tdk-export-range"),
      monthWrap: document.getElementById("tdk-export-month-wrap"),
      month: document.getElementById("tdk-export-month"),
      layoutWrap: document.getElementById("tdk-export-layout-wrap"),
      monthLayout: document.getElementById("tdk-export-layout"),
      scheduleType: document.getElementById("tdk-export-schedule-type"),
      status: document.getElementById("tdk-export-status"),
      submit: document.getElementById("tdk-export-submit")
    };
  }

  function setExportStatus(message, type) {
    const modal = exportModalElements();
    if (!modal.status) return;

    modal.status.textContent = message || "";
    modal.status.style.color =
      type === "error"
        ? "#b91c1c"
        : type === "success"
          ? "#065f46"
          : type === "info"
            ? "#1d4ed8"
            : "#334155";
  }

  function getExportContext() {
    const dateInput = document.querySelector("#dateNgayXemLich");
    const selectedDate = parseDDMMYYYY((dateInput || {}).value) || new Date();
    const currentMonth = toMonthInputValue(selectedDate);
    const scheduleType = getCurrentScheduleTypeValue();
    const table = getScheduleTable();

    return {
      isSchedulePage: isSchedulePage(),
      hasTable: Boolean(table),
      selectedDate: toDDMMYYYY(selectedDate),
      currentMonth,
      scheduleType,
      scheduleTypeLabel: scheduleTypeLabel(scheduleType),
      pageUrl: location.href
    };
  }

  function normalizeExportOptions(rawOptions, context) {
    const source = rawOptions && typeof rawOptions === "object" ? rawOptions : {};
    const fallbackContext = context || getExportContext();

    const format = EXPORT_FORMATS.includes(text(source.format).toLowerCase())
      ? text(source.format).toLowerCase()
      : "xlsx";

    const rangeMode = EXPORT_RANGE_MODES.includes(text(source.rangeMode).toLowerCase())
      ? text(source.rangeMode).toLowerCase()
      : "week";

    const month = parseMonthInput(source.month) ? text(source.month) : fallbackContext.currentMonth;

    const monthLayout = EXPORT_MONTH_LAYOUTS.includes(text(source.monthLayout).toLowerCase())
      ? text(source.monthLayout).toLowerCase()
      : "timetable";

    return {
      format,
      rangeMode,
      month,
      monthLayout,
      scheduleTypeMode: "currentRadio"
    };
  }

  function syncExportModalByRange() {
    const modal = exportModalElements();
    if (!modal.rangeMode || !modal.monthWrap || !modal.layoutWrap) return;

    const isMonth = modal.rangeMode.value === "month";
    modal.monthWrap.style.display = isMonth ? "" : "none";
    modal.layoutWrap.style.display = isMonth ? "" : "none";
  }

  function fillExportModalDefaults(context) {
    const modal = exportModalElements();
    if (!modal.wrap) return;

    const currentContext = context || getExportContext();
    if (modal.month) {
      modal.month.value = currentContext.currentMonth;
    }
    if (modal.scheduleType) {
      modal.scheduleType.value = currentContext.scheduleTypeLabel;
    }
    if (modal.format) modal.format.value = modal.format.value || "xlsx";
    if (modal.rangeMode) modal.rangeMode.value = modal.rangeMode.value || "week";
    if (modal.monthLayout) modal.monthLayout.value = modal.monthLayout.value || "timetable";

    syncExportModalByRange();
    setExportStatus("", "");
  }

  function disableExportSubmit(disabled) {
    const modal = exportModalElements();
    if (!modal.submit) return;
    modal.submit.disabled = disabled;
    modal.submit.style.opacity = disabled ? "0.7" : "1";
  }

  function assertExportLibraries(format) {
    if (format === "xlsx" && !(window.ExcelJS && window.ExcelJS.Workbook)) {
      throw new Error("Thiếu thư viện ExcelJS để xuất XLSX.");
    }

    if ((format === "png" || format === "pdf") && typeof window.html2canvas !== "function") {
      throw new Error("Thiếu thư viện html2canvas để xuất ảnh/PDF.");
    }

    if (format === "pdf" && !(window.jspdf && window.jspdf.jsPDF)) {
      throw new Error("Thiếu thư viện jsPDF để xuất PDF.");
    }
  }

  function ensureExportModal() {
    if (!isSchedulePage() || document.getElementById("tdk-export-modal-wrap")) return;

    ensureStyle();

    const wrap = document.createElement("div");
    wrap.id = "tdk-export-modal-wrap";
    wrap.innerHTML =
      `<div id="tdk-export-modal">` +
        `<div style="display:flex;justify-content:space-between;align-items:center">` +
          `<h3 style="margin:0">Xuất lịch học</h3>` +
          `<button id="tdk-export-close" class="tdk-btn">Đóng</button>` +
        `</div>` +
        `<form id="tdk-export-form" style="margin-top:10px">` +
          `<div class="tdk-export-grid">` +
            `<label>Định dạng` +
              `<select id="tdk-export-format">` +
                `<option value="xlsx">XLSX</option>` +
                `<option value="csv">CSV</option>` +
                `<option value="json">JSON</option>` +
                `<option value="html">HTML</option>` +
                `<option value="png">PNG</option>` +
                `<option value="pdf">PDF</option>` +
              `</select>` +
            `</label>` +
            `<label>Phạm vi` +
              `<select id="tdk-export-range">` +
                `<option value="week">Theo tuần đang xem</option>` +
                `<option value="month">Theo tháng</option>` +
              `</select>` +
            `</label>` +
            `<label id="tdk-export-month-wrap" style="display:none">Tháng` +
              `<input id="tdk-export-month" type="month">` +
            `</label>` +
            `<label id="tdk-export-layout-wrap" style="display:none">Bố cục tháng` +
              `<select id="tdk-export-layout">` +
                `<option value="timetable">Thời khóa biểu</option>` +
                `<option value="list">Danh sách</option>` +
              `</select>` +
            `</label>` +
            `<label class="tdk-full">Loại lịch hiện tại` +
              `<input id="tdk-export-schedule-type" type="text" readonly>` +
            `</label>` +
          `</div>` +
          `<div class="tdk-act">` +
            `<button id="tdk-export-submit" class="tdk-primary" type="submit">Xuất file</button>` +
            `<span id="tdk-export-status" class="tdk-export-status"></span>` +
          `</div>` +
          `<p class="tdk-export-note">Loại lịch xuất luôn theo radio hiện tại trên trang (Tất cả/Lịch học/Lịch thi).</p>` +
        `</form>` +
      `</div>`;

    document.body.appendChild(wrap);

    const modal = exportModalElements();
    modal.close.addEventListener("click", () => {
      wrap.style.display = "none";
    });

    modal.rangeMode.addEventListener("change", syncExportModalByRange);
    wrap.addEventListener("click", (event) => {
      if (event.target === wrap) {
        wrap.style.display = "none";
      }
    });

    modal.form.addEventListener("submit", async (event) => {
      event.preventDefault();

      try {
        const context = getExportContext();
        if (!context.isSchedulePage || !context.hasTable) {
          throw new Error("Không tìm thấy bảng lịch để xuất.");
        }

        const options = normalizeExportOptions(
          {
            format: modal.format.value,
            rangeMode: modal.rangeMode.value,
            month: modal.month.value,
            monthLayout: modal.monthLayout.value
          },
          context
        );

        await executeScheduleExport(options, (message, type) => {
          setExportStatus(message, type || "info");
        });

        setExportStatus("Xuất lịch thành công.", "success");
      } catch (error) {
        setExportStatus(error.message || "Xuất lịch thất bại.", "error");
      }
    });
  }

  async function openExportModal() {
    ensureExportModal();
    const modal = exportModalElements();
    if (!modal.wrap) return;

    fillExportModalDefaults(getExportContext());
    modal.wrap.style.display = "block";
  }

  function ensureExportButton() {
    if (!isSchedulePage()) return;

    const actions = document.querySelector(".portlet .actions");
    if (!actions || document.getElementById("tdk-export-open-btn")) return;

    const button = document.createElement("a");
    button.id = "tdk-export-open-btn";
    button.href = "javascript:;";
    button.className = "btn btn-action";
    button.innerHTML = '<i class="fa fa-download" aria-hidden="true"></i> Xuất lịch';
    button.addEventListener("click", () => {
      void openExportModal();
    });

    actions.appendChild(button);
  }

  function buildExportStamp() {
    return new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
  }

  function toDateKey(ddmmyyyy) {
    const dateObj = parseDDMMYYYY(ddmmyyyy);
    if (!dateObj) return "unknown";
    const year = String(dateObj.getFullYear());
    const month = String(dateObj.getMonth() + 1).padStart(2, "0");
    const day = String(dateObj.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function sourceLabel(source) {
    return source === "manual" ? "TH bổ sung" : "Hệ thống";
  }

  function categoryLabel(category) {
    if (category === "exam") return "Lịch thi";
    if (category === "study") return "Lịch học";
    return "Không rõ";
  }

  function downloadBlob(fileName, blob) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  function buildExportFileBaseName(model) {
    const options = model.options;
    const stamp = buildExportStamp();
    const rangeToken =
      options.rangeMode === "month"
        ? options.month.replace("-", "")
        : toDateKey(model.meta.rangeStart);
    return `uneti-time-mapper-schedule-${options.rangeMode}-${rangeToken}-${stamp}`;
  }

  function escapeCsvCell(value) {
    const raw = String(value == null ? "" : value);
    if (raw.includes("\"") || raw.includes(",") || raw.includes("\n")) {
      return `"${raw.replace(/"/g, "\"\"")}"`;
    }
    return raw;
  }

  function buildCsvContent(events) {
    const headers = [
      "Ngay",
      "Thu",
      "Ca hoc",
      "Bat dau",
      "Ket thuc",
      "Khung gio",
      "Mon hoc",
      "Lop",
      "Phong",
      "Giang vien",
      "Nguon",
      "Loai"
    ];

    const rows = [headers];
    for (const event of events) {
      rows.push([
        text(event.date),
        WEEKDAY_LABEL[event.weekday] || "",
        text(event.session),
        text(event.startTime),
        text(event.endTime),
        text(event.timeRange),
        text(event.subject),
        text(event.class),
        text(event.room),
        text(event.teacher),
        sourceLabel(event.source),
        categoryLabel(event.category)
      ]);
    }

    return `\uFEFF${rows.map((row) => row.map(escapeCsvCell).join(",")).join("\r\n")}`;
  }

  async function fetchWeekScheduleHtml(targetDate, scheduleType) {
    const params = new URLSearchParams();
    params.set("pNgayHienTai", targetDate.toISOString());
    params.set("pLoaiLich", String(scheduleType));

    const response = await fetch("/SinhVien/GetDanhSachLichTheoTuan", {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
      },
      body: params.toString(),
      credentials: "include"
    });

    if (!response.ok) {
      throw new Error(`Không tải được dữ liệu tuần (${response.status}).`);
    }

    return response.text();
  }

  function parseEventsFromHtmlContent(html, options = {}) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(String(html || ""), "text/html");
    const table = getScheduleTable(doc);
    if (!table) return [];
    return parseScheduleEventsFromTable(table, options);
  }

  function buildMonthWeeks(monthValue) {
    const mondays = getWeekMondaysForMonth(monthValue);
    if (!mondays.length) {
      throw new Error("Tháng xuất không hợp lệ.");
    }
    return mondays;
  }

  function getWeekRangeFromCurrentPage() {
    const { monday, sunday } = getWeekRange();
    return {
      rangeStart: monday,
      rangeEnd: sunday,
      weekMondays: [monday]
    };
  }

  function getMonthRangeFromValue(monthValue) {
    const range = getMonthRange(monthValue);
    if (!range) {
      throw new Error("Giá trị tháng không hợp lệ.");
    }

    return {
      rangeStart: range.firstDay,
      rangeEnd: range.lastDay,
      weekMondays: buildMonthWeeks(monthValue)
    };
  }

  async function collectSystemEventsByOptions(options, context, reportProgress) {
    if (options.rangeMode === "week") {
      const table = getScheduleTable(document);
      if (!table) throw new Error("Không tìm thấy bảng lịch tuần hiện tại.");

      const weekRange = getWeekRangeFromCurrentPage();
      const events = parseScheduleEventsFromTable(table, {
        scheduleType: context.scheduleType,
        weekStart: weekRange.rangeStart
      });

      return {
        systemEvents: events,
        rangeStart: weekRange.rangeStart,
        rangeEnd: weekRange.rangeEnd,
        weekMondays: weekRange.weekMondays
      };
    }

    const monthRange = getMonthRangeFromValue(options.month);
    const allWeekEvents = [];

    for (let i = 0; i < monthRange.weekMondays.length; i += 1) {
      const monday = monthRange.weekMondays[i];
      if (typeof reportProgress === "function") {
        reportProgress(`Đang tải dữ liệu tuần ${i + 1}/${monthRange.weekMondays.length}...`, "info");
      }

      const html = await fetchWeekScheduleHtml(monday, context.scheduleType);
      const weekEvents = parseEventsFromHtmlContent(html, {
        scheduleType: context.scheduleType,
        weekStart: monday
      });

      allWeekEvents.push(...weekEvents);
    }

    const filtered = allWeekEvents.filter((event) => {
      const dateObj = parseDDMMYYYY(event.date);
      return isDateBetween(dateObj, monthRange.rangeStart, monthRange.rangeEnd);
    });

    return {
      systemEvents: filtered,
      rangeStart: monthRange.rangeStart,
      rangeEnd: monthRange.rangeEnd,
      weekMondays: monthRange.weekMondays
    };
  }

  async function buildScheduleExportModel(options, reportProgress) {
    const context = getExportContext();
    if (!context.isSchedulePage) {
      throw new Error("Chỉ hỗ trợ xuất trên trang lịch UNETI.");
    }

    const config = await loadConfig();
    const normalizedOptions = normalizeExportOptions(options, context);
    const systemResult = await collectSystemEventsByOptions(normalizedOptions, context, reportProgress);
    const manualEvents = buildManualEventsForRange(
      config.manualPracticalSchedules,
      systemResult.rangeStart,
      systemResult.rangeEnd
    );

    const mergedEvents = dedupeScheduleEvents([...systemResult.systemEvents, ...manualEvents]);
    const sortedEvents = sortScheduleEvents(mergedEvents);
    const weeks = buildTimetableWeeks(sortedEvents, systemResult.weekMondays);

    return {
      options: normalizedOptions,
      context,
      events: sortedEvents,
      weeks,
      meta: {
        rangeStart: toDDMMYYYY(systemResult.rangeStart),
        rangeEnd: toDDMMYYYY(systemResult.rangeEnd),
        totalEvents: sortedEvents.length,
        totalWeeks: systemResult.weekMondays.length,
        scheduleType: context.scheduleType,
        scheduleTypeLabel: context.scheduleTypeLabel,
        exportedAt: new Date().toISOString(),
        sourcePage: location.href
      }
    };
  }

  function renderEventCardHtml(event) {
    const classes = [
      "tdk-export-event",
      event.category === "exam" ? "exam" : "",
      event.source === "manual" ? "manual" : ""
    ]
      .filter(Boolean)
      .join(" ");

    const lines = [
      `<div><b>${esc(event.subject)}</b></div>`,
      event.class ? `<div>Lớp: ${esc(event.class)}</div>` : "",
      event.timeRange ? `<div>Giờ: ${esc(event.timeRange)}</div>` : "",
      event.room ? `<div>Phòng: ${esc(event.room)}</div>` : "",
      event.teacher ? `<div>GV: ${esc(event.teacher)}</div>` : "",
      event.source === "manual" ? "<div><i>TH bổ sung</i></div>" : ""
    ]
      .filter(Boolean)
      .join("");

    return `<div class="${classes}">${lines}</div>`;
  }

  function buildTimetableSectionHtml(model) {
    if (!model.weeks.length) {
      return `<p>Không có dữ liệu lịch trong phạm vi đã chọn.</p>`;
    }

    return model.weeks
      .map((week) => {
        const headerCells = week.dayColumns.map((day) => `<th>${esc(day.label)}</th>`).join("");
        const sessions = week.sessions.length ? week.sessions : ["Sáng", "Chiều", "Tối"];
        const rows = sessions
          .map((session) => {
            const dayCells = week.dayColumns
              .map((day) => {
                const key = `${day.date}__${session}`;
                const events = week.byDayAndSession.get(key) || [];
                if (!events.length) {
                  return `<td><div class="tdk-export-cell-empty">-</div></td>`;
                }
                return `<td>${events.map(renderEventCardHtml).join("")}</td>`;
              })
              .join("");

            return `<tr><td><b>${esc(session)}</b></td>${dayCells}</tr>`;
          })
          .join("");

        return (
          `<h3>${esc(week.title)}</h3>` +
          `<table class="tdk-export-table">` +
            `<thead><tr><th>Ca học</th>${headerCells}</tr></thead>` +
            `<tbody>${rows}</tbody>` +
          `</table>`
        );
      })
      .join("");
  }

  function buildListSectionHtml(model) {
    if (!model.events.length) {
      return `<p>Không có dữ liệu lịch trong phạm vi đã chọn.</p>`;
    }

    const rows = model.events
      .map((event, index) => {
        return (
          "<tr>" +
            `<td>${index + 1}</td>` +
            `<td>${esc(event.date)}</td>` +
            `<td>${esc(WEEKDAY_LABEL[event.weekday] || "")}</td>` +
            `<td>${esc(event.session)}</td>` +
            `<td>${esc(event.timeRange)}</td>` +
            `<td>${esc(event.subject)}</td>` +
            `<td>${esc(event.class)}</td>` +
            `<td>${esc(event.room)}</td>` +
            `<td>${esc(event.teacher)}</td>` +
            `<td>${esc(sourceLabel(event.source))}</td>` +
            `<td>${esc(categoryLabel(event.category))}</td>` +
          "</tr>"
        );
      })
      .join("");

    return (
      `<table class="tdk-export-table">` +
        "<thead><tr><th>#</th><th>Ngày</th><th>Thứ</th><th>Ca học</th><th>Giờ</th><th>Môn</th><th>Lớp</th><th>Phòng</th><th>GV</th><th>Nguồn</th><th>Loại</th></tr></thead>" +
        `<tbody>${rows}</tbody>` +
      "</table>"
    );
  }

  function buildExportRenderableMarkup(model) {
    const useListLayout = model.options.rangeMode === "month" && model.options.monthLayout === "list";
    const contentHtml = useListLayout ? buildListSectionHtml(model) : buildTimetableSectionHtml(model);

    return (
      `<div class="tdk-export-sheet">` +
        "<h2>UNETI Time Mapper - Xuất lịch học</h2>" +
        `<div class="tdk-export-meta">` +
          `<div>Thời gian xuất: ${esc(new Date(model.meta.exportedAt).toLocaleString("vi-VN"))}</div>` +
          `<div>Phạm vi: ${esc(model.meta.rangeStart)} - ${esc(model.meta.rangeEnd)}</div>` +
          `<div>Loại lịch: ${esc(model.meta.scheduleTypeLabel)}</div>` +
          `<div>Tổng sự kiện: ${model.meta.totalEvents}</div>` +
        "</div>" +
        contentHtml +
      "</div>"
    );
  }

  function buildExportHtmlCss() {
    return (
      "body{margin:0;padding:20px;background:#fff;color:#0f172a;font-family:Arial,sans-serif}" +
      ".tdk-export-sheet h2{margin:0 0 8px;font-size:20px}" +
      ".tdk-export-sheet h3{margin:16px 0 8px;font-size:14px}" +
      ".tdk-export-meta{font-size:12px;color:#334155;margin-bottom:8px;line-height:1.6}" +
      ".tdk-export-table{width:100%;border-collapse:collapse;font-size:12px}" +
      ".tdk-export-table th,.tdk-export-table td{border:1px solid #cbd5e1;padding:6px;vertical-align:top}" +
      ".tdk-export-table th{background:#e2e8f0;text-align:center}" +
      ".tdk-export-event{border-radius:6px;padding:4px 6px;margin:4px 0;background:#f8fafc;border:1px solid #e2e8f0}" +
      ".tdk-export-event.exam{background:#eef7ff;border-color:#bfdbfe}" +
      ".tdk-export-event.manual{background:#f0fdf4;border-color:#86efac}" +
      ".tdk-export-cell-empty{color:#94a3b8;text-align:center}"
    );
  }

  function buildStandaloneHtmlDocument(model) {
    const markup = buildExportRenderableMarkup(model);
    const css = buildExportHtmlCss();

    return (
      "<!doctype html>" +
      "<html lang=\"vi\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">" +
      `<title>UNETI Time Mapper - Export</title><style>${css}</style></head><body>${markup}</body></html>`
    );
  }

  async function renderExportCanvas(model) {
    assertExportLibraries("png");

    const root = document.createElement("div");
    root.className = "tdk-render-root";
    root.innerHTML = `<style>${buildExportHtmlCss()}</style>${buildExportRenderableMarkup(model)}`;
    document.body.appendChild(root);

    try {
      await new Promise((resolve) => setTimeout(resolve, 60));
      return await window.html2canvas(root, {
        scale: 2,
        useCORS: true,
        backgroundColor: "#ffffff",
        windowWidth: Math.max(1400, root.scrollWidth),
        windowHeight: Math.max(900, root.scrollHeight)
      });
    } finally {
      root.remove();
    }
  }

  function canvasToBlob(canvas, type = "image/png") {
    return new Promise((resolve, reject) => {
      canvas.toBlob((blob) => {
        if (!blob) {
          reject(new Error("Không thể tạo blob từ canvas."));
          return;
        }
        resolve(blob);
      }, type);
    });
  }

  async function exportAsXlsx(model, fileName) {
    assertExportLibraries("xlsx");

    const workbook = new window.ExcelJS.Workbook();
    workbook.creator = "UNETI Time Mapper";
    workbook.created = new Date();

    const useListLayout = model.options.rangeMode === "month" && model.options.monthLayout === "list";
    const sheet = workbook.addWorksheet(useListLayout ? "DanhSachLich" : "ThoiKhoaBieu");

    if (useListLayout) {
      const headers = [
        "Ngày",
        "Thứ",
        "Ca học",
        "Bắt đầu",
        "Kết thúc",
        "Giờ",
        "Môn",
        "Lớp",
        "Phòng",
        "GV",
        "Nguồn",
        "Loại"
      ];

      sheet.columns = [
        { width: 12 },
        { width: 10 },
        { width: 10 },
        { width: 10 },
        { width: 10 },
        { width: 14 },
        { width: 34 },
        { width: 24 },
        { width: 22 },
        { width: 22 },
        { width: 14 },
        { width: 12 }
      ];

      sheet.addRow(headers);
      const headerRow = sheet.getRow(1);
      headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
      headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1D4ED8" } };
      headerRow.alignment = { vertical: "middle", horizontal: "center" };
      headerRow.eachCell((cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "CBD5E1" } },
          left: { style: "thin", color: { argb: "CBD5E1" } },
          bottom: { style: "thin", color: { argb: "CBD5E1" } },
          right: { style: "thin", color: { argb: "CBD5E1" } }
        };
      });

      for (const event of model.events) {
        const row = sheet.addRow([
          text(event.date),
          WEEKDAY_LABEL[event.weekday] || "",
          text(event.session),
          text(event.startTime),
          text(event.endTime),
          text(event.timeRange),
          text(event.subject),
          text(event.class),
          text(event.room),
          text(event.teacher),
          sourceLabel(event.source),
          categoryLabel(event.category)
        ]);

        row.alignment = { vertical: "top", horizontal: "left", wrapText: true };
        row.eachCell((cell) => {
          cell.border = {
            top: { style: "thin", color: { argb: "E2E8F0" } },
            left: { style: "thin", color: { argb: "E2E8F0" } },
            bottom: { style: "thin", color: { argb: "E2E8F0" } },
            right: { style: "thin", color: { argb: "E2E8F0" } }
          };
        });
      }
    } else {
      sheet.columns = [
        { width: 14 },
        { width: 22 },
        { width: 22 },
        { width: 22 },
        { width: 22 },
        { width: 22 },
        { width: 22 },
        { width: 22 }
      ];

      let rowPointer = 1;
      for (const week of model.weeks) {
        sheet.mergeCells(`A${rowPointer}:H${rowPointer}`);
        const titleCell = sheet.getCell(`A${rowPointer}`);
        titleCell.value = `Lịch tuần ${week.monday} - ${week.sunday}`;
        titleCell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 12 };
        titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1E3A8A" } };
        titleCell.alignment = { horizontal: "center", vertical: "middle" };
        rowPointer += 1;

        const headerRow = sheet.getRow(rowPointer);
        headerRow.values = ["Ca học", ...week.dayColumns.map((day) => day.label)];
        headerRow.font = { bold: true };
        headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "DBEAFE" } };
        headerRow.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        headerRow.eachCell((cell) => {
          cell.border = {
            top: { style: "thin", color: { argb: "CBD5E1" } },
            left: { style: "thin", color: { argb: "CBD5E1" } },
            bottom: { style: "thin", color: { argb: "CBD5E1" } },
            right: { style: "thin", color: { argb: "CBD5E1" } }
          };
        });
        rowPointer += 1;

        const sessions = week.sessions.length ? week.sessions : ["Sáng", "Chiều", "Tối"];
        for (const session of sessions) {
          const row = sheet.getRow(rowPointer);
          row.getCell(1).value = session;
          row.getCell(1).font = { bold: true };
          row.getCell(1).alignment = { horizontal: "center", vertical: "top" };

          for (let i = 0; i < week.dayColumns.length; i += 1) {
            const day = week.dayColumns[i];
            const key = `${day.date}__${session}`;
            const events = week.byDayAndSession.get(key) || [];
            const cell = row.getCell(i + 2);

            if (!events.length) {
              cell.value = "-";
              cell.alignment = { horizontal: "center", vertical: "middle" };
            } else {
              const value = events
                .map((event) => {
                  return [
                    event.subject,
                    event.class ? `Lớp: ${event.class}` : "",
                    event.timeRange ? `Giờ: ${event.timeRange}` : "",
                    event.room ? `Phòng: ${event.room}` : "",
                    event.teacher ? `GV: ${event.teacher}` : "",
                    event.source === "manual" ? "[TH bổ sung]" : ""
                  ]
                    .filter(Boolean)
                    .join("\n");
                })
                .join("\n\n");

              cell.value = value;
              cell.alignment = { horizontal: "left", vertical: "top", wrapText: true };
            }
          }

          row.eachCell((cell) => {
            cell.border = {
              top: { style: "thin", color: { argb: "CBD5E1" } },
              left: { style: "thin", color: { argb: "CBD5E1" } },
              bottom: { style: "thin", color: { argb: "CBD5E1" } },
              right: { style: "thin", color: { argb: "CBD5E1" } }
            };
          });

          rowPointer += 1;
        }

        rowPointer += 1;
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    downloadBlob(
      fileName,
      new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      })
    );
  }

  async function exportAsPng(model, fileName) {
    const canvas = await renderExportCanvas(model);
    const blob = await canvasToBlob(canvas, "image/png");
    downloadBlob(fileName, blob);
  }

  async function exportAsPdf(model, fileName) {
    assertExportLibraries("pdf");
    const canvas = await renderExportCanvas(model);

    const pdf = new window.jspdf.jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "a4"
    });

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const imgData = canvas.toDataURL("image/png");
    const imgWidth = pageWidth;
    const imgHeight = (canvas.height * imgWidth) / canvas.width;

    let heightLeft = imgHeight;
    let position = 0;

    pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
    heightLeft -= pageHeight;

    while (heightLeft > 0) {
      position -= pageHeight;
      pdf.addPage();
      pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
      heightLeft -= pageHeight;
    }

    pdf.save(fileName);
  }

  async function executeScheduleExport(options, reportProgress) {
    if (exportInProgress) {
      throw new Error("Đang có tác vụ xuất khác. Vui lòng thử lại sau.");
    }

    const context = getExportContext();
    if (!context.isSchedulePage || !context.hasTable) {
      throw new Error("Vui lòng mở trang lịch UNETI để xuất.");
    }

    const normalizedOptions = normalizeExportOptions(options, context);
    assertExportLibraries(normalizedOptions.format);

    exportInProgress = true;
    disableExportSubmit(true);

    const progress =
      typeof reportProgress === "function"
        ? reportProgress
        : () => {};

    try {
      progress("Đang chuẩn bị dữ liệu xuất...", "info");
      const model = await buildScheduleExportModel(normalizedOptions, progress);
      const baseName = buildExportFileBaseName(model);

      if (normalizedOptions.format === "json") {
        const payload = {
          version: 1,
          exportedAt: model.meta.exportedAt,
          sourcePage: model.meta.sourcePage,
          options: model.options,
          events: model.events,
          meta: model.meta
        };

        downloadBlob(
          `${baseName}.json`,
          new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" })
        );
      } else if (normalizedOptions.format === "csv") {
        const csv = buildCsvContent(model.events);
        downloadBlob(`${baseName}.csv`, new Blob([csv], { type: "text/csv;charset=utf-8" }));
      } else if (normalizedOptions.format === "html") {
        const html = buildStandaloneHtmlDocument(model);
        downloadBlob(`${baseName}.html`, new Blob([html], { type: "text/html;charset=utf-8" }));
      } else if (normalizedOptions.format === "xlsx") {
        progress("Đang tạo file XLSX...", "info");
        await exportAsXlsx(model, `${baseName}.xlsx`);
      } else if (normalizedOptions.format === "png") {
        progress("Đang render PNG...", "info");
        await exportAsPng(model, `${baseName}.png`);
      } else if (normalizedOptions.format === "pdf") {
        progress("Đang render PDF...", "info");
        await exportAsPdf(model, `${baseName}.pdf`);
      } else {
        throw new Error("Định dạng xuất không được hỗ trợ.");
      }

      progress(`Đã xuất ${model.events.length} sự kiện.`, "success");
      return {
        fileBase: baseName,
        totalEvents: model.events.length,
        options: normalizedOptions,
        meta: model.meta
      };
    } finally {
      exportInProgress = false;
      disableExportSubmit(false);
    }
  }

  function getGradeTable(root = document) {
    const direct = root.querySelector("table#xemDiem_aaa");
    if (direct) return direct;

    const host = root.querySelector("#xemDiem_aaa");
    if (!host) return null;
    if (host.tagName && host.tagName.toLowerCase() === "table") {
      return host;
    }

    return host.querySelector("table");
  }

  function getGradeExportContext() {
    const table = getGradeTable();
    let totalRows = 0;
    if (table && table.tBodies && table.tBodies.length > 0) {
      for (const tbody of Array.from(table.tBodies)) {
        totalRows += tbody.rows ? tbody.rows.length : 0;
      }
    }

    return {
      isGradePage: isGradePage(),
      hasTable: Boolean(table),
      totalRows,
      pageUrl: location.href
    };
  }

  function parseLocaleNumber(value) {
    const raw = text(value);
    if (!raw) return null;

    const compact = raw.replace(/\s+/g, "");
    const matches = compact.match(/-?\d+(?:[.,]\d+)?/g);
    if (!matches || !matches.length) return null;

    const token = matches[matches.length - 1];
    const lastComma = token.lastIndexOf(",");
    const lastDot = token.lastIndexOf(".");
    let normalized = token;

    if (lastComma >= 0 && lastDot >= 0) {
      if (lastComma > lastDot) {
        normalized = token.replace(/\./g, "").replace(",", ".");
      } else {
        normalized = token.replace(/,/g, "");
      }
    } else if (lastComma >= 0) {
      normalized = token.replace(",", ".");
    }

    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function extractLastNumberFromText(value) {
    const raw = text(value);
    if (!raw) return null;
    const matches = raw.match(/-?\d+(?:[.,]\d+)?/g);
    if (!matches || !matches.length) return null;
    return parseLocaleNumber(matches[matches.length - 1]);
  }

  function mapGradeHeaderToKey(label) {
    const n = norm(label);
    if (!n) return "";

    if (n === "stt" || n.startsWith("stt ")) return "stt";
    if (n.includes("ma lop hoc phan") || n.includes("ma lop hp") || n.includes("ma lop")) return "classCode";
    if (n.includes("ten mon")) return "courseName";
    if (n.includes("so tin chi") || n === "tin chi") return "credits";
    if (n.includes("diem tong ket") || (n.includes("diem") && n.includes("he 10"))) return "total10";
    if (n.includes("thang diem 4") || n.includes("diem tin chi") || (n.includes("diem") && n.includes("he 4"))) return "total4";
    if (n.includes("diem chu")) return "letter";
    if (n.includes("ghi chu")) return "note";
    if (n === "dat" || n.endsWith(" dat")) return "pass";
    return "";
  }

  function detectGradeHeader(table) {
    const allRows = Array.from(table.querySelectorAll("tr"));
    const fallbackByPosition = {
      stt: 0,
      classCode: 1,
      courseName: 2,
      credits: 3
    };

    let best = null;
    for (const row of allRows) {
      const cells = Array.from(row.cells || []);
      if (cells.length < 5) continue;

      const indexes = {};
      let score = 0;
      for (let i = 0; i < cells.length; i += 1) {
        const key = mapGradeHeaderToKey(cells[i].textContent);
        if (!key || typeof indexes[key] === "number") continue;
        indexes[key] = i;
        score += 1;
      }

      if (!best || score > best.score) {
        best = {
          row,
          indexes,
          score,
          cellCount: cells.length
        };
      }
    }

    const header = best || {
      row: null,
      indexes: {},
      score: 0,
      cellCount: 0
    };

    for (const key of Object.keys(fallbackByPosition)) {
      if (typeof header.indexes[key] === "number") continue;
      const fallbackIndex = fallbackByPosition[key];
      if (fallbackIndex < header.cellCount || (!best && fallbackIndex < 9)) {
        header.indexes[key] = fallbackIndex;
      }
    }

    return header;
  }

  function parseGradeTermHeading(textValue) {
    const raw = text(textValue);
    if (!raw) return null;

    const direct = raw.match(/(\d+)\s*\(\s*([^)]+)\s*\)/);
    if (direct) {
      return {
        termLabel: text(direct[1]),
        yearLabel: text(direct[2]).replace(/\s*-\s*/g, "-")
      };
    }

    const normalized = norm(raw);
    const alt = normalized.match(/hoc ky\s*(\d+).*(\d{4})\s*-\s*(\d{4})/);
    if (alt) {
      return {
        termLabel: text(alt[1]),
        yearLabel: `${alt[2]}-${alt[3]}`
      };
    }

    return null;
  }

  function isGradeSummaryRow(rowNorm) {
    return (
      rowNorm.includes("diem trung binh hoc ky he 10") ||
      rowNorm.includes("diem trung binh hoc ky he 4") ||
      rowNorm.includes("diem trung binh tich luy he 10") ||
      rowNorm.includes("diem trung binh tich luy he 4")
    );
  }

  function buildGradeTermKey(yearLabel, termLabel) {
    return `${text(yearLabel)}||${text(termLabel)}`;
  }

  function getGradeCellByTitle(row, titleList) {
    if (!row) return null;
    const titles = Array.isArray(titleList)
      ? titleList.map((item) => text(item)).filter(Boolean)
      : [text(titleList)].filter(Boolean);
    if (!titles.length) return null;

    const cells = Array.from(row.cells || []);
    for (const cell of cells) {
      const titleValue = text(cell.getAttribute("title"));
      if (!titleValue) continue;
      for (const expected of titles) {
        if (titleValue.toLowerCase() === expected.toLowerCase()) {
          return cell;
        }
      }
    }

    return null;
  }

  function pickGradeFieldCells(row, indexes) {
    const cells = Array.from(row.cells || []);
    const safeCell = (idx, fallbackIdx) => {
      if (typeof idx === "number" && idx >= 0 && idx < cells.length) return cells[idx];
      if (typeof fallbackIdx === "number" && fallbackIdx >= 0 && fallbackIdx < cells.length) return cells[fallbackIdx];
      return null;
    };

    const sttCell = safeCell(indexes.stt, 0);
    const classCodeCell = safeCell(indexes.classCode, 1);
    const courseNameCell = safeCell(indexes.courseName, 2);
    const creditsCell = safeCell(indexes.credits, 3);

    const total10Cell =
      getGradeCellByTitle(row, ["DiemTongKet", "DiemTongKet1"]) ||
      (typeof indexes.total10 === "number" ? safeCell(indexes.total10) : null);
    const total4Cell =
      getGradeCellByTitle(row, ["DiemTinChi"]) ||
      (typeof indexes.total4 === "number" ? safeCell(indexes.total4) : null);
    const letterCell =
      getGradeCellByTitle(row, ["DiemChu"]) ||
      (typeof indexes.letter === "number" ? safeCell(indexes.letter) : null);
    const noteCell =
      getGradeCellByTitle(row, ["GhiChu"]) ||
      (typeof indexes.note === "number" ? safeCell(indexes.note) : null);
    const passCell =
      getGradeCellByTitle(row, ["IsDat", "Dat"]) ||
      (row.querySelector("div.check, div.no-check, input[type='checkbox']")
        ? row.querySelector("div.check, div.no-check, input[type='checkbox']").closest("td")
        : null) ||
      (typeof indexes.pass === "number" ? safeCell(indexes.pass) : null);

    return {
      sttCell,
      classCodeCell,
      courseNameCell,
      creditsCell,
      total10Cell,
      total4Cell,
      letterCell,
      noteCell,
      passCell
    };
  }

  function parsePassFromCell(cell, note, letter, total4, total10) {
    if (cell) {
      const checkbox = cell.querySelector("input[type='checkbox']");
      if (checkbox) {
        return checkbox.checked;
      }

      const htmlNorm = norm(cell.innerHTML);
      const valueNorm = norm(cell.textContent);
      if (
        htmlNorm.includes("fa-times") ||
        htmlNorm.includes("fa-close") ||
        htmlNorm.includes("glyphicon-remove") ||
        valueNorm.includes("khong dat") ||
        valueNorm.includes("rot")
      ) {
        return false;
      }
      if (
        htmlNorm.includes("fa-check") ||
        htmlNorm.includes("glyphicon-ok") ||
        valueNorm === "x" ||
        valueNorm.includes("dat")
      ) {
        return true;
      }
    }

    const noteNorm = norm(note);
    const letterRaw = text(letter).toUpperCase();
    if (noteNorm.includes("thi lai")) return false;
    if (letterRaw.startsWith("F")) return false;

    if (Number.isFinite(total4)) return total4 >= 1;
    if (Number.isFinite(total10)) return total10 >= 4;
    return true;
  }

  function gradePoint4From10(total10) {
    if (!Number.isFinite(total10)) return null;
    if (total10 >= 8.5) return 4.0;
    if (total10 >= 8.0) return 3.5;
    if (total10 >= 7.0) return 3.0;
    if (total10 >= 6.5) return 2.5;
    if (total10 >= 5.5) return 2.0;
    if (total10 >= 5.0) return 1.5;
    if (total10 >= 4.0) return 1.0;
    if (total10 >= 3.0) return 0.5;
    return 0.0;
  }

  function normalizeGradeNumberValue(value) {
    return Number.isFinite(value) ? Number(value.toFixed(2)) : null;
  }

  function scrapeGradeRecords() {
    const table = getGradeTable();
    if (!table) {
      throw new Error("Không tìm thấy bảng điểm #xemDiem_aaa.");
    }

    const headerInfo = detectGradeHeader(table);
    const indexes = headerInfo.indexes || {};
    const rows = [];
    if (table.tBodies && table.tBodies.length > 0) {
      for (const tbody of Array.from(table.tBodies)) {
        rows.push(...Array.from(tbody.rows || []));
      }
    } else {
      rows.push(...Array.from(table.querySelectorAll("tr")));
    }

    const webTermSummaryMap = new Map();
    const webCumulativeSummary = {
      gpa10: null,
      gpa4: null
    };
    const records = [];

    let currentTermLabel = "";
    let currentYearLabel = "";

    for (const row of rows) {
      if (row === headerInfo.row) continue;

      const headingCell = row.querySelector("td.row-head");
      if (headingCell) {
        const termHeading = parseGradeTermHeading(headingCell.textContent);
        if (termHeading) {
          currentTermLabel = termHeading.termLabel;
          currentYearLabel = termHeading.yearLabel;
        }
        continue;
      }

      const cells = Array.from(row.cells || []);
      if (!cells.length) continue;

      if (cells.length === 1 && cells[0].colSpan > 1) {
        const termHeading = parseGradeTermHeading(cells[0].textContent);
        if (termHeading) {
          currentTermLabel = termHeading.termLabel;
          currentYearLabel = termHeading.yearLabel;
          continue;
        }
      }

      const rowTextNorm = norm(row.textContent);
      if (!rowTextNorm) continue;

      if (isGradeSummaryRow(rowTextNorm)) {
        const summaryValue = extractLastNumberFromText(row.textContent);
        if (rowTextNorm.includes("diem trung binh hoc ky he 10")) {
          const key = buildGradeTermKey(currentYearLabel, currentTermLabel);
          const termSummary = webTermSummaryMap.get(key) || {
            yearLabel: currentYearLabel,
            termLabel: currentTermLabel,
            gpa10: null,
            gpa4: null
          };
          termSummary.gpa10 = normalizeGradeNumberValue(summaryValue);
          webTermSummaryMap.set(key, termSummary);
        } else if (rowTextNorm.includes("diem trung binh hoc ky he 4")) {
          const key = buildGradeTermKey(currentYearLabel, currentTermLabel);
          const termSummary = webTermSummaryMap.get(key) || {
            yearLabel: currentYearLabel,
            termLabel: currentTermLabel,
            gpa10: null,
            gpa4: null
          };
          termSummary.gpa4 = normalizeGradeNumberValue(summaryValue);
          webTermSummaryMap.set(key, termSummary);
        } else if (rowTextNorm.includes("diem trung binh tich luy he 4")) {
          webCumulativeSummary.gpa4 = normalizeGradeNumberValue(summaryValue);
        } else if (rowTextNorm.includes("diem trung binh tich luy")) {
          webCumulativeSummary.gpa10 = normalizeGradeNumberValue(summaryValue);
        }

        continue;
      }

      const fields = pickGradeFieldCells(row, indexes);
      const sttCell = fields.sttCell;
      const classCodeCell = fields.classCodeCell;
      const courseNameCell = fields.courseNameCell;
      const creditsCell = fields.creditsCell;
      const total10Cell = fields.total10Cell;
      const total4Cell = fields.total4Cell;
      const letterCell = fields.letterCell;
      const noteCell = fields.noteCell;
      const passCell = fields.passCell;

      const classCode = text(classCodeCell && classCodeCell.textContent);
      const courseName = text(courseNameCell && courseNameCell.textContent);
      if (!classCode && !courseName) continue;

      const sttValue = parseLocaleNumber(sttCell && sttCell.textContent);
      const creditsValue = parseLocaleNumber(creditsCell && creditsCell.textContent);
      const total10Value = parseLocaleNumber(total10Cell && total10Cell.textContent);
      const total4Value = parseLocaleNumber(total4Cell && total4Cell.textContent);
      const letter = text(letterCell && letterCell.textContent);
      const note = text(noteCell && noteCell.textContent);
      const pass = parsePassFromCell(passCell, note, letter, total4Value, total10Value);
      const includeInGPA =
        Number.isFinite(creditsValue) &&
        creditsValue > 0 &&
        Number.isFinite(total4Value);

      records.push({
        yearLabel: text(currentYearLabel),
        termLabel: text(currentTermLabel),
        stt: Number.isFinite(sttValue) ? Math.round(sttValue) : null,
        classCode,
        courseName,
        credits: normalizeGradeNumberValue(creditsValue),
        total10: normalizeGradeNumberValue(total10Value),
        total4: normalizeGradeNumberValue(total4Value),
        letter,
        note,
        pass: pass === true,
        includeInGPA
      });
    }

    if (!records.length) {
      throw new Error("Không tìm thấy dữ liệu môn học trong bảng điểm.");
    }

    return {
      records,
      webSummary: {
        terms: Array.from(webTermSummaryMap.values()),
        cumulative: webCumulativeSummary
      }
    };
  }

  function buildEmptyGradeSummaryBucket() {
    return {
      totalCreditsGPA: 0,
      weighted4Current: 0,
      weighted10Current: 0,
      weighted4Predicted: 0,
      weighted10Predicted: 0,
      passedCreditsCurrent: 0,
      debtCreditsCurrent: 0,
      passedCreditsPredicted: 0,
      debtCreditsPredicted: 0
    };
  }

  function finalizeGradeSummaryBucket(bucket) {
    const baseCredits = bucket.totalCreditsGPA;
    const gpa4Current = baseCredits > 0 ? bucket.weighted4Current / baseCredits : null;
    const gpa10Current = baseCredits > 0 ? bucket.weighted10Current / baseCredits : null;
    const gpa4Predicted = baseCredits > 0 ? bucket.weighted4Predicted / baseCredits : null;
    const gpa10Predicted = baseCredits > 0 ? bucket.weighted10Predicted / baseCredits : null;

    return {
      totalCreditsGPA: normalizeGradeNumberValue(bucket.totalCreditsGPA),
      gpa4Current: normalizeGradeNumberValue(gpa4Current),
      gpa10Current: normalizeGradeNumberValue(gpa10Current),
      gpa4Predicted: normalizeGradeNumberValue(gpa4Predicted),
      gpa10Predicted: normalizeGradeNumberValue(gpa10Predicted),
      passedCreditsCurrent: normalizeGradeNumberValue(bucket.passedCreditsCurrent),
      debtCreditsCurrent: normalizeGradeNumberValue(bucket.debtCreditsCurrent),
      passedCreditsPredicted: normalizeGradeNumberValue(bucket.passedCreditsPredicted),
      debtCreditsPredicted: normalizeGradeNumberValue(bucket.debtCreditsPredicted)
    };
  }

  function calculateGradeSummary(records) {
    const termMap = new Map();
    const cumulativeBucket = buildEmptyGradeSummaryBucket();

    const getTermBucket = (record) => {
      const key = buildGradeTermKey(record.yearLabel, record.termLabel);
      if (!termMap.has(key)) {
        termMap.set(key, {
          key,
          yearLabel: text(record.yearLabel),
          termLabel: text(record.termLabel),
          bucket: buildEmptyGradeSummaryBucket()
        });
      }
      return termMap.get(key);
    };

    const accumulate = (bucket, record) => {
      const credits = Number(record.credits);
      const total10Current = Number.isFinite(record.total10) ? record.total10 : null;
      const total4Current = Number.isFinite(record.total4) ? record.total4 : null;
      const total10Predicted = Number.isFinite(total10Current) ? total10Current : null;
      const total4Predicted =
        Number.isFinite(total10Predicted)
          ? gradePoint4From10(total10Predicted)
          : Number.isFinite(total4Current)
            ? total4Current
            : null;

      if (Number.isFinite(credits) && credits > 0) {
        if (record.pass === true) {
          bucket.passedCreditsCurrent += credits;
        } else {
          bucket.debtCreditsCurrent += credits;
        }
      }

      if (!(Number.isFinite(credits) && credits > 0 && record.includeInGPA === true)) {
        return;
      }

      bucket.totalCreditsGPA += credits;
      if (Number.isFinite(total4Current)) {
        bucket.weighted4Current += credits * total4Current;
      }
      if (Number.isFinite(total10Current)) {
        bucket.weighted10Current += credits * total10Current;
      }
      if (Number.isFinite(total4Predicted)) {
        bucket.weighted4Predicted += credits * total4Predicted;
      }
      if (Number.isFinite(total10Predicted)) {
        bucket.weighted10Predicted += credits * total10Predicted;
      }

      if (Number.isFinite(total4Predicted) && total4Predicted >= 1) {
        bucket.passedCreditsPredicted += credits;
      } else {
        bucket.debtCreditsPredicted += credits;
      }
    };

    for (const record of records) {
      const term = getTermBucket(record);
      accumulate(term.bucket, record);
      accumulate(cumulativeBucket, record);
    }

    return {
      terms: Array.from(termMap.values()).map((item) => {
        return {
          yearLabel: item.yearLabel,
          termLabel: item.termLabel,
          ...finalizeGradeSummaryBucket(item.bucket)
        };
      }),
      cumulative: finalizeGradeSummaryBucket(cumulativeBucket)
    };
  }

  function buildGradeExportStamp() {
    const now = new Date();
    const yyyy = String(now.getFullYear());
    const mm = String(now.getMonth() + 1).padStart(2, "0");
    const dd = String(now.getDate()).padStart(2, "0");
    const hh = String(now.getHours()).padStart(2, "0");
    const min = String(now.getMinutes()).padStart(2, "0");
    return `${yyyy}${mm}${dd}-${hh}${min}`;
  }

  function buildGradeExportBaseName() {
    return `UNETI-Grades-${buildGradeExportStamp()}`;
  }

  function buildGradeCsvContent(records) {
    const headers = [
      "Year",
      "Term",
      "STT",
      "ClassCode",
      "CourseName",
      "Credits",
      "Total10Current",
      "Total4Current",
      "Letter",
      "Note",
      "Pass",
      "IncludeInGPA"
    ];
    const rows = [headers];

    for (const record of records) {
      rows.push([
        text(record.yearLabel),
        text(record.termLabel),
        Number.isFinite(record.stt) ? record.stt : "",
        text(record.classCode),
        text(record.courseName),
        Number.isFinite(record.credits) ? record.credits.toFixed(2) : "",
        Number.isFinite(record.total10) ? record.total10.toFixed(2) : "",
        Number.isFinite(record.total4) ? record.total4.toFixed(2) : "",
        text(record.letter),
        text(record.note),
        record.pass ? "TRUE" : "FALSE",
        record.includeInGPA ? "TRUE" : "FALSE"
      ]);
    }

    return `\uFEFF${rows.map((row) => row.map(escapeCsvCell).join(",")).join("\r\n")}`;
  }

  function autoFitWorksheetColumns(sheet, minWidth = 10, maxWidth = 44) {
    if (!sheet || !Array.isArray(sheet.columns)) return;

    for (const column of sheet.columns) {
      let width = minWidth;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const rawValue = cell && cell.value;
        if (rawValue == null) return;

        let rendered = "";
        if (typeof rawValue === "object" && rawValue.formula) {
          rendered = String(rawValue.result == null ? "" : rawValue.result);
        } else if (typeof rawValue === "object" && Array.isArray(rawValue.richText)) {
          rendered = rawValue.richText.map((part) => part && part.text ? part.text : "").join("");
        } else {
          rendered = String(rawValue);
        }

        width = Math.max(width, Math.min(maxWidth, rendered.length + 2));
      });
      column.width = Math.max(minWidth, Math.min(maxWidth, width));
    }
  }

  async function exportGradesAsXlsx(records, summary, webSummary, fileName) {
    assertExportLibraries("xlsx");

    const workbook = new window.ExcelJS.Workbook();
    workbook.creator = "UNETI Time Mapper";
    workbook.created = new Date();

    const gradesSheet = workbook.addWorksheet("Grades");
    gradesSheet.columns = [
      { header: "Year", key: "yearLabel", width: 12 },
      { header: "Term", key: "termLabel", width: 10 },
      { header: "STT", key: "stt", width: 8 },
      { header: "ClassCode", key: "classCode", width: 18 },
      { header: "CourseName", key: "courseName", width: 36 },
      { header: "Credits", key: "credits", width: 12 },
      { header: "Total10 (Current)", key: "total10Current", width: 16 },
      { header: "Total4 (Current)", key: "total4Current", width: 16 },
      { header: "Letter", key: "letter", width: 10 },
      { header: "Note", key: "note", width: 18 },
      { header: "Pass", key: "pass", width: 10 },
      { header: "Total10 (Edit)", key: "total10Edit", width: 14 },
      { header: "Total4 (Calc)", key: "total4Calc", width: 14 },
      { header: "IncludeInGPA", key: "includeInGPA", width: 14 }
    ];

    const headerRow = gradesSheet.getRow(1);
    headerRow.height = 24;
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1D4ED8" } };
    headerRow.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    headerRow.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "CBD5E1" } },
        left: { style: "thin", color: { argb: "CBD5E1" } },
        bottom: { style: "thin", color: { argb: "CBD5E1" } },
        right: { style: "thin", color: { argb: "CBD5E1" } }
      };
    });

    gradesSheet.views = [{ state: "frozen", ySplit: 1 }];

    for (const record of records) {
      const row = gradesSheet.addRow({
        yearLabel: text(record.yearLabel),
        termLabel: text(record.termLabel),
        stt: Number.isFinite(record.stt) ? record.stt : "",
        classCode: text(record.classCode),
        courseName: text(record.courseName),
        credits: Number.isFinite(record.credits) ? record.credits : "",
        total10Current: Number.isFinite(record.total10) ? record.total10 : "",
        total4Current: Number.isFinite(record.total4) ? record.total4 : "",
        letter: text(record.letter),
        note: text(record.note),
        pass: record.pass === true,
        total10Edit: Number.isFinite(record.total10) ? record.total10 : "",
        total4Calc: "",
        includeInGPA: record.includeInGPA === true
      });

      const rowNumber = row.number;
      row.getCell("L").value = Number.isFinite(record.total10) ? record.total10 : "";
      row.getCell("M").value = {
        formula: `IF(L${rowNumber}="","",IF(ABS(L${rowNumber}-G${rowNumber})<0.0001,H${rowNumber},LOOKUP(L${rowNumber},Lookup!$A$2:$A$10,Lookup!$B$2:$B$10)))`
      };
      row.getCell("N").value = {
        formula: `AND(F${rowNumber}>0,ISNUMBER(M${rowNumber}))`
      };

      row.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "E2E8F0" } },
          left: { style: "thin", color: { argb: "E2E8F0" } },
          bottom: { style: "thin", color: { argb: "E2E8F0" } },
          right: { style: "thin", color: { argb: "E2E8F0" } }
        };
      });

      row.getCell("A").alignment = { horizontal: "center", vertical: "middle" };
      row.getCell("B").alignment = { horizontal: "center", vertical: "middle" };
      row.getCell("C").alignment = { horizontal: "center", vertical: "middle" };
      row.getCell("D").alignment = { horizontal: "center", vertical: "middle" };
      row.getCell("F").alignment = { horizontal: "center", vertical: "middle" };
      row.getCell("G").alignment = { horizontal: "center", vertical: "middle" };
      row.getCell("H").alignment = { horizontal: "center", vertical: "middle" };
      row.getCell("I").alignment = { horizontal: "center", vertical: "middle" };
      row.getCell("K").alignment = { horizontal: "center", vertical: "middle" };
      row.getCell("L").alignment = { horizontal: "center", vertical: "middle" };
      row.getCell("M").alignment = { horizontal: "center", vertical: "middle" };
      row.getCell("N").alignment = { horizontal: "center", vertical: "middle" };

      const noteNorm = norm(record.note);
      if (noteNorm.includes("thi lai") || record.pass === false) {
        row.eachCell((cell) => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFF1F2" }
          };
        });
      } else if (rowNumber % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "F8FAFC" }
          };
        });
      }
    }

    const gradesDataEnd = Math.max(2, gradesSheet.rowCount);
    gradesSheet.autoFilter = {
      from: "A1",
      to: "N1"
    };
    for (const columnId of ["F", "G", "H", "L", "M"]) {
      gradesSheet.getColumn(columnId).numFmt = "0.00";
    }
    autoFitWorksheetColumns(gradesSheet, 10, 40);

    const gpaSheet = workbook.addWorksheet("GPA");
    gpaSheet.views = [{ state: "frozen", ySplit: 1 }];
    gpaSheet.columns = [
      { header: "Year", key: "yearLabel", width: 12 },
      { header: "Term", key: "termLabel", width: 10 },
      { header: "Credits(GPA)", key: "creditsGpa", width: 14 },
      { header: "GPA4 Current", key: "gpa4Current", width: 13 },
      { header: "GPA10 Current", key: "gpa10Current", width: 13 },
      { header: "GPA4 Predicted", key: "gpa4Predicted", width: 14 },
      { header: "GPA10 Predicted", key: "gpa10Predicted", width: 14 },
      { header: "Passed Credits Current", key: "passedCurrent", width: 17 },
      { header: "Debt Credits Current", key: "debtCurrent", width: 15 },
      { header: "Passed Credits Predicted", key: "passedPredicted", width: 18 },
      { header: "Debt Credits Predicted", key: "debtPredicted", width: 16 },
      { header: "Web GPA10", key: "webGpa10", width: 11 },
      { header: "Web GPA4", key: "webGpa4", width: 11 }
    ];

    const gpaHeaderRow = gpaSheet.getRow(1);
    gpaHeaderRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    gpaHeaderRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "0F766E" } };
    gpaHeaderRow.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    gpaHeaderRow.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "CBD5E1" } },
        left: { style: "thin", color: { argb: "CBD5E1" } },
        bottom: { style: "thin", color: { argb: "CBD5E1" } },
        right: { style: "thin", color: { argb: "CBD5E1" } }
      };
    });

    const yearRange = `Grades!$A$2:$A$${gradesDataEnd}`;
    const termRange = `Grades!$B$2:$B$${gradesDataEnd}`;
    const creditsRange = `Grades!$F$2:$F$${gradesDataEnd}`;
    const total10CurrentRange = `Grades!$G$2:$G$${gradesDataEnd}`;
    const total4CurrentRange = `Grades!$H$2:$H$${gradesDataEnd}`;
    const passRange = `Grades!$K$2:$K$${gradesDataEnd}`;
    const total10EditRange = `Grades!$L$2:$L$${gradesDataEnd}`;
    const total4CalcRange = `Grades!$M$2:$M$${gradesDataEnd}`;
    const includeRange = `Grades!$N$2:$N$${gradesDataEnd}`;

    const webTermMap = new Map();
    for (const term of (webSummary && Array.isArray(webSummary.terms) ? webSummary.terms : [])) {
      webTermMap.set(buildGradeTermKey(term.yearLabel, term.termLabel), term);
    }

    const sortedTerms = [...summary.terms].sort((a, b) => {
      const yearDiff = text(a.yearLabel).localeCompare(text(b.yearLabel), "vi", { numeric: true });
      if (yearDiff !== 0) return yearDiff;
      return text(a.termLabel).localeCompare(text(b.termLabel), "vi", { numeric: true });
    });

    let gpaRowPointer = 2;
    for (const term of sortedTerms) {
      const rowNumber = gpaRowPointer;
      const termKey = buildGradeTermKey(term.yearLabel, term.termLabel);
      const webTerm = webTermMap.get(termKey) || {};

      const row = gpaSheet.getRow(rowNumber);
      row.getCell("A").value = text(term.yearLabel);
      row.getCell("B").value = text(term.termLabel);
      row.getCell("C").value = {
        formula: `SUMIFS(${creditsRange},${yearRange},A${rowNumber},${termRange},B${rowNumber},${includeRange},TRUE)`
      };
      row.getCell("D").value = {
        formula: `IF(C${rowNumber}=0,"",SUMPRODUCT((${yearRange}=A${rowNumber})*(${termRange}=B${rowNumber})*(${includeRange}=TRUE)*(${creditsRange})*(${total4CurrentRange}))/C${rowNumber})`
      };
      row.getCell("E").value = {
        formula: `IF(C${rowNumber}=0,"",SUMPRODUCT((${yearRange}=A${rowNumber})*(${termRange}=B${rowNumber})*(${includeRange}=TRUE)*(${creditsRange})*(${total10CurrentRange}))/C${rowNumber})`
      };
      row.getCell("F").value = {
        formula: `IF(C${rowNumber}=0,"",SUMPRODUCT((${yearRange}=A${rowNumber})*(${termRange}=B${rowNumber})*(${includeRange}=TRUE)*(${creditsRange})*(${total4CalcRange}))/C${rowNumber})`
      };
      row.getCell("G").value = {
        formula: `IF(C${rowNumber}=0,"",SUMPRODUCT((${yearRange}=A${rowNumber})*(${termRange}=B${rowNumber})*(${includeRange}=TRUE)*(${creditsRange})*(${total10EditRange}))/C${rowNumber})`
      };
      row.getCell("H").value = {
        formula: `SUMPRODUCT((${yearRange}=A${rowNumber})*(${termRange}=B${rowNumber})*(${creditsRange}>0)*(${passRange}=TRUE)*(${creditsRange}))`
      };
      row.getCell("I").value = {
        formula: `SUMPRODUCT((${yearRange}=A${rowNumber})*(${termRange}=B${rowNumber})*(${creditsRange}>0)*(${passRange}=FALSE)*(${creditsRange}))`
      };
      row.getCell("J").value = {
        formula: `SUMPRODUCT((${yearRange}=A${rowNumber})*(${termRange}=B${rowNumber})*(${includeRange}=TRUE)*(${creditsRange}>0)*(${total4CalcRange}>=1)*(${creditsRange}))`
      };
      row.getCell("K").value = {
        formula: `SUMPRODUCT((${yearRange}=A${rowNumber})*(${termRange}=B${rowNumber})*(${includeRange}=TRUE)*(${creditsRange}>0)*(${total4CalcRange}<1)*(${creditsRange}))`
      };
      row.getCell("L").value = Number.isFinite(webTerm.gpa10) ? webTerm.gpa10 : "";
      row.getCell("M").value = Number.isFinite(webTerm.gpa4) ? webTerm.gpa4 : "";
      row.alignment = { horizontal: "center", vertical: "middle" };
      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "E2E8F0" } },
          left: { style: "thin", color: { argb: "E2E8F0" } },
          bottom: { style: "thin", color: { argb: "E2E8F0" } },
          right: { style: "thin", color: { argb: "E2E8F0" } }
        };
      });
      if (rowNumber % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "F8FAFC" }
          };
        });
      }

      gpaRowPointer += 1;
    }

    const cumulativeTitleRow = gpaRowPointer + 1;
    gpaSheet.mergeCells(`A${cumulativeTitleRow}:D${cumulativeTitleRow}`);
    const cumulativeTitleCell = gpaSheet.getCell(`A${cumulativeTitleRow}`);
    cumulativeTitleCell.value = "Tổng hợp tích lũy";
    cumulativeTitleCell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cumulativeTitleCell.alignment = { horizontal: "center", vertical: "middle" };
    cumulativeTitleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1E3A8A" } };

    const cumulativeHeaderRow = cumulativeTitleRow + 1;
    gpaSheet.getCell(`A${cumulativeHeaderRow}`).value = "Chỉ số";
    gpaSheet.getCell(`B${cumulativeHeaderRow}`).value = "Current";
    gpaSheet.getCell(`C${cumulativeHeaderRow}`).value = "Predicted";
    gpaSheet.getCell(`D${cumulativeHeaderRow}`).value = "Web";
    const cumulativeHeader = gpaSheet.getRow(cumulativeHeaderRow);
    cumulativeHeader.font = { bold: true };
    cumulativeHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "DBEAFE" } };
    cumulativeHeader.alignment = { horizontal: "center", vertical: "middle" };

    const cumulativeRows = [
      "Tổng tín chỉ tính GPA",
      "GPA hệ 4",
      "GPA hệ 10",
      "Tổng tín chỉ đạt",
      "Số tín chỉ nợ"
    ];

    const cumulativeWeb = webSummary && webSummary.cumulative ? webSummary.cumulative : {};
    for (let i = 0; i < cumulativeRows.length; i += 1) {
      const rowNumber = cumulativeHeaderRow + 1 + i;
      const label = cumulativeRows[i];
      gpaSheet.getCell(`A${rowNumber}`).value = label;

      if (label === "Tổng tín chỉ tính GPA") {
        gpaSheet.getCell(`B${rowNumber}`).value = {
          formula: `SUMIF(${includeRange},TRUE,${creditsRange})`
        };
        gpaSheet.getCell(`C${rowNumber}`).value = {
          formula: `SUMIF(${includeRange},TRUE,${creditsRange})`
        };
      } else if (label === "GPA hệ 4") {
        gpaSheet.getCell(`B${rowNumber}`).value = {
          formula: `IF(B${cumulativeHeaderRow + 1}=0,"",SUMPRODUCT((${includeRange}=TRUE)*(${creditsRange})*(${total4CurrentRange}))/B${cumulativeHeaderRow + 1})`
        };
        gpaSheet.getCell(`C${rowNumber}`).value = {
          formula: `IF(C${cumulativeHeaderRow + 1}=0,"",SUMPRODUCT((${includeRange}=TRUE)*(${creditsRange})*(${total4CalcRange}))/C${cumulativeHeaderRow + 1})`
        };
        gpaSheet.getCell(`D${rowNumber}`).value = Number.isFinite(cumulativeWeb.gpa4) ? cumulativeWeb.gpa4 : "";
      } else if (label === "GPA hệ 10") {
        gpaSheet.getCell(`B${rowNumber}`).value = {
          formula: `IF(B${cumulativeHeaderRow + 1}=0,"",SUMPRODUCT((${includeRange}=TRUE)*(${creditsRange})*(${total10CurrentRange}))/B${cumulativeHeaderRow + 1})`
        };
        gpaSheet.getCell(`C${rowNumber}`).value = {
          formula: `IF(C${cumulativeHeaderRow + 1}=0,"",SUMPRODUCT((${includeRange}=TRUE)*(${creditsRange})*(${total10EditRange}))/C${cumulativeHeaderRow + 1})`
        };
        gpaSheet.getCell(`D${rowNumber}`).value = Number.isFinite(cumulativeWeb.gpa10) ? cumulativeWeb.gpa10 : "";
      } else if (label === "Tổng tín chỉ đạt") {
        gpaSheet.getCell(`B${rowNumber}`).value = {
          formula: `SUMPRODUCT((${creditsRange}>0)*(${passRange}=TRUE)*(${creditsRange}))`
        };
        gpaSheet.getCell(`C${rowNumber}`).value = {
          formula: `SUMPRODUCT((${includeRange}=TRUE)*(${creditsRange}>0)*(${total4CalcRange}>=1)*(${creditsRange}))`
        };
      } else if (label === "Số tín chỉ nợ") {
        gpaSheet.getCell(`B${rowNumber}`).value = {
          formula: `SUMPRODUCT((${creditsRange}>0)*(${passRange}=FALSE)*(${creditsRange}))`
        };
        gpaSheet.getCell(`C${rowNumber}`).value = {
          formula: `SUMPRODUCT((${includeRange}=TRUE)*(${creditsRange}>0)*(${total4CalcRange}<1)*(${creditsRange}))`
        };
      }

      const row = gpaSheet.getRow(rowNumber);
      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "E2E8F0" } },
          left: { style: "thin", color: { argb: "E2E8F0" } },
          bottom: { style: "thin", color: { argb: "E2E8F0" } },
          right: { style: "thin", color: { argb: "E2E8F0" } }
        };
      });
    }

    for (const columnId of ["C", "D", "E", "F", "G", "L", "M", "B"]) {
      gpaSheet.getColumn(columnId).numFmt = "0.00";
    }
    autoFitWorksheetColumns(gpaSheet, 11, 34);

    const lookupSheet = workbook.addWorksheet("Lookup");
    lookupSheet.columns = [
      { header: "Threshold10", key: "threshold10", width: 12 },
      { header: "Point4", key: "point4", width: 10 }
    ];
    const lookupHeader = lookupSheet.getRow(1);
    lookupHeader.font = { bold: true };
    lookupHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "E2E8F0" } };

    const lookupRows = [
      [0.0, 0.0],
      [3.0, 0.5],
      [4.0, 1.0],
      [5.0, 1.5],
      [5.5, 2.0],
      [6.5, 2.5],
      [7.0, 3.0],
      [8.0, 3.5],
      [8.5, 4.0]
    ];
    for (const [threshold10, point4] of lookupRows) {
      const row = lookupSheet.addRow({ threshold10, point4 });
      row.getCell(1).numFmt = "0.00";
      row.getCell(2).numFmt = "0.00";
    }
    autoFitWorksheetColumns(lookupSheet, 10, 18);

    const buffer = await workbook.xlsx.writeBuffer();
    downloadBlob(
      fileName,
      new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      })
    );
  }

  function normalizeGradeExportOptions(rawOptions) {
    const source = rawOptions && typeof rawOptions === "object" ? rawOptions : {};
    const format = text(source.format).toLowerCase();
    if (format === "csv" || format === "json" || format === "xlsx") {
      return { format };
    }
    return { format: "xlsx" };
  }

  async function executeGradeExport(options) {
    if (exportInProgress) {
      throw new Error("Đang có tác vụ xuất khác. Vui lòng thử lại sau.");
    }

    const context = getGradeExportContext();
    if (!context.isGradePage || !context.hasTable) {
      throw new Error("Vui lòng mở trang Kết quả học tập để xuất.");
    }

    const normalizedOptions = normalizeGradeExportOptions(options);
    if (normalizedOptions.format === "xlsx") {
      assertExportLibraries("xlsx");
    }

    exportInProgress = true;
    try {
      const scraped = scrapeGradeRecords();
      const records = scraped.records;
      const webSummary = scraped.webSummary;
      const summary = calculateGradeSummary(records);
      const baseName = buildGradeExportBaseName();

      if (normalizedOptions.format === "json") {
        const payload = {
          version: 1,
          exportedAt: new Date().toISOString(),
          sourcePage: location.href,
          records,
          summary,
          webSummary
        };
        downloadBlob(
          `${baseName}.json`,
          new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" })
        );
      } else if (normalizedOptions.format === "csv") {
        const csv = buildGradeCsvContent(records);
        downloadBlob(`${baseName}.csv`, new Blob([csv], { type: "text/csv;charset=utf-8" }));
      } else if (normalizedOptions.format === "xlsx") {
        await exportGradesAsXlsx(records, summary, webSummary, `${baseName}.xlsx`);
      } else {
        throw new Error("Định dạng xuất điểm không được hỗ trợ.");
      }

      return {
        fileBase: baseName,
        totalRecords: records.length,
        options: normalizedOptions,
        summary
      };
    } finally {
      exportInProgress = false;
    }
  }

  function ensureGradeInlineStyle() {
    if (document.getElementById("tdk-grade-inline-style")) return;
    const style = document.createElement("style");
    style.id = "tdk-grade-inline-style";
    style.textContent =
      "#tdk-grade-inline-controls{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-left:8px}" +
      "#tdk-grade-open-btn,#tdk-grade-reset-all-btn,#tdk-grade-calc-btn{height:32px;line-height:30px;padding:0 12px;border-radius:6px;text-decoration:none}" +
      "#tdk-grade-calc-scope{height:32px;border:1px solid #cbd5e1;border-radius:6px;padding:0 8px;background:#fff}" +
      "#tdk-grade-open-btn.tdk-active{background:#1d4ed8;color:#fff;border-color:#1d4ed8}" +
      "#tdk-grade-inline-status{font-size:12px;color:#334155}" +
      "#tdk-grade-inline-status.error{color:#b91c1c}" +
      "#tdk-grade-inline-status.success{color:#065f46}" +
      "#xemDiem_aaa.tdk-grade-inline-enabled td.tdk-grade-inline-editable{cursor:text;outline:1px dashed #bfdbfe;outline-offset:-3px}" +
      "#xemDiem_aaa td.tdk-grade-cell-changed{background:#fef9c3 !important}" +
      "#xemDiem_aaa tr.tdk-grade-row-changed td{box-shadow: inset 0 0 0 1px rgba(234,179,8,.18)}" +
      "#xemDiem_aaa tr.tdk-grade-row-selected td{box-shadow: inset 0 0 0 2px rgba(37,99,235,.28)}" +
      "#xemDiem_aaa span.tdk-grade-summary-changed{background:#fef3c7;padding:1px 4px;border-radius:4px}" +
      ".tdk-grade-inline-input{width:100%;min-width:54px;height:28px;border:1px solid #93c5fd;border-radius:4px;padding:0 6px;font-size:12px;line-height:28px}";
    document.head.appendChild(style);
  }

  function setGradeInlineStatus(message, type) {
    const statusEl = document.getElementById("tdk-grade-inline-status");
    if (!statusEl) return;
    statusEl.textContent = message || "";
    statusEl.className = "";
    if (type) statusEl.classList.add(type);
  }

  function createEmptySummaryRefs() {
    return {
      terms: new Map(),
      cumulativeByTerm: new Map(),
      cumulative: {
        gpa10: [],
        gpa4: [],
        credits: [],
        debt: []
      }
    };
  }

  function createCumulativeRefBucket() {
    return {
      gpa10: [],
      gpa4: [],
      credits: [],
      debt: []
    };
  }

  function appendSummaryRef(bucket, key, field, node) {
    if (!(node instanceof Element)) return;
    if (!bucket.terms.has(key)) {
      bucket.terms.set(key, {
        gpa10: [],
        gpa4: []
      });
    }
    bucket.terms.get(key)[field].push(node);
  }

  function appendCumulativeSummaryRef(bucket, key, field, node) {
    if (!(node instanceof Element)) return;
    const normalizedKey = text(key);
    if (!normalizedKey) {
      bucket.cumulative[field].push(node);
      return;
    }
    if (!bucket.cumulativeByTerm.has(normalizedKey)) {
      bucket.cumulativeByTerm.set(normalizedKey, createCumulativeRefBucket());
    }
    bucket.cumulativeByTerm.get(normalizedKey)[field].push(node);
  }

  function collectSummaryRefsFromTable(table, state) {
    const refs = createEmptySummaryRefs();
    let currentTermLabel = "";
    let currentYearLabel = "";
    const rows = table ? Array.from(table.querySelectorAll("tbody tr")) : [];

    for (const row of rows) {
      const headingCell = row.querySelector("td.row-head");
      if (headingCell) {
        const heading = parseGradeTermHeading(headingCell.textContent);
        if (heading) {
          currentTermLabel = heading.termLabel;
          currentYearLabel = heading.yearLabel;
        }
        continue;
      }

      const spans = Array.from(row.querySelectorAll("span[lang]"));
      if (!spans.length) continue;

      for (const span of spans) {
        const lang = text(span.getAttribute("lang")).toLowerCase();
        const valueEl = span.nextElementSibling instanceof Element ? span.nextElementSibling : null;
        if (!(valueEl instanceof Element)) continue;

        if (!state.summaryOriginalValues.has(valueEl)) {
          state.summaryOriginalValues.set(valueEl, parseLocaleNumber(valueEl.textContent));
        }

        if (lang.includes("kqht-tkhk-diemtbhocluc") && !lang.includes("tichluy")) {
          appendSummaryRef(refs, buildGradeTermKey(currentYearLabel, currentTermLabel), "gpa10", valueEl);
        } else if (lang.includes("kqht-tkhk-diemtbtinchi") && !lang.includes("tichluy")) {
          appendSummaryRef(refs, buildGradeTermKey(currentYearLabel, currentTermLabel), "gpa4", valueEl);
        } else if (lang.includes("kqht-tkhk-diemtbhocluctichluy")) {
          appendCumulativeSummaryRef(refs, buildGradeTermKey(currentYearLabel, currentTermLabel), "gpa10", valueEl);
        } else if (lang.includes("kqht-tkhk-diemtbtinchitichluy")) {
          appendCumulativeSummaryRef(refs, buildGradeTermKey(currentYearLabel, currentTermLabel), "gpa4", valueEl);
        } else if (lang.includes("kqht-tkhk-sotctichluy")) {
          appendCumulativeSummaryRef(refs, buildGradeTermKey(currentYearLabel, currentTermLabel), "credits", valueEl);
        } else if (lang.includes("kqht-tkhk-sotckhongdat")) {
          appendCumulativeSummaryRef(refs, buildGradeTermKey(currentYearLabel, currentTermLabel), "debt", valueEl);
        }
      }
    }

    return refs;
  }

  function setSummaryNodeValue(state, node, value) {
    if (!(node instanceof Element)) return;
    node.textContent = Number.isFinite(value) ? ` ${formatGradeSummaryNumber(value)}` : " --";
    const base = state.summaryOriginalValues.get(node);
    const changed = Number.isFinite(base) ? !isSameNumericValue(base, value) : Number.isFinite(value);
    node.classList.toggle("tdk-grade-summary-changed", changed);
  }

  function setPassCellValue(cell, passValue) {
    if (!cell) return;
    let marker = cell.querySelector("div.check, div.no-check");
    if (!marker) {
      cell.innerHTML = `<div><div class="${passValue ? "check" : "no-check"}"></div></div>`;
      return;
    }
    marker.classList.remove("check", "no-check");
    marker.classList.add(passValue ? "check" : "no-check");
  }

  function setCellDisplayValue(cell, value) {
    if (!(cell instanceof HTMLTableCellElement)) return;
    if (cell.dataset.tdkEditing === "1") return;
    cell.textContent = Number.isFinite(value) ? formatGradeNumberForDisplay(value) : "";
  }

  function isSameTextValue(left, right) {
    return norm(left) === norm(right);
  }

  function updateRecordComponentValue(record, titleValue, scoreValue) {
    const title = text(titleValue);
    if (!record.edit.componentFieldValues || typeof record.edit.componentFieldValues !== "object") {
      record.edit.componentFieldValues = { ...(record.baseComponentFieldValues || buildComponentFieldValuesFromRecord(record)) };
    }
    if (title === "DiemThi" || title.startsWith("DiemChuyenCan") || title.startsWith("DiemThuongKy") || title.startsWith("DiemHeSo") || title.startsWith("DiemThucHanh")) {
      record.edit.componentFieldValues[title] = Number.isFinite(scoreValue) ? scoreValue : null;
      recomputeComponentEditFromFieldValues(record.edit);
      const derived = deriveGradeFromComponentEdits(record.edit.componentEdit);
      if (Number.isFinite(derived.total10Calc)) {
        record.edit.total10Edit = normalizeScore10(derived.total10Calc);
        record.edit.lastEditedSource = "component";
        record.predicted = buildGradePrediction(record);
        record.predicted.tbThuongKyCalc = derived.tbThuongKyCalc;
      } else {
        record.edit.lastEditedSource = "component_pending";
        record.predicted = buildGradePrediction(record);
        if (Number.isFinite(derived.tbThuongKyCalc)) {
          record.predicted.tbThuongKyCalc = derived.tbThuongKyCalc;
        }
      }
      return;
    }
    if (GRADE_TOTAL10_TITLES.includes(title)) {
      if (Number.isFinite(scoreValue)) {
        record.edit.total10Edit = scoreValue;
        syncPredictionFromQuick(record);
      } else {
        record.edit.total10Edit = Number.isFinite(record.total10)
          ? normalizeScore10(record.total10)
          : Number.isFinite(record.rawTotal10)
            ? normalizeScore10(record.rawTotal10)
            : null;
        syncPredictionFromQuick(record);
      }
    }
  }

  function buildInlineBindingsForState(state, table) {
    const bindings = new Map();
    const cellToMeta = new WeakMap();
    const rows = table ? Array.from(table.querySelectorAll("tbody tr")) : [];
    const records = Array.isArray(state.records) ? state.records : [];
    let index = 0;

    for (const row of rows) {
      const headingCell = row.querySelector("td.row-head");
      if (headingCell) continue;

      const cells = Array.from(row.cells || []);
      if (!cells.length) continue;
      const fieldMap = getGradeFieldCellMap(row);
      const classCode = text(cells[1] && cells[1].textContent);
      const courseName = text(cells[2] && cells[2].textContent);
      if (!isGradeRecordCandidate(row, classCode, courseName, fieldMap)) continue;
      if (index >= records.length) break;

      let record = records[index];
      if (!record || text(record.classCode) !== classCode || text(record.courseName) !== courseName) {
        const fallback = records.find((item) => !bindings.has(item.recordKey) && text(item.classCode) === classCode && text(item.courseName) === courseName);
        if (fallback) {
          record = fallback;
        }
      }
      if (!record) continue;
      index += 1;

      const rowBinding = {
        row,
        cellsByTitle: {},
        tbCell: firstGradeFieldCell(fieldMap, "DiemTBThuongKy"),
        total4Cell: firstGradeFieldCell(fieldMap, "DiemTinChi"),
        letterCell: firstGradeFieldCell(fieldMap, "DiemChu"),
        xepLoaiCell: firstGradeFieldCell(fieldMap, "XepLoai"),
        passCell: firstGradeFieldCell(fieldMap, ["IsDat", "Dat"])
      };

      if (!rowBinding.passCell) {
        for (let i = cells.length - 1; i >= 0; i -= 1) {
          if (cells[i].querySelector("div.check, div.no-check, input[type='checkbox']")) {
            rowBinding.passCell = cells[i];
            break;
          }
        }
      }

      for (const title of GRADE_EDITABLE_TITLES) {
        const cell = firstGradeFieldCell(fieldMap, title);
        if (!cell) continue;
        rowBinding.cellsByTitle[title] = cell;
        cell.classList.add("tdk-grade-inline-editable");
        cellToMeta.set(cell, {
          recordKey: record.recordKey,
          title
        });
      }

      if (rowBinding.tbCell) cellToMeta.set(rowBinding.tbCell, { recordKey: record.recordKey, title: "DiemTBThuongKy" });
      if (rowBinding.total4Cell) cellToMeta.set(rowBinding.total4Cell, { recordKey: record.recordKey, title: "DiemTinChi" });
      if (rowBinding.letterCell) cellToMeta.set(rowBinding.letterCell, { recordKey: record.recordKey, title: "DiemChu" });
      if (rowBinding.xepLoaiCell) cellToMeta.set(rowBinding.xepLoaiCell, { recordKey: record.recordKey, title: "XepLoai" });
      if (rowBinding.passCell) cellToMeta.set(rowBinding.passCell, { recordKey: record.recordKey, title: "Dat" });

      row.setAttribute("data-tdk-record-key", record.recordKey);
      bindings.set(record.recordKey, rowBinding);
    }

    state.inlineBindings = bindings;
    state.cellToMeta = cellToMeta;
    state.summaryRefs = collectSummaryRefsFromTable(table, state);
    state.tableRef = table;

    const validKeys = new Set(bindings.keys());
    const selectedKeys = new Set();
    if (state.selectedRecordKeys instanceof Set) {
      for (const key of state.selectedRecordKeys) {
        const normalizedKey = text(key);
        if (normalizedKey && validKeys.has(normalizedKey)) {
          selectedKeys.add(normalizedKey);
        }
      }
    }
    state.selectedRecordKeys = selectedKeys;

    if (!text(state.selectedRecordKey) || !validKeys.has(text(state.selectedRecordKey))) {
      const firstRecord = state.records && state.records.length ? state.records[0] : null;
      state.selectedRecordKey = firstRecord ? firstRecord.recordKey : "";
    }
    if (text(state.selectedRecordKey)) {
      state.selectedRecordKeys.add(text(state.selectedRecordKey));
    }
  }

  function refreshSelectedRowVisual(state) {
    if (!(state && state.inlineBindings instanceof Map)) return;
    if (!(state.selectedRecordKeys instanceof Set)) {
      state.selectedRecordKeys = new Set(text(state.selectedRecordKey) ? [text(state.selectedRecordKey)] : []);
    }
    for (const [recordKey, binding] of state.inlineBindings.entries()) {
      if (!binding || !(binding.row instanceof HTMLTableRowElement)) continue;
      const isSelected = state.selectedRecordKeys.has(text(recordKey));
      binding.row.classList.toggle("tdk-grade-row-selected", isSelected);
    }
  }

  function setSelectedRecord(state, recordKey, options = {}) {
    if (!state) return;
    const key = text(recordKey);
    if (!key) return;

    if (!(state.selectedRecordKeys instanceof Set)) {
      state.selectedRecordKeys = new Set();
    }

    const additive = options && options.additive === true;
    const toggle = options && options.toggle === true;
    if (!additive) {
      state.selectedRecordKeys.clear();
    }

    if (toggle && state.selectedRecordKeys.has(key)) {
      state.selectedRecordKeys.delete(key);
      if (text(state.selectedRecordKey) === key) {
        const fallback = state.selectedRecordKeys.values().next();
        state.selectedRecordKey = fallback.done ? "" : fallback.value;
      }
    } else {
      state.selectedRecordKeys.add(key);
      state.selectedRecordKey = key;
    }

    if (!state.selectedRecordKeys.size && state.inlineBindings instanceof Map && state.inlineBindings.size) {
      const fallbackKey = state.inlineBindings.keys().next().value;
      if (fallbackKey) {
        state.selectedRecordKeys.add(text(fallbackKey));
        state.selectedRecordKey = text(fallbackKey);
      }
    }
    refreshSelectedRowVisual(state);
  }

  function applyRecordRowVisual(state, record) {
    if (!(state.inlineBindings instanceof Map)) return;
    const binding = state.inlineBindings.get(record.recordKey);
    if (!binding) return;

    const componentValues = record.edit && record.edit.componentFieldValues ? record.edit.componentFieldValues : {};
    const baseComponentValues = record.baseComponentFieldValues || buildComponentFieldValuesFromRecord(record);
    let rowChanged = false;

    for (const title of GRADE_COMPONENT_TITLES) {
      const cell = binding.cellsByTitle[title];
      if (!cell) continue;
      const value = componentValues[title];
      const changed = !isSameNumericValue(value, baseComponentValues[title]);
      rowChanged = rowChanged || changed;
      setCellDisplayValue(cell, value);
      cell.classList.toggle("tdk-grade-cell-changed", changed);
    }

    const baseTotal10Display = Number.isFinite(record.total10)
      ? record.total10
      : Number.isFinite(record.rawTotal10)
        ? normalizeGradeNumberValue(record.rawTotal10)
        : null;
    const total10 = record.predicted && Number.isFinite(record.predicted.total10Edit)
      ? record.predicted.total10Edit
      : baseTotal10Display;
    const totalChanged = !isSameNumericValue(total10, baseTotal10Display);
    rowChanged = rowChanged || totalChanged;
    for (const title of GRADE_TOTAL10_TITLES) {
      const cell = binding.cellsByTitle[title];
      if (!cell) continue;
      setCellDisplayValue(cell, total10);
      cell.classList.toggle("tdk-grade-cell-changed", totalChanged);
    }

    if (binding.tbCell) {
      const useCalculatedTb = record.edit && record.edit.lastEditedSource === "component";
      const tbValue = useCalculatedTb && Number.isFinite(record.predicted && record.predicted.tbThuongKyCalc)
        ? record.predicted.tbThuongKyCalc
        : record.components.tbThuongKy;
      const tbChanged = useCalculatedTb
        ? !isSameNumericValue(tbValue, record.components.tbThuongKy)
        : false;
      rowChanged = rowChanged || tbChanged;
      setCellDisplayValue(binding.tbCell, tbValue);
      binding.tbCell.classList.toggle("tdk-grade-cell-changed", tbChanged);
    }

    if (binding.total4Cell) {
      const total4Value = Number.isFinite(record.predicted && record.predicted.total4Calc)
        ? record.predicted.total4Calc
        : record.total4;
      const changed = !isSameNumericValue(total4Value, record.total4);
      rowChanged = rowChanged || changed;
      setCellDisplayValue(binding.total4Cell, total4Value);
      binding.total4Cell.classList.toggle("tdk-grade-cell-changed", changed);
    }

    if (binding.letterCell) {
      const letterValue = text(record.predicted && record.predicted.letterCalc ? record.predicted.letterCalc : record.letter);
      const changed = !isSameTextValue(letterValue, record.letter);
      rowChanged = rowChanged || changed;
      binding.letterCell.textContent = letterValue;
      binding.letterCell.classList.toggle("tdk-grade-cell-changed", changed);
    }

    if (binding.xepLoaiCell) {
      const xepLoaiValue = text(record.predicted && record.predicted.xepLoaiCalc ? record.predicted.xepLoaiCalc : record.xepLoai);
      const changed = !isSameTextValue(xepLoaiValue, record.xepLoai);
      rowChanged = rowChanged || changed;
      binding.xepLoaiCell.textContent = xepLoaiValue;
      binding.xepLoaiCell.classList.toggle("tdk-grade-cell-changed", changed);
    }

    if (binding.passCell) {
      const usePredictedPass = record.predicted && record.predicted.includeInGPA === true;
      const passValue = usePredictedPass ? record.predicted.passPredicted === true : record.pass === true;
      const changed = usePredictedPass ? passValue !== (record.pass === true) : false;
      rowChanged = rowChanged || changed;
      setPassCellValue(binding.passCell, passValue);
      binding.passCell.classList.toggle("tdk-grade-cell-changed", changed);
    }

    const usePredictedPass = record.predicted && record.predicted.includeInGPA === true;
    const passVisual = usePredictedPass ? record.predicted.passPredicted === true : record.pass === true;
    const shouldMarkRed = passVisual === false;
    for (const title of GRADE_TOTAL10_TITLES) {
      const cell = binding.cellsByTitle[title];
      if (cell) cell.classList.toggle("cl-red", shouldMarkRed);
    }
    if (binding.total4Cell) binding.total4Cell.classList.toggle("cl-red", shouldMarkRed);
    if (binding.letterCell) binding.letterCell.classList.toggle("cl-red", shouldMarkRed);
    if (binding.xepLoaiCell) binding.xepLoaiCell.classList.toggle("cl-red", shouldMarkRed);

    binding.row.classList.toggle("tdk-grade-row-changed", rowChanged);
  }

  function buildCumulativeByTermForVisual(records, orderedTerms) {
    const sourceRecords = Array.isArray(records) ? records : [];
    const terms = Array.isArray(orderedTerms) ? orderedTerms : [];
    const map = new Map();
    const running = {
      current: {
        earnedCredits: 0,
        debtCredits: 0,
        weighted4: 0,
        weighted10: 0
      },
      predicted: {
        earnedCredits: 0,
        debtCredits: 0,
        weighted4: 0,
        weighted10: 0
      }
    };

    for (const term of terms) {
      const key = buildGradeTermKey(term.yearLabel, term.termLabel);
      const termRecords = sourceRecords.filter((record) => buildGradeTermKey(record.yearLabel, record.termLabel) === key);
      for (const record of termRecords) {
        const credits = Number(record.credits);
        if (!(Number.isFinite(credits) && credits > 0)) continue;

        if (record.includeInGPA === true) {
          if (record.pass === true && Number.isFinite(record.total4) && Number.isFinite(record.total10)) {
            running.current.earnedCredits += credits;
            running.current.weighted4 += credits * record.total4;
            running.current.weighted10 += credits * record.total10;
          } else if (record.pass !== true) {
            running.current.debtCredits += credits;
          }
        }

        const hasPrediction = Boolean(record.predicted && record.predicted.includeInGPA === true);
        if (hasPrediction) {
          const passPred = record.predicted.passPredicted === true;
          const total4Pred = Number.isFinite(record.predicted.total4Calc) ? record.predicted.total4Calc : null;
          const total10Pred = Number.isFinite(record.predicted.total10Edit) ? record.predicted.total10Edit : null;
          if (passPred && Number.isFinite(total4Pred) && Number.isFinite(total10Pred)) {
            running.predicted.earnedCredits += credits;
            running.predicted.weighted4 += credits * total4Pred;
            running.predicted.weighted10 += credits * total10Pred;
          } else if (!passPred) {
            running.predicted.debtCredits += credits;
          }
        }
      }

      const currentGpa4 = running.current.earnedCredits > 0 ? running.current.weighted4 / running.current.earnedCredits : null;
      const currentGpa10 = running.current.earnedCredits > 0 ? running.current.weighted10 / running.current.earnedCredits : null;
      const predictedGpa4 = running.predicted.earnedCredits > 0 ? running.predicted.weighted4 / running.predicted.earnedCredits : null;
      const predictedGpa10 = running.predicted.earnedCredits > 0 ? running.predicted.weighted10 / running.predicted.earnedCredits : null;
      map.set(key, {
        current: {
          gpa4: normalizeGradeNumberValue(currentGpa4),
          gpa10: normalizeGradeNumberValue(currentGpa10),
          earnedCredits: normalizeGradeNumberValue(running.current.earnedCredits),
          debtCredits: normalizeGradeNumberValue(running.current.debtCredits)
        },
        predicted: {
          gpa4: normalizeGradeNumberValue(predictedGpa4),
          gpa10: normalizeGradeNumberValue(predictedGpa10),
          earnedCredits: normalizeGradeNumberValue(running.predicted.earnedCredits),
          debtCredits: normalizeGradeNumberValue(running.predicted.debtCredits)
        }
      });
    }
    return map;
  }

  function applySummaryVisual(state) {
    if (!state.summaryRefs) return;
    const summary = calculateGradeSummary(state.records, state.webSummary);
    state.summary = summary;
    const cumulativeByTerm = buildCumulativeByTermForVisual(state.records, summary.terms);

    for (const term of summary.terms) {
      const key = buildGradeTermKey(term.yearLabel, term.termLabel);
      const refs = state.summaryRefs.terms.get(key);
      if (!refs) continue;
      for (const node of refs.gpa10) {
        setSummaryNodeValue(state, node, term.predicted.gpa10);
      }
      for (const node of refs.gpa4) {
        setSummaryNodeValue(state, node, term.predicted.gpa4);
      }

      const cumulativeRefs = state.summaryRefs.cumulativeByTerm.get(key);
      const cumulativeTerm = cumulativeByTerm.get(key);
      if (cumulativeRefs && cumulativeTerm) {
        for (const node of cumulativeRefs.gpa10) {
          setSummaryNodeValue(state, node, cumulativeTerm.predicted.gpa10);
        }
        for (const node of cumulativeRefs.gpa4) {
          setSummaryNodeValue(state, node, cumulativeTerm.predicted.gpa4);
        }
        for (const node of cumulativeRefs.credits) {
          setSummaryNodeValue(state, node, cumulativeTerm.predicted.earnedCredits);
        }
        for (const node of cumulativeRefs.debt) {
          setSummaryNodeValue(state, node, cumulativeTerm.predicted.debtCredits);
        }
      }
    }

    // Fallback cho trang chỉ có 1 cụm tích lũy dùng chung
    for (const node of state.summaryRefs.cumulative.gpa10) {
      setSummaryNodeValue(state, node, summary.cumulative.predicted.gpa10);
    }
    for (const node of state.summaryRefs.cumulative.gpa4) {
      setSummaryNodeValue(state, node, summary.cumulative.predicted.gpa4);
    }
    for (const node of state.summaryRefs.cumulative.credits) {
      setSummaryNodeValue(state, node, summary.cumulative.predicted.passedCredits);
    }
    for (const node of state.summaryRefs.cumulative.debt) {
      setSummaryNodeValue(state, node, summary.cumulative.predicted.debtCredits);
    }
  }

  function renderGradeInlineState(state) {
    if (!state || !Array.isArray(state.records)) return;
    for (const record of state.records) {
      applyRecordRowVisual(state, record);
    }
    refreshSelectedRowVisual(state);
    applySummaryVisual(state);
  }

  function ensureGradeInlineState(forceRefresh = false) {
    const state = ensureGradeEditorState(forceRefresh);
    if (!state) return null;
    if (!(state.summaryOriginalValues instanceof WeakMap)) {
      state.summaryOriginalValues = new WeakMap();
    }
    if (typeof state.inlineEnabled !== "boolean") {
      state.inlineEnabled = false;
    }

    const table = getGradeTable();
    if (!table) return state;
    if (forceRefresh || state.tableRef !== table || !(state.inlineBindings instanceof Map) || !state.inlineBindings.size) {
      buildInlineBindingsForState(state, table);
      renderGradeInlineState(state);
    }

    return state;
  }

  function closeActiveInlineEditor(state, commitValue) {
    if (!state || !state.activeEditor) return;
    const active = state.activeEditor;
    state.activeEditor = null;

    const raw = text(active.input.value);
    if (commitValue === false) {
      delete active.cell.dataset.tdkEditing;
      renderGradeInlineState(state);
      return;
    }

    if (!raw) {
      updateRecordComponentValue(active.record, active.title, null);
      delete active.cell.dataset.tdkEditing;
      renderGradeInlineState(state);
      return;
    }

    const parsed = toScoreOrNull(raw);
    if (!Number.isFinite(parsed)) {
      setGradeInlineStatus("Điểm không hợp lệ. Chỉ nhận số 0 - 10.", "error");
      delete active.cell.dataset.tdkEditing;
      renderGradeInlineState(state);
      return;
    }

    updateRecordComponentValue(active.record, active.title, parsed);
    delete active.cell.dataset.tdkEditing;
    renderGradeInlineState(state);
    if (GRADE_COMPONENT_TITLES.includes(active.title)) {
      setGradeInlineStatus("Đã cập nhật điểm và GPA tự động. Nút 'Tính TB thường kỳ' chỉ hỗ trợ tự tính TB cho môn đã chọn.", "success");
    } else {
      setGradeInlineStatus("Đã cập nhật điểm dự đoán trên bảng hiện tại.", "success");
    }
  }

  function openInlineEditorForCell(state, cell, record, title) {
    if (!(cell instanceof HTMLTableCellElement) || !record) return;
    closeActiveInlineEditor(state, true);

    const input = document.createElement("input");
    input.type = "text";
    input.className = "tdk-grade-inline-input";
    const currentValue = GRADE_TOTAL10_TITLES.includes(title)
      ? record.predicted && Number.isFinite(record.predicted.total10Edit) ? record.predicted.total10Edit : null
      : record.edit && record.edit.componentFieldValues ? record.edit.componentFieldValues[title] : null;
    input.value = Number.isFinite(currentValue) ? String(currentValue).replace(".", ",") : "";

    cell.dataset.tdkEditing = "1";
    cell.innerHTML = "";
    cell.appendChild(input);
    input.focus();
    input.select();

    state.activeEditor = {
      cell,
      input,
      record,
      title
    };

    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        closeActiveInlineEditor(state, true);
      } else if (event.key === "Escape") {
        event.preventDefault();
        closeActiveInlineEditor(state, false);
      }
    });
    input.addEventListener("blur", () => {
      closeActiveInlineEditor(state, true);
    });
  }

  function applyComponentCalculation(record) {
    if (!record || !record.edit) return false;
    recomputeComponentEditFromFieldValues(record.edit);
    const derived = deriveGradeFromComponentEdits(record.edit.componentEdit);
    if (!Number.isFinite(derived.tbThuongKyCalc)) {
      return false;
    }

    // Nút này chỉ hỗ trợ tự tính TB thường kỳ, không ghi đè Total10.
    record.edit.lastEditedSource = "component";
    record.predicted = buildGradePrediction(record);
    record.predicted.tbThuongKyCalc = derived.tbThuongKyCalc;
    return true;
  }

  function applyComponentCalculationByScope(scope) {
    const state = ensureGradeInlineState(false);
    if (!state) return;
    closeActiveInlineEditor(state, true);

    const scopeKey = text(scope).toLowerCase();
    const normalizedScope = scopeKey === "selected_set" ? "selected_set" : "selected";
    const targets = [];
    if (normalizedScope === "selected_set") {
      if (state.selectedRecordKeys instanceof Set && state.selectedRecordKeys.size) {
        for (const key of state.selectedRecordKeys) {
          const target = state.records.find((item) => item.recordKey === key);
          if (target) {
            targets.push(target);
          }
        }
      } else {
        const fallback = state.records.find((item) => item.recordKey === state.selectedRecordKey);
        if (fallback) {
          targets.push(fallback);
        }
      }
    } else {
      const selected = state.records.find((item) => item.recordKey === state.selectedRecordKey);
      if (selected) {
        targets.push(selected);
      }
    }

    if (!targets.length) {
      setGradeInlineStatus("Chưa chọn môn để tính. Hãy click dòng môn (giữ Ctrl/Cmd để chọn nhiều).", "error");
      return;
    }

    let appliedCount = 0;
    for (const record of targets) {
      if (applyComponentCalculation(record)) {
        appliedCount += 1;
      }
    }

    renderGradeInlineState(state);
    if (appliedCount === 0) {
      setGradeInlineStatus("Không đủ dữ liệu thành phần để tự tính TB thường kỳ.", "error");
      return;
    }

    const scopeText = normalizedScope === "selected_set" ? "các môn đã chọn" : "môn đang chọn";
    setGradeInlineStatus(`Đã tự tính TB thường kỳ cho ${scopeText} (${appliedCount} môn).`, "success");
  }

  function refreshGradeInlineControlState(state) {
    const toggleBtn = document.getElementById("tdk-grade-open-btn");
    if (!toggleBtn) return;
    const enabled = Boolean(state && state.inlineEnabled);
    toggleBtn.textContent = enabled ? "Tắt sửa điểm inline" : "Bật sửa điểm inline";
    toggleBtn.classList.toggle("tdk-active", enabled);

    const table = getGradeTable();
    if (table) {
      table.classList.toggle("tdk-grade-inline-enabled", enabled);
    }
  }

  function bindGradeInlineHandlers() {
    if (gradeInlineHandlersBound) return;
    gradeInlineHandlersBound = true;

    document.addEventListener(
      "click",
      (event) => {
        if (!isGradePage()) return;
        const state = ensureGradeInlineState(false);
        if (!state || !state.inlineEnabled) return;

        const target = event.target;
        if (!(target instanceof Element)) return;
        if (target.id === "tdk-grade-open-btn" || target.id === "tdk-grade-reset-all-btn") return;
        if (target.closest(".tdk-grade-inline-input")) return;

        const cell = target.closest("td");
        if (!(cell instanceof HTMLTableCellElement)) return;
        if (!cell.closest("table#xemDiem_aaa")) return;
        if (cell.dataset.tdkEditing === "1") return;
        const isMultiSelect = event.ctrlKey === true || event.metaKey === true;

        const selectedRow = cell.closest("tr[data-tdk-record-key]");
        if (selectedRow) {
          const selectedKey = text(selectedRow.getAttribute("data-tdk-record-key"));
          if (selectedKey) {
            setSelectedRecord(
              state,
              selectedKey,
              isMultiSelect
                ? { additive: true, toggle: true }
                : undefined
            );
          }
        }
        if (isMultiSelect) {
          event.preventDefault();
          event.stopPropagation();
          return;
        }

        const meta = state.cellToMeta instanceof WeakMap ? state.cellToMeta.get(cell) : null;
        if (!meta || !isGradeEditableTitle(meta.title)) return;

        const record = state.records.find((item) => item.recordKey === meta.recordKey);
        if (!record) return;
        setSelectedRecord(state, record.recordKey);

        event.preventDefault();
        event.stopPropagation();
        openInlineEditorForCell(state, cell, record, meta.title);
      },
      true
    );
  }

  function resetAllInlineGradeEdits() {
    const state = ensureGradeInlineState(false);
    if (!state) return;
    closeActiveInlineEditor(state, false);
    for (const record of state.records) {
      resetGradeRecord(record);
    }
    renderGradeInlineState(state);
    setGradeInlineStatus("Đã khôi phục tất cả điểm về mặc định từ web.", "success");
  }

  function toggleInlineGradeEditMode() {
    const state = ensureGradeInlineState(false);
    if (!state) {
      setGradeInlineStatus("Không đọc được dữ liệu bảng điểm.", "error");
      return;
    }
    state.inlineEnabled = !state.inlineEnabled;
    if (!state.inlineEnabled) {
      closeActiveInlineEditor(state, true);
    }
    refreshGradeInlineControlState(state);
    setGradeInlineStatus(
      state.inlineEnabled
        ? "Đã bật sửa inline. Sửa điểm đến đâu GPA cập nhật đến đó; có thể giữ Ctrl/Cmd + click rồi bấm 'Tính TB thường kỳ'."
        : "Đã tắt sửa inline.",
      state.inlineEnabled ? "success" : ""
    );
  }

  function ensureGradeManageButton() {
    if (!isGradePage()) return;
    ensureGradeInlineStyle();
    bindGradeInlineHandlers();

    const host = getGradeActionsHost();
    if (!host) return;

    const legacyWrap = document.getElementById("tdk-grade-wrap");
    if (legacyWrap && legacyWrap.parentElement) {
      legacyWrap.parentElement.removeChild(legacyWrap);
    }
    const legacyButton = document.getElementById("tdk-grade-open-btn");
    if (legacyButton && !legacyButton.closest("#tdk-grade-inline-controls")) {
      legacyButton.remove();
    }

    let controls = document.getElementById("tdk-grade-inline-controls");
    if (!controls) {
      controls = document.createElement("div");
      controls.id = "tdk-grade-inline-controls";
      controls.innerHTML =
        `<a id="tdk-grade-open-btn" href="javascript:;" class="btn btn-action">Bật sửa điểm inline</a>` +
        `<select id="tdk-grade-calc-scope"><option value="selected_set">Các môn đã chọn</option></select>` +
        `<a id="tdk-grade-calc-btn" href="javascript:;" class="btn btn-default">Tính TB thường kỳ</a>` +
        `<a id="tdk-grade-reset-all-btn" href="javascript:;" class="btn btn-default">Khôi phục tất cả</a>` +
        `<span id="tdk-grade-inline-status"></span>`;
      host.appendChild(controls);

      const toggleBtn = controls.querySelector("#tdk-grade-open-btn");
      const calcBtn = controls.querySelector("#tdk-grade-calc-btn");
      const calcScope = controls.querySelector("#tdk-grade-calc-scope");
      const resetBtn = controls.querySelector("#tdk-grade-reset-all-btn");
      if (toggleBtn) {
        toggleBtn.addEventListener("click", (event) => {
          event.preventDefault();
          toggleInlineGradeEditMode();
        });
      }
      if (calcBtn) {
        calcBtn.addEventListener("click", (event) => {
          event.preventDefault();
          const scope = calcScope instanceof HTMLSelectElement ? calcScope.value : "selected_set";
          applyComponentCalculationByScope(scope);
        });
      }
      if (resetBtn) {
        resetBtn.addEventListener("click", (event) => {
          event.preventDefault();
          resetAllInlineGradeEdits();
        });
      }
    }

    const state = ensureGradeInlineState(false);
    refreshGradeInlineControlState(state);
  }

  function ensureGradeEditorPanel() {
    // Giữ hàm này để tương thích nhánh cũ, nhưng không tạo bảng editor riêng.
  }

  function ensureGradeEditorStyle() {
    if (document.getElementById("tdk-grade-style")) return;
    const style = document.createElement("style");
    style.id = "tdk-grade-style";
    style.textContent =
      "#tdk-grade-open-btn{margin-left:8px}" +
      "#tdk-grade-wrap{position:fixed;inset:0;background:rgba(15,23,42,.45);display:none;z-index:2147483642}" +
      "#tdk-grade-modal{width:min(1200px,96vw);max-height:92vh;overflow:auto;margin:3vh auto;background:#fff;border-radius:10px;padding:12px}" +
      ".tdk-grade-top{display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap}" +
      ".tdk-grade-top h3{margin:0;font-size:16px}" +
      ".tdk-grade-top .tdk-grade-ctrl{display:flex;gap:8px;align-items:center;flex-wrap:wrap}" +
      ".tdk-grade-top select,.tdk-grade-top button{height:32px;border:1px solid #cbd5e1;border-radius:8px;padding:0 8px;background:#fff}" +
      ".tdk-grade-top button.tdk-primary{background:#2563eb;border-color:#2563eb;color:#fff}" +
      ".tdk-grade-status{margin:8px 0 0;font-size:12px;color:#334155;min-height:18px}" +
      ".tdk-grade-status.error{color:#b91c1c}" +
      ".tdk-grade-status.success{color:#065f46}" +
      ".tdk-grade-summary{margin-top:8px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;background:#f8fafc;font-size:12px}" +
      ".tdk-grade-summary table{width:100%;border-collapse:collapse}" +
      ".tdk-grade-summary th,.tdk-grade-summary td{border:1px solid #dbe4ef;padding:4px;text-align:center}" +
      ".tdk-grade-table-wrap{overflow:auto;margin-top:8px}" +
      "#tdk-grade-table{width:100%;border-collapse:collapse;font-size:12px}" +
      "#tdk-grade-table th,#tdk-grade-table td{border:1px solid #e2e8f0;padding:5px;vertical-align:middle}" +
      "#tdk-grade-table th{background:#f1f5f9;position:sticky;top:0;z-index:2}" +
      "#tdk-grade-table input{width:74px;height:28px;border:1px solid #cbd5e1;border-radius:6px;padding:0 6px}" +
      "#tdk-grade-table .tdk-cell-wide{min-width:280px;text-align:left}" +
      "#tdk-grade-table .tdk-row-changed{background:#fffbeb}" +
      "#tdk-grade-table .tdk-input-changed{border-color:#eab308;background:#fef9c3}" +
      "#tdk-grade-table .tdk-reset-row{height:28px;border:1px solid #cbd5e1;border-radius:6px;background:#fff;cursor:pointer}" +
      ".tdk-comp-cell{display:flex;gap:4px;flex-wrap:wrap;min-width:320px}" +
      ".tdk-comp-item{display:flex;align-items:center;gap:3px}" +
      ".tdk-comp-item label{font-size:11px;color:#475569}" +
      ".tdk-comp-item input{width:58px}";
    document.head.appendChild(style);
  }

  function gradeEditorElements() {
    return {
      wrap: document.getElementById("tdk-grade-wrap"),
      mode: document.getElementById("tdk-grade-mode"),
      summary: document.getElementById("tdk-grade-summary"),
      tbody: document.getElementById("tdk-grade-tbody"),
      status: document.getElementById("tdk-grade-status")
    };
  }

  function setGradeEditorStatus(message, type) {
    const el = gradeEditorElements();
    if (!el.status) return;
    el.status.textContent = message || "";
    el.status.className = "tdk-grade-status";
    if (type) el.status.classList.add(type);
  }

  function getGradeActionsHost() {
    return document.querySelector(".portlet .portlet-title .col-md-3") ||
      document.querySelector(".portlet .portlet-title .caption") ||
      document.querySelector(".portlet .actions");
  }

  function sanitizeEditorScore(value, max = 10) {
    const parsed = parseLocaleNumber(value);
    if (!Number.isFinite(parsed)) return null;
    if (parsed < 0 || parsed > max) return null;
    return normalizeScore10(parsed);
  }

  function syncPredictionFromQuick(record) {
    record.edit.lastEditedSource = "total10";
    record.predicted = buildGradePrediction(record);
  }

  function syncPredictionFromComponents(record) {
    recomputeComponentEditFromFieldValues(record.edit);
    const derived = deriveGradeFromComponentEdits(record.edit.componentEdit);
    record.edit.total10Edit = Number.isFinite(derived.total10Calc) ? derived.total10Calc : null;
    record.edit.lastEditedSource = "component";
    record.predicted = buildGradePrediction(record);
    record.predicted.tbThuongKyCalc = derived.tbThuongKyCalc;
  }

  function resetGradeRecord(record) {
    record.edit = createDefaultGradeEdit(record);
    record.predicted = buildGradePrediction(record);
  }

  function isSameNumericValue(a, b) {
    const left = Number.isFinite(a) ? Number(a) : null;
    const right = Number.isFinite(b) ? Number(b) : null;
    if (left == null && right == null) return true;
    if (left == null || right == null) return false;
    return Math.abs(left - right) < 0.001;
  }

  function formatGradeCellNumber(value) {
    return Number.isFinite(value) ? value.toFixed(2) : "";
  }

  function buildGradeEditorSummaryHtml(summary) {
    const termRows = summary.terms
      .map((term) => (
        "<tr>" +
          `<td>${esc(term.yearLabel)}</td>` +
          `<td>${esc(term.termLabel)}</td>` +
          `<td>${esc(formatGradeCellNumber(term.current.creditsGPA))}</td>` +
          `<td>${esc(formatGradeCellNumber(term.current.gpa4))}</td>` +
          `<td>${esc(formatGradeCellNumber(term.current.gpa10))}</td>` +
          `<td>${esc(formatGradeCellNumber(term.predicted.gpa4))}</td>` +
          `<td>${esc(formatGradeCellNumber(term.predicted.gpa10))}</td>` +
          `<td>${esc(formatGradeCellNumber(term.current.passedCredits))}</td>` +
          `<td>${esc(formatGradeCellNumber(term.current.debtCredits))}</td>` +
          `<td>${esc(formatGradeCellNumber(term.predicted.passedCredits))}</td>` +
          `<td>${esc(formatGradeCellNumber(term.predicted.debtCredits))}</td>` +
        "</tr>"
      ))
      .join("");

    const c = summary.cumulative;
    return (
      "<table>" +
        "<thead><tr>" +
          "<th>Year</th><th>Term</th><th>Credits</th><th>GPA4 Cur</th><th>GPA10 Cur</th><th>GPA4 Pred</th><th>GPA10 Pred</th><th>Đạt Cur</th><th>Nợ Cur</th><th>Đạt Pred</th><th>Nợ Pred</th>" +
        "</tr></thead>" +
        `<tbody>${termRows}</tbody>` +
      "</table>" +
      "<table style='margin-top:8px'>" +
        "<thead><tr><th>Chỉ số tích lũy</th><th>Current</th><th>Predicted</th><th>Web</th></tr></thead>" +
        "<tbody>" +
          `<tr><td>Tổng tín chỉ tính GPA</td><td>${esc(formatGradeCellNumber(c.current.creditsGPA))}</td><td>${esc(formatGradeCellNumber(c.predicted.creditsGPA))}</td><td>${esc(formatGradeCellNumber(c.web.accumulatedCredits))}</td></tr>` +
          `<tr><td>GPA hệ 4</td><td>${esc(formatGradeCellNumber(c.current.gpa4))}</td><td>${esc(formatGradeCellNumber(c.predicted.gpa4))}</td><td>${esc(formatGradeCellNumber(c.web.gpa4))}</td></tr>` +
          `<tr><td>GPA hệ 10</td><td>${esc(formatGradeCellNumber(c.current.gpa10))}</td><td>${esc(formatGradeCellNumber(c.predicted.gpa10))}</td><td>${esc(formatGradeCellNumber(c.web.gpa10))}</td></tr>` +
          `<tr><td>Tổng tín chỉ đạt</td><td>${esc(formatGradeCellNumber(c.current.passedCredits))}</td><td>${esc(formatGradeCellNumber(c.predicted.passedCredits))}</td><td></td></tr>` +
          `<tr><td>Số tín chỉ nợ</td><td>${esc(formatGradeCellNumber(c.current.debtCredits))}</td><td>${esc(formatGradeCellNumber(c.predicted.debtCredits))}</td><td>${esc(formatGradeCellNumber(c.web.debtCredits))}</td></tr>` +
        "</tbody>" +
      "</table>"
    );
  }

  function buildGradeEditorRowsHtml(state) {
    return state.records.map((record) => {
      const compEdit = record.edit.componentEdit;
      const totalChanged = !isSameNumericValue(record.predicted.total10Edit, record.total10);
      const componentChanged =
        !isSameNumericValue(compEdit.cc, averageFinite(record.components.chuyenCan)) ||
        !isSameNumericValue(compEdit.tx, averageFinite(record.components.thuongKy)) ||
        !isSameNumericValue(compEdit.hs1, averageFinite(record.components.heSo1)) ||
        !isSameNumericValue(compEdit.hs2, averageFinite(record.components.heSo2)) ||
        !isSameNumericValue(compEdit.th, averageFinite(record.components.thucHanh)) ||
        !isSameNumericValue(compEdit.exam, record.components.exam);

      const rowClass = totalChanged || componentChanged ? "tdk-row-changed" : "";
      const totalInputClass = totalChanged ? "tdk-input-changed" : "";
      const compInputClass = componentChanged ? "tdk-input-changed" : "";

      return (
        `<tr class="${rowClass}" data-key="${esc(record.recordKey)}">` +
          `<td>${esc(record.yearLabel)}</td>` +
          `<td>${esc(record.termLabel)}</td>` +
          `<td>${Number.isFinite(record.stt) ? record.stt : ""}</td>` +
          `<td>${esc(record.classCode)}</td>` +
          `<td class="tdk-cell-wide">${esc(record.courseName)}</td>` +
          `<td>${esc(formatGradeCellNumber(record.credits))}</td>` +
          `<td>${esc(formatGradeCellNumber(record.total10))}</td>` +
          `<td>${esc(formatGradeCellNumber(record.total4))}</td>` +
          `<td>${esc(record.pass ? "Đạt" : "Nợ")}</td>` +
          `<td><input class="tdk-grade-quick ${totalInputClass}" data-field="total10Edit" type="number" step="0.01" min="0" max="10" value="${esc(Number.isFinite(record.edit.total10Edit) ? record.edit.total10Edit : "")}"></td>` +
          `<td>${esc(formatGradeCellNumber(record.predicted.total4Calc))}</td>` +
          `<td>${esc(record.predicted.passPredicted ? "Đạt" : "Nợ")}</td>` +
          "<td class='tdk-comp-cell'>" +
            `<span class="tdk-comp-item"><label>CC</label><input class="${compInputClass}" data-field="cc" type="number" step="0.01" min="0" max="10" value="${esc(Number.isFinite(compEdit.cc) ? compEdit.cc : "")}"></span>` +
            `<span class="tdk-comp-item"><label>TX</label><input class="${compInputClass}" data-field="tx" type="number" step="0.01" min="0" max="10" value="${esc(Number.isFinite(compEdit.tx) ? compEdit.tx : "")}"></span>` +
            `<span class="tdk-comp-item"><label>HS1</label><input class="${compInputClass}" data-field="hs1" type="number" step="0.01" min="0" max="10" value="${esc(Number.isFinite(compEdit.hs1) ? compEdit.hs1 : "")}"></span>` +
            `<span class="tdk-comp-item"><label>HS2</label><input class="${compInputClass}" data-field="hs2" type="number" step="0.01" min="0" max="10" value="${esc(Number.isFinite(compEdit.hs2) ? compEdit.hs2 : "")}"></span>` +
            `<span class="tdk-comp-item"><label>TH</label><input class="${compInputClass}" data-field="th" type="number" step="0.01" min="0" max="10" value="${esc(Number.isFinite(compEdit.th) ? compEdit.th : "")}"></span>` +
            `<span class="tdk-comp-item"><label>Thi</label><input class="${compInputClass}" data-field="exam" type="number" step="0.01" min="0" max="10" value="${esc(Number.isFinite(compEdit.exam) ? compEdit.exam : "")}"></span>` +
          "</td>" +
          `<td>${esc(formatGradeCellNumber(record.predicted.tbThuongKyCalc))}</td>` +
          `<td><button class="tdk-reset-row" data-action="reset-row">Reset môn</button></td>` +
        "</tr>"
      );
    }).join("");
  }

  function renderGradeEditor() {
    const state = gradeEditorState;
    if (!state) return;
    const el = gradeEditorElements();
    if (!el.wrap || !el.mode || !el.summary || !el.tbody) return;

    el.mode.value = state.mode === "component" ? "component" : "quick";
    const summary = calculateGradeSummary(state.records, state.webSummary);
    el.summary.innerHTML = buildGradeEditorSummaryHtml(summary);
    el.tbody.innerHTML = buildGradeEditorRowsHtml(state);

    const isComponentMode = state.mode === "component";
    for (const input of Array.from(el.tbody.querySelectorAll("input.tdk-grade-quick"))) {
      input.disabled = isComponentMode;
    }
  }

  function ensureGradeEditorState(forceRefresh = false) {
    if (!isGradePage()) return null;
    if (!forceRefresh && gradeEditorState && gradeEditorState.pageUrl === location.href) {
      return gradeEditorState;
    }

    const scraped = scrapeGradeRecords();
    gradeEditorState = {
      pageUrl: location.href,
      mode: gradeEditorState && gradeEditorState.mode ? gradeEditorState.mode : "quick",
      records: hydrateRecordsForPrediction(scraped.records),
      webSummary: scraped.webSummary
    };
    return gradeEditorState;
  }

  function ensureGradeEditorPanel() {
    if (!isGradePage() || document.getElementById("tdk-grade-wrap")) return;
    ensureGradeEditorStyle();

    const wrap = document.createElement("div");
    wrap.id = "tdk-grade-wrap";
    wrap.innerHTML =
      `<div id="tdk-grade-modal">` +
        `<div class="tdk-grade-top">` +
          `<h3>Chỉnh sửa điểm / GPA dự đoán</h3>` +
          `<div class="tdk-grade-ctrl">` +
            `<label>Chế độ: <select id="tdk-grade-mode"><option value="quick">Nhanh (Total10)</option><option value="component">Thành phần</option></select></label>` +
            `<button id="tdk-grade-reset-all" type="button">Reset tất cả</button>` +
            `<button id="tdk-grade-close" class="tdk-primary" type="button">Đóng</button>` +
          `</div>` +
        `</div>` +
        `<p id="tdk-grade-status" class="tdk-grade-status"></p>` +
        `<div id="tdk-grade-summary" class="tdk-grade-summary"></div>` +
        `<div class="tdk-grade-table-wrap">` +
          `<table id="tdk-grade-table">` +
            `<thead><tr>` +
              `<th>Year</th><th>Term</th><th>STT</th><th>ClassCode</th><th>CourseName</th><th>Credits</th>` +
              `<th>Total10 Cur</th><th>Total4 Cur</th><th>Pass Cur</th><th>Total10 Edit</th><th>Total4 Calc</th><th>Pass Pred</th><th>Thành phần chỉnh sửa</th><th>TB TK Calc</th><th>Thao tác</th>` +
            `</tr></thead>` +
            `<tbody id="tdk-grade-tbody"></tbody>` +
          `</table>` +
        `</div>` +
      `</div>`;
    document.body.appendChild(wrap);

    const el = gradeEditorElements();
    if (!el.wrap || !el.mode || !el.tbody) return;

    wrap.addEventListener("click", (event) => {
      if (event.target === wrap) {
        wrap.style.display = "none";
      }
    });
    document.getElementById("tdk-grade-close").addEventListener("click", () => {
      wrap.style.display = "none";
    });
    document.getElementById("tdk-grade-reset-all").addEventListener("click", () => {
      const state = ensureGradeEditorState(false);
      if (!state) return;
      for (const record of state.records) {
        resetGradeRecord(record);
      }
      setGradeEditorStatus("Đã reset toàn bộ chỉnh sửa về mặc định.", "success");
      renderGradeEditor();
    });
    el.mode.addEventListener("change", () => {
      const state = ensureGradeEditorState(false);
      if (!state) return;
      state.mode = el.mode.value === "component" ? "component" : "quick";
      renderGradeEditor();
    });

    el.tbody.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) return;
      const button = target.closest("[data-action='reset-row']");
      if (!button) return;
      const row = button.closest("tr[data-key]");
      if (!row) return;
      const key = text(row.getAttribute("data-key"));
      const state = ensureGradeEditorState(false);
      if (!state) return;
      const record = state.records.find((item) => item.recordKey === key);
      if (!record) return;
      resetGradeRecord(record);
      setGradeEditorStatus("Đã reset môn về mặc định.", "success");
      renderGradeEditor();
    });

    el.tbody.addEventListener("input", (event) => {
      const input = event.target;
      if (!(input instanceof HTMLInputElement)) return;
      const row = input.closest("tr[data-key]");
      if (!row) return;
      const key = text(row.getAttribute("data-key"));
      const field = text(input.getAttribute("data-field"));
      if (!field) return;

      const state = ensureGradeEditorState(false);
      if (!state) return;
      const record = state.records.find((item) => item.recordKey === key);
      if (!record) return;

      if (field === "total10Edit") {
        record.edit.total10Edit = sanitizeEditorScore(input.value, 10);
        syncPredictionFromQuick(record);
      } else {
        record.edit.componentEdit[field] = sanitizeEditorScore(input.value, 10);
        syncPredictionFromComponents(record);
      }
      renderGradeEditor();
    });
  }

  async function openGradeEditor() {
    if (!isGradePage()) return;
    ensureGradeEditorPanel();
    const state = ensureGradeEditorState(false);
    if (!state) {
      setGradeEditorStatus("Không đọc được dữ liệu bảng điểm.", "error");
      return;
    }
    renderGradeEditor();
    const el = gradeEditorElements();
    if (el.wrap) {
      el.wrap.style.display = "block";
    }
    setGradeEditorStatus("Sửa điểm tại đây sẽ cập nhật GPA realtime và chỉ lưu trong tab hiện tại.", "");
  }

  function ensureGradeManageButton() {
    if (!isGradePage()) return;
    const host = getGradeActionsHost();
    if (!host || document.getElementById("tdk-grade-open-btn")) return;

    const button = document.createElement("a");
    button.id = "tdk-grade-open-btn";
    button.href = "javascript:;";
    button.className = "btn btn-action";
    button.innerHTML = '<i class="fa fa-calculator" aria-hidden="true"></i> Chỉnh sửa điểm / GPA';
    button.addEventListener("click", () => {
      void openGradeEditor();
    });
    host.appendChild(button);
  }

  // ===== Grade V2 (override) =====
  function getGradeExportContext() {
    const table = getGradeTable();
    const rows = table ? Array.from(table.querySelectorAll("tbody tr")) : [];
    const headingRows = rows.filter((row) => row.querySelector("td.row-head")).length;
    return {
      isGradePage: isGradePage(),
      hasTable: Boolean(table),
      totalRows: rows.length,
      headingRows,
      totalCourseRows: Math.max(0, rows.length - headingRows),
      pageUrl: location.href,
      hasEditorSession: Boolean(gradeEditorState && gradeEditorState.pageUrl === location.href)
    };
  }

  function normalizeScore10(value) {
    if (!Number.isFinite(value)) return null;
    if (value < 0 || value > 10) return null;
    return normalizeGradeNumberValue(value);
  }

  function normalizeScore4(value) {
    if (!Number.isFinite(value)) return null;
    if (value < 0 || value > 4) return null;
    return normalizeGradeNumberValue(value);
  }

  function roundToOneDecimal(value) {
    if (!Number.isFinite(value)) return null;
    return Math.round(value * 10) / 10;
  }

  function averageFinite(values) {
    const list = Array.isArray(values) ? values.filter((item) => Number.isFinite(item)) : [];
    if (!list.length) return null;
    const sum = list.reduce((acc, item) => acc + item, 0);
    return normalizeScore10(sum / list.length);
  }

  const GRADE_TOTAL10_TITLES = ["DiemTongKet", "DiemTongKet1"];
  const GRADE_COMPONENT_TITLES = [
    "DiemChuyenCan1",
    "DiemThuongKy1",
    "DiemThuongKy2",
    "DiemThuongKy3",
    "DiemHeSo11",
    "DiemHeSo12",
    "DiemHeSo13",
    "DiemHeSo14",
    "DiemHeSo15",
    "DiemHeSo16",
    "DiemHeSo17",
    "DiemHeSo18",
    "DiemHeSo19",
    "DiemHeSo21",
    "DiemHeSo22",
    "DiemHeSo23",
    "DiemHeSo24",
    "DiemHeSo25",
    "DiemHeSo26",
    "DiemHeSo27",
    "DiemHeSo28",
    "DiemHeSo29",
    "DiemThucHanh1",
    "DiemThucHanh2",
    "DiemThi"
  ];
  const GRADE_EDITABLE_TITLES = new Set([...GRADE_TOTAL10_TITLES, ...GRADE_COMPONENT_TITLES]);
  const GRADE_EXCLUDED_COURSE_PATTERNS = [
    /toeic|diem test tieng anh dau vao/i,
    /giao duc the chat|quoc phong|an ninh/i,
    /hoc phan dieu kien|dieu kien tot nghiep|ky nang mem/i
  ];

  function isGradeEditableTitle(titleValue) {
    return GRADE_EDITABLE_TITLES.has(text(titleValue));
  }

  function isExcludedCourseNameForGPA(courseName) {
    const courseNorm = norm(courseName);
    if (!courseNorm) return false;
    return GRADE_EXCLUDED_COURSE_PATTERNS.some((pattern) => pattern.test(courseNorm));
  }

  function isExcludedFromGPA(record, options = {}) {
    const source = record && typeof record === "object" ? record : {};
    const opts = options && typeof options === "object" ? options : {};
    const creditsValue = Number.isFinite(opts.credits) ? opts.credits : source.credits;
    const total4Value = Number.isFinite(opts.total4) ? opts.total4 : source.total4;
    const letterValue = text(typeof opts.letter === "string" ? opts.letter : source.letter);
    const courseNameValue = text(typeof opts.courseName === "string" ? opts.courseName : source.courseName);

    if (!(Number.isFinite(creditsValue) && creditsValue > 0)) return true;
    if (isExcludedCourseNameForGPA(courseNameValue)) return true;
    if (!letterValue && !Number.isFinite(total4Value)) return true;
    if (!Number.isFinite(total4Value)) return true;
    return false;
  }

  function formatGradeNumberForDisplay(value) {
    if (!Number.isFinite(value)) return "";
    return value.toFixed(2).replace(".", ",");
  }

  function formatGradeSummaryNumber(value) {
    if (!Number.isFinite(value)) return "";
    const rounded = Math.abs(value - Math.round(value)) < 0.001 ? String(Math.round(value)) : value.toFixed(2);
    return rounded.replace(".", ",");
  }

  function point4ToLetter(point4) {
    if (!Number.isFinite(point4)) return "";
    if (point4 >= 4.0) return "A";
    if (point4 >= 3.5) return "B+";
    if (point4 >= 3.0) return "B";
    if (point4 >= 2.5) return "C+";
    if (point4 >= 2.0) return "C";
    if (point4 >= 1.5) return "D+";
    if (point4 >= 1.0) return "D";
    if (point4 >= 0.5) return "F+";
    return "F";
  }

  function score10ToXepLoai(total10, letter) {
    const letterNorm = text(letter).toUpperCase();
    if (letterNorm.startsWith("A")) return "Giỏi";
    if (letterNorm.startsWith("B+")) return "Khá Giỏi";
    if (letterNorm.startsWith("B")) return "Khá";
    if (letterNorm.startsWith("C+")) return "Trung bình Khá";
    if (letterNorm.startsWith("C")) return "Trung bình";
    if (letterNorm.startsWith("D+")) return "Trung bình";
    if (letterNorm.startsWith("D")) return "Trung bình";
    if (letterNorm.startsWith("F")) return "Kém";

    if (!Number.isFinite(total10)) return "";
    if (total10 >= 8.5) return "Giỏi";
    if (total10 >= 8.0) return "Khá Giỏi";
    if (total10 >= 7.0) return "Khá";
    if (total10 >= 6.5) return "Trung bình Khá";
    if (total10 >= 5.0) return "Trung bình";
    return "Kém";
  }

  function isVisualRedColor(colorValue) {
    const raw = text(colorValue).toLowerCase();
    if (!raw) return false;
    if (raw.includes("red")) return true;
    const rgbMatch = raw.match(/rgba?\(([^)]+)\)/i);
    if (!rgbMatch) return false;
    const parts = rgbMatch[1].split(",").map((item) => Number.parseFloat(item.trim()));
    if (parts.length < 3) return false;
    const [r, g, b] = parts;
    if (![r, g, b].every((value) => Number.isFinite(value))) return false;
    return r >= 170 && g <= 110 && b <= 110;
  }

  function isFailVisualCell(cell) {
    if (!(cell instanceof Element)) return false;
    const classNorm = norm(cell.className);
    const textNorm = norm(cell.textContent);
    const styleInline = text(cell.getAttribute("style"));
    if (
      classNorm.includes("cl-red") ||
      classNorm.includes("text-danger") ||
      classNorm.includes("fail") ||
      styleInline.toLowerCase().includes("color:red")
    ) {
      return true;
    }
    if (textNorm.includes("khong dat") || textNorm.includes("rot")) {
      return true;
    }

    if (typeof window.getComputedStyle === "function") {
      try {
        const color = window.getComputedStyle(cell).color;
        if (isVisualRedColor(color)) {
          return true;
        }
      } catch (_error) {
        // ignore style read errors
      }
    }

    const highlighted = cell.querySelector(".cl-red, .text-danger, .fail");
    if (highlighted) return true;
    return false;
  }

  function hasFailVisualInCells(cells) {
    const list = Array.isArray(cells) ? cells : [];
    return list.some((cell) => isFailVisualCell(cell));
  }

  function toScoreOrNull(value) {
    const parsed = parseLocaleNumber(value);
    if (!Number.isFinite(parsed)) return null;
    if (parsed < 0 || parsed > 10) return null;
    return normalizeScore10(parsed);
  }

  function isGradeSummaryRow(rowNorm) {
    return (
      rowNorm.includes("diem trung binh hoc ky he 10") ||
      rowNorm.includes("diem trung binh hoc ky he 4") ||
      rowNorm.includes("diem trung binh tich luy") ||
      rowNorm.includes("tong so tin chi tich luy") ||
      rowNorm.includes("tong so tin chi no")
    );
  }

  function parseGradeCellTitle(cell) {
    if (!cell) return "";
    return text(
      cell.getAttribute("title") ||
      cell.getAttribute("data-original-title") ||
      cell.dataset.title ||
      ""
    );
  }

  function getGradeFieldCellMap(row) {
    const map = {};
    for (const cell of Array.from(row.cells || [])) {
      const titleValue = parseGradeCellTitle(cell);
      if (!titleValue) continue;
      if (!Array.isArray(map[titleValue])) {
        map[titleValue] = [];
      }
      map[titleValue].push(cell);
    }
    return map;
  }

  function firstGradeFieldCell(fieldMap, titleList) {
    const titles = Array.isArray(titleList) ? titleList : [titleList];
    for (const title of titles) {
      const key = text(title);
      const list = fieldMap[key];
      if (Array.isArray(list) && list.length) return list[0];

      const matchedKey = Object.keys(fieldMap).find((candidate) => candidate.toLowerCase() === key.toLowerCase());
      if (matchedKey && Array.isArray(fieldMap[matchedKey]) && fieldMap[matchedKey].length) {
        return fieldMap[matchedKey][0];
      }
    }
    return null;
  }

  function parseScore10FromCell(cell) {
    return normalizeScore10(parseLocaleNumber(cell && cell.textContent));
  }

  function parseScore4FromCell(cell) {
    return normalizeScore4(parseLocaleNumber(cell && cell.textContent));
  }

  function buildComponentList(fieldMap, prefix, from, to) {
    const output = [];
    for (let i = from; i <= to; i += 1) {
      output.push(parseScore10FromCell(firstGradeFieldCell(fieldMap, `${prefix}${i}`)));
    }
    return output;
  }

  function buildGradeComponentBundle(fieldMap) {
    const chuyenCan = [parseScore10FromCell(firstGradeFieldCell(fieldMap, "DiemChuyenCan1"))];
    const thuongKy = buildComponentList(fieldMap, "DiemThuongKy", 1, 3);
    const heSo1 = buildComponentList(fieldMap, "DiemHeSo1", 1, 9);
    const heSo2 = buildComponentList(fieldMap, "DiemHeSo2", 1, 9);
    const thucHanh = buildComponentList(fieldMap, "DiemThucHanh", 1, 2);

    const tbThuongKy = parseScore10FromCell(firstGradeFieldCell(fieldMap, "DiemTBThuongKy"));
    const exam = parseScore10FromCell(firstGradeFieldCell(fieldMap, "DiemThi"));
    const total10Lan1Raw = parseLocaleNumber(firstGradeFieldCell(fieldMap, "DiemTongKet1") && firstGradeFieldCell(fieldMap, "DiemTongKet1").textContent);
    const total10Raw = parseLocaleNumber(firstGradeFieldCell(fieldMap, "DiemTongKet") && firstGradeFieldCell(fieldMap, "DiemTongKet").textContent);

    return {
      chuyenCan,
      thuongKy,
      heSo1,
      heSo2,
      thucHanh,
      tbThuongKy,
      exam,
      total10Lan1: normalizeScore10(total10Lan1Raw),
      total10Raw: Number.isFinite(total10Raw) ? normalizeGradeNumberValue(total10Raw) : Number.isFinite(total10Lan1Raw) ? normalizeGradeNumberValue(total10Lan1Raw) : null,
      total10: normalizeScore10(Number.isFinite(total10Raw) ? total10Raw : total10Lan1Raw)
    };
  }

  function parsePassFromCell(cell, note, letter, total4, total10, xepLoai, hintCells) {
    if (cell) {
      const checkbox = cell.querySelector("input[type='checkbox']");
      if (checkbox) return checkbox.checked;
      if (cell.querySelector("div.no-check")) return false;
      if (cell.querySelector("div.check")) return true;

      const htmlNorm = norm(cell.innerHTML);
      const valueNorm = norm(cell.textContent);
      if (
        htmlNorm.includes("fa-times") ||
        htmlNorm.includes("fa-close") ||
        htmlNorm.includes("glyphicon-remove") ||
        valueNorm.includes("khong dat") ||
        valueNorm.includes("rot")
      ) {
        return false;
      }
      if (
        htmlNorm.includes("fa-check") ||
        htmlNorm.includes("glyphicon-ok") ||
        valueNorm === "x" ||
        valueNorm.includes("dat")
      ) {
        return true;
      }
    }

    const noteNorm = norm(note);
    const letterRaw = text(letter).toUpperCase();
    const xepLoaiNorm = norm(xepLoai);
    if (noteNorm.includes("thi lai")) return false;
    if (letterRaw.startsWith("F")) return false;
    if (xepLoaiNorm.includes("kem") || xepLoaiNorm.includes("khong dat") || xepLoaiNorm.includes("rot")) return false;
    if (hasFailVisualInCells([cell, ...(Array.isArray(hintCells) ? hintCells : [])])) return false;
    return true;
  }

  function parseWebSummaryFromRow(row, yearLabel, termLabel, termMap, cumulative) {
    const spans = Array.from(row.querySelectorAll("span[lang]"));
    if (!spans.length) return;

    for (const span of spans) {
      const lang = text(span.getAttribute("lang")).toLowerCase();
      const holder = span.nextElementSibling || span.parentElement;
      const value = extractLastNumberFromText(holder ? holder.textContent : span.textContent);
      if (!Number.isFinite(value)) continue;
      const normalized = normalizeGradeNumberValue(value);

      if (lang.includes("kqht-tkhk-diemtbhocluc") && !lang.includes("tichluy")) {
        const key = buildGradeTermKey(yearLabel, termLabel);
        const termSummary = termMap.get(key) || {
          yearLabel: text(yearLabel),
          termLabel: text(termLabel),
          gpa10: null,
          gpa4: null
        };
        termSummary.gpa10 = normalized;
        termMap.set(key, termSummary);
      } else if (lang.includes("kqht-tkhk-diemtbtinchi") && !lang.includes("tichluy")) {
        const key = buildGradeTermKey(yearLabel, termLabel);
        const termSummary = termMap.get(key) || {
          yearLabel: text(yearLabel),
          termLabel: text(termLabel),
          gpa10: null,
          gpa4: null
        };
        termSummary.gpa4 = normalized;
        termMap.set(key, termSummary);
      } else if (lang.includes("kqht-tkhk-diemtbhocluctichluy")) {
        cumulative.gpa10 = normalized;
      } else if (lang.includes("kqht-tkhk-diemtbtinchitichluy")) {
        cumulative.gpa4 = normalized;
      } else if (lang.includes("kqht-tkhk-sotctichluy")) {
        cumulative.accumulatedCredits = normalized;
      } else if (lang.includes("kqht-tkhk-sotckhongdat")) {
        cumulative.debtCredits = normalized;
      }
    }
  }

  function isGradeRecordCandidate(row, classCode, courseName, fieldMap) {
    if (!classCode && !courseName) return false;
    if (!courseName) return false;
    if (row.querySelector("td.row-head")) return false;
    if (!firstGradeFieldCell(fieldMap, ["DiemTongKet", "DiemTongKet1", "DiemTinChi", "DiemChu", "GhiChu"])) return false;

    const courseNorm = norm(courseName);
    if (
      courseNorm.includes("diem trung binh") ||
      courseNorm.includes("tong so tin chi") ||
      courseNorm.includes("xu ly hoc vu")
    ) {
      return false;
    }
    return true;
  }

  function buildGradeRecordKey(record, index) {
    return [
      text(record.yearLabel),
      text(record.termLabel),
      text(record.classCode),
      text(record.courseName),
      Number.isFinite(record.stt) ? String(record.stt) : "",
      String(index + 1)
    ].join("||");
  }

  function scrapeGradeRecords() {
    const table = getGradeTable();
    if (!table) {
      throw new Error("Không tìm thấy bảng điểm #xemDiem_aaa.");
    }

    const rows = [];
    if (table.tBodies && table.tBodies.length > 0) {
      for (const tbody of Array.from(table.tBodies)) {
        rows.push(...Array.from(tbody.rows || []));
      }
    } else {
      rows.push(...Array.from(table.querySelectorAll("tr")));
    }

    const records = [];
    const webTermSummaryMap = new Map();
    const webCumulativeSummary = {
      gpa10: null,
      gpa4: null,
      accumulatedCredits: null,
      debtCredits: null
    };

    let currentTermLabel = "";
    let currentYearLabel = "";

    for (const row of rows) {
      const headingCell = row.querySelector("td.row-head");
      if (headingCell) {
        const termHeading = parseGradeTermHeading(headingCell.textContent);
        if (termHeading) {
          currentTermLabel = termHeading.termLabel;
          currentYearLabel = termHeading.yearLabel;
        }
        continue;
      }

      const cells = Array.from(row.cells || []);
      if (!cells.length) continue;

      if (cells.length === 1 && cells[0].colSpan > 1) {
        const termHeading = parseGradeTermHeading(cells[0].textContent);
        if (termHeading) {
          currentTermLabel = termHeading.termLabel;
          currentYearLabel = termHeading.yearLabel;
        }
        continue;
      }

      const rowNorm = norm(row.textContent);
      if (rowNorm && (isGradeSummaryRow(rowNorm) || row.querySelector("span[lang^='kqht-tkhk-']"))) {
        parseWebSummaryFromRow(row, currentYearLabel, currentTermLabel, webTermSummaryMap, webCumulativeSummary);
        continue;
      }

      const fieldMap = getGradeFieldCellMap(row);
      const stt = parseLocaleNumber(cells[0] && cells[0].textContent);
      const classCode = text(cells[1] && cells[1].textContent);
      const courseName = text(cells[2] && cells[2].textContent);
      const creditsRaw = parseLocaleNumber(cells[3] && cells[3].textContent);
      const credits = Number.isFinite(creditsRaw) && creditsRaw > 0 ? normalizeGradeNumberValue(creditsRaw) : null;

      if (!isGradeRecordCandidate(row, classCode, courseName, fieldMap)) continue;

      const components = buildGradeComponentBundle(fieldMap);
      const total4Cell = firstGradeFieldCell(fieldMap, "DiemTinChi");
      const letterCell = firstGradeFieldCell(fieldMap, "DiemChu");
      const xepLoaiCell = firstGradeFieldCell(fieldMap, "XepLoai");
      const total10Cell = firstGradeFieldCell(fieldMap, "DiemTongKet") || firstGradeFieldCell(fieldMap, "DiemTongKet1");
      const total4 = parseScore4FromCell(total4Cell);
      const letter = text(letterCell && letterCell.textContent);
      const xepLoai = text(xepLoaiCell && xepLoaiCell.textContent);
      const note = text(firstGradeFieldCell(fieldMap, "GhiChu") && firstGradeFieldCell(fieldMap, "GhiChu").textContent);

      let passCell = firstGradeFieldCell(fieldMap, ["IsDat", "Dat"]);
      if (!passCell) {
        for (let i = cells.length - 1; i >= 0; i -= 1) {
          if (cells[i].querySelector("div.check, div.no-check, input[type='checkbox']")) {
            passCell = cells[i];
            break;
          }
        }
      }
      const pass = parsePassFromCell(
        passCell,
        note,
        letter,
        total4,
        components.total10,
        xepLoai,
        [total10Cell, total4Cell, letterCell, xepLoaiCell]
      );
      const includeInGPA = !isExcludedFromGPA({
        courseName,
        credits,
        total4,
        letter
      });

      records.push({
        recordKey: "",
        yearLabel: text(currentYearLabel),
        termLabel: text(currentTermLabel),
        stt: Number.isFinite(stt) ? Math.round(stt) : null,
        classCode,
        courseName,
        credits,
        total10: components.total10,
        rawTotal10: components.total10Raw,
        total4,
        letter,
        xepLoai,
        note,
        pass: pass === true,
        includeInGPA,
        components,
        edit: null,
        predicted: null
      });
    }

    if (!records.length) {
      throw new Error("Không tìm thấy dữ liệu môn học trong bảng điểm.");
    }

    for (let i = 0; i < records.length; i += 1) {
      records[i].recordKey = buildGradeRecordKey(records[i], i);
    }

    return {
      records,
      webSummary: {
        terms: Array.from(webTermSummaryMap.values()),
        cumulative: webCumulativeSummary
      }
    };
  }

  function buildDefaultComponentEdits(record) {
    return {
      cc: averageFinite(record.components.chuyenCan),
      tx: averageFinite(record.components.thuongKy),
      hs1: averageFinite(record.components.heSo1),
      hs2: averageFinite(record.components.heSo2),
      th: averageFinite(record.components.thucHanh),
      exam: Number.isFinite(record.components.exam) ? record.components.exam : null
    };
  }

  function buildComponentFieldValuesFromRecord(record) {
    const values = {
      DiemChuyenCan1: Number.isFinite(record.components.chuyenCan[0]) ? record.components.chuyenCan[0] : null,
      DiemThi: Number.isFinite(record.components.exam) ? record.components.exam : null
    };
    for (let i = 0; i < 3; i += 1) {
      values[`DiemThuongKy${i + 1}`] = Number.isFinite(record.components.thuongKy[i]) ? record.components.thuongKy[i] : null;
    }
    for (let i = 0; i < 9; i += 1) {
      values[`DiemHeSo1${i + 1}`] = Number.isFinite(record.components.heSo1[i]) ? record.components.heSo1[i] : null;
      values[`DiemHeSo2${i + 1}`] = Number.isFinite(record.components.heSo2[i]) ? record.components.heSo2[i] : null;
    }
    for (let i = 0; i < 2; i += 1) {
      values[`DiemThucHanh${i + 1}`] = Number.isFinite(record.components.thucHanh[i]) ? record.components.thucHanh[i] : null;
    }
    return values;
  }

  function recomputeComponentEditFromFieldValues(edit) {
    const values = edit && edit.componentFieldValues ? edit.componentFieldValues : {};
    const tx = [values.DiemThuongKy1, values.DiemThuongKy2, values.DiemThuongKy3];
    const hs1 = [];
    const hs2 = [];
    const th = [];
    for (let i = 1; i <= 9; i += 1) {
      hs1.push(values[`DiemHeSo1${i}`]);
      hs2.push(values[`DiemHeSo2${i}`]);
    }
    for (let i = 1; i <= 2; i += 1) {
      th.push(values[`DiemThucHanh${i}`]);
    }

    edit.componentEdit = {
      cc: averageFinite([values.DiemChuyenCan1]),
      tx: averageFinite(tx),
      hs1: averageFinite(hs1),
      hs2: averageFinite(hs2),
      th: averageFinite(th),
      exam: Number.isFinite(values.DiemThi) ? normalizeScore10(values.DiemThi) : null
    };
  }

  function isComponentFieldChanged(record, edit) {
    const base = record.baseComponentFieldValues || buildComponentFieldValuesFromRecord(record);
    const current = edit && edit.componentFieldValues ? edit.componentFieldValues : {};
    for (const title of GRADE_COMPONENT_TITLES) {
      if (!isSameNumericValue(current[title], base[title])) {
        return true;
      }
    }
    return false;
  }

  function deriveGradeFromComponentEdits(componentEdit) {
    const groups = [
      { value: componentEdit.cc, weight: 1 },
      { value: componentEdit.tx, weight: 1 },
      { value: componentEdit.hs1, weight: 1 },
      { value: componentEdit.hs2, weight: 2 },
      { value: componentEdit.th, weight: 1 }
    ];
    let numerator = 0;
    let denominator = 0;
    for (const group of groups) {
      if (!Number.isFinite(group.value)) continue;
      numerator += group.value * group.weight;
      denominator += group.weight;
    }

    const tbThuongKy = denominator > 0 ? normalizeScore10(numerator / denominator) : null;
    const exam = Number.isFinite(componentEdit.exam) ? normalizeScore10(componentEdit.exam) : null;
    const total10 = Number.isFinite(tbThuongKy) && Number.isFinite(exam)
      ? normalizeScore10(roundToOneDecimal(tbThuongKy * 0.4 + exam * 0.6))
      : null;
    return {
      tbThuongKyCalc: tbThuongKy,
      exam,
      total10Calc: total10
    };
  }

  function createDefaultGradeEdit(record) {
    const componentFieldValues = record.baseComponentFieldValues
      ? { ...record.baseComponentFieldValues }
      : buildComponentFieldValuesFromRecord(record);
    const componentEdit = buildDefaultComponentEdits(record);
    return {
      total10Edit: Number.isFinite(record.total10) ? record.total10 : null,
      componentEdit,
      componentFieldValues,
      lastEditedSource: "total10"
    };
  }

  function buildGradePrediction(record) {
    const edit = record.edit || createDefaultGradeEdit(record);
    if (!edit.componentFieldValues || typeof edit.componentFieldValues !== "object") {
      edit.componentFieldValues = buildComponentFieldValuesFromRecord(record);
    }
    if (!edit.componentEdit || typeof edit.componentEdit !== "object") {
      recomputeComponentEditFromFieldValues(edit);
    }
    const derived = deriveGradeFromComponentEdits(edit.componentEdit);
    const baseTotal10 = Number.isFinite(record.total10)
      ? normalizeScore10(record.total10)
      : Number.isFinite(record.rawTotal10)
        ? normalizeScore10(record.rawTotal10)
        : null;
    const total10Edit = Number.isFinite(edit.total10Edit) ? normalizeScore10(edit.total10Edit) : baseTotal10;
    const componentChanged = isComponentFieldChanged(record, edit);
    const totalChanged = !(
      (!Number.isFinite(total10Edit) && !Number.isFinite(baseTotal10)) ||
      (Number.isFinite(total10Edit) && Number.isFinite(baseTotal10) && Math.abs(total10Edit - baseTotal10) < 0.001)
    );
    const predictionChanged = totalChanged;
    const changed = totalChanged || componentChanged;

    const total4FromLookup = Number.isFinite(total10Edit)
      ? normalizeGradeNumberValue(gradePoint4From10(total10Edit))
      : null;
    const total4Calc = predictionChanged
      ? total4FromLookup
      : (Number.isFinite(record.total4) ? normalizeGradeNumberValue(record.total4) : total4FromLookup);
    const includePred = predictionChanged
      ? !isExcludedFromGPA(record, { total4: total4Calc, letter: point4ToLetter(total4Calc) })
      : record.includeInGPA === true;
    const passPredicted = predictionChanged
      ? (Number.isFinite(total4Calc) ? total4Calc >= 1 : record.pass === true)
      : record.pass === true;
    const letterCalc = predictionChanged
      ? point4ToLetter(total4Calc)
      : text(record.letter) || point4ToLetter(total4Calc);
    const xepLoaiCalc = predictionChanged
      ? score10ToXepLoai(total10Edit, letterCalc)
      : text(record.xepLoai) || score10ToXepLoai(total10Edit, letterCalc);

    return {
      tbThuongKyCalc: derived.tbThuongKyCalc,
      examCalc: derived.exam,
      total10Edit,
      total4Calc,
      letterCalc,
      xepLoaiCalc,
      passPredicted,
      includeInGPA: includePred,
      changed
    };
  }

  function hydrateRecordsForPrediction(records) {
    return records.map((record) => {
      const hydrated = {
        ...record,
        baseComponentFieldValues: record.baseComponentFieldValues
          ? { ...record.baseComponentFieldValues }
          : buildComponentFieldValuesFromRecord(record),
        edit: record.edit && typeof record.edit === "object" ? record.edit : createDefaultGradeEdit(record)
      };
      if (!hydrated.edit.componentFieldValues || typeof hydrated.edit.componentFieldValues !== "object") {
        hydrated.edit.componentFieldValues = { ...hydrated.baseComponentFieldValues };
      }
      recomputeComponentEditFromFieldValues(hydrated.edit);
      hydrated.predicted = buildGradePrediction(hydrated);
      return hydrated;
    });
  }

  function buildEmptyGradeSummaryBucket() {
    return {
      attemptedCreditsCurrent: 0,
      attemptedCreditsPredicted: 0,
      earnedCreditsCurrent: 0,
      earnedCreditsPredicted: 0,
      excludedCreditsCurrent: 0,
      excludedCreditsPredicted: 0,
      debtCreditsCurrent: 0,
      debtCreditsPredicted: 0,
      weighted4TermCurrent: 0,
      weighted10TermCurrent: 0,
      weighted4TermPredicted: 0,
      weighted10TermPredicted: 0,
      weighted4CumCurrent: 0,
      weighted10CumCurrent: 0,
      weighted4CumPredicted: 0,
      weighted10CumPredicted: 0
    };
  }

  function finalizeGradeSummaryBucket(bucket, mode) {
    const useEarnedForGpa = mode === "cumulative";
    const denominatorCurrent = useEarnedForGpa ? bucket.earnedCreditsCurrent : bucket.attemptedCreditsCurrent;
    const denominatorPredicted = useEarnedForGpa ? bucket.earnedCreditsPredicted : bucket.attemptedCreditsPredicted;
    const weighted4Current = useEarnedForGpa ? bucket.weighted4CumCurrent : bucket.weighted4TermCurrent;
    const weighted10Current = useEarnedForGpa ? bucket.weighted10CumCurrent : bucket.weighted10TermCurrent;
    const weighted4Predicted = useEarnedForGpa ? bucket.weighted4CumPredicted : bucket.weighted4TermPredicted;
    const weighted10Predicted = useEarnedForGpa ? bucket.weighted10CumPredicted : bucket.weighted10TermPredicted;
    const gpa4Current = denominatorCurrent > 0 ? weighted4Current / denominatorCurrent : null;
    const gpa10Current = denominatorCurrent > 0 ? weighted10Current / denominatorCurrent : null;
    const gpa4Predicted = denominatorPredicted > 0 ? weighted4Predicted / denominatorPredicted : null;
    const gpa10Predicted = denominatorPredicted > 0 ? weighted10Predicted / denominatorPredicted : null;
    return {
      current: {
        creditsGPA: normalizeGradeNumberValue(denominatorCurrent),
        attemptedCredits: normalizeGradeNumberValue(bucket.attemptedCreditsCurrent),
        earnedCredits: normalizeGradeNumberValue(bucket.earnedCreditsCurrent),
        excludedCredits: normalizeGradeNumberValue(bucket.excludedCreditsCurrent),
        gpa4: normalizeGradeNumberValue(gpa4Current),
        gpa10: normalizeGradeNumberValue(gpa10Current),
        passedCredits: normalizeGradeNumberValue(bucket.earnedCreditsCurrent),
        debtCredits: normalizeGradeNumberValue(bucket.debtCreditsCurrent)
      },
      predicted: {
        creditsGPA: normalizeGradeNumberValue(denominatorPredicted),
        attemptedCredits: normalizeGradeNumberValue(bucket.attemptedCreditsPredicted),
        earnedCredits: normalizeGradeNumberValue(bucket.earnedCreditsPredicted),
        excludedCredits: normalizeGradeNumberValue(bucket.excludedCreditsPredicted),
        gpa4: normalizeGradeNumberValue(gpa4Predicted),
        gpa10: normalizeGradeNumberValue(gpa10Predicted),
        passedCredits: normalizeGradeNumberValue(bucket.earnedCreditsPredicted),
        debtCredits: normalizeGradeNumberValue(bucket.debtCreditsPredicted)
      }
    };
  }

  function calculateGradeSummary(records, webSummary) {
    const termMap = new Map();
    const cumulative = buildEmptyGradeSummaryBucket();
    const webTermMap = new Map();
    for (const term of (webSummary && Array.isArray(webSummary.terms) ? webSummary.terms : [])) {
      webTermMap.set(buildGradeTermKey(term.yearLabel, term.termLabel), term);
    }

    const ensureTerm = (record) => {
      const key = buildGradeTermKey(record.yearLabel, record.termLabel);
      if (!termMap.has(key)) {
        termMap.set(key, {
          key,
          yearLabel: text(record.yearLabel),
          termLabel: text(record.termLabel),
          bucket: buildEmptyGradeSummaryBucket()
        });
      }
      return termMap.get(key);
    };

    const accumulate = (bucket, record) => {
      const credits = Number(record.credits);
      if (!(Number.isFinite(credits) && credits > 0)) return;

      const includeCurrent = record.includeInGPA === true;
      const passCurrent = record.pass === true;
      const total4Current = Number.isFinite(record.total4) ? record.total4 : null;
      const total10Current = Number.isFinite(record.total10) ? record.total10 : null;

      const includePred = Boolean(record.predicted && record.predicted.includeInGPA === true);
      const passPred = Boolean(record.predicted ? record.predicted.passPredicted === true : passCurrent);
      const total4Pred = Number.isFinite(record.predicted && record.predicted.total4Calc)
        ? record.predicted.total4Calc
        : null;
      const total10Pred = Number.isFinite(record.predicted && record.predicted.total10Edit)
        ? record.predicted.total10Edit
        : null;

      if (includeCurrent) {
        bucket.attemptedCreditsCurrent += credits;
        if (Number.isFinite(total4Current) && Number.isFinite(total10Current)) {
          bucket.weighted4TermCurrent += credits * total4Current;
          bucket.weighted10TermCurrent += credits * total10Current;
        }
        if (passCurrent) {
          bucket.earnedCreditsCurrent += credits;
          if (Number.isFinite(total4Current) && Number.isFinite(total10Current)) {
            bucket.weighted4CumCurrent += credits * total4Current;
            bucket.weighted10CumCurrent += credits * total10Current;
          }
        } else {
          bucket.debtCreditsCurrent += credits;
        }
      } else {
        bucket.excludedCreditsCurrent += credits;
      }

      if (includePred) {
        bucket.attemptedCreditsPredicted += credits;
        if (Number.isFinite(total4Pred) && Number.isFinite(total10Pred)) {
          bucket.weighted4TermPredicted += credits * total4Pred;
          bucket.weighted10TermPredicted += credits * total10Pred;
        }
        if (passPred) {
          bucket.earnedCreditsPredicted += credits;
          if (Number.isFinite(total4Pred) && Number.isFinite(total10Pred)) {
            bucket.weighted4CumPredicted += credits * total4Pred;
            bucket.weighted10CumPredicted += credits * total10Pred;
          }
        } else {
          bucket.debtCreditsPredicted += credits;
        }
      } else {
        bucket.excludedCreditsPredicted += credits;
      }
    };

    for (const record of records) {
      const termHolder = ensureTerm(record);
      accumulate(termHolder.bucket, record);
      accumulate(cumulative, record);
    }

    const terms = Array.from(termMap.values())
      .map((item) => {
        const web = webTermMap.get(item.key) || {};
        const finalized = finalizeGradeSummaryBucket(item.bucket, "term");
        return {
          yearLabel: item.yearLabel,
          termLabel: item.termLabel,
          current: finalized.current,
          predicted: finalized.predicted,
          web: {
            gpa10: Number.isFinite(web.gpa10) ? web.gpa10 : null,
            gpa4: Number.isFinite(web.gpa4) ? web.gpa4 : null
          }
        };
      })
      .sort((a, b) => {
        const yearDiff = text(a.yearLabel).localeCompare(text(b.yearLabel), "vi", { numeric: true });
        if (yearDiff !== 0) return yearDiff;
        return text(a.termLabel).localeCompare(text(b.termLabel), "vi", { numeric: true });
      });

    const cumulativeFinal = finalizeGradeSummaryBucket(cumulative, "cumulative");
    const cumulativeWeb = webSummary && webSummary.cumulative ? webSummary.cumulative : {};
    const termGPA4 = terms.map((term) => ({
      yearLabel: term.yearLabel,
      termLabel: term.termLabel,
      value: term.current.gpa4
    }));
    const termGPA10 = terms.map((term) => ({
      yearLabel: term.yearLabel,
      termLabel: term.termLabel,
      value: term.current.gpa10
    }));
    return {
      terms,
      termGPA4,
      termGPA10,
      cumGPA4: cumulativeFinal.current.gpa4,
      cumGPA10: cumulativeFinal.current.gpa10,
      attemptedCredits: cumulativeFinal.current.attemptedCredits,
      earnedCredits: cumulativeFinal.current.earnedCredits,
      excludedCredits: cumulativeFinal.current.excludedCredits,
      cumulative: {
        current: {
          ...cumulativeFinal.current,
          cumGPA4: cumulativeFinal.current.gpa4,
          cumGPA10: cumulativeFinal.current.gpa10
        },
        predicted: {
          ...cumulativeFinal.predicted,
          cumGPA4: cumulativeFinal.predicted.gpa4,
          cumGPA10: cumulativeFinal.predicted.gpa10
        },
        web: {
          gpa10: Number.isFinite(cumulativeWeb.gpa10) ? cumulativeWeb.gpa10 : null,
          gpa4: Number.isFinite(cumulativeWeb.gpa4) ? cumulativeWeb.gpa4 : null,
          accumulatedCredits: Number.isFinite(cumulativeWeb.accumulatedCredits)
            ? cumulativeWeb.accumulatedCredits
            : null,
          debtCredits: Number.isFinite(cumulativeWeb.debtCredits)
            ? cumulativeWeb.debtCredits
            : null
        }
      }
    };
  }

  function mergeRecordsWithEditorState(scrapedRecords) {
    const records = scrapedRecords.map((record) => ({ ...record }));
    if (!(gradeEditorState && gradeEditorState.pageUrl === location.href && Array.isArray(gradeEditorState.records))) {
      return hydrateRecordsForPrediction(records);
    }

    const editorMap = new Map();
    for (const record of gradeEditorState.records) {
      editorMap.set(record.recordKey, record);
    }

    for (let i = 0; i < records.length; i += 1) {
      const key = buildGradeRecordKey(records[i], i);
      records[i].recordKey = key;
      const edited = editorMap.get(key);
      records[i].edit = edited && edited.edit ? edited.edit : createDefaultGradeEdit(records[i]);
    }

    return hydrateRecordsForPrediction(records);
  }

  function buildGradeCsvContent(records) {
    const headers = [
      "Year",
      "Term",
      "STT",
      "ClassCode",
      "CourseName",
      "Credits",
      "CC_Avg",
      "TX_Avg",
      "HS1_Avg",
      "HS2_Avg",
      "TH_Avg",
      "TBThuongKy_Current",
      "Exam_Current",
      "Total10_Current",
      "Total4_Current",
      "Letter",
      "Note",
      "Pass_Current",
      "Total10_Edit",
      "Total4_Calc",
      "Pass_Predicted",
      "IncludeInGPA_Current",
      "IncludeInGPA_Predicted",
      "Changed"
    ];
    const rows = [headers];
    for (const record of records) {
      const editComp = record.edit && record.edit.componentEdit ? record.edit.componentEdit : {};
      rows.push([
        text(record.yearLabel),
        text(record.termLabel),
        Number.isFinite(record.stt) ? record.stt : "",
        text(record.classCode),
        text(record.courseName),
        Number.isFinite(record.credits) ? record.credits.toFixed(2) : "",
        Number.isFinite(editComp.cc) ? editComp.cc.toFixed(2) : "",
        Number.isFinite(editComp.tx) ? editComp.tx.toFixed(2) : "",
        Number.isFinite(editComp.hs1) ? editComp.hs1.toFixed(2) : "",
        Number.isFinite(editComp.hs2) ? editComp.hs2.toFixed(2) : "",
        Number.isFinite(editComp.th) ? editComp.th.toFixed(2) : "",
        Number.isFinite(record.components && record.components.tbThuongKy) ? record.components.tbThuongKy.toFixed(2) : "",
        Number.isFinite(record.components && record.components.exam) ? record.components.exam.toFixed(2) : "",
        Number.isFinite(record.total10) ? record.total10.toFixed(2) : "",
        Number.isFinite(record.total4) ? record.total4.toFixed(2) : "",
        text(record.letter),
        text(record.note),
        record.pass ? "TRUE" : "FALSE",
        Number.isFinite(record.predicted && record.predicted.total10Edit) ? record.predicted.total10Edit.toFixed(2) : "",
        Number.isFinite(record.predicted && record.predicted.total4Calc) ? record.predicted.total4Calc.toFixed(2) : "",
        record.predicted && record.predicted.passPredicted ? "TRUE" : "FALSE",
        record.includeInGPA ? "TRUE" : "FALSE",
        record.predicted && record.predicted.includeInGPA ? "TRUE" : "FALSE",
        record.predicted && record.predicted.changed ? "TRUE" : "FALSE"
      ]);
    }
    return `\uFEFF${rows.map((row) => row.map(escapeCsvCell).join(",")).join("\r\n")}`;
  }

  function setWorksheetHeaderStyle(row, headerColor) {
    row.height = 24;
    row.font = { bold: true, color: { argb: "FFFFFFFF" } };
    row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: headerColor } };
    row.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "CBD5E1" } },
        left: { style: "thin", color: { argb: "CBD5E1" } },
        bottom: { style: "thin", color: { argb: "CBD5E1" } },
        right: { style: "thin", color: { argb: "CBD5E1" } }
      };
    });
  }

  function setWorksheetRowBorder(row) {
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "E2E8F0" } },
        left: { style: "thin", color: { argb: "E2E8F0" } },
        bottom: { style: "thin", color: { argb: "E2E8F0" } },
        right: { style: "thin", color: { argb: "E2E8F0" } }
      };
    });
  }

  function buildGradeDetailColumns() {
    const columns = [
      { header: "Year", key: "yearLabel", width: 12 },
      { header: "Term", key: "termLabel", width: 10 },
      { header: "STT", key: "stt", width: 8 },
      { header: "ClassCode", key: "classCode", width: 16 },
      { header: "CourseName", key: "courseName", width: 34 },
      { header: "Credits", key: "credits", width: 10 },
      { header: "CC1", key: "cc1", width: 9 }
    ];

    for (let i = 1; i <= 3; i += 1) columns.push({ header: `TX${i}`, key: `tx${i}`, width: 9 });
    for (let i = 1; i <= 9; i += 1) columns.push({ header: `HS1_${i}`, key: `hs1_${i}`, width: 9 });
    for (let i = 1; i <= 9; i += 1) columns.push({ header: `HS2_${i}`, key: `hs2_${i}`, width: 9 });
    for (let i = 1; i <= 2; i += 1) columns.push({ header: `TH${i}`, key: `th${i}`, width: 9 });

    columns.push(
      { header: "TBThuongKy(Current)", key: "tbCurrent", width: 14 },
      { header: "Exam(Current)", key: "examCurrent", width: 12 },
      { header: "Total10(Current)", key: "total10Current", width: 14 },
      { header: "RawTotal10", key: "rawTotal10", width: 12 },
      { header: "Total4(Current)", key: "total4Current", width: 14 },
      { header: "Letter", key: "letter", width: 8 },
      { header: "Note", key: "note", width: 14 },
      { header: "Pass(Current)", key: "passCurrent", width: 12 },
      { header: "CC(EditAvg)", key: "ccEdit", width: 11 },
      { header: "TX(EditAvg)", key: "txEdit", width: 11 },
      { header: "HS1(EditAvg)", key: "hs1Edit", width: 11 },
      { header: "HS2(EditAvg)", key: "hs2Edit", width: 11 },
      { header: "TH(EditAvg)", key: "thEdit", width: 11 },
      { header: "Exam(Edit)", key: "examEdit", width: 10 },
      { header: "TBThuongKy(Calc)", key: "tbCalc", width: 14 },
      { header: "Total10(Edit)", key: "total10Edit", width: 12 },
      { header: "Total4(Calc)", key: "total4Calc", width: 12 },
      { header: "Pass(Predicted)", key: "passPred", width: 13 },
      { header: "Include(Current)", key: "includeCurrent", width: 12 },
      { header: "Include(Predicted)", key: "includePred", width: 13 },
      { header: "Changed", key: "changed", width: 10 }
    );
    return columns;
  }

  async function exportGradesAsXlsx(records, summary, webSummary, fileName) {
    assertExportLibraries("xlsx");

    const workbook = new window.ExcelJS.Workbook();
    workbook.creator = "UNETI Time Mapper";
    workbook.created = new Date();

    const gradesSheet = workbook.addWorksheet("Grades");
    gradesSheet.columns = [
      { header: "Year", key: "yearLabel", width: 12 },
      { header: "Term", key: "termLabel", width: 10 },
      { header: "STT", key: "stt", width: 8 },
      { header: "ClassCode", key: "classCode", width: 18 },
      { header: "CourseName", key: "courseName", width: 36 },
      { header: "Credits", key: "credits", width: 11 },
      { header: "Total10 (Current)", key: "total10Current", width: 16 },
      { header: "Total4 (Current)", key: "total4Current", width: 16 },
      { header: "Letter", key: "letter", width: 10 },
      { header: "Note", key: "note", width: 16 },
      { header: "Pass", key: "pass", width: 9 },
      { header: "Total10 (Edit)", key: "total10Edit", width: 14 },
      { header: "Total4 (Calc)", key: "total4Calc", width: 14 },
      { header: "IncludeInGPA (Current)", key: "includeCurrent", width: 17 },
      { header: "IncludeInGPA (Predicted)", key: "includePred", width: 19 }
    ];
    setWorksheetHeaderStyle(gradesSheet.getRow(1), "1D4ED8");
    gradesSheet.views = [{ state: "frozen", ySplit: 1 }];

    for (const record of records) {
      const row = gradesSheet.addRow({
        yearLabel: text(record.yearLabel),
        termLabel: text(record.termLabel),
        stt: Number.isFinite(record.stt) ? record.stt : "",
        classCode: text(record.classCode),
        courseName: text(record.courseName),
        credits: Number.isFinite(record.credits) ? record.credits : "",
        total10Current: Number.isFinite(record.total10) ? record.total10 : "",
        total4Current: Number.isFinite(record.total4) ? record.total4 : "",
        letter: text(record.letter),
        note: text(record.note),
        pass: record.pass === true,
        total10Edit: Number.isFinite(record.predicted && record.predicted.total10Edit) ? record.predicted.total10Edit : "",
        total4Calc: "",
        includeCurrent: record.includeInGPA === true,
        includePred: ""
      });

      const rowNum = row.number;
      row.getCell("M").value = {
        formula: `IF(L${rowNum}="","",IF(ABS(L${rowNum}-G${rowNum})<0.0001,H${rowNum},LOOKUP(L${rowNum},Lookup!$A$2:$A$10,Lookup!$B$2:$B$10)))`,
        result: Number.isFinite(record.predicted && record.predicted.total4Calc) ? record.predicted.total4Calc : null
      };
      row.getCell("O").value = {
        formula: `AND(N${rowNum}=TRUE,F${rowNum}>0,ISNUMBER(M${rowNum}))`,
        result: record.predicted && record.predicted.includeInGPA === true
      };

      row.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      for (const cellId of ["A", "B", "C", "D", "F", "G", "H", "I", "K", "L", "M", "N", "O"]) {
        row.getCell(cellId).alignment = { horizontal: "center", vertical: "middle" };
      }
      setWorksheetRowBorder(row);

      const noteNorm = norm(record.note);
      if (noteNorm.includes("thi lai") || record.pass === false) {
        row.eachCell((cell) => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F2" } };
        });
      } else if (record.predicted && record.predicted.changed) {
        row.eachCell((cell) => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FEF9C3" } };
        });
      } else if (rowNum % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F8FAFC" } };
        });
      }
    }

    const gradesDataEnd = Math.max(2, gradesSheet.rowCount);
    gradesSheet.autoFilter = { from: "A1", to: "O1" };
    for (const columnId of ["F", "G", "H", "L", "M"]) {
      gradesSheet.getColumn(columnId).numFmt = "0.00";
    }
    autoFitWorksheetColumns(gradesSheet, 10, 40);

    const detailSheet = workbook.addWorksheet("Grades_Detail");
    detailSheet.columns = buildGradeDetailColumns();
    setWorksheetHeaderStyle(detailSheet.getRow(1), "0F766E");
    detailSheet.views = [{ state: "frozen", ySplit: 1 }];

    for (const record of records) {
      const componentEdit = record.edit && record.edit.componentEdit ? record.edit.componentEdit : {};
      const rowPayload = {
        yearLabel: text(record.yearLabel),
        termLabel: text(record.termLabel),
        stt: Number.isFinite(record.stt) ? record.stt : "",
        classCode: text(record.classCode),
        courseName: text(record.courseName),
        credits: Number.isFinite(record.credits) ? record.credits : "",
        cc1: Number.isFinite(record.components.chuyenCan[0]) ? record.components.chuyenCan[0] : "",
        tbCurrent: Number.isFinite(record.components.tbThuongKy) ? record.components.tbThuongKy : "",
        examCurrent: Number.isFinite(record.components.exam) ? record.components.exam : "",
        total10Current: Number.isFinite(record.total10) ? record.total10 : "",
        rawTotal10: Number.isFinite(record.rawTotal10) ? record.rawTotal10 : "",
        total4Current: Number.isFinite(record.total4) ? record.total4 : "",
        letter: text(record.letter),
        note: text(record.note),
        passCurrent: record.pass === true,
        ccEdit: Number.isFinite(componentEdit.cc) ? componentEdit.cc : "",
        txEdit: Number.isFinite(componentEdit.tx) ? componentEdit.tx : "",
        hs1Edit: Number.isFinite(componentEdit.hs1) ? componentEdit.hs1 : "",
        hs2Edit: Number.isFinite(componentEdit.hs2) ? componentEdit.hs2 : "",
        thEdit: Number.isFinite(componentEdit.th) ? componentEdit.th : "",
        examEdit: Number.isFinite(componentEdit.exam) ? componentEdit.exam : "",
        tbCalc: Number.isFinite(record.predicted && record.predicted.tbThuongKyCalc) ? record.predicted.tbThuongKyCalc : "",
        total10Edit: Number.isFinite(record.predicted && record.predicted.total10Edit) ? record.predicted.total10Edit : "",
        total4Calc: Number.isFinite(record.predicted && record.predicted.total4Calc) ? record.predicted.total4Calc : "",
        passPred: record.predicted && record.predicted.passPredicted === true,
        includeCurrent: record.includeInGPA === true,
        includePred: record.predicted && record.predicted.includeInGPA === true,
        changed: record.predicted && record.predicted.changed === true
      };

      for (let i = 0; i < 3; i += 1) {
        rowPayload[`tx${i + 1}`] = Number.isFinite(record.components.thuongKy[i]) ? record.components.thuongKy[i] : "";
      }
      for (let i = 0; i < 9; i += 1) {
        rowPayload[`hs1_${i + 1}`] = Number.isFinite(record.components.heSo1[i]) ? record.components.heSo1[i] : "";
        rowPayload[`hs2_${i + 1}`] = Number.isFinite(record.components.heSo2[i]) ? record.components.heSo2[i] : "";
      }
      for (let i = 0; i < 2; i += 1) {
        rowPayload[`th${i + 1}`] = Number.isFinite(record.components.thucHanh[i]) ? record.components.thucHanh[i] : "";
      }

      const row = detailSheet.addRow(rowPayload);
      setWorksheetRowBorder(row);
      if (record.predicted && record.predicted.changed) {
        row.eachCell((cell) => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FEF9C3" } };
        });
      } else if (row.number % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F8FAFC" } };
        });
      }
    }
    const detailLastCol = detailSheet.getColumn(detailSheet.columnCount).letter || "AX";
    detailSheet.autoFilter = { from: "A1", to: `${detailLastCol}1` };
    for (const column of detailSheet.columns) {
      const headerText = String((column && column.header) || "");
      if (/Credits|CC|TX|HS|TH|TB|Exam|Total10|Total4/i.test(headerText)) {
        column.numFmt = "0.00";
      }
    }
    autoFitWorksheetColumns(detailSheet, 8, 24);

    const gpaSheet = workbook.addWorksheet("GPA");
    gpaSheet.views = [{ state: "frozen", ySplit: 1 }];
    gpaSheet.columns = [
      { header: "Year", key: "yearLabel", width: 12 },
      { header: "Term", key: "termLabel", width: 10 },
      { header: "Attempted Credits Current", key: "creditsCurrent", width: 18 },
      { header: "Attempted Credits Predicted", key: "creditsPredicted", width: 19 },
      { header: "GPA4 Current", key: "gpa4Current", width: 12 },
      { header: "GPA10 Current", key: "gpa10Current", width: 13 },
      { header: "GPA4 Predicted", key: "gpa4Predicted", width: 13 },
      { header: "GPA10 Predicted", key: "gpa10Predicted", width: 14 },
      { header: "Passed Credits Current", key: "passedCurrent", width: 18 },
      { header: "Debt Credits Current", key: "debtCurrent", width: 16 },
      { header: "Passed Credits Predicted", key: "passedPred", width: 19 },
      { header: "Debt Credits Predicted", key: "debtPred", width: 17 },
      { header: "Web GPA10", key: "webGpa10", width: 11 },
      { header: "Web GPA4", key: "webGpa4", width: 11 }
    ];
    setWorksheetHeaderStyle(gpaSheet.getRow(1), "0F766E");

    const yearRange = `Grades!$A$2:$A$${gradesDataEnd}`;
    const termRange = `Grades!$B$2:$B$${gradesDataEnd}`;
    const creditsRange = `Grades!$F$2:$F$${gradesDataEnd}`;
    const total10CurrentRange = `Grades!$G$2:$G$${gradesDataEnd}`;
    const total4CurrentRange = `Grades!$H$2:$H$${gradesDataEnd}`;
    const passRange = `Grades!$K$2:$K$${gradesDataEnd}`;
    const total10EditRange = `Grades!$L$2:$L$${gradesDataEnd}`;
    const total4CalcRange = `Grades!$M$2:$M$${gradesDataEnd}`;
    const includeCurrentRange = `Grades!$N$2:$N$${gradesDataEnd}`;
    const includePredRange = `Grades!$O$2:$O$${gradesDataEnd}`;

    let gpaRowPointer = 2;
    for (const term of summary.terms) {
      const rowNum = gpaRowPointer;
      const row = gpaSheet.getRow(rowNum);
      row.getCell("A").value = text(term.yearLabel);
      row.getCell("B").value = text(term.termLabel);
      row.getCell("C").value = {
        formula: `SUMIFS(${creditsRange},${yearRange},A${rowNum},${termRange},B${rowNum},${includeCurrentRange},TRUE)`,
        result: Number.isFinite(term.current.attemptedCredits) ? term.current.attemptedCredits : null
      };
      row.getCell("D").value = {
        formula: `SUMIFS(${creditsRange},${yearRange},A${rowNum},${termRange},B${rowNum},${includePredRange},TRUE)`,
        result: Number.isFinite(term.predicted.attemptedCredits) ? term.predicted.attemptedCredits : null
      };
      row.getCell("E").value = {
        formula: `IF(C${rowNum}=0,"",SUMPRODUCT((${yearRange}=A${rowNum})*(${termRange}=B${rowNum})*(${includeCurrentRange}=TRUE)*(${creditsRange})*(${total4CurrentRange}))/C${rowNum})`,
        result: Number.isFinite(term.current.gpa4) ? term.current.gpa4 : null
      };
      row.getCell("F").value = {
        formula: `IF(C${rowNum}=0,"",SUMPRODUCT((${yearRange}=A${rowNum})*(${termRange}=B${rowNum})*(${includeCurrentRange}=TRUE)*(${creditsRange})*(${total10CurrentRange}))/C${rowNum})`,
        result: Number.isFinite(term.current.gpa10) ? term.current.gpa10 : null
      };
      row.getCell("G").value = {
        formula: `IF(D${rowNum}=0,"",SUMPRODUCT((${yearRange}=A${rowNum})*(${termRange}=B${rowNum})*(${includePredRange}=TRUE)*(${creditsRange})*(${total4CalcRange}))/D${rowNum})`,
        result: Number.isFinite(term.predicted.gpa4) ? term.predicted.gpa4 : null
      };
      row.getCell("H").value = {
        formula: `IF(D${rowNum}=0,"",SUMPRODUCT((${yearRange}=A${rowNum})*(${termRange}=B${rowNum})*(${includePredRange}=TRUE)*(${creditsRange})*(${total10EditRange}))/D${rowNum})`,
        result: Number.isFinite(term.predicted.gpa10) ? term.predicted.gpa10 : null
      };
      row.getCell("I").value = {
        formula: `SUMPRODUCT((${yearRange}=A${rowNum})*(${termRange}=B${rowNum})*(${includeCurrentRange}=TRUE)*(${creditsRange}>0)*(${passRange}=TRUE)*(${creditsRange}))`,
        result: Number.isFinite(term.current.passedCredits) ? term.current.passedCredits : null
      };
      row.getCell("J").value = {
        formula: `SUMPRODUCT((${yearRange}=A${rowNum})*(${termRange}=B${rowNum})*(${includeCurrentRange}=TRUE)*(${creditsRange}>0)*(${passRange}=FALSE)*(${creditsRange}))`,
        result: Number.isFinite(term.current.debtCredits) ? term.current.debtCredits : null
      };
      row.getCell("K").value = {
        formula: `SUMPRODUCT((${yearRange}=A${rowNum})*(${termRange}=B${rowNum})*(${includePredRange}=TRUE)*(${creditsRange}>0)*(${total4CalcRange}>=1)*(${creditsRange}))`,
        result: Number.isFinite(term.predicted.passedCredits) ? term.predicted.passedCredits : null
      };
      row.getCell("L").value = {
        formula: `SUMPRODUCT((${yearRange}=A${rowNum})*(${termRange}=B${rowNum})*(${includePredRange}=TRUE)*(${creditsRange}>0)*(${total4CalcRange}<1)*(${creditsRange}))`,
        result: Number.isFinite(term.predicted.debtCredits) ? term.predicted.debtCredits : null
      };
      row.getCell("M").value = Number.isFinite(term.web.gpa10) ? term.web.gpa10 : "";
      row.getCell("N").value = Number.isFinite(term.web.gpa4) ? term.web.gpa4 : "";
      row.alignment = { horizontal: "center", vertical: "middle" };
      setWorksheetRowBorder(row);
      if (rowNum % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F8FAFC" } };
        });
      }
      gpaRowPointer += 1;
    }

    const cumulativeTitleRow = gpaRowPointer + 1;
    gpaSheet.mergeCells(`A${cumulativeTitleRow}:E${cumulativeTitleRow}`);
    const cumulativeTitleCell = gpaSheet.getCell(`A${cumulativeTitleRow}`);
    cumulativeTitleCell.value = "Tổng hợp tích lũy";
    cumulativeTitleCell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cumulativeTitleCell.alignment = { horizontal: "center", vertical: "middle" };
    cumulativeTitleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1E3A8A" } };

    const cumulativeHeaderRow = cumulativeTitleRow + 1;
    gpaSheet.getCell(`A${cumulativeHeaderRow}`).value = "Chỉ số";
    gpaSheet.getCell(`B${cumulativeHeaderRow}`).value = "Current";
    gpaSheet.getCell(`C${cumulativeHeaderRow}`).value = "Predicted";
    gpaSheet.getCell(`D${cumulativeHeaderRow}`).value = "Web";
    const cumulativeHeader = gpaSheet.getRow(cumulativeHeaderRow);
    cumulativeHeader.font = { bold: true };
    cumulativeHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "DBEAFE" } };
    cumulativeHeader.alignment = { horizontal: "center", vertical: "middle" };
    setWorksheetRowBorder(cumulativeHeader);

    const cumulativeRows = [
      "Tín chỉ attempted (include)",
      "Tín chỉ earned (pass, dùng tính GPA)",
      "Tín chỉ excluded",
      "GPA hệ 4",
      "GPA hệ 10",
      "Tổng tín chỉ đạt",
      "Số tín chỉ nợ"
    ];
    const earnedRow = cumulativeHeaderRow + 2;
    for (let i = 0; i < cumulativeRows.length; i += 1) {
      const rowNum = cumulativeHeaderRow + 1 + i;
      const label = cumulativeRows[i];
      gpaSheet.getCell(`A${rowNum}`).value = label;
      if (label === "Tín chỉ attempted (include)") {
        gpaSheet.getCell(`B${rowNum}`).value = {
          formula: `SUMIF(${includeCurrentRange},TRUE,${creditsRange})`,
          result: Number.isFinite(summary.cumulative.current.attemptedCredits) ? summary.cumulative.current.attemptedCredits : null
        };
        gpaSheet.getCell(`C${rowNum}`).value = {
          formula: `SUMIF(${includePredRange},TRUE,${creditsRange})`,
          result: Number.isFinite(summary.cumulative.predicted.attemptedCredits) ? summary.cumulative.predicted.attemptedCredits : null
        };
      } else if (label === "Tín chỉ earned (pass, dùng tính GPA)") {
        gpaSheet.getCell(`B${rowNum}`).value = {
          formula: `SUMPRODUCT((${includeCurrentRange}=TRUE)*(${creditsRange}>0)*(${passRange}=TRUE)*(${creditsRange}))`,
          result: Number.isFinite(summary.cumulative.current.earnedCredits) ? summary.cumulative.current.earnedCredits : null
        };
        gpaSheet.getCell(`C${rowNum}`).value = {
          formula: `SUMPRODUCT((${includePredRange}=TRUE)*(${creditsRange}>0)*(${total4CalcRange}>=1)*(${creditsRange}))`,
          result: Number.isFinite(summary.cumulative.predicted.earnedCredits) ? summary.cumulative.predicted.earnedCredits : null
        };
        gpaSheet.getCell(`D${rowNum}`).value = Number.isFinite(summary.cumulative.web.accumulatedCredits) ? summary.cumulative.web.accumulatedCredits : "";
      } else if (label === "Tín chỉ excluded") {
        gpaSheet.getCell(`B${rowNum}`).value = {
          formula: `SUMPRODUCT((${creditsRange}>0)*(${includeCurrentRange}<>TRUE)*(${creditsRange}))`,
          result: Number.isFinite(summary.cumulative.current.excludedCredits) ? summary.cumulative.current.excludedCredits : null
        };
        gpaSheet.getCell(`C${rowNum}`).value = {
          formula: `SUMPRODUCT((${creditsRange}>0)*(${includePredRange}<>TRUE)*(${creditsRange}))`,
          result: Number.isFinite(summary.cumulative.predicted.excludedCredits) ? summary.cumulative.predicted.excludedCredits : null
        };
      } else if (label === "GPA hệ 4") {
        gpaSheet.getCell(`B${rowNum}`).value = {
          formula: `IF(B${earnedRow}=0,"",SUMPRODUCT((${includeCurrentRange}=TRUE)*(${passRange}=TRUE)*(${creditsRange})*(${total4CurrentRange}))/B${earnedRow})`,
          result: Number.isFinite(summary.cumulative.current.gpa4) ? summary.cumulative.current.gpa4 : null
        };
        gpaSheet.getCell(`C${rowNum}`).value = {
          formula: `IF(C${earnedRow}=0,"",SUMPRODUCT((${includePredRange}=TRUE)*(${total4CalcRange}>=1)*(${creditsRange})*(${total4CalcRange}))/C${earnedRow})`,
          result: Number.isFinite(summary.cumulative.predicted.gpa4) ? summary.cumulative.predicted.gpa4 : null
        };
        gpaSheet.getCell(`D${rowNum}`).value = Number.isFinite(summary.cumulative.web.gpa4) ? summary.cumulative.web.gpa4 : "";
      } else if (label === "GPA hệ 10") {
        gpaSheet.getCell(`B${rowNum}`).value = {
          formula: `IF(B${earnedRow}=0,"",SUMPRODUCT((${includeCurrentRange}=TRUE)*(${passRange}=TRUE)*(${creditsRange})*(${total10CurrentRange}))/B${earnedRow})`,
          result: Number.isFinite(summary.cumulative.current.gpa10) ? summary.cumulative.current.gpa10 : null
        };
        gpaSheet.getCell(`C${rowNum}`).value = {
          formula: `IF(C${earnedRow}=0,"",SUMPRODUCT((${includePredRange}=TRUE)*(${total4CalcRange}>=1)*(${creditsRange})*(${total10EditRange}))/C${earnedRow})`,
          result: Number.isFinite(summary.cumulative.predicted.gpa10) ? summary.cumulative.predicted.gpa10 : null
        };
        gpaSheet.getCell(`D${rowNum}`).value = Number.isFinite(summary.cumulative.web.gpa10) ? summary.cumulative.web.gpa10 : "";
      } else if (label === "Tổng tín chỉ đạt") {
        gpaSheet.getCell(`B${rowNum}`).value = {
          formula: `SUMPRODUCT((${includeCurrentRange}=TRUE)*(${creditsRange}>0)*(${passRange}=TRUE)*(${creditsRange}))`,
          result: Number.isFinite(summary.cumulative.current.passedCredits) ? summary.cumulative.current.passedCredits : null
        };
        gpaSheet.getCell(`C${rowNum}`).value = {
          formula: `SUMPRODUCT((${includePredRange}=TRUE)*(${creditsRange}>0)*(${total4CalcRange}>=1)*(${creditsRange}))`,
          result: Number.isFinite(summary.cumulative.predicted.passedCredits) ? summary.cumulative.predicted.passedCredits : null
        };
      } else if (label === "Số tín chỉ nợ") {
        gpaSheet.getCell(`B${rowNum}`).value = {
          formula: `SUMPRODUCT((${includeCurrentRange}=TRUE)*(${creditsRange}>0)*(${passRange}=FALSE)*(${creditsRange}))`,
          result: Number.isFinite(summary.cumulative.current.debtCredits) ? summary.cumulative.current.debtCredits : null
        };
        gpaSheet.getCell(`C${rowNum}`).value = {
          formula: `SUMPRODUCT((${includePredRange}=TRUE)*(${creditsRange}>0)*(${total4CalcRange}<1)*(${creditsRange}))`,
          result: Number.isFinite(summary.cumulative.predicted.debtCredits) ? summary.cumulative.predicted.debtCredits : null
        };
        gpaSheet.getCell(`D${rowNum}`).value = Number.isFinite(summary.cumulative.web.debtCredits) ? summary.cumulative.web.debtCredits : "";
      }
      setWorksheetRowBorder(gpaSheet.getRow(rowNum));
    }

    for (const columnId of ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]) {
      gpaSheet.getColumn(columnId).numFmt = "0.00";
    }
    autoFitWorksheetColumns(gpaSheet, 11, 34);

    const lookupSheet = workbook.addWorksheet("Lookup");
    lookupSheet.columns = [
      { header: "Threshold10", key: "threshold10", width: 12 },
      { header: "Point4", key: "point4", width: 10 }
    ];
    const lookupHeader = lookupSheet.getRow(1);
    lookupHeader.font = { bold: true };
    lookupHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "E2E8F0" } };

    const lookupRows = [
      [0.0, 0.0],
      [3.0, 0.5],
      [4.0, 1.0],
      [5.0, 1.5],
      [5.5, 2.0],
      [6.5, 2.5],
      [7.0, 3.0],
      [8.0, 3.5],
      [8.5, 4.0]
    ];
    for (const [threshold10, point4] of lookupRows) {
      const row = lookupSheet.addRow({ threshold10, point4 });
      row.getCell(1).numFmt = "0.00";
      row.getCell(2).numFmt = "0.00";
    }
    autoFitWorksheetColumns(lookupSheet, 10, 18);

    const buffer = await workbook.xlsx.writeBuffer();
    downloadBlob(
      fileName,
      new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      })
    );
  }

  function normalizeGradeExportOptions(rawOptions) {
    const source = rawOptions && typeof rawOptions === "object" ? rawOptions : {};
    const format = text(source.format).toLowerCase();
    if (format === "csv" || format === "json" || format === "xlsx") {
      return { format };
    }
    return { format: "xlsx" };
  }

  async function executeGradeExport(options) {
    if (exportInProgress) {
      throw new Error("Đang có tác vụ xuất khác. Vui lòng thử lại sau.");
    }

    const context = getGradeExportContext();
    if (!context.isGradePage || !context.hasTable) {
      throw new Error("Vui lòng mở trang Kết quả học tập để xuất.");
    }

    const normalizedOptions = normalizeGradeExportOptions(options);
    if (normalizedOptions.format === "xlsx") {
      assertExportLibraries("xlsx");
    }

    exportInProgress = true;
    try {
      const scraped = scrapeGradeRecords();
      const records = mergeRecordsWithEditorState(scraped.records);
      const webSummary = scraped.webSummary;
      const summary = calculateGradeSummary(records, webSummary);
      const baseName = buildGradeExportBaseName();

      if (normalizedOptions.format === "json") {
        const payload = {
          version: 2,
          exportedAt: new Date().toISOString(),
          sourcePage: location.href,
          records,
          summary,
          webSummary
        };
        downloadBlob(
          `${baseName}.json`,
          new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" })
        );
      } else if (normalizedOptions.format === "csv") {
        const csv = buildGradeCsvContent(records);
        downloadBlob(`${baseName}.csv`, new Blob([csv], { type: "text/csv;charset=utf-8" }));
      } else if (normalizedOptions.format === "xlsx") {
        await exportGradesAsXlsx(records, summary, webSummary, `${baseName}.xlsx`);
      } else {
        throw new Error("Định dạng xuất điểm không được hỗ trợ.");
      }

      return {
        fileBase: baseName,
        totalRecords: records.length,
        options: normalizedOptions,
        summary
      };
    } finally {
      exportInProgress = false;
    }
  }

  function ensureGradeManageButton() {
    if (!isGradePage()) return;
    ensureGradeInlineStyle();
    bindGradeInlineHandlers();

    const host = getGradeActionsHost();
    if (!host) return;

    const legacyWrap = document.getElementById("tdk-grade-wrap");
    if (legacyWrap && legacyWrap.parentElement) {
      legacyWrap.parentElement.removeChild(legacyWrap);
    }
    const legacyButton = document.getElementById("tdk-grade-open-btn");
    if (legacyButton && !legacyButton.closest("#tdk-grade-inline-controls")) {
      legacyButton.remove();
    }

    let controls = document.getElementById("tdk-grade-inline-controls");
    if (!controls) {
      controls = document.createElement("div");
      controls.id = "tdk-grade-inline-controls";
      controls.innerHTML =
        `<a id="tdk-grade-open-btn" href="javascript:;" class="btn btn-action">Bật sửa điểm inline</a>` +
        `<select id="tdk-grade-calc-scope"><option value="selected_set">Các môn đã chọn</option></select>` +
        `<a id="tdk-grade-calc-btn" href="javascript:;" class="btn btn-default">Tính TB thường kỳ</a>` +
        `<a id="tdk-grade-reset-all-btn" href="javascript:;" class="btn btn-default">Khôi phục tất cả</a>` +
        `<span id="tdk-grade-inline-status"></span>`;
      host.appendChild(controls);

      const toggleBtn = controls.querySelector("#tdk-grade-open-btn");
      const calcBtn = controls.querySelector("#tdk-grade-calc-btn");
      const calcScope = controls.querySelector("#tdk-grade-calc-scope");
      const resetBtn = controls.querySelector("#tdk-grade-reset-all-btn");
      if (toggleBtn) {
        toggleBtn.addEventListener("click", (event) => {
          event.preventDefault();
          toggleInlineGradeEditMode();
        });
      }
      if (calcBtn) {
        calcBtn.addEventListener("click", (event) => {
          event.preventDefault();
          const scope = calcScope instanceof HTMLSelectElement ? calcScope.value : "selected_set";
          applyComponentCalculationByScope(scope);
        });
      }
      if (resetBtn) {
        resetBtn.addEventListener("click", (event) => {
          event.preventDefault();
          resetAllInlineGradeEdits();
        });
      }
    }

    const state = ensureGradeInlineState(false);
    refreshGradeInlineControlState(state);
  }

  function ensureGradeEditorPanel() {
    // no-op: chỉ sửa inline trên bảng điểm gốc.
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
      node.id === "tdk-export-modal-wrap" ||
      node.id === "tdk-brand-footer" ||
      node.id === "tdk-open-btn" ||
      node.id === "tdk-export-open-btn" ||
      node.id === "tdk-grade-open-btn" ||
      node.id === "tdk-grade-calc-btn" ||
      node.id === "tdk-grade-calc-scope" ||
      node.id === "tdk-grade-reset-all-btn" ||
      node.id === "tdk-grade-inline-controls" ||
      node.id === "tdk-grade-inline-status" ||
      node.id === "tdk-grade-wrap" ||
      node.id === "tdk-grade-modal" ||
      node.classList.contains("tdk-render-root") ||
      node.classList.contains("tdk-manual-card") ||
      node.classList.contains("tdk-grade-cell-changed") ||
      node.classList.contains("tdk-grade-row-changed") ||
      node.classList.contains("tdk-grade-row-selected") ||
      node.classList.contains("tdk-grade-summary-changed") ||
      node.classList.contains("tdk-grade-inline-editable") ||
      node.classList.contains("tdk-grade-inline-input") ||
      node.classList.contains("tdk-grade-summary") ||
      node.classList.contains("tdk-grade-table-wrap")
    ) {
      return true;
    }

    return Boolean(
      node.closest(
        "#tdk-modal-wrap, #tdk-export-modal-wrap, #tdk-brand-footer, #tdk-open-btn, #tdk-export-open-btn, #tdk-grade-open-btn, #tdk-grade-calc-btn, #tdk-grade-calc-scope, #tdk-grade-reset-all-btn, #tdk-grade-inline-controls, #tdk-grade-inline-status, #tdk-grade-wrap, #tdk-grade-modal, .tdk-manual-card, .tdk-render-root, .tdk-grade-summary, .tdk-grade-table-wrap, .tdk-grade-inline-editable, .tdk-grade-inline-input, .tdk-grade-cell-changed, .tdk-grade-row-changed, .tdk-grade-row-selected, .tdk-grade-summary-changed"
      )
    );
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
        ensureExportButton();
        ensureModal();
        ensureExportModal();
        ensureBrandFooter(config);
        renderManualCards(config);
      }

      if (isGradePage()) {
        ensureGradeManageButton();
        ensureGradeInlineState(false);
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

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (!message || typeof message !== "object") return;

    if (message.type === "tdk_grade_ping") {
      sendResponse({
        ok: true,
        page: location.href,
        isGradePage: isGradePage()
      });
      return;
    }

    if (message.type === "tdk_grade_get_context") {
      sendResponse({
        ok: true,
        context: getGradeExportContext()
      });
      return;
    }

    if (message.type === "tdk_grade_export_execute") {
      void (async () => {
        try {
          if (!isGradePage()) {
            throw new Error("Tab hiện tại không phải trang điểm UNETI.");
          }

          const result = await executeGradeExport(message.options || {});
          sendResponse({
            ok: true,
            result
          });
        } catch (error) {
          sendResponse({
            ok: false,
            error: error && error.message ? error.message : "Xuất điểm thất bại."
          });
        }
      })();

      return true;
    }

    if (message.type === "tdk_export_ping") {
      sendResponse({
        ok: true,
        page: location.href,
        isSchedulePage: isSchedulePage()
      });
      return;
    }

    if (message.type === "tdk_export_get_context") {
      sendResponse({
        ok: true,
        context: getExportContext()
      });
      return;
    }

    if (message.type === "tdk_export_open_panel") {
      void (async () => {
        try {
          if (!isSchedulePage()) {
            throw new Error("Tab hiện tại không phải trang lịch UNETI.");
          }

          await openExportModal();
          sendResponse({
            ok: true
          });
        } catch (error) {
          sendResponse({
            ok: false,
            error: error && error.message ? error.message : "Không mở được panel xuất."
          });
        }
      })();

      return true;
    }

    if (message.type === "tdk_export_execute") {
      void (async () => {
        try {
          if (!isSchedulePage()) {
            throw new Error("Tab hiện tại không phải trang lịch UNETI.");
          }

          const result = await executeScheduleExport(message.options || {}, null);
          sendResponse({
            ok: true,
            result
          });
        } catch (error) {
          sendResponse({
            ok: false,
            error: error && error.message ? error.message : "Xuất lịch thất bại."
          });
        }
      })();

      return true;
    }
  });

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
