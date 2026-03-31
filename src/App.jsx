
import React, { useState, useMemo, useEffect, useRef } from 'react';
import {
  Plus, Minus, Settings, Sparkles, Loader2,
  ArrowUp, ArrowDown, Save, History as Clock, Download,
  FileSpreadsheet, FileText, X, Check, Calendar, CalendarDays,
  User, Lock, Info, Layout, ShieldCheck, Grid, UserCheck,
  Database, Cpu, Monitor, ArrowLeft, ChevronRight, ChevronDown, ChevronUp, CheckCircle2, Trash2
} from 'lucide-react';

// ==========================================
// 1. 系統代碼字典
// ==========================================
const DICT = {
  SHIFTS: ['D', 'E', 'N', '白8-8', '夜8-8', '8-12', '12-16'],
  LEAVES: ['off', '例', '休', '特', '補', '國', '喪', '婚', '產', '病', '事', '陪產', 'AM', 'PM']
};

const SMART_RULES = {
  maxConsecutiveWorkDays: 5,
  allowCrossGroupAssignment: false,
  disallowedNextShiftMap: {
    N: ['D', 'E'],
    E: ['N', 'D'],
    '白8-8': ['D', 'N'],
    '夜8-8': ['E', 'N']
  },
  blockedLeavePrefixes: ['off', '例', '休', '特', '補', '國', '喪', '婚', '產', '病', '事', '陪產', 'AM', 'PM'],
  pregnancyRestrictedShifts: ['N', '夜8-8'],
  fillPriorityWeights: {
    sameShiftCount: 3,
    totalShiftCount: 2,
    sameGroup: 1
  }
};

const SHIFT_GROUPS = ['白班', '小夜', '大夜'];

const ANNOUNCED_CALENDAR_OVERRIDES = {
  2024: {
    holidays: ['2024-01-01', '2024-02-08', '2024-02-09', '2024-02-10', '2024-02-11', '2024-02-12', '2024-02-13', '2024-02-14', '2024-02-28', '2024-04-04', '2024-04-05', '2024-06-10', '2024-09-17', '2024-10-10'],
    workdays: []
  },
  2025: {
    holidays: ['2025-01-01', '2025-01-27', '2025-01-28', '2025-01-29', '2025-01-30', '2025-01-31', '2025-02-28', '2025-04-03', '2025-04-04', '2025-05-31', '2025-10-06', '2025-10-10'],
    workdays: []
  },
  2026: {
    holidays: ['2026-01-01', '2026-02-16', '2026-02-17', '2026-02-18', '2026-02-19', '2026-02-20', '2026-02-28', '2026-04-04', '2026-04-05', '2026-06-19', '2026-09-25', '2026-10-10'],
    workdays: []
  }
};

const CHINESE_MONTH_MAP = {
  '正月': 1, '一月': 1, '二月': 2, '三月': 3, '四月': 4, '五月': 5, '六月': 6,
  '七月': 7, '八月': 8, '九月': 9, '十月': 10, '十一月': 11, '十二月': 12,
  '臘月': 12
};

const CHINESE_DAY_MAP = {
  '初一': 1, '初二': 2, '初三': 3, '初四': 4, '初五': 5, '初六': 6, '初七': 7, '初八': 8, '初九': 9, '初十': 10,
  '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15, '十六': 16, '十七': 17, '十八': 18, '十九': 19, '二十': 20,
  '廿一': 21, '廿二': 22, '廿三': 23, '廿四': 24, '廿五': 25, '廿六': 26, '廿七': 27, '廿八': 28, '廿九': 29, '三十': 30
};

const normalizeLunarMonth = (label = '') => {
  const cleaned = String(label).trim().replace(/^閏/, '').replace(/月$/, '');
  if (CHINESE_MONTH_MAP[String(label).trim().replace(/^閏/, '')]) {
    return CHINESE_MONTH_MAP[String(label).trim().replace(/^閏/, '')];
  }
  if (/^\d+$/.test(cleaned)) return Number(cleaned);
  if (CHINESE_MONTH_MAP[`${cleaned}月`]) return CHINESE_MONTH_MAP[`${cleaned}月`];
  return null;
};

const normalizeLunarDay = (label = '') => {
  const cleaned = String(label).trim();
  if (CHINESE_DAY_MAP[cleaned]) return CHINESE_DAY_MAP[cleaned];
  if (/^\d+$/.test(cleaned)) return Number(cleaned);
  return null;
};

const formatDateKey = (date) => {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

const parseDateKey = (dateKey) => {
  const [y, m, d] = dateKey.split('-').map(Number);
  return new Date(y, m - 1, d);
};

const addDays = (date, amount) => {
  const next = new Date(date);
  next.setDate(next.getDate() + amount);
  return next;
};

const isWeekendDate = (date) => date.getDay() === 0 || date.getDay() === 6;

const uniqueSortedDates = (dates = []) => Array.from(new Set(dates)).sort();

const getChineseCalendarInfo = (date) => {
  const formatter = new Intl.DateTimeFormat('zh-TW-u-ca-chinese', {
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
  const parts = formatter.formatToParts(date);
  const yearPart = parts.find((part) => part.type === 'relatedYear');
  const monthPart = parts.find((part) => part.type === 'month');
  const dayPart = parts.find((part) => part.type === 'day');
  const rawMonth = String(monthPart?.value || '').trim();
  const rawDay = String(dayPart?.value || '').trim();

  return {
    relatedYear: Number(yearPart?.value || date.getFullYear()),
    leapMonth: rawMonth.startsWith('閏'),
    monthLabel: rawMonth.replace(/^閏/, ''),
    dayLabel: rawDay,
    monthNumber: normalizeLunarMonth(rawMonth),
    dayNumber: normalizeLunarDay(rawDay)
  };
};

const matchesLunarDate = (date, lunarMonth, lunarDay, options = {}) => {
  const info = getChineseCalendarInfo(date);
  return info.monthNumber === lunarMonth && info.dayNumber === lunarDay && Boolean(info.leapMonth) === Boolean(options.leap);
};

const findGregorianDateByLunarInYear = (year, lunarMonth, lunarDay, options = {}) => {
  const start = new Date(year, 0, 1);
  const end = new Date(year, 11, 31);
  for (let cursor = new Date(start); cursor <= end; cursor = addDays(cursor, 1)) {
    if (matchesLunarDate(cursor, lunarMonth, lunarDay, options)) return new Date(cursor);
  }
  return null;
};

const getQingmingDate = (year) => {
  if (year >= 2000 && year <= 2099) {
    const y = year % 100;
    const day = Math.floor(y * 0.2422 + 4.81) - Math.floor((y - 1) / 4);
    return new Date(year, 3, day);
  }
  return new Date(year, 3, 4);
};

const getFixedSolarHolidayDates = (year) => {
  const fixed = [
    [1, 1],   // 開國紀念日
    [2, 28],  // 和平紀念日
    [4, 4],   // 兒童節
    [5, 1],   // 勞動節
    [9, 28],  // 孔子誕辰紀念日 / 教師節
    [10, 10], // 國慶日
    [10, 25], // 臺灣光復暨金門古寧頭大捷紀念日
    [12, 25], // 行憲紀念日
  ];
  return fixed.map(([month, day]) => formatDateKey(new Date(year, month - 1, day)));
};

const getRuleBasedHolidayDates = (year) => {
  const holidaySet = new Set(getFixedSolarHolidayDates(year));

  const qingming = getQingmingDate(year);
  holidaySet.add(formatDateKey(qingming));

  const dragonBoat = findGregorianDateByLunarInYear(year, 5, 5);
  if (dragonBoat) holidaySet.add(formatDateKey(dragonBoat));

  const midAutumn = findGregorianDateByLunarInYear(year, 8, 15);
  if (midAutumn) holidaySet.add(formatDateKey(midAutumn));

  const lunarNewYearDay = findGregorianDateByLunarInYear(year, 1, 1);
  if (lunarNewYearDay) {
    const springHoliday = [
      addDays(lunarNewYearDay, -2),
      addDays(lunarNewYearDay, -1),
      lunarNewYearDay,
      addDays(lunarNewYearDay, 1),
      addDays(lunarNewYearDay, 2)
    ];
    springHoliday
      .filter((date) => date.getFullYear() === year)
      .forEach((date) => holidaySet.add(formatDateKey(date)));
  }

  const childrensDayKey = formatDateKey(new Date(year, 3, 4));
  const qingmingKey = formatDateKey(qingming);
  if (childrensDayKey === qingmingKey) {
    const qingmingWeekday = qingming.getDay();
    const extraHoliday = qingmingWeekday === 4 ? addDays(qingming, 1) : addDays(qingming, -1);
    holidaySet.add(formatDateKey(extraHoliday));
  }

  return uniqueSortedDates([...holidaySet]);
};

const findNearestWorkday = (startDate, direction, occupiedHolidays, workdayOverrides) => {
  let cursor = addDays(startDate, direction);
  while (true) {
    const key = formatDateKey(cursor);
    const weekend = isWeekendDate(cursor);
    const isWorkday = (!weekend || workdayOverrides.has(key)) && !occupiedHolidays.has(key);
    if (isWorkday) return key;
    cursor = addDays(cursor, direction);
  }
};

const applyCompensatoryHolidays = (holidayDates, workdayDates = []) => {
  const holidaySet = new Set(uniqueSortedDates(holidayDates));
  const workdaySet = new Set(uniqueSortedDates(workdayDates));

  const baseDates = [...holidaySet].sort();
  baseDates.forEach((dateKey) => {
    const date = parseDateKey(dateKey);
    if (date.getDay() === 6) {
      holidaySet.add(findNearestWorkday(date, -1, holidaySet, workdaySet));
    } else if (date.getDay() === 0) {
      holidaySet.add(findNearestWorkday(date, 1, holidaySet, workdaySet));
    }
  });

  return uniqueSortedDates([...holidaySet]);
};

const getSystemHolidayCalendar = (year, options = {}) => {
  const {
    customHolidays = [],
    announcedOverrides = ANNOUNCED_CALENDAR_OVERRIDES,
    specialWorkdays = [],
    unitAdjustments = { holidays: [], workdays: [] }
  } = options;

  const announced = announcedOverrides[year];
  const overrideHolidays = announced?.holidays || [];
  const overrideWorkdays = announced?.workdays || [];

  const unitHolidayDates = (unitAdjustments.holidays || []).filter((date) => date.startsWith(`${year}-`));
  const unitWorkdayDates = (unitAdjustments.workdays || []).filter((date) => date.startsWith(`${year}-`));
  const customHolidayDates = customHolidays.filter((date) => date.startsWith(`${year}-`));
  const specialWorkdayDates = specialWorkdays.filter((date) => date.startsWith(`${year}-`));

  let holidayDates = overrideHolidays.length > 0 ? overrideHolidays : applyCompensatoryHolidays(getRuleBasedHolidayDates(year), [...overrideWorkdays, ...specialWorkdayDates, ...unitWorkdayDates]);
  let workdayDates = uniqueSortedDates([...overrideWorkdays, ...specialWorkdayDates, ...unitWorkdayDates]);

  holidayDates = uniqueSortedDates([...holidayDates, ...customHolidayDates, ...unitHolidayDates]).filter((date) => !workdayDates.includes(date));

  return {
    holidays: holidayDates,
    workdays: workdayDates
  };
};
const STORAGE_KEY = 'schedule_app_history';
const ACTIVE_DRAFT_KEY = 'schedule_app_active_draft';

// 外部套件載入：ExcelJS 用於高品質 Excel 樣式輸出
const loadExcelJS = () => {
  return new Promise((resolve) => {
    if (window.ExcelJS) return resolve(window.ExcelJS);
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js";
    script.onload = () => resolve(window.ExcelJS);
    document.head.appendChild(script);
  });
};

const loadSheetJS = () => {
  return new Promise((resolve, reject) => {
    if (window.XLSX) return resolve(window.XLSX);
    const existing = document.querySelector('script[data-sheetjs-loader="true"]');
    if (existing) {
      existing.addEventListener('load', () => resolve(window.XLSX), { once: true });
      existing.addEventListener('error', () => reject(new Error('SheetJS 載入失敗')), { once: true });
      return;
    }
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    script.dataset.sheetjsLoader = 'true';
    script.onload = () => resolve(window.XLSX);
    script.onerror = () => reject(new Error('SheetJS 載入失敗'));
    document.head.appendChild(script);
  });
};

const normalizeImportedShiftCode = (rawValue = '') => {
  const value = String(rawValue ?? '').trim();
  if (!value) return '';

  const normalizedWhitespace = value.replace(/\s+/g, '');
  const lower = normalizedWhitespace.toLowerCase();

  const directMap = {
    d: 'D',
    e: 'E',
    n: 'N',
    off: 'off',
    of: 'off',
    am: 'AM',
    pm: 'PM',
    '8-12': '8-12',
    '12-16': '12-16',
    '白8-8': '白8-8',
    '夜8-8': '夜8-8'
  };

  if (directMap[lower]) return directMap[lower];
  if (DICT.LEAVES.includes(normalizedWhitespace) || DICT.SHIFTS.includes(normalizedWhitespace)) return normalizedWhitespace;
  return value;
};

const buildMonthKey = (year, month) => `${year}-${String(month).padStart(2, '0')}`;


const normalizeManualShiftCode = (rawValue = '', allowedLeaveCodes = []) => {
  const value = String(rawValue ?? '').trim();
  if (!value) return { normalized: '', isValid: true };

  const collapsed = value.replace(/\s+/g, '');
  const lower = collapsed.toLowerCase();
  const allowedCodes = Array.from(new Set([...(DICT.SHIFTS || []), ...(allowedLeaveCodes || [])])).filter(Boolean);
  const directMap = {
    d: 'D',
    e: 'E',
    n: 'N',
    off: 'off',
    of: 'off',
    am: 'AM',
    pm: 'PM',
    '8-12': '8-12',
    '12-16': '12-16',
    '白8-8': '白8-8',
    '夜8-8': '夜8-8'
  };

  if (directMap[lower]) return { normalized: directMap[lower], isValid: allowedCodes.includes(directMap[lower]) };
  const directAllowed = allowedCodes.find(code => String(code).toLowerCase() === lower);
  if (directAllowed) return { normalized: directAllowed, isValid: true };

  return { normalized: value, isValid: false };
};

const makeCellKey = (staffId, dateStr) => `${staffId}__${dateStr}`;

const parseClipboardGrid = (text = '') => {
  const raw = String(text || '').replace(/\r/g, '');
  if (!raw.trim()) return [];
  return raw.split('\n').map(row => row.split('\t'));
};

const getSelectionGroupStaffs = (selection, staffs = []) => {
  const selectionGroup = selection?.start?.group || selection?.end?.group || '';
  if (!selectionGroup) return staffs;
  return staffs.filter((staff) => (staff.group || '白班') === selectionGroup);
};

const getRectFromSelection = (selection, staffs = [], daysInMonth = []) => {
  if (!selection?.start || !selection?.end) return null;
  const scopedStaffs = getSelectionGroupStaffs(selection, staffs);
  const staffIndexMap = new Map(scopedStaffs.map((staff, index) => [staff.id, index]));
  const dayIndexMap = new Map(daysInMonth.map((day, index) => [day.date, index]));

  const startRow = staffIndexMap.get(selection.start.staffId);
  const endRow = staffIndexMap.get(selection.end.staffId);
  const startCol = dayIndexMap.get(selection.start.dateStr);
  const endCol = dayIndexMap.get(selection.end.dateStr);

  if ([startRow, endRow, startCol, endCol].some(v => v === undefined)) return null;

  return {
    rowStart: Math.min(startRow, endRow),
    rowEnd: Math.max(startRow, endRow),
    colStart: Math.min(startCol, endCol),
    colEnd: Math.max(startCol, endCol),
    scopedStaffs
  };
};

const expandSelectionCells = (selection, staffs = [], daysInMonth = []) => {
  const rect = getRectFromSelection(selection, staffs, daysInMonth);
  if (!rect) return [];
  const cells = [];
  for (let rowIndex = rect.rowStart; rowIndex <= rect.rowEnd; rowIndex += 1) {
    for (let colIndex = rect.colStart; colIndex <= rect.colEnd; colIndex += 1) {
      const staff = rect.scopedStaffs[rowIndex];
      const day = daysInMonth[colIndex];
      if (staff && day) cells.push({ staffId: staff.id, dateStr: day.date, rowIndex, colIndex });
    }
  }
  return cells;
};

const isCellInSelectionRect = (selection, staffs = [], daysInMonth = [], staffId, dateStr) => {
  const rect = getRectFromSelection(selection, staffs, daysInMonth);
  if (!rect) return false;
  const rowIndex = rect.scopedStaffs.findIndex(staff => staff.id === staffId);
  const colIndex = daysInMonth.findIndex(day => day.date === dateStr);
  if (rowIndex === -1 || colIndex === -1) return false;
  return rowIndex >= rect.rowStart && rowIndex <= rect.rowEnd && colIndex >= rect.colStart && colIndex <= rect.colEnd;
};

const extractYearMonthCandidates = (...sources) => {
  const patterns = [
    /(\d{4})\s*年\s*(\d{1,2})\s*月/,
    /(\d{4})[\/_\-.](\d{1,2})/,
    /(\d{1,2})\s*月/
  ];

  for (const source of sources) {
    const text = String(source || '').trim();
    if (!text) continue;

    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (!match) continue;
      if (match.length >= 3 && pattern !== patterns[2]) {
        const year = Number(match[1]);
        const month = Number(match[2]);
        if (year >= 1900 && month >= 1 && month <= 12) return { year, month };
      }
      if (pattern === patterns[2]) {
        const month = Number(match[1]);
        if (month >= 1 && month <= 12) return { year: null, month };
      }
    }
  }

  return { year: null, month: null };
};

const detectImportedDayNumber = (label = '') => {
  const text = String(label ?? '').replace(/\r/g, '').trim();
  if (!text) return null;
  const firstLine = text.split('\n').map(part => part.trim()).find(Boolean) || text;
  const compact = text.replace(/\s+/g, '');
  const firstCompact = firstLine.replace(/\s+/g, '');
  const patterns = [
    /^(\d{1,2})日$/,
    /^(\d{1,2})$/,
    /^(\d{1,2})\(.+\)$/,
  ];
  for (const source of [firstCompact, compact]) {
    for (const pattern of patterns) {
      const match = source.match(pattern);
      if (match) {
        const day = Number(match[1]);
        if (day >= 1 && day <= 31) return day;
      }
    }
  }
  const looseMatch = compact.match(/^(\d{1,2})/);
  if (looseMatch) {
    const day = Number(looseMatch[1]);
    if (day >= 1 && day <= 31) return day;
  }
  return null;
};

const inferImportedGroupFromCodes = (dayMap = {}) => {
  const counts = { 白班: 0, 小夜: 0, 大夜: 0 };
  Object.values(dayMap || {}).forEach((cell) => {
    const code = typeof cell === 'object' && cell !== null ? (cell.value || '') : String(cell || '');
    const group = getShiftGroupByCode(code);
    if (group && counts[group] !== undefined) counts[group] += 1;
  });
  const ranked = Object.entries(counts).sort((a, b) => {
    if (b[1] !== a[1]) return b[1] - a[1];
    return SHIFT_GROUPS.indexOf(a[0]) - SHIFT_GROUPS.indexOf(b[0]);
  });
  return ranked[0]?.[1] > 0 ? ranked[0][0] : '';
};

const parseImportedWorksheet = ({ rows, sheetName, fileName, fallbackYear, customLeaveCodes = [] }) => {
  if (!Array.isArray(rows) || rows.length === 0) return null;

  let headerRowIndex = -1;
  let nameColumnIndex = -1;
  let groupColumnIndex = -1;
  let dayColumnPairs = [];

  for (let rowIndex = 0; rowIndex < Math.min(20, rows.length); rowIndex += 1) {
    const row = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
    let detectedNameIndex = -1;
    let detectedGroupIndex = -1;
    const detectedDayPairs = [];

    row.forEach((cellValue, columnIndex) => {
      const value = String(cellValue ?? '').trim();
      if (!value) return;
      if (value === '姓名') detectedNameIndex = columnIndex;
      if (value === '班別群組') detectedGroupIndex = columnIndex;
      const dayNumber = detectImportedDayNumber(value);
      if (dayNumber) detectedDayPairs.push({ day: dayNumber, colNumber: columnIndex });
    });

    const uniqueDayPairs = Array.from(
      new Map(detectedDayPairs.sort((a, b) => a.day - b.day).map(item => [item.day, item])).values()
    );

    if (detectedNameIndex !== -1 && uniqueDayPairs.length > 0) {
      headerRowIndex = rowIndex;
      nameColumnIndex = detectedNameIndex;
      groupColumnIndex = detectedGroupIndex;
      dayColumnPairs = uniqueDayPairs;
      break;
    }
  }

  if (headerRowIndex === -1 || nameColumnIndex === -1 || dayColumnPairs.length === 0) return null;

  const validGroups = new Set(SHIFT_GROUPS);
  const validCodes = new Set([...DICT.SHIFTS, ...DICT.LEAVES, ...(customLeaveCodes || [])]);

  const importedStaffs = [];
  const importedSchedule = {};
  const invalidMessages = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
    const rawName = String(row[nameColumnIndex] ?? '').trim();
    const rawGroup = groupColumnIndex === -1 ? '' : String(row[groupColumnIndex] ?? '').trim();

    const hasAnyContent = row.some((value) => String(value ?? '').trim() !== '');
    if (!hasAnyContent || !rawName) continue;

    const rowNumber = rowIndex + 1;
    const hasAnyDayContent = dayColumnPairs.some(({ colNumber }) => String(row[colNumber] ?? '').trim() !== '');
    if (!hasAnyDayContent) continue;

    const staffId = `import_${Date.now()}_${sheetName}_${rowNumber}`;
    importedSchedule[staffId] = {};

    dayColumnPairs.forEach(({ day, colNumber }) => {
      const rawValue = String(row[colNumber] ?? '').trim();
      if (!rawValue) return;

      const normalizedCode = normalizeImportedShiftCode(rawValue);
      if (!validCodes.has(normalizedCode)) {
        invalidMessages.push(`工作表「${sheetName}」第 ${rowNumber} 列「${rawName}」的 ${day}日 代碼「${rawValue}」無法匯入，已略過`);
        return;
      }

      importedSchedule[staffId][day] = { value: normalizedCode, source: 'manual' };
    });

    let normalizedGroup = '白班';
    if (rawGroup) {
      normalizedGroup = validGroups.has(rawGroup) ? rawGroup : '白班';
      if (!validGroups.has(rawGroup)) {
        invalidMessages.push(`工作表「${sheetName}」第 ${rowNumber} 列「${rawName}」的班別群組不是白班／小夜／大夜，已自動改為白班`);
      }
    } else {
      normalizedGroup = inferImportedGroupFromCodes(importedSchedule[staffId]);
      if (!normalizedGroup) {
        delete importedSchedule[staffId];
        invalidMessages.push(`工作表「${sheetName}」第 ${rowNumber} 列「${rawName}」沒有可判定群組的班別代碼，已略過`);
        continue;
      }
    }

    importedStaffs.push({
      id: staffId,
      name: rawName,
      group: normalizedGroup,
      pregnant: false
    });
  }

  if (importedStaffs.length === 0) return null;

  const scanTexts = [];
  const maxRowsToScan = Math.min(rows.length, 8);
  for (let r = 0; r < maxRowsToScan; r += 1) {
    const row = Array.isArray(rows[r]) ? rows[r] : [];
    for (let c = 0; c < Math.min(row.length, 8); c += 1) {
      const cellText = String(row[c] ?? '').trim();
      if (cellText) scanTexts.push(cellText);
    }
  }

  const detected = extractYearMonthCandidates(...scanTexts, sheetName, fileName);
  const month = detected.month;
  const year = detected.year || fallbackYear;

  if (!month) {
    throw new Error(`工作表「${sheetName}」無法辨識月份，請確認表頭、sheet 名稱或檔名包含幾月資訊`);
  }

  const monthKey = buildMonthKey(year, month);
  const importedScheduleByDate = Object.fromEntries(
    Object.entries(importedSchedule).map(([staffId, dayMap]) => [
      staffId,
      Object.fromEntries(
        Object.entries(dayMap || {}).map(([day, cell]) => {
          const dateKey = `${monthKey}-${String(Number(day)).padStart(2, '0')}`;
          return [dateKey, cell];
        })
      )
    ])
  );

  return {
    year,
    month,
    staffs: importedStaffs,
    scheduleData: importedScheduleByDate,
    customColumnValues: {},
    schedulingRulesText: '',
    warnings: invalidMessages,
    importMeta: {
      sourceType: 'excel',
      sourceFiles: [fileName],
      sourceSheets: [sheetName],
      importedAt: new Date().toISOString(),
      lastUpdatedAt: new Date().toISOString()
    }
  };
};

const parseImportedExcelFiles = async (files = [], fallbackYear = new Date().getFullYear(), options = {}) => {
  const XLSX = await loadSheetJS();
  const fileList = Array.from(files || []);
  const monthlySchedules = {};
  const warnings = [];
  let firstMonthKey = '';

  for (const file of fileList) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet || !worksheet['!ref']) continue;

      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        raw: false,
        defval: ''
      });

      try {
        const parsed = parseImportedWorksheet({
          rows,
          sheetName,
          fileName: file.name,
          fallbackYear
        });

        if (!parsed) continue;

        const monthKey = buildMonthKey(parsed.year, parsed.month);
        const existing = monthlySchedules[monthKey];
        if (existing) {
          monthlySchedules[monthKey] = {
            ...parsed,
            importMeta: {
              ...parsed.importMeta,
              sourceFiles: Array.from(new Set([...(existing.importMeta?.sourceFiles || []), file.name])),
              sourceSheets: Array.from(new Set([...(existing.importMeta?.sourceSheets || []), sheetName]))
            }
          };
          warnings.push(`月份 ${parsed.year}年${parsed.month}月 重複匯入，已以最後讀取的內容覆蓋`);
        } else {
          monthlySchedules[monthKey] = parsed;
          if (!firstMonthKey) firstMonthKey = monthKey;
        }

        warnings.push(...(parsed.warnings || []));
      } catch (error) {
        warnings.push(error?.message || `檔案「${file.name}」工作表「${sheetName}」無法匯入`);
      }
    }
  }

  const keys = Object.keys(monthlySchedules).sort();
  if (keys.length === 0) {
    throw new Error('匯入失敗：找不到可匯入的月份資料，請確認檔案至少包含「姓名」與日期欄（可為 1日~31日，或系統匯出格式的 1\n(六) 這類表頭）；班別群組欄位可省略');
  }

  return {
    monthlySchedules,
    firstMonthKey: firstMonthKey || keys[0],
    warnings
  };
};



const normalizeStaffGroup = (staffList = []) => {
  if (!Array.isArray(staffList) || staffList.length === 0) return [];

  const fallbackGroups = [
    '白班', '白班', '白班', '白班', '白班',
    '小夜', '小夜', '小夜', '小夜', '小夜',
    '大夜', '大夜', '大夜', '大夜', '大夜'
  ];

  return staffList.map((staff, index) => ({
    ...staff,
    pregnant: Boolean(staff.pregnant),
    group: SHIFT_GROUPS.includes(staff.group) ? staff.group : (fallbackGroups[index] || '白班')
  }));
};

const createBlankMonthStaffs = (targetYear, targetMonth) => {
  const monthKey = buildMonthKey(targetYear, targetMonth);
  const blankStaffs = Array.from({ length: 15 }, (_, index) => ({
    id: `blank_${monthKey}_${index + 1}`,
    name: '新成員',
    group: SHIFT_GROUPS[Math.floor(index / 5)] || '白班',
    pregnant: false
  }));
  return normalizeStaffGroup(blankStaffs);
};

const createBlankMonthState = (targetYear, targetMonth) => {
  const blankStaffs = createBlankMonthStaffs(targetYear, targetMonth);
  const blankSchedule = blankStaffs.reduce((acc, staff) => {
    acc[staff.id] = {};
    return acc;
  }, {});
  return {
    staffs: blankStaffs,
    schedule: blankSchedule,
    customColumnValues: {},
    schedulingRulesText: ''
  };
};


const getCodePrefix = (rawCode = '') => {
  const code = String(rawCode || '').trim();
  if (!code) return '';
  if (code === 'off') return 'off';
  const direct = SMART_RULES.blockedLeavePrefixes.find((prefix) => code === prefix || code.startsWith(prefix));
  if (direct) return direct;
  return code;
};

const getShiftGroupByCode = (code = '') => {
  if (['D', '白8-8', '8-12', '12-16'].includes(code)) return '白班';
  if (['E', '夜8-8'].includes(code)) return '小夜';
  if (['N'].includes(code)) return '大夜';
  return null;
};

const isLeaveCode = (code = '') => SMART_RULES.blockedLeavePrefixes.includes(getCodePrefix(code));
const isShiftCode = (code = '') => DICT.SHIFTS.includes(code);

const GROUP_TO_DEMAND_KEY = {
  '白班': 'white',
  '小夜': 'evening',
  '大夜': 'night'
};

const DEFAULT_SHIFT_BY_GROUP = {
  '白班': 'D',
  '小夜': 'E',
  '大夜': 'N'
};

const RULE_FILL_MAIN_SHIFTS = ['D', 'E', 'N'];

const HOSPITAL_LEVEL_LABELS = {
  medical: '醫學中心',
  regional: '區域醫院',
  local: '地區醫院'
};

const HOSPITAL_RATIO_HINTS = {
  medical: { white: '1:6', evening: '1:9', night: '1:11' },
  regional: { white: '1:7', evening: '1:11', night: '1:13' },
  local: { white: '1:10', evening: '1:13', night: '1:15' }
};


const UI_FONT_SIZE_OPTIONS = {
  small: { label: '小', className: 'text-xs', shiftLabelSize: '1.45rem', shiftCellLabelSize: '1.55rem' },
  medium: { label: '標準', className: 'text-sm', shiftLabelSize: '1.95rem', shiftCellLabelSize: '2.05rem' },
  large: { label: '大', className: 'text-base', shiftLabelSize: '2.45rem', shiftCellLabelSize: '2.55rem' }
};

const getUiFontSizeClass = (sizeKey = 'medium') => UI_FONT_SIZE_OPTIONS[sizeKey]?.className || UI_FONT_SIZE_OPTIONS.medium.className;
const getShiftLabelFontSize = (sizeKey = 'medium') => UI_FONT_SIZE_OPTIONS[sizeKey]?.shiftLabelSize || UI_FONT_SIZE_OPTIONS.medium.shiftLabelSize;
const getShiftCellLabelFontSize = (sizeKey = 'medium') => UI_FONT_SIZE_OPTIONS[sizeKey]?.shiftCellLabelSize || UI_FONT_SIZE_OPTIONS.medium.shiftCellLabelSize;

const UI_DENSITY_OPTIONS = {
  compact: {
    shiftWidth: 58,
    nameWidth: 84,
    dayMinWidth: 32,
    dayHeaderClass: 'px-0.5 py-1 text-[11px]',
    statHeaderClass: 'p-1.5',
    leaveHeaderClass: 'p-1',
    cellHeightClass: 'h-8',
    nameCellPaddingClass: 'px-0.5 py-0.5',
    footCellPaddingClass: 'p-1.5',
    groupLabelClass: '',
    selectorDotClass: 'w-1.5 h-1.5',
    rowMinHeight: 72
  },
  standard: {
    shiftWidth: 68,
    nameWidth: 84,
    dayMinWidth: 52,
    dayHeaderClass: 'px-1.5 py-2 text-xs',
    statHeaderClass: 'px-1 py-1',
    leaveHeaderClass: 'px-1 py-1.5',
    cellHeightClass: 'h-9',
    nameCellPaddingClass: 'px-1 py-1',
    footCellPaddingClass: 'px-1 py-1',
    groupLabelClass: '',
    selectorDotClass: 'w-2 h-2',
    rowMinHeight: 72
  },
  relaxed: {
    shiftWidth: 100,
    nameWidth: 156,
    dayMinWidth: 68,
    dayHeaderClass: 'px-1 py-1.5 text-sm',
    statHeaderClass: 'p-4',
    leaveHeaderClass: 'p-2',
    cellHeightClass: 'h-12',
    nameCellPaddingClass: 'px-2 py-2',
    footCellPaddingClass: 'p-3',
    groupLabelClass: '',
    selectorDotClass: 'w-3 h-3',
    rowMinHeight: 96
  }
};

const getUiDensityConfig = (densityKey = 'standard') => UI_DENSITY_OPTIONS[densityKey] || UI_DENSITY_OPTIONS.standard;


const UI_THEME_PRESETS = {
  classic: {
    pageBackgroundColor: '#f8fafc',
    weekendColor: '#dcfce7',
    holidayColor: '#fca5a5',
    tableFontColor: '#1f2937',
    shiftColumnBgColor: '#ffffff',
    nameDateColumnBgColor: '#ffffff',
    shiftColumnFontColor: '#1e293b',
    nameDateColumnFontColor: '#1e293b',
    demandOverColor: '#fde68a'
  },
  soft: {
    pageBackgroundColor: '#f7faf7',
    weekendColor: '#e7f7ec',
    holidayColor: '#f6c7c7',
    tableFontColor: '#334155',
    shiftColumnBgColor: '#f7fbf8',
    nameDateColumnBgColor: '#fcfdfc',
    shiftColumnFontColor: '#365314',
    nameDateColumnFontColor: '#334155',
    demandOverColor: '#fde68a'
  },
  warm: {
    pageBackgroundColor: '#fffaf5',
    weekendColor: '#fef3c7',
    holidayColor: '#fecaca',
    tableFontColor: '#44403c',
    shiftColumnBgColor: '#fff7ed',
    nameDateColumnBgColor: '#fffbeb',
    shiftColumnFontColor: '#7c2d12',
    nameDateColumnFontColor: '#44403c',
    demandOverColor: '#fdba74'
  },
  dark: {
    pageBackgroundColor: '#0f172a',
    weekendColor: '#334155',
    holidayColor: '#7f1d1d',
    tableFontColor: '#e2e8f0',
    shiftColumnBgColor: '#1e293b',
    nameDateColumnBgColor: '#172033',
    shiftColumnFontColor: '#f8fafc',
    nameDateColumnFontColor: '#e2e8f0',
    demandOverColor: '#78350f'
  }
};

const WIDTH_ADJUST_MAP = { narrow: -12, standard: 0, wide: 12 };
const HEIGHT_ADJUST_MAP = { compact: -4, standard: 0, roomy: 4 };
const getAdjustedDensityConfig = (baseConfig, uiSettings = {}) => {
  const shiftAdjust = WIDTH_ADJUST_MAP[uiSettings.shiftColumnWidthMode || 'standard'] || 0;
  const nameAdjust = WIDTH_ADJUST_MAP[uiSettings.nameDateColumnWidthMode || 'standard'] || 0;
  const dayAdjust = WIDTH_ADJUST_MAP[uiSettings.dayColumnWidthMode || 'standard'] || 0;
  const heightAdjust = HEIGHT_ADJUST_MAP[uiSettings.cellHeightMode || 'standard'] || 0;
  const dotClassMap = { compact: 'w-1.5 h-1.5', standard: 'w-2 h-2', roomy: 'w-2.5 h-2.5' };
  return {
    ...baseConfig,
    shiftWidth: Math.max(48, baseConfig.shiftWidth + shiftAdjust),
    nameWidth: Math.max(76, baseConfig.nameWidth + nameAdjust),
    dayMinWidth: Math.max(28, baseConfig.dayMinWidth + dayAdjust),
    rowMinHeight: Math.max(72, (baseConfig.rowMinHeight || 80) + heightAdjust * 4),
    selectorDotClass: dotClassMap[uiSettings.cellHeightMode || 'standard'] || baseConfig.selectorDotClass
  };
};

const clampColorChannel = (value) => Math.max(0, Math.min(255, Math.round(value)));

const normalizeHexColor = (hex, fallback = '#000000') => {
  const raw = String(hex || '').trim();
  if (/^#[0-9a-fA-F]{6}$/.test(raw)) return raw;
  if (/^#[0-9a-fA-F]{3}$/.test(raw)) {
    return `#${raw[1]}${raw[1]}${raw[2]}${raw[2]}${raw[3]}${raw[3]}`;
  }
  return fallback;
};

const hexToRgbObject = (hex, fallback = '#000000') => {
  const normalized = normalizeHexColor(hex, fallback).replace('#', '');
  return {
    r: parseInt(normalized.slice(0, 2), 16),
    g: parseInt(normalized.slice(2, 4), 16),
    b: parseInt(normalized.slice(4, 6), 16)
  };
};

const rgbObjectToHex = ({ r, g, b }) => {
  const toHex = (value) => clampColorChannel(value).toString(16).padStart(2, '0');
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
};

const blendHexColors = (baseHex, mixHex, mixRatio = 0.5) => {
  const ratio = Math.max(0, Math.min(1, Number(mixRatio) || 0));
  const base = hexToRgbObject(baseHex, '#ffffff');
  const mix = hexToRgbObject(mixHex, '#ffffff');
  return rgbObjectToHex({
    r: base.r * (1 - ratio) + mix.r * ratio,
    g: base.g * (1 - ratio) + mix.g * ratio,
    b: base.b * (1 - ratio) + mix.b * ratio
  });
};

const hexToExcelArgb = (hex, fallback = '#FFFFFF') => {
  return `FF${normalizeHexColor(hex, fallback).replace('#', '').toUpperCase()}`;
};

const FOUR_WEEK_CYCLE_START = '2026-04-13';
const FOUR_WEEK_CYCLE_DAYS = 28;

const isFourWeekCycleEndDate = (dateStr, cycleStart = FOUR_WEEK_CYCLE_START) => {
  if (!dateStr) return false;
  const target = parseDateKey(dateStr);
  const start = parseDateKey(cycleStart);
  target.setHours(0, 0, 0, 0);
  start.setHours(0, 0, 0, 0);
  const diffDays = Math.floor((target.getTime() - start.getTime()) / 86400000);
  const cycleOffset = ((diffDays + 1) % FOUR_WEEK_CYCLE_DAYS + FOUR_WEEK_CYCLE_DAYS) % FOUR_WEEK_CYCLE_DAYS;
  return cycleOffset === 0;
};


function ScheduleView({ changeScreen, colors, setColors, customHolidays, setCustomHolidays, specialWorkdays, setSpecialWorkdays, medicalCalendarAdjustments, setMedicalCalendarAdjustments, staffingConfig, setStaffingConfig, uiSettings, setUiSettings, customLeaveCodes, setCustomLeaveCodes, customColumns, setCustomColumns, customColumnValues, setCustomColumnValues, schedulingRulesText, setSchedulingRulesText, loadLatestOnEnter, onLatestLoaded, importedSchedulePayload, onImportedScheduleApplied, monthlySchedules, setMonthlySchedules, pendingOpenMonthKey, onPendingOpenHandled, year, setYear, month, setMonth, staffs, setStaffs, schedule, setSchedule, onDownloadDraftFile, onImportDraftFileClick, draftImportInputRef, onImportDraftFileChange }) {
  // ==========================================
  // 2. 核心 State 定義
  // ==========================================
  const [isRuleFillLoading, setIsRuleFillLoading] = useState(false);


  const [unitAdjustmentDraft, setUnitAdjustmentDraft] = useState({ holidays: [], workdays: [] });

  const [showRuleFillControl, setShowRuleFillControl] = useState(false);
  const [ruleFillFeedback, setRuleFillFeedback] = useState("");

  const [showHistoryModal, setShowHistoryModal] = useState(false);
  const [historyList, setHistoryList] = useState([]);
  const [showDraftPrompt, setShowDraftPrompt] = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [selectedFillCell, setSelectedFillCell] = useState(null);
  const [fillCandidates, setFillCandidates] = useState([]);
  const [showFillModal, setShowFillModal] = useState(false);
  const [selectedGridCell, setSelectedGridCell] = useState(null);
  const [rangeClearMode, setRangeClearMode] = useState('autoOnly');
  const [cellDrafts, setCellDrafts] = useState({});
  const [invalidCellKeys, setInvalidCellKeys] = useState({});
  const [rangeSelection, setRangeSelection] = useState(null);
  const [selectionAnchor, setSelectionAnchor] = useState(null);
  const [isRangeDragging, setIsRangeDragging] = useState(false);
  const [clipboardGrid, setClipboardGrid] = useState([]);
  const [keyInputBuffer, setKeyInputBuffer] = useState('');
  const [collapsedGroups, setCollapsedGroups] = useState({ 白班: false, 小夜: false, 大夜: false });
  const [editingStaffId, setEditingStaffId] = useState(null);
  const [editingNameDraft, setEditingNameDraft] = useState('');

  // 規則補空指定設定
  const [ruleFillConfig, setRuleFillConfig] = useState({
    selectedStaffs: [],
    dateRange: { start: 1, end: 31 },
    targetShift: ''
  });

  const tableFontSizeClass = getUiFontSizeClass(uiSettings?.tableFontSize);
  const shiftColumnFontSizeClass = getUiFontSizeClass(uiSettings?.shiftColumnFontSize);
  const nameDateColumnFontSizeClass = getUiFontSizeClass(uiSettings?.nameDateColumnFontSize);
  const shiftLabelFontSize = getShiftLabelFontSize(uiSettings?.shiftColumnFontSize);
  const shiftCellLabelFontSize = getShiftCellLabelFontSize(uiSettings?.shiftColumnFontSize);
  const densityConfig = getAdjustedDensityConfig(getUiDensityConfig(uiSettings?.tableDensity), uiSettings);
  const dynamicNameWidth = useMemo(() => {
    const longestNameLength = Math.max(
      ...((staffs || []).map((staff) => String(staff?.name || '').trim().length || 0)),
      0
    );
    const safeLength = Math.max(1, longestNameLength);
    const compactControlSpace = 54;
    const estimatedNameWidth = compactControlSpace + (safeLength * 16);
    return Math.max(96, Math.min(196, estimatedNameWidth));
  }, [staffs]);
  const effectiveDensityConfig = useMemo(() => ({
    ...densityConfig,
    nameWidth: dynamicNameWidth
  }), [densityConfig, dynamicNameWidth]);
  const monthLoadSkipRef = useRef(false);
  const initializedMonthRef = useRef(false);
  const monthSwitchSeedRef = useRef('');
  const keyInputTimerRef = useRef(null);

  const pageBackgroundColor = uiSettings?.pageBackgroundColor || '#f8fafc';
  const tableFontColor = uiSettings?.tableFontColor || '#1f2937';
  const shiftColumnFontColor = uiSettings?.shiftColumnFontColor || '#1e293b';
  const nameDateColumnFontColor = uiSettings?.nameDateColumnFontColor || '#1e293b';
  const shiftColumnBgColor = uiSettings?.shiftColumnBgColor || '#ffffff';
  const nameDateColumnBgColor = uiSettings?.nameDateColumnBgColor || '#ffffff';
  const demandOverColor = uiSettings?.demandOverColor || '#fde68a';
  const stickyGroupSummaryTop = 44;
  const stickyGroupSummaryShadow = '0 6px 12px rgba(15, 23, 42, 0.08)';
  const fourWeekDividerBaseColor = nameDateColumnFontColor || shiftColumnFontColor || tableFontColor || '#1e293b';
  const fourWeekDividerColor = blendHexColors(fourWeekDividerBaseColor, pageBackgroundColor, 0.18);

  const getFourWeekDividerStyle = (dateStr) => (
    isFourWeekCycleEndDate(dateStr)
      ? { boxShadow: `inset -3px 0 0 ${fourWeekDividerColor}` }
      : null
  );

  const getWordCycleDividerStyle = (dateStr) => (
    isFourWeekCycleEndDate(dateStr)
      ? `border-right:3pt solid ${fourWeekDividerColor};mso-border-right-alt:3pt solid ${fourWeekDividerColor};`
      : ''
  );

  const applyExcelFourWeekDivider = (border = {}, dateStr) => {
    if (!isFourWeekCycleEndDate(dateStr)) return border;
    return {
      ...border,
      right: {
        style: 'thick',
        color: { argb: hexToExcelArgb(fourWeekDividerColor, '#64748B') }
      }
    };
  };
  const showRightStats = uiSettings?.showRightStats ?? uiSettings?.showStats ?? true;
  const showLeaveStats = uiSettings?.showLeaveStats ?? uiSettings?.showStats ?? true;
  const showBottomStats = uiSettings?.showBottomStats ?? true;
  const showBlueDots = uiSettings?.showBlueDots ?? true;
  const showShiftLabels = uiSettings?.showShiftLabels ?? true;
  const defaultAutoLeaveCode = uiSettings?.defaultAutoLeaveCode || 'off';
  const selectionMode = uiSettings?.selectionMode || 'dot';
  const mergedLeaveCodes = useMemo(() => Array.from(new Set([...DICT.LEAVES, ...(customLeaveCodes || [])])).filter(Boolean), [customLeaveCodes]);
  const isConfiguredLeaveCode = (code = '') => {
    if (!code) return false;
    const prefix = getCodePrefix(code);
    return mergedLeaveCodes.includes(code) || mergedLeaveCodes.includes(prefix);
  };

  // ==========================================
  // 3. 初始載入與自動帶入
  // ==========================================
  useEffect(() => {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
      try {
        const parsed = JSON.parse(stored);
        if (parsed && parsed.length > 0) {
          setHistoryList(parsed);
          setShowDraftPrompt(true);
        }
      } catch (e) {
        console.error("歷史紀錄解析失敗");
      }
    }
  }, []);

  useEffect(() => {
    if (!loadLatestOnEnter) return;

    const stored = localStorage.getItem(STORAGE_KEY);
    if (!stored) {
      onLatestLoaded?.();
      return;
    }

    try {
      const parsed = JSON.parse(stored);
      if (parsed && parsed.length > 0) {
        setHistoryList(parsed);
        loadHistory(parsed[0]);
      }
    } catch (e) {
      console.error("自動載入最新歷史紀錄失敗");
    } finally {
      onLatestLoaded?.();
    }
  }, [loadLatestOnEnter, onLatestLoaded]);

  useEffect(() => {
    if (!pendingOpenMonthKey) return;
    const [nextYear, nextMonth] = String(pendingOpenMonthKey).split('-').map(Number);
    if (!Number.isFinite(nextYear) || !Number.isFinite(nextMonth)) {
      onPendingOpenHandled?.();
      return;
    }
    monthSwitchSeedRef.current = pendingOpenMonthKey;
    setYear(nextYear);
    setMonth(nextMonth);
    onPendingOpenHandled?.();
  }, [pendingOpenMonthKey, onPendingOpenHandled]);

  useEffect(() => {
    if (!importedSchedulePayload || !importedSchedulePayload.monthlySchedules) return;

    const mergedSchedules = {
      ...(monthlySchedules || {}),
      ...importedSchedulePayload.monthlySchedules
    };

    setMonthlySchedules(mergedSchedules);

    const totalMonths = Object.keys(importedSchedulePayload.monthlySchedules || {}).length;
    const targetMonthKey = pendingOpenMonthKey || importedSchedulePayload.firstMonthKey || buildMonthKey(year, month);
    const [targetYear, targetMonth] = String(targetMonthKey).split('-').map(Number);

    if (Number.isFinite(targetYear) && Number.isFinite(targetMonth)) {
      monthSwitchSeedRef.current = targetMonthKey;
      if (year !== targetYear) setYear(targetYear);
      if (month !== targetMonth) setMonth(targetMonth);
      loadMonthState(targetYear, targetMonth, mergedSchedules);
      initializedMonthRef.current = true;
    }

    if (Array.isArray(importedSchedulePayload.warnings) && importedSchedulePayload.warnings.length > 0) {
      setRuleFillFeedback(`✅ 匯入完成，共載入 ${totalMonths} 個月份；另有 ${importedSchedulePayload.warnings.length} 筆資料已自動略過、修正或覆蓋`);
    } else {
      setRuleFillFeedback(`✅ 匯入完成，共載入 ${totalMonths} 個月份`);
    }
    onImportedScheduleApplied?.();
  }, [importedSchedulePayload, monthlySchedules, onImportedScheduleApplied, pendingOpenMonthKey, setMonthlySchedules, year, month]);

  useEffect(() => {
    const currentKey = buildMonthKey(year, month);
    if (!initializedMonthRef.current) {
      loadMonthState(year, month);
      initializedMonthRef.current = true;
      monthSwitchSeedRef.current = currentKey;
      return;
    }

    if (monthSwitchSeedRef.current !== currentKey) {
      loadMonthState(year, month);
      monthSwitchSeedRef.current = currentKey;
    }
  }, [year, month]);

  useEffect(() => {
    if (monthLoadSkipRef.current) return;
    const monthKey = buildMonthKey(year, month);
    setMonthlySchedules(prev => ({
      ...prev,
      [monthKey]: {
        ...(prev?.[monthKey] || {}),
        year,
        month,
        staffs: normalizeStaffGroup(staffs),
        scheduleData: schedule,
        customColumnValues: customColumnValues || {},
        schedulingRulesText: typeof schedulingRulesText === 'string' ? schedulingRulesText : '',
        importMeta: {
          ...(prev?.[monthKey]?.importMeta || {}),
          sourceType: prev?.[monthKey]?.importMeta?.sourceType || 'manual',
          sourceFiles: prev?.[monthKey]?.importMeta?.sourceFiles || [],
          sourceSheets: prev?.[monthKey]?.importMeta?.sourceSheets || [],
          importedAt: prev?.[monthKey]?.importMeta?.importedAt || new Date().toISOString(),
          lastUpdatedAt: new Date().toISOString()
        }
      }
    }));
  }, [year, month, staffs, schedule, customColumnValues, schedulingRulesText, setMonthlySchedules]);

  const holidayCalendar = useMemo(() => {
    return getSystemHolidayCalendar(year, {
      customHolidays,
      specialWorkdays,
      unitAdjustments: medicalCalendarAdjustments
    });
  }, [year, customHolidays, specialWorkdays, medicalCalendarAdjustments]);

  const holidays = holidayCalendar.holidays;
  const workdays = holidayCalendar.workdays;

  const daysInMonth = useMemo(() => {
    const days = [];
    const daysCount = new Date(year, month, 0).getDate();
    const weekNames = ['日', '一', '二', '三', '四', '五', '六'];
    const holidaySet = new Set(holidays);
    const workdaySet = new Set(workdays);

    for (let i = 1; i <= daysCount; i++) {
      const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
      const weekNum = new Date(year, month - 1, i).getDay();
      const rawWeekend = weekNum === 0 || weekNum === 6;
      const isAdjustedWorkday = workdaySet.has(dateStr);
      days.push({
        day: i,
        date: dateStr,
        weekStr: weekNames[weekNum],
        isWeekend: rawWeekend && !isAdjustedWorkday,
        isHoliday: holidaySet.has(dateStr),
        isAdjustedWorkday
      });
    }
    return days;
  }, [year, month, holidays, workdays]);

  const requiredLeaves = useMemo(
    () => daysInMonth.filter(d => d.isWeekend || d.isHoliday).length,
    [daysInMonth]
  );

  const createBlankScheduleForStaffs = (staffList = []) => {
    return staffList.reduce((acc, staff) => {
      acc[staff.id] = {};
      return acc;
    }, {});
  };

  const loadMonthState = (targetYear, targetMonth, schedulesSource = monthlySchedules) => {
    const monthKey = buildMonthKey(targetYear, targetMonth);
    const monthData = schedulesSource?.[monthKey];
    monthLoadSkipRef.current = true;

    if (monthData) {
      const normalizedMonthStaffs = normalizeStaffGroup(monthData.staffs || []);
      const legacyScheduleByDay = monthData.scheduleByDay || {};
      const rebuiltScheduleData = monthData.scheduleData || Object.fromEntries(
        Object.entries(legacyScheduleByDay).map(([staffId, dayMap]) => [
          staffId,
          Object.fromEntries(
            Object.entries(dayMap || {}).map(([day, cell]) => {
              const dateKey = `${monthKey}-${String(Number(day)).padStart(2, '0')}`;
              return [dateKey, cell];
            })
          )
        ])
      );

      setStaffs(normalizedMonthStaffs);
      setSchedule(rebuiltScheduleData || createBlankScheduleForStaffs(normalizedMonthStaffs));
      setCustomColumnValues(monthData.customColumnValues || {});
      setSchedulingRulesText(typeof monthData.schedulingRulesText === 'string' ? monthData.schedulingRulesText : '');
    } else {
      const blankMonthState = createBlankMonthState(targetYear, targetMonth);
      setStaffs(blankMonthState.staffs);
      setSchedule(blankMonthState.schedule);
      setCustomColumnValues(blankMonthState.customColumnValues);
      setSchedulingRulesText(blankMonthState.schedulingRulesText);
    }

    setSelectedGridCell(null);
    setRangeSelection(null);
    setSelectionAnchor(null);
    setCellDrafts({});
    setInvalidCellKeys({});
    setKeyInputBuffer('');
    setEditingStaffId(null);
    setEditingNameDraft('');
    setTimeout(() => {
      monthLoadSkipRef.current = false;
    }, 0);
  };

  const clearInvalidCellLater = (cellKey) => {
    window.setTimeout(() => {
      setInvalidCellKeys(prev => {
        const next = { ...prev };
        delete next[cellKey];
        return next;
      });
    }, 1200);
  };

  const resetKeyInputBuffer = () => {
    setKeyInputBuffer('');
    if (keyInputTimerRef.current) {
      window.clearTimeout(keyInputTimerRef.current);
      keyInputTimerRef.current = null;
    }
  };

  const keepKeyInputBufferAlive = () => {
    if (keyInputTimerRef.current) window.clearTimeout(keyInputTimerRef.current);
    keyInputTimerRef.current = window.setTimeout(() => {
      setKeyInputBuffer('');
      keyInputTimerRef.current = null;
    }, 1500);
  };

  const getEffectiveSelection = () => {
    if (rangeSelection?.start && rangeSelection?.end) return rangeSelection;
    if (selectedGridCell?.staff?.id && selectedGridCell?.dateStr) {
      return {
        start: { staffId: selectedGridCell.staff.id, dateStr: selectedGridCell.dateStr },
        end: { staffId: selectedGridCell.staff.id, dateStr: selectedGridCell.dateStr }
      };
    }
    return null;
  };

  const selectedRangeCells = useMemo(
    () => expandSelectionCells(getEffectiveSelection(), staffs, daysInMonth),
    [rangeSelection, selectedGridCell, staffs, daysInMonth]
  );

  const clearSelectionContents = () => {
    if (selectedRangeCells.length === 0) return false;
    setSchedule(prev => {
      const next = JSON.parse(JSON.stringify(prev));
      selectedRangeCells.forEach(({ staffId, dateStr }) => {
        if (!next[staffId]) next[staffId] = {};
        next[staffId][dateStr] = null;
      });
      return next;
    });
    resetKeyInputBuffer();
    return true;
  };

  const applyValueToCells = (cells, normalized) => {
    if (!cells || cells.length === 0) return false;
    setSchedule(prev => {
      const next = JSON.parse(JSON.stringify(prev));
      cells.forEach(({ staffId, dateStr }) => {
        if (!next[staffId]) next[staffId] = {};
        next[staffId][dateStr] = normalized ? { value: normalized, source: 'manual' } : null;
      });
      return next;
    });
    return true;
  };

  const tryApplyBufferedCode = (buffer) => {
    if (!buffer || selectedRangeCells.length === 0) return false;
    const { normalized, isValid } = normalizeManualShiftCode(buffer, mergedLeaveCodes);
    if (!isValid) return false;
    applyValueToCells(selectedRangeCells, normalized);
    resetKeyInputBuffer();
    return true;
  };

  const commitCellValue = (staffId, dateStr, rawValue) => {
    const cellKey = makeCellKey(staffId, dateStr);
    const { normalized, isValid } = normalizeManualShiftCode(rawValue, mergedLeaveCodes);

    if (!isValid) {
      setInvalidCellKeys(prev => ({ ...prev, [cellKey]: true }));
      clearInvalidCellLater(cellKey);
      setCellDrafts(prev => {
        const next = { ...prev };
        delete next[cellKey];
        return next;
      });
      return false;
    }

    handleCellChange(staffId, dateStr, normalized);
    setCellDrafts(prev => {
      const next = { ...prev };
      delete next[cellKey];
      return next;
    });
    setInvalidCellKeys(prev => {
      const next = { ...prev };
      delete next[cellKey];
      return next;
    });
    return true;
  };

  const startRangeSelection = (staff, dateStr, event = {}) => {
    const mouseButton = typeof event.button === 'number' ? event.button : 0;
    if (mouseButton !== 0) return;

    const point = { staffId: staff.id, dateStr, group: staff.group || '白班' };
    if (event.shiftKey && selectionAnchor) {
      setRangeSelection({ start: selectionAnchor, end: point });
    } else {
      setSelectionAnchor(point);
      setRangeSelection({ start: point, end: point });
    }
    setSelectedGridCell({ staff, dateStr });
    resetKeyInputBuffer();
  };

  const updateRangeSelection = (staff, dateStr) => {
    if (!isRangeDragging || !selectionAnchor) return;
    const anchorGroup = selectionAnchor.group || '白班';
    const targetGroup = staff.group || '白班';
    if (anchorGroup !== targetGroup) return;
    setRangeSelection({ start: selectionAnchor, end: { staffId: staff.id, dateStr, group: targetGroup } });
  };

  const copySelectionToClipboard = async () => {
    const selection = getEffectiveSelection();
    const rect = getRectFromSelection(selection, staffs, daysInMonth);
    if (!rect) return;

    const grid = [];
    for (let rowIndex = rect.rowStart; rowIndex <= rect.rowEnd; rowIndex += 1) {
      const row = [];
      const staff = rect.scopedStaffs[rowIndex];
      for (let colIndex = rect.colStart; colIndex <= rect.colEnd; colIndex += 1) {
        const day = daysInMonth[colIndex];
        row.push(getCellCode(staff?.id, day?.date) || '');
      }
      grid.push(row);
    }

    setClipboardGrid(grid);
    const text = grid.map(row => row.join('\t')).join('\n');
    try {
      if (navigator?.clipboard?.writeText) await navigator.clipboard.writeText(text);
    } catch (error) {
      console.error('寫入剪貼簿失敗', error);
    }
  };

  const pasteGridToSelection = async () => {
    const selection = getEffectiveSelection();
    const rect = getRectFromSelection(selection, staffs, daysInMonth);
    if (!rect) return;

    let grid = clipboardGrid;
    if (!grid || grid.length === 0) {
      try {
        if (navigator?.clipboard?.readText) {
          const text = await navigator.clipboard.readText();
          grid = parseClipboardGrid(text);
        }
      } catch (error) {
        console.error('讀取剪貼簿失敗', error);
      }
    }
    if (!grid || grid.length === 0) return;

    const selectionRowCount = rect.rowEnd - rect.rowStart + 1;
    const selectionColCount = rect.colEnd - rect.colStart + 1;
    const sourceRowCount = grid.length;
    const sourceColCount = Math.max(...grid.map(row => (row || []).length), 0);

    let targetRowCount = sourceRowCount;
    let targetColCount = sourceColCount;

    if (selectionRowCount > 1 || selectionColCount > 1) {
      if (sourceRowCount === 1 && sourceColCount === 1) {
        targetRowCount = selectionRowCount;
        targetColCount = selectionColCount;
      } else {
        targetRowCount = Math.min(sourceRowCount, selectionRowCount);
        targetColCount = Math.min(sourceColCount, selectionColCount);
      }
    }

    const updates = [];
    for (let rowOffset = 0; rowOffset < targetRowCount; rowOffset += 1) {
      for (let colOffset = 0; colOffset < targetColCount; colOffset += 1) {
        const sourceRow = sourceRowCount === 1 && sourceColCount === 1 ? 0 : rowOffset;
        const sourceCol = sourceRowCount === 1 && sourceColCount === 1 ? 0 : colOffset;
        const targetRow = rect.rowStart + rowOffset;
        const targetCol = rect.colStart + colOffset;
        const staff = rect.scopedStaffs[targetRow];
        const day = daysInMonth[targetCol];
        if (!staff || !day) continue;
        const rawValue = grid[sourceRow]?.[sourceCol] ?? '';
        const { normalized, isValid } = normalizeManualShiftCode(rawValue, mergedLeaveCodes);
        if (!isValid) continue;
        updates.push({ staffId: staff.id, dateStr: day.date, value: normalized });
      }
    }

    if (updates.length === 0) return;

    setSchedule(prev => {
      const next = JSON.parse(JSON.stringify(prev));
      updates.forEach(({ staffId, dateStr, value }) => {
        if (!next[staffId]) next[staffId] = {};
        next[staffId][dateStr] = value ? { value, source: 'manual' } : null;
      });
      return next;
    });
    resetKeyInputBuffer();
  };

  useEffect(() => {
    const stopDrag = () => setIsRangeDragging(false);
    window.addEventListener('mouseup', stopDrag);
    return () => window.removeEventListener('mouseup', stopDrag);
  }, []);

  useEffect(() => {
    const handleGlobalKeyDown = (event) => {
      const target = event.target;
      const tagName = String(target?.tagName || '').toLowerCase();
      const isTypingTarget = tagName === 'input' || tagName === 'textarea' || tagName === 'select' || target?.isContentEditable;
      const hasSelection = selectedRangeCells.length > 0;

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c' && hasSelection) {
        event.preventDefault();
        copySelectionToClipboard();
        return;
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'v' && hasSelection) {
        event.preventDefault();
        pasteGridToSelection();
        return;
      }

      if (!hasSelection) return;
      if (isTypingTarget) return;

      if (event.key === 'Escape') {
        event.preventDefault();
        resetKeyInputBuffer();
        setRangeSelection(null);
        setSelectionAnchor(null);
        setSelectedGridCell(null);
        return;
      }

      if (event.key === 'Delete' || event.key === 'Backspace') {
        event.preventDefault();
        clearSelectionContents();
        return;
      }

      if (event.key.length === 1 && !event.altKey && !event.ctrlKey && !event.metaKey) {
        event.preventDefault();
        const nextBuffer = `${keyInputBuffer}${event.key}`;
        setKeyInputBuffer(nextBuffer);
        keepKeyInputBufferAlive();
        tryApplyBufferedCode(nextBuffer);
      }
    };

    window.addEventListener('keydown', handleGlobalKeyDown);
    return () => window.removeEventListener('keydown', handleGlobalKeyDown);
  }, [selectedRangeCells, keyInputBuffer, clipboardGrid, rangeSelection, selectedGridCell, staffs, daysInMonth, mergedLeaveCodes]);

  // ==========================================
  // 4. Excel 匯出 (ExcelJS 實現)
  // ==========================================
  const exportToExcel = async () => {
    setRuleFillFeedback("📊 正在產生高品質 Excel 報表...");
    const ExcelJS = await loadExcelJS();
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`${year}年${month}月班表`);

    const exportTheme = {
      pageBg: uiSettings?.pageBackgroundColor || '#f8fafc',
      tableFont: uiSettings?.tableFontColor || '#1f2937',
      shiftBg: uiSettings?.shiftColumnBgColor || '#ffffff',
      shiftFont: uiSettings?.shiftColumnFontColor || '#1e293b',
      nameBg: uiSettings?.nameDateColumnBgColor || '#ffffff',
      nameFont: uiSettings?.nameDateColumnFontColor || '#1e293b',
      weekdayHeadBg: blendHexColors(uiSettings?.nameDateColumnBgColor || '#ffffff', '#f1f5f9', 0.7),
      weekendHeadBg: colors.weekend || '#dcfce7',
      holidayHeadBg: colors.holiday || '#fca5a5',
      weekendCellBg: blendHexColors(colors.weekend || '#dcfce7', '#ffffff', 0.35),
      holidayCellBg: blendHexColors(colors.holiday || '#fca5a5', '#ffffff', 0.35),
      monthTitleBg: blendHexColors(uiSettings?.pageBackgroundColor || '#f8fafc', '#ffffff', 0.55),
      summaryBg: '#fef3c7',
      leaveRowBg: blendHexColors(uiSettings?.pageBackgroundColor || '#f8fafc', '#ffffff', 0.2)
    };

    const statHeaders = ['上班', '假日休', '總休', ...mergedLeaveCodes, ...(customColumns || [])];
    const totalColumns = 1 + daysInMonth.length + statHeaders.length;
    const lastDateColumn = daysInMonth.length + 1;

    const monthTitleRow = worksheet.addRow([]);
    monthTitleRow.height = 26;

    const titleStartCol = 2;
    const titleEndCol = Math.max(2, lastDateColumn - 2);
    const leaveStartCol = Math.max(titleEndCol + 1, lastDateColumn - 1);
    const leaveEndCol = lastDateColumn;

    if (titleEndCol >= titleStartCol) {
      worksheet.mergeCells(1, titleStartCol, 1, titleEndCol);
      const titleCell = monthTitleRow.getCell(titleStartCol);
      titleCell.value = `${month}月班表`;
      titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.monthTitleBg, '#EFF6FF') } };
      titleCell.font = { bold: true, size: 14, color: { argb: hexToExcelArgb(exportTheme.tableFont, '#1F2937') } };
    }

    if (leaveEndCol >= leaveStartCol) {
      if (leaveEndCol > leaveStartCol) worksheet.mergeCells(1, leaveStartCol, 1, leaveEndCol);
      const leaveCell = monthTitleRow.getCell(leaveStartCol);
      leaveCell.value = `應休${requiredLeaves}天`;
      leaveCell.alignment = { vertical: 'middle', horizontal: 'right' };
      leaveCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.monthTitleBg, '#EFF6FF') } };
      leaveCell.font = { bold: true, size: 11, color: { argb: hexToExcelArgb(exportTheme.tableFont, '#1F2937') } };
    }

    for (let col = 1; col <= totalColumns; col += 1) {
      const cell = monthTitleRow.getCell(col);
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
      if (col === 1 || col > lastDateColumn) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.pageBg, '#FFFFFF') } };
      }
    }

    const headerRow = ['姓名', ...daysInMonth.map(d => `${d.day}\n(${d.weekStr})`), ...statHeaders];
    const header = worksheet.addRow(headerRow);
    header.height = 30;

    header.eachCell((cell, colNumber) => {
      cell.font = { bold: true, size: 10, color: { argb: hexToExcelArgb(exportTheme.tableFont, '#1F2937') } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      const baseBorder = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
      cell.border = baseBorder;

      if (colNumber >= 2 && colNumber <= daysInMonth.length + 1) {
        const d = daysInMonth[colNumber - 2];
        cell.border = applyExcelFourWeekDivider(baseBorder, d.date);
        if (d.isHoliday) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.holidayHeadBg, '#FFCACA') } };
        else if (d.isWeekend) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.weekendHeadBg, '#DCFCE7') } };
        else cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.weekdayHeadBg, '#F1F5F9') } };
      } else if (colNumber === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.shiftBg, '#FFFFFF') } };
        cell.font = { ...(cell.font || {}), color: { argb: hexToExcelArgb(exportTheme.shiftFont, '#1E293B') } };
      } else {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.pageBg, '#F8FAFC') } };
      }
    });

    const makeBaseBorder = () => ({
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' }
    });

    const applyStandardCellStyle = (cell, colNumber, dateObj = null) => {
      const baseBorder = makeBaseBorder();
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.font = { color: { argb: hexToExcelArgb(colNumber === 1 ? exportTheme.nameFont : exportTheme.tableFont, '#1F2937') } };
      cell.border = baseBorder;
      if (dateObj) {
        cell.numFmt = '@';
        cell.border = applyExcelFourWeekDivider(baseBorder, dateObj.date);
        if (dateObj.isHoliday) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.holidayCellBg, '#FFE4E4') } };
        else if (dateObj.isWeekend) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.weekendCellBg, '#F0FDF4') } };
      }
    };

    const addStaffRow = (staff) => {
      const stats = getStaffStats(staff.id);
      const rowData = [
        staff.name,
        ...daysInMonth.map(d => {
          const cellData = schedule[staff.id]?.[d.date];
          return typeof cellData === 'object' ? (cellData?.value || '') : (cellData || '');
        }),
        stats.work,
        stats.holidayLeave,
        stats.totalLeave,
        ...mergedLeaveCodes.map(l => stats.leaveDetails[l] || ''),
        ...(customColumns || []).map(col => customColumnValues?.[staff.id]?.[col] || '')
      ];
      const row = worksheet.addRow(rowData);
      row.eachCell((cell, colNumber) => {
        const dateObj = (colNumber >= 2 && colNumber <= daysInMonth.length + 1) ? daysInMonth[colNumber - 2] : null;
        applyStandardCellStyle(cell, colNumber, dateObj);
      });
      return row;
    };

    const addSummaryRow = (summaryKey, includeRightStats = false) => {
      const rowData = [
        '',
        ...daysInMonth.map(d => getDailyStats(d.date)[summaryKey] || ''),
        ...(includeRightStats ? Array(statHeaders.length).fill('') : Array(statHeaders.length).fill(''))
      ];
      const row = worksheet.addRow(rowData);
      row.eachCell((cell, colNumber) => {
        const baseBorder = makeBaseBorder();
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { bold: true, color: { argb: hexToExcelArgb(exportTheme.tableFont, '#1F2937') } };
        cell.border = baseBorder;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.summaryBg, '#FEF3C7') } };
        if (colNumber >= 2 && colNumber <= daysInMonth.length + 1) {
          const d = daysInMonth[colNumber - 2];
          cell.border = applyExcelFourWeekDivider(baseBorder, d.date);
        }
      });
      return row;
    };

    groupedStaffs.forEach(({ group, staffs: groupStaffList }) => {
      groupStaffList.forEach(addStaffRow);
      const summaryKey = group === '白班' ? 'D' : group === '小夜' ? 'E' : 'N';
      addSummaryRow(summaryKey);
    });

    const leaveRowData = [
      '',
      ...daysInMonth.map(d => getDailyStats(d.date).totalLeave || ''),
      ...Array(statHeaders.length).fill('')
    ];
    const leaveRow = worksheet.addRow(leaveRowData);
    leaveRow.eachCell((cell, colNumber) => {
      const baseBorder = makeBaseBorder();
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.font = { bold: true, color: { argb: hexToExcelArgb(exportTheme.tableFont, '#1F2937') } };
      cell.border = baseBorder;
      if (colNumber >= 1 && colNumber <= totalColumns) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.leaveRowBg, '#FFFFFF') } };
      }
      if (colNumber >= 2 && colNumber <= daysInMonth.length + 1) {
        const d = daysInMonth[colNumber - 2];
        cell.border = applyExcelFourWeekDivider(baseBorder, d.date);
      }
    });

    worksheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }];

    worksheet.getColumn(1).width = 15;
    for (let i = 2; i <= daysInMonth.length + 1; i += 1) worksheet.getColumn(i).width = 5;
    for (let i = daysInMonth.length + 2; i <= totalColumns; i += 1) worksheet.getColumn(i).width = 8;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `排班表_${year}年${month}月.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    setShowExportMenu(false);
    setRuleFillFeedback("✅ Excel 導出成功！");
  };

  const exportToWord = () => {
    const statHeaders = ['上班', '假日休', '總休'];
    const exportTheme = {
      pageBg: uiSettings?.pageBackgroundColor || '#f8fafc',
      tableFont: uiSettings?.tableFontColor || '#1f2937',
      shiftBg: uiSettings?.shiftColumnBgColor || '#ffffff',
      shiftFont: uiSettings?.shiftColumnFontColor || '#1e293b',
      nameBg: uiSettings?.nameDateColumnBgColor || '#ffffff',
      nameFont: uiSettings?.nameDateColumnFontColor || '#1e293b',
      weekdayHeadBg: blendHexColors(uiSettings?.nameDateColumnBgColor || '#ffffff', '#f1f5f9', 0.7),
      weekendHeadBg: colors.weekend || '#dcfce7',
      holidayHeadBg: colors.holiday || '#fca5a5',
      weekendCellBg: blendHexColors(colors.weekend || '#dcfce7', '#ffffff', 0.35),
      holidayCellBg: blendHexColors(colors.holiday || '#fca5a5', '#ffffff', 0.35),
      statWorkBg: blendHexColors(uiSettings?.pageBackgroundColor || '#f8fafc', '#93c5fd', 0.45),
      statHolidayBg: blendHexColors(uiSettings?.pageBackgroundColor || '#f8fafc', colors.weekend || '#dcfce7', 0.5),
      statTotalBg: blendHexColors(uiSettings?.pageBackgroundColor || '#f8fafc', colors.holiday || '#fca5a5', 0.45)
    };

    const titleColSpan = daysInMonth.length;
    const leaveColSpan = statHeaders.length;
    const totalColumns = 1 + daysInMonth.length + statHeaders.length;
    const schedulingRuleLines = String(schedulingRulesText || '')
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean);
    const schedulingRulesHtml = schedulingRuleLines.length > 0
      ? `排班規則：<br/>${schedulingRuleLines.map((line, index) => `${index + 1}. ${line}`).join('<br/>')}`
      : '排班規則：';

    const webSummaryRowBg = '#fef3c7';
    const summaryRows = [
      { group: '白班', key: 'D', label: '白班上班', bg: webSummaryRowBg },
      { group: '小夜', key: 'E', label: '小夜上班', bg: webSummaryRowBg },
      { group: '大夜', key: 'N', label: '大夜上班', bg: webSummaryRowBg }
    ];

    const groupedExportRowsHtml = SHIFT_GROUPS.map((group) => {
      const groupStaffsForExport = staffs.filter((staff) => (staff.group || '白班') === group);

      const staffRowsHtml = groupStaffsForExport.map((staff) => {
        const stats = getStaffStats(staff.id);
        return `
                <tr>
                  <td class="name-col" style="background:${exportTheme.nameBg}; color:${exportTheme.nameFont}; mso-pattern:auto none;">${staff.name}</td>
                  ${daysInMonth.map(d => {
                    const cellData = schedule[staff.id]?.[d.date];
                    const value = typeof cellData === 'object' ? (cellData?.value || '') : (cellData || '');
                    const cellClass = d.isHoliday ? 'holiday-cell' : (d.isWeekend ? 'weekend-cell' : '');
                    const cellBg = d.isHoliday ? exportTheme.holidayCellBg : (d.isWeekend ? exportTheme.weekendCellBg : exportTheme.pageBg);
                    return `<td class="day-col ${cellClass}" style="background:${cellBg}; mso-pattern:auto none;${getWordCycleDividerStyle(d.date)}">${value}</td>`;
                  }).join('')}
                  <td class="stat-col stat-work-cell" style="background:${exportTheme.statWorkBg}; color:${exportTheme.tableFont}; mso-pattern:auto none;">${stats.work || ''}</td>
                  <td class="stat-col stat-holiday-cell" style="background:${exportTheme.statHolidayBg}; color:${exportTheme.tableFont}; mso-pattern:auto none;">${stats.holidayLeave || ''}</td>
                  <td class="stat-col stat-total-cell" style="background:${exportTheme.statTotalBg}; color:${exportTheme.tableFont}; mso-pattern:auto none;">${stats.totalLeave || ''}</td>
                </tr>`;
      }).join('');

      const summaryConfig = summaryRows.find((item) => item.group === group);
      const summaryRowHtml = summaryConfig ? `
                <tr>
                  <td class="name-col summary-label-cell" style="background:${summaryConfig.bg}; color:${exportTheme.nameFont}; mso-pattern:auto none;"></td>
                  ${daysInMonth.map(d => {
                    const count = getDailyStats(d.date)[summaryConfig.key];
                    return `<td class="day-col summary-value-cell" style="background:${summaryConfig.bg}; color:${exportTheme.tableFont}; mso-pattern:auto none;${getWordCycleDividerStyle(d.date)}">${count || ''}</td>`;
                  }).join('')}
                  <td class="stat-col summary-value-cell" style="background:${summaryConfig.bg}; mso-pattern:auto none;"></td>
                  <td class="stat-col summary-value-cell" style="background:${summaryConfig.bg}; mso-pattern:auto none;"></td>
                  <td class="stat-col summary-value-cell" style="background:${summaryConfig.bg}; mso-pattern:auto none;"></td>
                </tr>` : '';

      return `${staffRowsHtml}${summaryRowHtml}`;
    }).join('');

    const html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office"
          xmlns:w="urn:schemas-microsoft-com:office:word"
          xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8">
        <meta name="ProgId" content="Word.Document">
        <meta name="Generator" content="Microsoft Word 15">
        <meta name="Originator" content="Microsoft Word 15">
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>90</w:Zoom>
            <w:DoNotOptimizeForBrowser/>
          </w:WordDocument>
        </xml>
        <style>
          @page WordSection1 {
            size: 841.9pt 595.3pt;
            mso-page-orientation: landscape;
            margin: 28.35pt 28.35pt 28.35pt 28.35pt;
          }
          div.WordSection1 { page: WordSection1; }
          body {
            font-family: "Microsoft JhengHei", Arial, sans-serif;
            margin: 0;
            padding: 0;
            color: ${exportTheme.tableFont};
            background: ${exportTheme.pageBg};
          }
          table {
            border-collapse: collapse;
            table-layout: fixed;
            width: auto;
            margin: 0 auto;
            font-size: 9pt;
          }
          th, td {
            border: 1px solid #000;
            padding: 2px 3px;
            text-align: center;
            vertical-align: middle;
            word-break: break-all;
          }
          .month-row td {
            height: 24pt;
            font-weight: 700;
            background: ${exportTheme.pageBg};
          }
          .month-title-cell {
            height: 24pt;
            padding: 0 8pt;
            text-align: center;
          }
          .month-title {
            display: block;
            width: 100%;
            text-align: center;
            font-size: 14pt;
            font-weight: 700;
            line-height: 24pt;
          }
          .leave-title-cell {
            height: 24pt;
            padding: 0 8pt;
            text-align: center;
            font-size: 10.5pt;
            font-weight: 700;
            white-space: nowrap;
          }
          .name-col {
            width: 54pt;
            min-width: 54pt;
            font-weight: 700;
          }
          .day-col {
            width: 18pt;
            min-width: 18pt;
          }
          .stat-col {
            width: 28pt;
            min-width: 28pt;
          }
          .header-cell {
            font-weight: 700;
            line-height: 1.1;
          }
          .weekday-head { background-color: ${exportTheme.weekdayHeadBg}; }
          .holiday-head { background-color: ${exportTheme.holidayHeadBg}; }
          .weekend-head { background-color: ${exportTheme.weekendHeadBg}; }
          .holiday-cell { background-color: ${exportTheme.holidayCellBg}; }
          .weekend-cell { background-color: ${exportTheme.weekendCellBg}; }
          .stat-work-head { background-color: ${exportTheme.statWorkBg}; color: ${exportTheme.tableFont}; }
          .stat-holiday-head { background-color: ${exportTheme.statHolidayBg}; color: ${exportTheme.tableFont}; }
          .stat-total-head { background-color: ${exportTheme.statTotalBg}; color: ${exportTheme.tableFont}; }
          .stat-work-cell { background-color: ${exportTheme.statWorkBg}; }
          .stat-holiday-cell { background-color: ${exportTheme.statHolidayBg}; }
          .stat-total-cell { background-color: ${exportTheme.statTotalBg}; }
          .summary-label-cell, .summary-value-cell {
            font-weight: 700;
          }
          .rules-row td {
            padding: 8pt 10pt;
            text-align: left;
            vertical-align: top;
            line-height: 1.7;
            font-size: 10pt;
            background: ${exportTheme.pageBg};
          }
        </style>
      </head>
      <body>
        <div class="WordSection1">
          <table>
            <thead>
              <tr class="month-row">
                <td class="name-col"></td>
                <td class="month-title-cell" colspan="${titleColSpan}">
                  <span class="month-title">${month}月班表</span>
                </td>
                <td class="leave-title-cell" colspan="${leaveColSpan}">應休${requiredLeaves}天</td>
              </tr>
              <tr>
                <th class="name-col header-cell" style="background:${exportTheme.nameBg}; color:${exportTheme.nameFont}; mso-pattern:auto none;">姓名</th>
                ${daysInMonth.map(d => {
                  const headClass = d.isHoliday ? 'holiday-head' : (d.isWeekend ? 'weekend-head' : 'weekday-head');
                  const headBg = d.isHoliday ? exportTheme.holidayHeadBg : (d.isWeekend ? exportTheme.weekendHeadBg : exportTheme.weekdayHeadBg);
                  return `<th class="day-col header-cell ${headClass}" style="background:${headBg}; mso-pattern:auto none;${getWordCycleDividerStyle(d.date)}">${d.day}<br/>(${d.weekStr})</th>`;
                }).join('')}
                <th class="stat-col header-cell stat-work-head">上班</th>
                <th class="stat-col header-cell stat-holiday-head">假日休</th>
                <th class="stat-col header-cell stat-total-head">總休</th>
              </tr>
            </thead>
            <tbody>
              ${groupedExportRowsHtml}
            </tbody>
            <tfoot>
              <tr class="rules-row">
                <td colspan="${totalColumns}">${schedulingRulesHtml}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </body>
    </html>`;

    const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `列印班表_${year}年${month}月.doc`;
    a.click();
    URL.revokeObjectURL(url);
    setShowExportMenu(false);
    setRuleFillFeedback("✅ Word 導出成功！");
  };

  // ==========================================
  // 5. 規則式半智慧補班功能
  // ==========================================
  const handleRuleBasedAutoSchedule = async (isPartial = false) => {
    setIsRuleFillLoading(true);
    setRuleFillFeedback(isPartial ? "🧩 系統正在依指定範圍進行規則補空..." : "🧩 系統正在依人力需求執行整月規則補空...");

    try {
      const mergedSchedule = JSON.parse(JSON.stringify(schedule));
      const targetStaffIds = isPartial && ruleFillConfig.selectedStaffs.length > 0
        ? new Set(ruleFillConfig.selectedStaffs)
        : new Set(staffs.map(s => s.id));

      const targetDays = daysInMonth.filter(d => {
        if (!isPartial) return true;
        return d.day >= ruleFillConfig.dateRange.start && d.day <= ruleFillConfig.dateRange.end;
      });

      const normalizedTargetShift = RULE_FILL_MAIN_SHIFTS.includes(ruleFillConfig.targetShift) ? ruleFillConfig.targetShift : '';
      const restrictedGroup = normalizedTargetShift ? getShiftGroupByCode(normalizedTargetShift) : null;
      const summary = { workFilled: 0, leaveFilled: 0, skipped: 0 };

      const getScheduleCode = (snapshot, staffId, dateStr) => {
        const cellData = snapshot[staffId]?.[dateStr];
        return typeof cellData === 'object' && cellData !== null ? (cellData.value || '') : (cellData || '');
      };

      const setScheduleCode = (snapshot, staffId, dateStr, value, source = 'auto') => {
        if (!snapshot[staffId]) snapshot[staffId] = {};
        snapshot[staffId][dateStr] = value ? { value, source } : null;
      };

      const getDemandType = (day) => (day.isWeekend || day.isHoliday) ? 'holiday' : 'weekday';
      const getDemandForGroup = (day, group) => {
        const bucket = getDemandType(day);
        const key = GROUP_TO_DEMAND_KEY[group];
        return Number(staffingConfig?.requiredStaffing?.[bucket]?.[key] || 0);
      };

      const getAssignedCountByGroup = (snapshot, dateStr, group) => {
        return staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => {
          const code = getScheduleCode(snapshot, s.id, dateStr);
          return sum + (getShiftGroupByCode(code) === group ? 1 : 0);
        }, 0);
      };

      const countConsecutiveBeforeFromSnapshot = (snapshot, staffId, dateStr) => {
        let count = 0;
        let cursor = addDays(parseDateKey(dateStr), -1);
        while (true) {
          const key = formatDateKey(cursor);
          const code = getScheduleCode(snapshot, staffId, key);
          if (!isShiftCode(code)) break;
          count += 1;
          cursor = addDays(cursor, -1);
        }
        return count;
      };

      const canAssignWithSnapshot = (snapshot, staff, dateStr, shiftCode) => {
        const reasons = [];
        const currentCode = getScheduleCode(snapshot, staff.id, dateStr);
        if (currentCode) reasons.push('該格已有排班或休假代碼');
        const prefix = getCodePrefix(currentCode);
        if (isConfiguredLeaveCode(currentCode)) reasons.push('該格已有休假，不可再排班');
        const staffGroup = staff.group || '白班';
        const shiftGroup = getShiftGroupByCode(shiftCode);
        if (!SMART_RULES.allowCrossGroupAssignment && shiftGroup && staffGroup !== shiftGroup) reasons.push('不可跨群組排班');
        const prevKey = formatDateKey(addDays(parseDateKey(dateStr), -1));
        const prevCode = getScheduleCode(snapshot, staff.id, prevKey);
        const disallowed = SMART_RULES.disallowedNextShiftMap[prevCode] || [];
        if (disallowed.includes(shiftCode)) reasons.push(`${prevCode} 後不可接 ${shiftCode}`);
        const consecutiveBefore = countConsecutiveBeforeFromSnapshot(snapshot, staff.id, dateStr);
        if (consecutiveBefore + 1 > SMART_RULES.maxConsecutiveWorkDays) reasons.push(`連續上班不可超過 ${SMART_RULES.maxConsecutiveWorkDays} 天`);
        if (staff.pregnant && SMART_RULES.pregnancyRestrictedShifts.includes(shiftCode)) reasons.push('懷孕標記人員不可排 N / 夜8-8');
        return { allowed: reasons.length === 0, reasons };
      };

      const getWorkCountFromSnapshot = (snapshot, staffId) => {
        return daysInMonth.reduce((sum, d) => sum + (isShiftCode(getScheduleCode(snapshot, staffId, d.date)) ? 1 : 0), 0);
      };

      const getShiftCountFromSnapshot = (snapshot, staffId, shiftCode) => {
        return daysInMonth.reduce((sum, d) => sum + (getScheduleCode(snapshot, staffId, d.date) === shiftCode ? 1 : 0), 0);
      };

      const getLeaveCountFromSnapshot = (snapshot, staffId) => {
        return daysInMonth.reduce((sum, d) => sum + (isConfiguredLeaveCode(getScheduleCode(snapshot, staffId, d.date)) ? 1 : 0), 0);
      };

      const getBlankCountFromSnapshot = (snapshot, staffId) => {
        return daysInMonth.reduce((sum, d) => sum + (!getScheduleCode(snapshot, staffId, d.date) ? 1 : 0), 0);
      };

      const canStillMeetRequiredLeavesAfterAssign = (snapshot, staffId) => {
        const currentLeaves = getLeaveCountFromSnapshot(snapshot, staffId);
        const remainingBlanks = getBlankCountFromSnapshot(snapshot, staffId);
        const remainingLeavesNeeded = Math.max(0, requiredLeaves - currentLeaves);
        return remainingBlanks >= remainingLeavesNeeded;
      };

      const canStillMeetRequiredLeavesIfAssignShift = (snapshot, staffId) => {
        const currentLeaves = getLeaveCountFromSnapshot(snapshot, staffId);
        const remainingBlanks = getBlankCountFromSnapshot(snapshot, staffId);
        const remainingLeavesNeeded = Math.max(0, requiredLeaves - currentLeaves);
        return (remainingBlanks - 1) >= remainingLeavesNeeded;
      };

      const getRecentWorkPressure = (snapshot, staffId, dateStr, lookback = 3) => {
        let count = 0;
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 0; i < lookback; i += 1) {
          const code = getScheduleCode(snapshot, staffId, formatDateKey(cursor));
          if (isShiftCode(code)) count += 1;
          cursor = addDays(cursor, -1);
        }
        return count;
      };

      const getRecentLeavePressure = (snapshot, staffId, dateStr, lookback = 4) => {
        let count = 0;
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 0; i < lookback; i += 1) {
          const code = getScheduleCode(snapshot, staffId, formatDateKey(cursor));
          if (isLeaveCode(code)) count += 1;
          cursor = addDays(cursor, -1);
        }
        return count;
      };

      const getNearbyLeavePressure = (snapshot, staffId, dateStr, radius = 2) => {
        let count = 0;
        const center = parseDateKey(dateStr);
        for (let offset = -radius; offset <= radius; offset += 1) {
          if (offset === 0) continue;
          const key = formatDateKey(addDays(center, offset));
          const code = getScheduleCode(snapshot, staffId, key);
          if (isLeaveCode(code)) count += 1;
        }
        return count;
      };

      const getDaysSinceLastLeave = (snapshot, staffId, dateStr, maxLookback = 10) => {
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 1; i <= maxLookback; i += 1) {
          const code = getScheduleCode(snapshot, staffId, formatDateKey(cursor));
          if (isLeaveCode(code)) return i;
          cursor = addDays(cursor, -1);
        }
        return maxLookback + 1;
      };

      const getGroupLeaveLoad = (snapshot, dateStr, group) => {
        return staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => {
          const code = getScheduleCode(snapshot, s.id, dateStr);
          return sum + (isLeaveCode(code) ? 1 : 0);
        }, 0);
      };

      const getConsecutiveLeavePattern = (snapshot, staffId, dateStr) => {
        const prevCode = getScheduleCode(snapshot, staffId, formatDateKey(addDays(parseDateKey(dateStr), -1)));
        const nextCode = getScheduleCode(snapshot, staffId, formatDateKey(addDays(parseDateKey(dateStr), 1)));
        const prevIsLeave = isLeaveCode(prevCode);
        const nextIsLeave = isLeaveCode(nextCode);
        return {
          prevIsLeave,
          nextIsLeave,
          adjacentLeaveCount: (prevIsLeave ? 1 : 0) + (nextIsLeave ? 1 : 0)
        };
      };

      const scoreCandidateWithSnapshot = (snapshot, staff, dateStr, shiftCode) => {
        let score = 0;
        score += (999 - getShiftCountFromSnapshot(snapshot, staff.id, shiftCode)) * SMART_RULES.fillPriorityWeights.sameShiftCount;
        score += (999 - getWorkCountFromSnapshot(snapshot, staff.id)) * SMART_RULES.fillPriorityWeights.totalShiftCount;
        if (getShiftGroupByCode(shiftCode) === (staff.group || '白班')) score += 100 * SMART_RULES.fillPriorityWeights.sameGroup;
        score -= getRecentWorkPressure(snapshot, staff.id, dateStr, 3) * 18;
        return score;
      };

      const scoreLeaveCandidateWithSnapshot = (snapshot, staff, dateStr) => {
        const leaveDeficit = Math.max(0, requiredLeaves - getLeaveCountFromSnapshot(snapshot, staff.id));
        const workCount = getWorkCountFromSnapshot(snapshot, staff.id);
        const group = staff.group || '白班';
        const sameDayLeaveLoad = getGroupLeaveLoad(snapshot, dateStr, group);
        const recentLeavePressure = getRecentLeavePressure(snapshot, staff.id, dateStr, 4);
        const nearbyLeavePressure = getNearbyLeavePressure(snapshot, staff.id, dateStr, 2);
        const daysSinceLastLeave = getDaysSinceLastLeave(snapshot, staff.id, dateStr, 10);
        const consecutiveLeavePattern = getConsecutiveLeavePattern(snapshot, staff.id, dateStr);
        let score = 0;
        score += leaveDeficit * 120;
        score += workCount * 5;
        score += getRecentWorkPressure(snapshot, staff.id, dateStr, 3) * 18;
        score += Math.min(daysSinceLastLeave, 10) * 8;
        score -= sameDayLeaveLoad * 30;
        score -= recentLeavePressure * 18;
        score -= nearbyLeavePressure * 28;
        if (consecutiveLeavePattern.adjacentLeaveCount === 1) score += 22;
        if (consecutiveLeavePattern.adjacentLeaveCount >= 2) score -= 12;
        return score;
      };

      for (const day of targetDays) {
        for (const group of SHIFT_GROUPS) {
          if (restrictedGroup && restrictedGroup !== group) continue;

          const shiftCode = normalizedTargetShift && getShiftGroupByCode(normalizedTargetShift) === group
            ? normalizedTargetShift
            : DEFAULT_SHIFT_BY_GROUP[group];

          const demand = getDemandForGroup(day, group);
          const alreadyAssigned = getAssignedCountByGroup(mergedSchedule, day.date, group);
          const needed = Math.max(0, demand - alreadyAssigned);

          const groupStaffs = staffs.filter(s => (s.group || '白班') === group && targetStaffIds.has(s.id));
          const groupStaffIds = new Set(groupStaffs.map(s => s.id));

          // 逐格補主班別，補到需求就停
          for (let slot = 0; slot < needed; slot += 1) {
            const assignableCandidates = groupStaffs
              .filter(staff => !getScheduleCode(mergedSchedule, staff.id, day.date))
              .map(staff => {
                const result = canAssignWithSnapshot(mergedSchedule, staff, day.date, shiftCode);
                const canKeepLeaveTarget = result.allowed ? canStillMeetRequiredLeavesIfAssignShift(mergedSchedule, staff.id) : false;
                return {
                  staff,
                  allowed: result.allowed && canKeepLeaveTarget,
                  score: result.allowed && canKeepLeaveTarget ? scoreCandidateWithSnapshot(mergedSchedule, staff, day.date, shiftCode) : -1
                };
              })
              .filter(item => item.allowed)
              .sort((a, b) => b.score - a.score);

            if (assignableCandidates.length === 0) {
              summary.skipped += 1;
              continue;
            }

            const picked = assignableCandidates[0];
            setScheduleCode(mergedSchedule, picked.staff.id, day.date, shiftCode, 'auto');
            summary.workFilled += 1;
          }

          // 需求已滿後，只替休假不足者補 off；其他空白保留
          const leaveCandidates = groupStaffs
            .filter(staff => !getScheduleCode(mergedSchedule, staff.id, day.date))
            .filter(staff => getLeaveCountFromSnapshot(mergedSchedule, staff.id) < requiredLeaves)
            .map(staff => ({ staff, score: scoreLeaveCandidateWithSnapshot(mergedSchedule, staff, day.date) }))
            .sort((a, b) => b.score - a.score);

          if (leaveCandidates.length > 0) {
            const currentLeaveLoad = getGroupLeaveLoad(mergedSchedule, day.date, group);
            const maxLeaveForDay = Math.max(0, groupStaffIds.size - demand);
            if (currentLeaveLoad < maxLeaveForDay) {
              const bestLeaveCandidate = leaveCandidates[0];
              if (bestLeaveCandidate && canStillMeetRequiredLeavesAfterAssign(mergedSchedule, bestLeaveCandidate.staff.id)) {
                setScheduleCode(mergedSchedule, bestLeaveCandidate.staff.id, day.date, defaultAutoLeaveCode, 'auto');
                summary.leaveFilled += 1;
              }
            }
          }
        }
      }

      setSchedule(mergedSchedule);
      saveToHistory(isPartial ? '規則指定補空' : '規則全月補空', mergedSchedule);
      setRuleFillFeedback(`✅ 補空完成：上班 ${summary.workFilled} 格、休假 ${summary.leaveFilled} 格、未補成功 ${summary.skipped} 格`);
    } catch (error) {
      console.error(error);
      setRuleFillFeedback("❌ 規則補空失敗，請檢查設定。");
    } finally {
      setIsRuleFillLoading(false);
    }
  };


  // ==========================================
  // 6. 輔助統計與操作
  // ==========================================
  const getStaffStats = (staffId) => {
    const stats = {
      work: 0,
      holidayLeave: 0,
      totalLeave: 0,
      leaveDetails: Object.fromEntries(mergedLeaveCodes.map(l => [l, 0]))
    };

    const mySchedule = schedule[staffId] || {};
    daysInMonth.forEach(d => {
      const cellData = mySchedule[d.date];
      const code = typeof cellData === 'object' && cellData !== null ? cellData.value : cellData;
      if (!code) return;

      if (DICT.SHIFTS.includes(code)) stats.work += 1;
      if (isConfiguredLeaveCode(code)) {
        stats.totalLeave += 1;
        if (stats.leaveDetails[code] !== undefined) stats.leaveDetails[code] += 1;
        if (d.isWeekend || d.isHoliday) stats.holidayLeave += 1;
      }
    });
    return stats;
  };

  const getDailyStats = (dateStr) => {
    const stats = { D: 0, E: 0, N: 0, totalLeave: 0 };

    staffs.forEach(staff => {
      const cellData = schedule[staff.id]?.[dateStr];
      const code = typeof cellData === 'object' && cellData !== null ? cellData.value : cellData;
      if (!code) return;

      if (['D', '白8-8', '8-12', '12-16'].includes(code)) {
        stats.D += 1;
      } else if (['E', '夜8-8'].includes(code)) {
        stats.E += 1;
      } else if (code === 'N') {
        stats.N += 1;
      } else if (isConfiguredLeaveCode(code)) {
        stats.totalLeave += 1;
      }
    });

    return stats;
  };

  const getRequiredCountForDate = (dateStr, rowKey) => {
    const dayInfo = daysInMonth.find(d => d.date === dateStr);
    if (!dayInfo) return null;
    const bucket = (dayInfo.isWeekend || dayInfo.isHoliday) ? 'holiday' : 'weekday';
    if (rowKey === 'D') return Number(staffingConfig?.requiredStaffing?.[bucket]?.white || 0);
    if (rowKey === 'E') return Number(staffingConfig?.requiredStaffing?.[bucket]?.evening || 0);
    if (rowKey === 'N') return Number(staffingConfig?.requiredStaffing?.[bucket]?.night || 0);
    return null;
  };

  const getDemandHighlightStyle = (dateStr, rowKey, actualCount) => {
    if (!['D', 'E', 'N'].includes(rowKey)) return {};
    const requiredCount = getRequiredCountForDate(dateStr, rowKey);
    if (requiredCount === null) return {};
    if (actualCount > requiredCount) return { backgroundColor: demandOverColor };
    return {};
  };

  const saveToHistory = (label, currentSchedule = schedule) => {
    const newRecord = {
      id: Date.now(),
      label,
      timestamp: new Date().toLocaleString(),
      state: { year, month, staffs, schedule: currentSchedule, colors, customHolidays, specialWorkdays, medicalCalendarAdjustments, staffingConfig, uiSettings, customLeaveCodes, customColumns, customColumnValues, schedulingRulesText }
    };

    setHistoryList(prev => {
      const updated = [newRecord, ...prev].slice(0, 10);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
      return updated;
    });
  };

  const loadHistory = (record) => {
    const { state } = record;
    setYear(state.year);
    setMonth(state.month);
    setCustomHolidays(Array.isArray(state.customHolidays) ? state.customHolidays : []);
    setSpecialWorkdays(Array.isArray(state.specialWorkdays) ? state.specialWorkdays : []);
    setMedicalCalendarAdjustments(state.medicalCalendarAdjustments || { holidays: [], workdays: [] });
    if (state.staffingConfig) setStaffingConfig(state.staffingConfig);
    if (state.uiSettings) setUiSettings(state.uiSettings);
    if (Array.isArray(state.customLeaveCodes)) setCustomLeaveCodes(state.customLeaveCodes);
    if (Array.isArray(state.customColumns)) setCustomColumns(state.customColumns);
    if (state.customColumnValues) setCustomColumnValues(state.customColumnValues);
    if (typeof state.schedulingRulesText === 'string') setSchedulingRulesText(state.schedulingRulesText);
    setStaffs(normalizeStaffGroup(state.staffs));
    setSchedule(state.schedule);
    if (state.colors) setColors(state.colors);
    setShowHistoryModal(false);
    setShowDraftPrompt(false);
  };

  const clearHistory = () => {
    if (window.confirm("確定要清空所有歷史紀錄嗎？")) {
      localStorage.removeItem(STORAGE_KEY);
      setHistoryList([]);
    }
  };

  const handleCellChange = (staffId, dateStr, value) => {
    setSchedule(prev => ({
      ...prev,
      [staffId]: { ...prev[staffId], [dateStr]: value ? { value, source: 'manual' } : null }
    }));
  };

  const getCellCode = (staffId, dateStr) => {
    const cellData = schedule[staffId]?.[dateStr];
    return typeof cellData === 'object' && cellData !== null ? (cellData.value || '') : (cellData || '');
  };

  const getCellSource = (staffId, dateStr) => {
    const cellData = schedule[staffId]?.[dateStr];
    if (!cellData) return '';
    if (typeof cellData === 'object' && cellData !== null) return cellData.source || 'manual';
    return 'manual';
  };

  const countConsecutiveWorkDaysBefore = (staffId, dateStr) => {
    let count = 0;
    let cursor = addDays(parseDateKey(dateStr), -1);
    while (true) {
      const key = formatDateKey(cursor);
      const code = getCellCode(staffId, key);
      if (!isShiftCode(code)) break;
      count += 1;
      cursor = addDays(cursor, -1);
    }
    return count;
  };

  const canAssign = (staff, dateStr, shiftCode) => {
    const reasons = [];
    const currentCode = getCellCode(staff.id, dateStr);
    if (currentCode) {
      reasons.push('該格已有排班或休假代碼');
    }

    const prefix = getCodePrefix(currentCode);
    if (isConfiguredLeaveCode(currentCode)) {
      reasons.push('該格已有休假，不可再排班');
    }

    const staffGroup = staff.group || '白班';
    const shiftGroup = getShiftGroupByCode(shiftCode);
    if (!SMART_RULES.allowCrossGroupAssignment && shiftGroup && staffGroup !== shiftGroup) {
      reasons.push('不可跨群組排班');
    }

    const prevKey = formatDateKey(addDays(parseDateKey(dateStr), -1));
    const prevCode = getCellCode(staff.id, prevKey);
    const disallowed = SMART_RULES.disallowedNextShiftMap[prevCode] || [];
    if (disallowed.includes(shiftCode)) {
      reasons.push(`${prevCode} 後不可接 ${shiftCode}`);
    }

    const consecutiveBefore = countConsecutiveWorkDaysBefore(staff.id, dateStr);
    if (consecutiveBefore + 1 > SMART_RULES.maxConsecutiveWorkDays) {
      reasons.push(`連續上班不可超過 ${SMART_RULES.maxConsecutiveWorkDays} 天`);
    }

    if (staff.pregnant && SMART_RULES.pregnancyRestrictedShifts.includes(shiftCode)) {
      reasons.push('懷孕標記人員不可排 N / 夜8-8');
    }

    return { allowed: reasons.length === 0, reasons };
  };

  const getShiftCountForStaff = (staffId, shiftCode) => {
    return daysInMonth.reduce((sum, d) => sum + (getCellCode(staffId, d.date) === shiftCode ? 1 : 0), 0);
  };

  const scoreCandidate = (staff, dateStr, shiftCode) => {
    let score = 0;
    const stats = getStaffStats(staff.id);
    score += (999 - getShiftCountForStaff(staff.id, shiftCode)) * SMART_RULES.fillPriorityWeights.sameShiftCount;
    score += (999 - stats.work) * SMART_RULES.fillPriorityWeights.totalShiftCount;
    if (getShiftGroupByCode(shiftCode) === (staff.group || '白班')) {
      score += 100 * SMART_RULES.fillPriorityWeights.sameGroup;
    }
    return score;
  };

  const getCurrentConsecutiveLeavePattern = (staffId, dateStr) => {
    const prevCode = getCellCode(staffId, formatDateKey(addDays(parseDateKey(dateStr), -1)));
    const nextCode = getCellCode(staffId, formatDateKey(addDays(parseDateKey(dateStr), 1)));
    const prevIsLeave = isLeaveCode(prevCode);
    const nextIsLeave = isLeaveCode(nextCode);
    return {
      prevIsLeave,
      nextIsLeave,
      adjacentLeaveCount: (prevIsLeave ? 1 : 0) + (nextIsLeave ? 1 : 0)
    };
  };

  const openFillModal = (staff, dateStr) => {
    const group = staff.group || '白班';
    const shiftCode = DEFAULT_SHIFT_BY_GROUP[group];
    const dayInfo = daysInMonth.find(d => d.date === dateStr);
    const demand = dayInfo ? Number(staffingConfig?.requiredStaffing?.[(dayInfo.isWeekend || dayInfo.isHoliday) ? 'holiday' : 'weekday']?.[GROUP_TO_DEMAND_KEY[group]] || 0) : 0;
    const alreadyAssigned = staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => {
      const code = getCellCode(s.id, dateStr);
      return sum + (getShiftGroupByCode(code) === group ? 1 : 0);
    }, 0);
    const leaveCount = daysInMonth.reduce((sum, d) => sum + (isLeaveCode(getCellCode(staff.id, d.date)) ? 1 : 0), 0);
    const leaveDeficit = Math.max(0, requiredLeaves - leaveCount);

    const shiftResult = canAssign(staff, dateStr, shiftCode);
    const currentLeaves = leaveCount;
    const remainingBlanks = daysInMonth.reduce((sum, d) => sum + (!getCellCode(staff.id, d.date) ? 1 : 0), 0);
    const remainingLeavesNeeded = Math.max(0, requiredLeaves - currentLeaves);
    const canKeepLeaveTarget = (remainingBlanks - 1) >= remainingLeavesNeeded;

    const candidates = [];

    if (alreadyAssigned < demand && shiftResult.allowed && canKeepLeaveTarget) {
      const reasonBits = [
        `${group}缺額尚未補滿`,
        `${shiftCode} 為此群組主班別`
      ];
      candidates.push({
        type: 'self-shift',
        staffId: staff.id,
        staffName: staff.name,
        group,
        shiftCode,
        allowed: true,
        score: scoreCandidate(staff, dateStr, shiftCode),
        reasons: reasonBits
      });
    }

    if (alreadyAssigned >= demand && leaveDeficit > 0) {
      const sameDayLeaveLoad = staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => sum + (isLeaveCode(getCellCode(s.id, dateStr)) ? 1 : 0), 0);
      const recentWorkPressure = (() => {
        let count = 0;
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 0; i < 3; i += 1) {
          if (isShiftCode(getCellCode(staff.id, formatDateKey(cursor)))) count += 1;
          cursor = addDays(cursor, -1);
        }
        return count;
      })();
      const consecutiveLeavePattern = getCurrentConsecutiveLeavePattern(staff.id, dateStr);
      let offScore = leaveDeficit * 100 + recentWorkPressure * 20 - sameDayLeaveLoad * 25;
      if (consecutiveLeavePattern.adjacentLeaveCount === 1) offScore += 18;
      if (consecutiveLeavePattern.adjacentLeaveCount >= 2) offScore -= 10;
      candidates.push({
        type: 'self-leave',
        staffId: staff.id,
        staffName: staff.name,
        group,
        shiftCode: defaultAutoLeaveCode,
        allowed: true,
        score: offScore,
        reasons: ['本月休假尚未達標', `當日群組需求已滿，優先補休（${defaultAutoLeaveCode}）`]
      });
    }

    candidates.sort((a, b) => b.score - a.score);

    setSelectedFillCell({ staffId: staff.id, staffName: staff.name, dateStr, group: staff.group });
    setFillCandidates(candidates);
    setShowFillModal(true);
  };

const openSelectedCellFillModal = () => {
    if (!selectedGridCell) return;
    openFillModal(selectedGridCell.staff, selectedGridCell.dateStr);
  };

  const clearSelectedCell = () => {
    if (!selectedGridCell) return;
    const { staff, dateStr } = selectedGridCell;
    const currentCode = getCellCode(staff.id, dateStr);
    if (!currentCode) return;
    if (!window.confirm(`確定清除此格內容？\n${staff.name}｜${dateStr}｜${currentCode}`)) return;

    setSchedule(prev => ({
      ...prev,
      [staff.id]: { ...prev[staff.id], [dateStr]: null }
    }));
    setSelectedGridCell(null);
    setRuleFillFeedback(`🧹 已清除 ${staff.name} 在 ${dateStr} 的內容`);
  };

  const clearRangeCells = () => {
    if (ruleFillConfig.selectedStaffs.length === 0) {
      setRuleFillFeedback('⚠️ 請先選擇要清除的人員');
      return;
    }

    const start = Number(ruleFillConfig.dateRange.start || 1);
    const end = Number(ruleFillConfig.dateRange.end || 31);
    const targetStaffIds = new Set(ruleFillConfig.selectedStaffs);

    let cleared = 0;
    setSchedule(prev => {
      const next = JSON.parse(JSON.stringify(prev));
      staffs.forEach(staff => {
        if (!targetStaffIds.has(staff.id)) return;
        daysInMonth.forEach(day => {
          if (day.day < start || day.day > end) return;
          const cellData = next[staff.id]?.[day.date];
          if (!cellData) return;

          const source = typeof cellData === 'object' && cellData !== null ? (cellData.source || 'manual') : 'manual';
          if (rangeClearMode === 'autoOnly' && source !== 'auto') return;

          next[staff.id][day.date] = null;
          cleared += 1;
        });
      });
      return next;
    });

    setRuleFillFeedback(cleared > 0 ? `🧹 已清除 ${cleared} 格內容` : 'ℹ️ 指定範圍內沒有可清除的內容');
  };

  const applyFillCandidate = (candidate) => {
    if (!selectedFillCell) return;
    handleCellChange(candidate.staffId, selectedFillCell.dateStr, candidate.shiftCode);
    setShowFillModal(false);
    setSelectedFillCell(null);
    setFillCandidates([]);
    setSelectedGridCell(null);
  };

  const addStaff = (group = '白班') => {
    const newId = 's' + Date.now();
    setStaffs(prev => [...prev, { id: newId, name: '新成員', group }]);
    setSchedule(prev => ({ ...prev, [newId]: {} }));
    setSelectedGridCell(null);
    setRangeSelection(null);
    setSelectionAnchor(null);
    setIsRangeDragging(false);
    resetKeyInputBuffer();
  };

  const removeStaff = (staffId) => {
    if (!window.confirm("確定要刪除此人員嗎？")) return;
    setStaffs(prev => prev.filter(s => s.id !== staffId));
    setSchedule(prev => {
      const next = { ...prev };
      delete next[staffId];
      return next;
    });
    setSelectedGridCell(null);
    setRangeSelection(null);
    setSelectionAnchor(null);
    setIsRangeDragging(false);
    resetKeyInputBuffer();
  };

  const moveStaffInGroup = (staffId, direction) => {
    const newStaffs = [...staffs];
    const currentIndex = newStaffs.findIndex(s => s.id === staffId);
    if (currentIndex === -1) return;

    const currentGroup = newStaffs[currentIndex].group;
    const groupIndexes = newStaffs
      .map((staff, index) => ({ staff, index }))
      .filter(item => item.staff.group === currentGroup)
      .map(item => item.index);

    const currentGroupPos = groupIndexes.indexOf(currentIndex);
    const targetGroupPos = direction === 'up' ? currentGroupPos - 1 : currentGroupPos + 1;
    if (targetGroupPos < 0 || targetGroupPos >= groupIndexes.length) return;

    const targetIndex = groupIndexes[targetGroupPos];
    [newStaffs[currentIndex], newStaffs[targetIndex]] = [newStaffs[targetIndex], newStaffs[currentIndex]];
    setStaffs(newStaffs);
    setSelectedGridCell(null);
    setRangeSelection(null);
    setSelectionAnchor(null);
    setIsRangeDragging(false);
    resetKeyInputBuffer();
  };

  const commitEditingStaffName = (staffId, nextName) => {
    const trimmedName = String(nextName ?? '').trim();
    setStaffs(prev => {
      const next = [...prev];
      const currentIndex = next.findIndex(s => s.id === staffId);
      if (currentIndex !== -1) next[currentIndex].name = trimmedName || '新成員';
      return next;
    });
    setEditingStaffId(null);
    setEditingNameDraft('');
  };

  const groupedStaffs = useMemo(() => {
    return SHIFT_GROUPS.map(group => ({
      group,
      staffs: staffs.filter(staff => (staff.group || '白班') === group)
    }));
  }, [staffs]);

  useEffect(() => {
    setSelectedGridCell(null);
    setRangeSelection(null);
    setSelectionAnchor(null);
    setIsRangeDragging(false);
    setEditingStaffId(null);
    setEditingNameDraft('');
    resetKeyInputBuffer();
  }, [staffs.length]);

  return (
    <div className="min-h-screen text-slate-900 p-3 font-sans overflow-x-hidden relative" style={{ backgroundColor: pageBackgroundColor }}>
      <style>{`
        @keyframes pulse-once { 0% { transform: translateY(-10px); opacity: 0; } 100% { transform: translateY(0); opacity: 1; } }
        @keyframes fade-in-down { 0% { opacity: 0; transform: translateY(-5px); } 100% { opacity: 1; transform: translateY(0); } }
        .animate-pulse-once { animation: pulse-once 0.5s ease-out forwards; }
        .animate-fade-in-down { animation: fade-in-down 0.3s ease-out forwards; }
      `}</style>

      {showDraftPrompt && (
        <div className="max-w-[98vw] mx-auto mb-3 bg-amber-50 border border-amber-200 text-amber-800 p-3 rounded-xl flex items-center justify-between shadow-sm animate-fade-in-down">
          <div className="flex items-center gap-1.5">
            <Clock size={18} className="text-amber-600" />
            <span className="text-sm font-bold">偵測到先前暫存紀錄。</span>
          </div>
          <div className="flex gap-2">
            <button onClick={() => loadHistory(historyList[0])} className="text-sm bg-amber-600 text-white px-3 py-1.5 rounded-lg hover:bg-amber-700 transition font-bold">載入最新</button>
            <button onClick={() => setShowDraftPrompt(false)} className="text-sm text-amber-700 hover:bg-amber-100 px-3 py-1.5 rounded-lg transition">忽略</button>
          </div>
        </div>
      )}

      <div className="max-w-[98vw] mx-auto mb-4">
        <div className="bg-white rounded-2xl p-4 shadow-sm border border-slate-200 flex flex-col xl:flex-row xl:items-center justify-between gap-3">
          <div>
            <h1 className="text-xl font-black text-slate-800 flex items-center gap-2">
              智慧補班系統｜開發版
              <span className="text-blue-500 text-xs font-normal px-2 py-0.5 bg-blue-50 rounded-lg border border-blue-100">PRO v1.6.0</span>
            </h1>
            <p className="text-slate-500 text-xs mt-1 italic">開發測試使用</p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <input
              ref={draftImportInputRef}
              type="file"
              accept=".json,application/json"
              className="hidden"
              onChange={onImportDraftFileChange}
            />
            <button onClick={() => saveToHistory('手動暫存')} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-1.5 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm">
              <Save size={16} /> 暫存
            </button>
            <button onClick={() => setShowHistoryModal(true)} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-1.5 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm">
              <Clock size={16} /> 歷史
            </button>
            <button onClick={onDownloadDraftFile} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-1.5 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm">
              <Download size={16} /> 下載工作檔
            </button>
            <button onClick={onImportDraftFileClick} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-1.5 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm">
              <Database size={16} /> 開啟工作檔
            </button>

            <div className="relative">
              <button onClick={() => setShowExportMenu(!showExportMenu)} className="flex items-center gap-1.5 bg-slate-800 text-white px-3 py-1.5 rounded-xl font-bold hover:bg-slate-900 transition-all text-sm">
                <Download size={16} /> 匯出
              </button>
              {showExportMenu && (
                <div className="absolute right-0 mt-2 w-48 bg-white border border-slate-200 rounded-xl shadow-xl z-50 overflow-hidden animate-fade-in-down">
                  <button onClick={exportToExcel} className="w-full text-left px-4 py-3 text-sm font-bold text-slate-700 hover:bg-green-50 hover:text-green-700 flex items-center gap-2 transition-colors border-b">
                    <FileSpreadsheet size={16} /> 高品質 Excel
                  </button>
                  <button onClick={exportToWord} className="w-full text-left px-4 py-3 text-sm font-bold text-slate-700 hover:bg-blue-50 hover:text-blue-700 flex items-center gap-2 transition-colors">
                    <FileText size={16} /> 橫向 Word (列印)
                  </button>
                </div>
              )}
            </div>

            <div className="w-px h-8 bg-slate-200 mx-2 hidden sm:block"></div>

            <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200 flex-wrap">
              <button onClick={() => handleRuleBasedAutoSchedule(false)} disabled={isRuleFillLoading} className="flex items-center gap-2 bg-white text-blue-600 px-3 py-1.5 rounded-lg font-bold hover:bg-blue-50 transition-all disabled:opacity-50 text-xs">
                {isRuleFillLoading ? <Loader2 className="animate-spin" size={14} /> : <Sparkles size={14} />} 規則全月補空
              </button>
              <button onClick={() => setShowRuleFillControl(!showRuleFillControl)} className={`flex items-center gap-2 px-3 py-1.5 rounded-lg font-bold transition-all text-xs ${showRuleFillControl ? 'bg-blue-600 text-white shadow-inner' : 'text-slate-600 hover:bg-slate-200'}`}>
                <Calendar size={14} /> 規則指定補空
              </button>
              <button
                type="button"
                onClick={openSelectedCellFillModal}
                disabled={!selectedGridCell}
                className={`flex items-center gap-2 px-3 py-1.5 rounded-lg font-bold transition-all text-xs ${selectedGridCell ? 'text-slate-700 hover:bg-slate-200' : 'text-slate-400 cursor-not-allowed'}`}
              >
                <Check size={14} /> 補此格
              </button>
              <button
                type="button"
                onClick={clearSelectedCell}
                disabled={!selectedGridCell}
                className={`flex items-center gap-2 px-3 py-1.5 rounded-lg font-bold transition-all text-xs ${selectedGridCell ? 'text-red-600 hover:bg-red-50' : 'text-slate-400 cursor-not-allowed'}`}
              >
                <Trash2 size={14} /> 清除此格
              </button>
            </div>
          </div>
        </div>
      </div>

      {showRuleFillControl && (
        <div className="max-w-[98vw] mx-auto mb-4 rounded-3xl border border-slate-200 bg-slate-100/90 px-4 py-4 shadow-sm animate-fade-in-down lg:px-5">
          <div className="mb-3 flex items-center gap-2 text-slate-800">
            <Sparkles size={18} className="text-blue-600" />
            <h3 className="font-black">規則指定補空設定</h3>
          </div>

          <div className="grid gap-5 xl:grid-cols-[1.35fr_0.95fr_1.05fr_1.15fr]">
            <div className="min-w-0">
              <label className="mb-2 block text-xs font-bold text-blue-700">1. 選擇人員（補空範圍）</label>
              <div className="max-h-[296px] space-y-2 overflow-y-auto pr-1">
                {groupedStaffs.map(({ group, staffs: groupStaffs }) => {
                  const groupIds = groupStaffs.map(s => s.id);
                  const isGroupFullySelected = groupIds.length > 0 && groupIds.every(id => ruleFillConfig.selectedStaffs.includes(id));

                  return (
                    <div key={group} className="rounded-xl border border-blue-100 bg-white/75 px-3 py-2 shadow-sm">
                      <div className="mb-1.5 flex items-center gap-2">
                        <span className="text-sm font-black text-slate-700">{group}：</span>
                        <button
                          type="button"
                          onClick={() => {
                            const next = isGroupFullySelected
                              ? ruleFillConfig.selectedStaffs.filter(id => !groupIds.includes(id))
                              : Array.from(new Set([...ruleFillConfig.selectedStaffs, ...groupIds]));
                            setRuleFillConfig({ ...ruleFillConfig, selectedStaffs: next });
                          }}
                          className={`rounded-lg border px-2.5 py-1 text-[11px] font-black transition-all ${isGroupFullySelected ? 'border-blue-600 bg-blue-600 text-white shadow-sm' : 'border-blue-200 bg-blue-50 text-blue-700 hover:bg-blue-100'}`}
                        >
                          全選
                        </button>
                      </div>

                      <div className="grid grid-cols-3 gap-1.5">
                        {groupStaffs.map(s => (
                          <button
                            key={s.id}
                            type="button"
                            onClick={() => {
                              const next = ruleFillConfig.selectedStaffs.includes(s.id)
                                ? ruleFillConfig.selectedStaffs.filter(id => id !== s.id)
                                : [...ruleFillConfig.selectedStaffs, s.id];
                              setRuleFillConfig({ ...ruleFillConfig, selectedStaffs: next });
                            }}
                            className={`min-h-[38px] rounded-lg border px-2.5 py-1.5 text-left text-xs font-bold leading-tight transition-all ${ruleFillConfig.selectedStaffs.includes(s.id) ? 'bg-blue-600 border-blue-600 text-white shadow-sm' : 'bg-white border-blue-200 text-blue-600 hover:bg-blue-50'}`}
                          >
                            {s.name}
                          </button>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            <div>
              <label className="mb-2 block text-xs font-bold text-blue-700">2. 日期範圍（1 ~ 31 號）</label>
              <div className="flex items-center gap-3">
                <input
                  type="number"
                  min="1"
                  max="31"
                  value={ruleFillConfig.dateRange.start}
                  onChange={(e) => setRuleFillConfig({ ...ruleFillConfig, dateRange: { ...ruleFillConfig.dateRange, start: parseInt(e.target.value, 10) || 1 } })}
                  className="w-full rounded-xl border border-blue-200 bg-white px-2 py-2 text-center text-sm font-bold text-slate-800"
                />
                <span className="shrink-0 text-sm font-bold text-slate-500">至</span>
                <input
                  type="number"
                  min="1"
                  max="31"
                  value={ruleFillConfig.dateRange.end}
                  onChange={(e) => setRuleFillConfig({ ...ruleFillConfig, dateRange: { ...ruleFillConfig.dateRange, end: parseInt(e.target.value, 10) || 31 } })}
                  className="w-full rounded-xl border border-blue-200 bg-white px-2 py-2 text-center text-sm font-bold text-slate-800"
                />
              </div>
            </div>

            <div>
              <label className="mb-2 block text-xs font-bold text-blue-700">3. 指定班別（選填）</label>
              <select
                value={ruleFillConfig.targetShift}
                onChange={(e) => setRuleFillConfig({ ...ruleFillConfig, targetShift: e.target.value })}
                className="w-full rounded-xl border border-blue-200 bg-white px-2 py-2 text-sm font-bold text-slate-800"
              >
                <option value="">依群組需求自動補空</option>
                {RULE_FILL_MAIN_SHIFTS.map(s => <option key={s} value={s}>{s} 班</option>)}
              </select>
            </div>

            <div className="space-y-3">
              <div>
                <label className="mb-2 block text-xs font-bold text-blue-700">4. 範圍清除模式</label>
                <select
                  value={rangeClearMode}
                  onChange={(e) => setRangeClearMode(e.target.value)}
                  className="w-full rounded-xl border border-blue-200 bg-white px-2 py-2 text-sm font-bold text-slate-800"
                >
                  <option value="autoOnly">只清除自動補入內容</option>
                  <option value="all">清除範圍內全部內容</option>
                </select>
              </div>

              <div className="grid grid-cols-1 gap-3 pt-1">
                <button
                  disabled={isRuleFillLoading || ruleFillConfig.selectedStaffs.length === 0}
                  onClick={() => handleRuleBasedAutoSchedule(true)}
                  className="flex w-full items-center justify-center gap-2 rounded-2xl bg-slate-400 py-2.5 font-black text-white transition-all hover:bg-slate-500 active:scale-95 disabled:opacity-50 disabled:grayscale"
                >
                  {isRuleFillLoading ? <Loader2 className="animate-spin" size={18} /> : <Check size={18} />} 套用規則補空
                </button>
                <button
                  type="button"
                  disabled={ruleFillConfig.selectedStaffs.length === 0}
                  onClick={clearRangeCells}
                  className="flex w-full items-center justify-center gap-2 rounded-2xl border border-slate-300 bg-slate-200/80 py-2.5 font-black text-slate-500 transition-all hover:bg-slate-300 disabled:opacity-50 disabled:grayscale"
                >
                  <Trash2 size={18} /> 範圍清除
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-[98vw] mx-auto mb-4 grid grid-cols-1 lg:grid-cols-12 gap-3">
        <div className="lg:col-span-7 bg-white p-3 rounded-xl shadow-sm border border-slate-200">
          <div className="flex flex-col lg:flex-row lg:items-center gap-4 lg:gap-5">
            <div className="flex flex-wrap items-center gap-x-2 gap-y-3 text-sm font-bold text-slate-700">
              <input
                type="number"
                value={year}
                onChange={(e) => setYear(Number(e.target.value) || new Date().getFullYear())}
                className="w-24 border border-slate-300 rounded-lg px-2 py-1.5 text-center font-bold bg-white text-slate-800"
              />
              <span className="shrink-0">年</span>

              <input
                type="number"
                min="1"
                max="12"
                step="1"
                value={month}
                onChange={(e) => {
                  const nextMonth = Number(e.target.value);
                  if (!Number.isFinite(nextMonth)) return;
                  setMonth(Math.min(12, Math.max(1, nextMonth)));
                }}
                className="w-20 border border-slate-300 rounded-lg px-2 py-1.5 text-center font-bold bg-white text-slate-800"
              />
              <span className="shrink-0">月</span>

              <div className="hidden lg:block w-px h-7 bg-slate-200 mx-2"></div>

              <span className="shrink-0 text-slate-600">應休天數</span>
              <span className="text-base font-black text-slate-800 tabular-nums">{requiredLeaves}</span>
              <span className="shrink-0">天</span>
            </div>
          </div>
        </div>

        <div className="lg:col-span-5 flex items-center justify-end gap-2">
          <button onClick={() => changeScreen('entry')} className="bg-white border px-2 py-2 rounded-xl hover:bg-slate-50 transition-colors text-sm font-bold text-slate-700 flex items-center gap-2">
            <ArrowLeft size={18} className="text-slate-600" /> 回入口頁
          </button>
          <button onClick={() => changeScreen('settings')} className="bg-white border px-2 py-2 rounded-xl hover:bg-slate-50 transition-colors text-sm font-bold text-slate-700 flex items-center gap-2">
            <Settings size={18} className="text-slate-600" /> 系統設定
          </button>
        </div>
      </div>

      <div className="max-w-[98vw] mx-auto rounded-2xl shadow-xl border border-slate-200 bg-white">
        <div className="overflow-auto rounded-2xl max-h-[calc(100vh-150px)]">
          <table className="w-max min-w-full border-collapse select-none">
            <thead>
              <tr className="bg-slate-100 border-b-2 border-slate-200 shadow-sm">
                <th className={`sticky left-0 top-0 z-50 border-r font-black shadow-sm ${shiftColumnFontSizeClass}`} style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor, color: shiftColumnFontColor }}>班別</th>
                <th className={`sticky top-0 z-50 border-r font-black shadow-sm ${nameDateColumnFontSizeClass}`} style={{ left: densityConfig.shiftWidth, width: effectiveDensityConfig.nameWidth, minWidth: effectiveDensityConfig.nameWidth, backgroundColor: nameDateColumnBgColor, color: nameDateColumnFontColor }}>姓名/日期</th>
                {daysInMonth.map(d => (
                  <th
                    key={d.day}
                    className={`sticky top-0 z-40 ${densityConfig.dayHeaderClass} border-r text-center shadow-sm`}
                    style={{
                      minWidth: densityConfig.dayMinWidth,
                      backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : '#f1f5f9'),
                      ...getFourWeekDividerStyle(d.date)
                    }}
                  >
                    <div className={`${tableFontSizeClass} opacity-60 uppercase`} style={{ color: tableFontColor }}>{d.weekStr}</div>
                    <div className={`${tableFontSizeClass} font-black`} style={{ color: tableFontColor }}>{d.day}</div>
                  </th>
                ))}
                {showRightStats && (
                <>
                <th className={`sticky top-0 z-40 ${densityConfig.statHeaderClass} border-r min-w-[52px] bg-blue-50 text-blue-700 font-bold shadow-sm`}>上班</th>
                <th className={`sticky top-0 z-40 ${densityConfig.statHeaderClass} border-r min-w-[52px] bg-green-50 text-green-700 font-bold shadow-sm`}>假日休</th>
                <th className={`sticky top-0 z-40 ${densityConfig.statHeaderClass} border-r min-w-[52px] bg-red-50 text-red-700 font-bold shadow-sm`}>總休</th>
                </>
                )}
                {showLeaveStats && mergedLeaveCodes.map(l => (
                  <th key={l} className={`sticky top-0 z-40 ${densityConfig.leaveHeaderClass} border-r min-w-[34px] bg-slate-50 text-[10px] uppercase text-slate-500 font-bold shadow-sm`}>{l}</th>
                ))}
                {(customColumns || []).map(col => (
                  <th key={col} className={`sticky top-0 z-40 ${densityConfig.leaveHeaderClass} border-r min-w-[60px] bg-violet-50 text-[10px] uppercase text-violet-600 font-bold shadow-sm`}>{col}</th>
                ))}
              </tr>
            </thead>

            <tbody>
              {groupedStaffs.map(({ group, staffs: groupStaffList }) => {
                const isCollapsed = Boolean(collapsedGroups[group]);
                const visibleGroupStaffList = isCollapsed ? [] : groupStaffList;
                return (
                <React.Fragment key={group}>
                  {visibleGroupStaffList.map((staff, index) => {
                    const stats = getStaffStats(staff.id);
                    const groupCount = visibleGroupStaffList.length + 1;
                    const groupIndex = visibleGroupStaffList.findIndex(s => s.id === staff.id);

                    return (
                      <tr key={staff.id} className="border-b border-slate-100 hover:bg-slate-50/50 transition-colors">
                        {index === 0 && (
                          <td rowSpan={groupCount} className="sticky left-0 z-20 border-r text-center shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)]" style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor }}>
                            <div className="flex items-center justify-center h-full" style={{ minHeight: densityConfig.rowMinHeight }}>
                              {showShiftLabels && (
                                <span
                                  className={`${shiftColumnFontSizeClass} font-black leading-none tracking-0 [writing-mode:vertical-rl]`}
                                  style={{ color: shiftColumnFontColor, fontSize: shiftCellLabelFontSize }}
                                >
                                  {group}
                                </span>
                              )}
                            </div>
                          </td>
                        )}

                        <td className="sticky z-30 border-r shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)] px-0.5 py-0.5" style={{ left: densityConfig.shiftWidth, width: effectiveDensityConfig.nameWidth, minWidth: effectiveDensityConfig.nameWidth, backgroundColor: nameDateColumnBgColor }}>
                          <div className="flex items-center gap-0.5">
                            <div className="flex flex-col items-center justify-center shrink-0 w-3">
                              <button
                                onClick={() => moveStaffInGroup(staff.id, 'up')}
                                disabled={groupIndex === 0}
                                className="text-slate-400 hover:text-blue-500 disabled:opacity-10 leading-none"
                              >
                                <ArrowUp size={14} />
                              </button>
                              <button
                                onClick={() => moveStaffInGroup(staff.id, 'down')}
                                disabled={groupIndex === visibleGroupStaffList.length - 1}
                                className="text-slate-400 hover:text-blue-500 disabled:opacity-10 leading-none"
                              >
                                <ArrowDown size={14} />
                              </button>
                            </div>

                            {editingStaffId === staff.id ? (
                            <div
                              contentEditable
                              suppressContentEditableWarning
                              ref={(node) => {
                                if (node && editingStaffId === staff.id) {
                                  requestAnimationFrame(() => {
                                    node.focus();
                                    const selection = window.getSelection();
                                    const range = document.createRange();
                                    range.selectNodeContents(node);
                                    range.collapse(false);
                                    selection.removeAllRanges();
                                    selection.addRange(range);
                                  });
                                }
                              }}
                              onInput={(e) => setEditingNameDraft(e.currentTarget.textContent || '')}
                              onBlur={(e) => commitEditingStaffName(staff.id, e.currentTarget.textContent || editingNameDraft)}
                              onMouseDown={(e) => e.stopPropagation()}
                              onClick={(e) => e.stopPropagation()}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') {
                                  e.preventDefault();
                                  commitEditingStaffName(staff.id, e.currentTarget.textContent || editingNameDraft);
                                } else if (e.key === 'Escape') {
                                  e.preventDefault();
                                  setEditingStaffId(null);
                                  setEditingNameDraft('');
                                }
                              }}
                              className={`flex-1 min-w-0 text-center py-0 px-0.5 font-bold bg-transparent whitespace-nowrap outline-none ${nameDateColumnFontSizeClass}`}
                              style={{ color: nameDateColumnFontColor, letterSpacing: "-0.02em", maxWidth: '100%' }}
                            >
                              {editingNameDraft}
                            </div>
                          ) : (
                            <button
                              type="button"
                              onClick={() => {
                                setEditingStaffId(staff.id);
                                setEditingNameDraft(staff.name || '');
                              }}
                              className={`flex-1 min-w-0 bg-transparent border-none text-center font-bold whitespace-nowrap px-0.5 py-0 ${nameDateColumnFontSizeClass}`}
                              style={{ color: nameDateColumnFontColor, letterSpacing: "-0.02em" }}
                              title="點擊編輯姓名"
                            >
                              <span className="block truncate">{staff.name}</span>
                            </button>
                          )}

                            <button
                              onClick={() => removeStaff(staff.id)}
                              className="text-slate-400 hover:text-red-500 shrink-0 w-3 flex items-center justify-center"
                            >
                              <Minus size={14} />
                            </button>
                          </div>
                        </td>

                        {daysInMonth.map(d => {
                          const cellData = schedule[staff.id]?.[d.date];
                          const val = typeof cellData === 'object' && cellData !== null ? (cellData?.value || '') : (cellData || '');
                          const cellKey = makeCellKey(staff.id, d.date);
                          const draftValue = cellDrafts[cellKey];
                          const displayValue = draftValue !== undefined ? draftValue : val;
                          const effectiveSelection = rangeSelection?.start && rangeSelection?.end ? rangeSelection : (selectedGridCell?.staff?.id && selectedGridCell?.dateStr ? {
                            start: { staffId: selectedGridCell.staff.id, dateStr: selectedGridCell.dateStr },
                            end: { staffId: selectedGridCell.staff.id, dateStr: selectedGridCell.dateStr }
                          } : null);
                          const inRangeSelection = isCellInSelectionRect(effectiveSelection, staffs, daysInMonth, staff.id, d.date);
                          const isPrimarySelected = selectedGridCell?.staff?.id === staff.id && selectedGridCell?.dateStr === d.date;
                          const isInvalid = Boolean(invalidCellKeys[cellKey]);

                          return (
                            <td
                              key={d.date}
                              className={`border-r p-0 relative overflow-hidden ${inRangeSelection ? 'ring-2 ring-violet-400 ring-inset' : isPrimarySelected ? 'ring-2 ring-blue-500 ring-inset' : ''} ${isInvalid ? 'ring-2 ring-red-400 ring-inset' : ''}`}
                              style={{
                                backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent'),
                                opacity: d.isHoliday || d.isWeekend ? 0.9 : 1,
                                ...getFourWeekDividerStyle(d.date)
                              }}
                              onMouseDown={(e) => {
                                if (e.button !== 0) return;
                                e.preventDefault();
                                setIsRangeDragging(true);
                                startRangeSelection(staff, d.date, e);
                              }}
                              onMouseEnter={() => updateRangeSelection(staff, d.date)}
                              onClick={() => {
                                startRangeSelection(staff, d.date);
                              }}
                            >
                              <div className="relative">
                                <div
                                  className={`w-full ${densityConfig.cellHeightClass} text-center bg-transparent border-none font-bold flex items-center justify-center ${tableFontSizeClass}`}
                                  style={{ color: tableFontColor, pointerEvents: 'none' }}
                                >
                                  {val}
                                </div>
                                <div
                                  className="absolute right-1 top-1/2 -translate-y-1/2 z-0 w-3.5 h-3.5 flex items-center justify-center"
                                  title="選擇班別/假別"
                                >
                                  <select
                                    value={val}
                                    onChange={(e) => {
                                      handleCellChange(staff.id, d.date, e.target.value);
                                      startRangeSelection(staff, d.date);
                                      e.currentTarget.blur();
                                    }}
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      startRangeSelection(staff, d.date);
                                    }}
                                    onMouseDown={(e) => {
                                      if (e.button !== 0) return;
                                      e.stopPropagation();
                                      startRangeSelection(staff, d.date, e);
                                    }}
                                    className="absolute inset-0 w-full h-full border-none bg-transparent cursor-pointer opacity-0"
                                    style={{ color: tableFontColor }}
                                    aria-label={`選擇 ${staff.name} ${d.date} 班別/假別`}
                                  >
                                    <option value=""></option>
                                    <optgroup label="上班">
                                      {DICT.SHIFTS.map(s => <option key={s} value={s}>{s}</option>)}
                                    </optgroup>
                                    <optgroup label="休假">
                                      {mergedLeaveCodes.map(l => <option key={l} value={l}>{l}</option>)}
                                    </optgroup>
                                  </select>

                                  {showBlueDots && (
                                    <span
                                      className={`${densityConfig.selectorDotClass} rounded-full transition-all pointer-events-none ${inRangeSelection ? 'bg-violet-600 scale-110' : isPrimarySelected ? 'bg-blue-700 scale-110' : 'bg-blue-300/90'}`}
                                    ></span>
                                  )}
                                </div>
                              </div>
                            </td>
                          );
                        })}

                        {showRightStats && (
                        <>
                        <td className={`border-r text-center font-black bg-blue-50/30 ${tableFontSizeClass}`} style={{ color: tableFontColor }}>{stats.work}</td>
                        <td className={`border-r text-center font-black bg-green-50/30 ${tableFontSizeClass}`} style={{ color: tableFontColor }}>{stats.holidayLeave}</td>
                        <td className={`border-r text-center font-black bg-red-50/30 ${tableFontSizeClass}`} style={{ color: tableFontColor }}>{stats.totalLeave}</td>
                        </>
                        )}
                        {showLeaveStats && mergedLeaveCodes.map(l => (
                          <td key={l} className={`border-r text-center bg-slate-50/20 ${tableFontSizeClass}`} style={{ color: tableFontColor }}>
                            {stats.leaveDetails[l] || ''}
                          </td>
                        ))}
                        {(customColumns || []).map(col => (
                          <td key={col} className={`border-r bg-violet-50/10 ${tableFontSizeClass}`} style={{ color: tableFontColor }}>
                            <input
                              type="text"
                              value={customColumnValues?.[staff.id]?.[col] || ''}
                              onChange={(e) => setCustomColumnValues(prev => ({ ...prev, [staff.id]: { ...(prev?.[staff.id] || {}), [col]: e.target.value } }))}
                              className="w-full h-full px-2 py-1 text-center bg-transparent border-none focus:ring-1 focus:ring-violet-300"
                              placeholder=""
                            />
                          </td>
                        ))}
                      </tr>
                    );
                  })}

                  {!isCollapsed && (
                  <tr className="border-b border-slate-200 bg-slate-50/70">
                    <td className="sticky z-30 border-r shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)] px-0.5 py-0.5" style={{ left: densityConfig.shiftWidth, width: effectiveDensityConfig.nameWidth, minWidth: effectiveDensityConfig.nameWidth, backgroundColor: nameDateColumnBgColor }}>
                      <div className="flex items-center justify-center">
                        <button
                          onClick={() => addStaff(group)}
                          className="bg-blue-600 text-white px-4 py-2 rounded-xl hover:bg-blue-700 transition-colors flex items-center gap-2 font-bold text-sm"
                        >
                          <Plus size={16} /> 新增
                        </button>
                      </div>
                    </td>

                    {daysInMonth.map(d => (
                      <td
                        key={d.date}
                        className="border-r"
                        style={{
                          backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent'),
                          opacity: d.isHoliday || d.isWeekend ? 0.9 : 1,
                          ...getFourWeekDividerStyle(d.date)
                        }}
                      />
                    ))}

                    <td colSpan={(showRightStats ? 3 : 0) + (showLeaveStats ? mergedLeaveCodes.length : 0) + (customColumns?.length || 0)}></td>
                  </tr>
                  )}

                  <tr className="bg-amber-50/95 border-b border-slate-200">
                    <td className={`sticky left-0 z-30 border-r ${densityConfig.footCellPaddingClass}`} style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor, top: stickyGroupSummaryTop, boxShadow: stickyGroupSummaryShadow }}>
                      <div className="flex items-center justify-center">
                        <button
                          type="button"
                          onClick={() => setCollapsedGroups(prev => ({ ...prev, [group]: !prev[group] }))}
                          className="inline-flex items-center justify-center w-7 h-7 rounded-md border border-slate-300 bg-white text-slate-600 hover:bg-slate-50 hover:text-slate-800 transition-colors"
                          title={isCollapsed ? `展開${group}` : `收合${group}`}
                          aria-label={isCollapsed ? `展開${group}` : `收合${group}`}
                        >
                          {isCollapsed ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
                        </button>
                      </div>
                    </td>
                    <td className={`sticky z-30 border-r text-right font-bold ${nameDateColumnFontSizeClass} ${densityConfig.footCellPaddingClass}`} style={{ left: densityConfig.shiftWidth, width: effectiveDensityConfig.nameWidth, minWidth: effectiveDensityConfig.nameWidth, backgroundColor: nameDateColumnBgColor, color: nameDateColumnFontColor, top: stickyGroupSummaryTop, boxShadow: stickyGroupSummaryShadow }}>
                      {group === '白班' ? '白班上班' : group === '小夜' ? '小夜上班' : '大夜上班'}
                    </td>
                    {daysInMonth.map(d => {
                      const count = getDailyStats(d.date)[group === '白班' ? 'D' : group === '小夜' ? 'E' : 'N'];
                      const rowKey = group === '白班' ? 'D' : group === '小夜' ? 'E' : 'N';
                      return (
                        <td
                          key={d.date}
                          className={`sticky z-20 border-r p-2 text-center font-black ${tableFontSizeClass}`}
                          style={{ backgroundColor: '#fef3c7', color: tableFontColor, top: stickyGroupSummaryTop, boxShadow: stickyGroupSummaryShadow, ...getDemandHighlightStyle(d.date, rowKey, count), ...getFourWeekDividerStyle(d.date) }}
                        >
                          {count || ''}
                        </td>
                      );
                    })}
                    <td
                      colSpan={(showRightStats ? 3 : 0) + (showLeaveStats ? mergedLeaveCodes.length : 0) + (customColumns?.length || 0)}
                      className="sticky z-20"
                      style={{ backgroundColor: '#fef3c7', top: stickyGroupSummaryTop, boxShadow: stickyGroupSummaryShadow }}
                    ></td>
                  </tr>
                </React.Fragment>
              )})}
            </tbody>

            {showBottomStats && (
            <tfoot className="bg-slate-100 border-t-2 border-slate-200">
              <tr>
                <td className={`sticky left-0 z-10 border-r ${densityConfig.footCellPaddingClass}`} style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor }}></td>
                <td className={`sticky z-10 border-r text-right font-bold ${nameDateColumnFontSizeClass} ${densityConfig.footCellPaddingClass}`} style={{ left: densityConfig.shiftWidth, width: effectiveDensityConfig.nameWidth, minWidth: effectiveDensityConfig.nameWidth, backgroundColor: nameDateColumnBgColor, color: nameDateColumnFontColor }}>
                  休假人數
                </td>
                {daysInMonth.map(d => {
                  const count = getDailyStats(d.date).totalLeave;
                  return (
                    <td
                      key={d.date}
                      className={`border-r p-2 text-center font-black ${tableFontSizeClass}`}
                      style={{ color: tableFontColor, ...getFourWeekDividerStyle(d.date) }}
                    >
                      {count || ''}
                    </td>
                  );
                })}
                <td colSpan={3 + mergedLeaveCodes.length + (customColumns?.length || 0)}></td>
              </tr>
            </tfoot>
            )}
          </table>
        </div>
      </div>


      {showFillModal && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-[110] flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-xl overflow-hidden shadow-2xl animate-pulse-once">
            <div className="p-6 border-b flex justify-between items-center bg-slate-50">
              <div>
                <h3 className="font-black text-slate-800">補此格</h3>
                <p className="text-sm text-slate-500 mt-1">{selectedFillCell?.staffName}｜{selectedFillCell?.dateStr}</p>
              </div>
              <button onClick={() => { setShowFillModal(false); setSelectedFillCell(null); setFillCandidates([]); }} className="p-2 hover:bg-slate-200 rounded-full transition-colors">
                <X />
              </button>
            </div>
            <div className="p-4 max-h-[60vh] overflow-y-auto space-y-3">
              {fillCandidates.length === 0 ? (
                <div className="p-6 text-center text-slate-500">目前沒有可直接建議的班別或人員。</div>
              ) : fillCandidates.map((candidate, index) => (
                <div key={`${candidate.staffId}-${candidate.shiftCode}`} className="border rounded-2xl p-4 flex items-center justify-between gap-4 hover:border-blue-400 hover:bg-blue-50/40 transition-all">
                  <div>
                    <div className="font-black text-slate-800">{candidate.staffName} → {candidate.shiftCode}</div>
                    <div className="text-xs text-slate-500 mt-1">群組：{candidate.group}｜排序分數：{candidate.score}</div>
                    {candidate.reasons?.length > 0 && (
                      <div className="text-xs text-slate-500 mt-1">{candidate.reasons.join('｜')}</div>
                    )}
                  </div>
                  <button onClick={() => applyFillCandidate(candidate)} className="px-4 py-2 rounded-xl bg-blue-600 text-white font-bold hover:bg-blue-700">
                    {index === 0 ? '選擇推薦' : '選擇'}
                  </button>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {showHistoryModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-lg overflow-hidden shadow-2xl animate-pulse-once">
            <div className="p-6 border-b flex justify-between items-center bg-slate-50">
              <h3 className="font-black text-slate-800 flex items-center gap-2"><Clock /> 歷史存檔紀錄</h3>
              <button onClick={() => setShowHistoryModal(false)} className="p-2 hover:bg-slate-200 rounded-full transition-colors">
                <X />
              </button>
            </div>

            <div className="max-h-[60vh] overflow-y-auto p-4 space-y-3">
              {historyList.length === 0 ? (
                <p className="text-center py-10 text-slate-400 font-bold">目前尚無存檔紀錄</p>
              ) : (
                historyList.map(record => (
                  <div key={record.id} className="flex items-center justify-between p-4 border rounded-2xl hover:border-blue-400 hover:bg-blue-50/30 transition-all group">
                    <div>
                      <div className="font-black text-slate-800">
                        {record.label}
                        <span className="text-xs font-normal text-slate-400 ml-2">{record.state.year}/{record.state.month}</span>
                      </div>
                      <div className="text-xs text-slate-500">{record.timestamp}</div>
                    </div>
                    <button onClick={() => loadHistory(record)} className="bg-slate-100 text-slate-700 px-4 py-2 rounded-xl text-sm font-bold hover:bg-blue-600 hover:text-white transition-all">
                      載入
                    </button>
                  </div>
                ))
              )}
            </div>

            <div className="p-4 bg-slate-50 border-t flex gap-3">
              <button onClick={clearHistory} className="flex-1 py-3 text-sm font-bold text-red-500 hover:bg-red-50 rounded-xl transition-colors">清空所有紀錄</button>
              <button onClick={() => setShowHistoryModal(false)} className="flex-1 py-3 text-sm font-bold text-slate-600 bg-white border rounded-xl hover:bg-slate-100 transition-colors">關閉</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


function SettingRow({ icon: Icon, title, desc, children, iconBg = 'bg-blue-50', iconColor = 'text-blue-600' }) {
  return (
    <div className="bg-white border border-gray-200 rounded-2xl shadow-sm overflow-hidden">
      <div className="grid grid-cols-1 lg:grid-cols-[260px_minmax(0,1fr)]">
        <div className="px-6 py-6 bg-gray-50/80 border-b lg:border-b-0 lg:border-r border-gray-200">
          <div className="flex items-start gap-3">
            <div className={`p-2 rounded-xl ${iconBg}`}><Icon className={`w-4 h-5 ${iconColor}`} /></div>
            <div><h3 className="font-bold text-gray-800">{title}</h3><p className="text-sm text-gray-500 mt-1 leading-relaxed">{desc}</p></div>
          </div>
        </div>
        <div className="px-6 py-6">{children}</div>
      </div>
    </div>
  );
}

function SettingsView({ changeScreen, colors, setColors, customHolidays, setCustomHolidays, specialWorkdays, setSpecialWorkdays, medicalCalendarAdjustments, setMedicalCalendarAdjustments, staffingConfig, setStaffingConfig, uiSettings, setUiSettings, customLeaveCodes, setCustomLeaveCodes, customColumns, setCustomColumns, schedulingRulesText, setSchedulingRulesText }) {
  const [holidayInput, setHolidayInput] = useState({ year: '', month: '', day: '' });
  const mergedLeaveCodes = useMemo(() => Array.from(new Set([...(DICT.LEAVES || []), ...((customLeaveCodes || []))])), [customLeaveCodes]);
  const addCustomLeaveCode = () => {
    const raw = window.prompt('請輸入自訂休假代碼');
    const code = String(raw || "").trim();
    if (!code) return;
    if (DICT.LEAVES.includes(code) || (customLeaveCodes || []).includes(code)) return;
    setCustomLeaveCodes(prev => [...prev, code]);
  };
  const removeCustomLeaveCode = (code) => setCustomLeaveCodes(prev => prev.filter(item => item !== code));
  const addCustomColumn = () => {
    const raw = window.prompt('請輸入自訂欄位名稱');
    const name = String(raw || "").trim();
    if (!name) return;
    if ((customColumns || []).includes(name)) return;
    setCustomColumns(prev => [...prev, name]);
  };
  const removeCustomColumn = (name) => setCustomColumns(prev => prev.filter(item => item !== name));
  const addCustomHoliday = () => {
    const y = holidayInput.year.trim();
    const m = holidayInput.month.trim();
    const d = holidayInput.day.trim();
    if (!y || !m || !d) return;
    const dateStr = `${y.padStart(4,'0')}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;
    if (customHolidays.includes(dateStr)) return;
    setCustomHolidays(prev => [...prev, dateStr].sort());
    setHolidayInput({ year: '', month: '', day: '' });
  };
  const removeCustomHoliday = (dateStr) => setCustomHolidays(prev => prev.filter(item => item !== dateStr));
  return (
    <div className="min-h-screen bg-gray-50 text-slate-800 font-sans selection:bg-blue-100">
      <header className="sticky top-0 z-50 bg-white border-b border-gray-200 px-8 py-4 flex justify-between items-center shadow-sm">
        <div><div className="flex items-center gap-2 mb-0.5"><Settings className="w-6 h-6 text-blue-600" /><h1 className="text-xl font-bold tracking-tight text-gray-900">系統設定</h1></div><p className="text-sm text-gray-500">可調整使用者設定與畫面顯示參數。</p></div>
        <div className="flex items-center gap-3">
          <button onClick={() => changeScreen('entry')} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-600 bg-white border border-gray-300 rounded-lg hover:bg-gray-50"><ArrowLeft className="w-4 h-4" />返回入口頁</button>
          <button onClick={() => changeScreen('schedule')} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-600 bg-white border border-gray-300 rounded-lg hover:bg-gray-50"><Calendar className="w-4 h-4" />返回排班頁</button>
          <button onClick={() => changeScreen('schedule')} className="flex items-center gap-2 px-6 py-2 text-sm font-semibold text-white bg-blue-600 rounded-lg hover:bg-blue-700 shadow-md shadow-blue-100"><Save className="w-4 h-4" />儲存設定</button>
        </div>
      </header>
      <main className="max-w-7xl mx-auto p-8 space-y-10">
        <section className="space-y-6">
          <div className="flex items-center gap-2"><div className="w-1 h-6 bg-blue-600 rounded-full"></div><h2 className="text-lg font-bold text-gray-800">使用者偏好設定</h2></div>
          <div className="space-y-5">
            <SettingRow icon={Monitor} title="外觀與顯示" desc="調整班表顏色、顯示大小與統計欄位呈現。">
              <div className="space-y-6">
                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">色彩標示</label>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">週末顏色</span><input type="color" value={colors.weekend} onChange={(e) => setColors(prev => ({ ...prev, weekend: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">假日顏色</span><input type="color" value={colors.holiday} onChange={(e) => setColors(prev => ({ ...prev, holiday: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">主頁背景顏色</span><input type="color" value={uiSettings.pageBackgroundColor} onChange={(e) => setUiSettings(prev => ({ ...prev, pageBackgroundColor: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">表格字體顏色</span><input type="color" value={uiSettings.tableFontColor} onChange={(e) => setUiSettings(prev => ({ ...prev, tableFontColor: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">班別欄背景顏色</span><input type="color" value={uiSettings.shiftColumnBgColor} onChange={(e) => setUiSettings(prev => ({ ...prev, shiftColumnBgColor: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">姓名/日期欄背景顏色</span><input type="color" value={uiSettings.nameDateColumnBgColor} onChange={(e) => setUiSettings(prev => ({ ...prev, nameDateColumnBgColor: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                  </div>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="flex items-center justify-between"><span className="text-sm font-medium">表格顯示大小</span><select value={uiSettings.tableDensity} onChange={(e) => setUiSettings(prev => ({ ...prev, tableDensity: e.target.value }))} className="text-sm border-none bg-gray-100 rounded-md px-3 py-2"><option value="standard">標準 (預設)</option><option value="compact">緊湊</option><option value="relaxed">寬鬆</option></select></div>
                  <div className="flex items-center justify-between"><span className="text-sm font-medium">表格字體大小</span><select value={uiSettings.tableFontSize} onChange={(e) => setUiSettings(prev => ({ ...prev, tableFontSize: e.target.value }))} className="text-sm border-none bg-gray-100 rounded-md px-3 py-2"><option value="small">小</option><option value="medium">標準</option><option value="large">大</option></select></div>
                  <div className="flex items-center justify-between"><span className="text-sm font-medium">班別欄字體大小</span><select value={uiSettings.shiftColumnFontSize} onChange={(e) => setUiSettings(prev => ({ ...prev, shiftColumnFontSize: e.target.value }))} className="text-sm border-none bg-gray-100 rounded-md px-3 py-2"><option value="small">小</option><option value="medium">標準</option><option value="large">大</option></select></div>
                  <div className="flex items-center justify-between"><span className="text-sm font-medium">班別欄字體顏色</span><input type="color" value={uiSettings.shiftColumnFontColor} onChange={(e) => setUiSettings(prev => ({ ...prev, shiftColumnFontColor: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                  <div className="flex items-center justify-between"><span className="text-sm font-medium">姓名/日期欄字體大小</span><select value={uiSettings.nameDateColumnFontSize} onChange={(e) => setUiSettings(prev => ({ ...prev, nameDateColumnFontSize: e.target.value }))} className="text-sm border-none bg-gray-100 rounded-md px-3 py-2"><option value="small">小</option><option value="medium">標準</option><option value="large">大</option></select></div>
                  <div className="flex items-center justify-between"><span className="text-sm font-medium">姓名/日期欄字體顏色</span><input type="color" value={uiSettings.nameDateColumnFontColor} onChange={(e) => setUiSettings(prev => ({ ...prev, nameDateColumnFontColor: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                  <div className="flex items-center justify-between"><span className="text-sm font-medium">快速切換全部統計</span><button type="button" onClick={() => setUiSettings(prev => ({ ...prev, showStats: !prev.showStats, showRightStats: !prev.showStats, showLeaveStats: !prev.showStats, showBottomStats: !prev.showStats }))} className={`w-10 h-5 rounded-full relative transition-colors ${uiSettings.showStats ? 'bg-blue-600' : 'bg-gray-300'}`}><div className={`absolute top-1 w-3 h-3 bg-white rounded-full transition-all ${uiSettings.showStats ? 'right-1' : 'left-1'}`}></div></button></div>
                </div>
                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">需求警示顯示</label>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">需求超編顏色</span><input type="color" value={uiSettings.demandOverColor} onChange={(e) => setUiSettings(prev => ({ ...prev, demandOverColor: e.target.value, themePreset: 'custom' }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                  </div>
                </div>
              </div>

            </SettingRow>
            <SettingRow icon={Settings} title="使用偏好" desc="快速切換主題、欄位顯示、操作方式與預設補休代碼。" iconBg="bg-violet-50" iconColor="text-violet-600">
              <div className="space-y-6">
                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">主題預設</label>
                  <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
                    {[
                      ['classic','經典藍白'],
                      ['soft','柔和綠灰'],
                      ['warm','米色護理站'],
                      ['dark','深色模式'],
                      ['custom','自訂主題']
                    ].map(([key,label]) => (
                      <button
                        key={key}
                        type="button"
                        onClick={() => {
                          if (key === 'custom') {
                            setUiSettings(prev => ({ ...prev, themePreset: 'custom' }));
                            return;
                          }
                          const preset = UI_THEME_PRESETS[key];
                          setColors(prev => ({ ...prev, weekend: preset.weekendColor, holiday: preset.holidayColor }));
                          setUiSettings(prev => ({
                            ...prev,
                            themePreset: key,
                            pageBackgroundColor: preset.pageBackgroundColor,
                            tableFontColor: preset.tableFontColor,
                            shiftColumnBgColor: preset.shiftColumnBgColor,
                            nameDateColumnBgColor: preset.nameDateColumnBgColor,
                            shiftColumnFontColor: preset.shiftColumnFontColor,
                            nameDateColumnFontColor: preset.nameDateColumnFontColor
                          }));
                        }}
                        className={`px-3 py-1.5 rounded-xl border text-sm font-medium transition ${uiSettings.themePreset === key ? 'bg-violet-600 border-violet-600 text-white' : 'bg-white border-violet-200 text-violet-700 hover:bg-violet-50'}`}
                      >
                        {label}
                      </button>
                    ))}
                  </div>
                </div>

                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">欄位顯示</label>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    {[
                      ['showRightStats','顯示右側統計'],
                      ['showLeaveStats','顯示休假別統計'],
                      ['showBottomStats','顯示下方每日統計'],
                      ['showBlueDots','顯示藍點提示'],
                      ['showShiftLabels','顯示班別群組大字']
                    ].map(([key,label]) => (
                      <div key={key} className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100">
                        <span className="text-sm font-medium">{label}</span>
                        <button
                          type="button"
                          onClick={() => setUiSettings(prev => ({ ...prev, [key]: !prev[key] }))}
                          className={`w-10 h-5 rounded-full relative transition-colors ${uiSettings[key] ? 'bg-violet-600' : 'bg-gray-300'}`}
                        >
                          <div className={`absolute top-1 w-3 h-3 bg-white rounded-full transition-all ${uiSettings[key] ? 'right-1' : 'left-1'}`}></div>
                        </button>
                      </div>
                    ))}
                  </div>
                </div>

                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">欄寬與高度</label>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">班別欄寬</span><select value={uiSettings.shiftColumnWidthMode} onChange={(e)=>setUiSettings(prev=>({...prev, shiftColumnWidthMode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-3 py-2"><option value="narrow">窄</option><option value="standard">標準</option><option value="wide">寬</option></select></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">姓名/日期欄寬</span><select value={uiSettings.nameDateColumnWidthMode} onChange={(e)=>setUiSettings(prev=>({...prev, nameDateColumnWidthMode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-3 py-2"><option value="narrow">窄</option><option value="standard">標準</option><option value="wide">寬</option></select></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">日期欄寬</span><select value={uiSettings.dayColumnWidthMode} onChange={(e)=>setUiSettings(prev=>({...prev, dayColumnWidthMode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-3 py-2"><option value="narrow">窄</option><option value="standard">標準</option><option value="wide">寬</option></select></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">儲存格高度</span><select value={uiSettings.cellHeightMode} onChange={(e)=>setUiSettings(prev=>({...prev, cellHeightMode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-3 py-2"><option value="compact">緊湊</option><option value="standard">標準</option><option value="roomy">寬鬆</option></select></div>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">系統補休預設代碼</span><select value={uiSettings.defaultAutoLeaveCode} onChange={(e)=>setUiSettings(prev=>({...prev, defaultAutoLeaveCode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-3 py-2">{mergedLeaveCodes.filter(code => ['off','休','例'].includes(code)).map(code => <option key={code} value={code}>{code}</option>)}</select></div>
                  <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">選格操作模式</span><select value={uiSettings.selectionMode} onChange={(e)=>setUiSettings(prev=>({...prev, selectionMode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-3 py-2"><option value="dot">點藍點選格</option><option value="cell">點格選取</option></select></div>
                </div>
              </div>
            </SettingRow>
            <SettingRow icon={Layout} title="班表內容自訂" desc="設定自訂休假代碼與延伸欄位。" iconBg="bg-indigo-50" iconColor="text-indigo-600">
              <div className="space-y-5">
                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">自訂休假代碼</label>
                  <div className="flex flex-wrap items-center gap-2">
                    <button type="button" onClick={addCustomLeaveCode} className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-md border border-gray-300 bg-white text-sm font-medium text-gray-700 hover:bg-gray-50">
                      <Plus className="w-3.5 h-3.5" /> 新增
                    </button>
                    {(customLeaveCodes || []).length === 0 ? <div className="text-xs text-gray-400">尚未新增自訂休假代碼</div> : (customLeaveCodes || []).map(code => (
                      <span key={code} className="inline-flex items-center gap-2 px-3 py-1.5 bg-gray-100 text-gray-600 text-xs font-bold rounded-md border border-gray-200">{code}<button type="button" onClick={() => removeCustomLeaveCode(code)} className="text-red-500 hover:text-red-600"><Minus className="w-3.5 h-3.5" /></button></span>
                    ))}
                  </div>
                  <div className="mt-2 text-xs text-gray-500">新增後會同步出現在主頁休假下拉選單，並視為休假類代碼。</div>
                </div>
                <div className="pt-3 border-t border-gray-100 space-y-3">
                  <div>
                    <button type="button" onClick={addCustomColumn} className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-md border border-gray-300 bg-white text-sm font-medium text-gray-700 hover:bg-gray-50">
                      <Plus className="w-3.5 h-3.5" /> 新增自訂欄位
                    </button>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {(customColumns || []).length === 0 ? <div className="text-xs text-gray-400">尚未新增自訂欄位</div> : (customColumns || []).map(col => (
                      <span key={col} className="inline-flex items-center gap-2 px-3 py-1.5 bg-violet-50 text-violet-700 text-xs font-bold rounded-md border border-violet-200">{col}<button type="button" onClick={() => removeCustomColumn(col)} className="text-red-500 hover:text-red-600"><Minus className="w-3.5 h-3.5" /></button></span>
                    ))}
                  </div>
                  <div className="text-xs text-gray-500">新增後會同步出現在主頁右側，作為延伸紀錄欄位。可用來記錄如門診、支援、教學、行政或其他單位自訂資訊。</div>
                </div>
                <div className="pt-3 border-t border-gray-100 space-y-3">
                  <div>
                    <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">匯出用排班規則</label>
                    <textarea
                      value={schedulingRulesText}
                      onChange={(e) => setSchedulingRulesText(e.target.value)}
                      rows={6}
                      placeholder={`請逐行輸入排班規則\n例如：\n白班每日至少 6 人\n小夜不跨白班支援`}
                      className="w-full rounded-xl border border-gray-200 bg-gray-50 px-3 py-3 text-sm text-gray-700 focus:outline-none focus:ring-2 focus:ring-indigo-100"
                    />
                  </div>
                  <div className="text-xs text-gray-500">這裡輸入的內容不會同步到主頁顯示，但會在匯出 Word 時顯示於最下方，格式為：排班規則：1.XXX 2.XXX。</div>
                </div>
              </div>
            </SettingRow>
            <SettingRow icon={UserCheck} title="規則補空需求設定" desc="設定平日 / 假日各班需求，作為規則全月補空與規則指定補空的直接依據。" iconBg="bg-sky-50" iconColor="text-sky-600">
              <div className="space-y-6">
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                  <div className="rounded-2xl border border-gray-200 p-4 bg-gray-50/50">
                    <h4 className="font-bold text-gray-800 mb-4">平日需求</h4>
                    <div className="grid grid-cols-3 gap-3">
                      <div>
                        <label className="text-xs font-bold text-gray-400 block mb-2">白班</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.white}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, white: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-2 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-bold text-gray-400 block mb-2">小夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.evening}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, evening: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-2 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-bold text-gray-400 block mb-2">大夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.night}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, night: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-2 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                    </div>
                  </div>

                  <div className="rounded-2xl border border-gray-200 p-4 bg-gray-50/50">
                    <h4 className="font-bold text-gray-800 mb-4">假日需求</h4>
                    <div className="grid grid-cols-3 gap-3">
                      <div>
                        <label className="text-xs font-bold text-gray-400 block mb-2">白班</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.holiday.white}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, holiday: { ...prev.requiredStaffing.holiday, white: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-2 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-bold text-gray-400 block mb-2">小夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.holiday.evening}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, holiday: { ...prev.requiredStaffing.holiday, evening: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-2 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-bold text-gray-400 block mb-2">大夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.holiday.night}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, holiday: { ...prev.requiredStaffing.holiday, night: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-2 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                    </div>
                  </div>
                </div>

                <div className="rounded-xl border border-gray-200 bg-white px-4 py-3 text-sm text-gray-600">
                  <div className="font-semibold text-gray-800 mb-1">目前補空依據</div>
                  <div>平日：白班 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.weekday.white}</span> 人、小夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.weekday.evening}</span> 人、大夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.weekday.night}</span> 人</div>
                  <div>假日：白班 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.holiday.white}</span> 人、小夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.holiday.evening}</span> 人、大夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.holiday.night}</span> 人</div>
                </div>
              </div>
            </SettingRow>
            <SettingRow icon={Calendar} title="假期新增" desc="使用西曆年月日新增自訂假期，並可個別刪除。">
              <div className="space-y-5"><div><label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">西曆年月日</label><div className="grid grid-cols-1 md:grid-cols-3 gap-3"><input type="number" placeholder="年" value={holidayInput.year} onChange={(e)=>setHolidayInput({ ...holidayInput, year: e.target.value })} className="w-full px-2 py-2 text-sm border border-gray-200 rounded-xl bg-gray-50 focus:outline-none focus:ring-2 focus:ring-blue-100" /><input type="number" placeholder="月" value={holidayInput.month} onChange={(e)=>setHolidayInput({ ...holidayInput, month: e.target.value })} className="w-full px-2 py-2 text-sm border border-gray-200 rounded-xl bg-gray-50 focus:outline-none focus:ring-2 focus:ring-blue-100" /><input type="number" placeholder="日" value={holidayInput.day} onChange={(e)=>setHolidayInput({ ...holidayInput, day: e.target.value })} className="w-full px-2 py-2 text-sm border border-gray-200 rounded-xl bg-gray-50 focus:outline-none focus:ring-2 focus:ring-blue-100" /></div></div><button onClick={addCustomHoliday} className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-md border border-gray-300 bg-white text-sm font-medium text-gray-700 hover:bg-gray-50"><Plus className="w-4 h-4" /> 新增假期</button><div className="pt-3 border-t border-gray-100"><label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">已新增假期</label><div className="space-y-2 max-h-52 overflow-y-auto pr-1">{customHolidays.length === 0 ? <div className="text-xs text-gray-400 p-4 bg-gray-50 border border-dashed border-gray-300 rounded-xl text-center">尚未新增自訂假期</div> : customHolidays.map(dateStr => <div key={dateStr} className="flex items-center justify-between p-3 bg-gray-50 border border-gray-200 rounded-xl"><span className="text-sm text-gray-700 font-medium">{dateStr}</span><button onClick={() => removeCustomHoliday(dateStr)} className="w-8 h-8 flex items-center justify-center rounded-full border border-red-200 text-red-500 hover:bg-red-50 font-bold">-</button></div>)}</div></div></div>
            </SettingRow>
          </div>
        </section>
      </main>
      <footer className="max-w-7xl mx-auto px-8 py-12 border-t border-gray-200 flex flex-col md:flex-row justify-between items-center gap-4"><div className="flex items-center gap-4"><div className="flex items-center gap-2"><div className="w-3 h-3 bg-green-500 rounded-full animate-pulse"></div><span className="text-sm font-semibold text-gray-700">系統核心版本: v2.5.0-PRO</span></div><span className="text-gray-300">|</span><span className="text-sm text-gray-500">規則式補班引擎開發測試版</span></div><div className="text-sm text-gray-400">© 2024 Intelligent Scheduling System PRO. All rights reserved.</div></footer>
    </div>
  );
}

function EntryView({ changeScreen, goToLatestHistory, onImportFiles, hasActiveDraft, activeDraftMeta, restoreActiveDraft, discardActiveDraft }) {
  const importInputRef = useRef(null);

  const handleImportButtonClick = () => {
    importInputRef.current?.click();
  };

  const handleImportInputChange = async (event) => {
    const files = Array.from(event.target.files || []);
    event.target.value = '';
    if (files.length === 0) return;

    const invalidFile = files.find((file) => {
      const fileName = String(file.name || '').toLowerCase();
      return !fileName.endsWith('.xlsx') && !fileName.endsWith('.xls');
    });
    if (invalidFile) {
      window.alert('目前僅支援 Excel 檔案（.xlsx / .xls）');
      return;
    }

    try {
      await onImportFiles?.(files);
    } catch (error) {
      console.error('匯入檔案失敗:', error);
      window.alert(error?.message || '匯入檔案失敗，請確認是否使用系統範本。');
    }
  };

  const handleDownloadImportTemplate = async () => {
    try {
      const ExcelJS = await loadExcelJS();
      const workbook = new ExcelJS.Workbook();
      const dataSheet = workbook.addWorksheet('班表匯入範本');
      const guideSheet = workbook.addWorksheet('填寫說明');

      const templateMonth = new Date().getMonth() + 1;
      const dayHeaders = Array.from({ length: 31 }, (_, i) => `${i + 1}日`);
      const headers = ['姓名', '班別群組', ...dayHeaders];

      const templateTheme = {
        titleBg: '#EFF6FF',
        titleFont: '#1F2937',
        shiftBg: '#F8FAFC',
        shiftFont: '#1E293B',
        nameBg: '#FFFFFF',
        nameFont: '#1E293B',
        dayBg: '#F8FAFC'
      };

      dataSheet.addRow([`${templateMonth}月班表匯入範本`]);
      dataSheet.mergeCells(1, 1, 1, headers.length);
      const titleCell = dataSheet.getCell(1, 1);
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(templateTheme.titleBg, '#EFF6FF') } };
      titleCell.font = { bold: true, size: 14, color: { argb: hexToExcelArgb(templateTheme.titleFont, '#1F2937') } };
      dataSheet.getRow(1).height = 24;

      const headerRow = dataSheet.addRow(headers);
      headerRow.height = 24;
      headerRow.eachCell((cell, colNumber) => {
        cell.font = { bold: true, size: 10, color: { argb: hexToExcelArgb(templateTheme.titleFont, '#1F2937') } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' },
          bottom: { style: 'thin' }, right: { style: 'thin' }
        };

        if (colNumber === 1) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(templateTheme.shiftBg, '#F8FAFC') } };
          cell.font = { ...(cell.font || {}), color: { argb: hexToExcelArgb(templateTheme.shiftFont, '#1E293B') } };
        } else if (colNumber === 2) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(templateTheme.nameBg, '#FFFFFF') } };
          cell.font = { ...(cell.font || {}), color: { argb: hexToExcelArgb(templateTheme.nameFont, '#1E293B') } };
        } else {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(templateTheme.dayBg, '#F8FAFC') } };
        }
      });

      const sampleRows = [
        ['王小美', '白班'],
        ['李小芳', '小夜'],
        ['陳小君', '大夜']
      ];
      sampleRows.forEach((rowData) => {
        const row = dataSheet.addRow(rowData);
        row.eachCell((cell) => {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' },
            bottom: { style: 'thin' }, right: { style: 'thin' }
          };
        });
      });

      dataSheet.columns = [
        { width: 14 },
        { width: 12 },
        ...Array.from({ length: 31 }, () => ({ width: 6 }))
      ];
      dataSheet.views = [{ state: 'frozen', ySplit: 2, xSplit: 2 }];

      const guideRows = [
        ['排班匯入範本說明'],
        ['1. 請依此範本填寫，避免欄位遺漏或格式不一致。'],
        ['2. 必填欄位為：姓名、班別群組。'],
        ['3. 班別群組請填：白班 / 小夜 / 大夜。'],
        ['4. 日期欄可填班別代碼或假別代碼，例如：D、E、N、off、例、休。'],
        ['5. 若當月不足31天，超出日期欄可留白。'],
        ['6. 匯入功能完成後，建議保留此範本格式，不要自行增刪欄位。']
      ];
      guideRows.forEach(r => guideSheet.addRow(r));
      guideSheet.getCell('A1').font = { bold: true, size: 13 };
      guideSheet.getColumn(1).width = 90;

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = '排班匯入範本.xlsx';
      a.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error('下載匯入範本失敗:', error);
      window.alert('下載匯入範本失敗，請稍後再試。');
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 font-sans text-slate-800">
      <div className="mb-10 text-center">
        <div className="flex items-center justify-center gap-3 mb-4">
          <div className="bg-blue-600 p-2.5 rounded-xl shadow-md shadow-blue-200/50">
            <CalendarDays className="w-8 h-8 text-white" />
          </div>
          <div className="text-left">
            <h1 className="text-2xl font-bold tracking-tight text-slate-900">智慧排班系統</h1>
            <p className="text-slate-500 text-sm font-medium tracking-wide mt-1">護理排班管理平台</p>
          </div>
        </div>
      </div>

      <div className="w-full max-w-[440px] bg-white rounded-2xl shadow-[0_12px_40px_rgba(0,0,0,0.03)] border border-slate-200/60 overflow-hidden">
        <div className="p-10">
          {hasActiveDraft && (
            <div className="mb-6 rounded-2xl border border-amber-200 bg-amber-50 p-4">
              <div className="flex items-start gap-2">
                <Clock className="mt-0.5 h-4 w-4 text-amber-600" />
                <div className="min-w-0">
                  <div className="text-sm font-bold text-amber-800">偵測到上次未完成工作</div>
                  <div className="mt-1 text-xs leading-relaxed text-amber-700">
                    {activeDraftMeta?.savedAtText ? `上次自動暫存：${activeDraftMeta.savedAtText}` : '可恢復上次未完成的排班進度。'}
                  </div>
                </div>
              </div>
              <div className="mt-3 flex gap-2">
                <button
                  type="button"
                  onClick={restoreActiveDraft}
                  className="flex-1 rounded-xl bg-amber-600 px-3 py-2 text-sm font-bold text-white hover:bg-amber-700"
                >
                  恢復未完成進度
                </button>
                <button
                  type="button"
                  onClick={discardActiveDraft}
                  className="rounded-xl border border-amber-200 bg-white px-3 py-2 text-sm font-bold text-amber-700 hover:bg-amber-100"
                >
                  捨棄
                </button>
              </div>
            </div>
          )}
          <div className="mb-8 text-center">
            <h2 className="text-xl font-bold text-slate-900 mb-2">系統入口</h2>
            <p className="text-sm text-slate-500 leading-relaxed">
              請選擇要進入的功能。
            </p>
          </div>

          <input
            ref={importInputRef}
            type="file"
            multiple
            accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
            className="hidden"
            onChange={handleImportInputChange}
          />

          <div className="space-y-4">
            <button
              type="button"
              onClick={() => changeScreen('schedule')}
              className="w-full flex justify-center items-center gap-2 py-4 px-4 rounded-xl text-sm font-bold text-white bg-blue-600 hover:bg-blue-700"
            >
              <ShieldCheck className="w-4.5 h-4.5" />
              進入排班系統
            </button>

            <button
              type="button"
              onClick={goToLatestHistory}
              className="w-full flex justify-center items-center gap-2 py-3.5 px-4 border border-slate-200 rounded-xl shadow-sm text-sm font-bold text-slate-700 bg-white hover:bg-slate-50"
            >
              <Clock className="w-4 h-4 text-slate-500" />
              開啟最近班表
            </button>

            <button
              type="button"
              onClick={handleImportButtonClick}
              className="w-full flex justify-center items-center gap-2 py-3.5 px-4 border border-slate-200 rounded-xl shadow-sm text-sm font-bold text-slate-700 bg-white hover:bg-slate-50"
            >
              <Database className="w-4 h-4 text-slate-500" />
              匯入檔案
            </button>

            <button
              type="button"
              onClick={handleDownloadImportTemplate}
              className="w-full flex justify-center items-center gap-2 py-3.5 px-4 border border-slate-200 rounded-xl shadow-sm text-sm font-bold text-slate-700 bg-white hover:bg-slate-50"
            >
              <Download className="w-4 h-4 text-slate-500" />
              下載匯入範本
            </button>

            <button
              type="button"
              onClick={() => window.alert('建議定期匯出 Excel / Word 或後續備份檔，以保留班表資料。')}
              className="w-full flex justify-center items-center gap-2 py-3 px-4 rounded-xl text-xs font-bold text-amber-700 bg-amber-50 border border-amber-200 hover:bg-amber-100"
            >
              <Info className="w-4 h-4" />
              建議定期備份
            </button>
          </div>
        </div>
      </div>

      <div className="mt-12 text-center">
        <p className="text-xs text-slate-400 font-medium tracking-[0.1em] uppercase">&copy; {new Date().getFullYear()} 智能排班系統 PRO. ALL RIGHTS RESERVED.</p>
      </div>
    </div>
  );
}

export default function App() {
  const [screen, setScreen] = useState('entry');
  const [colors, setColors] = useState({ weekend: '#dcfce7', holiday: '#fca5a5' });
  const [customHolidays, setCustomHolidays] = useState([]);
  const [specialWorkdays, setSpecialWorkdays] = useState([]);
  const [medicalCalendarAdjustments, setMedicalCalendarAdjustments] = useState({ holidays: [], workdays: [] });
  const [uiSettings, setUiSettings] = useState({
    pageBackgroundColor: '#f8fafc',
    tableFontSize: 'medium',
    tableFontColor: '#1f2937',
    shiftColumnFontSize: 'medium',
    shiftColumnFontColor: '#1e293b',
    nameDateColumnFontSize: 'medium',
    nameDateColumnFontColor: '#1e293b',
    shiftColumnBgColor: '#ffffff',
    nameDateColumnBgColor: '#ffffff',
    tableDensity: 'standard',
    showStats: true,
    themePreset: 'custom',
    showRightStats: true,
    showLeaveStats: true,
    showBottomStats: true,
    showBlueDots: true,
    showShiftLabels: true,
    defaultAutoLeaveCode: 'off',
    selectionMode: 'dot',
    shiftColumnWidthMode: 'standard',
    nameDateColumnWidthMode: 'standard',
    dayColumnWidthMode: 'standard',
    cellHeightMode: 'standard',
    demandOverColor: '#fde68a'
  });
  const [staffingConfig, setStaffingConfig] = useState({
    hospitalLevel: 'regional',
    totalBeds: 60,
    totalNurses: 20,
    requiredStaffing: {
      weekday: { white: 6, evening: 3, night: 2 },
      holiday: { white: 4, evening: 2, night: 2 }
    }
  });
  const [customLeaveCodes, setCustomLeaveCodes] = useState([]);
  const [customColumns, setCustomColumns] = useState([]);
  const [customColumnValues, setCustomColumnValues] = useState({});
  const [schedulingRulesText, setSchedulingRulesText] = useState('');
  const [loadLatestOnEnter, setLoadLatestOnEnter] = useState(false);
  const [importedSchedulePayload, setImportedSchedulePayload] = useState(null);
  const [monthlySchedules, setMonthlySchedules] = useState({});
  const [pendingOpenMonthKey, setPendingOpenMonthKey] = useState('');
  const [year, setYear] = useState(2025);
  const [month, setMonth] = useState(3);
  const [staffs, setStaffs] = useState(() => createBlankMonthState(2025, 3).staffs);
  const [schedule, setSchedule] = useState(() => createBlankMonthState(2025, 3).schedule);
  const [hasActiveDraft, setHasActiveDraft] = useState(false);
  const [activeDraftMeta, setActiveDraftMeta] = useState(null);
  const activeDraftHydratedRef = useRef(false);
  const activeDraftSaveReadyRef = useRef(false);
  const draftImportInputRef = useRef(null);

  const formatDraftSavedAt = (isoString) => {
    if (!isoString) return '';
    const date = new Date(isoString);
    if (Number.isNaN(date.getTime())) return '';
    return date.toLocaleString();
  };

  const buildWorkspaceState = () => ({
    colors,
    customHolidays,
    specialWorkdays,
    medicalCalendarAdjustments,
    uiSettings,
    staffingConfig,
    customLeaveCodes,
    customColumns,
    customColumnValues,
    schedulingRulesText,
    monthlySchedules,
    year,
    month,
    staffs,
    schedule
  });

  const applyWorkspaceState = (state = {}) => {
    setColors(state.colors || { weekend: '#dcfce7', holiday: '#fca5a5' });
    setCustomHolidays(Array.isArray(state.customHolidays) ? state.customHolidays : []);
    setSpecialWorkdays(Array.isArray(state.specialWorkdays) ? state.specialWorkdays : []);
    setMedicalCalendarAdjustments(state.medicalCalendarAdjustments || { holidays: [], workdays: [] });
    setUiSettings(state.uiSettings || {
      pageBackgroundColor: '#f8fafc',
      tableFontSize: 'medium',
      tableFontColor: '#1f2937',
      shiftColumnFontSize: 'medium',
      shiftColumnFontColor: '#1e293b',
      nameDateColumnFontSize: 'medium',
      nameDateColumnFontColor: '#1e293b',
      shiftColumnBgColor: '#ffffff',
      nameDateColumnBgColor: '#ffffff',
      tableDensity: 'standard',
      showStats: true,
      themePreset: 'custom',
      showRightStats: true,
      showLeaveStats: true,
      showBottomStats: true,
      showBlueDots: true,
      showShiftLabels: true,
      defaultAutoLeaveCode: 'off',
      selectionMode: 'dot',
      shiftColumnWidthMode: 'standard',
      nameDateColumnWidthMode: 'standard',
      dayColumnWidthMode: 'standard',
      cellHeightMode: 'standard',
      demandOverColor: '#fde68a'
    });
    setStaffingConfig(state.staffingConfig || {
      hospitalLevel: 'regional',
      totalBeds: 60,
      totalNurses: 20,
      requiredStaffing: {
        weekday: { white: 6, evening: 3, night: 2 },
        holiday: { white: 4, evening: 2, night: 2 }
      }
    });
    setCustomLeaveCodes(Array.isArray(state.customLeaveCodes) ? state.customLeaveCodes : []);
    setCustomColumns(Array.isArray(state.customColumns) ? state.customColumns : []);
    setCustomColumnValues(state.customColumnValues || {});
    setSchedulingRulesText(typeof state.schedulingRulesText === 'string' ? state.schedulingRulesText : '');
    setMonthlySchedules(state.monthlySchedules || {});
    setYear(Number(state.year) || 2025);
    setMonth(Number(state.month) || 3);
    setStaffs(normalizeStaffGroup(state.staffs || createBlankMonthState(Number(state.year) || 2025, Number(state.month) || 3).staffs));
    setSchedule(state.schedule || createBlankMonthState(Number(state.year) || 2025, Number(state.month) || 3).schedule);
    setImportedSchedulePayload(null);
    setPendingOpenMonthKey('');
    setLoadLatestOnEnter(false);
  };

  useEffect(() => {
    try {
      const storedDraft = localStorage.getItem(ACTIVE_DRAFT_KEY);
      if (!storedDraft) {
        activeDraftHydratedRef.current = true;
        return;
      }
      const parsed = JSON.parse(storedDraft);
      if (parsed?.state) {
        setHasActiveDraft(true);
        setActiveDraftMeta({
          savedAt: parsed.savedAt || '',
          savedAtText: formatDraftSavedAt(parsed.savedAt),
          year: parsed.state.year,
          month: parsed.state.month
        });
      }
    } catch (error) {
      console.error('讀取自動暫存失敗', error);
    } finally {
      activeDraftHydratedRef.current = true;
    }
  }, []);

  useEffect(() => {
    if (!activeDraftHydratedRef.current) return;
    if (!activeDraftSaveReadyRef.current) {
      activeDraftSaveReadyRef.current = true;
      return;
    }

    const savedAt = new Date().toISOString();
    const payload = {
      savedAt,
      state: {
        colors,
        customHolidays,
        specialWorkdays,
        medicalCalendarAdjustments,
        uiSettings,
        staffingConfig,
        customLeaveCodes,
        customColumns,
        customColumnValues,
        schedulingRulesText,
        monthlySchedules,
        year,
        month,
        staffs,
        schedule
      }
    };

    try {
      localStorage.setItem(ACTIVE_DRAFT_KEY, JSON.stringify(payload));
      setHasActiveDraft(true);
      setActiveDraftMeta({ savedAt, savedAtText: formatDraftSavedAt(savedAt), year, month });
    } catch (error) {
      console.error('寫入自動暫存失敗', error);
    }
  }, [colors, customHolidays, specialWorkdays, medicalCalendarAdjustments, uiSettings, staffingConfig, customLeaveCodes, customColumns, customColumnValues, schedulingRulesText, monthlySchedules, year, month, staffs, schedule]);

  const restoreActiveDraft = () => {
    try {
      const storedDraft = localStorage.getItem(ACTIVE_DRAFT_KEY);
      if (!storedDraft) return;
      const parsed = JSON.parse(storedDraft);
      if (!parsed?.state) return;
      applyWorkspaceState(parsed.state);
      setHasActiveDraft(true);
      setActiveDraftMeta({
        savedAt: parsed.savedAt || '',
        savedAtText: formatDraftSavedAt(parsed.savedAt),
        year: parsed.state.year,
        month: parsed.state.month
      });
      setScreen('schedule');
    } catch (error) {
      console.error('恢復自動暫存失敗', error);
      window.alert('恢復未完成進度失敗，請稍後再試。');
    }
  };

  const discardActiveDraft = () => {
    if (!window.confirm('確定要捨棄上次未完成進度嗎？')) return;
    localStorage.removeItem(ACTIVE_DRAFT_KEY);
    setHasActiveDraft(false);
    setActiveDraftMeta(null);
  };

  const handleDownloadDraftFile = () => {
    try {
      const exportedAt = new Date().toISOString();
      const payload = {
        type: 'schedule-draft',
        version: '1.6.0',
        exportedAt,
        state: buildWorkspaceState()
      };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `排班工作檔_${year}年${String(month).padStart(2, '0')}月.json`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error('下載工作檔失敗', error);
      window.alert('下載工作檔失敗，請稍後再試。');
    }
  };

  const handleImportDraftFileClick = () => {
    if (draftImportInputRef.current) draftImportInputRef.current.click();
  };

  const handleImportDraftFileChange = async (event) => {
    const file = event.target?.files?.[0];
    if (!file) return;

    try {
      const rawText = await file.text();
      const parsed = JSON.parse(rawText);
      const importedState = parsed?.state || parsed;
      if (!importedState || typeof importedState !== 'object') {
        throw new Error('暫存檔格式不正確');
      }

      applyWorkspaceState(importedState);
      const savedAt = new Date().toISOString();
      const payload = { savedAt, state: importedState };
      localStorage.setItem(ACTIVE_DRAFT_KEY, JSON.stringify(payload));
      setHasActiveDraft(true);
      setActiveDraftMeta({
        savedAt,
        savedAtText: formatDraftSavedAt(savedAt),
        year: importedState.year,
        month: importedState.month
      });
      setScreen('schedule');
      window.alert('開啟工作檔成功，已載入目前工作內容。');
    } catch (error) {
      console.error('開啟工作檔失敗', error);
      window.alert('開啟工作檔失敗，請確認檔案格式是否正確。');
    } finally {
      if (event.target) event.target.value = '';
    }
  };

  const handleImportFiles = async (files) => {
    const imported = await parseImportedExcelFiles(files, new Date().getFullYear(), { customLeaveCodes });
    setImportedSchedulePayload(imported);
    setPendingOpenMonthKey(imported.firstMonthKey || '');
    setLoadLatestOnEnter(false);
    setScreen('schedule');
  };

  const goToSchedule = () => {
    setLoadLatestOnEnter(false);
    setPendingOpenMonthKey('');
    setScreen('schedule');
  };

  const goToLatestHistory = () => {
    setLoadLatestOnEnter(true);
    setPendingOpenMonthKey('');
    setScreen('schedule');
  };

  if (screen === 'schedule') {
    return (
      <ScheduleView
        changeScreen={setScreen}
        colors={colors}
        setColors={setColors}
        customHolidays={customHolidays}
        setCustomHolidays={setCustomHolidays}
        specialWorkdays={specialWorkdays}
        setSpecialWorkdays={setSpecialWorkdays}
        medicalCalendarAdjustments={medicalCalendarAdjustments}
        setMedicalCalendarAdjustments={setMedicalCalendarAdjustments}
        staffingConfig={staffingConfig}
        setStaffingConfig={setStaffingConfig}
        uiSettings={uiSettings}
        setUiSettings={setUiSettings}
        customLeaveCodes={customLeaveCodes}
        setCustomLeaveCodes={setCustomLeaveCodes}
        customColumns={customColumns}
        setCustomColumns={setCustomColumns}
        customColumnValues={customColumnValues}
        setCustomColumnValues={setCustomColumnValues}
        schedulingRulesText={schedulingRulesText}
        setSchedulingRulesText={setSchedulingRulesText}
        loadLatestOnEnter={loadLatestOnEnter}
        onLatestLoaded={() => setLoadLatestOnEnter(false)}
        importedSchedulePayload={importedSchedulePayload}
        onImportedScheduleApplied={() => setImportedSchedulePayload(null)}
        monthlySchedules={monthlySchedules}
        setMonthlySchedules={setMonthlySchedules}
        pendingOpenMonthKey={pendingOpenMonthKey}
        onPendingOpenHandled={() => setPendingOpenMonthKey('')}
        year={year}
        setYear={setYear}
        month={month}
        setMonth={setMonth}
        staffs={staffs}
        setStaffs={setStaffs}
        schedule={schedule}
        setSchedule={setSchedule}
        onDownloadDraftFile={handleDownloadDraftFile}
        onImportDraftFileClick={handleImportDraftFileClick}
        draftImportInputRef={draftImportInputRef}
        onImportDraftFileChange={handleImportDraftFileChange}
      />
    );
  }

  if (screen === 'settings') {
    return (
      <SettingsView
        changeScreen={setScreen}
        colors={colors}
        setColors={setColors}
        customHolidays={customHolidays}
        setCustomHolidays={setCustomHolidays}
        specialWorkdays={specialWorkdays}
        setSpecialWorkdays={setSpecialWorkdays}
        medicalCalendarAdjustments={medicalCalendarAdjustments}
        setMedicalCalendarAdjustments={setMedicalCalendarAdjustments}
        staffingConfig={staffingConfig}
        setStaffingConfig={setStaffingConfig}
        uiSettings={uiSettings}
        setUiSettings={setUiSettings}
        customLeaveCodes={customLeaveCodes}
        setCustomLeaveCodes={setCustomLeaveCodes}
        customColumns={customColumns}
        setCustomColumns={setCustomColumns}
        schedulingRulesText={schedulingRulesText}
        setSchedulingRulesText={setSchedulingRulesText}
      />
    );
  }

  return (
    <EntryView
      changeScreen={(target) => {
        if (target === 'schedule') goToSchedule();
        else setScreen(target);
      }}
      goToLatestHistory={goToLatestHistory}
      onImportFiles={handleImportFiles}
      hasActiveDraft={hasActiveDraft}
      activeDraftMeta={activeDraftMeta}
      restoreActiveDraft={restoreActiveDraft}
      discardActiveDraft={discardActiveDraft}
    />
  );
}
