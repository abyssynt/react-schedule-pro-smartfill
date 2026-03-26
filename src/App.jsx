
import React, { useState, useMemo, useEffect } from 'react';
import {
  Plus, Minus, Settings, Sparkles, Loader2,
  ArrowUp, ArrowDown, Save, History as Clock, Download,
  FileSpreadsheet, FileText, X, Check, Calendar, CalendarDays,
  User, Lock, Info, Layout, ShieldCheck, Grid, UserCheck,
  Database, Cpu, Monitor, ArrowLeft, ChevronRight, CheckCircle2, Trash2
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

const apiKey = "";

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

const AI_MAIN_SHIFTS = ['D', 'E', 'N'];

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
    nameWidth: 98,
    dayMinWidth: 32,
    dayHeaderClass: 'px-0.5 py-1 text-[11px]',
    statHeaderClass: 'p-1.5',
    leaveHeaderClass: 'p-1',
    cellHeightClass: 'h-8',
    nameCellPaddingClass: 'px-1.5 py-1',
    footCellPaddingClass: 'p-1.5',
    groupLabelClass: '',
    selectorDotClass: 'w-1.5 h-1.5',
    rowMinHeight: 72
  },
  standard: {
    shiftWidth: 76,
    nameWidth: 122,
    dayMinWidth: 42,
    dayHeaderClass: 'px-1.5 py-2 text-xs',
    statHeaderClass: 'p-3',
    leaveHeaderClass: 'p-1.5',
    cellHeightClass: 'h-10',
    nameCellPaddingClass: 'px-2 py-2',
    footCellPaddingClass: 'p-2.5',
    groupLabelClass: '',
    selectorDotClass: 'w-2 h-2',
    rowMinHeight: 80
  },
  relaxed: {
    shiftWidth: 100,
    nameWidth: 156,
    dayMinWidth: 56,
    dayHeaderClass: 'px-2 py-2.5 text-sm',
    statHeaderClass: 'p-4',
    leaveHeaderClass: 'p-2',
    cellHeightClass: 'h-12',
    nameCellPaddingClass: 'px-3 py-2.5',
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
    nameDateColumnFontColor: '#1e293b'
  },
  soft: {
    pageBackgroundColor: '#f7faf7',
    weekendColor: '#e7f7ec',
    holidayColor: '#f6c7c7',
    tableFontColor: '#334155',
    shiftColumnBgColor: '#f7fbf8',
    nameDateColumnBgColor: '#fcfdfc',
    shiftColumnFontColor: '#365314',
    nameDateColumnFontColor: '#334155'
  },
  warm: {
    pageBackgroundColor: '#fffaf5',
    weekendColor: '#fef3c7',
    holidayColor: '#fecaca',
    tableFontColor: '#44403c',
    shiftColumnBgColor: '#fff7ed',
    nameDateColumnBgColor: '#fffbeb',
    shiftColumnFontColor: '#7c2d12',
    nameDateColumnFontColor: '#44403c'
  },
  dark: {
    pageBackgroundColor: '#0f172a',
    weekendColor: '#334155',
    holidayColor: '#7f1d1d',
    tableFontColor: '#e2e8f0',
    shiftColumnBgColor: '#1e293b',
    nameDateColumnBgColor: '#172033',
    shiftColumnFontColor: '#f8fafc',
    nameDateColumnFontColor: '#e2e8f0'
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
    nameWidth: Math.max(90, baseConfig.nameWidth + nameAdjust),
    dayMinWidth: Math.max(28, baseConfig.dayMinWidth + dayAdjust),
    rowMinHeight: Math.max(72, (baseConfig.rowMinHeight || 80) + heightAdjust * 4),
    selectorDotClass: dotClassMap[uiSettings.cellHeightMode || 'standard'] || baseConfig.selectorDotClass
  };
};

function ScheduleView({ changeScreen, colors, setColors, customHolidays, setCustomHolidays, specialWorkdays, setSpecialWorkdays, medicalCalendarAdjustments, setMedicalCalendarAdjustments, staffingConfig, setStaffingConfig, uiSettings, setUiSettings, customLeaveCodes, setCustomLeaveCodes, customColumns, setCustomColumns, customColumnValues, setCustomColumnValues, loadLatestOnEnter, onLatestLoaded }) {
  // ==========================================
  // 2. 核心 State 定義
  // ==========================================
  const [year, setYear] = useState(2025);
  const [month, setMonth] = useState(3);
  const [isAiLoading, setIsAiLoading] = useState(false);

  const [staffs, setStaffs] = useState(normalizeStaffGroup([
    { id: 's1', name: '新成員' }, { id: 's2', name: '新成員' }, { id: 's3', name: '新成員' }, { id: 's4', name: '新成員' }, { id: 's5', name: '新成員' },
    { id: 's6', name: '新成員' }, { id: 's7', name: '新成員' }, { id: 's8', name: '新成員' }, { id: 's9', name: '新成員' }, { id: 's10', name: '新成員' },
    { id: 's11', name: '新成員' }, { id: 's12', name: '新成員' }, { id: 's13', name: '新成員' }, { id: 's14', name: '新成員' }, { id: 's15', name: '新成員' }
  ]));
  const [schedule, setSchedule] = useState({
    s1: {}, s2: {}, s3: {}, s4: {}, s5: {},
    s6: {}, s7: {}, s8: {}, s9: {}, s10: {},
    s11: {}, s12: {}, s13: {}, s14: {}, s15: {}
  });

  const [unitAdjustmentDraft, setUnitAdjustmentDraft] = useState({ holidays: [], workdays: [] });

  const [showAiControl, setShowAiControl] = useState(false);
  const [aiFeedback, setAiFeedback] = useState("");

  const [showHistoryModal, setShowHistoryModal] = useState(false);
  const [historyList, setHistoryList] = useState([]);
  const [showDraftPrompt, setShowDraftPrompt] = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [selectedFillCell, setSelectedFillCell] = useState(null);
  const [fillCandidates, setFillCandidates] = useState([]);
  const [showFillModal, setShowFillModal] = useState(false);
  const [selectedGridCell, setSelectedGridCell] = useState(null);
  const [rangeClearMode, setRangeClearMode] = useState('autoOnly');

  // AI 指定排班設定
  const [aiConfig, setAiConfig] = useState({
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

  const pageBackgroundColor = uiSettings?.pageBackgroundColor || '#f8fafc';
  const tableFontColor = uiSettings?.tableFontColor || '#1f2937';
  const shiftColumnFontColor = uiSettings?.shiftColumnFontColor || '#1e293b';
  const nameDateColumnFontColor = uiSettings?.nameDateColumnFontColor || '#1e293b';
  const shiftColumnBgColor = uiSettings?.shiftColumnBgColor || '#ffffff';
  const nameDateColumnBgColor = uiSettings?.nameDateColumnBgColor || '#ffffff';
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
  }, [loadLatestOnEnter]);

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

  // ==========================================
  // 4. Excel 匯出 (ExcelJS 實現)
  // ==========================================
  const exportToExcel = async () => {
    setAiFeedback("📊 正在產生高品質 Excel 報表...");
    const ExcelJS = await loadExcelJS();
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`${year}年${month}月班表`);

    const headerRow = ['班別', '日期/姓名', ...daysInMonth.map(d => `${d.day}\n(${d.weekStr})`), '上班', '假日休', '總休', ...mergedLeaveCodes, ...(customColumns || [])];
    const header = worksheet.addRow(headerRow);
    header.height = 30;

    header.eachCell((cell, colNumber) => {
      cell.font = { bold: true, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
      if (colNumber > 2 && colNumber <= daysInMonth.length + 2) {
        const d = daysInMonth[colNumber - 3];
        if (d.isHoliday) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCACA' } };
        else if (d.isWeekend) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
      }
    });

    staffs.forEach(staff => {
      const stats = getStaffStats(staff.id);
      const rowData = [
        staff.group,
        staff.name,
        ...daysInMonth.map(d => {
          const cellData = schedule[staff.id]?.[d.date];
          return typeof cellData === 'object' ? (cellData?.value || '') : (cellData || '');
        }),
        stats.work, stats.holidayLeave, stats.totalLeave,
        ...mergedLeaveCodes.map(l => stats.leaveDetails[l] || ''), ...(customColumns || []).map(col => customColumnValues?.[staff.id]?.[col] || '')
      ];
      const row = worksheet.addRow(rowData);

      row.eachCell((cell, colNumber) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' },
          bottom: { style: 'thin' }, right: { style: 'thin' }
        };
        if (colNumber > 2 && colNumber <= daysInMonth.length + 2) {
          cell.numFmt = '@';
          const d = daysInMonth[colNumber - 3];
          if (d.isHoliday) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE4E4' } };
          else if (d.isWeekend) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FDF4' } };
        }
      });
    });

    ['D', 'E', 'N', 'totalLeave'].forEach(rowKey => {
      const label = rowKey === 'totalLeave' ? '當日休假' : `${rowKey} 班人數`;
      const rowData = ['', label, ...daysInMonth.map(d => getDailyStats(d.date)[rowKey] || '')];
      const row = worksheet.addRow(rowData);
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' },
          bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });
    });

    worksheet.getColumn(1).width = 10;
    worksheet.getColumn(2).width = 15;
    for (let i = 3; i <= daysInMonth.length + 2; i++) worksheet.getColumn(i).width = 5;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `排班表_${year}年${month}月.xlsx`;
    a.click();
    setShowExportMenu(false);
    setAiFeedback("✅ Excel 導出成功！");
  };

  const exportToWord = () => {
    const html = `
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          @page { size: landscape; margin: 1cm; }
          body { font-family: sans-serif; }
          table { border-collapse: collapse; width: 100%; font-size: 9pt; }
          th, td { border: 1px solid #000; padding: 4px; text-align: center; }
          .holiday { background-color: ${colors.holiday}; }
          .weekend { background-color: ${colors.weekend}; }
        </style>
      </head>
      <body>
        <h2 style="text-align:center;">${year}年${month}月 班表</h2>
        <table>
          <thead>
            <tr>
              <th>班別</th>
              <th>姓名</th>
              ${daysInMonth.map(d => `<th class="${d.isHoliday ? 'holiday' : (d.isWeekend ? 'weekend' : '')}">${d.day}<br/>(${d.weekStr})</th>`).join('')}
              <th>上班</th>
              <th>總休</th>
            </tr>
          </thead>
          <tbody>
            ${staffs.map(staff => {
              const stats = getStaffStats(staff.id);
              return `
                <tr>
                  <td>${staff.group}</td>
                  <td>${staff.name}</td>
                  ${daysInMonth.map(d => {
                    const cellData = schedule[staff.id]?.[d.date];
                    return `<td>${typeof cellData === 'object' ? (cellData?.value || '') : (cellData || '')}</td>`;
                  }).join('')}
                  <td>${stats.work}</td>
                  <td>${stats.totalLeave}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      </body>
    </html>`;

    const blob = new Blob([html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `列印班表_${year}年${month}月.doc`;
    a.click();
    setShowExportMenu(false);
  };

  // ==========================================
  // 5. AI 指定排班功能
  // ==========================================
  const handleAiAutoSchedule = async (isPartial = false) => {
    setIsAiLoading(true);
    setAiFeedback(isPartial ? "🧩 系統正在依指定範圍補空..." : "🧩 系統正在依人力需求補全整月空白...");

    try {
      const mergedSchedule = JSON.parse(JSON.stringify(schedule));
      const targetStaffIds = isPartial && aiConfig.selectedStaffs.length > 0
        ? new Set(aiConfig.selectedStaffs)
        : new Set(staffs.map(s => s.id));

      const targetDays = daysInMonth.filter(d => {
        if (!isPartial) return true;
        return d.day >= aiConfig.dateRange.start && d.day <= aiConfig.dateRange.end;
      });

      const normalizedTargetShift = AI_MAIN_SHIFTS.includes(aiConfig.targetShift) ? aiConfig.targetShift : '';
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
      setAiFeedback(`✅ 補空完成：上班 ${summary.workFilled} 格、休假 ${summary.leaveFilled} 格、未補成功 ${summary.skipped} 格`);
    } catch (error) {
      console.error(error);
      setAiFeedback("❌ 規則補空失敗，請檢查設定。");
    } finally {
      setIsAiLoading(false);
    }
  };

const callGemini = async (prompt, systemInstruction = "") => {
    let delay = 1000;
    for (let i = 0; i < 5; i++) {
      try {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=${apiKey}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            systemInstruction: systemInstruction ? { parts: [{ text: systemInstruction }] } : undefined,
            generationConfig: { responseMimeType: "application/json" }
          })
        });
        if (!response.ok) throw new Error('API Error');
        const data = await response.json();
        return JSON.parse(data.candidates[0].content.parts[0].text);
      } catch (err) {
        if (i === 4) throw err;
        await new Promise(resolve => setTimeout(resolve, delay));
        delay *= 2;
      }
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

  const saveToHistory = (label, currentSchedule = schedule) => {
    const newRecord = {
      id: Date.now(),
      label,
      timestamp: new Date().toLocaleString(),
      state: { year, month, staffs, schedule: currentSchedule, colors, customHolidays, specialWorkdays, medicalCalendarAdjustments, staffingConfig, uiSettings, customLeaveCodes, customColumns, customColumnValues }
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
    setAiFeedback(`🧹 已清除 ${staff.name} 在 ${dateStr} 的內容`);
  };

  const clearRangeCells = () => {
    if (aiConfig.selectedStaffs.length === 0) {
      setAiFeedback('⚠️ 請先選擇要清除的人員');
      return;
    }

    const start = Number(aiConfig.dateRange.start || 1);
    const end = Number(aiConfig.dateRange.end || 31);
    const targetStaffIds = new Set(aiConfig.selectedStaffs);

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

    setAiFeedback(cleared > 0 ? `🧹 已清除 ${cleared} 格內容` : 'ℹ️ 指定範圍內沒有可清除的內容');
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
  };

  const removeStaff = (staffId) => {
    if (!window.confirm("確定要刪除此人員嗎？")) return;
    setStaffs(prev => prev.filter(s => s.id !== staffId));
    setSchedule(prev => {
      const next = { ...prev };
      delete next[staffId];
      return next;
    });
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
  };

  const groupedStaffs = useMemo(() => {
    return SHIFT_GROUPS.map(group => ({
      group,
      staffs: staffs.filter(staff => (staff.group || '白班') === group)
    }));
  }, [staffs]);

  return (
    <div className="min-h-screen text-slate-900 p-4 font-sans overflow-x-hidden relative" style={{ backgroundColor: pageBackgroundColor }}>
      <style>{`
        @keyframes pulse-once { 0% { transform: translateY(-10px); opacity: 0; } 100% { transform: translateY(0); opacity: 1; } }
        @keyframes fade-in-down { 0% { opacity: 0; transform: translateY(-5px); } 100% { opacity: 1; transform: translateY(0); } }
        .animate-pulse-once { animation: pulse-once 0.5s ease-out forwards; }
        .animate-fade-in-down { animation: fade-in-down 0.3s ease-out forwards; }
      `}</style>

      {showDraftPrompt && (
        <div className="max-w-[95vw] mx-auto mb-4 bg-amber-50 border border-amber-200 text-amber-800 p-4 rounded-xl flex items-center justify-between shadow-sm animate-fade-in-down">
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

      <div className="max-w-[95vw] mx-auto mb-6">
        <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-200 flex flex-col xl:flex-row xl:items-center justify-between gap-4">
          <div>
            <h1 className="text-2xl font-black text-slate-800 flex items-center gap-2">
              智能排班｜智慧排班開發版
              <span className="text-blue-500 text-sm font-normal px-2 py-1 bg-blue-50 rounded-lg border border-blue-100">PRO v1.6.0</span>
            </h1>
            <p className="text-slate-500 text-xs mt-1 italic">開發版開發使用</p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button onClick={() => saveToHistory('手動暫存')} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-2 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm">
              <Save size={16} /> 暫存
            </button>
            <button onClick={() => setShowHistoryModal(true)} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-2 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm">
              <Clock size={16} /> 歷史
            </button>

            <div className="relative">
              <button onClick={() => setShowExportMenu(!showExportMenu)} className="flex items-center gap-1.5 bg-slate-800 text-white px-3 py-2 rounded-xl font-bold hover:bg-slate-900 transition-all text-sm">
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
              <button onClick={() => handleAiAutoSchedule(false)} disabled={isAiLoading} className="flex items-center gap-2 bg-white text-blue-600 px-3 py-2 rounded-lg font-bold hover:bg-blue-50 transition-all disabled:opacity-50 text-xs">
                {isAiLoading ? <Loader2 className="animate-spin" size={14} /> : <Sparkles size={14} />} 全月補空
              </button>
              <button onClick={() => setShowAiControl(!showAiControl)} className={`flex items-center gap-2 px-3 py-2 rounded-lg font-bold transition-all text-xs ${showAiControl ? 'bg-blue-600 text-white shadow-inner' : 'text-slate-600 hover:bg-slate-200'}`}>
                <Calendar size={14} /> 指定補空
              </button>
              <button
                type="button"
                onClick={openSelectedCellFillModal}
                disabled={!selectedGridCell}
                className={`flex items-center gap-2 px-3 py-2 rounded-lg font-bold transition-all text-xs ${selectedGridCell ? 'text-slate-700 hover:bg-slate-200' : 'text-slate-400 cursor-not-allowed'}`}
              >
                <Check size={14} /> 補此格
              </button>
              <button
                type="button"
                onClick={clearSelectedCell}
                disabled={!selectedGridCell}
                className={`flex items-center gap-2 px-3 py-2 rounded-lg font-bold transition-all text-xs ${selectedGridCell ? 'text-red-600 hover:bg-red-50' : 'text-slate-400 cursor-not-allowed'}`}
              >
                <Trash2 size={14} /> 清除此格
              </button>
            </div>
          </div>
        </div>
        {aiFeedback && (
          <div className="mt-4 bg-indigo-50 border border-indigo-100 p-4 rounded-xl text-indigo-900 text-sm animate-pulse-once flex items-center gap-2">
            <Check size={16} className="text-green-600" />
            {aiFeedback}
          </div>
        )}
        {selectedGridCell && (
          <div className="mt-4 bg-blue-50 border border-blue-200 p-3 rounded-xl text-blue-900 text-sm flex items-center justify-between gap-3">
            <div>已選取儲存格：<span className="font-bold">{selectedGridCell.staff.name}</span>｜{selectedGridCell.dateStr}</div>
            <button
              type="button"
              onClick={() => setSelectedGridCell(null)}
              className="px-3 py-1.5 rounded-lg border border-blue-200 bg-white text-blue-700 hover:bg-blue-100 transition-colors text-xs font-bold"
            >
              取消選取
            </button>
          </div>
        )}
      </div>

      {showAiControl && (
        <div className="max-w-[95vw] mx-auto mb-6 bg-blue-50 border border-blue-200 p-6 rounded-2xl shadow-sm animate-fade-in-down">
          <h3 className="font-black text-blue-900 mb-4 flex items-center gap-2"><Sparkles size={18} /> 指定補空設定</h3>
          <div className="grid lg:grid-cols-4 gap-6">
            <div>
              <label className="block text-xs font-bold text-blue-700 mb-2 uppercase">1. 選擇人員（補空範圍）</label>
              <div className="flex flex-wrap gap-2">
                {staffs.map(s => (
                  <button
                    key={s.id}
                    onClick={() => {
                      const next = aiConfig.selectedStaffs.includes(s.id)
                        ? aiConfig.selectedStaffs.filter(id => id !== s.id)
                        : [...aiConfig.selectedStaffs, s.id];
                      setAiConfig({ ...aiConfig, selectedStaffs: next });
                    }}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold border transition-all ${aiConfig.selectedStaffs.includes(s.id) ? 'bg-blue-600 border-blue-600 text-white shadow-md' : 'bg-white border-blue-200 text-blue-600 hover:bg-blue-100'}`}
                  >
                    {s.name}（{s.group || '白班'}）
                  </button>
                ))}
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold text-blue-700 mb-2 uppercase">2. 日期範圍 ({aiConfig.dateRange.start} ~ {aiConfig.dateRange.end} 號)</label>
              <div className="flex items-center gap-2">
                <input
                  type="number"
                  min="1"
                  max="31"
                  value={aiConfig.dateRange.start}
                  onChange={(e) => setAiConfig({ ...aiConfig, dateRange: { ...aiConfig.dateRange, start: parseInt(e.target.value, 10) || 1 } })}
                  className="w-full border-blue-200 border p-2 rounded-lg text-sm text-center font-bold"
                />
                <span>至</span>
                <input
                  type="number"
                  min="1"
                  max="31"
                  value={aiConfig.dateRange.end}
                  onChange={(e) => setAiConfig({ ...aiConfig, dateRange: { ...aiConfig.dateRange, end: parseInt(e.target.value, 10) || 31 } })}
                  className="w-full border-blue-200 border p-2 rounded-lg text-sm text-center font-bold"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold text-blue-700 mb-2 uppercase">3. 指定班別（選填）</label>
              <select
                value={aiConfig.targetShift}
                onChange={(e) => setAiConfig({ ...aiConfig, targetShift: e.target.value })}
                className="w-full border-blue-200 border p-2 rounded-lg text-sm font-bold bg-white"
              >
                <option value="">依群組需求自動補空</option>
                {AI_MAIN_SHIFTS.map(s => <option key={s} value={s}>{s} 班</option>)}
              </select>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-bold text-blue-700 mb-2 uppercase">4. 範圍清除模式</label>
                <select
                  value={rangeClearMode}
                  onChange={(e) => setRangeClearMode(e.target.value)}
                  className="w-full border-blue-200 border p-2 rounded-lg text-sm font-bold bg-white"
                >
                  <option value="autoOnly">只清除自動補入內容</option>
                  <option value="all">清除範圍內全部內容</option>
                </select>
              </div>
              <div className="grid grid-cols-1 gap-3">
                <button
                  disabled={isAiLoading || aiConfig.selectedStaffs.length === 0}
                  onClick={() => handleAiAutoSchedule(true)}
                  className="w-full bg-blue-600 hover:bg-blue-700 text-white font-black py-2 rounded-xl transition-all shadow-lg active:scale-95 disabled:opacity-50 disabled:grayscale flex items-center justify-center gap-2"
                >
                  {isAiLoading ? <Loader2 className="animate-spin" size={18} /> : <Check size={18} />} 套用並補空
                </button>
                <button
                  type="button"
                  disabled={aiConfig.selectedStaffs.length === 0}
                  onClick={clearRangeCells}
                  className="w-full bg-white border border-red-200 text-red-600 hover:bg-red-50 font-black py-2 rounded-xl transition-all disabled:opacity-50 disabled:grayscale flex items-center justify-center gap-2"
                >
                  <Trash2 size={18} /> 範圍清除
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-[95vw] mx-auto mb-6 grid grid-cols-1 lg:grid-cols-12 gap-4">
        <div className="lg:col-span-4 bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex items-center gap-4">
          <input type="number" value={year} onChange={(e) => setYear(Number(e.target.value))} className="w-24 border rounded-lg p-2 text-center font-bold" />
          <select value={month} onChange={(e) => setMonth(Number(e.target.value))} className="w-20 border rounded-lg p-2 text-center font-bold">
            {[...Array(12).keys()].map(m => <option key={m + 1} value={m + 1}>{m + 1}月</option>)}
          </select>
        </div>

        <div className="lg:col-span-3 bg-blue-600 p-4 rounded-xl shadow-md text-white flex flex-col justify-center">
          <span className="text-xs opacity-80 uppercase tracking-wider">本月應休天數</span>
          <span className="text-2xl font-black">{requiredLeaves} <small className="text-sm font-normal">DAYS</small></span>
        </div>

        <div className="lg:col-span-5 flex items-center justify-end gap-2">
          <button onClick={() => changeScreen('entry')} className="bg-white border px-4 py-3 rounded-xl hover:bg-slate-50 transition-colors text-sm font-bold text-slate-700 flex items-center gap-2">
            <ArrowLeft size={18} className="text-slate-600" /> 回入口頁
          </button>
          <button onClick={() => changeScreen('settings')} className="bg-white border px-4 py-3 rounded-xl hover:bg-slate-50 transition-colors text-sm font-bold text-slate-700 flex items-center gap-2">
            <Settings size={18} className="text-slate-600" /> 系統設定
          </button>
        </div>
      </div>

      <div className="max-w-[95vw] mx-auto bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto pb-4">
          <table className="w-max min-w-full border-collapse">
            <thead>
              <tr className="bg-slate-100 border-b-2 border-slate-200">
                <th className={`sticky left-0 z-30 border-r font-black ${shiftColumnFontSizeClass}`} style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor, color: shiftColumnFontColor }}>班別</th>
                <th className={`sticky z-30 border-r font-black ${nameDateColumnFontSizeClass}`} style={{ left: densityConfig.shiftWidth, width: densityConfig.nameWidth, minWidth: densityConfig.nameWidth, backgroundColor: nameDateColumnBgColor, color: nameDateColumnFontColor }}>姓名/日期</th>
                {daysInMonth.map(d => (
                  <th
                    key={d.day}
                    className={`${densityConfig.dayHeaderClass} border-r text-center`}
                    style={{ minWidth: densityConfig.dayMinWidth, backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent') }}
                  >
                    <div className={`${tableFontSizeClass} opacity-60 uppercase`} style={{ color: tableFontColor }}>{d.weekStr}</div>
                    <div className={`${tableFontSizeClass} font-black`} style={{ color: tableFontColor }}>{d.day}</div>
                  </th>
                ))}
                {showRightStats && (
                <>
                <th className={`${densityConfig.statHeaderClass} border-r min-w-[60px] bg-blue-50 text-blue-700 font-bold`}>上班</th>
                <th className={`${densityConfig.statHeaderClass} border-r min-w-[60px] bg-green-50 text-green-700 font-bold`}>假日休</th>
                <th className={`${densityConfig.statHeaderClass} border-r min-w-[60px] bg-red-50 text-red-700 font-bold`}>總休</th>
                </>
                )}
                {showLeaveStats && mergedLeaveCodes.map(l => (
                  <th key={l} className={`${densityConfig.leaveHeaderClass} border-r min-w-[40px] bg-slate-50 text-[10px] uppercase text-slate-500 font-bold`}>{l}</th>
                ))}
                {(customColumns || []).map(col => (
                  <th key={col} className={`${densityConfig.leaveHeaderClass} border-r min-w-[70px] bg-violet-50 text-[10px] uppercase text-violet-600 font-bold`}>{col}</th>
                ))}
              </tr>
            </thead>

            <tbody>
              {groupedStaffs.map(({ group, staffs: groupStaffList }) => (
                <React.Fragment key={group}>
                  {groupStaffList.map((staff, index) => {
                    const stats = getStaffStats(staff.id);
                    const groupCount = groupStaffList.length + 1;
                    const groupIndex = groupStaffList.findIndex(s => s.id === staff.id);

                    return (
                      <tr key={staff.id} className="border-b border-slate-100 hover:bg-slate-50/50 transition-colors">
                        {index === 0 && (
                          <td rowSpan={groupCount} className="sticky left-0 z-20 border-r text-center shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)]" style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor }}>
                            <div className="flex items-center justify-center h-full" style={{ minHeight: densityConfig.rowMinHeight }}>
                              {showShiftLabels && (<span className={`font-black leading-tight tracking-[0.14em] [writing-mode:vertical-rl]`} style={{ color: shiftColumnFontColor, fontSize: shiftCellLabelFontSize }}>
                                {group}
                              </span>)}
                            </div>
                          </td>
                        )}

                        <td className={`sticky z-30 border-r shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)] ${densityConfig.nameCellPaddingClass}`} style={{ left: densityConfig.shiftWidth, width: densityConfig.nameWidth, minWidth: densityConfig.nameWidth, backgroundColor: nameDateColumnBgColor }}>
                          <div className="flex items-center gap-2">
                            <div className="flex flex-col items-center justify-center shrink-0 w-5">
                              <button
                                onClick={() => moveStaffInGroup(staff.id, 'up')}
                                disabled={groupIndex === 0}
                                className="text-slate-400 hover:text-blue-500 disabled:opacity-10 leading-none"
                              >
                                <ArrowUp size={14} />
                              </button>
                              <button
                                onClick={() => moveStaffInGroup(staff.id, 'down')}
                                disabled={groupIndex === groupStaffList.length - 1}
                                className="text-slate-400 hover:text-blue-500 disabled:opacity-10 leading-none"
                              >
                                <ArrowDown size={14} />
                              </button>
                            </div>

                            <input
                              type="text"
                              value={staff.name}
                              onChange={(e) => {
                                const next = [...staffs];
                                const currentIndex = next.findIndex(s => s.id === staff.id);
                                if (currentIndex !== -1) next[currentIndex].name = e.target.value;
                                setStaffs(next);
                              }}
                              className={`flex-1 min-w-0 text-center py-1.5 font-bold border-none rounded-lg focus:ring-2 focus:ring-blue-400 bg-transparent ${nameDateColumnFontSizeClass}`} style={{ color: nameDateColumnFontColor }}
                            />

                            <button
                              onClick={() => removeStaff(staff.id)}
                              className="text-slate-400 hover:text-red-500 shrink-0 w-6 flex items-center justify-center"
                            >
                              <Minus size={14} />
                            </button>
                          </div>
                        </td>

                        {daysInMonth.map(d => {
                          const cellData = schedule[staff.id]?.[d.date];
                          const val = typeof cellData === 'object' && cellData !== null ? (cellData?.value || '') : (cellData || '');
                          return (
                            <td
                              key={d.date}
                              className={`border-r p-0 relative overflow-hidden ${selectedGridCell?.staff?.id === staff.id && selectedGridCell?.dateStr === d.date ? 'ring-2 ring-blue-500 ring-inset' : ''}`}
                              style={{ backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent'), opacity: d.isHoliday || d.isWeekend ? 0.9 : 1 }}
                            onClick={() => { if (selectionMode === 'cell') setSelectedGridCell({ staff, dateStr: d.date }); }}
                            >
                              <div className="relative">
                                <select
                                  value={val}
                                  onChange={(e) => handleCellChange(staff.id, d.date, e.target.value)}
                                  className={`w-full ${densityConfig.cellHeightClass} text-center bg-transparent border-none cursor-pointer font-bold appearance-none hover:bg-black/5 ${tableFontSizeClass}`} style={{ color: tableFontColor }}
                                >
                                  <option value=""></option>
                                  <optgroup label="上班">
                                    {DICT.SHIFTS.map(s => <option key={s} value={s}>{s}</option>)}
                                  </optgroup>
                                  <optgroup label="休假">
                                    {mergedLeaveCodes.map(l => <option key={l} value={l}>{l}</option>)}
                                  </optgroup>
                                </select>

                                {showBlueDots && selectionMode === 'dot' && (
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setSelectedGridCell({ staff, dateStr: d.date });
                                  }}
                                  className="absolute right-1 top-1/2 -translate-y-1/2 z-0 w-3.5 h-3.5 flex items-center justify-center"
                                  aria-label={`選取 ${staff.name} ${d.date} 儲存格`}
                                  title={`選取 ${staff.name} ${d.date} 儲存格`}
                                >
                                  <span className={`${densityConfig.selectorDotClass} rounded-full transition-all ${selectedGridCell?.staff?.id === staff.id && selectedGridCell?.dateStr === d.date ? 'bg-blue-700 scale-110' : 'bg-blue-300/90 hover:bg-blue-500'}`}></span>
                                </button>
                                )}
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

                  <tr className="border-b border-slate-200 bg-slate-50/70">
                    <td className={`sticky z-30 border-r shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)] ${densityConfig.nameCellPaddingClass}`} style={{ left: densityConfig.shiftWidth, width: densityConfig.nameWidth, minWidth: densityConfig.nameWidth, backgroundColor: nameDateColumnBgColor }}>
                      <div className="flex items-center justify-center">
                        <button
                          onClick={() => addStaff(group)}
                          className="bg-blue-600 text-white px-4 py-2 rounded-xl hover:bg-blue-700 transition-colors flex items-center gap-2 font-bold text-sm"
                        >
                          <Plus size={16} /> 新增人員
                        </button>
                      </div>
                    </td>

                    {daysInMonth.map(d => (
                      <td
                        key={d.date}
                        className="border-r"
                        style={{ backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent'), opacity: d.isHoliday || d.isWeekend ? 0.9 : 1 }}
                      ></td>
                    ))}

                    <td colSpan={(showRightStats ? 3 : 0) + (showLeaveStats ? mergedLeaveCodes.length : 0) + (customColumns?.length || 0)}></td>
                  </tr>
                </React.Fragment>
              ))}
            </tbody>

            {showBottomStats && (
            <tfoot className="bg-slate-100 border-t-2 border-slate-200">
              {['D', 'E', 'N', 'totalLeave'].map((rowKey) => (
                <tr key={rowKey}>
                  <td className={`sticky left-0 z-10 border-r ${densityConfig.footCellPaddingClass}`} style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor }}></td>
                  <td className={`sticky z-10 border-r text-right font-bold ${nameDateColumnFontSizeClass} ${densityConfig.footCellPaddingClass}`} style={{ left: densityConfig.shiftWidth, width: densityConfig.nameWidth, minWidth: densityConfig.nameWidth, backgroundColor: nameDateColumnBgColor, color: nameDateColumnFontColor }}>
                    {rowKey === 'totalLeave' ? '當日休假' : `${rowKey} 班人數`}
                  </td>
                  {daysInMonth.map(d => {
                    const count = getDailyStats(d.date)[rowKey];
                    return (
                      <td key={d.date} className={`border-r p-2 text-center font-black ${tableFontSizeClass}`} style={{ color: tableFontColor }}>
                        {count || ''}
                      </td>
                    );
                  })}
                  <td colSpan={3 + mergedLeaveCodes.length + (customColumns?.length || 0)}></td>
                </tr>
              ))}
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
            <div className={`p-2 rounded-xl ${iconBg}`}><Icon className={`w-5 h-5 ${iconColor}`} /></div>
            <div><h3 className="font-bold text-gray-800">{title}</h3><p className="text-sm text-gray-500 mt-1 leading-relaxed">{desc}</p></div>
          </div>
        </div>
        <div className="px-6 py-6">{children}</div>
      </div>
    </div>
  );
}

function SettingsView({ changeScreen, colors, setColors, customHolidays, setCustomHolidays, specialWorkdays, setSpecialWorkdays, medicalCalendarAdjustments, setMedicalCalendarAdjustments, staffingConfig, setStaffingConfig, uiSettings, setUiSettings, customLeaveCodes, setCustomLeaveCodes, customColumns, setCustomColumns }) {
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
  const fixedGroups = [
    { icon: ShieldCheck, title: '排班核心規則', items: ['每日至少一位主管級別值班', '夜班後必須接續至少 24 小時休息', '特定危險單位需雙人以上配置'] },
    { icon: Clock, title: '班別接續限制', items: ['禁止「花班」：白班不得直接接續大夜', '小夜轉白班間隔需大於 12 小時', '連班上限不得超過 6 天'] },
    { icon: UserCheck, title: '請假與休假限制', items: ['法定假日優先排休計算', '年度特休不得低於勞基法標準', '病假/喪假排位權重調整限制'] },
    { icon: Database, title: '核心資料格式', items: ['員工編號固定為 8 位數字', '班表週期固定為 1 個自然月', '資料匯出格式限定為 .xlsx / .pdf'] }
  ];
  const extGroups = [
    { title: '權限控管', desc: '各級主管登入權限', icon: ShieldCheck },
    { title: '雲端同步', desc: '即時異地備份機制', icon: Database },
    { title: '單位自訂規則', desc: '細化至科別的規則', icon: Plus },
    { title: '進階 AI 條件', desc: '複雜人員偏好學習', icon: Cpu }
  ];
  return (
    <div className="min-h-screen bg-gray-50 text-slate-800 font-sans selection:bg-blue-100">
      <header className="sticky top-0 z-50 bg-white border-b border-gray-200 px-8 py-4 flex justify-between items-center shadow-sm">
        <div><div className="flex items-center gap-2 mb-0.5"><Settings className="w-6 h-6 text-blue-600" /><h1 className="text-xl font-bold tracking-tight text-gray-900">系統設定</h1></div><p className="text-sm text-gray-500">可調整使用者設定，核心規則由系統固定管理</p></div>
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
                        className={`px-3 py-2 rounded-xl border text-sm font-medium transition ${uiSettings.themePreset === key ? 'bg-violet-600 border-violet-600 text-white' : 'bg-white border-violet-200 text-violet-700 hover:bg-violet-50'}`}
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
                    <button type="button" onClick={addCustomColumn} className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-md border border-gray-300 bg-white text-sm font-medium text-blue-600 hover:bg-blue-50">
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
              </div>
            </SettingRow>
            <SettingRow icon={UserCheck} title="人力需求設定" desc="獨立設定平日 / 假日各班需求，作為全月補空與指定補空的直接依據。" iconBg="bg-sky-50" iconColor="text-sky-600">
              <div className="space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div>
                    <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">醫院層級</label>
                    <select
                      value={staffingConfig.hospitalLevel}
                      onChange={(e) => setStaffingConfig(prev => ({ ...prev, hospitalLevel: e.target.value }))}
                      className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-gray-50"
                    >
                      <option value="medical">醫學中心</option>
                      <option value="regional">區域醫院</option>
                      <option value="local">地區醫院</option>
                    </select>
                  </div>
                  <div>
                    <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">總床數</label>
                    <input
                      type="number"
                      min="0"
                      value={staffingConfig.totalBeds}
                      onChange={(e) => setStaffingConfig(prev => ({ ...prev, totalBeds: parseInt(e.target.value, 10) || 0 }))}
                      className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-gray-50"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">護理師總數</label>
                    <input
                      type="number"
                      min="0"
                      value={staffingConfig.totalNurses}
                      onChange={(e) => setStaffingConfig(prev => ({ ...prev, totalNurses: parseInt(e.target.value, 10) || 0 }))}
                      className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-gray-50"
                    />
                  </div>
                </div>

                <div className="rounded-xl bg-sky-50 border border-sky-100 px-4 py-3 text-xs text-sky-700">
                  參考護病比：{HOSPITAL_LEVEL_LABELS[staffingConfig.hospitalLevel]}｜白班 {HOSPITAL_RATIO_HINTS[staffingConfig.hospitalLevel].white}、小夜 {HOSPITAL_RATIO_HINTS[staffingConfig.hospitalLevel].evening}、大夜 {HOSPITAL_RATIO_HINTS[staffingConfig.hospitalLevel].night}
                </div>

                <div className="rounded-xl border border-gray-200 bg-white px-4 py-3 text-sm text-gray-600">
                  <div className="font-semibold text-gray-800 mb-1">目前補空依據</div>
                  <div>平日：白班 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.weekday.white}</span> 人、小夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.weekday.evening}</span> 人、大夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.weekday.night}</span> 人</div>
                  <div>假日：白班 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.holiday.white}</span> 人、小夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.holiday.evening}</span> 人、大夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.holiday.night}</span> 人</div>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                  <div className="rounded-2xl border border-gray-200 p-4 bg-gray-50/50">
                    <h4 className="font-bold text-gray-800 mb-4">平日需求</h4>
                    <div className="grid grid-cols-3 gap-3">
                      <div>
                        <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">白班</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.white}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, white: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">小夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.evening}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, evening: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">大夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.night}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, night: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                    </div>
                  </div>

                  <div className="rounded-2xl border border-gray-200 p-4 bg-gray-50/50">
                    <h4 className="font-bold text-gray-800 mb-4">假日需求</h4>
                    <div className="grid grid-cols-3 gap-3">
                      <div>
                        <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">白班</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.holiday.white}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, holiday: { ...prev.requiredStaffing.holiday, white: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">小夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.holiday.evening}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, holiday: { ...prev.requiredStaffing.holiday, evening: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">大夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.holiday.night}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, holiday: { ...prev.requiredStaffing.holiday, night: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </SettingRow>
            <SettingRow icon={Calendar} title="假期新增" desc="使用西曆年月日新增自訂假期，並可個別刪除。">
              <div className="space-y-5"><div><label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">西曆年月日</label><div className="grid grid-cols-1 md:grid-cols-3 gap-3"><input type="number" placeholder="年" value={holidayInput.year} onChange={(e)=>setHolidayInput({ ...holidayInput, year: e.target.value })} className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-gray-50 focus:outline-none focus:ring-2 focus:ring-blue-100" /><input type="number" placeholder="月" value={holidayInput.month} onChange={(e)=>setHolidayInput({ ...holidayInput, month: e.target.value })} className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-gray-50 focus:outline-none focus:ring-2 focus:ring-blue-100" /><input type="number" placeholder="日" value={holidayInput.day} onChange={(e)=>setHolidayInput({ ...holidayInput, day: e.target.value })} className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-gray-50 focus:outline-none focus:ring-2 focus:ring-blue-100" /></div></div><button onClick={addCustomHoliday} className="w-full md:w-auto flex items-center justify-center gap-2 px-5 py-2.5 text-sm font-semibold text-white bg-blue-600 rounded-xl hover:bg-blue-700"><Plus className="w-4 h-4" /> 新增假期</button><div className="pt-3 border-t border-gray-100"><label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">已新增假期</label><div className="space-y-2 max-h-52 overflow-y-auto pr-1">{customHolidays.length === 0 ? <div className="text-xs text-gray-400 p-4 bg-gray-50 border border-dashed border-gray-300 rounded-xl text-center">尚未新增自訂假期</div> : customHolidays.map(dateStr => <div key={dateStr} className="flex items-center justify-between p-3 bg-gray-50 border border-gray-200 rounded-xl"><span className="text-sm text-gray-700 font-medium">{dateStr}</span><button onClick={() => removeCustomHoliday(dateStr)} className="w-8 h-8 flex items-center justify-center rounded-full border border-red-200 text-red-500 hover:bg-red-50 font-bold">-</button></div>)}</div></div></div>
            </SettingRow>
            <SettingRow icon={Grid} title="補空優先" desc="設定補空排序偏好，作為後續智慧補班的參考方向。" iconBg="bg-teal-50" iconColor="text-teal-600">
              <div className="space-y-3">{[{ label: '優先補缺額最多的班別', active: true }, { label: '優先達成個人班數平均', active: true }, { label: '優先減少跨天連班', active: false }, { label: '優先分配假日班補空', active: true }].map((item, idx) => <div key={idx} className="flex items-center gap-3 p-3 hover:bg-gray-50 rounded-xl cursor-pointer"><div className={`w-4 h-4 rounded border flex items-center justify-center shrink-0 ${item.active ? 'bg-blue-600 border-blue-600' : 'bg-white border-gray-300'}`}>{item.active && <div className="w-1.5 h-1.5 bg-white rounded-full"></div>}</div><span className="text-sm text-gray-700">{item.label}</span></div>)}<div className="pt-3 border-t border-gray-100"><button className="text-sm text-blue-600 font-medium flex items-center gap-1 hover:underline"><Plus className="w-3.5 h-3.5" /> 新增自訂偏好</button></div></div>
            </SettingRow>
          </div>
        </section>
        <section className="space-y-6">
          <div className="flex items-center justify-between"><div className="flex items-center gap-2"><div className="w-1 h-6 bg-amber-500 rounded-full"></div><h2 className="text-lg font-bold text-gray-800">系統固定規則</h2></div><div className="flex items-center gap-1.5 px-3 py-1 bg-amber-50 text-amber-700 rounded-full border border-amber-100"><Lock className="w-3.5 h-3.5" /><span className="text-xs font-bold uppercase tracking-tight">唯讀模式 / 使用者不可修改</span></div></div>
          <div className="bg-slate-50 border border-slate-200 rounded-3xl overflow-hidden shadow-inner"><div className="grid grid-cols-1 md:grid-cols-2 gap-px bg-slate-200">{fixedGroups.map((group, i) => <div key={i} className="bg-white p-6 hover:bg-slate-50 transition-colors"><div className="flex items-center gap-3 mb-4"><group.icon className="w-5 h-5 text-slate-400" /><h4 className="font-bold text-slate-700">{group.title}</h4></div><ul className="space-y-3">{group.items.map((item, j) => <li key={j} className="flex items-start gap-3"><CheckCircle2 className="w-4 h-4 text-green-500 mt-0.5 shrink-0" /><span className="text-sm text-slate-600 leading-relaxed">{item}</span></li>)}</ul></div>)}<div className="md:col-span-2 bg-slate-900 text-white p-8"><div className="flex items-center gap-3 mb-6"><div className="p-2 bg-blue-500/20 rounded-xl border border-blue-400/30"><Cpu className="w-5 h-5 text-blue-400" /></div><div><h4 className="font-bold text-lg">AI 智能補班基礎限制 (核心邏輯)</h4><p className="text-xs text-slate-400 uppercase tracking-widest mt-0.5">Core Intelligent Logic Constraints</p></div></div><div className="grid grid-cols-1 md:grid-cols-3 gap-8"><div className="space-y-2"><div className="text-blue-400 text-xs font-bold">權重演算限制</div><p className="text-sm text-slate-300">AI 禁止將權重設為負值，所有補空邏輯皆需符合人員適任性評分，系統會鎖定基底適任性分數不被手動覆蓋。</p></div><div className="space-y-2"><div className="text-blue-400 text-xs font-bold">覆蓋衝突限制</div><p className="text-sm text-slate-300">手動排班具有 100% 絕對優先權。AI 補空演算時自動迴避「手動鎖定」儲存格，且不主動修改現有固定排程。</p></div><div className="space-y-2"><div className="text-blue-400 text-xs font-bold">合規性校準</div><p className="text-sm text-slate-300">補空結果產出前需通過內部「三重合規檢驗」，若違反勞基法工時限制，AI 將拒絕輸出該時段排程。</p></div></div></div></div></div>
        </section>
        <section className="space-y-6"><div className="flex items-center gap-2"><div className="w-1 h-6 bg-purple-600 rounded-full"></div><h2 className="text-lg font-bold text-gray-800">未來擴充功能</h2></div><div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">{extGroups.map((ext, idx) => <div key={idx} className="group p-5 bg-white border border-gray-100 rounded-2xl shadow-sm hover:border-purple-200 transition-all cursor-not-allowed opacity-75"><div className="mb-4 p-2 w-fit bg-purple-50 rounded-xl group-hover:bg-purple-100 transition-colors"><ext.icon className="w-5 h-5 text-purple-600" /></div><h4 className="font-bold text-gray-800 mb-1">{ext.title}</h4><p className="text-xs text-gray-500 mb-4">{ext.desc}</p><div className="flex items-center text-[10px] font-bold text-purple-400 uppercase tracking-tighter"><span className="bg-purple-50 px-2 py-0.5 rounded">Coming Soon</span></div></div>)}</div></section>
      </main>
      <footer className="max-w-7xl mx-auto px-8 py-12 border-t border-gray-200 flex flex-col md:flex-row justify-between items-center gap-4"><div className="flex items-center gap-4"><div className="flex items-center gap-2"><div className="w-3 h-3 bg-green-500 rounded-full animate-pulse"></div><span className="text-sm font-semibold text-gray-700">系統核心版本: v2.5.0-PRO</span></div><span className="text-gray-300">|</span><span className="text-sm text-gray-500">智慧排班引擎開發測試版</span></div><div className="text-sm text-gray-400">© 2024 Intelligent Scheduling System PRO. All rights reserved.</div></footer>
    </div>
  );
}

function EntryView({ changeScreen, goToLatestHistory }) {
  const [account, setAccount] = useState('');
  const [password, setPassword] = useState('');
  const handleLogin = (e) => { e.preventDefault(); changeScreen('schedule'); };
  return (
    <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 font-sans text-slate-800">
      <div className="mb-10 text-center"><div className="flex items-center justify-center gap-3 mb-4"><div className="bg-blue-600 p-2.5 rounded-xl shadow-md shadow-blue-200/50"><CalendarDays className="w-8 h-8 text-white" /></div><div className="text-left"><h1 className="text-2xl font-bold tracking-tight text-slate-900 flex items-center">智能排班｜智慧排班開發版<span className="ml-3 inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-100 text-blue-800 border border-blue-200">PRO v1.6.0</span></h1><p className="text-slate-500 text-xs font-medium tracking-wide mt-1">開發版開發使用</p></div></div></div>
      <div className="w-full max-w-[440px] bg-white rounded-2xl shadow-[0_12px_40px_rgba(0,0,0,0.03)] border border-slate-200/60 overflow-hidden"><div className="p-10"><div className="mb-10"><h2 className="text-xl font-bold text-slate-900 mb-2">歡迎使用智能排班系統</h2><p className="text-sm text-slate-500 leading-relaxed">您可以直接進入排班系統、開啟最近編輯過的班表，或前往調整系統設定參數。</p></div><form onSubmit={handleLogin} className="space-y-6 mb-10"><div className="space-y-2"><label className="block text-sm font-semibold text-slate-700">帳號</label><div className="relative group"><div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none"><User className="h-4.5 w-4.5 text-slate-400" /></div><input type="text" className="block w-full pl-11 pr-4 py-3 border border-slate-200 rounded-xl bg-slate-50/50 outline-none placeholder:text-slate-400" placeholder="請輸入您的管理帳號" value={account} onChange={(e)=>setAccount(e.target.value)} /></div></div><div className="space-y-2"><label className="block text-sm font-semibold text-slate-700">密碼</label><div className="relative group"><div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none"><Lock className="h-4.5 w-4.5 text-slate-400" /></div><input type="password" className="block w-full pl-11 pr-4 py-3 border border-slate-200 rounded-xl bg-slate-50/50 outline-none placeholder:text-slate-400" placeholder="••••••••" value={password} onChange={(e)=>setPassword(e.target.value)} /></div></div><div className="bg-slate-50 rounded-xl p-4 border border-slate-100 flex items-start gap-3"><Info className="w-5 h-5 text-blue-500 shrink-0 mt-0.5" /><p className="text-xs text-slate-500 leading-normal">目前為開發版入口頁，登入功能尚未正式啟用。點擊下方按鈕即可直接進入系統環境。</p></div></form><div className="space-y-4"><button type="button" onClick={() => changeScreen('schedule')} className="w-full flex justify-center items-center gap-2 py-3.5 px-4 rounded-xl text-sm font-bold text-white bg-blue-600 hover:bg-blue-700"><ShieldCheck className="w-4.5 h-4.5" />進入排班系統</button><div className="grid grid-cols-2 gap-4"><button type="button" onClick={goToLatestHistory} className="flex justify-center items-center gap-2 py-3 px-4 border border-slate-200 rounded-xl shadow-sm text-sm font-bold text-slate-700 bg-white hover:bg-slate-50"><Clock className="w-4 h-4 text-slate-500" />最近班表</button><button type="button" onClick={() => changeScreen('settings')} className="flex justify-center items-center gap-2 py-3 px-4 border border-slate-200 rounded-xl shadow-sm text-sm font-bold text-slate-700 bg-white hover:bg-slate-50"><Settings className="w-4 h-4 text-slate-500" />系統設定</button></div></div></div><div className="bg-slate-50/80 px-10 py-5 border-t border-slate-100 flex justify-between items-center"><button className="text-xs text-slate-500 hover:text-blue-600 font-bold flex items-center gap-1 transition-colors group">開發版說明<ChevronRight className="w-3.5 h-3.5 group-hover:translate-x-0.5 transition-transform" /></button><button className="text-xs text-slate-500 hover:text-blue-600 font-bold transition-colors">查看版本資訊</button></div></div><div className="mt-12 text-center"><p className="text-xs text-slate-400 font-medium tracking-[0.1em] uppercase">&copy; {new Date().getFullYear()} 智能排班系統 PRO. ALL RIGHTS RESERVED.</p></div>
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
    shiftColumnFontSize: 'large',
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
    cellHeightMode: 'standard'
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
  const [loadLatestOnEnter, setLoadLatestOnEnter] = useState(false);

  const goToSchedule = () => {
    setLoadLatestOnEnter(false);
    setScreen('schedule');
  };

  const goToLatestHistory = () => {
    setLoadLatestOnEnter(true);
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
        loadLatestOnEnter={loadLatestOnEnter}
        onLatestLoaded={() => setLoadLatestOnEnter(false)}
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
    />
  );
}
