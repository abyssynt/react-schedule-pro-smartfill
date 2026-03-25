
import React, { useState, useMemo, useEffect } from 'react';
進口 {
  加號、減號、設定、閃光燈、載入器2
  上箭頭、下箭頭、保存、歷史記錄（時鐘）、下載
  文件電子表格、文件文字、X、複選框、日曆、日曆天數、
  使用者、鎖定、資訊、佈局、盾牌檢查、網格、使用者檢查
  資料庫、CPU、顯示器、左箭頭、右箭頭、複選框2、垃圾桶2
來自 'lucide-react'；

// ==========================================
// 1. 系統程式碼字典
// ==========================================
const DICT = {
  班次: ['D', 'E', 'N', '白8-8', '夜8-8', '8-12', '12-16'],
  LEAVES: ['休', '例', '休', '特', '補', '國', '喪', '婚', '產', '病', '事', '陪產', 'AM', 'PM']
};

const SMART_RULES = {
  最大連續工作日：5，
  allowCrossGroupAssignment: false,
  不允許的下一班映射：{
    N: ['D', 'E'],
    結尾']，
    '白8-8': ['D', 'N'],
    '夜8-8': ['E', 'N']
  },
  blockedLeavePrefixes: ['off', '例', '休', '特', '補', '國', '喪', '婚', '產', '病', '事', '陪產', 'AM', 'PM'],
  懷孕限制班次：['N', '夜8-8'],
  填充優先權權重：{
    sameShiftCount: 3,
    總班次數：2，
    sameGroup: 1
  }
};

const SHIFT_GROUPS = ['白班', '小夜', '大夜'];

const ANNOUNCED_CALENDAR_OVERRIDES = {
  2024年：{
    holidays: ['2024-01-01', '2024-02-08', '2024-02-09', '2024-02-10', '2024-02-11', '2024-02-12', '2024-02-11', '2024-02-12', '2024-02-11', '2024-02-12', '2024-02-13' '2024-02-28', '2024-04-04', '2024-04-05', '2024-06-10', '2024-09-17', '2024-10-10'],
    工作日：[]
  },
  2025年：{
    holidays: ['2025-01-01', '2025-01-27', '2025-01-28', '2025-01-29', '2025-01-30', '2025-01-31', '2025-01-30', '2025-01-31', '2025-01-30', '2025-01-31', '2025-01-283' '2025-04-04', '2025-05-31', '2025-10-06', '2025-10-10'],
    工作日：[]
  },
  2026年：{
    holidays: ['2026-01-01', '2026-02-16', '2026-02-17', '2026-02-18', '2026-02-19', '2026-02-20', '2026-02-19', '2026-02-20', '2026-02-19', '2026-02-20', '2026-026,0200 '2026-04-05', '2026-06-19', '2026-09-25', '2026-10-10'],
    工作日：[]
  }
};

const CHINESE_MONTH_MAP = {
  '正月': 1, '一月': 1, '二月': 2, '三月': 3, '四月': 4, '五月': 5, '六月': 6,
  '七月': 7, '八月': 8, '九月': 9, '十月': 10, '十一月': 11, '十二月': 12,
  「臘月」：12
};

const CHINESE_DAY_MAP = {
  '初一': 1, '初二': 2, '初三': 3, '初四': 4, '初五': 5, '初六': 6, '初七': 7, '初八': 8, '初九': 9, '初十': 10,
  '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15, '十六': 16, '十七': 17, '十八': 18, '十九': 19, '二十': 20,
  '二十一': 21, '二十二': 22, '二十三': 23, '二十四': 24, '二十五': 25, '二十六': 26, '二十七': 27, '二十八': 28, '九二十': 29, '二十七': 27, '二十八': 28, '九二十': 29, '二十七': 30
};

const normalizeLunarMonth = (label = '') => {
  const cleaned = String(label).trim().replace(/^閏/, '').replace(/月$/, '');
  if (CHINESE_MONTH_MAP[String(label).trim().replace(/^閏/, '')]) {
    return CHINESE_MONTH_MAP[String(label).trim().replace(/^閏/, '')];
  }
  如果 (/^\d+$/.test(cleaned)) 回傳 Number(cleaned);
  如果 (CHINESE_MONTH_MAP[`${cleaned}月`]) 回傳 CHINESE_MONTH_MAP[`${cleaned}月`];
  返回空值；
};

const normalizeLunarDay = (label = '') => {
  const cleaned = String(label).trim();
  如果 (CHINESE_DAY_MAP[cleaned]) 回傳 CHINESE_DAY_MAP[cleaned];
  如果 (/^\d+$/.test(cleaned)) 回傳 Number(cleaned);
  返回空值；
};

const apiKey = "";

const formatDateKey = (date) => {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  返回 `${y}-${m}-${d}`；
};

const parseDateKey = (dateKey) => {
  const [y, m, d] = dateKey.split('-').map(Number);
  返回新的 Date(y, m - 1, d)；
};

const addDays = (date, amount) => {
  const next = new Date(date);
  next.setDate(next.getDate() + amount);
  返回下一個；
};

const isWeekendDate = (date) => date.getDay() === 0 || date.getDay() === 6;

const uniqueSortedDates = (dates = []) => Array.from(new Set(dates)).sort();

const getChineseCalendarInfo = (date) => {
  const formatter = new Intl.DateTimeFormat('zh-TW-u-ca-chinese', {
    年份：'數字'，
    月份：'長'，
    日期：'數字'
  });
  const parts = formatter.formatToParts(date);
  const yearPart = parts.find((part) => part.type === 'relatedYear');
  const monthPart = parts.find((part) => part.type === 'month');
  const dayPart = parts.find((part) => part.type === 'day');
  const rawMonth = String(monthPart?.value || '').trim();
  const rawDay = String(dayPart?.value || '').trim();

  返回 {
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
  回傳 info.monthNumber === lunarMonth && info.dayNumber === lunarDay && Boolean(info.leapMonth) === Boolean(options.leap);
};

const findGregorianDateByLunarInYear = (year, lunarMonth, lunarDay, options = {}) => {
  const start = new Date(year, 0, 1);
  const end = new Date(year, 11, 31);
  for (let cursor = new Date(start); cursor <= end; cursor = addDays(cursor, 1)) {
    如果 (matchesLunarDate(cursor, lunarMonth, lunarDay, options)) 傳回新的 Date(cursor)；
  }
  返回空值；
};

const getQingmingDate = (year) => {
  如果（年份 >= 2000 && 年份 <= 2099）{
    const y = year % 100;
    const day = Math.floor(y * 0.2422 + 4.81) - Math.floor((y - 1) / 4);
    返回新的 Date(year, 3, day)；
  }
  返回新的 Date(year, 3, 4)；
};

const getFixedSolarHolidayDates = (year) => {
  const fixed = [
    [1, 1], // 開國紀念日
    [2, 28], // 和平紀念日
    [4, 4], // 兒童節
    [5, 1], // 勞動節
    [9, 28], // 孔子誕辰紀念日 / 教師節
    [10, 10], // 國慶日
    [10, 25], // 台灣光復暨金門古寧頭大捷紀念日
    [12, 25], // 行憲紀念日
  ];
  return fixed.map(([month, day]) => formatDateKey(new Date(year, month - 1, day)));
};

const getRuleBasedHolidayDates = (year) => {
  const holidaySet = new Set(getFixedSolarHolidayDates(year));

  const qingming = getQingmingDate(year);
  holidaySet.add(formatDateKey(qingming));

  const dragonBoat = findGregorianDateByLunarInYear(year, 5, 5);
  如果 (dragonBoat) holidaySet.add(formatDateKey(dragonBoat));

  const midAutumn = findGregorianDateByLunarInYear(year, 8, 15);
  如果 (midAutumn) holidaySet.add(formatDateKey(midAutumn));

  const lunarNewYearDay = findGregorianDateByLunarInYear(year, 1, 1);
  如果（農曆新年）{
    const springHoliday = [
      addDays(lunarNewYearDay, -2),
      addDays(lunarNewYearDay, -1),
      農曆新年
      addDays(lunarNewYearDay, 1),
      addDays(lunarNewYearDay, 2)
    ];
    春假
      .filter((date) => date.getFullYear() === year)
      .forEach((date) => holidaySet.add(formatDateKey(date)));
  }

  const childrensDayKey = formatDateKey(new Date(year, 3, 4));
  const qingmingKey = formatDateKey(qingming);
  如果 (childrensDayKey === qingmingKey) {
    const qingmingWeekday = qingming.getDay();
    const extraHoliday = qingmingWeekday === 4 ? addDays(qingming, 1) : addDays(qingming, -1);
    holidaySet.add(formatDateKey(extraHoliday));
  }

  返回 uniqueSortedDates([...holidaySet]);
};

const findNearestWorkday = (startDate, direction, occupiedHolidays, workdayOverrides) => {
  let cursor = addDays(startDate, direction);
  while (true) {
    const key = formatDateKey(cursor);
    const weekend = isWeekendDate(cursor);
    const isWorkday = (!weekend || workdayOverrides.has(key)) && !occupiedHolidays.has(key);
    如果（是工作日）返回鍵；
    cursor = addDays(cursor, direction);
  }
};

const applyCompensatoryHolidays = (holidayDates, workdayDates = []) => {
  const holidaySet = new Set(uniqueSortedDates(holidayDates));
  const workdaySet = new Set(uniqueSortedDates(workdayDates));

  const baseDates = [...holidaySet].sort();
  baseDates.forEach((dateKey) => {
    const date = parseDateKey(dateKey);
    如果 (date.getDay() === 6) {
      holidaySet.add(findNearestWorkday(date, -1, holidaySet, workdaySet));
    } else if (date.getDay() === 0) {
      holidaySet.add(findNearestWorkday(date, 1, holidaySet, workdaySet));
    }
  });

  返回 uniqueSortedDates([...holidaySet]);
};

const getSystemHolidayCalendar = (year, options = {}) => {
  const {
    customHolidays = [],
    announceOverrides = ANNOUNCED_CALENDAR_OVERRIDES,
    specialWorkdays = [],
    unitAdjustments = { holidays: [], workdays: [] }
  } = 選項；

  const announce = announceOverrides[year];
  const overrideHolidays = announce?.holidays || [];
  const overrideWorkdays = announce?.workdays || [];

  const unitHolidayDates = (unitAdjustments.holidays || []).filter((date) => date.startsWith(`${year}-`));
  const unitWorkdayDates = (unitAdjustments.workdays || []).filter((date) => date.startsWith(`${year}-`));
  const customHolidayDates = customHolidays.filter((date) => date.startsWith(`${year}-`));
  const specialWorkdayDates = specialWorkdays.filter((date) => date.startsWith(`${year}-`));

  let holidayDates = overrideHolidays.length > 0 ? overrideHolidays : applyCompensatoryHolidays(getRuleBasedHolidayDates(year), [...overrideWorkdays, ...specialWorkdayDates, ...unitWorkdayDates]);
  let workdayDates = uniqueSortedDates([...overrideWorkdays, ...specialWorkdayDates, ...unitWorkdayDates]);

  holidayDates = uniqueSortedDates([...holidayDates, ...customHolidayDates, ...unitHolidayDates]).filter((date) => !workdayDates.includes(date));

  返回 {
    假期：假期日期，
    工作日：工作日期
  };
};
const STORAGE_KEY = 'schedule_app_history';

// 外部套件匯入：ExcelJS用於高品質Excel樣式輸出
const loadExcelJS = () => {
  傳回一個新的 Promise((resolve) => {
    如果 (window.ExcelJS) 返回 resolve(window.ExcelJS);
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js";
    script.onload = () => resolve(window.ExcelJS);
    document.head.appendChild(script);
  });
};

const normalizeStaffGroup = (staffList = []) => {
  如果 (!Array.isArray(staffList) || staffList.length === 0) 回傳 [];

  const fallbackGroups = [
    '白班', '白班', '白班', '白班', '白班',
    '小夜', '小夜', '小夜', '小夜', '小夜',
    '大夜', '大夜', '大夜', '大夜', '大夜'
  ];

  回傳 staffList.map((staff, index) => ({
    ....職員，
    懷孕：布爾值（員工.懷孕），
    group: SHIFT_GROUPS.includes(staff.group) ? staff.group : (fallbackGroups[index] || '白班')
  }));
};


const getCodePrefix = (rawCode = '') => {
  const code = String(rawCode || '').trim();
  如果 (!code) 回傳 '';
  如果 (代碼 === 'off') 返回 'off'；
  const direct = SMART_RULES.blockedLeavePrefixes.find((prefix) => code === prefix || code.startsWith(prefix));
  如果（直接）返回直接；
  返回碼；
};

const getShiftGroupByCode = (code = '') => {
  if (['D', '白8-8', '8-12', '12-16'].includes(code)) return '白班';
  if (['E', '夜8-8'].includes(code)) return '小夜';
  如果 (['N'].includes(code)) 回傳 '大夜';
  返回空值；
};

const isLeaveCode = (code = '') => SMART_RULES.blockedLeavePrefixes.includes(getCodePrefix(code));
const isShiftCode = (code = '') => DICT.SHIFTS.includes(code);

const GROUP_TO_DEMAND_KEY = {
  '白班': '白色',
  '小夜': '晚上',
  '大夜': '夜晚'
};

const DEFAULT_SHIFT_BY_GROUP = {
  '白班': 'D',
  '小夜': 'E',
  '大夜': 'N'
};

const AI_MAIN_SHIFTS = ['D', 'E', 'N'];

const HOSPITAL_LEVEL_LABELS = {
  醫療：“醫學中心”，
  區域： '區域醫院',
  local: '地區醫院'
};

const HOSPITAL_RATIO_HINTS = {
  醫療：{ 白色：'1:6'， 晚上：'1:9'， 夜間：'1:11' }，
  區域：{ 白色：'1:7'， 傍晚：'1:11'， 夜晚：'1:13' }，
  local: { white: '1:10', evening: '1:13', night: '1:15' }
};


const UI_FONT_SIZE_OPTIONS = {
  small: { label: '小', className: 'text-xs' },
  中：{標籤：'標準'，className：'text-sm'}，
  large: { label: '大', className: 'text-base' }
};

const getUiFontSizeClass = (sizeKey = 'medium') => UI_FONT_SIZE_OPTIONS[sizeKey]?.className || UI_FONT_SIZE_OPTIONS.medium.className;

function ScheduleView({ changeScreen, colors, setColors, customHolidays, setCustomHolidays, specialWorkdays, setSpecialWorkdays, medicalCalendarAdjustments, setMedicalCalendarAdjustments, stuiingConfigettSaffidjustments. loadLatestOnEnter, onLatestLoaded }) {
  // ==========================================
  // 2. 核心狀態定義
  // ==========================================
  const [year, setYear] = useState(2025);
  const [month, setMonth] = useState(3);
  const [isAiLoading, setIsAiLoading] = useState(false);

  const [staffs, setStaffs] = useState(normalizeStaffGroup([
    { id: 's1', name: '新成員' }, { id: 's2', name: '新成員' }, { id: 's3', name: '新成員' }, { id: 's4', name: '新成員' }, { id: 's5', name: '新成員' }, { id: 's5', name: '新成員' }, { id: 's5', name: '新成員' }, { id: 's5', name: '新成員' }, { id: 's5', name: '新成員' }, { id: 's5', name: '新成員' }, { id: 's5', name: '新成員' }, { id: 's5', name: '新成員' }
    { id: 's6', name: '新成員' }, { id: 's7', name: '新成員' }, { id: 's8', name: '新成員' }, { id: 's9', name: '新成員' }, { id: 's10', name: '新成員' }, { id: 's10', name: '新成員' }, { id: 's10', name: '新成員' }, { id: 's10', name: '新成員' },
    { id: 's11', name: '新成員' }, { id: 's12', name: '新成員' }, { id: 's13', name: '新成員' }, { id: 's14', name: '新成員' }, { id: 's15', name: '新成員' }, { id: 's15', name: '新成員'
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

  // AI指定排班設定
  const [aiConfig, setAiConfig] = useState({
    selectedStaffs: [],
    日期範圍：{開始：1，結束：31}，
    targetShift:''
  });

  const tableFontSizeClass = getUiFontSizeClass(uiSettings?.tableFontSize);
  const pageBackgroundColor = uiSettings?.pageBackgroundColor || '#f8fafc';
  const tableFontColor = uiSettings?.tableFontColor || '#1f2937';
  const shiftColumnBgColor = uiSettings?.shiftColumnBgColor || '#ffffff';
  const nameDateColumnBgColor = uiSettings?.nameDateColumnBgColor || '#ffffff';

  // ==========================================
  // 3.初始載入與自動帶入
  // ==========================================
  useEffect(() => {
    const stored = localStorage.getItem(STORAGE_KEY);
    如果（已儲存）{
      嘗試 {
        const parsed = JSON.parse(stored);
        如果（已解析且解析後的長度大於 0）{
          設定歷史記錄清單（已解析）；
          setShowDraftPrompt(true);
        }
      } catch (e) {
        console.error("歷史記錄解析失敗");
      }
    }
  }, []);

  useEffect(() => {
    如果 (!loadLatestOnEnter) 返回；

    const stored = localStorage.getItem(STORAGE_KEY);
    如果 (!stored) {
      onLatestLoaded?.();
      返回;
    }

    嘗試 {
      const parsed = JSON.parse(stored);
      如果（已解析且解析長度大於 0）{
        設定歷史記錄清單（已解析）；
        載入歷史記錄(parsed[0]);
      }
    } catch (e) {
      console.error("自動載入最新歷史記錄失敗");
    } 最後 {
      onLatestLoaded?.();
    }
  }, [loadLatestOnEnter]);

  const holidayCalendar = useMemo(() => {
    返回 getSystemHolidayCalendar(year, {
      自訂節日
      特殊工作日
      單位調整：醫療日曆調整
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
        日期：i，
        日期：dateStr，
        weekStr: weekNames[weekNum],
        isWeekend: rawWeekend && !isAdjustedWorkday,
        isHoliday: holidaySet.has(dateStr),
        調整後的工作日
      });
    }
    返回天數；
  }, [年，月，假日，工作日]);

  const requiredLeaves = useMemo(
    () => daysInMonth.filter(d => d.isWeekend || d.isHoliday).length,
    [每月天數]
  ）；

  // ==========================================
  // 4. Excel 匯出 (ExcelJS 實作)
  // ==========================================
  const exportToExcel = async () => {
    setAiFeedback("📊正在產生高品質Excel報表...");
    const ExcelJS = await loadExcelJS();
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`${year}年${month}月班表`);

    const headerRow = ['班別', '日期/姓名', ...daysInMonth.map(d => `${d.day}\n(${d.weekStr})`), '上班', '假日休', '總休息', ...DICT.LEAVES];
    const header = worksheet.addRow(headerRow);
    標題高度 = 30；

    header.eachCell((cell, colNumber) => {
      cell.font = { bold: true, size: 10 };
      cell.alignment = { vertical: 'middle', horizo​​ntal: 'center', wrapText: true };
      單元格邊框 = {
        上：{樣式：'thin'}，左：{樣式：'thin'}，
        底部：{樣式：'細線'}，右側：{樣式：'細線'}
      };
      如果 (colNumber > 2 && colNumber <= daysInMonth.length + 2) {
        const d = daysInMonth[列號 - 3];
        如果 (d.isHoliday) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCACA' } };
        否則如果 (d.isWeekend) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
      }
    });

    staffs.forEach(staff => {
      const stats = getStaffStats(staff.id);
      const rowData = [
        員工組
        員工姓名
        ...daysInMonth.map(d => {
          const cellData = schedule[staff.id]?.[d.date];
          return typeof cellData === 'object' ? (cellData?.value || '') : (cellData || '');
        }),
        stats.work、stats.holidayLeave、stats.totalLeave、
        ...DICT.LEAVES.map(l => stats.leaveDetails[l] || '')
      ];
      const row = worksheet.addRow(rowData);

      row.eachCell((cell, colNumber) => {
        cell.alignment = { vertical: 'middle', horizo​​ntal: 'center' };
        單元格邊框 = {
          上：{樣式：'thin'}，左：{樣式：'thin'}，
          底部：{樣式：'細線'}，右側：{樣式：'細線'}
        };
        如果 (colNumber > 2 && colNumber <= daysInMonth.length + 2) {
          cell.numFmt = '@';
          const d = daysInMonth[列號 - 3];
          如果 (d.isHoliday) 儲存格.填色 = { 類型: '圖案'，圖案: '實線'，fgColor: { argb: 'FFFFE4E4' } };
          否則如果 (d.isWeekend) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FDF4' } };
        }
      });
    });

    ['D', 'E', 'N', 'totalLeave'].forEach(rowKey => {
      const label = rowKey === 'totalLeave' ？ '當日休假' : `${rowKey} 班次`;
      const rowData = ['', label, ...daysInMonth.map(d => getDailyStats(d.date)[rowKey] || '')];
      const row = worksheet.addRow(rowData);
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
        單元格邊框 = {
          上：{樣式：'thin'}，左：{樣式：'thin'}，
          底部：{樣式：'細線'}，右側：{樣式：'細線'}
        };
      });
    });

    worksheet.getColumn(1).width = 10;
    worksheet.getColumn(2).width = 15;
    for (let i = 3; i <= daysInMonth.length + 2; i++) worksheet.getColumn(i).width = 5;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      類型：'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `排班表_${年}年${月}月.xlsx`;
    a.點擊();
    setShowExportMenu(false);
    setAiFeedback("✅ Excel匯出成功！");
  };

  const exportToWord = () => {
    const html = `
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          @page { size: landscape; margin: 1cm; }
          body { font-family: sans-serif; }
          表格 { 邊框折疊：折疊；寬度：100%；字體大小：9pt； }
          th, td { border: 1px solid #000; padding: 4px; text-align: center; }
          .holiday { background-color: ${colors.holiday}; }
          .weekend { background-color: ${colors.weekend}; }
        </style>
      </head>
      <body>
        <h2 style="text-align:center;">${year}年${month}月班表</h2>
        <table>
          <thead>
            <tr>
              <th>班別</th>
              <th>姓名</th>
              ${daysInMonth.map(d => `<th class="${d.isHoliday ? 'holiday' : (d.isWeekend ? 'weekend' : '')}">${d.day}<br/>(${d.weekStr})</th>`).join('')}<br/>(${d.weekStr})</th>`).join('')}}
              <th>上班</th>
              <th>總休</th>
            </tr>
          </thead>
          <tbody>
            ${staffs.map(staff => {
              const stats = getStaffStats(staff.id);
              回傳`
                <tr>
                  <td>${staff.group}</td>
                  <td>${staff.name}</td>
                  ${daysInMonth.map(d => {
                    const cellData = schedule[staff.id]?.[d.date];
                    return `<td>${typeof cellData === 'object' ? (cellData?.value || '') : (cellData || '')}</td>`;
                  }）。加入（''）}
                  <td>${stats.work}</td>
                  <td>${stats.totalLeave}</td>
                </tr>
              `;
            }）。加入（''）}
          </tbody>
        </table>
      </body>
    </html>`;

    const blob = new Blob([html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `打印班表_${年}年${月}月.doc`;
    a.點擊();
    setShowExportMenu(false);
  };

  // ==========================================
  // 5. AI指定排班功能
  // ==========================================
  const handleAiAutoSchedule = async (isPartial = false) => {
    setIsAiLoading(true);
    setAiFeedback(isPartial ? "🧩系統正在依指定範圍補空..." : "🧩系統正在依人力需求補全月空白...");

    嘗試 {
      const mergedSchedule = JSON.parse(JSON.stringify(schedule));
      const targetStaffIds = isPartial && aiConfig.selectedStaffs.length > 0
        ? new Set(aiConfig.selectedStaffs)
        : new Set(staffs.map(s => s.id));

      const targetDays = daysInMonth.filter(d => {
        如果 (!isPartial) 回傳 true；
        返回 d.day >= aiConfig.dateRange.start && d.day <= aiConfig.dateRange.end;
      });

      const normalizedTargetShift = AI_MAIN_SHIFTS.includes(aiConfig.targetShift) ? aiConfig.targetShift : '';
      const restrictedGroup = normalizedTargetShift ? getShiftGroupByCode(normalizedTargetShift) : null;
      const summary = { workFilled: 0, leaveFilled: 0, skipped: 0 };

      const getScheduleCode = (snapshot, staffId, dateStr) => {
        const cellData = snapshot[staffId]?.[dateStr];
        return typeof cellData === 'object' && cellData !== null ? (cellData.value || '') : (cellData || '');
      };

      const setScheduleCode = (snapshot, staffId, dateStr, value, source = 'auto') => {
        如果 (!snapshot[staffId]) snapshot[staffId] = {};
        snapshot[staffId][dateStr] = value ? { value, source } : null;
      };

      const getDemandType = (day) => (day.isWeekend || day.isHoliday) ? 'holiday' : 'weekday';
      const getDemandForGroup = (day, group) => {
        const bucket = getDemandType(day);
        const key = GROUP_TO_DEMAND_KEY[group];
        返回 Number(staffingConfig?.requiredStaffing?.[bucket]?.[key] || 0);
      };

      const getAssignedCountByGroup = (snapshot, dateStr, group) => {
        return staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => {
          const code = getScheduleCode(snapshot, s.id, dateStr);
          傳回總和 + (getShiftGroupByCode(code) === group ? 1 : 0);
        }, 0);
      };

      const countConsecutiveBeforeFromSnapshot = (snapshot, staffId, dateStr) => {
        設計數為0；
        let cursor = addDays(parseDateKey(dateStr), -1);
        while (true) {
          const key = formatDateKey(cursor);
          const code = getScheduleCode(snapshot, staffId, key);
          如果 (!isShiftCode(code)) 則跳出；
          計數加 1；
          cursor = addDays(cursor, -1);
        }
        返回計數；
      };

      const canAssignWithSnapshot = (snapshot, staff, dateStr, shiftCode) => {
        const reasons = [];
        const currentCode = getScheduleCode(snapshot, staff.id, dateStr);
        if (currentCode) Reasons.push('該格排已有班級或休假代碼');
        const prefix = getCodePrefix(currentCode);
        if (prefix && SMART_RULES.blockedLeavePrefixes.includes(prefix)) Reasons.push('該格現有休假，不可再排班');
        const StaffGroup = 員工.group || '白班';
        const shiftGroup = getShiftGroupByCode(shiftCode);
        if (!SMART_RULES.allowCrossGroupAssignment && shiftGroup && StaffGroup !== shiftGroup)reasons.push('不可跨群體排班');
        const prevKey = formatDateKey(addDays(parseDateKey(dateStr), -1));
        const prevCode = getScheduleCode(snapshot, staff.id, prevKey);
        const disallowed = SMART_RULES.disallowedNextShiftMap[prevCode] || [];
        if (disallowed.includes(shiftCode))reasons.push(`${prevCode} 後不可接 ${shiftCode}`);
        const consecutiveBefore = countConsecutiveBeforeFromSnapshot(snapshot, staff.id, dateStr);
        if (consecutiveBefore + 1 > SMART_RULES.maxConsecutiveWorkDays) Reasons.push(`連續上班不可超過 ${SMART_RULES.maxConsecutiveWorkDays} 天`);
        if (staff.pregnant && SMART_RULES.pregnancyRestrictedShifts.includes(shiftCode)) Reasons.push('懷孕標記人員不可排 N / 夜 8-8');
        返回 { allowed: reasons.length === 0, reasons };
      };

      const getWorkCountFromSnapshot = (snapshot, staffId) => {
        return daysInMonth.reduce((sum, d) => sum + (isShiftCode(getScheduleCode(snapshot, staffId, d.date)) ? 1 : 0), 0);
      };

      const getShiftCountFromSnapshot = (snapshot, staffId, shiftCode) => {
        return daysInMonth.reduce((sum, d) => sum + (getScheduleCode(snapshot, staffId, d.date) === shiftCode ? 1 : 0), 0);
      };

      const getLeaveCountFromSnapshot = (snapshot, staffId) => {
        return daysInMonth.reduce((sum, d) => sum + (isLeaveCode(getScheduleCode(snapshot, staffId, d.date)) ? 1 : 0), 0);
      };

      const getBlankCountFromSnapshot = (snapshot, staffId) => {
        return daysInMonth.reduce((sum, d) => sum + (!getScheduleCode(snapshot, staffId, d.date) ? 1 : 0), 0);
      };

      const canStillMeetRequiredLeavesAfterAssign = (snapshot, staffId) => {
        const currentLeaves = getLeaveCountFromSnapshot(snapshot, staffId);
        const remainingBlanks = getBlankCountFromSnapshot(snapshot, staffId);
        const remainingLeavesNeeded = Math.max(0, requiredLeaves - currentLeaves);
        返回剩餘空白頁數 >= 剩餘所需頁數；
      };

      const canStillMeetRequiredLeavesIfAssignShift = (snapshot, staffId) => {
        const currentLeaves = getLeaveCountFromSnapshot(snapshot, staffId);
        const remainingBlanks = getBlankCountFromSnapshot(snapshot, staffId);
        const remainingLeavesNeeded = Math.max(0, requiredLeaves - currentLeaves);
        返回（剩餘空白處 - 1）>= 剩餘所需葉數；
      };

      const getRecentWorkPressure = (snapshot, staffId, dateStr, lookback = 3) => {
        設計數為0；
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 0; i < lookback; i += 1) {
          const code = getScheduleCode(snapshot, staffId, formatDateKey(cursor));
          如果 (isShiftCode(code)) 計數 += 1;
          cursor = addDays(cursor, -1);
        }
        返回計數；
      };

      const getRecentLeavePressure = (snapshot, staffId, dateStr, lookback = 4) => {
        設計數為0；
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 0; i < lookback; i += 1) {
          const code = getScheduleCode(snapshot, staffId, formatDateKey(cursor));
          如果 (isLeaveCode(code)) 計數 += 1;
          cursor = addDays(cursor, -1);
        }
        返回計數；
      };

      const getNearbyLeavePressure = (snapshot, staffId, dateStr, radius = 2) => {
        設計數為0；
        const center = parseDateKey(dateStr);
        for (let offset = -radius; offset <= radius; offset += 1) {
          如果（偏移量 === 0）繼續；
          const key = formatDateKey(addDays(center, offset));
          const code = getScheduleCode(snapshot, staffId, key);
          如果 (isLeaveCode(code)) 計數 += 1;
        }
        返回計數；
      };

      const getDaysSinceLastLeave = (snapshot, staffId, dateStr, maxLookback = 10) => {
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 1; i <= maxLookback; i += 1) {
          const code = getScheduleCode(snapshot, staffId, formatDateKey(cursor));
          如果 (isLeaveCode(code)) 回傳 i;
          cursor = addDays(cursor, -1);
        }
        返回 maxLookback + 1；
      };

      const getGroupLeaveLoad = (snapshot, dateStr, group) => {
        return staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => {
          const code = getScheduleCode(snapshot, s.id, dateStr);
          返回 sum + (isLeaveCode(code) ? 1 : 0);
        }, 0);
      };

      const getConsecutiveLeavePattern = (snapshot, staffId, dateStr) => {
        const prevCode = getScheduleCode(snapshot, staffId, formatDateKey(addDays(parseDateKey(dateStr), -1)));
        const nextCode = getScheduleCode(snapshot, staffId, formatDateKey(addDays(parseDateKey(dateStr), 1)));
        const prevIsLeave = isLeaveCode(prevCode);
        const nextIsLeave = isLeaveCode(nextCode);
        返回 {
          prevIsLeave，
          nextIsLeave，
          相鄰休假次數：（上一個休假？1 : 0）+（下一個休假？1 : 0）
        };
      };

      const scoreCandidateWithSnapshot = (snapshot, staff, dateStr, shiftCode) => {
        設得分 = 0；
        score += (999 - getShiftCountFromSnapshot(snapshot, staff.id, shiftCode)) * SMART_RULES.fillPriorityWeights.sameShiftCount;
        score += (999 - getWorkCountFromSnapshot(snapshot, staff.id)) * SMART_RULES.fillPriorityWeights.totalShiftCount;
        如果 (getShiftGroupByCode(shiftCode) === (staff.group || '白班')) score += 100 * SMART_RULES.fillPriorityWeights.sameGroup;
        score -= getRecentWorkPressure(snapshot, staff.id, dateStr, 3) * 18;
        返回得分；
      };

      const scoreLeaveCandidateWithSnapshot = (snapshot, staff, dateStr) => {
        const leaveDeficit = Math.max(0, requiredLeaves - getLeaveCountFromSnapshot(snapshot, staff.id));
        const workCount = getWorkCountFromSnapshot(snapshot, staff.id);
        const 群組 = Staff.group || '白班';
        const sameDayLeaveLoad = getGroupLeaveLoad(snapshot, dateStr, group);
        const recentLeavePressure = getRecentLeavePressure(snapshot, staff.id, dateStr, 4);
        const nearbyLeavePressure = getNearbyLeavePressure(snapshot, staff.id, dateStr, 2);
        const daysSinceLastLeave = getDaysSinceLastLeave(snapshot, staff.id, dateStr, 10);
        const consecutiveLeavePattern = getConsecutiveLeavePattern(snapshot, staff.id, dateStr);
        設得分 = 0；
        得分 += 剩餘赤字 * 120；
        得分 += 工作次數 * 5；
        score += getRecentWorkPressure(snapshot, staff.id, dateStr, 3) * 18;
        score += Math.min(daysSinceLastLeave, 10) * 8;
        得分 -= sameDayLeaveLoad * 30;
        得分 -= 最近離職壓力 * 18；
        分 -= 附近離境壓力 * 28；
        如果 (consecutiveLeavePattern.adjacentLeaveCount === 1) score += 22;
        如果（連續休假模式.相鄰休假次數 >= 2）得分 -= 12；
        返回得分；
      };

      for (const day of targetDays) {
        for (const group of SHIFT_GROUPS) {
          如果（受限組 && 受限組 !== 組）繼續；

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
                返回 {
                  職員，
                  允許：result.allowed && canKeepLeaveTarget，
                  分數：result.allowed && canKeepLeaveTarget ? scoreCandidateWithSnapshot(mergedSchedule, staff, day.date, shiftCode) : -1
                };
              })
              .filter(item => item.allowed)
              .sort((a, b) => b.score - a.score);

            如果（可分配候選人.長度 === 0）{
              summary.skipped += 1;
              繼續;
            }

            const picked = assignableCandidates[0];
            setScheduleCode(mergedSchedule, picked.staff.id, day.date, shiftCode, 'auto');
            summary.workFilled += 1;
          }

          // 需求已滿後，只替補缺口，不足者補掉；其他空白保留
          const leaveCandidates = groupStaffs
            .filter(staff => !getScheduleCode(mergedSchedule, staff.id, day.date))
            .filter(staff => getLeaveCountFromSnapshot(mergedSchedule, staff.id) < requiredLeaves)
            .map(staff => ({ staff, score: scoreLeaveCandidateWithSnapshot(mergedSchedule, staff, day.date) }))
            .sort((a, b) => b.score - a.score);

          如果 (leaveCandidates.length > 0) {
            const currentLeaveLoad = getGroupLeaveLoad(mergedSchedule, day.date, group);
            const maxLeaveForDay = Math.max(0, groupStaffIds.size - demand);
            如果（目前請假負荷 < 最大請假天數）{
              const bestLeaveCandidate = LeaveCandidates[0];
              若 (最佳請假候選人 && 合併行程安排後仍能滿足所需請假) {
                setScheduleCode(mergedSchedule, bestLeaveCandidate.staff.id, day.date, 'off', 'auto');
                summary.leaveFilled += 1;
              }
            }
          }
        }
      }

      設定日程（合併後的日程）​​；
      saveToHistory(isPartial ? '規則指定補空' : '規則全月補空', mergedSchedule);
      setAiFeedback(`✅ 補空完成：上班 ${summary.workFilled} 格休假、 ${summary.leaveFilled} 格、未補成功 ${summary.skipped} 格`);
    } catch (error) {
      console.error(error);
      setAiFeedback("❌規則補空失敗，請檢查設定。");
    } 最後 {
      setIsAiLoading(false);
    }
  };

const callGemini = async (prompt, systemInstruction = "") => {
    令延遲時間為 1000；
    for (let i = 0; i < 5; i++) {
      嘗試 {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=${apiKey}`, {
          方法：'POST'，
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            內容：[{ parts: [{ text: prompt }] }],
            systemInstruction: systemInstruction ? { parts: [{ text: systemInstruction }] } : undefined,
            generationConfig: { responseMimeType: "application/json" }
          })
        });
        如果回應未成功，則拋出新的錯誤（'API 錯誤'）。
        const data = await response.json();
        return JSON.parse(data.candidates[0].content.parts[0].text);
      } catch (err) {
        如果 (i === 4) 拋出錯誤；
        await new Promise(resolve => setTimeout(resolve, delay));
        延遲 *= 2；
      }
    }
  };

  // ==========================================
  // 6.輔助統計與操作
  // ==========================================
  const getStaffStats = (staffId) => {
    const stats = {
      工作量：0，
      holidayLeave：0，
      總計休假：0，
      leaveDetails: Object.fromEntries(DICT.LEAVES.map(l => [l, 0]))
    };

    const mySchedule = schedule[staffId] || {};
    daysInMonth.forEach(d => {
      const cellData = mySchedule[d.date];
      const code = typeof cellData === 'object' && cellData !== null ? cellData.value : cellData;
      如果 (!code) 返回；

      如果 (DICT.SHIFTS.includes(code)) stats.work += 1;
      如果 (DICT.LEAVES.includes(code)) {
        stats.totalLeave += 1;
        如果 (stats.leaveDetails[code] !== undefined) stats.leaveDetails[code] += 1;
        如果 (d.isWeekend || d.isHoliday) stats.holidayLeave += 1;
      }
    });
    返回統計數據；
  };

  const getDailyStats = (dateStr) => {
    const stats = { D: 0, E: 0, N: 0, totalLeave: 0 };

    staffs.forEach(staff => {
      const cellData = schedule[staff.id]?.[dateStr];
      const code = typeof cellData === 'object' && cellData !== null ? cellData.value : cellData;
      如果 (!code) 返回；

      if (['D', '白8-8', '8-12', '12-16'].includes(code)) {
        stats.D += 1;
      } else if (['E', '夜8-8'].includes(code)) {
        stats.E += 1;
      } else if (code === 'N') {
        stats.N += 1;
      } else if (DICT.LEAVES.includes(code)) {
        stats.totalLeave += 1;
      }
    });

    返回統計數據；
  };

  const saveToHistory = (label, currentSchedule = schedule) => {
    const newRecord = {
      id: Date.now(),
      標籤，
      時間戳: new Date().toLocaleString(),
      狀態：{ 年、月、員工、日程安排：當前日程、顏色、自訂假日、特殊工作日、醫療日曆調整、人員配置、使用者介面設定 }
    };

    setHistoryList(prev => {
      const updated = [newRecord, ...prev].slice(0, 10);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
      返回更新後的結果；
    });
  };

  const loadHistory = (record) => {
    const { state } = record;
    設定年份(state.year);
    設定月份(state.month);
    setCustomHolidays(Array.isArray(state.customHolidays) ? state.customHolidays : []);
    setSpecialWorkdays(Array.isArray(state.specialWorkdays) ? state.specialWorkdays : []);
    setMedicalCalendarAdjustments(state.medicalCalendarAdjustments || { holidays: [], workdays: [] });
    如果 (state.staffingConfig) 設定 StaffingConfig(state.staffingConfig);
    如果 (state.uiSettings) setUiSettings(state.uiSettings);
    setStaffs(normalizeStaffGroup(state.staffs));
    設定日程（state.schedule）；
    如果 (state.colors) setColors(state.colors);
    設定顯示歷史記錄模式(false);
    setShowDraftPrompt(false);
  };

  const clearHistory = () => {
    if (window.confirm("確定要清空所有歷史記錄嗎？")) {
      localStorage.removeItem(STORAGE_KEY);
      設定歷史記錄列表([]);
    }
  };

  const handleCellChange = (staffId, dateStr, value) => {
    設定日程（上一個 => ({
      上一頁
      [staffId]: { ...prev[staffId], [dateStr]: value ? { value, source: 'manual' } : null }
    }));
  };

  const getCellCode = (staffId, dateStr) => {
    const cellData = schedule[staffId]?.[dateStr];
    return typeof cellData === 'object' && cellData !== null ? (cellData.value || '') : (cellData || '');
  };

  const getCellSource = (staffId, dateStr) => {
    const cellData = schedule[staffId]?.[dateStr];
    如果 (!cellData) 回傳 '';
    如果 (typeof cellData === 'object' && cellData !== null) 回傳 cellData.source || 'manual';
    返回“手動”；
  };

  const countConsecutiveWorkDaysBefore = (staffId, dateStr) => {
    設計數為0；
    let cursor = addDays(parseDateKey(dateStr), -1);
    while (true) {
      const key = formatDateKey(cursor);
      const code = getCellCode(staffId, key);
      如果 (!isShiftCode(code)) 則跳出；
      計數加 1；
      cursor = addDays(cursor, -1);
    }
    返回計數；
  };

  const canAssign = (staff, dateStr, shiftCode) => {
    const reasons = [];
    const currentCode = getCellCode(staff.id, dateStr);
    如果 (currentCode) {
      Reasons.push('該格已有排班或休假代碼');
    }

    const prefix = getCodePrefix(currentCode);
    如果（前綴 && SMART_RULES.blockedLeavePrefixes.includes(前綴)）{
      Reasons.push('該格已有休假，不可再排班');
    }

    const StaffGroup = 員工.group || '白班';
    const shiftGroup = getShiftGroupByCode(shiftCode);
    如果 (!SMART_RULES.allowCrossGroupAssignment && shiftGroup && staffGroup !== shiftGroup) {
      Reasons.push('不可跨群體排班');
    }

    const prevKey = formatDateKey(addDays(parseDateKey(dateStr), -1));
    const prevCode = getCellCode(staff.id, prevKey);
    const disallowed = SMART_RULES.disallowedNextShiftMap[prevCode] || [];
    如果 (disallowed.includes(shiftCode)) {
      Reasons.push(`${prevCode} 後不可接 ${shiftCode}`);
    }

    const consecutiveBefore = countConsecutiveWorkDaysBefore(staff.id, dateStr);
    如果 (consecutiveBefore + 1 > SMART_RULES.maxConsecutiveWorkDays) {
      Reasons.push(`連續上班不可超過 ${SMART_RULES.maxConsecutiveWorkDays} 天`);
    }

    如果 (staff.pregnant && SMART_RULES.pregnancyRestrictedShifts.includes(shiftCode)) {
      Reasons.push('懷孕標記人員不可排N / 夜8-8');
    }

    返回 { allowed: reasons.length === 0, reasons };
  };

  const getShiftCountForStaff = (staffId, shiftCode) => {
    return daysInMonth.reduce((sum, d) => sum + (getCellCode(staffId, d.date) === shiftCode ? 1 : 0), 0);
  };

  const scoreCandidate = (staff, dateStr, shiftCode) => {
    設得分 = 0；
    const stats = getStaffStats(staff.id);
    score += (999 - getShiftCountForStaff(staff.id, shiftCode)) * SMART_RULES.fillPriorityWeights.sameShiftCount;
    score += (999 - stats.work) * SMART_RULES.fillPriorityWeights.totalShiftCount;
    如果 (getShiftGroupByCode(shiftCode) === (staff.group || '白班')) {
      得分 += 100 * SMART_RULES.fillPriorityWeights.sameGroup;
    }
    返回得分；
  };

  const getCurrentConsecutiveLeavePattern = (staffId, dateStr) => {
    const prevCode = getCellCode(staffId, formatDateKey(addDays(parseDateKey(dateStr), -1)));
    const nextCode = getCellCode(staffId, formatDateKey(addDays(parseDateKey(dateStr), 1)));
    const prevIsLeave = isLeaveCode(prevCode);
    const nextIsLeave = isLeaveCode(nextCode);
    返回 {
      prevIsLeave，
      nextIsLeave，
      相鄰休假次數：（上一個休假？1 : 0）+（下一個休假？1 : 0）
    };
  };

  const openFillModal = (staff, dateStr) => {
    const 群組 = Staff.group || '白班';
    const shiftCode = DEFAULT_SHIFT_BY_GROUP[group];
    const dayInfo = daysInMonth.find(d => d.date === dateStr);
    const demand = dayInfo ? Number(staffingConfig?.requiredStaffing?.[(dayInfo.isWeekend || dayInfo.isHoliday) ? 'holiday' : 'weekday']?.[GROUP_TO_DEMAND_KEY[group]] || 0) : 0;
    const alreadyAssigned = staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => {
      const code = getCellCode(s.id, dateStr);
      傳回總和 + (getShiftGroupByCode(code) === group ? 1 : 0);
    }, 0);
    const LeaveCount = daysInMonth.reduce((sum, d) => sum + (isLeaveCode(getCellCode(staff.id, d.date)) ? 1 : 0), 0);
    const leaveDeficit = Math.max(0, requiredLeaves - leaveCount);

    const shiftResult = canAssign(staff, dateStr, shiftCode);
    const currentLeaves = leaveCount;
    const remainingBlanks = daysInMonth.reduce((sum, d) => sum + (!getCellCode(staff.id, d.date) ? 1 : 0), 0);
    const remainingLeavesNeeded = Math.max(0, requiredLeaves - currentLeaves);
    const canKeepLeaveTarget = (remainingBlanks - 1) >= remainingLeavesNeeded;

    const candidates = [];

    如果（已分配 < 需求 && 班次結果.允許 && 可以保留休假目標）{
      const reasonBits = [
        `${group}缺額尚未補滿`,
        `${shiftCode} 這組主班別`
      ];
      candidates.push({
        類型：'自動換檔'
        staffId: staff.id,
        員工姓名：員工姓名
        團體，
        shiftCode，
        允許：是，
        得分：scoreCandidate（員工，dateStr，shiftCode），
        原因：reasonBits
      });
    }

    如果（已分配 >= 需求 && 休假赤字 > 0）{
      const sameDayLeaveLoad = staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => sum + (isLeaveCode(getCellCode(s.id, dateStr)) ? 1 : 0), 0);
      const recentWorkPressure = (() => {
        設計數為0；
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 0; i < 3; i += 1) {
          如果 (isShiftCode(getCellCode(staff.id, formatDateKey(cursor)))) count += 1;
          cursor = addDays(cursor, -1);
        }
        返回計數；
      })();
      const consecutiveLeavePattern = getCurrentConsecutiveLeavePattern(staff.id, dateStr);
      let offScore = leaveDeficit * 100 + recentWorkPressure * 20 - sameDayLeaveLoad * 25;
      如果 (consecutiveLeavePattern.adjacentLeaveCount === 1) offScore += 18;
      如果（連續休假模式.相鄰休假次數 >= 2）offScore -= 10；
      candidates.push({
        類型：'自行離開'
        staffId: staff.id,
        員工姓名：員工姓名
        團體，
        shiftCode: 'off',
        允許：是，
        分數：offScore，
        Reasons: ['本月尚未達標', '當日負載需求已滿，優先補休']
      });
    }

    candidates.sort((a, b) => b.score - a.score);

    setSelectedFillCell({ staffId: staff.id, staffName: staff.name, dateStr, group: staff.group });
    setFillCandidates(candidates);
    設定顯示填滿模式(true);
  };

const openSelectedCellFillModal = () => {
    如果 (!selectedGridCell) 回傳;
    openFillModal(selectedGridCell.staff, selectedGridCell.dateStr);
  };

  const clearSelectedCell = () => {
    如果 (!selectedGridCell) 回傳;
    const { staff, dateStr } = selectedGridCell;
    const currentCode = getCellCode(staff.id, dateStr);
    如果 (!currentCode) 返回；
    if (!window.confirm(`確定清除此格內容？\n${staff.name}｜${dateStr}｜${currentCode}`)) return;

    設定日程（上一個 => ({
      上一頁
      [staff.id]: { ...prev[staff.id], [dateStr]: null }
    }));
    setSelectedGridCell(null);
    setAiFeedback(`🧹已清除${staff.name}在${dateStr}的內容`);
  };

  const clearRangeCells = () => {
    如果 (aiConfig.selectedStaffs.length === 0) {
      setAiFeedback('⚠️請先選擇要清除的人員');
      返回;
    }

    const start = Number(aiConfig.dateRange.start || 1);
    const end = Number(aiConfig.dateRange.end || 31);
    const targetStaffIds = new Set(aiConfig.selectedStaffs);

    let cleared = 0;
    設定日程（上一個 => {
      const next = JSON.parse(JSON.stringify(prev));
      staffs.forEach(staff => {
        如果 (!targetStaffIds.has(staff.id)) 回傳;
        daysInMonth.forEach(day => {
          如果（日期小於開始日期或日期大於結束日期）則傳回；
          const cellData = next[staff.id]?.[day.date];
          如果 (!cellData) 返回；

          const source = typeof cellData === 'object' && cellData !== null ? (cellData.source || 'manual') : 'manual';
          如果 (rangeClearMode === 'autoOnly' && source !== 'auto') 回傳;

          next[staff.id][day.date] = null;
          清除 += 1;
        });
      });
      返回下一個；
    });

    setAiFeedback(cleared > 0 ? `🧹已清除${cleared}格內容`: 'ℹ️指定範圍內沒有可清除的內容');
  };

  const applyFillCandidate = (candidate) => {
    如果 (!selectedFillCell) 回傳;
    handleCellChange(candidate.staffId, selectedFillCell.dateStr, candidate.shiftCode);
    setShowFillModal(false);
    setSelectedFillCell(null);
    設定候選列表([]);
    setSelectedGridCell(null);
  };

  const addStaff = (group = '白班') => {
    const newId = 's' + Date.now();
    setStaffs(prev => [...prev, { id: newId, name: '新成員', group }]);
    setSchedule(上一個 => ({ ...上一個, [newId]: {} }));
  };

  const removeStaff = (staffId) => {
    if (!window.confirm("確定要刪除此人員嗎？")) return;
    setStaffs(prev => prev.filter(s => s.id !== StaffId));
    設定日程（上一個 => {
      const next = { ...prev };
      刪除下一個[staffId]；
      返回下一個；
    });
  };

  const moveStaffInGroup = (staffId, direction) => {
    const newStaffs = [...staffs];
    const currentIndex = newStaffs.findIndex(s => s.id === staffId);
    如果 (currentIndex === -1) 返回；

    const currentGroup = newStaffs[currentIndex].group;
    const groupIndexes = newStaffs
      .map((staff, index) => ({ staff, index }))
      .filter(item => item.staff.group === currentGroup)
      .map(item => item.index);

    const currentGroupPos = groupIndexes.indexOf(currentIndex);
    const targetGroupPos = direction === 'up' ? currentGroupPos - 1 : currentGroupPos + 1;
    如果（targetGroupPos < 0 || targetGroupPos >= groupIndexes.length）返回；

    const targetIndex = groupIndexes[targetGroupPos];
    [newStaffs[currentIndex], newStaffs[targetIndex]] = [newStaffs[targetIndex], newStaffs[currentIndex]];
    設定員工（新員工）；
  };

  const groupedStaffs = useMemo(() => {
    返回 SHIFT_GROUPS.map(group => ({
      團體，
      staffs: staffs.filter(staff => (staff.group || '白班') === group)
    }));
  }, [員工]);

  返回 （
    <div className="min-h-screen text-slate-900 p-4 font-sans overflow-x-hidden relative" style={{ backgroundColor: pageBackgroundColor }}>
      <style>{`
        @keyframes pulse-once { 0% { transform: translateY(-10px); opacity: 0; } 100% { transform: translateY(0); opacity: 1; } }
        @keyframes fade-in-down { 0% { opacity: 0; transform: translateY(-5px); } 100% { opacity: 1; transform: translateY(0); } }
        .animate-pulse-once { animation: pulse-once 0.5s ease-out forwards; }
        .animate-fade-in-down { animation: fade-in-down 0.3s ease-out forwards; }
      `}</style>

      {showDraftPrompt && (
        <div className="max-w-[95vw] mx-auto mb-4 bg-amber-50 border border-amber-200 text-amber-800 p-4 rounded-xl flex items-center justify-between shadow-sm animate-fade-in-downdown">fadein-downdown">
          <div className="flex items-center gap-1.5">
            <Clock size={18} className="text-amber-600" />
            <span className="text-sm font-bold">偵測到先前暫存記錄。 </span>
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
              智慧排班｜智慧排班開發版
              <span className="text-blue-500 text-sm font-normal px-2 py-1 bg-blue-50 rounded-lg border border-blue-100">PRO v1.6.0</span>
            </h1>
            <p className="text-slate-500 text-xs mt-1 italic">開發版開發使用</p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            〈
              <儲存大小={16} /> 暫存
            </button>
            <button onClick={() => setShowHistoryModal(true)} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-2 rounded-xl font-bold.
              <時鐘大小={16} /> 歷史記錄
            </button>

            <div className="relative">
              <button onClick={() => setShowExportMenu(!showExportMenu)} className="flex items-center gap-1.5 bg-slate-800 text-white px-3 py-2 rounded-xl font-bold hover:bg-translate-9000 text-9000 text-text-900:bg-pold hover:bg-trans
                <Download size={16} /> 匯出
              </button>
              {showExportMenu && (
                <div className="absolute right-0 mt-2 w-48 bg-white border border-slate-200 rounded-xl shadow-xl z-50 overflow-hidden animate-fade-in-down">
                  <button onClick={exportToExcel} className="w-full text-left px-4 py-3 text-sm font-bold text-slate-700 hover:bg-green-50 hover:text-green-700 flex items-center gap-2 trans-colors bp">
                    <FileSpreadsheet size={16} /> Excel 高品質
                  </button>
                  <button onClick={exportToWord} className="w-full text-left px-4 py-3 text-sm font-bold text-slate-700 hover:bg-blue-50 hover:text-blue-700 flex items-center gap-2 transition-colorition-colorition-colorition-color-colorition-color-colors">
                    <FileText size={16} /> 橫向字（印刷）
                  </button>
                </div>
              )}
            </div>

            <div className="w-px h-8 bg-slate-200 mx-2 hidden sm:block"></div>

            <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200 flex-wrap">
              <button onClick={() => handleAiAutoSchedule(false)} disabled={isAiLoading} className="flex items-center gap-2 bg-white text-blue-600 px-3 py-2 rounded-fal-boldableed-blue-boldue. text-xs">
                {正在載入？ <Loader2 className="animate-spin" size={14} /> : <Sparkles size={14} />} 全月補空
              </button>
              <button onClick={() => setShowAiControl(!showAiControl)} className={`flex items-center gap-2 px-3 py-2 rounded-lg font-bold transition-all text-xs ${showAiControl ? 'bg-blue-600 transition-all text-xs ${showAiControl -bg-blue-600 集' hover:bg-slate-200'}`}>
                <Calendar size={14} /> 指定補空
              </button>
              <按鈕
                type="button"
                onClick={開啟選取儲存格填滿模態框}
                已停用={!selectedGridCell}
                className={`flex items-center gap-2 px-3 py-2 rounded-lg font-bold transition-all text-xs ${selectedGridCell ? 'text-slate-700 hover:bg-slate-200' : 'text-slate-400 cur-`not}
              >
                <Check size={14} /> 補此格
              </button>
              <按鈕
                type="button"
                onClick={清除選取儲存格}
                已停用={!selectedGridCell}
                className={`flex items-center gap-2 px-3 py-2 rounded-lg font-bold transition-all text-xs ${selectedGridCell ? 'text-red-600 hover:bg-red-50' : 'text-slate-400 cursor-not-allowed'
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
            <div>已大量儲存格：<span className="font-bold">{selectedGridCell.staff.name}</span>｜{selectedGridCell.dateStr}</div>
            <按鈕
              type="button"
              onClick={() => setSelectedGridCell(null)}
              className="px-3 py-1.5 rounded-lg border border-blue-200 bg-white text-blue-700 hover:bg-blue-100 transition-colors text-xs font-bold"
            >
              取消一些
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
                  <按鈕
                    key={s.id}
                    onClick={() => {
                      const next = aiConfig.selectedStaffs.includes(s.id)
                        aiConfig.selectedStaffs.filter(id => id !== s.id)
                        : [...aiConfig.selectedStaffs, s.id];
                      setAiConfig({ ...aiConfig, selectedStaffs: next });
                    }}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold border transition-all ${aiConfig.selectedStaffs.includes(s.id) ? 'bg-blue-600 border-blue-600 text-white shadowm-mdhwue text-blue-600 hover:bg-blue-100'}`}
                  >
                    {s.名稱}（{s.組 || '白班'}）
                  </button>
                ))}
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold text-blue-700 mb-2 uppercase">2. 日期範圍 ({aiConfig.dateRange.start} ~ {aiConfig.dateRange.end} 號碼)</label>
              <div className="flex items-center gap-2">
                <輸入
                  type="number"
                  最小值="1"
                  最大值="31"
                  值={aiConfig.dateRange.start}
                  onChange={(e) => setAiConfig({ ...aiConfig, dateRange: { ...aiConfig.dateRange, start: parseInt(e.target.value, 10) || 1 } })}
                  className="w-full border-blue-200 border p-2 rounded-lg text-sm text-center font-bold"
                />
                <span>至</span>
                <輸入
                  type="number"
                  最小值="1"
                  最大值="31"
                  值={aiConfig.dateRange.end}
                  onChange={(e) => setAiConfig({ ...aiConfig, dateRange: { ...aiConfig.dateRange, end: parseInt(e.target.value, 10) || 31 } })}
                  className="w-full border-blue-200 border p-2 rounded-lg text-sm text-center font-bold"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold text-blue-700 mb-2 uppercase">3. 指定班別（選填）</label>
              <選擇
                值={aiConfig.targetShift}
                onChange={(e) => setAiConfig({ ...aiConfig, targetShift: e.target.value })}
                className="w-full border-blue-200 border p-2 rounded-lg text-sm font-bold bg-white"
              >
                <option value="">依客戶需求自動補空</option>
                {AI_MAIN_SHIFTS.map(s => <option key={s} value={s}>{s} 班</option>)}
              </select>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-bold text-blue-700 mb-2 uppercase">4. 範圍清晰模式</label>
                <選擇
                  value={rangeClearMode}
                  onChange={(e) => setRangeClearMode(e.target.value)}
                  className="w-full border-blue-200 border p-2 rounded-lg text-sm font-bold bg-white"
                >
                  <option value="autoOnly">只清除自動補入內容</option>
                  <option value="all">清除範圍內全部內容</option>
                </select>
              </div>
              <div className="grid grid-cols-1 gap-3">
                <按鈕
                  disabled={isAiLoading || aiConfig.selectedStaffs.length === 0}
                  onClick={() => handleAiAutoSchedule(true)}
                  className="w-full bg-blue-600 hover:bg-blue-700 text-white font-black py-2 rounded-xl transition-all shadow-lg active:scale-95 disabled:opacity-50 disabled:grayscale flexd:grayscale 00acity-50 disabled:grayscale 0ite-mcenter g
                >
                  {正在載入？ <Loader2 className="animate-spin" size={18} /> : <Check size={18} />} 套用並補空
                </button>
                <按鈕
                  type="button"
                  disabled={aiConfig.selectedStaffs.length === 0}
                  onClick={clearRangeCells}
                  className="w-full bg-white border border-red-200 text-red-600 hover:bg-red-50 font-black py-2 rounded-xl transition-all disabled:opacity-50 disabled:grayscale flex apite-center
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
          <span className="text-xs opacity-80 uppercasetracking-wider">本月應休天數</span>
          <span className="text-2xl font-black">{requiredLeaves} <small className="text-sm font-normal">天數</small></span>
        </div>

        <div className="lg:col-span-5 flex items-center justify-end gap-2">
          <button onClick={() => changeScreen('entry')} className="bg-white border px-4 py-3 rounded-xl hover:bg-slate-50 transition-colors text-sm font-bold text-slate-700 flex items-center gap-2">
            <ArrowLeft size={18} className="text-slate-600" /> 回入口頁
          </button>
          <button onClick={() => changeScreen('settings')} className="bg-white border px-4 py-3 rounded-xl hover:bg-slate-50 transition-colors text-sm font-bold text-slate-700 flex 2ms-center gap-sm font-bold text-slate-700 flex 2ms-center-cent
            <Settings size={18} className="text-slate-600" /> 系統設定
          </button>
        </div>
      </div>

      <div className="max-w-[95vw] mx-auto bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto pb-4">
          <table className="w-max min-w-full border-collapse">
            <thead>
              <tr className="bg-slate-100 border-b-2 border-slate-200">
                <th className="sticky left-0 z-30 px-3 py-4 border-r font-black text-slate-700 w-20 min-w-[80px]" style={{ backgroundColor: shiftColumnBgColor }}>班次</th>
                <th className="sticky left-[80px] z-30 px-3 py-4 border-r font-black text-slate-700 w-32 min-w-[128px]" style={{ backgroundColor: nameDateColumnBgColth>
                {daysInMonth.map(d => (
                  <th
                    key={d.day}
                    className="px-1.5 py-2 border-r min-w-[44px] text-center"
                    style={{ backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent') }}
                  >
                    <div className={`${tableFontSizeClass} opacity-60 uppercase`} style={{ color: tableFontColor }}>{d.weekStr}</div>
                    <div className={`${tableFontSizeClass} font-black`} style={{ color: tableFontColor }}>{d.day}</div>
                  </th>
                ))}
                <th className="p-4 border-r min-w-[60px] bg-blue-50 text-blue-700 font-bold">上班</th>
                <th className="p-4 border-r min-w-[60px] bg-green-50 text-green-700 font-bold">假日休息</th>
                <th className="p-4 border-r min-w-[60px] bg-red-50 text-red-700 font-bold">總休</th>
                {DICT.LEAVES.map(l => (
                  <th key={l} className="p-2 border-r min-w-[40px] bg-slate-50 text-[10px] uppercase text-slate-500 font-bold">{l}</th>
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

                    返回 （
                      <tr key={staff.id} className="border-b border-slate-100 hover:bg-slate-50/50 transition-colors">
                        {index === 0 && (
                          <td rowSpan={groupCount} className="sticky left-0 z-20 border-r px-2 py-3 text-center shadow-[4px_0_10px_-5px_rgba(ColColBor:0,0.1)]" style={{ift>
                            <div className="flex items-center justify-center h-full min-h-[80px]">
                              <span className="text-[2rem] font-black leading-tight tracking-[0.14em] [writing-mode:vertical-rl]" style={{ color: tableFontColor }}>
                                {團體}
                              </span>
                            </div>
                          </td>
                        )}

                        <td className="sticky left-[80px] z-30 border-r px-2 py-2 shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)]" style={{ backgroundor: nameDate }ColumnBggor
                          <div className="flex items-center gap-2">
                            <div class Name="flex flex-col items-center justify-center shrink-0 w-6">
                              <按鈕
                                onClick={() => moveStaffInGroup(staff.id, 'up')}
                                disabled={groupIndex === 0}
                                className="text-slate-400 hover:text-blue-500 disabled:opacity-10 leading-none"
                              >
                                <ArrowUp size={14} />
                              </button>
                              <按鈕
                                onClick={() => moveStaffInGroup(staff.id, 'down')}
                                disabled={groupIndex === groupStaffList.length - 1}
                                className="text-slate-400 hover:text-blue-500 disabled:opacity-10 leading-none"
                              >
                                <ArrowDown size={14} />
                              </button>
                            </div>

                            <輸入
                              type="text"
                              value={員工姓名}
                              onChange={(e) => {
                                const next = [...]staffs];
                                const currentIndex = next.findIndex(s => s.id === staff.id);
                                如果 (currentIndex !== -1) next[currentIndex].name = e.target.value;
                                設定員工（下一個）；
                              }}
                              className={`flex-1 min-w-0 text-center py-1.5 font-bold border-none rounded-lg focus:ring-2 focus:ring-blue-400 bg-transparent ${tableFontSizeClass}`} style={{ color: tableFontor }}
                            />

                            <按鈕
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
                          返回 （
                            <td
                              key={d.date}
                              className={`border-r p-0 relative overflow-hidden ${selectedGridCell?.staff?.id === staff.id && selectedGridCell?.dateStr === d.date ? 'ring-2 ring-blue-500 ring-inset' : ''}`}
                              style={{ backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent'), opacity: d.isHoliday || d.isWeekend ? 0.9 : 1 }}
                            >
                              <div className="relative">
                                <選擇
                                  值={val}
                                  onChange={(e) => handleCellChange(staff.id, d.date, e.target.value)}
                                  className={`w-full h-10 text-center bg-transparent border-none cursor-pointer font-bold appearance-none hover:bg-black/5 ${tableFontSizeClass}`} style={{ color: tableFontColor }}
                                >
                                  <option value=""></option>
                                  <optgroup label="上班">
                                    {DICT.SHIFTS.map(s => <option key={s} value={s}>{s}</option>)}
                                  </optgroup>
                                  <optgroup label="休假">
                                    {DICT.LEAVES.map(l => <option key={l} value={l}>{l}</option>)}
                                  </optgroup>
                                </select>

                                <按鈕
                                  type="button"
                                  onClick={(e) => {
                                    e.停止傳播();
                                    setSelectedGridCell({ staff, dateStr: d.date });
                                  }}
                                  className="absolute right-1 top-1/2 -translate-y-1/2 z-0 w-3.5 h-3.5 flex items-center justify-center"
                                  aria-label={`一些 ${staff.name} ${d.date} 儲存格`}
                                  title={`大概${staff.name} ${d.date} 儲存格`}
                                >
                                  <span className={`w-2.5 h-2.5 rounded-full tr​​​​ansition-all ${selectedGridCell?.staff?.id === staff.id && selectedGridCell?.dateStr === d.date ? 'bg-blue-700 scale-110' hover:bg-blue-500'}`}></span>
                                </button>
                              </div>
                            </td>
                          ）；
                        })}

                        <td className={`border-r text-center font-black bg-blue-50/30 ${tableFontSizeClass}`} style={{ color: tableFontColor }}>{stats.work}</td>
                        <td className={`border-r text-center font-black bg-green-50/30 ${tableFontSizeClass}`} style={{ color: tableFontColor }}>{stats.holidayLeave}</td>
                        <td className={`border-r text-center font-black bg-red-50/30 ${tableFontSizeClass}`} style={{ color: tableFontColor }}>{stats.totalLeave}</td>
                        {DICT.LEAVES.map(l => (
                          <td key={l} className={`border-r text-center bg-slate-50/20 ${tableFontSizeClass}`} style={{ color: tableFontColor }}>
                            {stats.leaveDetails[l] || ''}
                          </td>
                        ))}
                      </tr>
                    ）；
                  })}

                  <tr className="border-b border-slate-200 bg-slate-50/70">
                    <td className="sticky left-[80px] z-30 border-r px-2 py-2 shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)]" style={{ backgroundor: nameDate }ColumnBggor
                      <div className="flex items-center justify-center">
                        <按鈕
                          onClick={() => 新增員工(群組)}
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

                    <td colSpan={3 + DICT.LEAVES.length}></td>
                  </tr>
                </React.Fragment>
              ))}
            </tbody>

            {uiSettings?.showStats && (
            <tfoot className="bg-slate-100 border-t-2 border-slate-200">
              {['D', 'E', 'N', 'totalLeave'].map((rowKey) => (
                <tr key={rowKey}>
                  <td className="sticky left-0 z-10 border-r p-3 w-24 min-w-[96px]" style={{ backgroundColor: shiftColumnBgColor }}></td>
                  <td className={`sticky left-[96px] z-10 border-r p-3 text-right font-bold min-w-[144px] ${tableFontSizeClass}`} style={{ backgroundColor: nameDateColumnBgColor, color: tableFontColor }}>
                    {rowKey === 'totalLeave' ? '當日休假' : `${rowKey} 班次`}
                  </td>
                  {daysInMonth.map(d => {
                    const count = getDailyStats(d.date)[rowKey];
                    返回 （
                      <td key={d.date} className={`border-r p-2 text-center font-black ${tableFontSizeClass}`} style={{ color: tableFontColor }}>
                        {count || ''}
                      </td>
                    ）；
                  })}
                  <td colSpan={3 + DICT.LEAVES.length}></td>
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
              <button onClick={() => { setShowFillModal(false); setSelectedFillCell(null); setFillCandidates([]); }} className="p-2 hover:bg-slate-200 rounded-full tr​​​​s">ansition-colors">ans
                <X />
              </button>
            </div>
            <div className="p-4 max-h-[60vh] overflow-y-auto space-y-3">
              {fillCandidates.length === 0 ? (
                <div className="p-6 text-center text-slate-500">目前沒有可直接建議的班別或人員。 </div>
              ): fillCandidates.map((candidate, index) => (
                <div key={`${candidate.staffId}-${candidate.shiftCode}`} className="border rounded-2xl p-4 flex items-center justify-between gap-4 hover:border-blue-400 hover:bg-blue-50/40+
                  <div>
                    <div className="font-black text-slate-800">{candidate.staffName} → {candidate.shiftCode}</div>
                    <div className="text-xs text-slate-500 mt-1"> 群體：{candidate.group}｜排序分數：{candidate.score}</div>
                    {candidate.reasons?.length > 0 && (
                      <div className="text-xs text-slate-500 mt-1">{candidate.reasons.join('｜')}</div>
                    )}
                  </div>
                  <button onClick={() => applyFillCandidate(candidate)} className="px-4 py-2 rounded-xl bg-blue-600 text-white font-bold hover:bg-blue-700">
                    {索引 === 0 ? '選擇推薦' : '選擇'}
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
              <h3 className="font-black text-slate-800 flex items-center gap-2"><Clock />歷史檔案記錄</h3>
              <button onClick={() => setShowHistoryModal(false)} className="p-2 hover:bg-slate-200 rounded-full tr​​​​ansition-colors">
                <X />
              </button>
            </div>

            <div className="max-h-[60vh] overflow-y-auto p-4 space-y-3">
              {historyList.length === 0 ? (
                <p className="text-center py-10 text-slate-400 font-bold">目前尚無檔案記錄</p>
              ): (
                historyList.map(record => (
                  <div key={record.id} className="flex items-center justify-between p-4 border rounded-2xl hover:border-blue-400 hover:bg-blue-50/30 transition-all group">
                    <div>
                      <div className="font-black text-slate-800">
                        {record.label}
                        <span className="text-xs font-normal text-slate-400 ml-2">{record.state.year}/{record.state.month}</span>
                      </div>
                      <div className="text-xs text-slate-500">{record.timestamp}</div>
                    </div>
                    <button onClick={() => loadHistory(record)} className="bg-slate-100 text-slate-700 px-4 py-2 rounded-xl text-sm font-bold hover:bg-blue-600 hover:text-wall-wall-wall-wall-wall-wall-wall
                      載入
                    </button>
                  </div>
                ））
              )}
            </div>

            <div className="p-4 bg-slate-50 border-t flex gap-3">
              <button onClick={clearHistory} className="flex-1 py-3 text-sm font-bold text-red-500 hide:bg-red-50 rounded-xltransition-colors">清空所有記錄</button>
              <button onClick={() => setShowHistoryModal(false)} className="flex-1 py-3 text-sm font-bold text-slate-600 bg-white border rounded-xl hover:bg-slate-100 transbutton>
            </div>
          </div>
        </div>
      )}
    </div>
  ）；
}


function SettingRow({ icon: Icon, title, desc, children, iconBg = 'bg-blue-50', iconColor = 'text-blue-600' }) {
  返回 （
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
  ）；
}

function SettingsView({ changeScreen, colors, setColors, customHolidays, setCustomHolidays, specialWorkdays, setSpecialWorkdays, medicalCalendarAdjustments, setMedicalCalendarAdjustments, stuiingConfigettSaffidjustments.
  const [holidayInput, setHolidayInput] = useState({ year: '', month: '', day: '' });
  const addCustomHoliday = () => {
    const y = holidayInput.year.trim();
    const m = holidayInput.month.trim();
    const d = holidayInput.day.trim();
    如果 (!y || !m || !d) 返回；
    const dateStr = `${y.padStart(4,'0')}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;
    如果 (customHolidays.includes(dateStr)) 回傳;
    setCustomHolidays(prev => [...prev, dateStr].sort());
    setHolidayInput({ year: '', month: '', day: '' });
  };
  const removeCustomHoliday = (dateStr) => setCustomHolidays(prev => prev.filter(item => item !== dateStr));
  const fixedGroups = [
    { icon: ShieldCheck, ti​​tle: '排班核心規則', items: ['每日至少一名主管級別值班', '夜班後必須小時接續至少 24 休息', '特定危險單位需雙人以上配置'] },
    { icon: Clock, ti​​tle: '班別接續續', items: ['禁止「花班」：白班不能直接接續大夜', '小夜轉白班間隔需大於12小時', '連班上限不得超過6天'] },
    { icon: UserCheck, ti​​tle: '請假與休假限制', items: ['法定假日優先排休計算', '年度特休不得低於勞基法標準', '病假/喪假排位權重調整限制'] },
    { icon: Database, title: '核心資料格式', items: ['員工編號固定為8位數字', '班表週期固定為1自然個月', '資料匯出格式限定為.xlsx / .pdf'] }
  ];
  const extGroups = [
    { title: '控制權限', desc: '各級主管登錄權限', icon: ShieldCheck },
    { title: '雲端同步', desc: '即時異地備份機制', icon: 資料庫 },
    { title: '單位自訂規則', desc: '細化至科其他規則', icon: 加號 },
    { title: '進階 AI 條件', desc: '複雜人員偏好學習', icon: Cpu }
  ];
  返回 （
    <div className="min-h-screen bg-gray-50 text-slate-800 font-sans selection:bg-blue-100">
      <header className="sticky top-0 z-50 bg-white border-b border-gray-200 px-8 py-4 flex justify-between items-center shadow-sm">
        <div><div className="flex items-center gap-2 mb-0.5"><Settings className="w-6 h-6 text-blue-600" /><h1 className="text-xl font-boldtracking-tight text-gray-900">系統設定</p1></divtext> text-gray-500">可調整使用者設置，核心規則由系統固定管理</p></div>
        <div className="flex items-center gap-3">
          <button onClick={() => changeScreen('entry')} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-600 bg-white border border-gray-300 rounded-class h0000 />返回入口頁</button>
          <button onClick={() => changeScreen('schedule')} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-600 bg-white border border-gray-3000 rounded-40000-FFy-50000-FFy-3000 />返回排班頁</button>
          <button onClick={() => changeScreen('schedule')} className="flex items-center gap-2 px-6 py-2 text-sm font-semibold text-white bg-blue-600 rounded-sm :bue-blue-700098<am<m<m-mddow:bue className="w-4 h-4" />儲存設定</button>
        </div>
      </header>
      <main className="max-w-7xl mx-auto p-8 space-y-10">
        <section className="space-y-6">
          <div className="flex items-center gap-2"><div className="w-1 h-6 bg-blue-600 rounded-full"></div><h2 className="text-lg font-bold text-gray-800">使用者偏好設定</h2></div>
          <div className="space-y-5">
            <SettingRow icon={Monitor} title="外觀與顯示" desc="調整班表顏色、顯示大小與統計欄位置顯示。">
              <div className="space-y-6">
                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">顏色標誌</label>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">週末顏色</span><input type="color" value={colors.weekends> 週末顏色</span><input type="color" value={colors.weekended={colors.>>. weekend: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">假日顏色</span><input type="color" value={colors.holi} 假日顏色</span><input type="color" value={colors.holi} 假日顏色</span>(prev. holiday: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">主頁背景顏色</span><input type="color" value={uiSettings.pagecm. setUiSettings(prev => ({ ...prev, pageBackgroundColor: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">表格字體顏色</span><input type="color" value={uiSett>.">表格字體顏色</span><input type="color" value={uiSett> tableue=Sat; => ({ ...prev, tableFontColor: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">班別背景顏色</span><input type="color" value={iftSettings.shue={iftS. setUiSettings(prev => ({ ...prev, shiftColumnBgColor: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">姓名/日期欄背景顏色</span><input type="color" value={Settings. setUiSettings(prev => ({ ...prev, nameDateColumnBgColor: e.target.value }))} className="w-10 h-8 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>>
                  </div>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="flex items-center justify- Between"><span className="text-sm font-medium">表格顯示大小</span><select value={uiSettings.tableDensity} onChange={(e) => setUiSettings(prev =>() class. border-none bg-gray-100 rounded-md px-3 py-2"><option value="standard">標準（預設）</option><option value="compact">壓縮</option><option value="relaxed">模組化</option></select></div>
                  <div className="flex items-center justify-between"><span className="text-sm font-medium">表格字體大小</span><select value={uiSettings.tableFontSize} onChange={(e) => setUiSettings(prevs=)s)(prev=)s_p. border-none bg-gray-100 rounded-md px-3 py-2"><option value="small">小</option><option value="medium">標準</option><option value="large">大</option></select></div>
                  <div className="flex items-center justify-between"><span className="text-sm font-medium">顯示統計位</span><button type="button" onClick={() => setUiSettings(prev => ({ ...prev, showStats: !prev. relative transition-colors ${uiSettings.showStats ? 'bg-blue-600' : 'bg-gray-300'}`}><div className={`absolute top-1 w-3 h-3 bg-white rounded-full tr​​trute top-1 w-3 h-3 bg-white rounded-full tr​​transition-all ${uiSettings: truiS. 'left-1'}`}></div></button></div>
                </div>
              </div>
            </SettingRow>
            <SettingRow icon={Layout} title="班表內容自訂" desc="設定自訂行程代碼、班級其他方向順序與延伸欄位。" iconBg="bg-indigo-50" iconColor="text-indigo-600">
              <div className="space-y-5"><div><label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">自訂彈性程式碼</label><div className="flex flex-wrap gap-2">{PL',</label><div className="flex flex-wrap gap-2">{PL', <.' key={code} className="px-3 py-1.5 bg-gray-100 text-gray-600 text-xs font-bold rounded-md border border-gray-200">{code}</span>)}<button className="p-1.50-blue">{code}</span>)}<button className="p-1.50. className="w-4 h-4" /></button></div></div><div><label className="text-sm font-medium block mb-2">班別顯示順序</label><div className="text-xs text-gray-500 p-4 bg-graordery-50 borderporder-d text-center">拖放排序功能開發</div></div><div className="pt-3 border-t border-gray-100"><button className="text-sm text-blue-600 font-medium flex items-center gap-1 hide:underline"><Plus className="w-div. className="rounded-xl bg-blue-50 border border-blue-100 px-4 py-3 text-xs text-blue-700">系統已支援固定國歷、補假規則、清明/端午/中秋/除夕與春節推算；2024–2026 清明/端午/中秋/除夕與春節推算；2024–2026 清明/端午/中秋/除夕與春節推算；2024–2026 清明/端午/中秋/除夕</div></div>
            </SettingRow>
            <SettingRow icon={UserCheck} title="人力需求設定" desc="獨立設定各平日/假日班需求，作為全月補空與指定補空的直接關聯。" iconBg="bg-sky-50" iconColor="text-sky-600">
              <div className="space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div>
                    <label className="text-xs font-semibold text-gray-400 uppercase Tracking-wider block mb-2">醫院體系</label>
                    <選擇
                      value={staffingConfig.hospitalLevel}
                      onChange={(e) => setStaffingConfig(prev => ({ ...prev, HospitalLevel: e.target.value }))}
                      className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-gray-50"
                    >
                      <選項值=“醫療”>醫學中心</選項>
                      <option value="regional">區域醫院</option>
                      <option value="local">地區醫院</option>
                    </select>
                  </div>
                  <div>
                    <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">總床數</label>
                    <輸入
                      type="number"
                      最小值="0"
                      值={staffingConfig.totalBeds}
                      onChange={(e) => setStaffingConfig(prev => ({ ...prev, TotalBeds: parseInt(e.target.value, 10) || 0 }))}
                      className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-gray-50"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-semibold text-gray-400 uppercase track-wider block mb-2">護理師總數</label>
                    <輸入
                      type="number"
                      最小值="0"
                      value={staffingConfig.totalNurses}
                      onChange={(e) => setStaffingConfig(prev => ({ ...prev, TotalNurses: parseInt(e.target.value, 10) || 0 }))}
                      className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-gray-50"
                    />
                  </div>
                </div>

                <div className="rounded-xl bg-sky-50 border border-sky-100 px-4 py-3 text-xs text-sky-700">
                  參考防護病比：{HOSPITAL_LEVEL_LABELS[staffingConfig.hospitalLevel]}｜白班 {HOSPITAL_RATIO_HINTS[staffingConfig.hospitalLevel]}｜白班 {HOSPITAL_RATIO_HINTS[staffingConfig.hospitalLevel].white}、小夜 {HOSPITAL_RATIO_H.Sf.Spi. {HOSPITAL_RATIO_HINTS[staffingConfig.hospitalLevel].night}
                </div>

                <div className="rounded-xl border border-gray-200 bg-white px-4 py-3 text-sm text-gray-600">
                  <div className="font-semibold text-gray-800 mb-1">目前補空樣本</div>
                  <div>平日：白班 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.weekday.white}</span> 人、小夜 <span className="font-bold text-sky-700">{staffingConfiging.reaffday. className="font-bold text-sky-700">{staffingConfig.requiredStaffing.weekday.night}</span> 人</div>
                  <div>假期：白班 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.holiday.white}</span> 人、小夜 <span className="font-bold text-sky-700">{staffingConfig.relidaffing. className="font-bold text-sky-700">{staffingConfig.requiredStaffing.holiday.night}</span> 人</div>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                  <div className="rounded-2xl border border-gray-200 p-4 bg-gray-50/50">
                    <h4 className="font-bold text-gray-800 mb-4">平日需求</h4>
                    <div className="grid grid-cols-3 gap-3">
                      <div>
                        <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">白班</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.white}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, white: parseInt(e.target.value, 10) ||}
                          className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">小夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.evening}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, evening: parseInt(e.target.value, 10) ||)
                          className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">大夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.night}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, night: parseInt(e.target.value, 10) ||}
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
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, holiday: { ...prev.requiredStaffing.holiday, white: parseInt(e.target.value, 10) ||}} } }) ||
                          className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">小夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.holiday.evening}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, holiday: { ...prev.requiredStaffing.holiday, evening: parseInt(e.target.value, 10) ||)
                          className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                      <div>
                        <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">大夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.holiday.night}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, holiday: { ...prev.requiredStaffing.holiday, night: parseInt(e.target.value, 10) ||}
                          className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-white" />
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </SettingRow>
            <SettingRow icon={Calendar} title="假期新增" desc="使用西曆年月日新增自訂假期，並可個別刪除。">
              <div className="space-y-5"><div><label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">西曆年、月value={holidayInput.year} onChange={(e)=>setHolidayInput({ ...holidayInput, year: e.target.value })} className="w-full px-3 py-2.5 text-smcus: border-gray-20000-d-ff-ex-ra-rs-v-gray-200-00000m-vt 500000 月 00000000000004mt focus:ring-blue-100" /><input type="number" placeholder="月" value={holidayInput.month} onChange={(e)=>setHolidayInput({ ...holidayInput, month: e.target.value })} class="p=wp. border-gray-200 rounded-xl bg-gray-50 focus:outline-none focus:ring-2 focus:ring-blue-100" /><input type="number" placeholder="日" value={holidayInput.day} onChange=( e.target.value })} className="w-full px-3 py-2.5 text-sm border border-gray-200 rounded-xl bg-gray-50 focus:outline-none focus:ring-2 focus:ring-blue-100" /></divdivi focus:ring-2 focus:ring-blue-100" /></divdivi> className="w-full md:w-auto flex items-center justify-center gap-2 px-5 py-2.5 text-sm font-semibold text-white bg-blue-600 rounded-smh hover:bg-blue-700">0 className="pt-3 border-t border-gray-100"><label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-3">已新增假期</label><div className="space-y-2 max-h-52 ">已新增假期</label><div className="space-y-2 max-h-52 ==-d ? <div className="text-xs text-gray-400 p-4 bg-gray-50 border border-dashed border-gray-300 rounded-xl text-center">尚未新增自訂假期</div> : customHolidays.map(dateStr => <divStrkey-date</div> class=dateHolidays.map(dateStr => <divStrkey-date</div>>>> flexm-date flexpwems> <Strkeys-datedatedate}o flex-3 flexpwem-Bad. bg-gray-50 border border-gray-200 rounded-xl"><span className="text-sm text-gray-700 font-medium">{dateStr}</span><button onClick={() => removeCustomHoliday(span><button onClick={() => removeCustomHoliday(dateStrs-8m="dateStr. border border-red-200 text-red-500 hover:bg-red-50 font-bold">-</button></div>)}</div></div></div>
            </SettingRow>
            <SettingRow icon={Grid} title="補空優先" desc="設定補空優先偏好，作為後續智慧補班的參考方向。" iconBg="bg-teal-50" iconColor="text-teal-600">
              <div className="space-y-3">{[{ label: '優先分配缺額最多的班別', active: true }, { label: '優先完成個人班數平均', active: true }, { label: '優先減少跨天連班', active: false }, { label: ' key={idx} className="flex items-center間隙-3 p-3懸停：bg-gray-50 rounded-xl遊標指針"><div className={`w-4 h-4圓形邊框flex items-center justify-center收縮-0 ${item.active/10g-blue'60g border-gray-300'}`}>{item.active && <div className="w-1.5 h-1.5 bg-white rounded-full"></div>}</div><span className="text-sm text-gray-70order">{item.label}</span></divtext-sm text-gray-70order">{item.Pptel}</span></div className="text-sm text-blue-600 font-medium flex items-center gap-1 hover:underline"><Plus className="w-3.5 h-3.5" /> 新增自訂偏好</button></div></div>
            </SettingRow>
          </div>
        </section>
        <section className="space-y-6">
          <div className="flex items-center justify- Between"><div className="flex items-center gap-2"><div className="w-1 h-6 bg-amber-500 rounded-full"></div><h2 className="text-lg font-bold="Apdiv-full"></divp-div items-center gap-1.5 px-3 py-1 bg-amber-50 text-amber-700 圓角邊框 border-amber-100"><Lock className="w-3.5 h-3.5" /><span className="text-xxs font-bold uppercapercaing-uprack0/uppercas 片段模式使用者不可修改</span></div></div>
          <div className="bg-slate-50 border border-slate-200 rounded-3xl overflow-hidden shadow-inner"><div className="grid grid-cols-1 md:grid-cols-2 gap-px bg-slate-200">{fixed,Pp. className="bg-white p-6 hover:bg-slate-50 transition-colors"><div className="flex items-center gap-3 mb-4"><group.icon className="w-5 h-5 text-slate-400" /><h4 className="font-bold" className="space-y-3">{group.items.map((item, j) => <li key={j} className="flex items-start gap-3"><CheckCircle2 className="w-4 h-4 text-green-500 mt-0.5 shrink-0" /><s 類型leading-relaxed">{item}</span></li>)}</ul></div>)}<div className="md:col-span-2 bg-slate-900 text-white p-8"><div className="flex items-center gap-3 mb-6">< class-pue="p-divg border-blue-400/30"><Cpu className="w-5 h-5 text-blue-400" /></div><div><h4 className="font-bold text-lg">AI 智慧型補班基限制 (核心邏輯)</h4><p className="text-xs text-slate-cm片. Intelligent Logic約束</p></div></div><div className="grid grid-cols-1 md:grid-cols-3 gap-8"><div className="space-y-2"><div className="text-blue-400 text-xs font-bold">">重播權</divtext<p text-slate-300">AI禁止將權重設為負值，所有補空邏輯均需符合人員適任性評分，系統會鎖定鎖定適任性分數不被手動覆蓋。 </p></div><div className="space-y-2"><div className="text-blue-400 text-xs font-bold">覆蓋衝突限制</div><p className="text-sm text-slate-300">手動排班擁有100%絕對優先權。 AI 補空演算時自動迴避「手動鎖定」儲存格，且不主動修改現有固定排程。 </p></div><div className="space-y-2"><div className="text-blue-400 text-xs font-bold">合規性佈置</div><p className="text-sm text-slate-300">補空結果尚需以「三重合規」，若違反勞基檢定」，若違反勞基測試時拒絕此工法。 </p></div></div></div></div></div>
        </section>
        <section className="space-y-6"><div className="flex items-center gap-2"><div className="w-1 h-6 bg-purple-600 rounded-full"></div><h2 className="text-lg font-bold text-gray-full"></div><h2 className="text-lg font-bold text-gray-8002-div8002 sm:grid-cols-2 lg:grid-cols-4 gap-4">{extGroups.map((ext, idx) => <div key={idx} className="group p-5 bg-white border border-gray-100 rounded-2group p-5 bg-white border border-gray-100 rounded-2group p-5 bg-white border border-gray-100 rounded-2group shadowple-200 opacity-75"><div className="mb-4 p-2 w-fit bg-purple-50 rounded-xl group-hover:bg-purple-100 transition-colors"><ext.icon className="w-5 h-5 text-purple-600" /><ext.icon<h4 classtext-0-bold text> mb-1">{ext.title}</h4><p className="text-xs text-gray-500 mb-4">{ext.desc}</p><div className="flex items-center text-[10px] font-bold text-purple-400 uppercase-cracking-cm-c​​yk-purple-400 uppercase-0000 月py-0.5 rounded">即將推出</span></div></div>)}</div></section>
      </main>
      <footer className="max-w-7xl mx-auto px-8 py-12 border-t border-gray-200 flex flex-col md:flex-row justify-between items-center gap-4"><div className="flex items-center gap-4"> divo gap-4"><div 類h-3 bg-green-500 rounded-full animate-pulse"></div><span className="text-sm font-semibold text-gray-700">系統核心版本：v2.5.0-PRO</span></div><span className="text-gray-300] class</poo<p; text-gray-500">智慧排班引擎開發測試版</span></div><div className="text-sm text-gray-400">© 2024 Intelligent Scheduling System PRO.版權所有。 </div></footer>
    </div>
  ）；
}

function EntryView({ changeScreen, goToLatestHistory }) {
  const [account, setAccount] = useState('');
  const [password, setPassword] = useState('');
  const handleLogin = (e) => { e.preventDefault(); changeScreen('schedule'); };
  返回 （
    <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 font-sans text-slate-800">
      <div className="mb-10 text-center"><div className="flex items-center justify-center gap-3 mb-4"><div className="bg-blue-600 p-2.5 rounded-xl shadow-md shadow-blue-200/50+Calwendo. /></div><div className="text-left"><h1 className="text-2xl font-bold tracking-tight text-slate-900 flex items-center">智慧排班｜智慧排班開發版<span className="ml-3 inline-flex items-center px-p. bg-blue-100 text-blue-800 border border-blue-200">PRO v1.6.0</span></h1><p className="text-slate-500 text-xs font-medium track-wide mt-1">開發版開發使用</div>
      <div className="w-full max-w-[440px] bg-white rounded-2xl Shadow-[0_12px_40px_rgba(0,0,0,0.03)] border border-slate-200/60 Overflow-hidden"><10class class> className="text-xl font-bold text-slate-900 mb-2">歡迎使用智慧排班系統</h2><p className="text-sm text-slate-500leading-relaxed">您可以直接進入排班系統、開啟最近編輯過的班表，或前往調整系統設定參數。 </p></div><form onSubmit={handleLogin} className="space-y-6 mb-10"><div className="space-y-2"><label className="block text-sm font-semibold text-slate-700">帳號</label>< class-00rel>divName-0rel-700">帳號</label>divvv-0 pl-3.5 flex items-center pointer-events-none"><User className="h-4.5 w-4.5 text-slate-400" /></div><input type="text" className="block w-full pl-11 pr-4 py-3 border borderorder outline-none placeholder:text-slate-400" placeholder="請輸入您的管理帳號" value={account} onChange={(e)=>setAccount(e.target.value)} /></div></div><font className="space-divy-2">label classNameb. text-slate-700">密碼</label><div className="relative group"><div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none"><Lock className="h-4.50-45 text-pdiv-400 /plate">">">">"> "> "><p. w-full pl-11 pr-4 py-3 border border-slate-200 rounded-xl bg-slate-50/50 outline-none placeholder:text-slate-400" placeholder="••••••••" value={password} onChange={edivue). 。點擊底部按鈕即可直接進入系統環境。 </p></div></form><div className="space-y-4"><button type="button" onClick={() => changeScreen('schedule')} className="w-full flex justify-center items-center gap-2 py-3.5-x-4-texted bg-blue-600 hoover:bg-blue-700"><ShieldCheck className="w-4.5 h-4.5" />進入班排系統</button><div className="grid grid-cols-2 gap-4"><button type="button" onCapify=gomosooo-Toooo-too-too​​y'm-akak-2m-testCapify-2ms-1 月）-2000 py-3 px-4 border border-slate-200 rounded-xl shadow-sm text-sm font-bold text-slate-700 bg-white hover:bg-slate-50"><Clock className="w-4 h-4 text-slate-500" />最近班表</button><button type="button" onClick={() => changeScreen('settings')} className=" justify-center ite b-center gaporder-20-pyplate-20-p-pyp-center-20-20-20-20p-20-20-20-20-20-20-20-20-20-20-20-20-20-20-20-20-20-20-20-20-20-20-20-20-20 pp-20-20-20-20-20p-20-20-20-20-20-20-20-20 pp; rounded-xl shadow-sm text-sm font-bold text-slate-700 bg-white hover:bg-slate-50"><Settings className="w-4 h-4 text-slate-500" />系統設定</button></div></divp></divp-divp-0p-00001/0 月> border-slate-100 flex justify-between items-center"><button className="text-xs text-slate-500 hover:text-blue-600 font-bold flex items-center gap-1transition-colors group">開發版本說明<Chevron-ight items-center gap-1transition-colors group">開發版本說明<Chevron-ight . transition-transform" /></button><button className="text-xs text-slate-500 hide:text-blue-600 font-boldtransition-colors">查看版本資訊</button></div></div><div className="mt-12 text-center"><p class 類font-mediumtracking-[0.1em] uppercase">© {new Date().getFullYear()} 智慧排班系統 PRO.</p></div>
    </div>
  ）；
}

export default function App() {
  const [screen, setScreen] = useState('entry');
  const [colors, setColors] = useState({ weekend: '#dcfce7', holiday: '#fca5a5' });
  const [customHolidays, setCustomHolidays] = useState([]);
  const [specialWorkdays, setSpecialWorkdays] = useState([]);
  const [medicalCalendarAdjustments, setMedicalCalendarAdjustments] = useState({ holidays: [], workdays: [] });
  const [uiSettings, setUiSettings] = useState({
    pageBackgroundColor: '#f8fafc',
    表格字體大小：'中等'
    表格字體顏色：'#1f2937'，
    shiftColumnBgColor: '#ffffff',
    nameDateColumnBgColor: '#ffffff',
    表格密度：'標準'
    顯示統計資訊：是
  });
  const [staffingConfig, setStaffingConfig] = useState({
    醫院級：'區域性'
    總床數：60
    護士總數：20
    所需人員配置：{
      平日：{ 白色：6， 傍晚：3， 夜晚：2 }，
      假日：{ 白色：4， 傍晚：2， 夜晚：2 }
    }
  });
  const [loadLatestOnEnter, setLoadLatestOnEnter] = useState(false);

  const goToSchedule = () => {
    setLoadLatestOnEnter(false);
    設定畫面('schedule');
  };

  const goToLatestHistory = () => {
    setLoadLatestOnEnter(true);
    設定畫面('schedule');
  };

  如果 (screen === 'schedule') {
    返回 （
      <ScheduleView
        changeScreen={setScreen}
        colors={colors}
        setColors={setColors}
        customHolidays={customHolidays}
        setCustomHolidays={setCustomHolidays}
        specialWorkdays={specialWorkdays}
        設定特殊工作日={setSpecialWorkdays}
        medicalCalendarAdjustments={medicalCalendarAdjustments}
        setMedicalCalendarAdjustments={setMedicalCalendarAdjustments}
        staffingConfig={staffingConfig}
        setStaffingConfig={setStaffingConfig}
        uiSettings={uiSettings}
        setUiSettings={setUiSettings}
        loadLatestOnEnter={loadLatestOnEnter}
        onLatestLoaded={() => setLoadLatestOnEnter(false)}
      />
    ）；
  }

  如果（螢幕 === '設定'）{
    返回 （
      <設定視圖
        changeScreen={setScreen}
        colors={colors}
        setColors={setColors}
        customHolidays={customHolidays}
        setCustomHolidays={setCustomHolidays}
        specialWorkdays={specialWorkdays}
        設定特殊工作日={setSpecialWorkdays}
        medicalCalendarAdjustments={medicalCalendarAdjustments}
        setMedicalCalendarAdjustments={setMedicalCalendarAdjustments}
        staffingConfig={staffingConfig}
        setStaffingConfig={setStaffingConfig}
        uiSettings={uiSettings}
        setUiSettings={setUiSettings}
      />
    ）；
  }

  返回 （
    <EntryView
      changeScreen={(target) => {
        如果 (target === 'schedule') goToSchedule();
        否則 setScreen(target);
      }}
      goToLatestHistory={goToLatestHistory}
    />
  ）；
}
