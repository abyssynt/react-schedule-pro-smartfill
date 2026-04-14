import { DICT, SHIFT_GROUPS } from './scheduleData';
import { buildMonthKey, buildExistingStaffGroupLookup } from './monthScheduleData';
import {
  normalizeImportedShiftCode,
  getImportedRawNumberedLeaveValue,
  isConfiguredImportedLeaveCode
} from './importCodeData';

export const extractYearMonthCandidates = (...sources) => {
  const fullPatterns = [
    /(\d{4})\s*年\s*(\d{1,2})\s*月/,
    /(\d{4})[\/_\-.](\d{1,2})/
  ];
  const monthOnlyPatterns = [
    /(\d{1,2})\s*月/
  ];

  let monthOnlyCandidate = { year: null, month: null };

  for (const source of sources) {
    const text = String(source || '').trim();
    if (!text) continue;

    for (const pattern of fullPatterns) {
      const match = text.match(pattern);
      if (!match) continue;
      const year = Number(match[1]);
      const month = Number(match[2]);
      if (year >= 1900 && month >= 1 && month <= 12) {
        return { year, month };
      }
    }

    if (!monthOnlyCandidate.month) {
      for (const pattern of monthOnlyPatterns) {
        const match = text.match(pattern);
        if (!match) continue;
        const month = Number(match[1]);
        if (month >= 1 && month <= 12) {
          monthOnlyCandidate = { year: null, month };
          break;
        }
      }
    }
  }

  return monthOnlyCandidate;
};

export const detectImportedDayNumber = (label = '') => {
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

export const inferImportedGroupFromCodes = (dayMap = {}, helpers = {}) => {
  const { getShiftGroupByCode = () => '' } = helpers;
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

export const parseImportedWorksheet = (
  { rows, sheetName, fileName, fallbackYear, customLeaveCodes = [], importMode = 'schedule', existingStaffGroupLookup = { byMonth: {}, fallback: {} } },
  helpers = {}
) => {
  const {
    getAllShiftCodes = () => [],
    getShiftGroupByCode = () => '',
    getCodePrefix = (value) => value
  } = helpers;

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
  const knownCodes = new Set([...getAllShiftCodes(), ...DICT.LEAVES, ...(customLeaveCodes || [])]);

  const scanTexts = [];
  const maxRowsToScan = Math.min(rows.length, 10);
  for (let r = 0; r < maxRowsToScan; r += 1) {
    const row = Array.isArray(rows[r]) ? rows[r] : [];
    for (let c = 0; c < row.length; c += 1) {
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
  const monthGroupLookup = existingStaffGroupLookup?.byMonth?.[monthKey] || {};
  const fallbackGroupLookup = existingStaffGroupLookup?.fallback || {};

  const importedStaffs = [];
  const importedSchedule = {};
  const invalidMessages = [];
  const unknownCodes = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
    const rawName = String(row[nameColumnIndex] ?? '').trim();
    const rawGroup = groupColumnIndex === -1 ? '' : String(row[groupColumnIndex] ?? '').trim();

    const hasAnyContent = row.some((value) => String(value ?? '').trim() !== '');
    if (!hasAnyContent || !rawName) continue;

    const rowNumber = rowIndex + 1;
    const hasAnyDayContent = dayColumnPairs.some(({ colNumber }) => String(row[colNumber] ?? '').trim() !== '');

    const staffId = `import_${Date.now()}_${sheetName}_${rowNumber}`;
    importedSchedule[staffId] = {};

    dayColumnPairs.forEach(({ day, colNumber }) => {
      const rawValue = String(row[colNumber] ?? '').trim();
      if (!rawValue) return;

      const normalizedCode = normalizeImportedShiftCode(rawValue, getAllShiftCodes);
      const rawImportedValue = getImportedRawNumberedLeaveValue(rawValue);
      const isKnownCode = knownCodes.has(normalizedCode);

      importedSchedule[staffId][day] = {
        value: normalizedCode,
        source: 'manual',
        ...(rawImportedValue ? { rawImportedValue } : {}),
        ...(!isKnownCode ? { isUnknownCode: true } : {})
      };

      if (!isKnownCode) {
        unknownCodes.push(normalizedCode);
      }
    });

    const importedCodes = Object.values(importedSchedule[staffId] || {}).map((cell) => {
      if (!cell) return '';
      return typeof cell === 'object' && cell !== null ? (cell.value || '') : String(cell || '').trim();
    }).filter(Boolean);
    const hasShiftCode = importedCodes.some(code => !isConfiguredImportedLeaveCode(code, customLeaveCodes, getCodePrefix));
    const hasLeaveCode = importedCodes.some(code => isConfiguredImportedLeaveCode(code, customLeaveCodes, getCodePrefix));

    let normalizedGroup = '';
    if (rawGroup) {
      if (validGroups.has(rawGroup)) {
        normalizedGroup = rawGroup;
      } else {
        invalidMessages.push(`工作表「${sheetName}」第 ${rowNumber} 列「${rawName}」的班別群組不是白班／小夜／大夜，將改用其他規則判定`);
      }
    }

    if (!normalizedGroup) {
      normalizedGroup = monthGroupLookup[rawName] || fallbackGroupLookup[rawName] || '';
    }

    if (!normalizedGroup) {
      normalizedGroup = inferImportedGroupFromCodes(importedSchedule[staffId], { getShiftGroupByCode });
    }

    if (!normalizedGroup) {
      if (!hasAnyDayContent) {
        normalizedGroup = monthGroupLookup[rawName] || fallbackGroupLookup[rawName] || '白班';
        invalidMessages.push(`工作表「${sheetName}」第 ${rowNumber} 列「${rawName}」日期區沒有排班內容，已先保留人員名單`);
      } else if (importMode === 'preSchedule') {
        normalizedGroup = '白班';
        if (hasLeaveCode && !hasShiftCode) {
          invalidMessages.push(`工作表「${sheetName}」第 ${rowNumber} 列「${rawName}」只有預假代碼，且無可對照群組，已先歸入白班保留資料`);
        } else {
          invalidMessages.push(`工作表「${sheetName}」第 ${rowNumber} 列「${rawName}」無法判定群組，已依預班規則先歸入白班`);
        }
      } else if (hasLeaveCode && !hasShiftCode) {
        normalizedGroup = '白班';
        invalidMessages.push(`工作表「${sheetName}」第 ${rowNumber} 列「${rawName}」只有休假代碼，已先歸入白班保留資料`);
      } else {
        normalizedGroup = monthGroupLookup[rawName] || fallbackGroupLookup[rawName] || '白班';
        invalidMessages.push(`工作表「${sheetName}」第 ${rowNumber} 列「${rawName}」沒有可判定群組的班別代碼，已先歸入${normalizedGroup}`);
      }
    }

    if (!hasAnyDayContent && importMode === 'preSchedule') {
      invalidMessages.push(`工作表「${sheetName}」第 ${rowNumber} 列「${rawName}」日期區沒有預班內容，已先建立人員骨架供後續預班使用`);
    }

    importedStaffs.push({
      id: staffId,
      name: rawName,
      group: normalizedGroup,
      pregnant: false
    });
  }

  if (importedStaffs.length === 0) return null;

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
    unknownCodes: Array.from(new Set(unknownCodes)).sort(),
    importMeta: {
      sourceType: importMode === 'preSchedule' ? 'preScheduleExcel' : 'excel',
      sourceFiles: [fileName],
      sourceSheets: [sheetName],
      importedAt: new Date().toISOString(),
      lastUpdatedAt: new Date().toISOString(),
      importMode
    }
  };
};

export const parseImportedExcelFiles = async (files = [], fallbackYear = new Date().getFullYear(), options = {}, helpers = {}) => {
  const { loadSheetJS, getAllShiftCodes = () => [], getShiftGroupByCode = () => '', getCodePrefix = (value) => value } = helpers;
  if (typeof loadSheetJS !== 'function') {
    throw new Error('缺少 loadSheetJS，無法匯入 Excel');
  }

  const XLSX = await loadSheetJS();
  const fileList = Array.from(files || []);
  const monthlySchedules = {};
  const warnings = [];
  const unknownCodes = [];
  let firstMonthKey = '';
  const importMode = options.importMode || 'schedule';
  const existingStaffGroupLookup = buildExistingStaffGroupLookup(options.existingMonthlySchedules || {});

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
        const parsed = parseImportedWorksheet(
          {
            rows,
            sheetName,
            fileName: file.name,
            fallbackYear,
            customLeaveCodes: options.customLeaveCodes || [],
            importMode,
            existingStaffGroupLookup
          },
          { getAllShiftCodes, getShiftGroupByCode, getCodePrefix }
        );

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
        unknownCodes.push(...(parsed.unknownCodes || []));
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
    warnings,
    unknownCodes: Array.from(new Set(unknownCodes.filter(Boolean))).sort(),
    importMode
  };
};
