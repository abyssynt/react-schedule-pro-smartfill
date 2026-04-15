import { SHIFT_GROUPS } from './scheduleData';

export const buildMonthKey = (year, month) => `${year}-${String(month).padStart(2, '0')}`;

export const normalizeStaffGroup = (staffList = []) => {
  if (!Array.isArray(staffList) || staffList.length === 0) return [];

  const fallbackGroups = [
    '白班', '白班', '白班', '白班', '白班',
    '小夜', '小夜', '小夜', '小夜', '小夜',
    '大夜', '大夜', '大夜', '大夜', '大夜'
  ];

  return staffList.map((staff, index) => ({
    ...staff,
    pregnant: Boolean(staff?.pregnant),
    group: SHIFT_GROUPS.includes(staff?.group) ? staff.group : (fallbackGroups[index] || '白班')
  }));
};

export const createBlankMonthStaffs = (targetYear, targetMonth) => {
  const monthKey = buildMonthKey(targetYear, targetMonth);
  const blankStaffs = Array.from({ length: 15 }, (_, index) => ({
    id: `blank_${monthKey}_${index + 1}`,
    name: '新成員',
    group: SHIFT_GROUPS[Math.floor(index / 5)] || '白班',
    pregnant: false
  }));
  return normalizeStaffGroup(blankStaffs);
};

export const createBlankMonthState = (targetYear, targetMonth) => {
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


const normalizeStoredScheduleCellValue = (rawValue = '', customLeaveCodes = [], helpers = {}) => {
  const {
    DICT,
    getAllShiftCodes,
    normalizeManualShiftCode,
    normalizeImportedHalfWidth
  } = helpers;

  const trimmedValue = String(rawValue ?? '').trim();
  if (!trimmedValue) return '';

  const allShiftCodes = typeof getAllShiftCodes === 'function' ? getAllShiftCodes() : [];
  const leaveCodes = Array.isArray(DICT?.LEAVES) ? DICT.LEAVES : [];
  const exactKnownCode = Array.from(new Set([...(allShiftCodes || []), ...leaveCodes, ...(customLeaveCodes || [])])).find((code) => {
    const normalizedCode = String(code || '').trim().replace(/\s+/g, '');
    const normalizedValue = String(trimmedValue || '').trim().replace(/\s+/g, '');
    return normalizedCode === normalizedValue;
  });
  if (exactKnownCode) return exactKnownCode;

  if (typeof normalizeManualShiftCode === 'function') {
    const { normalized, isValid } = normalizeManualShiftCode(trimmedValue, customLeaveCodes || []);
    if (isValid) return normalized;
  }

  return typeof normalizeImportedHalfWidth === 'function'
    ? normalizeImportedHalfWidth(trimmedValue)
    : trimmedValue;
};

export const reconcileScheduleDataMap = (scheduleData = {}, customLeaveCodes = [], helpers = {}) => {
  const nextScheduleData = {};
  Object.entries(scheduleData || {}).forEach(([staffId, dateMap]) => {
    if (!dateMap || typeof dateMap !== 'object') {
      nextScheduleData[staffId] = {};
      return;
    }
    nextScheduleData[staffId] = {};
    Object.entries(dateMap || {}).forEach(([dateStr, cell]) => {
      if (!cell) {
        nextScheduleData[staffId][dateStr] = null;
        return;
      }
      const rawValue = typeof cell === 'object' && cell !== null ? (cell.value || '') : String(cell || '');
      const normalizedValue = normalizeStoredScheduleCellValue(rawValue, customLeaveCodes, helpers);
      if (!normalizedValue) {
        nextScheduleData[staffId][dateStr] = null;
        return;
      }

      const allShiftCodes = typeof helpers.getAllShiftCodes === 'function' ? helpers.getAllShiftCodes() : [];
      const leaveCodes = Array.isArray(helpers.DICT?.LEAVES) ? helpers.DICT.LEAVES : [];
      const isKnownCode = allShiftCodes.includes(normalizedValue) || leaveCodes.includes(normalizedValue) || (customLeaveCodes || []).includes(normalizedValue);

      const existingMeta = typeof cell === 'object' && cell !== null ? { ...cell } : {};
      delete existingMeta.value;
      delete existingMeta.isUnknownCode;
      nextScheduleData[staffId][dateStr] = {
        ...existingMeta,
        value: normalizedValue,
        ...(isKnownCode ? {} : { isUnknownCode: true })
      };
    });
  });
  return nextScheduleData;
};

export const reconcileMonthStateCollections = (collection = {}, customLeaveCodes = [], helpers = {}) => {
  const nextCollection = {};
  Object.entries(collection || {}).forEach(([monthKey, monthState]) => {
    nextCollection[monthKey] = {
      ...(monthState || {}),
      scheduleData: reconcileScheduleDataMap(monthState?.scheduleData || monthState?.schedule || {}, customLeaveCodes, helpers)
    };
  });
  return nextCollection;
};
