export const buildExportNumberedValueMap = (staffOrId, deps = {}) => {
  const {
    daysInMonth = [],
    getExportCellPresentation = () => ({ displayValue: '' }),
    getRawImportedLeaveSeed = () => 0,
    isFourWeekCycleEndDate = () => false,
  } = deps;

  const valueMap = {};
  let leaveCounters = { 例: 0, 休: 0 };
  let isFirstSegment = true;
  let firstSegmentSeed = {
    例: getRawImportedLeaveSeed(staffOrId, '例'),
    休: getRawImportedLeaveSeed(staffOrId, '休')
  };
  let firstSegmentCarryUsed = { 例: false, 休: false };

  daysInMonth.forEach((dayInfo, index) => {
    const presentation = getExportCellPresentation(staffOrId, dayInfo);
    const displayValue = presentation.displayValue || '';

    if (displayValue === '例' || displayValue === '休') {
      const leaveType = displayValue;

      if (isFirstSegment && Number(firstSegmentSeed[leaveType] || 0) > 0) {
        if (!firstSegmentCarryUsed[leaveType]) {
          const continuedCount = Math.min(Number(firstSegmentSeed[leaveType] || 0) + 1, 4);
          valueMap[dayInfo.date] = `${leaveType}${continuedCount}`;
          firstSegmentCarryUsed[leaveType] = true;
        } else {
          valueMap[dayInfo.date] = leaveType;
        }
      } else {
        const nextCount = Number(leaveCounters[leaveType] || 0) + 1;
        leaveCounters[leaveType] = nextCount;
        valueMap[dayInfo.date] = nextCount <= 4 ? `${leaveType}${nextCount}` : leaveType;
      }
    } else {
      valueMap[dayInfo.date] = displayValue;
    }

    if (isFourWeekCycleEndDate(dayInfo.date)) {
      if (index < daysInMonth.length - 1) {
        leaveCounters = { 例: 0, 休: 0 };
        isFirstSegment = false;
        firstSegmentSeed = { 例: 0, 休: 0 };
        firstSegmentCarryUsed = { 例: false, 休: false };
      }
    }
  });

  return valueMap;
};

export const getExportNumberedValue = (staffOrId, dateStr, deps = {}) => {
  if (!dateStr) return '';
  const valueMap = buildExportNumberedValueMap(staffOrId, deps);
  return valueMap[dateStr] || '';
};

export const buildExportStaffStats = (staffId, deps = {}) => {
  const {
    mergedLeaveCodes = [],
    daysInMonth = [],
    getExportNumberedValue = () => '',
    getAllShiftCodes = () => [],
    isConfiguredLeaveCode = () => false,
    getCodePrefix = (code) => code
  } = deps;

  const stats = {
    work: 0,
    holidayLeave: 0,
    totalLeave: 0,
    leaveDetails: Object.fromEntries(mergedLeaveCodes.map((leaveCode) => [leaveCode, 0]))
  };

  daysInMonth.forEach((dayInfo) => {
    const displayValue = getExportNumberedValue(staffId, dayInfo.date);
    if (!displayValue) return;
    if (getAllShiftCodes().includes(displayValue)) stats.work += 1;
    if (isConfiguredLeaveCode(displayValue)) {
      stats.totalLeave += 1;
      const leavePrefix = getCodePrefix(displayValue);
      if (stats.leaveDetails[leavePrefix] !== undefined) stats.leaveDetails[leavePrefix] += 1;
      if (dayInfo.isWeekend || dayInfo.isHoliday) stats.holidayLeave += 1;
    }
  });

  return stats;
};

export const buildExportDailyStats = (dateStr, deps = {}) => {
  const {
    staffs = [],
    getExportNumberedValue = () => '',
    getShiftGroupByCode = () => '',
    isConfiguredLeaveCode = () => false
  } = deps;

  const stats = { D: 0, E: 0, N: 0, totalLeave: 0 };
  staffs.forEach((staff) => {
    const displayValue = getExportNumberedValue(staff.id, dateStr);
    if (!displayValue) return;

    const shiftGroup = getShiftGroupByCode(displayValue);
    if (shiftGroup === '白班') stats.D += 1;
    else if (shiftGroup === '小夜') stats.E += 1;
    else if (shiftGroup === '大夜') stats.N += 1;
    else if (isConfiguredLeaveCode(displayValue)) stats.totalLeave += 1;
  });
  return stats;
};

export const formatWordDayCellValue = (value = "") => {
  const text = String(value || '').trim();
  if (/^(例|休)[1-4]$/.test(text)) {
    return `<span class="word-numbered-leave">${text}</span>`;
  }
  return text;
};
