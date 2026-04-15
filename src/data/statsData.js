export const buildEmptyStaffStats = (mergedLeaveCodes = []) => ({
  work: 0,
  holidayLeave: 0,
  totalLeave: 0,
  leaveDetails: Object.fromEntries((mergedLeaveCodes || []).map((leaveCode) => [leaveCode, 0]))
});

export const buildStaffStatsMap = ({
  staffs = [],
  schedule = {},
  daysInMonth = [],
  mergedLeaveCodes = [],
  getAllShiftCodes = () => [],
  isConfiguredLeaveCode = () => false,
  getCodePrefix = (code) => code
} = {}) => {
  const next = {};
  (staffs || []).forEach((staff) => {
    const stats = buildEmptyStaffStats(mergedLeaveCodes);
    const mySchedule = schedule?.[staff.id] || {};
    (daysInMonth || []).forEach((dayInfo) => {
      const cellData = mySchedule[dayInfo.date];
      const code = typeof cellData === 'object' && cellData !== null ? cellData.value : cellData;
      if (!code) return;
      if ((getAllShiftCodes() || []).includes(code)) stats.work += 1;
      if (isConfiguredLeaveCode(code)) {
        stats.totalLeave += 1;
        const leavePrefix = getCodePrefix(code);
        if (stats.leaveDetails[leavePrefix] !== undefined) stats.leaveDetails[leavePrefix] += 1;
        if (dayInfo.isWeekend || dayInfo.isHoliday) stats.holidayLeave += 1;
      }
    });
    next[staff.id] = stats;
  });
  return next;
};

export const buildDailyStatsMap = ({
  staffs = [],
  schedule = {},
  daysInMonth = [],
  getShiftGroupByCode = () => '',
  isConfiguredLeaveCode = () => false
} = {}) => {
  const next = {};
  (daysInMonth || []).forEach((dayInfo) => {
    const stats = { D: 0, E: 0, N: 0, totalLeave: 0 };
    (staffs || []).forEach((staff) => {
      const cellData = schedule?.[staff.id]?.[dayInfo.date];
      const code = typeof cellData === 'object' && cellData !== null ? cellData.value : cellData;
      if (!code) return;
      const shiftGroup = getShiftGroupByCode(code);
      if (shiftGroup === '白班') stats.D += 1;
      else if (shiftGroup === '小夜') stats.E += 1;
      else if (shiftGroup === '大夜') stats.N += 1;
      else if (isConfiguredLeaveCode(code)) stats.totalLeave += 1;
    });
    next[dayInfo.date] = stats;
  });
  return next;
};

export const buildRequiredCountMap = (daysInMonth = [], staffingConfig = {}, getRequiredStaffingBucketByDay = () => 'weekday') => {
  const next = {};
  (daysInMonth || []).forEach((dayInfo) => {
    const bucket = getRequiredStaffingBucketByDay(dayInfo);
    next[dayInfo.date] = {
      D: Number(staffingConfig?.requiredStaffing?.[bucket]?.white || 0),
      E: Number(staffingConfig?.requiredStaffing?.[bucket]?.evening || 0),
      N: Number(staffingConfig?.requiredStaffing?.[bucket]?.night || 0)
    };
  });
  return next;
};

export const getDemandHighlightStyle = (dateStr, rowKey, actualCount, requiredCountMap = {}, demandOverColor = '') => {
  if (!['D', 'E', 'N'].includes(rowKey)) return {};
  const requiredCount = requiredCountMap?.[dateStr]?.[rowKey] ?? null;
  if (requiredCount === null) return {};
  if (actualCount > requiredCount) return { backgroundColor: demandOverColor };
  return {};
};
