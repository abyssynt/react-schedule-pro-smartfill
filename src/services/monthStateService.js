export const createBlankScheduleForStaffs = (staffList = []) => {
  return (Array.isArray(staffList) ? staffList : []).reduce((acc, staff) => {
    if (staff?.id) acc[staff.id] = {};
    return acc;
  }, {});
};

export const resolveRulesText = (...candidates) => {
  for (const candidate of candidates) {
    if (typeof candidate === 'string' && candidate.trim() !== '') return candidate;
  }
  for (const candidate of candidates) {
    if (typeof candidate === 'string') return candidate;
  }
  return '';
};

export const rebuildLegacyScheduleData = (monthKey, legacyScheduleByDay = {}) => {
  return Object.fromEntries(
    Object.entries(legacyScheduleByDay || {}).map(([staffId, dayMap]) => [
      staffId,
      Object.fromEntries(
        Object.entries(dayMap || {}).map(([day, cell]) => {
          const dateKey = `${monthKey}-${String(Number(day)).padStart(2, '0')}`;
          return [dateKey, cell];
        })
      )
    ])
  );
};

export const buildMonthStatePayload = ({
  targetYear,
  targetMonth,
  schedulesSource = {},
  preScheduleMonthlySchedules = {},
  schedulingRulesText = '',
  readSchedulingRulesTextFromLocalSettings = () => '',
  normalizeStaffGroup = (staffs = []) => staffs,
  createBlankMonthState = () => ({
    staffs: [],
    schedule: {},
    customColumnValues: {},
    schedulingRulesText: ''
  })
} = {}) => {
  const monthKey = `${targetYear}-${String(targetMonth).padStart(2, '0')}`;
  const monthData = schedulesSource?.[monthKey];
  const preMonthData = preScheduleMonthlySchedules?.[monthKey];
  const currentRulesText = typeof schedulingRulesText === 'string' ? schedulingRulesText : '';
  const savedLocalRulesText = readSchedulingRulesTextFromLocalSettings();

  if (monthData) {
    const normalizedMonthStaffs = normalizeStaffGroup(monthData.staffs || []);
    const legacyScheduleByDay = monthData.scheduleByDay || {};
    const rebuiltScheduleData = monthData.scheduleData || rebuildLegacyScheduleData(monthKey, legacyScheduleByDay);

    return {
      staffs: normalizedMonthStaffs,
      schedule: rebuiltScheduleData || createBlankScheduleForStaffs(normalizedMonthStaffs),
      customColumnValues: monthData.customColumnValues || {},
      schedulingRulesText: resolveRulesText(
        monthData.schedulingRulesText,
        currentRulesText,
        savedLocalRulesText
      )
    };
  }

  if (preMonthData) {
    const normalizedPreMonthStaffs = normalizeStaffGroup(preMonthData.staffs || []);
    return {
      staffs: normalizedPreMonthStaffs,
      schedule: createBlankScheduleForStaffs(normalizedPreMonthStaffs),
      customColumnValues: preMonthData.customColumnValues || {},
      schedulingRulesText: resolveRulesText(
        preMonthData.schedulingRulesText,
        currentRulesText,
        savedLocalRulesText
      )
    };
  }

  const blankMonthState = createBlankMonthState(targetYear, targetMonth);
  return {
    staffs: blankMonthState.staffs || [],
    schedule: blankMonthState.schedule || {},
    customColumnValues: blankMonthState.customColumnValues || {},
    schedulingRulesText: resolveRulesText(
      currentRulesText,
      savedLocalRulesText,
      blankMonthState.schedulingRulesText
    )
  };
};
