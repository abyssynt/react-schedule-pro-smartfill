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

export const buildExistingStaffGroupLookup = (monthlySchedules = {}) => {
  const byMonth = {};
  const fallback = {};

  Object.entries(monthlySchedules || {}).forEach(([monthKey, monthState]) => {
    const monthLookup = {};
    const monthStaffs = Array.isArray(monthState?.staffs) ? monthState.staffs : [];
    monthStaffs.forEach((staff) => {
      const nameKey = String(staff?.name || '').trim();
      const group = staff?.group || '';
      if (!nameKey || !group) return;
      if (!monthLookup[nameKey]) monthLookup[nameKey] = group;
      if (!fallback[nameKey]) fallback[nameKey] = group;
    });
    byMonth[monthKey] = monthLookup;
  });

  return { byMonth, fallback };
};

export const mergeImportedMonthStates = (baseState = null, incomingState = null) => {
  if (!baseState) return incomingState;
  if (!incomingState) return baseState;

  const mergedStaffs = [];
  const mergedSchedule = {};
  const staffKeyToId = new Map();
  const signatureToKey = new Map();

  const registerStaffs = (monthState, priority = 'base') => {
    const staffList = Array.isArray(monthState?.staffs) ? monthState.staffs : [];
    const scheduleData = monthState?.scheduleData || monthState?.schedule || {};

    staffList.forEach((staff) => {
      const name = String(staff?.name || '').trim();
      const group = staff?.group || '白班';
      if (!name) return;
      const signature = `${name}__${group}`;
      const fallbackSignature = `${name}__*`;
      const matchedSignature = signatureToKey.has(signature)
        ? signature
        : (signatureToKey.has(fallbackSignature) ? signatureToKey.get(fallbackSignature) : null);

      const staffKey = matchedSignature || signature;
      let existingId = staffKeyToId.get(staffKey);

      if (!existingId) {
        existingId = staff.id || `${priority}_${mergedStaffs.length + 1}`;
        staffKeyToId.set(staffKey, existingId);
        signatureToKey.set(signature, staffKey);
        signatureToKey.set(fallbackSignature, staffKey);
        mergedStaffs.push({
          ...staff,
          id: existingId,
          name,
          group
        });
        mergedSchedule[existingId] = { ...(scheduleData?.[staff.id] || {}) };
        return;
      }

      const existingIndex = mergedStaffs.findIndex((item) => item.id === existingId);
      if (existingIndex !== -1 && priority === 'incoming') {
        mergedStaffs[existingIndex] = {
          ...mergedStaffs[existingIndex],
          ...staff,
          id: existingId,
          name,
          group
        };
      }

      mergedSchedule[existingId] = {
        ...(mergedSchedule[existingId] || {}),
        ...(scheduleData?.[staff.id] || {})
      };
    });
  };

  registerStaffs(baseState, 'base');
  registerStaffs(incomingState, 'incoming');

  const baseImportMeta = baseState?.importMeta || {};
  const incomingImportMeta = incomingState?.importMeta || {};

  return {
    ...baseState,
    ...incomingState,
    staffs: mergedStaffs,
    scheduleData: mergedSchedule,
    importMeta: {
      ...baseImportMeta,
      ...incomingImportMeta,
      sourceType: incomingImportMeta.sourceType || baseImportMeta.sourceType || 'preScheduleExcel',
      sourceFiles: Array.from(new Set([...(baseImportMeta.sourceFiles || []), ...(incomingImportMeta.sourceFiles || [])])),
      sourceSheets: Array.from(new Set([...(baseImportMeta.sourceSheets || []), ...(incomingImportMeta.sourceSheets || [])])),
      importedAt: baseImportMeta.importedAt || incomingImportMeta.importedAt || new Date().toISOString(),
      lastUpdatedAt: new Date().toISOString()
    },
    warnings: [...(baseState?.warnings || []), ...(incomingState?.warnings || [])]
  };
};
