export const getCellCodeFromSnapshotService = ({
  snapshot = {},
  staffId,
  dateStr,
  currentMonthKey = '',
  getContextCellCode = () => ''
} = {}) => {
  const staffSnapshot = snapshot?.[staffId];
  if (staffSnapshot && Object.prototype.hasOwnProperty.call(staffSnapshot, dateStr)) {
    const cellData = staffSnapshot?.[dateStr];
    if (!cellData) return '';
    return typeof cellData === 'object' && cellData !== null ? (cellData.value || '') : (cellData || '');
  }

  const targetMonthKey = String(dateStr || '').slice(0, 7);
  if (!targetMonthKey || targetMonthKey === currentMonthKey) return '';
  return getContextCellCode(staffId, dateStr, { snapshot }) || '';
};

const countConsecutiveWorkDaysBeforeInSnapshotService = ({
  snapshot = {},
  staffId,
  dateStr,
  parseDateKey = (value) => new Date(value),
  formatDateKey = (date) => String(date),
  addDays = (date, amount) => date,
  getCellCodeFromSnapshot = () => '',
  isShiftCode = () => false
} = {}) => {
  let count = 0;
  let cursor = addDays(parseDateKey(dateStr), -1);
  while (true) {
    const key = formatDateKey(cursor);
    const code = getCellCodeFromSnapshot(snapshot, staffId, key);
    if (!isShiftCode(code)) break;
    count += 1;
    cursor = addDays(cursor, -1);
  }
  return count;
};

export const evaluateRuleWarningForCellInSnapshotService = ({
  snapshot = {},
  staff,
  dateStr,
  smartRules = {},
  isConfiguredLeaveCode = () => false,
  isShiftCode = () => false,
  parseDateKey = (value) => new Date(value),
  formatDateKey = (date) => String(date),
  addDays = (date, amount) => date,
  getCellCodeFromSnapshot = () => ''
} = {}) => {
  if (!staff || !dateStr) return [];
  const shiftCode = getCellCodeFromSnapshot(snapshot, staff.id, dateStr);
  if (!shiftCode || isConfiguredLeaveCode(shiftCode) || !isShiftCode(shiftCode)) return [];

  const reasons = [];
  const prevKey = formatDateKey(addDays(parseDateKey(dateStr), -1));
  const prevCode = getCellCodeFromSnapshot(snapshot, staff.id, prevKey);
  const disallowed = smartRules?.disallowedNextShiftMap?.[prevCode] || [];
  if (disallowed.includes(shiftCode)) {
    reasons.push(`${prevCode} 後不可接 ${shiftCode}`);
  }

  const consecutiveBefore = countConsecutiveWorkDaysBeforeInSnapshotService({
    snapshot,
    staffId: staff.id,
    dateStr,
    parseDateKey,
    formatDateKey,
    addDays,
    getCellCodeFromSnapshot,
    isShiftCode
  });
  if (consecutiveBefore + 1 > Number(smartRules?.maxConsecutiveWorkDays || 0)) {
    reasons.push(`連續上班不可超過 ${smartRules.maxConsecutiveWorkDays} 天`);
  }

  if (staff.pregnant && (smartRules?.pregnancyRestrictedShifts || []).includes(shiftCode)) {
    reasons.push('懷孕標記人員不可排 N / 夜8-8');
  }

  return reasons;
};

export const refreshRuleWarningsForCellsService = ({
  cells = [],
  snapshot = {},
  schedule = {},
  staffs = [],
  makeCellKey = (staffId, dateStr) => `${staffId}__${dateStr}`,
  setCellRuleWarnings = () => {},
  evaluateRuleWarningForCellInSnapshot = () => []
} = {}) => {
  const normalizedCells = Array.isArray(cells)
    ? Array.from(new Map(cells.filter((cell) => cell?.staffId && cell?.dateStr).map((cell) => [makeCellKey(cell.staffId, cell.dateStr), cell])).values())
    : [];

  if (normalizedCells.length === 0) return;

  setCellRuleWarnings((prev) => {
    const next = { ...prev };
    normalizedCells.forEach(({ staffId, dateStr }) => {
      const staff = staffs.find((item) => item.id === staffId);
      const reasons = evaluateRuleWarningForCellInSnapshot(snapshot || schedule, staff, dateStr);
      const cellKey = makeCellKey(staffId, dateStr);
      if (reasons.length > 0) next[cellKey] = reasons[0];
      else delete next[cellKey];
    });
    return next;
  });
};

export const expandCellsForRuleRecheckService = ({
  cells = [],
  daysInMonth = []
} = {}) => {
  const staffIds = Array.from(new Set(
    (Array.isArray(cells) ? cells : [])
      .map((cell) => cell?.staffId)
      .filter(Boolean)
  ));

  if (staffIds.length === 0) return [];

  const expanded = [];
  staffIds.forEach((staffId) => {
    (daysInMonth || []).forEach((day) => {
      expanded.push({ staffId, dateStr: day.date });
    });
  });

  return expanded;
};

export const scanScheduleRuleViolationsService = ({
  schedulesSource = {},
  options = {},
  smartRules = {},
  isShiftCode = () => false,
  parseDateKey = (value) => new Date(value),
  formatDateKey = (date) => String(date),
  addDays = (date, amount) => date
} = {}) => {
  const monthKeys = Object.keys(schedulesSource || {}).sort();
  const targetMonthKeys = new Set(Array.isArray(options?.targetMonthKeys) ? options.targetMonthKeys : monthKeys);
  const byName = new Map();

  monthKeys.forEach((monthKey) => {
    const monthState = schedulesSource?.[monthKey];
    const monthStaffs = Array.isArray(monthState?.staffs) ? monthState.staffs : [];
    const monthSchedule = monthState?.scheduleData || monthState?.schedule || {};
    monthStaffs.forEach((staff) => {
      const nameKey = String(staff?.name || '').trim();
      if (!nameKey) return;
      const staffSchedule = monthSchedule?.[staff.id] || {};
      if (!byName.has(nameKey)) byName.set(nameKey, { staff, cells: {} });
      const ref = byName.get(nameKey);
      Object.entries(staffSchedule).forEach(([dateStr, cell]) => {
        ref.cells[dateStr] = cell;
      });
    });
  });

  const violations = [];
  byName.forEach((ref, nameKey) => {
    const dateKeys = Object.keys(ref.cells || {}).sort();
    const getCode = (dateStr) => {
      const cell = ref.cells?.[dateStr];
      return typeof cell === 'object' && cell !== null ? (cell.value || '') : String(cell || '').trim();
    };

    dateKeys.forEach((dateStr) => {
      if (!targetMonthKeys.has(String(dateStr).slice(0, 7))) return;
      const code = getCode(dateStr);
      if (!isShiftCode(code)) return;

      const prevKey = formatDateKey(addDays(parseDateKey(dateStr), -1));
      const prevCode = getCode(prevKey);
      const disallowed = smartRules?.disallowedNextShiftMap?.[prevCode] || [];
      if (disallowed.includes(code)) {
        violations.push({
          staffName: nameKey,
          dateStr,
          code,
          reason: `${prevCode} 後不可接 ${code}`
        });
      }

      let consecutiveBefore = 0;
      let cursor = addDays(parseDateKey(dateStr), -1);
      while (true) {
        const cursorKey = formatDateKey(cursor);
        const cursorCode = getCode(cursorKey);
        if (!isShiftCode(cursorCode)) break;
        consecutiveBefore += 1;
        cursor = addDays(cursor, -1);
      }
      if (consecutiveBefore + 1 > Number(smartRules?.maxConsecutiveWorkDays || 0)) {
        violations.push({
          staffName: nameKey,
          dateStr,
          code,
          reason: `連續上班不可超過 ${smartRules.maxConsecutiveWorkDays} 天`
        });
      }
    });
  });

  return violations;
};
