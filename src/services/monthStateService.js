export const createBlankScheduleForStaffs = (staffList = []) => {
  return (Array.isArray(staffList) ? staffList : []).reduce((acc, staff) => {
    acc[staff.id] = {};
    return acc;
  }, {});
};

export const createLoadMonthState = (deps = {}) => {
  const {
    buildMonthKey = (year, month) => `${year}-${String(month).padStart(2, '0')}`,
    getPreScheduleMonthlySchedules = () => ({}),
    getSchedulingRulesText = () => '',
    readSchedulingRulesTextFromLocalSettings = () => '',
    normalizeStaffGroup = (staffs = []) => staffs,
    createBlankMonthState = () => ({ staffs: [], schedule: {}, customColumnValues: {}, schedulingRulesText: '' }),
    setStaffs = () => {},
    setSchedule = () => {},
    setCustomColumnValues = () => {},
    setSchedulingRulesText = () => {},
    setSelectedGridCell = () => {},
    setRangeSelection = () => {},
    setSelectionAnchor = () => {},
    setCellDrafts = () => {},
    setInvalidCellKeys = () => {},
    setCellRuleWarnings = () => {},
    setKeyInputBuffer = () => {},
    clearInputAssist = () => {},
    setEditingStaffId = () => {},
    setEditingNameDraft = () => {},
    setDraggingStaffId = () => {},
    setDragOverTarget = () => {},
    monthLoadSkipRef = { current: false },
    createBlankScheduleForStaffs = (staffList = []) => ({})
  } = deps;

  return (targetYear, targetMonth, schedulesSource = {}) => {
    const monthKey = buildMonthKey(targetYear, targetMonth);
    const monthData = schedulesSource?.[monthKey];
    const preMonthData = getPreScheduleMonthlySchedules()?.[monthKey];
    const currentRulesText = typeof getSchedulingRulesText() === 'string' ? getSchedulingRulesText() : '';
    const savedLocalRulesText = readSchedulingRulesTextFromLocalSettings();

    const resolveRulesText = (...candidates) => {
      for (const candidate of candidates) {
        if (typeof candidate === 'string' && candidate.trim() !== '') return candidate;
      }
      for (const candidate of candidates) {
        if (typeof candidate === 'string') return candidate;
      }
      return '';
    };

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
      setSchedulingRulesText(resolveRulesText(monthData.schedulingRulesText, currentRulesText, savedLocalRulesText));
    } else if (preMonthData) {
      const normalizedPreMonthStaffs = normalizeStaffGroup(preMonthData.staffs || []);
      setStaffs(normalizedPreMonthStaffs);
      setSchedule(createBlankScheduleForStaffs(normalizedPreMonthStaffs));
      setCustomColumnValues(preMonthData.customColumnValues || {});
      setSchedulingRulesText(resolveRulesText(preMonthData.schedulingRulesText, currentRulesText, savedLocalRulesText));
    } else {
      const blankMonthState = createBlankMonthState(targetYear, targetMonth);
      setStaffs(blankMonthState.staffs);
      setSchedule(blankMonthState.schedule);
      setCustomColumnValues(blankMonthState.customColumnValues);
      setSchedulingRulesText(resolveRulesText(currentRulesText, savedLocalRulesText, blankMonthState.schedulingRulesText));
    }

    setSelectedGridCell(null);
    setRangeSelection(null);
    setSelectionAnchor(null);
    setCellDrafts({});
    setInvalidCellKeys({});
    setCellRuleWarnings({});
    setKeyInputBuffer('');
    clearInputAssist();
    setEditingStaffId(null);
    setEditingNameDraft('');
    setDraggingStaffId(null);
    setDragOverTarget(null);

    setTimeout(() => {
      monthLoadSkipRef.current = false;
    }, 0);
  };
};
