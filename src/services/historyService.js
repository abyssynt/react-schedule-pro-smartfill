export const createSaveToHistory = (deps) => (label, currentSchedule = deps.schedule) => {
  const {
    buildMonthKey,
    year,
    month,
    staffs,
    normalizeStaffGroup,
    monthlySchedules,
    preScheduleMonthlySchedules,
    colors,
    customHolidays,
    specialWorkdays,
    medicalCalendarAdjustments,
    staffingConfig,
    uiSettings,
    customLeaveCodes,
    customWorkShifts,
    customColumns,
    customColumnValues,
    schedulingRulesText,
    setHistoryList,
    STORAGE_KEY
  } = deps;

    const currentMonthKey = buildMonthKey(year, month);
    const normalizedMonthStaffs = normalizeStaffGroup(staffs);
    const mergedMonthlySchedules = {
      ...(monthlySchedules || {}),
      [currentMonthKey]: {
        ...(monthlySchedules?.[currentMonthKey] || {}),
        year,
        month,
        staffs: normalizedMonthStaffs,
        scheduleData: currentSchedule,
        customColumnValues: customColumnValues || {},
        schedulingRulesText: typeof schedulingRulesText === 'string' ? schedulingRulesText : '',
        importMeta: {
          ...(monthlySchedules?.[currentMonthKey]?.importMeta || {}),
          sourceType: monthlySchedules?.[currentMonthKey]?.importMeta?.sourceType || 'manual',
          sourceFiles: monthlySchedules?.[currentMonthKey]?.importMeta?.sourceFiles || [],
          sourceSheets: monthlySchedules?.[currentMonthKey]?.importMeta?.sourceSheets || [],
          importedAt: monthlySchedules?.[currentMonthKey]?.importMeta?.importedAt || new Date().toISOString(),
          lastUpdatedAt: new Date().toISOString()
        }
      }
    };

    const newRecord = {
      id: Date.now(),
      label,
      timestamp: new Date().toLocaleString(),
      state: {
        year,
        month,
        staffs: normalizedMonthStaffs,
        schedule: currentSchedule,
        monthlySchedules: mergedMonthlySchedules,
        preScheduleMonthlySchedules: preScheduleMonthlySchedules || {},
        colors,
        customHolidays,
        specialWorkdays,
        medicalCalendarAdjustments,
        staffingConfig,
        uiSettings,
        customLeaveCodes,
        customWorkShifts,
        customColumns,
        customColumnValues,
        schedulingRulesText
      }
    };

    setHistoryList(prev => {
      const updated = [newRecord, ...prev].slice(0, 10);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
      return updated;
    });
};

export const createLoadHistory = (deps) => (record) => {
  const {
    buildMonthKey,
    normalizeStaffGroup,
    setMonthlySchedules,
    setPreScheduleMonthlySchedules,
    setYear,
    setMonth,
    setCustomHolidays,
    setSpecialWorkdays,
    setMedicalCalendarAdjustments,
    setStaffingConfig,
    setUiSettings,
    setCustomLeaveCodes,
    setCustomWorkShifts,
    setCustomColumns,
    setCustomColumnValues,
    setSchedulingRulesText,
    setStaffs,
    setSchedule,
    setColors,
    setShowHistoryModal,
    setShowDraftPrompt
  } = deps;

    const { state } = record;
    const nextMonthlySchedules = state.monthlySchedules || {};
    const nextPreScheduleMonthlySchedules = state.preScheduleMonthlySchedules || {};
    const targetMonthKey = buildMonthKey(state.year, state.month);
    const currentMonthState = nextMonthlySchedules?.[targetMonthKey] || null;

    setMonthlySchedules(nextMonthlySchedules);
    setPreScheduleMonthlySchedules(nextPreScheduleMonthlySchedules);
    setYear(state.year);
    setMonth(state.month);
    setCustomHolidays(Array.isArray(state.customHolidays) ? state.customHolidays : []);
    setSpecialWorkdays(Array.isArray(state.specialWorkdays) ? state.specialWorkdays : []);
    setMedicalCalendarAdjustments(state.medicalCalendarAdjustments || { holidays: [], workdays: [] });
    if (state.staffingConfig) setStaffingConfig(state.staffingConfig);
    if (state.uiSettings) setUiSettings(state.uiSettings);
    if (Array.isArray(state.customLeaveCodes)) setCustomLeaveCodes(state.customLeaveCodes);
    if (Array.isArray(state.customWorkShifts)) setCustomWorkShifts(state.customWorkShifts);
    if (Array.isArray(state.customColumns)) setCustomColumns(state.customColumns);
    setCustomColumnValues(currentMonthState?.customColumnValues || state.customColumnValues || {});
    if (typeof (currentMonthState?.schedulingRulesText ?? state.schedulingRulesText) === 'string') {
      setSchedulingRulesText(currentMonthState?.schedulingRulesText ?? state.schedulingRulesText);
    }
    setStaffs(normalizeStaffGroup(currentMonthState?.staffs || state.staffs));
    setSchedule(currentMonthState?.scheduleData || state.schedule);
    if (state.colors) setColors(state.colors);
    setShowHistoryModal(false);
    setShowDraftPrompt(false);
};

export const createClearHistory = (deps) => () => {
  const {
    STORAGE_KEY,
    setHistoryList
  } = deps;

    if (window.confirm("確定要清空所有本機暫存紀錄嗎？")) {
      localStorage.removeItem(STORAGE_KEY);
      setHistoryList([]);
    }
};
