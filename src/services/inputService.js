export const normalizePreScheduleInputService = (rawValue = '', deps = {}) => {
  const {
    getImportedRawNumberedLeaveValue = () => '',
    normalizeManualShiftCode = () => ({ normalized: '', isValid: false }),
    mergedLeaveCodes = [],
    mergedShiftCodes = []
  } = deps;

  const raw = String(rawValue ?? '').trim();
  if (!raw) return { normalized: '', isValid: true };

  const normalizedNumberedLeave = getImportedRawNumberedLeaveValue(raw);
  if (normalizedNumberedLeave) {
    return {
      normalized: normalizedNumberedLeave.startsWith('例') ? '例' : '休',
      isValid: true
    };
  }

  const { normalized, isValid } = normalizeManualShiftCode(raw, [...mergedLeaveCodes, ...mergedShiftCodes]);
  if (!isValid) return { normalized: '', isValid: false };
  return { normalized, isValid: true };
};

export const validateManualEntriesService = (entries = [], options = {}, deps = {}) => {
  const {
    schedule = {},
    isConfiguredLeaveCode = () => false,
    isShiftCode = () => false,
    staffs = [],
    evaluateRuleWarningForCellInSnapshot = () => [],
    setRuleWarningsForEntries = () => {},
    showInputAssist = () => {}
  } = deps;

  const allowedEntries = [];
  const warningEntries = [];
  const showFeedback = options.showFeedback !== false;
  const workingSnapshot = JSON.parse(JSON.stringify(schedule || {}));

  (Array.isArray(entries) ? entries : []).forEach((entry) => {
    const normalizedValue = String(entry?.value || '').trim();
    const normalizedEntry = { ...entry, value: normalizedValue, source: entry?.source || 'manual' };
    if (!normalizedEntry.staffId || !normalizedEntry.dateStr) return;

    if (!workingSnapshot[normalizedEntry.staffId]) workingSnapshot[normalizedEntry.staffId] = {};
    const existingCell = workingSnapshot[normalizedEntry.staffId]?.[normalizedEntry.dateStr];
    const existingMeta = typeof existingCell === 'object' && existingCell !== null ? existingCell : {};
    workingSnapshot[normalizedEntry.staffId][normalizedEntry.dateStr] = normalizedValue
      ? { ...existingMeta, value: normalizedValue, source: normalizedEntry.source }
      : null;

    if (!normalizedValue || isConfiguredLeaveCode(normalizedValue) || !isShiftCode(normalizedValue)) {
      allowedEntries.push(normalizedEntry);
      return;
    }

    const staff = (Array.isArray(staffs) ? staffs : []).find((item) => item.id === normalizedEntry.staffId);
    if (!staff) {
      allowedEntries.push(normalizedEntry);
      return;
    }

    const reasons = evaluateRuleWarningForCellInSnapshot(workingSnapshot, staff, normalizedEntry.dateStr);
    allowedEntries.push(normalizedEntry);
    if (reasons.length > 0) {
      warningEntries.push({ ...normalizedEntry, reasons });
    }
  });

  if (warningEntries.length > 0 && showFeedback) {
    setRuleWarningsForEntries(warningEntries);
    showInputAssist(`已寫入，但有 ${warningEntries.length} 格違反規則`, 'warning', 2200);
  }

  return { allowedEntries, warningEntries };
};

export const applyValueToCellsService = (cells, normalized, options = {}, deps = {}) => {
  const {
    validateManualEntries = () => ({ allowedEntries: [], warningEntries: [] }),
    clearRuleWarningCells = () => {},
    applyScheduleEntries = () => false
  } = deps;

  if (!cells || cells.length === 0) return false;
  const targetCells = Array.isArray(cells) ? cells : [];
  const source = options.source || 'manual';
  const rawEntries = targetCells.map(({ staffId, dateStr }) => ({
    staffId,
    dateStr,
    value: normalized,
    source
  }));

  const { allowedEntries, warningEntries } = source === 'manual'
    ? validateManualEntries(rawEntries, { showFeedback: options.clearAssist !== false })
    : { allowedEntries: rawEntries, warningEntries: [] };

  if (allowedEntries.length === 0) return false;

  const nonWarningCells = targetCells.filter(({ staffId, dateStr }) =>
    !warningEntries.some((entry) => entry.staffId === staffId && entry.dateStr === dateStr)
  );

  if (source === 'manual') clearRuleWarningCells(nonWarningCells);

  return applyScheduleEntries(allowedEntries, {
    ...options,
    clearAssist: options.clearAssist !== false && warningEntries.length === 0,
    preserveSelection: options.preserveSelection === true || warningEntries.length > 0,
    selectionCells: options.selectionCells || targetCells,
    activeCell: options.activeCell || targetCells[targetCells.length - 1]
  });
};

export const applyPreScheduleValueToCellsService = (cells = [], rawValue = '', options = {}, deps = {}) => {
  const {
    normalizePreScheduleInput = () => ({ normalized: '', isValid: false }),
    flashInvalidSelection = () => {},
    showInputAssist = () => {},
    updatePreScheduleEntries = () => 0,
    setSelectionRangeFromCells = () => false,
    clearInputAssist = () => {},
    resetKeyInputBuffer = () => {},
    moveSelectionAfterInput = () => false
  } = deps;

  if (!Array.isArray(cells) || cells.length === 0) return { applied: false, normalized: '' };
  const { normalized, isValid } = normalizePreScheduleInput(rawValue);
  if (!isValid) {
    flashInvalidSelection(cells);
    if (options.showFeedback !== false) showInputAssist('預班可輸入上班或休假代號', 'error');
    return { applied: false, normalized: '' };
  }

  const changedCount = updatePreScheduleEntries(cells.map(({ staffId, dateStr }) => ({
    staffId,
    dateStr,
    value: normalized
  })));

  if (changedCount <= 0 && normalized) return { applied: false, normalized: '' };

  const shouldAdvance = options.advance !== false && normalized && cells.length === 1;
  if (options.preserveSelection || !shouldAdvance) {
    setSelectionRangeFromCells(cells, { activeCell: options.activeCell || cells[cells.length - 1] });
  }

  if (options.clearAssist !== false) clearInputAssist();
  resetKeyInputBuffer();

  if (shouldAdvance) moveSelectionAfterInput(cells, options.direction === -1 ? -1 : 1);

  return { applied: true, normalized };
};

export const applySelectionValueService = (cells = [], rawValue = '', options = {}, deps = {}) => {
  const {
    normalizeManualShiftCode = () => ({ normalized: '', isValid: false }),
    mergedLeaveCodes = [],
    mergedShiftCodes = [],
    applyValueToCells = () => false,
    clearInputAssist = () => {},
    resetKeyInputBuffer = () => {},
    moveSelectionAfterInput = () => false
  } = deps;

  if (!Array.isArray(cells) || cells.length === 0) return { applied: false, normalized: '' };
  const { normalized, isValid } = normalizeManualShiftCode(rawValue, [...mergedLeaveCodes, ...mergedShiftCodes]);
  if (!isValid) return { applied: false, normalized: '' };

  const shouldAdvance = options.advance !== false && normalized && cells.length === 1;
  const applied = applyValueToCells(cells, normalized, {
    source: options.source || 'manual',
    preserveSelection: !shouldAdvance,
    selectionCells: cells,
    activeCell: cells[cells.length - 1],
    clearAssist: options.clearAssist,
    resetBuffer: options.resetBuffer
  });
  if (!applied) return { applied: false, normalized: '' };

  if (options.clearAssist !== false) clearInputAssist();
  resetKeyInputBuffer();

  if (shouldAdvance) {
    moveSelectionAfterInput(cells, options.direction === -1 ? -1 : 1);
  }

  return { applied: true, normalized };
};

export const handleCellChangeService = (staffId, dateStr, value, options = {}, deps = {}) => {
  const {
    validateManualEntries = () => ({ allowedEntries: [], warningEntries: [] }),
    clearRuleWarningCells = () => {},
    applyScheduleEntries = () => false
  } = deps;

  const source = options.source || 'manual';
  const rawEntries = [{ staffId, dateStr, value, source }];
  const { allowedEntries, warningEntries } = source === 'manual'
    ? validateManualEntries(rawEntries, { showFeedback: options.clearAssist !== false })
    : { allowedEntries: rawEntries, warningEntries: [] };

  if (allowedEntries.length === 0) return false;

  const nonWarningCells = rawEntries.filter(({ staffId, dateStr }) =>
    !warningEntries.some((entry) => entry.staffId === staffId && entry.dateStr === dateStr)
  );

  if (source === 'manual') clearRuleWarningCells(nonWarningCells);
  if (!options.allowWarningAssistClear && warningEntries.length > 0) options = { ...options, clearAssist: false };

  return applyScheduleEntries(allowedEntries, {
    clearAssist: options.clearAssist !== false,
    resetBuffer: options.resetBuffer !== false,
    preserveSelection: options.preserveSelection === true || warningEntries.length > 0,
    selectionCells: options.selectionCells || [{ staffId, dateStr }],
    activeCell: options.activeCell || { staffId, dateStr }
  });
};
