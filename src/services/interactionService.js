export const createStartRangeSelection = (deps = {}) => {
  const {
    selectionAnchor = null,
    setSelectionAnchor = () => {},
    setRangeSelection = () => {},
    setSelectedGridCell = () => {},
    resetKeyInputBuffer = () => {}
  } = deps;

  return (staff, dateStr, event = {}) => {
    const mouseButton = typeof event.button === 'number' ? event.button : 0;
    if (mouseButton !== 0) return;

    const point = { staffId: staff.id, dateStr, group: staff.group || '白班' };
    if (event.shiftKey && selectionAnchor) {
      setRangeSelection({ start: selectionAnchor, end: point });
    } else {
      setSelectionAnchor(point);
      setRangeSelection({ start: point, end: point });
    }
    setSelectedGridCell({ staff, dateStr });
    resetKeyInputBuffer();
  };
};

export const createUpdateRangeSelection = (deps = {}) => {
  const {
    isRangeDragging = false,
    selectionAnchor = null,
    setRangeSelection = () => {}
  } = deps;

  return (staff, dateStr) => {
    if (!isRangeDragging || !selectionAnchor) return;
    const anchorGroup = selectionAnchor.group || '白班';
    const targetGroup = staff.group || '白班';
    if (anchorGroup !== targetGroup) return;
    setRangeSelection({ start: selectionAnchor, end: { staffId: staff.id, dateStr, group: targetGroup } });
  };
};

export const createMoveSelectedCell = (deps = {}) => {
  const {
    getEffectiveSelection = () => null,
    selectedGridCell = null,
    staffs = [],
    daysInMonth = [],
    setSelectionAnchor = () => {},
    setRangeSelection = () => {},
    setSelectedGridCell = () => {},
    resetKeyInputBuffer = () => {}
  } = deps;

  return (rowDelta = 0, colDelta = 0) => {
    const selection = getEffectiveSelection();
    if (!selection?.start) return false;

    const activeStaffId = selectedGridCell?.staff?.id || selection.end?.staffId || selection.start.staffId;
    const activeDateStr = selectedGridCell?.dateStr || selection.end?.dateStr || selection.start.dateStr;
    const activeStaff = staffs.find((staff) => staff.id === activeStaffId);
    if (!activeStaff || !activeDateStr) return false;

    const scopedStaffs = staffs.filter((staff) => (staff.group || '白班') === (activeStaff.group || '白班'));
    const rowIndex = scopedStaffs.findIndex((staff) => staff.id === activeStaff.id);
    const colIndex = daysInMonth.findIndex((day) => day.date === activeDateStr);
    if (rowIndex === -1 || colIndex === -1) return false;

    let nextRowIndex = rowIndex + rowDelta;
    let nextColIndex = colIndex + colDelta;

    if (colDelta !== 0 && daysInMonth.length > 0) {
      while (nextColIndex >= daysInMonth.length) {
        if (nextRowIndex >= scopedStaffs.length - 1) {
          nextRowIndex = scopedStaffs.length - 1;
          nextColIndex = daysInMonth.length - 1;
          break;
        }
        nextRowIndex += 1;
        nextColIndex -= daysInMonth.length;
      }

      while (nextColIndex < 0) {
        if (nextRowIndex <= 0) {
          nextRowIndex = 0;
          nextColIndex = 0;
          break;
        }
        nextRowIndex -= 1;
        nextColIndex += daysInMonth.length;
      }
    }

    nextRowIndex = Math.max(0, Math.min(scopedStaffs.length - 1, nextRowIndex));
    nextColIndex = Math.max(0, Math.min(daysInMonth.length - 1, nextColIndex));

    const nextStaff = scopedStaffs[nextRowIndex];
    const nextDay = daysInMonth[nextColIndex];
    if (!nextStaff || !nextDay) return false;

    const nextPoint = { staffId: nextStaff.id, dateStr: nextDay.date, group: nextStaff.group || '白班' };
    setSelectionAnchor(nextPoint);
    setRangeSelection({ start: nextPoint, end: nextPoint });
    setSelectedGridCell({ staff: nextStaff, dateStr: nextDay.date });
    resetKeyInputBuffer();
    return true;
  };
};

export const createCopySelectionToClipboard = (deps = {}) => {
  const {
    getEffectiveSelection = () => null,
    staffs = [],
    daysInMonth = [],
    getRectFromSelection = () => null,
    preScheduleEditMode = false,
    getVisiblePreScheduleCode = () => '',
    getCellCode = () => '',
    setClipboardGrid = () => {}
  } = deps;

  return async () => {
    const selection = getEffectiveSelection();
    const rect = getRectFromSelection(selection, staffs, daysInMonth);
    if (!rect) return;

    const grid = [];
    for (let rowIndex = rect.rowStart; rowIndex <= rect.rowEnd; rowIndex += 1) {
      const row = [];
      const staff = rect.scopedStaffs[rowIndex];
      for (let colIndex = rect.colStart; colIndex <= rect.colEnd; colIndex += 1) {
        const day = daysInMonth[colIndex];
        const copiedValue = preScheduleEditMode
          ? (getVisiblePreScheduleCode(staff?.id, day?.date) || '')
          : (getCellCode(staff?.id, day?.date) || '');
        row.push(copiedValue);
      }
      grid.push(row);
    }

    setClipboardGrid(grid);
    const text = grid.map((row) => row.join('\t')).join('\n');
    try {
      if (navigator?.clipboard?.writeText) await navigator.clipboard.writeText(text);
    } catch (error) {
      console.error('寫入剪貼簿失敗', error);
    }
  };
};

export const buildPastePlan = (grid = [], rect = null, deps = {}) => {
  const {
    daysInMonth = [],
    normalizeManualShiftCode = () => ({ normalized: '', isValid: false }),
    mergedLeaveCodes = [],
    mergedShiftCodes = []
  } = deps;

  if (!rect || !Array.isArray(grid) || grid.length === 0) {
    return { updates: [], affectedCells: [], invalidCount: 0, clearCount: 0, writeCount: 0, clipped: false };
  }

  const selectionRowCount = rect.rowEnd - rect.rowStart + 1;
  const selectionColCount = rect.colEnd - rect.colStart + 1;
  const sourceRowCount = grid.length;
  const sourceColCount = Math.max(...grid.map((row) => (row || []).length), 0);
  const isSingleCellPaste = sourceRowCount === 1 && sourceColCount === 1;

  let targetRowCount = sourceRowCount;
  let targetColCount = sourceColCount;
  let clipped = false;

  if (selectionRowCount > 1 || selectionColCount > 1) {
    if (isSingleCellPaste) {
      targetRowCount = selectionRowCount;
      targetColCount = selectionColCount;
    } else {
      targetRowCount = Math.min(sourceRowCount, selectionRowCount);
      targetColCount = Math.min(sourceColCount, selectionColCount);
      clipped = sourceRowCount > selectionRowCount || sourceColCount > selectionColCount;
    }
  }

  const updates = [];
  const affectedCells = [];
  let invalidCount = 0;
  let clearCount = 0;
  let writeCount = 0;

  for (let rowOffset = 0; rowOffset < targetRowCount; rowOffset += 1) {
    for (let colOffset = 0; colOffset < targetColCount; colOffset += 1) {
      const sourceRow = isSingleCellPaste ? 0 : rowOffset;
      const sourceCol = isSingleCellPaste ? 0 : colOffset;
      const targetRow = rect.rowStart + rowOffset;
      const targetCol = rect.colStart + colOffset;
      const staff = rect.scopedStaffs[targetRow];
      const day = daysInMonth[targetCol];
      if (!staff || !day) continue;

      const targetCell = { staffId: staff.id, dateStr: day.date };
      affectedCells.push(targetCell);

      const rawValue = grid[sourceRow]?.[sourceCol] ?? '';
      const rawText = String(rawValue ?? '').trim();
      if (!rawText) {
        updates.push({ ...targetCell, value: '', source: 'manual' });
        clearCount += 1;
        continue;
      }

      const { normalized, isValid } = normalizeManualShiftCode(rawText, [...mergedLeaveCodes, ...mergedShiftCodes]);
      if (!isValid) {
        invalidCount += 1;
        continue;
      }

      updates.push({ ...targetCell, value: normalized, source: 'manual' });
      writeCount += 1;
    }
  }

  return { updates, affectedCells, invalidCount, clearCount, writeCount, clipped };
};

export const createPasteGridToSelection = (deps = {}) => {
  const {
    getEffectiveSelection = () => null,
    staffs = [],
    daysInMonth = [],
    getRectFromSelection = () => null,
    clipboardGrid = [],
    parseClipboardGrid = () => [],
    preScheduleEditMode = false,
    buildPreSchedulePastePlan = () => ({ entries: [], affectedCells: [], invalidCount: 0 }),
    updatePreScheduleEntries = () => 0,
    setSelectionRangeFromCells = () => false,
    resetKeyInputBuffer = () => {},
    clearInputAssist = () => {},
    flashInvalidSelection = () => {},
    buildPastePlan = () => ({ updates: [], affectedCells: [], invalidCount: 0 }),
    selectedRangeCells = [],
    validateManualEntries = () => ({ allowedEntries: [] }),
    applyScheduleEntries = () => {}
  } = deps;

  return async () => {
    const selection = getEffectiveSelection();
    const rect = getRectFromSelection(selection, staffs, daysInMonth);
    if (!rect) return;

    let grid = clipboardGrid;
    if (!grid || grid.length === 0) {
      try {
        if (navigator?.clipboard?.readText) {
          const text = await navigator.clipboard.readText();
          grid = parseClipboardGrid(text);
        }
      } catch (error) {
        console.error('讀取剪貼簿失敗', error);
      }
    }
    if (!grid || grid.length === 0) return;

    if (preScheduleEditMode) {
      const pastePlan = buildPreSchedulePastePlan(grid, rect);
      if (pastePlan.affectedCells.length === 0) return;

      if (pastePlan.entries.length === 0) {
        if (pastePlan.invalidCount > 0) flashInvalidSelection(pastePlan.affectedCells);
        return;
      }

      updatePreScheduleEntries(pastePlan.entries);
      setSelectionRangeFromCells(pastePlan.affectedCells, { activeCell: pastePlan.affectedCells[pastePlan.affectedCells.length - 1] });
      resetKeyInputBuffer();
      clearInputAssist();
      if (pastePlan.invalidCount > 0) flashInvalidSelection(pastePlan.affectedCells);
      return;
    }

    const pastePlan = buildPastePlan(grid, rect);
    if (pastePlan.updates.length === 0) {
      if (pastePlan.invalidCount > 0) flashInvalidSelection(selectedRangeCells);
      return;
    }

    const { allowedEntries } = validateManualEntries(pastePlan.updates, { showFeedback: false });
    if (allowedEntries.length === 0) {
      if (pastePlan.invalidCount > 0) flashInvalidSelection(pastePlan.affectedCells);
      return;
    }

    applyScheduleEntries(allowedEntries, {
      preserveSelection: true,
      selectionCells: pastePlan.affectedCells,
      activeCell: pastePlan.affectedCells[pastePlan.affectedCells.length - 1],
      clearAssist: false,
      resetBuffer: true
    });

    if (pastePlan.invalidCount > 0) flashInvalidSelection(pastePlan.affectedCells);
    clearInputAssist();
  };
};
