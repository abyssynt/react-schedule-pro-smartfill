export const makeCellKey = (staffId, dateStr) => `${staffId}__${dateStr}`;

export const parseClipboardGrid = (text = '') => {
  const raw = String(text || '').replace(/\r/g, '');
  if (!raw.trim()) return [];
  return raw.split('\n').map(row => row.split('\t'));
};

export const getSelectionGroupStaffs = (selection, staffs = []) => {
  const selectionGroup = selection?.start?.group || selection?.end?.group || '';
  if (!selectionGroup) return staffs;
  return staffs.filter((staff) => (staff.group || '白班') === selectionGroup);
};

export const getRectFromSelection = (selection, staffs = [], daysInMonth = []) => {
  if (!selection?.start || !selection?.end) return null;
  const scopedStaffs = getSelectionGroupStaffs(selection, staffs);
  const staffIndexMap = new Map(scopedStaffs.map((staff, index) => [staff.id, index]));
  const dayIndexMap = new Map(daysInMonth.map((day, index) => [day.date, index]));

  const startRow = staffIndexMap.get(selection.start.staffId);
  const endRow = staffIndexMap.get(selection.end.staffId);
  const startCol = dayIndexMap.get(selection.start.dateStr);
  const endCol = dayIndexMap.get(selection.end.dateStr);

  if ([startRow, endRow, startCol, endCol].some(v => v === undefined)) return null;

  return {
    rowStart: Math.min(startRow, endRow),
    rowEnd: Math.max(startRow, endRow),
    colStart: Math.min(startCol, endCol),
    colEnd: Math.max(startCol, endCol),
    scopedStaffs
  };
};

export const expandSelectionCells = (selection, staffs = [], daysInMonth = []) => {
  const rect = getRectFromSelection(selection, staffs, daysInMonth);
  if (!rect) return [];
  const cells = [];
  for (let rowIndex = rect.rowStart; rowIndex <= rect.rowEnd; rowIndex += 1) {
    for (let colIndex = rect.colStart; colIndex <= rect.colEnd; colIndex += 1) {
      const staff = rect.scopedStaffs[rowIndex];
      const day = daysInMonth[colIndex];
      if (staff && day) cells.push({ staffId: staff.id, dateStr: day.date, rowIndex, colIndex });
    }
  }
  return cells;
};

export const isCellInSelectionRect = (selection, staffs = [], daysInMonth = [], staffId, dateStr) => {
  const rect = getRectFromSelection(selection, staffs, daysInMonth);
  if (!rect) return false;
  const rowIndex = rect.scopedStaffs.findIndex(staff => staff.id === staffId);
  const colIndex = daysInMonth.findIndex(day => day.date === dateStr);
  if (rowIndex === -1 || colIndex === -1) return false;
  return rowIndex >= rect.rowStart && rowIndex <= rect.rowEnd && colIndex >= rect.colStart && colIndex <= rect.colEnd;
};
