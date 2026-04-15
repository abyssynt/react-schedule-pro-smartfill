import { DICT } from './scheduleData';

export const BLOCKED_LEAVE_PREFIXES = ['off', '例', '休', '特', '補', '國', '喪', '婚', '產', '病', '事', '陪產', 'AM', 'PM'];

let CUSTOM_SHIFT_DEFS = [];

export const getCustomShiftCodes = () =>
  CUSTOM_SHIFT_DEFS.map((item) => String(item?.code || '').trim()).filter(Boolean);

export const getAllShiftCodes = () =>
  Array.from(new Set([...(DICT.SHIFTS || []), ...getCustomShiftCodes()]));

export const setCustomShiftDefsRegistry = (defs = []) => {
  CUSTOM_SHIFT_DEFS = Array.isArray(defs) ? defs : [];
};

export const getCustomShiftGroup = (code = '') => {
  const normalized = String(code || '').trim();
  if (!normalized) return null;
  const matched = CUSTOM_SHIFT_DEFS.find((item) => String(item?.code || '').trim() === normalized);
  return matched?.group || null;
};
