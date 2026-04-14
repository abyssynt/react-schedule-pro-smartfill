import { DICT } from './scheduleData';

export const normalizeImportedHalfWidth = (input = '') => String(input ?? '')
  .replace(/[！-～]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xFEE0))
  .replace(/　/g, ' ')
  .trim();

export const getImportedRawNumberedLeaveValue = (rawValue = '') => {
  const normalized = normalizeImportedHalfWidth(rawValue).replace(/\s+/g, '');
  const match = normalized.match(/^(例|休)([1-4])$/);
  if (!match) return '';
  return `${match[1]}${match[2]}`;
};

export const normalizeImportedShiftCode = (rawValue = '', getAllShiftCodes = () => []) => {
  const value = String(rawValue ?? '').trim();
  if (!value) return '';

  const normalizedWhitespace = normalizeImportedHalfWidth(value).replace(/\s+/g, '');
  const lower = normalizedWhitespace.toLowerCase();
  const numberedLeaveValue = getImportedRawNumberedLeaveValue(normalizedWhitespace);
  if (numberedLeaveValue) return numberedLeaveValue.startsWith('例') ? '例' : '休';

  const directMap = {
    d: 'D',
    e: 'E',
    n: 'N',
    off: 'off',
    of: 'off',
    am: 'AM',
    pm: 'PM',
    '8-12': '8-12',
    '12-16': '12-16',
    '白8-8': '白8-8',
    '夜8-8': '夜8-8'
  };

  if (directMap[lower]) return directMap[lower];
  if (DICT.LEAVES.includes(normalizedWhitespace) || getAllShiftCodes().includes(normalizedWhitespace)) return normalizedWhitespace;
  return normalizeImportedHalfWidth(value);
};

export const isConfiguredImportedLeaveCode = (code = '', customLeaveCodes = [], getCodePrefix = (value) => value) => {
  const mergedLeaveCodes = Array.from(new Set([...(DICT.LEAVES || []), ...(customLeaveCodes || [])])).filter(Boolean);
  const prefix = getCodePrefix(code);
  return mergedLeaveCodes.includes(code) || mergedLeaveCodes.includes(prefix);
};

export const normalizeCodeComparisonValue = (input = '') => normalizeImportedHalfWidth(input).trim().replace(/\s+/g, '');
export const normalizeCodeComparisonCompact = (input = '') => normalizeCodeComparisonValue(input).replace(/[－—–~～_]/g, '-');
export const normalizeCodeComparisonCompactNoHyphen = (input = '') => normalizeCodeComparisonCompact(input).replace(/-/g, '');
export const isBuiltInCode = (code = '') => DICT.SHIFTS.includes(code) || DICT.LEAVES.includes(code);

export const normalizeManualShiftCode = (rawValue = '', allowedLeaveCodes = [], getAllShiftCodes = () => []) => {
  const value = String(rawValue ?? '').trim();
  if (!value) return { normalized: '', isValid: true };

  const normalizedBase = normalizeImportedHalfWidth(value).trim();
  const collapsed = normalizeCodeComparisonValue(normalizedBase);
  const lower = collapsed.toLowerCase();
  const compact = lower.replace(/[－—–~～_]/g, '-');
  const compactNoHyphen = compact.replace(/-/g, '');
  const exactCompact = normalizeCodeComparisonCompact(normalizedBase);
  const exactCompactNoHyphen = normalizeCodeComparisonCompactNoHyphen(normalizedBase);
  const allowedCodes = Array.from(new Set([...(getAllShiftCodes() || []), ...(allowedLeaveCodes || [])])).filter(Boolean);

  const exactAllowed = allowedCodes.find((code) => {
    const codeCompact = normalizeCodeComparisonCompact(code);
    const codeCompactNoHyphen = normalizeCodeComparisonCompactNoHyphen(code);
    return codeCompact === exactCompact || codeCompactNoHyphen === exactCompactNoHyphen;
  });
  if (exactAllowed) return { normalized: exactAllowed, isValid: true };

  const directMap = {
    d: 'D',
    e: 'E',
    n: 'N',
    off: 'off',
    of: 'off',
    o: 'off',
    am: 'AM',
    pm: 'PM',
    a: 'AM',
    p: 'PM',
    '8-12': '8-12',
    '812': '8-12',
    '08-12': '8-12',
    '0812': '8-12',
    '12-16': '12-16',
    '1216': '12-16',
    '白8-8': '白8-8',
    '白88': '白8-8',
    '白8-08': '白8-8',
    '白8to8': '白8-8',
    '夜8-8': '夜8-8',
    '夜88': '夜8-8',
    '夜8to8': '夜8-8'
  };

  const directCandidate = directMap[compact] || directMap[compactNoHyphen] || directMap[lower];
  if (directCandidate) return { normalized: directCandidate, isValid: allowedCodes.includes(directCandidate) };

  const directAllowed = allowedCodes.find((code) => {
    if (!isBuiltInCode(code)) return false;
    const codeLower = normalizeCodeComparisonValue(code).toLowerCase();
    const codeCompact = codeLower.replace(/[－—–~～_]/g, '-');
    const codeCompactNoHyphen = codeCompact.replace(/-/g, '');
    return codeLower === lower || codeCompact === compact || codeCompactNoHyphen === compactNoHyphen;
  });
  if (directAllowed) return { normalized: directAllowed, isValid: true };

  return { normalized: normalizedBase, isValid: false };
};

export const getNormalizedManualCodeCandidates = (rawValue = '', allowedLeaveCodes = [], getAllShiftCodes = () => []) => {
  const value = String(rawValue ?? '').trim();
  if (!value) return [];

  const normalizedBase = normalizeImportedHalfWidth(value).trim();
  const collapsed = normalizeCodeComparisonValue(normalizedBase);
  const lower = collapsed.toLowerCase();
  const compact = lower.replace(/[－—–~～_]/g, '-');
  const compactNoHyphen = compact.replace(/-/g, '');
  const aliases = new Set([lower, compact, compactNoHyphen]);
  const exactCompact = normalizeCodeComparisonCompact(normalizedBase);
  const exactCompactNoHyphen = normalizeCodeComparisonCompactNoHyphen(normalizedBase);
  const allowedCodes = Array.from(new Set([...(getAllShiftCodes() || []), ...(allowedLeaveCodes || [])])).filter(Boolean);

  const expandedAliases = new Set(aliases);
  if (aliases.has('o')) {
    expandedAliases.add('of');
    expandedAliases.add('off');
  }
  if (aliases.has('of')) expandedAliases.add('off');
  if (aliases.has('a')) expandedAliases.add('am');
  if (aliases.has('p')) expandedAliases.add('pm');
  if (aliases.has('8')) {
    expandedAliases.add('8-12');
    expandedAliases.add('812');
  }
  if (aliases.has('12')) {
    expandedAliases.add('12-16');
    expandedAliases.add('1216');
  }
  if (aliases.has('白8')) {
    expandedAliases.add('白8-8');
    expandedAliases.add('白88');
  }
  if (aliases.has('夜8')) {
    expandedAliases.add('夜8-8');
    expandedAliases.add('夜88');
  }

  return allowedCodes.filter((code) => {
    const codeCompact = normalizeCodeComparisonCompact(code);
    const codeCompactNoHyphen = normalizeCodeComparisonCompactNoHyphen(code);
    if (!isBuiltInCode(code)) {
      return codeCompact.startsWith(exactCompact) || codeCompactNoHyphen.startsWith(exactCompactNoHyphen);
    }
    const codeLower = normalizeCodeComparisonValue(code).toLowerCase();
    const codeLowerCompact = codeLower.replace(/[－—–~～_]/g, '-');
    const codeLowerCompactNoHyphen = codeLowerCompact.replace(/-/g, '');
    return Array.from(expandedAliases).some((alias) => codeLower.startsWith(alias) || codeLowerCompact.startsWith(alias) || codeLowerCompactNoHyphen.startsWith(alias));
  });
};

export const isPotentialManualShiftPrefix = (rawValue = '', allowedLeaveCodes = [], getAllShiftCodes = () => []) => {
  if (!String(rawValue ?? '').trim()) return true;
  return getNormalizedManualCodeCandidates(rawValue, allowedLeaveCodes, getAllShiftCodes).length > 0;
};
