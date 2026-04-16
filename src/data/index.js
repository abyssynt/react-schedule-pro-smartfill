export * from './scheduleData.js';
export * from './shiftResolverData.js';
export * from './calendarData.js';
export * from './runtimeLoaderData.js';
export * from './gridSelectionData.js';
export * from './statsData.js';
export * from './storageData.js';
export * from './uiConfigData.js';
export * from './exportDataHelpers.js';
export * from './ruleData.js';

import {
  normalizeImportedHalfWidth as baseNormalizeImportedHalfWidth,
  getImportedRawNumberedLeaveValue as baseGetImportedRawNumberedLeaveValue,
  normalizeImportedShiftCode as baseNormalizeImportedShiftCode,
  isConfiguredImportedLeaveCode as baseIsConfiguredImportedLeaveCode,
  normalizeCodeComparisonValue,
  normalizeCodeComparisonCompact,
  normalizeCodeComparisonCompactNoHyphen,
  isBuiltInCode,
  normalizeManualShiftCode as baseNormalizeManualShiftCode,
  getNormalizedManualCodeCandidates as baseGetNormalizedManualCodeCandidates
} from './importCodeData.js';
import {
  extractYearMonthCandidates,
  detectImportedDayNumber,
  inferImportedGroupFromCodes as baseInferImportedGroupFromCodes,
  parseImportedWorksheet as baseParseImportedWorksheet,
  parseImportedExcelFiles as baseParseImportedExcelFiles
} from './importMonthData.js';
import {
  buildMonthKey,
  normalizeStaffGroup,
  createBlankMonthStaffs,
  createBlankMonthState,
  buildExistingStaffGroupLookup,
  mergeImportedMonthStates,
  reconcileScheduleDataMap as baseReconcileScheduleDataMap,
  reconcileMonthStateCollections as baseReconcileMonthStateCollections,
  createBlankScheduleForStaffs,
  getMonthScheduleData,
  resolveRulesTextCandidates
} from './monthScheduleData.js';
import {
  clampColorChannel,
  normalizeHexColor,
  hexToRgbObject,
  rgbObjectToHex,
  blendHexColors,
  hexToExcelArgb,
  FOUR_WEEK_CYCLE_START,
  FOUR_WEEK_CYCLE_DAYS,
  RULE_CROSS_MONTH_CONTEXT_DAYS,
  isFourWeekCycleEndDate as baseIsFourWeekCycleEndDate
} from './colorCycleData.js';
import { loadSheetJS } from './runtimeLoaderData.js';
import { DICT } from './scheduleData.js';
import { getAllShiftCodes, getCodePrefix, getShiftGroupByCode } from './shiftResolverData.js';

export {
  baseNormalizeImportedHalfWidth as normalizeImportedHalfWidth,
  baseGetImportedRawNumberedLeaveValue as getImportedRawNumberedLeaveValue,
  normalizeCodeComparisonValue,
  normalizeCodeComparisonCompact,
  normalizeCodeComparisonCompactNoHyphen,
  isBuiltInCode,
  extractYearMonthCandidates,
  detectImportedDayNumber,
  buildMonthKey,
  normalizeStaffGroup,
  createBlankMonthStaffs,
  createBlankMonthState,
  buildExistingStaffGroupLookup,
  mergeImportedMonthStates,
  createBlankScheduleForStaffs,
  getMonthScheduleData,
  resolveRulesTextCandidates,
  clampColorChannel,
  normalizeHexColor,
  hexToRgbObject,
  rgbObjectToHex,
  blendHexColors,
  hexToExcelArgb,
  FOUR_WEEK_CYCLE_START,
  FOUR_WEEK_CYCLE_DAYS,
  RULE_CROSS_MONTH_CONTEXT_DAYS
};

export const normalizeImportedShiftCode = (rawValue = '') =>
  baseNormalizeImportedShiftCode(rawValue, getAllShiftCodes);

export const isConfiguredImportedLeaveCode = (code = '', customLeaveCodes = []) =>
  baseIsConfiguredImportedLeaveCode(code, customLeaveCodes, getCodePrefix);

export const normalizeManualShiftCode = (rawValue = '', allowedLeaveCodes = []) =>
  baseNormalizeManualShiftCode(rawValue, allowedLeaveCodes, getAllShiftCodes);

export const getNormalizedManualCodeCandidates = (rawValue = '', allowedLeaveCodes = []) =>
  baseGetNormalizedManualCodeCandidates(rawValue, allowedLeaveCodes, getAllShiftCodes);

export const isPotentialManualShiftPrefix = (rawValue = '', allowedLeaveCodes = []) =>
  !String(rawValue ?? '').trim() || getNormalizedManualCodeCandidates(rawValue, allowedLeaveCodes).length > 0;

export const inferImportedGroupFromCodes = (dayMap = {}) =>
  baseInferImportedGroupFromCodes(dayMap, { getShiftGroupByCode });

export const parseImportedWorksheet = (options = {}) =>
  baseParseImportedWorksheet(options, { getAllShiftCodes, getShiftGroupByCode, getCodePrefix });

export const parseImportedExcelFiles = (files = [], fallbackYear = new Date().getFullYear(), options = {}) =>
  baseParseImportedExcelFiles(files, fallbackYear, options, { loadSheetJS, getAllShiftCodes, getShiftGroupByCode, getCodePrefix });

export const isFourWeekCycleEndDate = (dateStr, cycleStart = FOUR_WEEK_CYCLE_START) =>
  baseIsFourWeekCycleEndDate(dateStr, (value) => {
    const [y, m, d] = String(value || '').split('-').map(Number);
    return new Date(y, (m || 1) - 1, d || 1);
  }, cycleStart);

export const reconcileScheduleDataMap = (scheduleData = {}, customLeaveCodes = []) =>
  baseReconcileScheduleDataMap(scheduleData, customLeaveCodes, {
    DICT,
    getAllShiftCodes,
    normalizeManualShiftCode: (rawValue = '', leaveCodes = []) => baseNormalizeManualShiftCode(rawValue, leaveCodes, getAllShiftCodes),
    normalizeImportedHalfWidth: baseNormalizeImportedHalfWidth
  });

export const reconcileMonthStateCollections = (collection = {}, customLeaveCodes = []) =>
  baseReconcileMonthStateCollections(collection, customLeaveCodes, {
    DICT,
    getAllShiftCodes,
    normalizeManualShiftCode: (rawValue = '', leaveCodes = []) => baseNormalizeManualShiftCode(rawValue, leaveCodes, getAllShiftCodes),
    normalizeImportedHalfWidth: baseNormalizeImportedHalfWidth
  });
