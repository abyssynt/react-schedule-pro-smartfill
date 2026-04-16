
import React, { useState, useMemo, useEffect, useRef } from 'react';
import {
  Plus, Minus, Settings, Sparkles, Loader2,
  Save, History as Clock, Download,
  FileSpreadsheet, FileText, X, Check, Calendar, CalendarDays,
  User, Lock, Info, Layout, ShieldCheck, Grid, UserCheck,
  Database, Cpu, Monitor, ArrowLeft, ChevronRight, CheckCircle2, Trash2, GripVertical, AlertTriangle
} from 'lucide-react';

import {
  DICT,
  SMART_RULES,
  SHIFT_GROUPS,
  getCustomShiftCodes,
  getAllShiftCodes,
  setCustomShiftDefsRegistry,
  getCustomShiftGroup,
  getCodePrefix,
  getShiftGroupByCode,
  getSystemHolidayCalendar,
  readSchedulingRulesTextFromLocalSettings,
  STORAGE_KEYS,
  loadExcelJS,
  loadSheetJS,
  normalizeImportedHalfWidth,
  getImportedRawNumberedLeaveValue,
  normalizeImportedShiftCode,
  buildMonthKey,
  parseDateKey,
  formatDateKey,
  addDays,
  isConfiguredImportedLeaveCode,
  normalizeCodeComparisonValue,
  normalizeCodeComparisonCompact,
  normalizeCodeComparisonCompactNoHyphen,
  isBuiltInCode,
  normalizeManualShiftCode,
  getNormalizedManualCodeCandidates,
  isPotentialManualShiftPrefix,
  makeCellKey,
  parseClipboardGrid,
  getSelectionGroupStaffs,
  getRectFromSelection,
  expandSelectionCells,
  isCellInSelectionRect,
  extractYearMonthCandidates,
  detectImportedDayNumber,
  inferImportedGroupFromCodes,
  buildExistingStaffGroupLookup,
  mergeImportedMonthStates,
  parseImportedWorksheet,
  parseImportedExcelFiles,
  normalizeStaffGroup,
  createBlankMonthStaffs,
  createBlankMonthState,
  isLeaveCode,
  isShiftCode,
  GROUP_TO_DEMAND_KEY,
  DEFAULT_REQUIRED_STAFFING,
  normalizeRequiredStaffingConfig,
  getRequiredStaffingBucketByDay,
  DEFAULT_SHIFT_BY_GROUP,
  RULE_FILL_MAIN_SHIFTS,
  HOSPITAL_LEVEL_LABELS,
  HOSPITAL_RATIO_HINTS,
  UI_FONT_SIZE_OPTIONS,
  getUiFontSizeClass,
  getShiftLabelFontSize,
  getShiftCellLabelFontSize,
  UI_DENSITY_OPTIONS,
  getUiDensityConfig,
  UI_THEME_PRESETS,
  getAdjustedDensityConfig,
  clampColorChannel,
  normalizeHexColor,
  hexToRgbObject,
  rgbObjectToHex,
  blendHexColors,
  hexToExcelArgb,
  FOUR_WEEK_CYCLE_START,
  FOUR_WEEK_CYCLE_DAYS,
  RULE_CROSS_MONTH_CONTEXT_DAYS,
  isFourWeekCycleEndDate,
  reconcileScheduleDataMap,
  reconcileMonthStateCollections
} from './data/index.js';
import { buildMonthStatePayload } from './services/monthStateService.js';

import { exportToExcelService, exportToWordService } from './services/exportService.js';
import { saveToHistoryService, loadHistoryService, clearHistoryService } from './services/historyService.js';
import { createHandleRuleBasedAutoSchedule } from './services/ruleFillService.js';
import {
  buildEmptyStaffStatsService,
  buildStaffStatsMapService,
  buildDailyStatsMapService,
  buildRequiredCountMapService,
  getDemandHighlightStyleService
} from './services/statsService.js';
import {
  getCellCodeFromSnapshotService,
  evaluateRuleWarningForCellInSnapshotService,
  refreshRuleWarningsForCellsService,
  expandCellsForRuleRecheckService,
  scanScheduleRuleViolationsService
} from './services/ruleWarningService.js';

const STORAGE_KEY = STORAGE_KEYS.HISTORY;
const ACTIVE_DRAFT_KEY = STORAGE_KEYS.ACTIVE_DRAFT;
const LOCAL_SETTINGS_KEY = STORAGE_KEYS.LOCAL_SETTINGS;

const ScheduleGridCell = React.memo(function ScheduleGridCell({
  dayInfo,
  densityConfig,
  tableFontSizeClass,
  tableFontColor,
  dangerTintColor,
  warningTintColor,
  preScheduleHintBorderColor,
  preScheduleTextColor,
  preScheduleDisplayTextColor,
  displayValue,
  cellTextColor,
  showPreScheduleAsMain,
  showPreScheduleAsHint,
  preScheduleCode,
  cellBackgroundColor,
  inRangeSelection,
  isPrimarySelected,
  isInvalid,
  isRuleWarning,
  ruleWarningMessage,
  rowInsertStyle,
  fourWeekDividerStyle,
  showBlueDots,
  selectionMode,
  showShiftLabels,
  selectorDotClass,
  selectValue,
  selectTitle,
  shiftOptionsNodes,
  leaveOptionsNodes,
  preScheduleEditMode,
  onMouseDownCell,
  onMouseEnterCell,
  onClickCell,
  onDotClick,
  onSelectChange,
  onSelectClick,
  onSelectMouseDown
}) {
  return (
    <td
      className={`border-r p-0 relative overflow-hidden ${inRangeSelection ? 'ring-2 ring-violet-400 ring-inset' : isPrimarySelected ? 'ring-2 ring-blue-500 ring-inset' : ''}`}
      style={{
        backgroundColor: cellBackgroundColor,
        opacity: dayInfo.isHoliday || dayInfo.isWeekend ? 0.9 : 1,
        boxShadow: `${inRangeSelection ? 'inset 0 0 0 2px #a78bfa' : isPrimarySelected ? 'inset 0 0 0 2px #3b82f6' : ''}${isInvalid ? `${inRangeSelection || isPrimarySelected ? ', ' : ''}inset 0 0 0 2px ${dangerTintColor}` : ''}${isRuleWarning ? `${inRangeSelection || isPrimarySelected || isInvalid ? ', ' : ''}inset 0 0 0 2px ${warningTintColor}` : ''}${showPreScheduleAsMain ? `${inRangeSelection || isPrimarySelected || isInvalid || isRuleWarning ? ', ' : ''}inset 0 0 0 1px ${preScheduleHintBorderColor}` : ''}` || undefined,
        ...fourWeekDividerStyle,
        ...(rowInsertStyle || {})
      }}
      onMouseDown={onMouseDownCell}
      onMouseEnter={onMouseEnterCell}
      onClick={onClickCell}
    >
      <div className="relative">
        <div
          className={`w-full ${densityConfig.cellHeightClass} text-center bg-transparent border-none font-bold flex items-center justify-center ${tableFontSizeClass}`}
          style={{ color: showPreScheduleAsMain ? (preScheduleDisplayTextColor || preScheduleTextColor || tableFontColor) : (cellTextColor || tableFontColor), pointerEvents: 'none' }}
        >
          {showPreScheduleAsMain ? preScheduleCode : displayValue}
        </div>
        {showPreScheduleAsHint && (
          <div
            className="absolute left-0 top-0 w-0 h-0 z-20 pointer-events-none border-r-[9px] border-r-transparent border-b-[9px] border-b-transparent"
            style={{ borderTop: `9px solid ${preScheduleHintBorderColor}` }}
            title={`預班：${preScheduleCode}`}
          />
        )}
        {isRuleWarning && (
          <div
            className="absolute top-0 right-0 w-0 h-0 border-l-[10px] border-l-transparent z-20"
            style={{ borderTop: `10px solid ${warningTintColor}` }}
            title={ruleWarningMessage}
          />
        )}
        <div
          className="absolute right-1 top-1/2 -translate-y-1/2 z-10 w-3.5 h-3.5 flex items-center justify-center"
          title={selectTitle}
        >
          <select
            value={selectValue}
            onChange={onSelectChange}
            onClick={onSelectClick}
            onMouseDown={onSelectMouseDown}
            className="absolute inset-0 w-full h-full border-none bg-transparent cursor-pointer opacity-0"
            style={{ color: tableFontColor }}
            aria-label={selectTitle}
          >
            <option value=""></option>
            {!preScheduleEditMode && (
              <optgroup label="上班">
                {shiftOptionsNodes}
              </optgroup>
            )}
            <optgroup label={preScheduleEditMode ? "預班／預假" : "休假"}>
              {leaveOptionsNodes}
            </optgroup>
          </select>

          {showBlueDots && (
            <span
              className={`${selectorDotClass} rounded-full transition-all pointer-events-none ${inRangeSelection ? 'bg-violet-600 scale-110' : isPrimarySelected ? 'bg-blue-700 scale-110' : 'bg-blue-300/90'}`}
            ></span>
          )}
        </div>
      </div>
    </td>
  );
});


function ScheduleView({ changeScreen, colors, setColors, customHolidays, setCustomHolidays, specialWorkdays, setSpecialWorkdays, medicalCalendarAdjustments, setMedicalCalendarAdjustments, staffingConfig, setStaffingConfig, uiSettings, setUiSettings, customLeaveCodes, setCustomLeaveCodes, customWorkShifts, setCustomWorkShifts, customColumns, setCustomColumns, customColumnValues, setCustomColumnValues, schedulingRulesText, setSchedulingRulesText, loadLatestOnEnter, onLatestLoaded, importedSchedulePayload, onImportedScheduleApplied, monthlySchedules, setMonthlySchedules, preScheduleMonthlySchedules, setPreScheduleMonthlySchedules, importedPreSchedulePayload, onImportedPreScheduleApplied, pendingOpenMonthKey, onPendingOpenHandled, year, setYear, month, setMonth, staffs, setStaffs, schedule, setSchedule, onDownloadDraftFile, onImportDraftFileClick, draftImportInputRef, onImportDraftFileChange }) {
  // ==========================================
  // 2. 核心 State 定義
  // ==========================================
  const [isRuleFillLoading, setIsRuleFillLoading] = useState(false);


  const [unitAdjustmentDraft, setUnitAdjustmentDraft] = useState({ holidays: [], workdays: [] });

  const [showRuleFillControl, setShowRuleFillControl] = useState(false);
  const [ruleFillFeedback, setRuleFillFeedback] = useState("");

  const [showHistoryModal, setShowHistoryModal] = useState(false);
  const [historyList, setHistoryList] = useState([]);
  const [showDraftPrompt, setShowDraftPrompt] = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [selectedFillCell, setSelectedFillCell] = useState(null);
  const [fillCandidates, setFillCandidates] = useState([]);
  const [showFillModal, setShowFillModal] = useState(false);
  const [selectedGridCell, setSelectedGridCell] = useState(null);
  const [rangeClearMode, setRangeClearMode] = useState('preScheduleOnly');
  const [cellDrafts, setCellDrafts] = useState({});
  const [invalidCellKeys, setInvalidCellKeys] = useState({});
  const [cellRuleWarnings, setCellRuleWarnings] = useState({});
  const [importRuleViolations, setImportRuleViolations] = useState([]);
  const [showImportViolationList, setShowImportViolationList] = useState(false);
  const [showImportViolationSummary, setShowImportViolationSummary] = useState(false);
  const [rangeSelection, setRangeSelection] = useState(null);
  const [selectionAnchor, setSelectionAnchor] = useState(null);
  const [isRangeDragging, setIsRangeDragging] = useState(false);
  const [clipboardGrid, setClipboardGrid] = useState([]);
  const [keyInputBuffer, setKeyInputBuffer] = useState('');
  const [collapsedGroups, setCollapsedGroups] = useState({ 白班: false, 小夜: false, 大夜: false });
  const [draggingStaffId, setDraggingStaffId] = useState(null);
  const [dragOverTarget, setDragOverTarget] = useState(null);
  const [dragOverGroup, setDragOverGroup] = useState('');
  const [editingStaffId, setEditingStaffId] = useState(null);
  const [editingNameDraft, setEditingNameDraft] = useState('');
  const [preScheduleEditMode, setPreScheduleEditMode] = useState(false);
  const [inputAssist, setInputAssist] = useState({ type: '', message: '' });

  // 規則補空指定設定
  const [ruleFillConfig, setRuleFillConfig] = useState({
    selectedStaffs: [],
    dateRange: { start: 1, end: 31 },
    targetShift: ''
  });

  const tableFontSizeClass = getUiFontSizeClass(uiSettings?.tableFontSize);
  const shiftColumnFontSizeClass = getUiFontSizeClass(uiSettings?.shiftColumnFontSize);
  const nameDateColumnFontSizeClass = getUiFontSizeClass(uiSettings?.nameDateColumnFontSize);
  const shiftLabelFontSize = getShiftLabelFontSize(uiSettings?.shiftColumnFontSize);
  const shiftCellLabelFontSize = getShiftCellLabelFontSize(uiSettings?.shiftColumnFontSize);
  const densityConfig = getAdjustedDensityConfig(getUiDensityConfig(uiSettings?.tableDensity), uiSettings);
  const dynamicNameWidth = useMemo(() => {
    const longestNameLength = Math.max(
      ...((staffs || []).map((staff) => String(staff?.name || '').trim().length || 0)),
      0
    );
    const safeLength = Math.max(1, longestNameLength);
    const compactControlSpace = 54;
    const estimatedNameWidth = compactControlSpace + (safeLength * 16);
    return Math.max(96, Math.min(196, estimatedNameWidth));
  }, [staffs]);
  const effectiveDensityConfig = useMemo(() => ({
    ...densityConfig,
    nameWidth: dynamicNameWidth
  }), [densityConfig, dynamicNameWidth]);
  const monthLoadSkipRef = useRef(false);
  const initializedMonthRef = useRef(false);
  const monthSwitchSeedRef = useRef('');
  const keyInputTimerRef = useRef(null);
  const inputAssistTimerRef = useRef(null);
  const tableScrollContainerRef = useRef(null);

  const pageBackgroundColor = uiSettings?.pageBackgroundColor || '#f8fafc';
  const tableFontColor = uiSettings?.tableFontColor || '#1f2937';
  const shiftColumnFontColor = uiSettings?.shiftColumnFontColor || '#1e293b';
  const nameDateColumnFontColor = uiSettings?.nameDateColumnFontColor || '#1e293b';
  const shiftColumnBgColor = uiSettings?.shiftColumnBgColor || '#ffffff';
  const nameDateColumnBgColor = uiSettings?.nameDateColumnBgColor || '#ffffff';
  const demandOverColor = uiSettings?.demandOverColor || '#fde68a';
  const groupSummaryRowBgColor = uiSettings?.groupSummaryRowBgColor || '#fef3c7';
  const warningTintColor = uiSettings?.warningTintColor || '#f59e0b';
  const warningTextColor = uiSettings?.warningTextColor || '#92400e';
  const infoTintColor = uiSettings?.infoTintColor || '#38bdf8';
  const infoTextColor = uiSettings?.infoTextColor || '#075985';
  const dangerTintColor = uiSettings?.dangerTintColor || '#ef4444';
  const dangerTextColor = uiSettings?.dangerTextColor || '#9f1239';
  const warningSoftBgColor = blendHexColors(pageBackgroundColor, warningTintColor, 0.16);
  const warningSoftBorderColor = blendHexColors(pageBackgroundColor, warningTintColor, 0.35);
  const infoSoftBgColor = blendHexColors(pageBackgroundColor, infoTintColor, 0.16);
  const infoSoftBorderColor = blendHexColors(pageBackgroundColor, infoTintColor, 0.35);
  const dangerSoftBgColor = blendHexColors(pageBackgroundColor, dangerTintColor, 0.14);
  const dangerSoftBorderColor = blendHexColors(pageBackgroundColor, dangerTintColor, 0.32);
  const preScheduleTintColor = uiSettings?.preScheduleTintColor || infoTintColor;
  const preScheduleTextColor = uiSettings?.preScheduleTextColor || infoTextColor;
  const preScheduleMainBgColor = blendHexColors(pageBackgroundColor, preScheduleTintColor, 0.18);
  const preScheduleHintBgColor = blendHexColors(pageBackgroundColor, preScheduleTintColor, 0.12);
  const preScheduleHintBorderColor = blendHexColors(pageBackgroundColor, preScheduleTintColor, 0.32);
  const stickyGroupSummaryTop = 44;
  const stickyGroupSummaryShadow = '0 6px 12px rgba(15, 23, 42, 0.08)';
  const fourWeekDividerBaseColor = nameDateColumnFontColor || shiftColumnFontColor || tableFontColor || '#1e293b';
  const fourWeekDividerColor = blendHexColors(fourWeekDividerBaseColor, pageBackgroundColor, 0.18);

  const getFourWeekDividerStyle = (dateStr) => (
    isFourWeekCycleEndDate(dateStr)
      ? { boxShadow: `inset -3px 0 0 ${fourWeekDividerColor}` }
      : null
  );

  const getWordCycleDividerStyle = (dateStr) => (
    isFourWeekCycleEndDate(dateStr)
      ? `border-right:3pt solid ${fourWeekDividerColor};mso-border-right-alt:3pt solid ${fourWeekDividerColor};`
      : ''
  );

  const applyExcelFourWeekDivider = (border = {}, dateStr) => {
    if (!isFourWeekCycleEndDate(dateStr)) return border;
    return {
      ...border,
      right: {
        style: 'thick',
        color: { argb: hexToExcelArgb(fourWeekDividerColor, '#64748B') }
      }
    };
  };
  const showRightStats = uiSettings?.showRightStats ?? uiSettings?.showStats ?? true;
  const showLeaveStats = uiSettings?.showLeaveStats ?? uiSettings?.showStats ?? true;
  const showBottomStats = uiSettings?.showBottomStats ?? true;
  const showBlueDots = uiSettings?.showBlueDots ?? true;
  const showShiftLabels = uiSettings?.showShiftLabels ?? true;
  const defaultAutoLeaveCode = uiSettings?.defaultAutoLeaveCode || 'off';
  const selectionMode = uiSettings?.selectionMode || 'dot';
  const mergedLeaveCodes = useMemo(() => Array.from(new Set([...DICT.LEAVES, ...(customLeaveCodes || [])])).filter(Boolean), [customLeaveCodes]);
  const mergedShiftCodes = useMemo(() => Array.from(new Set([...(DICT.SHIFTS || []), ...((customWorkShifts || []).map(item => String(item?.code || '').trim()).filter(Boolean))])).filter(Boolean), [customWorkShifts]);
  const isConfiguredLeaveCode = (code = '') => {
    if (!code) return false;
    const prefix = getCodePrefix(code);
    return mergedLeaveCodes.includes(code) || mergedLeaveCodes.includes(prefix);
  };

  // ==========================================
  // 3. 初始載入與自動帶入
  // ==========================================
  useEffect(() => {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
      try {
        const parsed = JSON.parse(stored);
        if (parsed && parsed.length > 0) {
          setHistoryList(parsed);
        }
      } catch (e) {
        console.error("本機暫存紀錄解析失敗");
      }
    }
  }, []);

  useEffect(() => {
    if (!loadLatestOnEnter) return;

    const stored = localStorage.getItem(STORAGE_KEY);
    if (!stored) {
      onLatestLoaded?.();
      return;
    }

    try {
      const parsed = JSON.parse(stored);
      if (parsed && parsed.length > 0) {
        setHistoryList(parsed);
        loadHistory(parsed[0]);
      }
    } catch (e) {
      console.error("自動載入最新本機暫存紀錄失敗");
    } finally {
      onLatestLoaded?.();
    }
  }, [loadLatestOnEnter, onLatestLoaded]);

  useEffect(() => {
    if (!pendingOpenMonthKey) return;
    const [nextYear, nextMonth] = String(pendingOpenMonthKey).split('-').map(Number);
    if (!Number.isFinite(nextYear) || !Number.isFinite(nextMonth)) {
      onPendingOpenHandled?.();
      return;
    }
    monthSwitchSeedRef.current = pendingOpenMonthKey;
    setYear(nextYear);
    setMonth(nextMonth);
    onPendingOpenHandled?.();
  }, [pendingOpenMonthKey, onPendingOpenHandled]);

  useEffect(() => {
    if (!importedSchedulePayload || !importedSchedulePayload.monthlySchedules) return;

    const mergedSchedules = {
      ...(monthlySchedules || {}),
      ...importedSchedulePayload.monthlySchedules
    };

    setMonthlySchedules(mergedSchedules);

    const totalMonths = Object.keys(importedSchedulePayload.monthlySchedules || {}).length;
    const targetMonthKey = pendingOpenMonthKey || importedSchedulePayload.firstMonthKey || buildMonthKey(year, month);
    const [targetYear, targetMonth] = String(targetMonthKey).split('-').map(Number);

    if (Number.isFinite(targetYear) && Number.isFinite(targetMonth)) {
      monthSwitchSeedRef.current = targetMonthKey;
      if (year !== targetYear) setYear(targetYear);
      if (month !== targetMonth) setMonth(targetMonth);
      loadMonthState(targetYear, targetMonth, mergedSchedules);
      initializedMonthRef.current = true;
    }

    const importedMonthKeys = Object.keys(importedSchedulePayload.monthlySchedules || {});
    const scannedViolations = scanScheduleRuleViolations(mergedSchedules, { targetMonthKeys: importedMonthKeys });
    setImportRuleViolations(scannedViolations);
    setShowImportViolationList(false);
    setShowImportViolationSummary(scannedViolations.length > 0);

    onImportedScheduleApplied?.();
  }, [importedSchedulePayload, monthlySchedules, onImportedScheduleApplied, pendingOpenMonthKey, setMonthlySchedules, year, month]);


  useEffect(() => {
    if (!importedPreSchedulePayload || !importedPreSchedulePayload.monthlySchedules) return;

    const targetMonthKey = pendingOpenMonthKey || importedPreSchedulePayload.firstMonthKey || buildMonthKey(year, month);
    const [targetYear, targetMonth] = String(targetMonthKey).split('-').map(Number);
    if (!Number.isFinite(targetYear) || !Number.isFinite(targetMonth)) return;

    monthSwitchSeedRef.current = targetMonthKey;
    if (year !== targetYear) setYear(targetYear);
    if (month !== targetMonth) setMonth(targetMonth);

    onImportedPreScheduleApplied?.();
  }, [importedPreSchedulePayload, pendingOpenMonthKey]);

  useEffect(() => {
    const currentKey = buildMonthKey(year, month);
    if (!initializedMonthRef.current) {
      loadMonthState(year, month);
      initializedMonthRef.current = true;
      monthSwitchSeedRef.current = currentKey;
      return;
    }

    if (monthSwitchSeedRef.current !== currentKey) {
      loadMonthState(year, month);
      monthSwitchSeedRef.current = currentKey;
    }
  }, [year, month]);

  useEffect(() => {
    if (monthLoadSkipRef.current) return;
    const monthKey = buildMonthKey(year, month);
    setMonthlySchedules(prev => ({
      ...prev,
      [monthKey]: {
        ...(prev?.[monthKey] || {}),
        year,
        month,
        staffs: normalizeStaffGroup(staffs),
        scheduleData: schedule,
        customColumnValues: customColumnValues || {},
        schedulingRulesText: typeof schedulingRulesText === 'string' ? schedulingRulesText : '',
        importMeta: {
          ...(prev?.[monthKey]?.importMeta || {}),
          sourceType: prev?.[monthKey]?.importMeta?.sourceType || 'manual',
          sourceFiles: prev?.[monthKey]?.importMeta?.sourceFiles || [],
          sourceSheets: prev?.[monthKey]?.importMeta?.sourceSheets || [],
          importedAt: prev?.[monthKey]?.importMeta?.importedAt || new Date().toISOString(),
          lastUpdatedAt: new Date().toISOString()
        }
      }
    }));
  }, [year, month, staffs, schedule, customColumnValues, schedulingRulesText, setMonthlySchedules]);


  const holidayCalendar = useMemo(() => {
    return getSystemHolidayCalendar(year, {
      customHolidays,
      specialWorkdays,
      unitAdjustments: medicalCalendarAdjustments
    });
  }, [year, customHolidays, specialWorkdays, medicalCalendarAdjustments]);

  const holidays = holidayCalendar.holidays;
  const workdays = holidayCalendar.workdays;

  const daysInMonth = useMemo(() => {
    const days = [];
    const daysCount = new Date(year, month, 0).getDate();
    const weekNames = ['日', '一', '二', '三', '四', '五', '六'];
    const holidaySet = new Set(holidays);
    const workdaySet = new Set(workdays);

    for (let i = 1; i <= daysCount; i++) {
      const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
      const weekNum = new Date(year, month - 1, i).getDay();
      const rawWeekend = weekNum === 0 || weekNum === 6;
      const isAdjustedWorkday = workdaySet.has(dateStr);
      days.push({
        day: i,
        date: dateStr,
        weekStr: weekNames[weekNum],
        isWeekend: rawWeekend && !isAdjustedWorkday,
        isHoliday: holidaySet.has(dateStr),
        isAdjustedWorkday
      });
    }
    return days;
  }, [year, month, holidays, workdays]);

  const requiredLeaves = useMemo(
    () => daysInMonth.filter(d => d.isWeekend || d.isHoliday).length,
    [daysInMonth]
  );


  useEffect(() => {
    const currentMonthPrefix = buildMonthKey(year, month);
    setCellRuleWarnings(prev => {
      const next = { ...prev };
      Object.keys(next).forEach((cellKey) => {
        if (String(cellKey).includes('__') && String(cellKey).split('__')[1]?.startsWith(currentMonthPrefix)) {
          delete next[cellKey];
        }
      });

      const monthSnapshot = schedule || {};
      (staffs || []).forEach((staff) => {
        daysInMonth.forEach((day) => {
          const reasons = evaluateRuleWarningForCellInSnapshot(monthSnapshot, staff, day.date);
          if (reasons.length > 0) {
            next[makeCellKey(staff.id, day.date)] = reasons[0];
          }
        });
      });

      importRuleViolations
        .filter((item) => String(item.dateStr || '').startsWith(currentMonthPrefix))
        .forEach((item) => {
          const matchedStaff = staffs.find((staff) => String(staff.name || '').trim() === String(item.staffName || '').trim());
          if (matchedStaff) next[makeCellKey(matchedStaff.id, item.dateStr)] = item.reason;
        });
      return next;
    });
  }, [importRuleViolations, year, month, staffs, schedule, daysInMonth]);

  const loadMonthState = (targetYear, targetMonth, schedulesSource = monthlySchedules) => {
    monthLoadSkipRef.current = true;

    const nextMonthState = buildMonthStatePayload({
      targetYear,
      targetMonth,
      schedulesSource,
      preScheduleMonthlySchedules,
      schedulingRulesText,
      readSchedulingRulesTextFromLocalSettings,
      normalizeStaffGroup,
      createBlankMonthState
    });

    setStaffs(nextMonthState.staffs);
    setSchedule(nextMonthState.schedule);
    setCustomColumnValues(nextMonthState.customColumnValues);
    setSchedulingRulesText(nextMonthState.schedulingRulesText);

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


  const currentMonthKey = buildMonthKey(year, month);

  const getRuleContextMonthState = (monthKey, options = {}) => {
    if (!monthKey) return null;
    if (monthKey === currentMonthKey) {
      return {
        year,
        month,
        staffs,
        scheduleData: options.snapshot || schedule
      };
    }
    return monthlySchedules?.[monthKey] || null;
  };

  const findComparableStaffInMonthState = (targetStaff, monthState) => {
    if (!targetStaff || !monthState?.staffs) return null;
    const monthStaffs = Array.isArray(monthState.staffs) ? monthState.staffs : [];
    const byId = monthStaffs.find((staff) => staff.id === targetStaff.id);
    if (byId) return byId;

    const targetName = String(targetStaff.name || '').trim();
    const targetGroup = targetStaff.group || '白班';
    if (!targetName) return null;

    return monthStaffs.find((staff) => {
      const candidateName = String(staff.name || '').trim();
      const candidateGroup = staff.group || '白班';
      return candidateName === targetName && candidateGroup === targetGroup;
    }) || monthStaffs.find((staff) => String(staff.name || '').trim() === targetName) || null;
  };

  const getStaffRefFromCurrentMonth = (staffOrId) => {
    if (!staffOrId) return null;
    if (typeof staffOrId === 'object') return staffOrId;
    return staffs.find((staff) => staff.id === staffOrId) || null;
  };

  const getContextCellData = (staffOrId, dateStr, options = {}) => {
    if (!dateStr) return null;
    const monthKey = String(dateStr).slice(0, 7);
    const monthState = getRuleContextMonthState(monthKey, options);
    if (!monthState) return null;

    const targetStaff = getStaffRefFromCurrentMonth(staffOrId);
    const matchedStaff = findComparableStaffInMonthState(targetStaff, monthState) || (typeof staffOrId === 'string'
      ? (Array.isArray(monthState.staffs) ? monthState.staffs.find((staff) => staff.id === staffOrId) : null)
      : null);
    if (!matchedStaff) return null;

    const monthSchedule = monthState.scheduleData || monthState.schedule || {};
    return monthSchedule?.[matchedStaff.id]?.[dateStr] || null;
  };

  const getContextCellCode = (staffOrId, dateStr, options = {}) => {
    const cellData = getContextCellData(staffOrId, dateStr, options);
    return typeof cellData === 'object' && cellData !== null ? (cellData.value || '') : (cellData || '');
  };

  const getPreScheduleMonthState = (monthKey) => {
    if (!monthKey) return null;
    return preScheduleMonthlySchedules?.[monthKey] || null;
  };

  const getPreScheduleCellData = (staffOrId, dateStr) => {
    if (!dateStr) return null;
    const monthKey = String(dateStr).slice(0, 7);
    const monthState = getPreScheduleMonthState(monthKey);
    if (!monthState) return null;

    const targetStaff = getStaffRefFromCurrentMonth(staffOrId);
    const matchedStaff = findComparableStaffInMonthState(targetStaff, monthState) || (typeof staffOrId === 'string'
      ? (Array.isArray(monthState.staffs) ? monthState.staffs.find((staff) => staff.id === staffOrId) : null)
      : null);
    if (!matchedStaff) return null;

    const monthSchedule = monthState.scheduleData || monthState.schedule || {};
    return monthSchedule?.[matchedStaff.id]?.[dateStr] || null;
  };

  const getPreScheduleCellCode = (staffOrId, dateStr) => {
    const cellData = getPreScheduleCellData(staffOrId, dateStr);
    return typeof cellData === 'object' && cellData !== null ? (cellData.value || '') : (cellData || '');
  };

  const getPreScheduleCellTextColor = (staffOrId, dateStr) => {
    const cellData = getPreScheduleCellData(staffOrId, dateStr);
    if (!cellData || typeof cellData !== 'object') return '';
    return String(cellData.textColor || '').trim();
  };

  const getEffectiveCellTextColor = (staffOrId, dateStr) => {
    const formalTextColor = getCellTextColor(staffOrId, dateStr);
    if (formalTextColor) return formalTextColor;
    const preScheduleText = getPreScheduleCellTextColor(staffOrId, dateStr);
    return preScheduleText || '';
  };

  const getVisiblePreScheduleCode = (staffOrId, dateStr) => {
    const code = getPreScheduleCellCode(staffOrId, dateStr);
    return code || '';
  };

  const resolvePreScheduleMatchedStaff = (staffOrId, monthState) => {
    if (!monthState) return null;
    const targetStaff = getStaffRefFromCurrentMonth(staffOrId);
    return findComparableStaffInMonthState(targetStaff, monthState) || (typeof staffOrId === 'string'
      ? (Array.isArray(monthState?.staffs) ? monthState.staffs.find((staff) => staff.id === staffOrId) : null)
      : null);
  };

  const updatePreScheduleEntries = (entries = []) => {
    const normalizedEntries = Array.isArray(entries)
      ? entries.filter((entry) => entry?.staffId && entry?.dateStr)
      : [];
    if (normalizedEntries.length === 0) return 0;

    const entriesByMonth = normalizedEntries.reduce((acc, entry) => {
      const monthKey = String(entry.dateStr).slice(0, 7);
      if (!acc[monthKey]) acc[monthKey] = [];
      acc[monthKey].push(entry);
      return acc;
    }, {});

    let changedCount = 0;
    setPreScheduleMonthlySchedules((prev) => {
      const next = { ...(prev || {}) };

      Object.entries(entriesByMonth).forEach(([monthKey, monthEntries]) => {
        const [targetYear, targetMonth] = String(monthKey).split('-').map(Number);
        const baseMonthState = next[monthKey] || {
          year: targetYear,
          month: targetMonth,
          staffs: normalizeStaffGroup((staffs || []).map((staff) => ({
            id: staff.id,
            name: staff.name,
            group: staff.group || '白班',
            pregnant: Boolean(staff.pregnant)
          }))),
          scheduleData: {},
          customColumnValues: {},
          schedulingRulesText: '',
          importMeta: {
            sourceType: 'manualPreSchedule',
            sourceFiles: [],
            sourceSheets: [],
            importedAt: new Date().toISOString(),
            lastUpdatedAt: new Date().toISOString(),
            importMode: 'preSchedule'
          }
        };

        const monthState = {
          ...baseMonthState,
          staffs: Array.isArray(baseMonthState.staffs) && baseMonthState.staffs.length > 0
            ? baseMonthState.staffs
            : normalizeStaffGroup((staffs || []).map((staff) => ({
                id: staff.id,
                name: staff.name,
                group: staff.group || '白班',
                pregnant: Boolean(staff.pregnant)
              })))
        };

        const monthSchedule = JSON.parse(JSON.stringify(monthState.scheduleData || monthState.schedule || {}));
        monthEntries.forEach(({ staffId, dateStr, value }) => {
          const matchedStaff = resolvePreScheduleMatchedStaff(staffId, monthState)
            || monthState.staffs.find((staff) => staff.id === staffId)
            || (() => {
              const currentStaff = staffs.find((staff) => staff.id === staffId);
              if (!currentStaff) return null;
              const clonedStaff = {
                id: currentStaff.id,
                name: currentStaff.name,
                group: currentStaff.group || '白班',
                pregnant: Boolean(currentStaff.pregnant)
              };
              monthState.staffs = [...(monthState.staffs || []), clonedStaff];
              return clonedStaff;
            })();
          if (!matchedStaff) return;
          if (!monthSchedule[matchedStaff.id]) monthSchedule[matchedStaff.id] = {};
          const previousCell = monthSchedule[matchedStaff.id]?.[dateStr];
          const previousValue = typeof previousCell === 'object' && previousCell !== null ? (previousCell.value || '') : String(previousCell || '');
          if (value) {
            monthSchedule[matchedStaff.id][dateStr] = { value, source: 'manual' };
            if (previousValue !== value) changedCount += 1;
          } else if (previousValue) {
            delete monthSchedule[matchedStaff.id][dateStr];
            changedCount += 1;
          }
        });

        next[monthKey] = {
          ...monthState,
          year: targetYear,
          month: targetMonth,
          scheduleData: monthSchedule,
          importMeta: {
            ...(monthState.importMeta || {}),
            sourceType: monthState.importMeta?.sourceType || 'manualPreSchedule',
            importMode: 'preSchedule',
            lastUpdatedAt: new Date().toISOString()
          }
        };
      });

      return next;
    });

    return changedCount;
  };

  const getPreScheduleBackgroundColor = (baseColor = 'transparent', hasFormalValue = false) => {
    const normalizedBaseColor = baseColor && baseColor !== 'transparent' ? baseColor : pageBackgroundColor;
    return blendHexColors(normalizedBaseColor, preScheduleTintColor, hasFormalValue ? 0.12 : 0.2);
  };


  const getExportCellPresentation = (staffOrId, dayInfo) => {
    const dateStr = dayInfo?.date;
    const formalCode = getCellCode(staffOrId, dateStr) || '';
    const preScheduleCode = getVisiblePreScheduleCode(staffOrId, dateStr) || '';
    const formalTextColor = getCellTextColor(staffOrId, dateStr) || '';
    const preScheduleCellTextColor = getPreScheduleCellTextColor(staffOrId, dateStr) || '';
    const baseBackgroundColor = dayInfo?.isHoliday
      ? colors.holiday
      : (dayInfo?.isWeekend ? colors.weekend : pageBackgroundColor);
    const hasFormalValue = Boolean(formalCode);
    const hasPreSchedule = Boolean(preScheduleCode);
    const displayValue = hasFormalValue ? formalCode : preScheduleCode;
    const backgroundColor = hasPreSchedule
      ? getPreScheduleBackgroundColor(baseBackgroundColor, hasFormalValue)
      : baseBackgroundColor;
    const textColor = hasFormalValue
      ? (formalTextColor || tableFontColor)
      : (hasPreSchedule ? (preScheduleCellTextColor || preScheduleTextColor) : tableFontColor);

    return {
      formalCode,
      preScheduleCode,
      displayValue,
      backgroundColor,
      textColor,
      hasFormalValue,
      hasPreSchedule
    };
  };


  const getRawImportedLeaveSeed = (staffOrId, leaveType) => {
    if (!leaveType || !['例', '休'].includes(leaveType)) return 0;
    const monthStartDate = daysInMonth?.[0]?.date ? parseDateKey(daysInMonth[0].date) : null;
    if (!monthStartDate) return 0;

    let cursor = addDays(monthStartDate, -1);
    for (let i = 0; i < 62; i += 1) {
      const dateKey = formatDateKey(cursor);
      const rawCell = getContextCellData(staffOrId, dateKey);
      const rawImportedValue = typeof rawCell === 'object' && rawCell !== null ? String(rawCell.rawImportedValue || '').trim() : '';
      const match = rawImportedValue.match(/^(例|休)([1-4])$/);
      if (match && match[1] === leaveType) return Number(match[2]) || 0;
      cursor = addDays(cursor, -1);
    }

    return 0;
  };

  const buildExportNumberedValueMap = (staffOrId) => {
    const valueMap = {};
    let leaveCounters = {
      例: 0,
      休: 0
    };
    let isFirstSegment = true;
    let firstSegmentSeed = {
      例: getRawImportedLeaveSeed(staffOrId, '例'),
      休: getRawImportedLeaveSeed(staffOrId, '休')
    };
    let firstSegmentCarryUsed = {
      例: false,
      休: false
    };

    daysInMonth.forEach((dayInfo, index) => {
      const presentation = getExportCellPresentation(staffOrId, dayInfo);
      const displayValue = presentation.displayValue || '';

      if (displayValue === '例' || displayValue === '休') {
        const leaveType = displayValue;

        if (isFirstSegment && Number(firstSegmentSeed[leaveType] || 0) > 0) {
          if (!firstSegmentCarryUsed[leaveType]) {
            const continuedCount = Math.min(Number(firstSegmentSeed[leaveType] || 0) + 1, 4);
            valueMap[dayInfo.date] = `${leaveType}${continuedCount}`;
            firstSegmentCarryUsed[leaveType] = true;
          } else {
            valueMap[dayInfo.date] = leaveType;
          }
        } else {
          const nextCount = Number(leaveCounters[leaveType] || 0) + 1;
          leaveCounters[leaveType] = nextCount;
          valueMap[dayInfo.date] = nextCount <= 4 ? `${leaveType}${nextCount}` : leaveType;
        }
      } else {
        valueMap[dayInfo.date] = displayValue;
      }

      if (isFourWeekCycleEndDate(dayInfo.date) && index < daysInMonth.length - 1) {
        leaveCounters = { 例: 0, 休: 0 };
        isFirstSegment = false;
        firstSegmentSeed = { 例: 0, 休: 0 };
        firstSegmentCarryUsed = { 例: false, 休: false };
      }
    });

    return valueMap;
  };

  const getExportNumberedValue = (staffOrId, dateStr) => {
    if (!dateStr) return '';
    const valueMap = buildExportNumberedValueMap(staffOrId);
    return valueMap[dateStr] || '';
  };

  const buildExportStaffStats = (staffId) => {
    const stats = {
      work: 0,
      holidayLeave: 0,
      totalLeave: 0,
      leaveDetails: Object.fromEntries(mergedLeaveCodes.map((leaveCode) => [leaveCode, 0]))
    };

    daysInMonth.forEach((dayInfo) => {
      const displayValue = getExportNumberedValue(staffId, dayInfo.date);
      if (!displayValue) return;
      if (getAllShiftCodes().includes(displayValue)) stats.work += 1;
      if (isConfiguredLeaveCode(displayValue)) {
        stats.totalLeave += 1;
        const leavePrefix = getCodePrefix(displayValue);
        if (stats.leaveDetails[leavePrefix] !== undefined) stats.leaveDetails[leavePrefix] += 1;
        if (dayInfo.isWeekend || dayInfo.isHoliday) stats.holidayLeave += 1;
      }
    });

    return stats;
  };

  const buildExportDailyStats = (dateStr) => {
    const stats = { D: 0, E: 0, N: 0, totalLeave: 0 };
    staffs.forEach((staff) => {
      const displayValue = getExportNumberedValue(staff.id, dateStr);
      if (!displayValue) return;

      const shiftGroup = getShiftGroupByCode(displayValue);
      if (shiftGroup === '白班') stats.D += 1;
      else if (shiftGroup === '小夜') stats.E += 1;
      else if (shiftGroup === '大夜') stats.N += 1;
      else if (isConfiguredLeaveCode(displayValue)) stats.totalLeave += 1;
    });
    return stats;
  };

  const clearRuleWarningCells = (cells = []) => {
    if (!Array.isArray(cells) || cells.length === 0) return;
    setCellRuleWarnings(prev => {
      const next = { ...prev };
      cells.forEach(({ staffId, dateStr }) => {
        delete next[makeCellKey(staffId, dateStr)];
      });
      return next;
    });
  };

  const setRuleWarningsForEntries = (warningEntries = []) => {
    const normalizedWarnings = Array.isArray(warningEntries) ? warningEntries : [];
    if (normalizedWarnings.length === 0) return;
    setCellRuleWarnings(prev => {
      const next = { ...prev };
      normalizedWarnings.forEach(({ staffId, dateStr, reasons = [] }) => {
        next[makeCellKey(staffId, dateStr)] = reasons?.[0] || '此格違反排班規則';
      });
      return next;
    });
  };

  const getCellCodeFromSnapshot = (snapshot = {}, staffId, dateStr) => {
    return getCellCodeFromSnapshotService({
      snapshot,
      staffId,
      dateStr,
      currentMonthKey,
      getContextCellCode
    });
  };

  const evaluateRuleWarningForCellInSnapshot = (snapshot = {}, staff, dateStr) => {
    return evaluateRuleWarningForCellInSnapshotService({
      snapshot,
      staff,
      dateStr,
      smartRules: SMART_RULES,
      isConfiguredLeaveCode,
      isShiftCode,
      parseDateKey,
      formatDateKey,
      addDays,
      getCellCodeFromSnapshot
    });
  };

  const refreshRuleWarningsForCells = (cells = [], snapshot = schedule) => {
    refreshRuleWarningsForCellsService({
      cells,
      snapshot,
      schedule,
      staffs,
      makeCellKey,
      setCellRuleWarnings,
      evaluateRuleWarningForCellInSnapshot
    });
  };

  const expandCellsForRuleRecheck = (cells = []) => {
    return expandCellsForRuleRecheckService({
      cells,
      daysInMonth
    });
  };

  const scanScheduleRuleViolations = (schedulesSource = {}, options = {}) => {
    return scanScheduleRuleViolationsService({
      schedulesSource,
      options,
      smartRules: SMART_RULES,
      isShiftCode,
      parseDateKey,
      formatDateKey,
      addDays
    });
  };

  const clearInvalidCellLater = (cellKey) => {
    window.setTimeout(() => {
      setInvalidCellKeys(prev => {
        const next = { ...prev };
        delete next[cellKey];
        return next;
      });
    }, 1200);
  };

  const flashInvalidSelection = (cells = []) => {
    cells.forEach(({ staffId, dateStr }) => {
      const cellKey = makeCellKey(staffId, dateStr);
      setInvalidCellKeys(prev => ({ ...prev, [cellKey]: true }));
      clearInvalidCellLater(cellKey);
    });
  };

  const showInputAssist = (message, type = 'error', duration = 1600) => {
    if (!message) return;
    setInputAssist({ type, message });
    if (inputAssistTimerRef.current) window.clearTimeout(inputAssistTimerRef.current);
    inputAssistTimerRef.current = window.setTimeout(() => {
      setInputAssist({ type: '', message: '' });
      inputAssistTimerRef.current = null;
    }, duration);
  };

  const clearInputAssist = () => {
    setInputAssist({ type: '', message: '' });
    if (inputAssistTimerRef.current) {
      window.clearTimeout(inputAssistTimerRef.current);
      inputAssistTimerRef.current = null;
    }
  };

  const resetKeyInputBuffer = () => {
    setKeyInputBuffer('');
    if (keyInputTimerRef.current) {
      window.clearTimeout(keyInputTimerRef.current);
      keyInputTimerRef.current = null;
    }
  };

  const keepKeyInputBufferAlive = () => {
    if (keyInputTimerRef.current) window.clearTimeout(keyInputTimerRef.current);
    keyInputTimerRef.current = window.setTimeout(() => {
      setKeyInputBuffer('');
      keyInputTimerRef.current = null;
    }, 1500);
  };

  const getEffectiveSelection = () => {
    if (rangeSelection?.start && rangeSelection?.end) return rangeSelection;
    if (selectedGridCell?.staff?.id && selectedGridCell?.dateStr) {
      return {
        start: { staffId: selectedGridCell.staff.id, dateStr: selectedGridCell.dateStr },
        end: { staffId: selectedGridCell.staff.id, dateStr: selectedGridCell.dateStr }
      };
    }
    return null;
  };

  const effectiveSelection = useMemo(() => {
    if (rangeSelection?.start && rangeSelection?.end) return rangeSelection;
    if (selectedGridCell?.staff?.id && selectedGridCell?.dateStr) {
      return {
        start: { staffId: selectedGridCell.staff.id, dateStr: selectedGridCell.dateStr, group: selectedGridCell.staff.group || '白班' },
        end: { staffId: selectedGridCell.staff.id, dateStr: selectedGridCell.dateStr, group: selectedGridCell.staff.group || '白班' }
      };
    }
    return null;
  }, [rangeSelection, selectedGridCell]);

  const selectedRangeCells = useMemo(
    () => expandSelectionCells(effectiveSelection, staffs, daysInMonth),
    [effectiveSelection, staffs, daysInMonth]
  );

  const selectedCellHasPreSchedule = useMemo(() => {
    if (!selectedGridCell?.staff?.id || !selectedGridCell?.dateStr) return false;
    return Boolean(getVisiblePreScheduleCode(selectedGridCell.staff.id, selectedGridCell.dateStr));
  }, [selectedGridCell?.staff?.id, selectedGridCell?.dateStr, preScheduleMonthlySchedules, staffs, year, month]);

  const selectedCellTextColor = useMemo(() => {
    if (!selectedGridCell?.staff?.id || !selectedGridCell?.dateStr) return '';
    return getEffectiveCellTextColor(selectedGridCell.staff.id, selectedGridCell.dateStr);
  }, [selectedGridCell?.staff?.id, selectedGridCell?.dateStr, schedule, monthlySchedules, preScheduleMonthlySchedules, staffs, year, month]);

  const selectedCellHasFormalValue = useMemo(() => {
    if (!selectedGridCell?.staff?.id || !selectedGridCell?.dateStr) return false;
    return Boolean(getCellCode(selectedGridCell.staff.id, selectedGridCell.dateStr));
  }, [selectedGridCell?.staff?.id, selectedGridCell?.dateStr, schedule, monthlySchedules, year, month]);

  const selectedCellSupportsTextColor = useMemo(() => {
    return selectedCellHasFormalValue || selectedCellHasPreSchedule;
  }, [selectedCellHasFormalValue, selectedCellHasPreSchedule]);

  const selectedRangeCellKeySet = useMemo(() => {
    return new Set(selectedRangeCells.map(({ staffId, dateStr }) => makeCellKey(staffId, dateStr)));
  }, [selectedRangeCells]);

  const primarySelectedCellKey = useMemo(() => {
    if (!selectedGridCell?.staff?.id || !selectedGridCell?.dateStr) return '';
    return makeCellKey(selectedGridCell.staff.id, selectedGridCell.dateStr);
  }, [selectedGridCell?.staff?.id, selectedGridCell?.dateStr]);

  const shiftOptionsNodes = useMemo(() => mergedShiftCodes.map(s => <option key={s} value={s}>{s}</option>), [mergedShiftCodes]);
  const leaveOptionsNodes = useMemo(() => mergedLeaveCodes.map(l => <option key={l} value={l}>{l}</option>), [mergedLeaveCodes]);

  useEffect(() => {
    if (selectedRangeCells.length > 0) clearInputAssist();
  }, [selectedGridCell?.staff?.id, selectedGridCell?.dateStr]);

  useEffect(() => {
    resetKeyInputBuffer();
    clearInputAssist();
  }, [preScheduleEditMode]);

  const setSelectionRangeFromCells = (cells = [], options = {}) => {
    if (!Array.isArray(cells) || cells.length === 0) return false;

    const cellMap = new Map(cells.map((cell) => [makeCellKey(cell.staffId, cell.dateStr), cell]));
    const orderedCells = expandSelectionCells(getEffectiveSelection(), staffs, daysInMonth);
    const orderedMatchedCells = orderedCells.filter((cell) => cellMap.has(makeCellKey(cell.staffId, cell.dateStr)));
    const targetCells = orderedMatchedCells.length > 0 ? orderedMatchedCells : cells;
    const firstCell = targetCells[0];
    const lastCell = targetCells[targetCells.length - 1];
    if (!firstCell || !lastCell) return false;

    const startStaff = staffs.find((staff) => staff.id === firstCell.staffId);
    const endStaff = staffs.find((staff) => staff.id === lastCell.staffId);
    if (!startStaff || !endStaff) return false;

    const anchorPoint = {
      staffId: firstCell.staffId,
      dateStr: firstCell.dateStr,
      group: startStaff.group || '白班'
    };
    const endPoint = {
      staffId: lastCell.staffId,
      dateStr: lastCell.dateStr,
      group: endStaff.group || '白班'
    };
    const activeCell = options.activeCell || lastCell;
    const activeStaff = staffs.find((staff) => staff.id === activeCell.staffId) || endStaff;

    setSelectionAnchor(anchorPoint);
    setRangeSelection({ start: anchorPoint, end: endPoint });
    setSelectedGridCell({ staff: activeStaff, dateStr: activeCell.dateStr });
    return true;
  };

  const clearSelectionContents = () => {
    if (selectedRangeCells.length === 0) return false;
    return applyScheduleEntries(
      selectedRangeCells.map(({ staffId, dateStr }) => ({
        staffId,
        dateStr,
        value: '',
        source: 'manual'
      })),
      {
        preserveSelection: true,
        selectionCells: selectedRangeCells,
        activeCell: selectedRangeCells[selectedRangeCells.length - 1]
      }
    );
  };

  const applyScheduleEntries = (entries = [], options = {}) => {
    if (!Array.isArray(entries) || entries.length === 0) return false;

    const normalizedEntries = entries
      .map((entry) => ({
        staffId: entry?.staffId,
        dateStr: entry?.dateStr,
        value: entry?.value || '',
        source: entry?.source || 'manual'
      }))
      .filter((entry) => entry.staffId && entry.dateStr);

    if (normalizedEntries.length === 0) return false;

    const dedupedEntries = Array.from(
      new Map(normalizedEntries.map((entry) => [makeCellKey(entry.staffId, entry.dateStr), entry])).values()
    );

    const touchedStaffIds = Array.from(new Set(dedupedEntries.map(({ staffId }) => staffId)));
    const nextSchedule = { ...(schedule || {}) };
    touchedStaffIds.forEach((staffId) => {
      const currentRow = schedule?.[staffId];
      nextSchedule[staffId] = currentRow && typeof currentRow === 'object' ? { ...currentRow } : {};
    });

    dedupedEntries.forEach(({ staffId, dateStr, value, source }) => {
      if (!nextSchedule[staffId]) nextSchedule[staffId] = {};
      const existingCell = nextSchedule[staffId]?.[dateStr];
      const existingMeta = typeof existingCell === 'object' && existingCell !== null ? existingCell : {};
      nextSchedule[staffId][dateStr] = value ? { ...existingMeta, value, source } : null;
    });

    setSchedule(nextSchedule);

    const recheckCells = expandCellsForRuleRecheck(dedupedEntries.map(({ staffId, dateStr }) => ({ staffId, dateStr })));
    refreshRuleWarningsForCells(recheckCells, nextSchedule);

    if (options.preserveSelection) {
      const selectionCells = Array.isArray(options.selectionCells) && options.selectionCells.length > 0
        ? options.selectionCells
        : dedupedEntries.map(({ staffId, dateStr }) => ({ staffId, dateStr }));
      setSelectionRangeFromCells(selectionCells, { activeCell: options.activeCell });
    }

    if (options.clearAssist !== false) clearInputAssist();
    if (options.resetBuffer !== false) resetKeyInputBuffer();

    return true;
  };


  const buildEntriesFromSnapshotDiff = (snapshot = {}, options = {}) => {
    const entries = [];
    const onlyCells = Array.isArray(options.onlyCells) ? new Set(options.onlyCells.map((cell) => makeCellKey(cell.staffId, cell.dateStr))) : null;

    staffs.forEach((staff) => {
      daysInMonth.forEach((day) => {
        const cellKey = makeCellKey(staff.id, day.date);
        if (onlyCells && !onlyCells.has(cellKey)) return;

        const currentCell = schedule[staff.id]?.[day.date];
        const nextCell = snapshot[staff.id]?.[day.date];
        const currentValue = typeof currentCell === 'object' && currentCell !== null ? (currentCell.value || '') : (currentCell || '');
        const currentSource = typeof currentCell === 'object' && currentCell !== null ? (currentCell.source || 'manual') : (currentValue ? 'manual' : '');
        const nextValue = typeof nextCell === 'object' && nextCell !== null ? (nextCell.value || '') : (nextCell || '');
        const nextSource = typeof nextCell === 'object' && nextCell !== null ? (nextCell.source || 'manual') : (nextValue ? 'manual' : '');

        if (currentValue === nextValue && currentSource === nextSource) return;
        entries.push({ staffId: staff.id, dateStr: day.date, value: nextValue, source: nextSource || 'manual' });
      });
    });

    return entries;
  };

  const applyRuleFillSnapshot = (snapshot = {}, options = {}) => {
    const entries = buildEntriesFromSnapshotDiff(snapshot, { onlyCells: options.onlyCells });
    if (entries.length === 0) return false;
    return applyScheduleEntries(entries, {
      preserveSelection: options.preserveSelection === true,
      selectionCells: options.selectionCells || entries.map(({ staffId, dateStr }) => ({ staffId, dateStr })),
      activeCell: options.activeCell || (entries.length > 0 ? { staffId: entries[entries.length - 1].staffId, dateStr: entries[entries.length - 1].dateStr } : null),
      clearAssist: options.clearAssist,
      resetBuffer: options.resetBuffer
    });
  };

  const applyRuleFillEntries = (entries = [], options = {}) => {
    if (!Array.isArray(entries) || entries.length === 0) return false;
    return applyScheduleEntries(
      entries.map((entry) => ({ ...entry, source: entry?.source || 'auto' })),
      {
        preserveSelection: options.preserveSelection === true,
        selectionCells: options.selectionCells || entries.map(({ staffId, dateStr }) => ({ staffId, dateStr })),
        activeCell: options.activeCell || (entries.length > 0 ? { staffId: entries[entries.length - 1].staffId, dateStr: entries[entries.length - 1].dateStr } : null),
        clearAssist: options.clearAssist,
        resetBuffer: options.resetBuffer
      }
    );
  };

  const filterChangedScheduleEntries = (entries = [], baseSchedule = schedule) => {
    return (Array.isArray(entries) ? entries : []).filter((entry) => {
      const currentCell = baseSchedule?.[entry?.staffId]?.[entry?.dateStr];
      const currentValue = typeof currentCell === 'object' && currentCell !== null ? (currentCell.value || '') : (currentCell || '');
      const currentSource = typeof currentCell === 'object' && currentCell !== null ? (currentCell.source || 'manual') : (currentValue ? 'manual' : '');
      const nextValue = String(entry?.value || '');
      const nextSource = entry?.source || (nextValue ? 'manual' : '');
      return currentValue !== nextValue || currentSource !== nextSource;
    });
  };

  const applyValueToCells = (cells, normalized, options = {}) => {
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
    const nonWarningCells = targetCells.filter(({ staffId, dateStr }) => !warningEntries.some((entry) => entry.staffId === staffId && entry.dateStr === dateStr));
    if (source === 'manual') clearRuleWarningCells(nonWarningCells);
    return applyScheduleEntries(
      allowedEntries,
      {
        ...options,
        clearAssist: options.clearAssist !== false && warningEntries.length === 0,
        preserveSelection: options.preserveSelection === true || warningEntries.length > 0,
        selectionCells: options.selectionCells || targetCells,
        activeCell: options.activeCell || targetCells[targetCells.length - 1]
      }
    );
  };

  const normalizePreScheduleInput = (rawValue = '') => {
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

  const buildPreSchedulePastePlan = (grid = [], rect = null) => {
    if (!rect || !Array.isArray(grid) || grid.length === 0) {
      return { entries: [], affectedCells: [], invalidCount: 0, clearCount: 0, writeCount: 0, clipped: false };
    }

    const selectionRowCount = rect.rowEnd - rect.rowStart + 1;
    const selectionColCount = rect.colEnd - rect.colStart + 1;
    const sourceRowCount = grid.length;
    const sourceColCount = Math.max(...grid.map(row => (row || []).length), 0);
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

    const entries = [];
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

        const rawValue = String(grid[sourceRow]?.[sourceCol] ?? '').trim();
        if (!rawValue) {
          entries.push({ ...targetCell, value: '' });
          clearCount += 1;
          continue;
        }

        const result = normalizePreScheduleInput(rawValue);
        if (!result.isValid) {
          invalidCount += 1;
          continue;
        }

        entries.push({ ...targetCell, value: result.normalized });
        writeCount += 1;
      }
    }

    return { entries, affectedCells, invalidCount, clearCount, writeCount, clipped };
  };

  const isPotentialPreSchedulePrefix = (rawValue = '') => {
    if (!String(rawValue ?? '').trim()) return true;
    return getNormalizedManualCodeCandidates(rawValue, [...mergedLeaveCodes, ...mergedShiftCodes]).length > 0;
  };

  const clearSelectedPreScheduleRangeByKeyboard = () => {
    if (!Array.isArray(selectedRangeCells) || selectedRangeCells.length === 0) return false;
    const changedCount = updatePreScheduleEntries(selectedRangeCells.map(({ staffId, dateStr }) => ({
      staffId,
      dateStr,
      value: ''
    })));
    if (changedCount <= 0) return false;
    setSelectionRangeFromCells(selectedRangeCells, { activeCell: selectedRangeCells[selectedRangeCells.length - 1] });
    clearInputAssist();
    resetKeyInputBuffer();
    return true;
  };

  const applyPreScheduleValueToCells = (cells = [], rawValue = '', options = {}) => {
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

  const moveSelectionAfterInput = (cells = [], direction = 1) => {
    if (!Array.isArray(cells) || cells.length !== 1) return false;
    return moveSelectedCell(0, direction);
  };

  const applySelectionValue = (cells = [], rawValue = '', options = {}) => {
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

  const navigateSelection = (rowDelta = 0, colDelta = 0) => {
    clearInputAssist();
    resetKeyInputBuffer();
    return moveSelectedCell(rowDelta, colDelta);
  };

  const tryApplyBufferedCode = (buffer) => {
    if (!buffer || selectedRangeCells.length === 0) return false;
    return (preScheduleEditMode
      ? applyPreScheduleValueToCells(selectedRangeCells, buffer, { advance: true, showFeedback: true })
      : applySelectionValue(selectedRangeCells, buffer, { advance: true })
    ).applied;
  };

  const applyShortcutCodeToSelection = (shortcutCode) => {
    if (!shortcutCode || selectedRangeCells.length === 0) return false;
    return (preScheduleEditMode
      ? applyPreScheduleValueToCells(selectedRangeCells, shortcutCode, { advance: true, showFeedback: true })
      : applySelectionValue(selectedRangeCells, shortcutCode, { advance: true })
    ).applied;
  };

  const moveSelectedCell = (rowDelta = 0, colDelta = 0) => {
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

  const commitCellValue = (staffId, dateStr, rawValue) => {
    const cellKey = makeCellKey(staffId, dateStr);
    const { normalized, isValid } = normalizeManualShiftCode(rawValue, [...mergedLeaveCodes, ...mergedShiftCodes]);

    if (!isValid) {
      setInvalidCellKeys(prev => ({ ...prev, [cellKey]: true }));
      clearInvalidCellLater(cellKey);
      setCellDrafts(prev => {
        const next = { ...prev };
        delete next[cellKey];
        return next;
      });
      return false;
    }

    const result = applySelectionValue([{ staffId, dateStr }], normalized, { advance: true });
    if (!result.applied) return false;

    setCellDrafts(prev => {
      const next = { ...prev };
      delete next[cellKey];
      return next;
    });
    setInvalidCellKeys(prev => {
      const next = { ...prev };
      delete next[cellKey];
      return next;
    });
    return true;
  };

  const startRangeSelection = (staff, dateStr, event = {}) => {
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

  const updateRangeSelection = (staff, dateStr) => {
    if (!isRangeDragging || !selectionAnchor) return;
    const anchorGroup = selectionAnchor.group || '白班';
    const targetGroup = staff.group || '白班';
    if (anchorGroup !== targetGroup) return;
    setRangeSelection({ start: selectionAnchor, end: { staffId: staff.id, dateStr, group: targetGroup } });
  };

  const copySelectionToClipboard = async () => {
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
    const text = grid.map(row => row.join('\t')).join('\n');
    try {
      if (navigator?.clipboard?.writeText) await navigator.clipboard.writeText(text);
    } catch (error) {
      console.error('寫入剪貼簿失敗', error);
    }
  };

  const buildPastePlan = (grid = [], rect = null) => {
    if (!rect || !Array.isArray(grid) || grid.length === 0) {
      return { updates: [], affectedCells: [], invalidCount: 0, clearCount: 0, writeCount: 0, clipped: false };
    }

    const selectionRowCount = rect.rowEnd - rect.rowStart + 1;
    const selectionColCount = rect.colEnd - rect.colStart + 1;
    const sourceRowCount = grid.length;
    const sourceColCount = Math.max(...grid.map(row => (row || []).length), 0);
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

  const pasteGridToSelection = async () => {
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

  useEffect(() => {
    const stopDrag = () => setIsRangeDragging(false);
    window.addEventListener('mouseup', stopDrag);
    return () => window.removeEventListener('mouseup', stopDrag);
  }, []);

  useEffect(() => {
    return () => {
      if (inputAssistTimerRef.current) window.clearTimeout(inputAssistTimerRef.current);
      if (keyInputTimerRef.current) window.clearTimeout(keyInputTimerRef.current);
    };
  }, []);

  useEffect(() => {
    const handleGlobalKeyDown = (event) => {
      const target = event.target;
      const tagName = String(target?.tagName || '').toLowerCase();
      const isTypingTarget = tagName === 'input' || tagName === 'textarea' || tagName === 'select' || target?.isContentEditable;
      const hasSelection = selectedRangeCells.length > 0;

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c' && hasSelection) {
        event.preventDefault();
        copySelectionToClipboard();
        return;
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'v' && hasSelection) {
        event.preventDefault();
        pasteGridToSelection();
        return;
      }

      if (!hasSelection) return;
      if (isTypingTarget) return;

      const shortcutCodeMap = {
        F1: 'D',
        F2: 'E',
        F3: 'N',
        F4: defaultAutoLeaveCode || 'off'
      };
      if (shortcutCodeMap[event.key]) {
        event.preventDefault();
        applyShortcutCodeToSelection(shortcutCodeMap[event.key]);
        return;
      }

      if (event.key === 'Escape') {
        event.preventDefault();
        if (keyInputBuffer) {
          resetKeyInputBuffer();
          clearInputAssist();
          return;
        }
        clearInputAssist();
        setRangeSelection(null);
        setSelectionAnchor(null);
        setSelectedGridCell(null);
        return;
      }

      if (event.key === 'Backspace' || event.key === 'Delete') {
        event.preventDefault();

        if (event.key === 'Backspace' && keyInputBuffer) {
          const nextBuffer = keyInputBuffer.slice(0, -1);
          if (!nextBuffer) {
            resetKeyInputBuffer();
            clearInputAssist();
            return;
          }
          setKeyInputBuffer(nextBuffer);
          keepKeyInputBufferAlive();
          if (!isPotentialManualShiftPrefix(nextBuffer, [...mergedLeaveCodes, ...mergedShiftCodes])) {
            flashInvalidSelection(selectedRangeCells);
          } else {
            clearInputAssist();
          }
          return;
        }

        if (preScheduleEditMode) {
          clearSelectedPreScheduleRangeByKeyboard();
        } else {
          clearSelectionContents();
        }
        return;
      }

      if (event.key === 'ArrowLeft') {
        event.preventDefault();
        navigateSelection(0, -1);
        return;
      }

      if (event.key === 'ArrowRight') {
        event.preventDefault();
        navigateSelection(0, 1);
        return;
      }

      if (event.key === 'ArrowUp') {
        event.preventDefault();
        navigateSelection(-1, 0);
        return;
      }

      if (event.key === 'ArrowDown') {
        event.preventDefault();
        navigateSelection(1, 0);
        return;
      }

      if (event.key === 'Enter') {
        event.preventDefault();
        navigateSelection(0, event.shiftKey ? -1 : 1);
        return;
      }

      if (event.key === 'Tab') {
        event.preventDefault();
        navigateSelection(0, event.shiftKey ? -1 : 1);
        return;
      }

      if (event.key.length === 1 && !event.altKey && !event.ctrlKey && !event.metaKey) {
        event.preventDefault();
        const nextBuffer = `${keyInputBuffer}${event.key}`;
        const applied = tryApplyBufferedCode(nextBuffer);
        if (applied) return;

        const prefixValid = preScheduleEditMode
          ? isPotentialPreSchedulePrefix(nextBuffer)
          : isPotentialManualShiftPrefix(nextBuffer, [...mergedLeaveCodes, ...mergedShiftCodes]);

        if (!prefixValid) {
          flashInvalidSelection(selectedRangeCells);
          resetKeyInputBuffer();
          if (preScheduleEditMode) showInputAssist('預班只能輸入預假代號', 'error');
          return;
        }

        clearInputAssist();
        setKeyInputBuffer(nextBuffer);
        keepKeyInputBufferAlive();
      }
    };

    window.addEventListener('keydown', handleGlobalKeyDown);
    return () => window.removeEventListener('keydown', handleGlobalKeyDown);
  }, [selectedRangeCells, keyInputBuffer, clipboardGrid, rangeSelection, selectedGridCell, staffs, daysInMonth, mergedLeaveCodes, defaultAutoLeaveCode]);

  // ==========================================
  // 4. Excel 匯出 (ExcelJS 實現)
  // ==========================================
  const exportToExcel = async () => {
    return exportToExcelService({
      loadExcelJS, year, month, uiSettings, blendHexColors, colors, hexToExcelArgb, mergedLeaveCodes, customColumns, daysInMonth, requiredLeaves, getExportNumberedValue, customColumnValues, getExportCellPresentation, buildExportStaffStats, groupedStaffs, buildExportDailyStats, applyExcelFourWeekDivider, tableFontColor, setShowExportMenu
    });
  };

  const formatWordDayCellValue = (value = "") => {
    const text = String(value || '').trim();
    if (/^(例|休)[1-4]$/.test(text)) {
      return `<span class="word-numbered-leave">${text}</span>`;
    }
    return text;
  };

  const exportToWord = () => {
    return exportToWordService({
      year, month, uiSettings, blendHexColors, colors, daysInMonth, requiredLeaves, schedulingRulesText, SHIFT_GROUPS, staffs, getExportCellPresentation, getExportNumberedValue, buildExportStaffStats, buildExportDailyStats, getWordCycleDividerStyle, tableFontColor, setShowExportMenu
    });
  };

  // ==========================================
  // 5. 規則式半智慧補班功能
  // ==========================================
  const handleRuleBasedAutoSchedule = createHandleRuleBasedAutoSchedule({
    setIsRuleFillLoading,
    setRuleFillFeedback,
    schedule,
    ruleFillConfig,
    staffs,
    daysInMonth,
    SHIFT_GROUPS,
    RULE_FILL_MAIN_SHIFTS,
    getShiftGroupByCode,
    GROUP_TO_DEMAND_KEY,
    staffingConfig,
    SMART_RULES,
    getVisiblePreScheduleCode,
    isConfiguredLeaveCode,
    requiredLeaves,
    isShiftCode,
    isLeaveCode,
    DEFAULT_SHIFT_BY_GROUP,
    makeCellKey,
    getContextCellCode,
    parseDateKey,
    addDays,
    formatDateKey,
    getRequiredStaffingBucketByDay,
    getStaffRefFromCurrentMonth,
    buildEntriesFromSnapshotDiff,
    applyRuleFillEntries,
    saveToHistory,
    defaultAutoLeaveCode
  });


  // ==========================================
  // 6. 輔助統計與操作
  // ==========================================
  const getStaffStats = (staffId) => {
    return staffStatsMap[staffId] || emptyStaffStats;
  };

  const getDailyStats = (dateStr) => {
    return dailyStatsMap[dateStr] || { D: 0, E: 0, N: 0, totalLeave: 0 };
  };

  const getRequiredCountForDate = (dateStr, rowKey) => {
    return requiredCountMap?.[dateStr]?.[rowKey] ?? null;
  };

  const getDemandHighlightStyle = (dateStr, rowKey, actualCount) => {
    return getDemandHighlightStyleService({
      dateStr,
      rowKey,
      actualCount,
      requiredCountMap,
      demandOverColor
    });
  };

  const saveToHistory = (label, currentSchedule = schedule) => {
    return saveToHistoryService(label, currentSchedule, {
      buildMonthKey, year, month, normalizeStaffGroup, staffs, monthlySchedules, customColumnValues, schedulingRulesText, preScheduleMonthlySchedules, colors, customHolidays, specialWorkdays, medicalCalendarAdjustments, staffingConfig, uiSettings, customLeaveCodes, customWorkShifts, customColumns, setHistoryList, storageKey: STORAGE_KEY, schedule
    });
  };

  const loadHistory = (record) => {
    return loadHistoryService(record, {
      buildMonthKey, setMonthlySchedules, setPreScheduleMonthlySchedules, setYear, setMonth, setCustomHolidays, setSpecialWorkdays, setMedicalCalendarAdjustments, setStaffingConfig, setUiSettings, setCustomLeaveCodes, setCustomWorkShifts, setCustomColumns, setCustomColumnValues, setSchedulingRulesText, setStaffs, normalizeStaffGroup, setSchedule, setColors, setShowHistoryModal, setShowDraftPrompt
    });
  };

  const clearHistory = () => {
    return clearHistoryService({ setHistoryList, storageKey: STORAGE_KEY });
  };


  const canAssignWithManualOverride = (staff, dateStr, shiftCode) => {
    const reasons = [];

    const prevKey = formatDateKey(addDays(parseDateKey(dateStr), -1));
    const prevCode = getContextCellCode(staff, prevKey);
    const disallowed = SMART_RULES.disallowedNextShiftMap[prevCode] || [];
    if (disallowed.includes(shiftCode)) {
      reasons.push(`${prevCode} 後不可接 ${shiftCode}`);
    }

    const existingCode = getContextCellCode(staff, dateStr);
    const existingIsShift = isShiftCode(existingCode);
    const consecutiveBefore = countConsecutiveWorkDaysBefore(staff.id, dateStr);
    const effectiveConsecutive = consecutiveBefore + (existingIsShift ? 0 : 1);
    if (effectiveConsecutive > SMART_RULES.maxConsecutiveWorkDays) {
      reasons.push(`連續上班不可超過 ${SMART_RULES.maxConsecutiveWorkDays} 天`);
    }

    if (staff.pregnant && SMART_RULES.pregnancyRestrictedShifts.includes(shiftCode)) {
      reasons.push('懷孕標記人員不可排 N / 夜8-8');
    }

    return { allowed: reasons.length === 0, reasons };
  };

  const validateManualEntries = (entries = [], options = {}) => {
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

      const staff = staffs.find((item) => item.id === normalizedEntry.staffId);
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

  const handleCellChange = (staffId, dateStr, value, options = {}) => {
    const source = options.source || 'manual';
    const rawEntries = [{ staffId, dateStr, value, source }];
    const { allowedEntries, warningEntries } = source === 'manual'
      ? validateManualEntries(rawEntries, { showFeedback: options.clearAssist !== false })
      : { allowedEntries: rawEntries, warningEntries: [] };
    if (allowedEntries.length === 0) return false;
    const nonWarningCells = rawEntries.filter(({ staffId, dateStr }) => !warningEntries.some((entry) => entry.staffId === staffId && entry.dateStr === dateStr));
    if (source === 'manual') clearRuleWarningCells(nonWarningCells);
    if (!options.allowWarningAssistClear && warningEntries.length > 0) options = { ...options, clearAssist: false };
    return applyScheduleEntries(
      allowedEntries,
      {
        clearAssist: options.clearAssist !== false,
        resetBuffer: options.resetBuffer !== false,
        preserveSelection: options.preserveSelection === true || warningEntries.length > 0,
        selectionCells: options.selectionCells || [{ staffId, dateStr }],
        activeCell: options.activeCell || { staffId, dateStr }
      }
    );
  };

  const applyCellTextColor = (staffId, dateStr, textColor = '') => {
    const currentCell = getContextCellData(staffId, dateStr);
    const normalizedValue = typeof currentCell === 'object' && currentCell !== null ? (currentCell.value || '') : (currentCell || '');
    const normalizedSource = typeof currentCell === 'object' && currentCell !== null ? (currentCell.source || 'manual') : 'manual';
    const preScheduleCode = getVisiblePreScheduleCode(staffId, dateStr);
    const nextColor = String(textColor || '').trim();

    if (!normalizedValue && !preScheduleCode) {
      showInputAssist('請先在此格填入班別、假別或預班，再設定字色', 'info', 1800);
      return false;
    }

    if (normalizedValue) {
      setSchedule(prev => {
        const next = { ...(prev || {}) };
        const currentRow = prev?.[staffId];
        next[staffId] = currentRow && typeof currentRow === 'object' ? { ...currentRow } : {};
        const existingCell = next[staffId]?.[dateStr];
        const existingMeta = typeof existingCell === 'object' && existingCell !== null ? existingCell : {};
        next[staffId][dateStr] = {
          ...existingMeta,
          value: normalizedValue,
          source: normalizedSource,
          ...(nextColor ? { textColor: nextColor } : {})
        };
        if (!nextColor) delete next[staffId][dateStr].textColor;
        return next;
      });

      setSelectionRangeFromCells([{ staffId, dateStr }], { activeCell: { staffId, dateStr } });
      return true;
    }

    let changed = false;
    setPreScheduleMonthlySchedules((prev) => {
      const next = { ...(prev || {}) };
      const monthKey = String(dateStr).slice(0, 7);
      const [targetYear, targetMonth] = String(monthKey).split('-').map(Number);
      const baseMonthState = next[monthKey] || {
        year: targetYear,
        month: targetMonth,
        staffs: normalizeStaffGroup((staffs || []).map((staff) => ({
          id: staff.id,
          name: staff.name,
          group: staff.group || '白班',
          pregnant: Boolean(staff.pregnant)
        }))),
        scheduleData: {},
        customColumnValues: {},
        schedulingRulesText: '',
        importMeta: {
          sourceType: 'manualPreSchedule',
          sourceFiles: [],
          sourceSheets: [],
          importedAt: new Date().toISOString(),
          lastUpdatedAt: new Date().toISOString(),
          importMode: 'preSchedule'
        }
      };

      const monthState = {
        ...baseMonthState,
        staffs: Array.isArray(baseMonthState.staffs) && baseMonthState.staffs.length > 0
          ? [...baseMonthState.staffs]
          : normalizeStaffGroup((staffs || []).map((staff) => ({
              id: staff.id,
              name: staff.name,
              group: staff.group || '白班',
              pregnant: Boolean(staff.pregnant)
            })))
      };

      const monthSchedule = JSON.parse(JSON.stringify(monthState.scheduleData || monthState.schedule || {}));
      const matchedStaff = resolvePreScheduleMatchedStaff(staffId, monthState)
        || monthState.staffs.find((staff) => staff.id === staffId)
        || (() => {
          const currentStaff = staffs.find((staff) => staff.id === staffId);
          if (!currentStaff) return null;
          const clonedStaff = {
            id: currentStaff.id,
            name: currentStaff.name,
            group: currentStaff.group || '白班',
            pregnant: Boolean(currentStaff.pregnant)
          };
          monthState.staffs = [...(monthState.staffs || []), clonedStaff];
          return clonedStaff;
        })();

      if (!matchedStaff) return prev;

      if (!monthSchedule[matchedStaff.id]) monthSchedule[matchedStaff.id] = {};
      const existingCell = monthSchedule[matchedStaff.id]?.[dateStr];
      const existingMeta = typeof existingCell === 'object' && existingCell !== null ? { ...existingCell } : {};
      monthSchedule[matchedStaff.id][dateStr] = {
        ...existingMeta,
        value: preScheduleCode,
        source: existingMeta.source || 'manual',
        ...(nextColor ? { textColor: nextColor } : {})
      };
      if (!nextColor) delete monthSchedule[matchedStaff.id][dateStr].textColor;
      changed = true;

      next[monthKey] = {
        ...monthState,
        year: targetYear,
        month: targetMonth,
        scheduleData: monthSchedule,
        importMeta: {
          ...(monthState.importMeta || {}),
          sourceType: monthState.importMeta?.sourceType || 'manualPreSchedule',
          importMode: 'preSchedule',
          lastUpdatedAt: new Date().toISOString()
        }
      };
      return next;
    });

    if (changed) {
      setSelectionRangeFromCells([{ staffId, dateStr }], { activeCell: { staffId, dateStr } });
      return true;
    }
    return false;
  };

  function getCellCode(staffId, dateStr) {
    return getContextCellCode(staffId, dateStr);
  }

  function getCellSource(staffId, dateStr) {
    const cellData = getContextCellData(staffId, dateStr);
    if (!cellData) return '';
    if (typeof cellData === 'object' && cellData !== null) return cellData.source || 'manual';
    return 'manual';
  }

  function getCellTextColor(staffId, dateStr) {
    const cellData = getContextCellData(staffId, dateStr);
    if (!cellData || typeof cellData !== 'object') return '';
    return String(cellData.textColor || '').trim();
  }

  const countConsecutiveWorkDaysBefore = (staffId, dateStr) => {
    let count = 0;
    let cursor = addDays(parseDateKey(dateStr), -1);
    while (true) {
      const key = formatDateKey(cursor);
      const code = getCellCode(staffId, key);
      if (!isShiftCode(code)) break;
      count += 1;
      cursor = addDays(cursor, -1);
    }
    return count;
  };

  const canAssign = (staff, dateStr, shiftCode) => {
    const reasons = [];
    const currentCode = getCellCode(staff.id, dateStr);
    if (currentCode) {
      reasons.push('該格已有排班或休假代碼');
    }

    const prefix = getCodePrefix(currentCode);
    if (isConfiguredLeaveCode(currentCode)) {
      reasons.push('該格已有休假，不可再排班');
    }

    const staffGroup = staff.group || '白班';
    const shiftGroup = getShiftGroupByCode(shiftCode);
    if (!SMART_RULES.allowCrossGroupAssignment && shiftGroup && staffGroup !== shiftGroup) {
      reasons.push('不可跨群組排班');
    }

    const prevKey = formatDateKey(addDays(parseDateKey(dateStr), -1));
    const prevCode = getCellCode(staff.id, prevKey);
    const disallowed = SMART_RULES.disallowedNextShiftMap[prevCode] || [];
    if (disallowed.includes(shiftCode)) {
      reasons.push(`${prevCode} 後不可接 ${shiftCode}`);
    }

    const consecutiveBefore = countConsecutiveWorkDaysBefore(staff.id, dateStr);
    if (consecutiveBefore + 1 > SMART_RULES.maxConsecutiveWorkDays) {
      reasons.push(`連續上班不可超過 ${SMART_RULES.maxConsecutiveWorkDays} 天`);
    }

    if (staff.pregnant && SMART_RULES.pregnancyRestrictedShifts.includes(shiftCode)) {
      reasons.push('懷孕標記人員不可排 N / 夜8-8');
    }

    return { allowed: reasons.length === 0, reasons };
  };

  const getShiftCountForStaff = (staffId, shiftCode) => {
    return daysInMonth.reduce((sum, d) => sum + (getCellCode(staffId, d.date) === shiftCode ? 1 : 0), 0);
  };

  const scoreCandidate = (staff, dateStr, shiftCode) => {
    let score = 0;
    const stats = getStaffStats(staff.id);
    score += (999 - getShiftCountForStaff(staff.id, shiftCode)) * SMART_RULES.fillPriorityWeights.sameShiftCount;
    score += (999 - stats.work) * SMART_RULES.fillPriorityWeights.totalShiftCount;
    if (getShiftGroupByCode(shiftCode) === (staff.group || '白班')) {
      score += 100 * SMART_RULES.fillPriorityWeights.sameGroup;
    }
    return score;
  };

  const getCurrentConsecutiveLeavePattern = (staffId, dateStr) => {
    const prevCode = getCellCode(staffId, formatDateKey(addDays(parseDateKey(dateStr), -1)));
    const nextCode = getCellCode(staffId, formatDateKey(addDays(parseDateKey(dateStr), 1)));
    const prevIsLeave = isLeaveCode(prevCode);
    const nextIsLeave = isLeaveCode(nextCode);
    return {
      prevIsLeave,
      nextIsLeave,
      adjacentLeaveCount: (prevIsLeave ? 1 : 0) + (nextIsLeave ? 1 : 0)
    };
  };

  const openFillModal = (staff, dateStr) => {
    const group = staff.group || '白班';
    const shiftCode = DEFAULT_SHIFT_BY_GROUP[group];
    const dayInfo = daysInMonth.find(d => d.date === dateStr);
    const demand = dayInfo ? Number(staffingConfig?.requiredStaffing?.[getRequiredStaffingBucketByDay(dayInfo)]?.[GROUP_TO_DEMAND_KEY[group]] || 0) : 0;
    const alreadyAssigned = staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => {
      const code = getCellCode(s.id, dateStr);
      return sum + (getShiftGroupByCode(code) === group ? 1 : 0);
    }, 0);
    const leaveCount = daysInMonth.reduce((sum, d) => sum + (isLeaveCode(getCellCode(staff.id, d.date)) ? 1 : 0), 0);
    const leaveDeficit = Math.max(0, requiredLeaves - leaveCount);

    const shiftResult = canAssign(staff, dateStr, shiftCode);
    const currentLeaves = leaveCount;
    const remainingBlanks = daysInMonth.reduce((sum, d) => sum + (!getCellCode(staff.id, d.date) ? 1 : 0), 0);
    const remainingLeavesNeeded = Math.max(0, requiredLeaves - currentLeaves);
    const canKeepLeaveTarget = (remainingBlanks - 1) >= remainingLeavesNeeded;

    const candidates = [];

    if (alreadyAssigned < demand && shiftResult.allowed && canKeepLeaveTarget) {
      const reasonBits = [
        `${group}缺額尚未補滿`,
        `${shiftCode} 為此群組主班別`
      ];
      candidates.push({
        type: 'self-shift',
        staffId: staff.id,
        staffName: staff.name,
        group,
        shiftCode,
        allowed: true,
        score: scoreCandidate(staff, dateStr, shiftCode),
        reasons: reasonBits
      });
    }

    if (alreadyAssigned >= demand && leaveDeficit > 0) {
      const sameDayLeaveLoad = staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => sum + (isLeaveCode(getCellCode(s.id, dateStr)) ? 1 : 0), 0);
      const recentWorkPressure = (() => {
        let count = 0;
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 0; i < 3; i += 1) {
          if (isShiftCode(getCellCode(staff.id, formatDateKey(cursor)))) count += 1;
          cursor = addDays(cursor, -1);
        }
        return count;
      })();
      const consecutiveLeavePattern = getCurrentConsecutiveLeavePattern(staff.id, dateStr);
      let offScore = leaveDeficit * 100 + recentWorkPressure * 20 - sameDayLeaveLoad * 25;
      if (consecutiveLeavePattern.adjacentLeaveCount === 1) offScore += 18;
      if (consecutiveLeavePattern.adjacentLeaveCount >= 2) offScore -= 10;
      candidates.push({
        type: 'self-leave',
        staffId: staff.id,
        staffName: staff.name,
        group,
        shiftCode: defaultAutoLeaveCode,
        allowed: true,
        score: offScore,
        reasons: ['本月休假尚未達標', `當日群組需求已滿，優先補休（${defaultAutoLeaveCode}）`]
      });
    }

    candidates.sort((a, b) => b.score - a.score);

    setSelectedFillCell({ staffId: staff.id, staffName: staff.name, dateStr, group: staff.group });
    setFillCandidates(candidates);
    setShowFillModal(true);
  };

const openSelectedCellFillModal = () => {
    if (!selectedGridCell) return;
    openFillModal(selectedGridCell.staff, selectedGridCell.dateStr);
  };

  const clearSelectedCell = () => {
    if (!selectedGridCell) return;
    const { staff, dateStr } = selectedGridCell;
    const currentCode = getCellCode(staff.id, dateStr);
    if (!currentCode) return;
    if (!window.confirm(`確定清除此格內容？\n${staff.name}｜${dateStr}｜${currentCode}`)) return;

    handleCellChange(staff.id, dateStr, '', { preserveSelection: true, selectionCells: [{ staffId: staff.id, dateStr }], activeCell: { staffId: staff.id, dateStr } });
    setSelectedGridCell(null);

  };

  const clearSelectedPreScheduleCell = () => {
    if (!selectedGridCell) return;
    const { staff, dateStr } = selectedGridCell;
    const preScheduleCode = getVisiblePreScheduleCode(staff.id, dateStr);
    if (!preScheduleCode) return;
    if (!window.confirm(`確定清除此格預班？\n${staff.name}｜${dateStr}｜${preScheduleCode}`)) return;

    const changedCount = updatePreScheduleEntries([{ staffId: staff.id, dateStr, value: '' }]);
    if (changedCount > 0) {

    }
  };

  const clearRangeCells = () => {
    if (ruleFillConfig.selectedStaffs.length === 0) {
      setRuleFillFeedback('⚠️ 請先選擇要清除的人員');
      return;
    }

    const startDay = Number(ruleFillConfig.dateRange.start || 1);
    const endDay = Number(ruleFillConfig.dateRange.end || 31);
    const targetStaffIds = new Set(ruleFillConfig.selectedStaffs);

    if (rangeClearMode === 'preScheduleOnly') {
      const clearEntries = [];
      staffs.forEach((staff) => {
        if (!targetStaffIds.has(staff.id)) return;
        daysInMonth.forEach((day) => {
          if (day.day < startDay || day.day > endDay) return;
          const preScheduleCode = getVisiblePreScheduleCode(staff.id, day.date);
          if (!preScheduleCode) return;
          clearEntries.push({ staffId: staff.id, dateStr: day.date, value: '' });
        });
      });

      if (clearEntries.length === 0) {
        setRuleFillFeedback('ℹ️ 指定範圍內沒有可清除的預班');
        return;
      }

      const changedCount = updatePreScheduleEntries(clearEntries);
      if (changedCount > 0) {

      }
      return;
    }

    const clearEntries = [];

    staffs.forEach(staff => {
      if (!targetStaffIds.has(staff.id)) return;
      daysInMonth.forEach(day => {
        if (day.day < startDay || day.day > endDay) return;
        const cellData = schedule[staff.id]?.[day.date];
        if (!cellData) return;

        const source = typeof cellData === 'object' && cellData !== null ? (cellData.source || 'manual') : 'manual';
        if (rangeClearMode === 'autoOnly' && source !== 'auto') return;

        clearEntries.push({ staffId: staff.id, dateStr: day.date, value: '', source: 'auto' });
      });
    });

    const changedClearEntries = filterChangedScheduleEntries(clearEntries);
    if (changedClearEntries.length === 0) {
      setRuleFillFeedback('ℹ️ 指定範圍內沒有可清除的內容');
      return;
    }

    applyRuleFillEntries(changedClearEntries, {
      preserveSelection: true,
      selectionCells: changedClearEntries.map(({ staffId, dateStr }) => ({ staffId, dateStr })),
      activeCell: changedClearEntries.length > 0 ? { staffId: changedClearEntries[changedClearEntries.length - 1].staffId, dateStr: changedClearEntries[changedClearEntries.length - 1].dateStr } : null,
      clearAssist: false,
      resetBuffer: true
    });


  };

  const applyFillCandidate = (candidate) => {
    if (!selectedFillCell) return;
    handleCellChange(candidate.staffId, selectedFillCell.dateStr, candidate.shiftCode, { source: 'auto', preserveSelection: true, selectionCells: [{ staffId: candidate.staffId, dateStr: selectedFillCell.dateStr }], activeCell: { staffId: candidate.staffId, dateStr: selectedFillCell.dateStr } });
    setShowFillModal(false);
    setSelectedFillCell(null);
    setFillCandidates([]);
    setSelectedGridCell(null);
  };

  const addStaff = (group = '白班') => {
    const newId = 's' + Date.now();
    setStaffs(prev => [...prev, { id: newId, name: '新成員', group }]);
    setSchedule(prev => ({ ...prev, [newId]: {} }));
    setSelectedGridCell(null);
    setRangeSelection(null);
    setSelectionAnchor(null);
    setIsRangeDragging(false);
    resetKeyInputBuffer();
  };

  const removeStaff = (staffId) => {
    if (!window.confirm("確定要刪除此人員嗎？")) return;
    setStaffs(prev => prev.filter(s => s.id !== staffId));
    setSchedule(prev => {
      const next = { ...prev };
      delete next[staffId];
      return next;
    });
    setSelectedGridCell(null);
    setRangeSelection(null);
    setSelectionAnchor(null);
    setIsRangeDragging(false);
    resetKeyInputBuffer();
  };


  const moveStaffWithinGroupByDrag = (draggedStaffId, targetStaffId, position = 'before') => {
    if (!draggedStaffId || !targetStaffId || draggedStaffId === targetStaffId) return;

    const draggedStaff = staffs.find(staff => staff.id === draggedStaffId);
    const targetStaff = staffs.find(staff => staff.id === targetStaffId);
    if (!draggedStaff || !targetStaff) return;
    if ((draggedStaff.group || '白班') !== (targetStaff.group || '白班')) return;

    const nextStaffs = [...staffs];
    const draggedIndex = nextStaffs.findIndex(staff => staff.id === draggedStaffId);
    const targetIndex = nextStaffs.findIndex(staff => staff.id === targetStaffId);
    if (draggedIndex === -1 || targetIndex === -1) return;

    const [draggedItem] = nextStaffs.splice(draggedIndex, 1);
    const targetIndexAfterRemoval = nextStaffs.findIndex(staff => staff.id === targetStaffId);
    let insertIndex = position === 'after' ? targetIndexAfterRemoval + 1 : targetIndexAfterRemoval;
    if (insertIndex < 0) insertIndex = nextStaffs.length;
    nextStaffs.splice(insertIndex, 0, draggedItem);

    setStaffs(nextStaffs);
    setSelectedGridCell(null);
    setRangeSelection(null);
    setSelectionAnchor(null);
    setIsRangeDragging(false);
    resetKeyInputBuffer();
  };

  const warnGroupMismatchIfNeeded = (staff, nextGroup) => {
    if (!staff || !nextGroup) return;
    const mismatchCodes = daysInMonth
      .map(day => getCellCode(staff.id, day.date))
      .filter(code => code && isShiftCode(code) && getShiftGroupByCode(code) && getShiftGroupByCode(code) !== nextGroup);
    if (mismatchCodes.length > 0) {
      const uniqueCodes = Array.from(new Set(mismatchCodes));
      setRuleFillFeedback(`⚠️ 已將 ${staff.name} 移到${nextGroup}，但此人仍包含不屬於新群組的班別：${uniqueCodes.join('、')}`);
    }
  };

  const moveStaffAcrossGroupsByDrag = (draggedStaffId, targetGroup, targetStaffId = '', position = 'before') => {
    if (!draggedStaffId || !targetGroup) return;
    const draggedStaff = staffs.find(staff => staff.id === draggedStaffId);
    if (!draggedStaff) return;

    const nextStaffs = [...staffs];
    const draggedIndex = nextStaffs.findIndex(staff => staff.id === draggedStaffId);
    if (draggedIndex === -1) return;

    const [draggedItem] = nextStaffs.splice(draggedIndex, 1);
    const updatedDraggedItem = { ...draggedItem, group: targetGroup };

    let insertIndex = nextStaffs.length;
    if (targetStaffId) {
      const targetIndex = nextStaffs.findIndex(staff => staff.id === targetStaffId);
      if (targetIndex !== -1) {
        insertIndex = position === 'after' ? targetIndex + 1 : targetIndex;
      }
    } else {
      const groupIndexes = nextStaffs
        .map((staff, index) => ({ staff, index }))
        .filter(item => (item.staff.group || '白班') === targetGroup)
        .map(item => item.index);
      insertIndex = groupIndexes.length > 0 ? groupIndexes[groupIndexes.length - 1] + 1 : nextStaffs.length;
    }

    nextStaffs.splice(insertIndex, 0, updatedDraggedItem);
    setStaffs(nextStaffs);
    setSelectedGridCell(null);
    setRangeSelection(null);
    setSelectionAnchor(null);
    setIsRangeDragging(false);
    resetKeyInputBuffer();
    warnGroupMismatchIfNeeded(updatedDraggedItem, targetGroup);
  };

  const handleStaffDragStart = (event, staff) => {
    if (!staff?.id) return;
    event.stopPropagation();
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', staff.id);
    setDraggingStaffId(staff.id);
    setDragOverTarget(null);
    setDragOverGroup(staff.group || '白班');
  };

  const handleStaffDragOver = (event, targetStaff) => {
    if (!draggingStaffId || !targetStaff?.id || draggingStaffId === targetStaff.id) return;
    const draggingStaff = staffs.find(staff => staff.id === draggingStaffId);
    if (!draggingStaff) return;

    event.preventDefault();
    const bounds = event.currentTarget.getBoundingClientRect();
    const nextPosition = (event.clientY - bounds.top) < (bounds.height / 2) ? 'before' : 'after';
    setDragOverTarget({ staffId: targetStaff.id, position: nextPosition, group: targetStaff.group || '白班' });
    setDragOverGroup(targetStaff.group || '白班');
  };

  const handleTableContainerDragOver = (event) => {
    if (!draggingStaffId) return;
    const container = tableScrollContainerRef.current;
    if (!container) return;

    const rect = container.getBoundingClientRect();
    const verticalThreshold = 72;
    const horizontalThreshold = 56;
    const maxStep = 26;

    let nextScrollTop = container.scrollTop;
    let nextScrollLeft = container.scrollLeft;

    if (event.clientY < rect.top + verticalThreshold) {
      const ratio = (rect.top + verticalThreshold - event.clientY) / verticalThreshold;
      nextScrollTop -= Math.max(8, ratio * maxStep);
    } else if (event.clientY > rect.bottom - verticalThreshold) {
      const ratio = (event.clientY - (rect.bottom - verticalThreshold)) / verticalThreshold;
      nextScrollTop += Math.max(8, ratio * maxStep);
    }

    if (event.clientX < rect.left + horizontalThreshold) {
      const ratio = (rect.left + horizontalThreshold - event.clientX) / horizontalThreshold;
      nextScrollLeft -= Math.max(6, ratio * 20);
    } else if (event.clientX > rect.right - horizontalThreshold) {
      const ratio = (event.clientX - (rect.right - horizontalThreshold)) / horizontalThreshold;
      nextScrollLeft += Math.max(6, ratio * 20);
    }

    if (nextScrollTop !== container.scrollTop) container.scrollTop = nextScrollTop;
    if (nextScrollLeft !== container.scrollLeft) container.scrollLeft = nextScrollLeft;
  };

  const handleStaffDrop = (event, targetStaff) => {
    event.preventDefault();
    event.stopPropagation();
    if (!draggingStaffId || !targetStaff?.id) return;
    const draggingStaff = staffs.find(staff => staff.id === draggingStaffId);
    if (!draggingStaff) return;

    const targetGroup = targetStaff.group || '白班';
    const dropPosition = dragOverTarget?.staffId === targetStaff.id ? dragOverTarget.position : 'before';

    if ((draggingStaff.group || '白班') === targetGroup) {
      moveStaffWithinGroupByDrag(draggingStaffId, targetStaff.id, dropPosition);
    } else {
      moveStaffAcrossGroupsByDrag(draggingStaffId, targetGroup, targetStaff.id, dropPosition);
    }
    setDraggingStaffId(null);
    setDragOverTarget(null);
    setDragOverGroup('');
  };

  const handleGroupDragOver = (event, group) => {
    if (!draggingStaffId) return;
    event.preventDefault();
    setDragOverTarget({ staffId: '', position: 'after', group });
    setDragOverGroup(group);
  };

  const handleGroupDrop = (event, group) => {
    event.preventDefault();
    event.stopPropagation();
    if (!draggingStaffId) return;
    const draggingStaff = staffs.find(staff => staff.id === draggingStaffId);
    if (!draggingStaff) return;
    moveStaffAcrossGroupsByDrag(draggingStaffId, group);
    setDraggingStaffId(null);
    setDragOverTarget(null);
    setDragOverGroup('');
  };

  const handleStaffDragEnd = () => {
    setDraggingStaffId(null);
    setDragOverTarget(null);
    setDragOverGroup('');
  };

  const commitEditingStaffName = (staffId, nextName) => {
    const trimmedName = String(nextName ?? '').trim();
    setStaffs(prev => {
      const next = [...prev];
      const currentIndex = next.findIndex(s => s.id === staffId);
      if (currentIndex !== -1) next[currentIndex].name = trimmedName || '新成員';
      return next;
    });
    setEditingStaffId(null);
    setEditingNameDraft('');
  };

  const groupedStaffs = useMemo(() => {
    return SHIFT_GROUPS.map(group => ({
      group,
      staffs: staffs.filter(staff => (staff.group || '白班') === group)
    }));
  }, [staffs]);


  const emptyStaffStats = useMemo(
    () => buildEmptyStaffStatsService(mergedLeaveCodes),
    [mergedLeaveCodes]
  );

  const staffStatsMap = useMemo(() => {
    return buildStaffStatsMapService({
      staffs,
      schedule,
      daysInMonth,
      mergedLeaveCodes,
      getAllShiftCodes,
      isConfiguredLeaveCode,
      getCodePrefix
    });
  }, [staffs, schedule, daysInMonth, mergedLeaveCodes]);

  const dailyStatsMap = useMemo(() => {
    return buildDailyStatsMapService({
      staffs,
      schedule,
      daysInMonth,
      getShiftGroupByCode,
      isConfiguredLeaveCode
    });
  }, [staffs, schedule, daysInMonth]);

  const requiredCountMap = useMemo(() => {
    return buildRequiredCountMapService({
      daysInMonth,
      staffingConfig,
      getRequiredStaffingBucketByDay
    });
  }, [daysInMonth, staffingConfig]);

  useEffect(() => {
    setSelectedGridCell(null);
    setRangeSelection(null);
    setSelectionAnchor(null);
    setIsRangeDragging(false);
    setEditingStaffId(null);
    setEditingNameDraft('');
    setDraggingStaffId(null);
    setDragOverTarget(null);
    setDragOverGroup('');
    resetKeyInputBuffer();
  }, [staffs.length]);

  return (
    <div className="min-h-screen text-slate-900 p-3 font-sans overflow-x-hidden relative" style={{ backgroundColor: pageBackgroundColor }}>
      <style>{`
        @keyframes pulse-once { 0% { transform: translateY(-10px); opacity: 0; } 100% { transform: translateY(0); opacity: 1; } }
        @keyframes fade-in-down { 0% { opacity: 0; transform: translateY(-5px); } 100% { opacity: 1; transform: translateY(0); } }
        .animate-pulse-once { animation: pulse-once 0.5s ease-out forwards; }
        .animate-fade-in-down { animation: fade-in-down 0.3s ease-out forwards; }
      `}</style>

      <div className="max-w-[98vw] mx-auto mb-4">
        <div className="bg-white rounded-2xl p-4 shadow-sm border border-slate-200 flex flex-col xl:flex-row xl:items-center justify-between gap-3">
          <div>
            <h1 className="text-xl font-black text-slate-800 flex items-center gap-2">
              智慧補班系統｜開發版
              <span className="text-blue-500 text-xs font-normal px-2 py-0.5 bg-blue-50 rounded-lg border border-blue-100">PRO v1.6.0</span>
            </h1>
            <p className="text-slate-500 text-xs mt-1 italic">開發測試使用</p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <input
              ref={draftImportInputRef}
              type="file"
              accept=".json,application/json"
              className="hidden"
              onChange={onImportDraftFileChange}
            />
            <button onClick={() => saveToHistory('手動暫存')} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-1.5 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm">
              <Save size={16} /> 暫存
            </button>
            <button onClick={() => setShowHistoryModal(true)} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-1.5 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm">
              <Clock size={16} /> 本機暫存紀錄
            </button>
            <button onClick={onDownloadDraftFile} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-1.5 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm">
              <Download size={16} /> 下載工作檔
            </button>
            <button onClick={onImportDraftFileClick} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-1.5 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm">
              <Database size={16} /> 匯入工作檔
            </button>

            <div className="relative">
              <button onClick={() => setShowExportMenu(!showExportMenu)} className="flex items-center gap-1.5 bg-slate-800 text-white px-3 py-1.5 rounded-xl font-bold hover:bg-slate-900 transition-all text-sm">
                <Download size={16} /> 匯出
              </button>
              {showExportMenu && (
                <div className="absolute right-0 mt-2 w-48 bg-white border border-slate-200 rounded-xl shadow-xl z-50 overflow-hidden animate-fade-in-down">
                  <button onClick={exportToExcel} className="w-full text-left px-4 py-3 text-sm font-bold text-slate-700 hover:bg-green-50 hover:text-green-700 flex items-center gap-2 transition-colors border-b">
                    <FileSpreadsheet size={16} /> 高品質 Excel
                  </button>
                  <button onClick={exportToWord} className="w-full text-left px-4 py-3 text-sm font-bold text-slate-700 hover:bg-blue-50 hover:text-blue-700 flex items-center gap-2 transition-colors">
                    <FileText size={16} /> 橫向 Word (列印)
                  </button>
                </div>
              )}
            </div>

            <div className="w-px h-8 bg-slate-200 mx-2 hidden sm:block"></div>

            <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200 flex-wrap">
              <button onClick={() => handleRuleBasedAutoSchedule(false)} disabled={isRuleFillLoading} className="flex items-center gap-2 bg-white text-blue-600 px-3 py-1.5 rounded-lg font-bold hover:bg-blue-50 transition-all disabled:opacity-50 text-xs">
                {isRuleFillLoading ? <Loader2 className="animate-spin" size={14} /> : <Sparkles size={14} />} 規則全月補空
              </button>
              <button onClick={() => setShowRuleFillControl(!showRuleFillControl)} className={`flex items-center gap-2 px-3 py-1.5 rounded-lg font-bold transition-all text-xs ${showRuleFillControl ? 'bg-blue-600 text-white shadow-inner' : 'text-slate-600 hover:bg-slate-200'}`}>
                <Calendar size={14} /> 規則指定補空
              </button>
              <button
                type="button"
                onClick={() => setPreScheduleEditMode(prev => !prev)}
                className={`flex items-center gap-2 px-3 py-1.5 rounded-lg font-bold transition-all text-xs ${preScheduleEditMode ? 'bg-cyan-600 text-white shadow-inner' : 'text-cyan-700 hover:bg-cyan-50'}`}
                title="開啟後，點格或鍵盤輸入只會寫入預班，不會改正式班表"
              >
                <CalendarDays size={14} /> 預班模式
              </button>
              <button
                type="button"
                onClick={openSelectedCellFillModal}
                disabled={!selectedGridCell}
                className={`flex items-center gap-2 px-3 py-1.5 rounded-lg font-bold transition-all text-xs ${selectedGridCell ? 'text-slate-700 hover:bg-slate-200' : 'text-slate-400 cursor-not-allowed'}`}
              >
                <Check size={14} /> 補此格
              </button>
              <button
                type="button"
                onClick={clearSelectedPreScheduleCell}
                disabled={!selectedCellHasPreSchedule}
                className={`flex items-center gap-2 px-3 py-1.5 rounded-lg font-bold transition-all text-xs ${selectedCellHasPreSchedule ? 'text-amber-700 hover:bg-amber-50' : 'text-slate-400 cursor-not-allowed'}`}
              >
                <Trash2 size={14} /> 清預班
              </button>
              <label className={`flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-bold border ${selectedCellHasFormalValue ? 'border-slate-200 text-slate-700 bg-white' : 'border-slate-200 text-slate-400 bg-slate-50 cursor-not-allowed'}`} title={selectedCellHasFormalValue ? '修改目前選取格子的字體顏色' : '請先選取已有班別或假別的正式班表格子'}>
                <span>字色</span>
                <input
                  type="color"
                  value={selectedCellTextColor || tableFontColor}
                  disabled={!selectedCellSupportsTextColor}
                  onChange={(e) => {
                    if (!selectedGridCell?.staff?.id || !selectedGridCell?.dateStr) return;
                    applyCellTextColor(selectedGridCell.staff.id, selectedGridCell.dateStr, e.target.value);
                  }}
                  className="w-8 h-6 rounded border border-slate-200 bg-transparent cursor-pointer disabled:cursor-not-allowed"
                />
                <button
                  type="button"
                  disabled={!selectedCellSupportsTextColor || !selectedCellTextColor}
                  onClick={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    if (!selectedGridCell?.staff?.id || !selectedGridCell?.dateStr) return;
                    applyCellTextColor(selectedGridCell.staff.id, selectedGridCell.dateStr, '');
                  }}
                  className={`px-2 py-0.5 rounded-md border ${selectedCellSupportsTextColor && selectedCellTextColor ? 'border-slate-200 text-slate-600 hover:bg-slate-50' : 'border-slate-200 text-slate-300 cursor-not-allowed'}`}
                >
                  清除
                </button>
              </label>
              <button
                type="button"
                onClick={clearSelectionContents}
                disabled={selectedRangeCells.length === 0}
                className={`flex items-center gap-2 px-3 py-1.5 rounded-lg font-bold transition-all text-xs ${selectedRangeCells.length > 0 ? 'text-red-600 hover:bg-red-50' : 'text-slate-400 cursor-not-allowed'}`}
                title={selectedRangeCells.length > 1 ? '清除目前選取範圍（等同 Del）' : '清除目前選取格（等同 Del）'}
              >
                <Trash2 size={14} /> Del清除
              </button>
              <button
                type="button"
                onClick={clearSelectedCell}
                disabled={!selectedGridCell}
                className={`flex items-center gap-2 px-3 py-1.5 rounded-lg font-bold transition-all text-xs ${selectedGridCell ? 'text-red-600 hover:bg-red-50' : 'text-slate-400 cursor-not-allowed'}`}
              >
                <Trash2 size={14} /> 清除此格
              </button>
            </div>
          </div>
        </div>
      </div>

      {showRuleFillControl && (
        <div className="max-w-[98vw] mx-auto mb-4 rounded-3xl border border-slate-200 bg-slate-100/90 px-4 py-4 shadow-sm animate-fade-in-down lg:px-5">
          <div className="mb-3 flex items-center gap-2 text-slate-800">
            <Sparkles size={18} className="text-blue-600" />
            <h3 className="font-black">規則指定補空設定</h3>
          </div>

          <div className="grid gap-5 xl:grid-cols-[1.35fr_0.95fr_1.05fr_1.15fr]">
            <div className="min-w-0">
              <label className="mb-2 block text-xs font-bold text-blue-700">1. 選擇人員（補空範圍）</label>
              <div className="max-h-[296px] space-y-2 overflow-y-auto pr-1">
                {groupedStaffs.map(({ group, staffs: groupStaffs }) => {
                  const groupIds = groupStaffs.map(s => s.id);
                  const isGroupFullySelected = groupIds.length > 0 && groupIds.every(id => ruleFillConfig.selectedStaffs.includes(id));

                  return (
                    <div key={group} className="rounded-xl border border-blue-100 bg-white/75 px-3 py-2 shadow-sm">
                      <div className="mb-1.5 flex items-center gap-2">
                        <span className="text-sm font-black text-slate-700">{group}：</span>
                        <button
                          type="button"
                          onClick={() => {
                            const next = isGroupFullySelected
                              ? ruleFillConfig.selectedStaffs.filter(id => !groupIds.includes(id))
                              : Array.from(new Set([...ruleFillConfig.selectedStaffs, ...groupIds]));
                            setRuleFillConfig({ ...ruleFillConfig, selectedStaffs: next });
                          }}
                          className={`rounded-lg border px-2.5 py-1 text-[11px] font-black transition-all ${isGroupFullySelected ? 'border-blue-600 bg-blue-600 text-white shadow-sm' : 'border-blue-200 bg-blue-50 text-blue-700 hover:bg-blue-100'}`}
                        >
                          全選
                        </button>
                      </div>

                      <div className="grid grid-cols-3 gap-1.5">
                        {groupStaffs.map(s => (
                          <button
                            key={s.id}
                            type="button"
                            onClick={() => {
                              const next = ruleFillConfig.selectedStaffs.includes(s.id)
                                ? ruleFillConfig.selectedStaffs.filter(id => id !== s.id)
                                : [...ruleFillConfig.selectedStaffs, s.id];
                              setRuleFillConfig({ ...ruleFillConfig, selectedStaffs: next });
                            }}
                            className={`min-h-[38px] rounded-lg border px-2.5 py-1.5 text-left text-xs font-bold leading-tight transition-all ${ruleFillConfig.selectedStaffs.includes(s.id) ? 'bg-blue-600 border-blue-600 text-white shadow-sm' : 'bg-white border-blue-200 text-blue-600 hover:bg-blue-50'}`}
                          >
                            {s.name}
                          </button>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            <div>
              <label className="mb-2 block text-xs font-bold text-blue-700">2. 日期範圍（1 ~ 31 號）</label>
              <div className="flex items-center gap-3">
                <input
                  type="number"
                  min="1"
                  max="31"
                  value={ruleFillConfig.dateRange.start}
                  onChange={(e) => setRuleFillConfig({ ...ruleFillConfig, dateRange: { ...ruleFillConfig.dateRange, start: parseInt(e.target.value, 10) || 1 } })}
                  className="w-full rounded-xl border border-blue-200 bg-white px-2 py-2 text-center text-sm font-bold text-slate-800"
                />
                <span className="shrink-0 text-sm font-bold text-slate-500">至</span>
                <input
                  type="number"
                  min="1"
                  max="31"
                  value={ruleFillConfig.dateRange.end}
                  onChange={(e) => setRuleFillConfig({ ...ruleFillConfig, dateRange: { ...ruleFillConfig.dateRange, end: parseInt(e.target.value, 10) || 31 } })}
                  className="w-full rounded-xl border border-blue-200 bg-white px-2 py-2 text-center text-sm font-bold text-slate-800"
                />
              </div>
            </div>

            <div>
              <label className="mb-2 block text-xs font-bold text-blue-700">3. 指定班別（選填）</label>
              <select
                value={ruleFillConfig.targetShift}
                onChange={(e) => setRuleFillConfig({ ...ruleFillConfig, targetShift: e.target.value })}
                className="w-full rounded-xl border border-blue-200 bg-white px-2 py-2 text-sm font-bold text-slate-800"
              >
                <option value="">依群組需求自動補空</option>
                {RULE_FILL_MAIN_SHIFTS.map(s => <option key={s} value={s}>{s} 班</option>)}
              </select>
            </div>

            <div className="space-y-3">
              <div>
                <label className="mb-2 block text-xs font-bold text-blue-700">4. 範圍清除模式</label>
                <select
                  value={rangeClearMode}
                  onChange={(e) => setRangeClearMode(e.target.value)}
                  className="w-full rounded-xl border border-blue-200 bg-white px-2 py-2 text-sm font-bold text-slate-800"
                >
                  <option value="preScheduleOnly">只清除預班</option>
                  <option value="autoOnly">只清除自動補入內容</option>
                  <option value="all">清除範圍內全部正式內容</option>
                </select>
              </div>

              <div className="grid grid-cols-1 gap-3 pt-1">
                <button
                  disabled={isRuleFillLoading || ruleFillConfig.selectedStaffs.length === 0}
                  onClick={() => handleRuleBasedAutoSchedule(true)}
                  className="flex w-full items-center justify-center gap-2 rounded-2xl bg-slate-400 py-2.5 font-black text-white transition-all hover:bg-slate-500 active:scale-95 disabled:opacity-50 disabled:grayscale"
                >
                  {isRuleFillLoading ? <Loader2 className="animate-spin" size={18} /> : <Check size={18} />} 套用規則補空
                </button>
                <button
                  type="button"
                  disabled={ruleFillConfig.selectedStaffs.length === 0}
                  onClick={clearRangeCells}
                  className="flex w-full items-center justify-center gap-2 rounded-2xl border border-slate-300 bg-slate-200/80 py-2.5 font-black text-slate-500 transition-all hover:bg-slate-300 disabled:opacity-50 disabled:grayscale"
                >
                  <Trash2 size={18} /> 範圍清除
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-[98vw] mx-auto mb-4 grid grid-cols-1 lg:grid-cols-12 gap-3">
        <div className="lg:col-span-7 bg-white p-3 rounded-xl shadow-sm border border-slate-200">
          <div className="flex flex-col lg:flex-row lg:items-center gap-4 lg:gap-5">
            <div className="flex flex-wrap items-center gap-x-2 gap-y-3 text-sm font-bold text-slate-700">
              <input
                type="number"
                value={year}
                onChange={(e) => setYear(Number(e.target.value) || new Date().getFullYear())}
                className="w-24 border border-slate-300 rounded-lg px-2 py-1.5 text-center font-bold bg-white text-slate-800"
              />
              <span className="shrink-0">年</span>

              <input
                type="number"
                min="1"
                max="12"
                step="1"
                value={month}
                onChange={(e) => {
                  const nextMonth = Number(e.target.value);
                  if (!Number.isFinite(nextMonth)) return;
                  setMonth(Math.min(12, Math.max(1, nextMonth)));
                }}
                className="w-20 border border-slate-300 rounded-lg px-2 py-1.5 text-center font-bold bg-white text-slate-800"
              />
              <span className="shrink-0">月</span>

              <div className="hidden lg:block w-px h-7 bg-slate-200 mx-2"></div>

              <span className="shrink-0 text-slate-600">應休天數</span>
              <span className="text-base font-black text-slate-800 tabular-nums">{requiredLeaves}</span>
              <span className="shrink-0">天</span>
            </div>
          </div>
        </div>

        <div className="lg:col-span-5 flex items-center justify-end gap-2">
          <button onClick={() => changeScreen('entry')} className="bg-white border px-2 py-2 rounded-xl hover:bg-slate-50 transition-colors text-sm font-bold text-slate-700 flex items-center gap-2">
            <ArrowLeft size={18} className="text-slate-600" /> 回入口頁
          </button>
          <button onClick={() => changeScreen('settings')} className="bg-white border px-2 py-2 rounded-xl hover:bg-slate-50 transition-colors text-sm font-bold text-slate-700 flex items-center gap-2">
            <Settings size={18} className="text-slate-600" /> 系統設定
          </button>
        </div>
      </div>

      <div className="max-w-[98vw] mx-auto rounded-2xl shadow-xl border border-slate-200 bg-white">
        <div ref={tableScrollContainerRef} onDragOver={handleTableContainerDragOver} className="overflow-auto rounded-2xl max-h-[calc(100vh-150px)]">
          <table className="w-max min-w-full border-collapse select-none">
            <thead>
              <tr className="bg-slate-100 border-b-2 border-slate-200 shadow-sm">
                <th className={`sticky left-0 top-0 z-50 border-r font-black shadow-sm ${shiftColumnFontSizeClass}`} style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor, color: shiftColumnFontColor }}>班別</th>
                <th className={`sticky top-0 z-50 border-r font-black shadow-sm ${nameDateColumnFontSizeClass}`} style={{ left: densityConfig.shiftWidth, width: effectiveDensityConfig.nameWidth, minWidth: effectiveDensityConfig.nameWidth, backgroundColor: nameDateColumnBgColor, color: nameDateColumnFontColor }}>姓名/日期</th>
                {daysInMonth.map(d => (
                  <th
                    key={d.day}
                    className={`sticky top-0 z-40 ${densityConfig.dayHeaderClass} border-r text-center shadow-sm`}
                    style={{
                      minWidth: densityConfig.dayMinWidth,
                      backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : '#f1f5f9'),
                      ...getFourWeekDividerStyle(d.date)
                    }}
                  >
                    <div className={`${tableFontSizeClass} opacity-60 uppercase`} style={{ color: tableFontColor }}>{d.weekStr}</div>
                    <div className={`${tableFontSizeClass} font-black`} style={{ color: tableFontColor }}>{d.day}</div>
                  </th>
                ))}
                {showRightStats && (
                <>
                <th className={`sticky top-0 z-40 ${densityConfig.statHeaderClass} border-r min-w-[52px] bg-blue-50 text-blue-700 font-bold shadow-sm`}>上班</th>
                <th className={`sticky top-0 z-40 ${densityConfig.statHeaderClass} border-r min-w-[52px] bg-green-50 text-green-700 font-bold shadow-sm`}>假日休</th>
                <th className={`sticky top-0 z-40 ${densityConfig.statHeaderClass} border-r min-w-[52px] bg-red-50 text-red-700 font-bold shadow-sm`}>總休</th>
                </>
                )}
                {showLeaveStats && mergedLeaveCodes.map(l => (
                  <th key={l} className={`sticky top-0 z-40 ${densityConfig.leaveHeaderClass} border-r min-w-[34px] bg-slate-50 text-[10px] uppercase text-slate-500 font-bold shadow-sm`}>{l}</th>
                ))}
                {(customColumns || []).map(col => (
                  <th key={col} className={`sticky top-0 z-40 ${densityConfig.leaveHeaderClass} border-r min-w-[60px] bg-violet-50 text-[10px] uppercase text-violet-600 font-bold shadow-sm`}>{col}</th>
                ))}
              </tr>
            </thead>

            <tbody>
              {groupedStaffs.map(({ group, staffs: groupStaffList }) => {
                const isCollapsed = Boolean(collapsedGroups[group]);
                const visibleGroupStaffList = isCollapsed ? [] : groupStaffList;
                const isGroupDropActive = Boolean(draggingStaffId) && dragOverGroup === group;
                return (
                <React.Fragment key={group}>
                  {visibleGroupStaffList.map((staff, index) => {
                    const stats = getStaffStats(staff.id);
                    const groupCount = visibleGroupStaffList.length + 1;

                    const isDraggingRow = draggingStaffId === staff.id;
                    const isDragOverBefore = dragOverTarget?.staffId === staff.id && dragOverTarget?.position === 'before';
                    const isDragOverAfter = dragOverTarget?.staffId === staff.id && dragOverTarget?.position === 'after';
                    const rowInsertStyle = isDragOverBefore
                      ? { boxShadow: 'inset 0 3px 0 #3b82f6' }
                      : isDragOverAfter
                        ? { boxShadow: 'inset 0 -3px 0 #10b981' }
                        : null;

                    return (
                      <tr
                        key={staff.id}
                        className={`relative border-b border-slate-100 hover:bg-slate-50/50 transition-colors ${isDraggingRow ? 'opacity-45 bg-blue-50/40' : ''} ${isDragOverBefore ? 'drag-insert-before' : ''} ${isDragOverAfter ? 'drag-insert-after' : ''}`}
                      >
                        {index === 0 && (
                          <td rowSpan={groupCount} className="sticky left-0 z-20 border-r text-center shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)]" style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor, ...(rowInsertStyle || {}) }}>
                            <div className="flex items-center justify-center h-full" style={{ minHeight: densityConfig.rowMinHeight }}>
                              {showShiftLabels && (
                                <span
                                  className={`${shiftColumnFontSizeClass} font-black leading-none tracking-0 [writing-mode:vertical-rl]`}
                                  style={{ color: shiftColumnFontColor, fontSize: shiftCellLabelFontSize }}
                                >
                                  {group}
                                </span>
                              )}
                            </div>
                          </td>
                        )}

                        <td
                          className={`sticky z-30 border-r shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)] px-0.5 py-0.5 relative`}
                          style={{ left: densityConfig.shiftWidth, width: effectiveDensityConfig.nameWidth, minWidth: effectiveDensityConfig.nameWidth, backgroundColor: nameDateColumnBgColor, ...(rowInsertStyle || {}) }}
                          onDragOver={(e) => handleStaffDragOver(e, staff)}
                          onDrop={(e) => handleStaffDrop(e, staff)}
                        >
                          <div className="flex items-center gap-1">
                            <button
                              type="button"
                              draggable
                              onDragStart={(e) => handleStaffDragStart(e, staff)}
                              onDragEnd={handleStaffDragEnd}
                              onMouseDown={(e) => e.stopPropagation()}
                              onClick={(e) => e.stopPropagation()}
                              className="shrink-0 w-4 h-8 flex items-center justify-center rounded-md text-slate-400 hover:text-blue-600 hover:bg-blue-50 cursor-grab active:cursor-grabbing"
                              aria-label={`拖曳排序 ${staff.name}`}
                            >
                              <GripVertical size={14} />
                            </button>
                            {editingStaffId === staff.id ? (
                            <div
                              contentEditable
                              suppressContentEditableWarning
                              ref={(node) => {
                                if (node && editingStaffId === staff.id) {
                                  requestAnimationFrame(() => {
                                    node.focus();
                                    const selection = window.getSelection();
                                    const range = document.createRange();
                                    range.selectNodeContents(node);
                                    range.collapse(false);
                                    selection.removeAllRanges();
                                    selection.addRange(range);
                                  });
                                }
                              }}
                              onInput={(e) => setEditingNameDraft(e.currentTarget.textContent || '')}
                              onBlur={(e) => commitEditingStaffName(staff.id, e.currentTarget.textContent || editingNameDraft)}
                              onMouseDown={(e) => e.stopPropagation()}
                              onClick={(e) => e.stopPropagation()}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') {
                                  e.preventDefault();
                                  commitEditingStaffName(staff.id, e.currentTarget.textContent || editingNameDraft);
                                } else if (e.key === 'Escape') {
                                  e.preventDefault();
                                  setEditingStaffId(null);
                                  setEditingNameDraft('');
                                }
                              }}
                              className={`flex-1 min-w-0 text-center py-0 px-0.5 font-bold bg-transparent whitespace-nowrap outline-none ${nameDateColumnFontSizeClass}`}
                              style={{ color: nameDateColumnFontColor, letterSpacing: "-0.02em", maxWidth: '100%' }}
                            >
                              {editingNameDraft}
                            </div>
                          ) : (
                            <button
                              type="button"
                              onClick={() => {
                                setEditingStaffId(staff.id);
                                setEditingNameDraft(staff.name || '');
                              }}
                              className={`flex-1 min-w-0 bg-transparent border-none text-center font-bold whitespace-nowrap px-0.5 py-0 ${nameDateColumnFontSizeClass}`}
                              style={{ color: nameDateColumnFontColor, letterSpacing: "-0.02em" }}
                              title="點擊編輯姓名"
                            >
                              <span className="block truncate">{staff.name}</span>
                            </button>
                          )}

                            <button
                              onClick={() => removeStaff(staff.id)}
                              className="text-slate-400 hover:text-red-500 shrink-0 w-3 flex items-center justify-center"
                            >
                              <Minus size={14} />
                            </button>
                          </div>
                        </td>

                        {daysInMonth.map(d => {
                          const cellData = schedule[staff.id]?.[d.date];
                          const val = typeof cellData === 'object' && cellData !== null ? (cellData?.value || '') : (cellData || '');
                          const cellTextColor = typeof cellData === 'object' && cellData !== null ? String(cellData?.textColor || '').trim() : '';
                          const cellKey = makeCellKey(staff.id, d.date);
                          const draftValue = cellDrafts[cellKey];
                          const displayValue = draftValue !== undefined ? draftValue : val;
                          const preScheduleCode = getVisiblePreScheduleCode(staff.id, d.date) || '';
                          const hasFormalValue = Boolean(displayValue);
                          const hasSameVisibleCode = hasFormalValue && Boolean(preScheduleCode) && String(displayValue).trim() === String(preScheduleCode).trim();
                          const showPreScheduleAsMain = !hasFormalValue && Boolean(preScheduleCode);
                          const showPreScheduleAsHint = hasFormalValue && Boolean(preScheduleCode) && !hasSameVisibleCode;
                          const baseCellBackgroundColor = d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent');
                          const cellBackgroundColor = showPreScheduleAsMain
                            ? getPreScheduleBackgroundColor(baseCellBackgroundColor, false)
                            : baseCellBackgroundColor;
                          const inRangeSelection = selectedRangeCellKeySet.has(cellKey);
                          const isPrimarySelected = primarySelectedCellKey === cellKey;
                          const isInvalid = Boolean(invalidCellKeys[cellKey]);
                          const ruleWarningMessage = cellRuleWarnings[cellKey] || '';
                          const isRuleWarning = Boolean(ruleWarningMessage) && !isInvalid;

                          return (
                            <ScheduleGridCell
                              key={cellKey}
                              dayInfo={d}
                              densityConfig={densityConfig}
                              tableFontSizeClass={tableFontSizeClass}
                              tableFontColor={tableFontColor}
                              dangerTintColor={dangerTintColor}
                              warningTintColor={warningTintColor}
                              preScheduleHintBorderColor={preScheduleHintBorderColor}
                              preScheduleTextColor={preScheduleTextColor}
                              preScheduleDisplayTextColor={showPreScheduleAsMain ? (getPreScheduleCellTextColor(staff.id, d.date) || '') : ''}
                              displayValue={displayValue}
                              cellTextColor={cellTextColor}
                              showPreScheduleAsMain={showPreScheduleAsMain}
                              showPreScheduleAsHint={showPreScheduleAsHint}
                              preScheduleCode={preScheduleCode}
                              cellBackgroundColor={cellBackgroundColor}
                              inRangeSelection={inRangeSelection}
                              isPrimarySelected={isPrimarySelected}
                              isInvalid={isInvalid}
                              isRuleWarning={isRuleWarning}
                              ruleWarningMessage={ruleWarningMessage}
                              rowInsertStyle={rowInsertStyle}
                              fourWeekDividerStyle={getFourWeekDividerStyle(d.date)}
                              showBlueDots={showBlueDots}
                              selectionMode={selectionMode}
                              showShiftLabels={showShiftLabels}
                              selectorDotClass={densityConfig.selectorDotClass}
                              selectValue={preScheduleEditMode ? preScheduleCode : val}
                              selectTitle={showPreScheduleAsHint ? `選擇班別/假別｜預班 ${preScheduleCode}` : '選擇班別/假別'}
                              shiftOptionsNodes={shiftOptionsNodes}
                              leaveOptionsNodes={leaveOptionsNodes}
                              preScheduleEditMode={preScheduleEditMode}
                              onMouseDownCell={(e) => {
                                if (e.button !== 0) return;
                                e.preventDefault();
                                setIsRangeDragging(true);
                                startRangeSelection(staff, d.date, e);
                              }}
                              onMouseEnterCell={() => updateRangeSelection(staff, d.date)}
                              onClickCell={() => {
                                if (selectionMode === 'cell' || !showBlueDots) {
                                  startRangeSelection(staff, d.date);
                                }
                              }}
                              onDotClick={(e) => {
                                e.stopPropagation();
                                startRangeSelection(staff, d.date);
                              }}
                              onSelectChange={(e) => {
                                startRangeSelection(staff, d.date);
                                const targetCells = [{ staffId: staff.id, dateStr: d.date }];
                                if (preScheduleEditMode) {
                                  applyPreScheduleValueToCells(targetCells, e.target.value, {
                                    advance: Boolean(e.target.value),
                                    direction: 1,
                                    preserveSelection: false,
                                    activeCell: targetCells[targetCells.length - 1],
                                    showFeedback: true
                                  });
                                } else {
                                  applySelectionValue(targetCells, e.target.value, {
                                    advance: Boolean(e.target.value),
                                    direction: 1,
                                    source: 'manual'
                                  });
                                }
                                e.currentTarget.blur();
                              }}
                              onSelectClick={(e) => {
                                e.stopPropagation();
                                startRangeSelection(staff, d.date);
                              }}
                              onSelectMouseDown={(e) => {
                                if (e.button !== 0) return;
                                e.stopPropagation();
                                startRangeSelection(staff, d.date, e);
                              }}
                            />
                          );
                        })}

                        {showRightStats && (
                        <>
                        <td className={`border-r text-center font-black bg-blue-50/30 ${tableFontSizeClass}`} style={{ color: tableFontColor, ...(rowInsertStyle || {}) }}>{stats.work}</td>
                        <td className={`border-r text-center font-black bg-green-50/30 ${tableFontSizeClass}`} style={{ color: tableFontColor, ...(rowInsertStyle || {}) }}>{stats.holidayLeave}</td>
                        <td className={`border-r text-center font-black bg-red-50/30 ${tableFontSizeClass}`} style={{ color: tableFontColor, ...(rowInsertStyle || {}) }}>{stats.totalLeave}</td>
                        </>
                        )}
                        {showLeaveStats && mergedLeaveCodes.map(l => (
                          <td key={l} className={`border-r text-center bg-slate-50/20 ${tableFontSizeClass}`} style={{ color: tableFontColor, ...(rowInsertStyle || {}) }}>
                            {stats.leaveDetails[l] || ''}
                          </td>
                        ))}
                        {(customColumns || []).map(col => (
                          <td key={col} className={`border-r bg-violet-50/10 ${tableFontSizeClass}`} style={{ color: tableFontColor, ...(rowInsertStyle || {}) }}>
                            <input
                              type="text"
                              value={customColumnValues?.[staff.id]?.[col] || ''}
                              onChange={(e) => setCustomColumnValues(prev => ({ ...prev, [staff.id]: { ...(prev?.[staff.id] || {}), [col]: e.target.value } }))}
                              className="w-full h-full px-2 py-1 text-center bg-transparent border-none focus:ring-1 focus:ring-violet-300"
                              placeholder=""
                            />
                          </td>
                        ))}
                      </tr>
                    );
                  })}

                  {!isCollapsed && (
                  <tr className={`border-b border-slate-200 transition-colors ${isGroupDropActive ? 'bg-blue-50/90 shadow-[inset_0_0_0_2px_rgba(59,130,246,0.28)]' : 'bg-slate-50/70'}`} onDragOver={(e) => handleGroupDragOver(e, group)} onDrop={(e) => handleGroupDrop(e, group)}>
                    {visibleGroupStaffList.length === 0 && (
                      <td rowSpan={2} className="sticky left-0 z-20 border-r text-center shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)]" style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor }}>
                        <div className="flex items-center justify-center h-full" style={{ minHeight: densityConfig.rowMinHeight }}>
                          {showShiftLabels && (
                            <span
                              className={`${shiftColumnFontSizeClass} font-black leading-none tracking-0 [writing-mode:vertical-rl]`}
                              style={{ color: shiftColumnFontColor, fontSize: shiftCellLabelFontSize }}
                            >
                              {group}
                            </span>
                          )}
                        </div>
                      </td>
                    )}
                    <td className="sticky z-30 border-r shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)] px-0.5 py-0.5" style={{ left: densityConfig.shiftWidth, width: effectiveDensityConfig.nameWidth, minWidth: effectiveDensityConfig.nameWidth, backgroundColor: nameDateColumnBgColor }}>
                      <div className="flex items-center justify-center">
                        <button
                          onClick={() => addStaff(group)}
                          className="bg-blue-600 text-white px-4 py-2 rounded-xl hover:bg-blue-700 transition-colors flex items-center gap-2 font-bold text-sm"
                        >
                          <Plus size={16} /> 新增
                        </button>
                      </div>
                    </td>

                    {daysInMonth.map(d => (
                      <td
                        key={d.date}
                        className="border-r"
                        style={{
                          backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent'),
                          opacity: d.isHoliday || d.isWeekend ? 0.9 : 1,
                          ...getFourWeekDividerStyle(d.date)
                        }}
                      />
                    ))}

                    <td colSpan={(showRightStats ? 3 : 0) + (showLeaveStats ? mergedLeaveCodes.length : 0) + (customColumns?.length || 0)}></td>
                  </tr>
                  )}

                  <tr className={`border-b border-slate-200 transition-colors ${isGroupDropActive ? 'bg-blue-100/80 shadow-[inset_0_0_0_2px_rgba(59,130,246,0.28)]' : 'bg-amber-50/95'}`} onDragOver={(e) => handleGroupDragOver(e, group)} onDrop={(e) => handleGroupDrop(e, group)}>
                    {visibleGroupStaffList.length > 0 || isCollapsed ? (
                    <td className={`sticky left-0 z-30 border-r ${densityConfig.footCellPaddingClass}`} style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor, top: stickyGroupSummaryTop, boxShadow: stickyGroupSummaryShadow }}>
                      <div className="flex items-center justify-center">
                        <button
                          type="button"
                          onClick={() => setCollapsedGroups(prev => ({ ...prev, [group]: !prev[group] }))}
                          className="inline-flex items-center justify-center w-7 h-7 rounded-md border border-slate-300 bg-white text-slate-600 hover:bg-slate-50 hover:text-slate-800 transition-colors"
                          title={isCollapsed ? `展開${group}` : `收合${group}`}
                          aria-label={isCollapsed ? `展開${group}` : `收合${group}`}
                        >
                          <span className="text-base font-black leading-none">{isCollapsed ? '+' : '−'}</span>
                        </button>
                      </div>
                    </td>
                    ) : null}
                    <td className={`sticky z-30 border-r text-center font-bold ${nameDateColumnFontSizeClass} ${densityConfig.footCellPaddingClass}`} style={{ left: densityConfig.shiftWidth, width: effectiveDensityConfig.nameWidth, minWidth: effectiveDensityConfig.nameWidth, backgroundColor: nameDateColumnBgColor, color: nameDateColumnFontColor, top: stickyGroupSummaryTop, boxShadow: stickyGroupSummaryShadow }}>
                      {group === '白班' ? '白班上班' : group === '小夜' ? '小夜上班' : '大夜上班'}
                    </td>
                    {daysInMonth.map(d => {
                      const count = buildExportDailyStats(d.date)[group === '白班' ? 'D' : group === '小夜' ? 'E' : 'N'];
                      const rowKey = group === '白班' ? 'D' : group === '小夜' ? 'E' : 'N';
                      return (
                        <td
                          key={d.date}
                          className={`sticky z-20 border-r p-2 text-center font-black ${tableFontSizeClass}`}
                          style={{ backgroundColor: groupSummaryRowBgColor, color: tableFontColor, top: stickyGroupSummaryTop, boxShadow: stickyGroupSummaryShadow, ...getDemandHighlightStyle(d.date, rowKey, count), ...getFourWeekDividerStyle(d.date) }}
                        >
                          {count || ''}
                        </td>
                      );
                    })}
                    <td
                      colSpan={(showRightStats ? 3 : 0) + (showLeaveStats ? mergedLeaveCodes.length : 0) + (customColumns?.length || 0)}
                      className="sticky z-20"
                      style={{ backgroundColor: groupSummaryRowBgColor, top: stickyGroupSummaryTop, boxShadow: stickyGroupSummaryShadow }}
                    ></td>
                  </tr>
                </React.Fragment>
              )})}
            </tbody>

            {showBottomStats && (
            <tfoot className="bg-slate-100 border-t-2 border-slate-200">
              <tr>
                <td className={`sticky left-0 z-10 border-r ${densityConfig.footCellPaddingClass}`} style={{ width: densityConfig.shiftWidth, minWidth: densityConfig.shiftWidth, backgroundColor: shiftColumnBgColor }}></td>
                <td className={`sticky z-10 border-r text-right font-bold ${nameDateColumnFontSizeClass} ${densityConfig.footCellPaddingClass}`} style={{ left: densityConfig.shiftWidth, width: effectiveDensityConfig.nameWidth, minWidth: effectiveDensityConfig.nameWidth, backgroundColor: nameDateColumnBgColor, color: nameDateColumnFontColor }}>
                  休假人數
                </td>
                {daysInMonth.map(d => {
                  const count = buildExportDailyStats(d.date).totalLeave;
                  return (
                    <td
                      key={d.date}
                      className={`border-r p-2 text-center font-black ${tableFontSizeClass}`}
                      style={{ color: tableFontColor, ...getFourWeekDividerStyle(d.date) }}
                    >
                      {count || ''}
                    </td>
                  );
                })}
                <td colSpan={3 + mergedLeaveCodes.length + (customColumns?.length || 0)}></td>
              </tr>
            </tfoot>
            )}
          </table>
        </div>
      </div>


      {showImportViolationList && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-[115] flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-3xl overflow-hidden shadow-2xl animate-pulse-once">
            <div className="p-5 border-b flex justify-between items-center bg-slate-50">
              <div>
                <h3 className="font-black text-slate-800">匯入規則衝突清單</h3>
                <p className="text-sm text-slate-500 mt-1">共 {importRuleViolations.length} 筆，僅提示，不影響匯入。</p>
              </div>
              <button onClick={() => { setShowImportViolationList(false); }} className="p-2 hover:bg-slate-200 rounded-full transition-colors">
                <X />
              </button>
            </div>
            <div className="max-h-[65vh] overflow-y-auto">
              <table className="w-full text-sm">
                <thead className="sticky top-0 bg-slate-100 text-slate-700">
                  <tr>
                    <th className="text-left px-4 py-3">人員</th>
                    <th className="text-left px-4 py-3">日期</th>
                    <th className="text-left px-4 py-3">代碼</th>
                    <th className="text-left px-4 py-3">原因</th>
                  </tr>
                </thead>
                <tbody>
                  {importRuleViolations.map((item, index) => (
                    <tr key={`${item.staffName}-${item.dateStr}-${item.code}-${index}`} className="border-t border-slate-200">
                      <td className="px-4 py-3">{item.staffName}</td>
                      <td className="px-4 py-3">{item.dateStr}</td>
                      <td className="px-4 py-3 font-bold">{item.code}</td>
                      <td className="px-4 py-3" style={{ color: warningTextColor }}>{item.reason}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {showFillModal && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-[110] flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-xl overflow-hidden shadow-2xl animate-pulse-once">
            <div className="p-6 border-b flex justify-between items-center bg-slate-50">
              <div>
                <h3 className="font-black text-slate-800">補此格</h3>
                <p className="text-sm text-slate-500 mt-1">{selectedFillCell?.staffName}｜{selectedFillCell?.dateStr}</p>
              </div>
              <button onClick={() => { setShowFillModal(false); setSelectedFillCell(null); setFillCandidates([]); }} className="p-2 hover:bg-slate-200 rounded-full transition-colors">
                <X />
              </button>
            </div>
            <div className="p-4 max-h-[60vh] overflow-y-auto space-y-3">
              {fillCandidates.length === 0 ? (
                <div className="p-6 text-center text-slate-500">目前沒有可直接建議的班別或人員。</div>
              ) : fillCandidates.map((candidate, index) => (
                <div key={`${candidate.staffId}-${candidate.shiftCode}`} className="border rounded-2xl p-4 flex items-center justify-between gap-4 hover:border-blue-400 hover:bg-blue-50/40 transition-all">
                  <div>
                    <div className="font-black text-slate-800">{candidate.staffName} → {candidate.shiftCode}</div>
                    <div className="text-xs text-slate-500 mt-1">群組：{candidate.group}｜排序分數：{candidate.score}</div>
                    {candidate.reasons?.length > 0 && (
                      <div className="text-xs text-slate-500 mt-1">{candidate.reasons.join('｜')}</div>
                    )}
                  </div>
                  <button onClick={() => applyFillCandidate(candidate)} className="px-4 py-2 rounded-xl bg-blue-600 text-white font-bold hover:bg-blue-700">
                    {index === 0 ? '選擇推薦' : '選擇'}
                  </button>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {showHistoryModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-lg overflow-hidden shadow-2xl animate-pulse-once">
            <div className="p-6 border-b flex justify-between items-center bg-slate-50">
              <h3 className="font-black text-slate-800 flex items-center gap-2"><Clock /> 本機暫存紀錄</h3>
              <button onClick={() => setShowHistoryModal(false)} className="p-2 hover:bg-slate-200 rounded-full transition-colors">
                <X />
              </button>
            </div>

            <div className="max-h-[60vh] overflow-y-auto p-4 space-y-3">
              {historyList.length === 0 ? (
                <p className="text-center py-10 text-slate-400 font-bold">目前尚無存檔紀錄</p>
              ) : (
                historyList.map(record => (
                  <div key={record.id} className="flex items-center justify-between p-4 border rounded-2xl hover:border-blue-400 hover:bg-blue-50/30 transition-all group">
                    <div>
                      <div className="font-black text-slate-800">
                        {record.label}
                        <span className="text-xs font-normal text-slate-400 ml-2">{record.state.year}/{record.state.month}</span>
                      </div>
                      <div className="text-xs text-slate-500">{record.timestamp}</div>
                    </div>
                    <button onClick={() => loadHistory(record)} className="bg-slate-100 text-slate-700 px-4 py-2 rounded-xl text-sm font-bold hover:bg-blue-600 hover:text-white transition-all">
                      載入
                    </button>
                  </div>
                ))
              )}
            </div>

            <div className="p-4 bg-slate-50 border-t flex gap-3">
              <button onClick={clearHistory} className="flex-1 py-3 text-sm font-bold text-red-500 hover:bg-red-50 rounded-xl transition-colors">清空所有紀錄</button>
              <button onClick={() => setShowHistoryModal(false)} className="flex-1 py-3 text-sm font-bold text-slate-600 bg-white border rounded-xl hover:bg-slate-100 transition-colors">關閉</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


function SettingRow({ icon: Icon, title, desc, children, iconBg = 'bg-blue-50', iconColor = 'text-blue-600' }) {
  return (
    <div className="bg-white border border-gray-200 rounded-2xl shadow-sm overflow-hidden">
      <div className="grid grid-cols-1 lg:grid-cols-[260px_minmax(0,1fr)]">
        <div className="px-6 py-6 bg-gray-50/80 border-b lg:border-b-0 lg:border-r border-gray-200">
          <div className="flex items-start gap-3">
            <div className={`p-2 rounded-xl ${iconBg}`}><Icon className={`w-4 h-5 ${iconColor}`} /></div>
            <div><h3 className="font-bold text-gray-800">{title}</h3><p className="text-sm text-gray-500 mt-1 leading-relaxed">{desc}</p></div>
          </div>
        </div>
        <div className="px-6 py-6">{children}</div>
      </div>
    </div>
  );
}

function SettingsView({ changeScreen, onSaveSettings, colors, setColors, customHolidays, setCustomHolidays, specialWorkdays, setSpecialWorkdays, medicalCalendarAdjustments, setMedicalCalendarAdjustments, staffingConfig, setStaffingConfig, uiSettings, setUiSettings, customLeaveCodes, setCustomLeaveCodes, customWorkShifts, setCustomWorkShifts, customColumns, setCustomColumns, schedulingRulesText, setSchedulingRulesText }) {
  const activeThemePreset = useMemo(() => {
    const presetKeys = [
      'pageBackgroundColor',
      'tableFontColor',
      'shiftColumnBgColor',
      'nameDateColumnBgColor',
      'shiftColumnFontColor',
      'nameDateColumnFontColor',
      'demandOverColor',
      'groupSummaryRowBgColor',
      'warningTintColor',
      'warningTextColor',
      'infoTintColor',
      'infoTextColor',
      'dangerTintColor',
      'dangerTextColor'
    ];

    for (const [presetKey, preset] of Object.entries(UI_THEME_PRESETS || {})) {
      const matchesCalendarColors = (colors?.weekend || '') === (preset?.weekendColor || '')
        && (colors?.holiday || '') === (preset?.holidayColor || '');
      const matchesUiSettings = presetKeys.every((key) => (uiSettings?.[key] || '') === (preset?.[key] || ''));
      if (matchesCalendarColors && matchesUiSettings) return presetKey;
    }
    return 'custom';
  }, [colors, uiSettings]);

  const [holidayInput, setHolidayInput] = useState({ year: '', month: '', day: '' });
  const systemToday = useMemo(() => new Date(), []);
  const systemYear = String(systemToday.getFullYear());
  const systemMonth = String(systemToday.getMonth() + 1).padStart(2, '0');
  const systemDay = String(systemToday.getDate()).padStart(2, '0');
  const sanitizeHolidayPart = (field, rawValue = '') => {
    const digits = String(rawValue || '').replace(/\D/g, '');
    if (field === 'year') {
      if (!digits) return '';
      if (digits.length > 4) return systemYear;
      const parsedYear = Number(digits);
      if (digits.length === 4 && parsedYear > Number(systemYear)) return systemYear;
      return digits.slice(0, 4);
    }
    if (!digits) return '';
    if (digits.length > 2) return field === 'month' ? systemMonth : systemDay;
    const numeric = Number(digits);
    if (field === 'month' && numeric > 12) return systemMonth;
    if (field === 'day' && numeric > 31) return systemDay;
    return digits.slice(0, 2);
  };
  const updateHolidayInput = (field, rawValue) => {
    setHolidayInput((prev) => ({ ...prev, [field]: sanitizeHolidayPart(field, rawValue) }));
  };
  const getHolidayInputClassName = (field) => {
    const value = String(holidayInput?.[field] || '');
    const hasValue = value.length > 0;
    const isYearInvalid = field === 'year' && hasValue && value.length < 4;
    return `w-full px-2.5 py-1.5 text-sm border rounded-lg bg-gray-50 focus:outline-none focus:ring-2 ${isYearInvalid ? 'border-red-400 ring-1 ring-red-100 focus:ring-red-100' : 'border-gray-200 focus:ring-blue-100'}`;
  };
  const mergedLeaveCodes = useMemo(() => Array.from(new Set([...(DICT.LEAVES || []), ...((customLeaveCodes || []))])), [customLeaveCodes]);
  const addCustomLeaveCode = () => {
    const raw = window.prompt('請輸入自訂休假代碼');
    const code = String(raw || "").trim();
    if (!code) return;
    if (DICT.LEAVES.includes(code) || (customLeaveCodes || []).includes(code)) return;
    setCustomLeaveCodes(prev => [...prev, code]);
  };
  const removeCustomLeaveCode = (code) => setCustomLeaveCodes(prev => prev.filter(item => item !== code));
  const addCustomWorkShift = (group) => {
    const raw = window.prompt(`請輸入${group}自訂上班代碼`);
    const code = String(raw || "").trim();
    if (!code) return;
    if (getAllShiftCodes().includes(code) || (customWorkShifts || []).some(item => String(item?.code || '').trim() === code)) return;
    setCustomWorkShifts(prev => [...(prev || []), { group, code }]);
  };
  const removeCustomWorkShift = (group, code) => setCustomWorkShifts(prev => (prev || []).filter(item => !(item.group === group && item.code === code)));
  const addCustomColumn = () => {
    const raw = window.prompt('請輸入自訂欄位名稱');
    const name = String(raw || "").trim();
    if (!name) return;
    if ((customColumns || []).includes(name)) return;
    setCustomColumns(prev => [...prev, name]);
  };
  const removeCustomColumn = (name) => setCustomColumns(prev => prev.filter(item => item !== name));
  const customWorkShiftGroups = [
    { group: '白班', hint: '例如：0800-1600、0700-1500' },
    { group: '小夜', hint: '例如：1600-2400、0700-1600' },
    { group: '大夜', hint: '例如：2400-0800' }
  ];
  const addCustomHoliday = () => {
    const y = holidayInput.year.trim();
    const m = holidayInput.month.trim();
    const d = holidayInput.day.trim();
    if (y.length !== 4 || !m || !d) return;
    const monthNumber = Number(m);
    const dayNumber = Number(d);
    if (!Number.isInteger(monthNumber) || !Number.isInteger(dayNumber)) return;
    if (monthNumber < 1 || monthNumber > 12 || dayNumber < 1 || dayNumber > 31) return;
    const normalizedMonth = String(monthNumber).padStart(2, '0');
    const normalizedDay = String(dayNumber).padStart(2, '0');
    const dateStr = `${y}-${normalizedMonth}-${normalizedDay}`;
    if (customHolidays.includes(dateStr)) return;
    setCustomHolidays(prev => [...prev, dateStr].sort());
    setHolidayInput({ year: '', month: '', day: '' });
  };
  const removeCustomHoliday = (dateStr) => setCustomHolidays(prev => prev.filter(item => item !== dateStr));
  return (
    <div className="min-h-screen bg-gray-50 text-slate-800 font-sans selection:bg-blue-100">
      <header className="sticky top-0 z-50 bg-white border-b border-gray-200 px-8 py-4 flex justify-between items-center shadow-sm">
        <div><div className="flex items-center gap-2 mb-0.5"><Settings className="w-6 h-6 text-blue-600" /><h1 className="text-xl font-bold tracking-tight text-gray-900">系統設定</h1></div><p className="text-sm text-gray-500">可調整使用者設定與畫面顯示參數。</p></div>
        <div className="flex items-center gap-3">
          <button onClick={() => changeScreen('entry')} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-600 bg-white border border-gray-300 rounded-lg hover:bg-gray-50"><ArrowLeft className="w-4 h-4" />返回入口頁</button>
          <button onClick={() => changeScreen('schedule')} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-600 bg-white border border-gray-300 rounded-lg hover:bg-gray-50"><Calendar className="w-4 h-4" />返回排班頁</button>
          <button onClick={onSaveSettings} className="flex items-center gap-2 px-6 py-2 text-sm font-semibold text-white bg-blue-600 rounded-lg hover:bg-blue-700 shadow-md shadow-blue-100"><Save className="w-4 h-4" />儲存設定</button>
        </div>
      </header>
      <main className="max-w-7xl mx-auto p-8 space-y-10">
        <section className="space-y-6">
          <div className="flex items-center gap-2"><div className="w-1 h-6 bg-blue-600 rounded-full"></div><h2 className="text-lg font-bold text-gray-800">使用者偏好設定</h2></div>
          <div className="space-y-5">
            <SettingRow icon={Monitor} title="外觀與顯示" desc="調整班表顏色、顯示大小與統計欄位呈現。">
              <div className="space-y-4">
                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">色彩標示</label>
                  <div className="grid grid-cols-1 md:grid-cols-4 gap-2">
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">週末顏色</span><input type="color" value={colors.weekend} onChange={(e) => { const nextColor = e.target.value; setColors(prev => ({ ...prev, weekend: nextColor })); setUiSettings(prev => ({ ...prev, themePreset: 'custom' })); }} className="w-9 h-7 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">國定假日</span><input type="color" value={colors.holiday} onChange={(e) => { const nextColor = e.target.value; setColors(prev => ({ ...prev, holiday: nextColor })); setUiSettings(prev => ({ ...prev, themePreset: 'custom' })); }} className="w-9 h-7 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">主頁背景色</span><input type="color" value={uiSettings.pageBackgroundColor} onChange={(e) => setUiSettings(prev => ({ ...prev, pageBackgroundColor: e.target.value, themePreset: 'custom' }))} className="w-9 h-7 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">表格字體色(全體)</span><input type="color" value={uiSettings.tableFontColor} onChange={(e) => setUiSettings(prev => ({ ...prev, tableFontColor: e.target.value, themePreset: 'custom' }))} className="w-9 h-7 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">班別欄背景顏色</span><input type="color" value={uiSettings.shiftColumnBgColor} onChange={(e) => setUiSettings(prev => ({ ...prev, shiftColumnBgColor: e.target.value, themePreset: 'custom' }))} className="w-9 h-7 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">班別欄字體顏色</span><input type="color" value={uiSettings.shiftColumnFontColor} onChange={(e) => setUiSettings(prev => ({ ...prev, shiftColumnFontColor: e.target.value, themePreset: 'custom' }))} className="w-9 h-7 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">姓名/日期欄背景顏色</span><input type="color" value={uiSettings.nameDateColumnBgColor} onChange={(e) => setUiSettings(prev => ({ ...prev, nameDateColumnBgColor: e.target.value, themePreset: 'custom' }))} className="w-9 h-7 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">姓名/日期欄字體顏色</span><input type="color" value={uiSettings.nameDateColumnFontColor} onChange={(e) => setUiSettings(prev => ({ ...prev, nameDateColumnFontColor: e.target.value, themePreset: 'custom' }))} className="w-9 h-7 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">人力需求超標色</span><input type="color" value={uiSettings.demandOverColor} onChange={(e) => setUiSettings(prev => ({ ...prev, demandOverColor: e.target.value, themePreset: 'custom' }))} className="w-9 h-7 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">各班別統計色</span><input type="color" value={uiSettings.groupSummaryRowBgColor || '#fef3c7'} onChange={(e) => setUiSettings(prev => ({ ...prev, groupSummaryRowBgColor: e.target.value, themePreset: 'custom' }))} className="w-9 h-7 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-xl border border-gray-100"><span className="text-sm font-medium">預班/假色</span><input type="color" value={uiSettings.infoTintColor || '#38bdf8'} onChange={(e) => setUiSettings(prev => ({ ...prev, infoTintColor: e.target.value, themePreset: 'custom' }))} className="w-9 h-7 rounded border border-gray-200 bg-transparent cursor-pointer" /></div>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-x-5 gap-y-3">
                  <div className="flex items-center justify-between gap-3"><span className="text-sm font-medium">表格顯示大小</span><select value={uiSettings.tableDensity} onChange={(e) => setUiSettings(prev => ({ ...prev, tableDensity: e.target.value }))} className="text-sm border-none bg-gray-100 rounded-md px-3 py-1.5"><option value="standard">標準 (預設)</option><option value="compact">緊湊</option><option value="relaxed">寬鬆</option></select></div>
                  <div className="flex items-center justify-between gap-3"><span className="text-sm font-medium">表格字體大小</span><select value={uiSettings.tableFontSize} onChange={(e) => setUiSettings(prev => ({ ...prev, tableFontSize: e.target.value }))} className="text-sm border-none bg-gray-100 rounded-md px-3 py-1.5"><option value="small">小</option><option value="medium">標準</option><option value="large">大</option></select></div>
                  <div className="flex items-center justify-between gap-3"><span className="text-sm font-medium">班別欄字體大小</span><select value={uiSettings.shiftColumnFontSize} onChange={(e) => setUiSettings(prev => ({ ...prev, shiftColumnFontSize: e.target.value }))} className="text-sm border-none bg-gray-100 rounded-md px-3 py-1.5"><option value="small">小</option><option value="medium">標準</option><option value="large">大</option></select></div>
                  <div className="flex items-center justify-between gap-3"><span className="text-sm font-medium">姓名/日期欄字體大小</span><select value={uiSettings.nameDateColumnFontSize} onChange={(e) => setUiSettings(prev => ({ ...prev, nameDateColumnFontSize: e.target.value }))} className="text-sm border-none bg-gray-100 rounded-md px-3 py-1.5"><option value="small">小</option><option value="medium">標準</option><option value="large">大</option></select></div>
                  <div className="flex items-center justify-between gap-3"><span className="text-sm font-medium">快速切換全部統計</span><button type="button" onClick={() => setUiSettings(prev => ({ ...prev, showStats: !prev.showStats, showRightStats: !prev.showStats, showLeaveStats: !prev.showStats, showBottomStats: !prev.showStats }))} className={`w-10 h-5 rounded-full relative transition-colors ${uiSettings.showStats ? 'bg-blue-600' : 'bg-gray-300'}`}><div className={`absolute top-1 w-3 h-3 bg-white rounded-full transition-all ${uiSettings.showStats ? 'right-1' : 'left-1'}`}></div></button></div>
                </div>
              </div>
            </SettingRow>
            <SettingRow icon={Settings} title="使用偏好" desc="快速切換主題、欄位顯示、操作方式與預設補休代碼。" iconBg="bg-violet-50" iconColor="text-violet-600">
              <div className="space-y-6">
                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">主題預設</label>
                  <div className="grid grid-cols-2 md:grid-cols-5 gap-2">
                    {[
                      ['classic','經典藍白'],
                      ['soft','柔和綠灰'],
                      ['warm','米色護理站'],
                      ['dark','深色模式'],
                      ['sky','晨霧藍'],
                      ['lavender','石墨灰'],
                      ['forest','森林綠'],
                      ['sakura','櫻花粉白'],
                      ['sand','暖杏橘'],
                      ['custom','自訂主題']
                    ].map(([key,label]) => (
                      <button
                        key={key}
                        type="button"
                        onClick={() => {
                          if (key === 'custom') {
                            setUiSettings(prev => ({ ...prev, themePreset: 'custom' }));
                            return;
                          }
                          const preset = UI_THEME_PRESETS[key];
                          if (!preset) return;
                          setColors(prev => ({ ...prev, weekend: preset.weekendColor, holiday: preset.holidayColor }));
                          setUiSettings(prev => ({
                            ...prev,
                            themePreset: key,
                            pageBackgroundColor: preset.pageBackgroundColor,
                            tableFontColor: preset.tableFontColor,
                            shiftColumnBgColor: preset.shiftColumnBgColor,
                            nameDateColumnBgColor: preset.nameDateColumnBgColor,
                            shiftColumnFontColor: preset.shiftColumnFontColor,
                            nameDateColumnFontColor: preset.nameDateColumnFontColor,
                            demandOverColor: preset.demandOverColor,
                            groupSummaryRowBgColor: preset.groupSummaryRowBgColor,
                            warningTintColor: preset.warningTintColor,
                            warningTextColor: preset.warningTextColor,
                            infoTintColor: preset.infoTintColor,
                            infoTextColor: preset.infoTextColor,
                            dangerTintColor: preset.dangerTintColor,
                            dangerTextColor: preset.dangerTextColor
                          }));
                        }}
                        className={`px-2.5 py-1.5 rounded-lg border text-[13px] font-medium transition leading-5 ${activeThemePreset === key ? 'bg-violet-600 border-violet-600 text-white shadow-sm' : 'bg-white border-violet-200 text-violet-700 hover:bg-violet-50'}`}
                      >
                        {label}
                      </button>
                    ))}
                  </div>
                </div>

                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">欄位顯示</label>
                  <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-2">
                    {[
                      ['showRightStats','上班總休統計'],
                      ['showLeaveStats','休假統計'],
                      ['showBottomStats','下方每日統計'],
                      ['showBlueDots','藍點提示']
                    ].map(([key,label]) => (
                      <div key={key} className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-lg border border-gray-100">
                        <span className="text-sm font-medium">{label}</span>
                        <button
                          type="button"
                          onClick={() => setUiSettings(prev => ({ ...prev, [key]: !prev[key] }))}
                          className={`w-9 h-5 rounded-full relative transition-colors ${uiSettings[key] ? 'bg-violet-600' : 'bg-gray-300'}`}
                        >
                          <div className={`absolute top-1 w-3 h-3 bg-white rounded-full transition-all ${uiSettings[key] ? 'right-1' : 'left-1'}`}></div>
                        </button>
                      </div>
                    ))}
                  </div>
                </div>

                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">寬與高度</label>
                  <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-2">
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-lg border border-gray-100 gap-3"><span className="text-sm font-medium">班別欄寬</span><select value={uiSettings.shiftColumnWidthMode} onChange={(e)=>setUiSettings(prev=>({...prev, shiftColumnWidthMode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-2.5 py-1.5"><option value="narrow">窄</option><option value="standard">標準</option><option value="wide">寬</option></select></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-lg border border-gray-100 gap-3"><span className="text-sm font-medium">姓名/日期欄寬</span><select value={uiSettings.nameDateColumnWidthMode} onChange={(e)=>setUiSettings(prev=>({...prev, nameDateColumnWidthMode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-2.5 py-1.5"><option value="narrow">窄</option><option value="standard">標準</option><option value="wide">寬</option></select></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-lg border border-gray-100 gap-3"><span className="text-sm font-medium">日期欄寬</span><select value={uiSettings.dayColumnWidthMode} onChange={(e)=>setUiSettings(prev=>({...prev, dayColumnWidthMode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-2.5 py-1.5"><option value="narrow">窄</option><option value="standard">標準</option><option value="wide">寬</option></select></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-lg border border-gray-100 gap-3"><span className="text-sm font-medium">儲存格高度</span><select value={uiSettings.cellHeightMode} onChange={(e)=>setUiSettings(prev=>({...prev, cellHeightMode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-2.5 py-1.5"><option value="compact">緊湊</option><option value="standard">標準</option><option value="roomy">寬鬆</option></select></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-lg border border-gray-100 gap-3"><span className="text-sm font-medium">預設補休代碼</span><select value={uiSettings.defaultAutoLeaveCode} onChange={(e)=>setUiSettings(prev=>({...prev, defaultAutoLeaveCode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-2.5 py-1.5">{mergedLeaveCodes.filter(code => ['off','休','例'].includes(code)).map(code => <option key={code} value={code}>{code}</option>)}</select></div>
                    <div className="flex items-center justify-between px-3 py-2 bg-gray-50 rounded-lg border border-gray-100 gap-3"><span className="text-sm font-medium">選取方式</span><select value={uiSettings.selectionMode} onChange={(e)=>setUiSettings(prev=>({...prev, selectionMode:e.target.value}))} className="text-sm border-none bg-white rounded-md px-2.5 py-1.5"><option value="dot">點藍點選格</option><option value="cell">點格選取</option></select></div>
                  </div>
                </div>
              </div>
            </SettingRow>
            <SettingRow icon={Layout} title="班表內容自訂" desc="設定自訂休假、上班代碼與備註欄位。" iconBg="bg-indigo-50" iconColor="text-indigo-600">
              <div className="space-y-4">
                <div className="space-y-2">
                  <div className="flex flex-wrap items-center gap-2">
                    <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mr-2">自訂休假</label>
                    <button type="button" onClick={addCustomLeaveCode} className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-md border border-gray-300 bg-white text-sm font-medium text-gray-700 hover:bg-gray-50">
                      <Plus className="w-3.5 h-3.5" /> 新增
                    </button>
                    {(customLeaveCodes || []).length === 0 ? <div className="text-xs text-gray-400">尚未新增自訂休假</div> : (customLeaveCodes || []).map(code => (
                      <span key={code} className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-gray-100 text-gray-600 text-xs font-bold rounded-md border border-gray-200">{code}<button type="button" onClick={() => removeCustomLeaveCode(code)} className="text-red-500 hover:text-red-600"><Minus className="w-3.5 h-3.5" /></button></span>
                    ))}
                  </div>
                  <div className="text-xs text-gray-500">新增後會同步出現在主頁休假下拉選單，並視為休假類代碼。</div>
                </div>
                <div className="pt-3 border-t border-gray-100 space-y-3">
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block">自訂上班</label>
                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-2">
                    {customWorkShiftGroups.map(({ group, hint }) => {
                      const groupItems = (customWorkShifts || []).filter(item => item.group === group);
                      return (
                        <div key={group} className="rounded-lg border border-gray-200 bg-gray-50 px-3 py-2.5 space-y-2">
                          <div className="flex items-start justify-between gap-2">
                            <div className="min-w-0">
                              <div className="text-sm font-semibold text-gray-700 leading-5">{group}</div>
                              <div className="text-[11px] text-gray-400 leading-4 mt-0.5">{hint}</div>
                            </div>
                            <button type="button" onClick={() => addCustomWorkShift(group)} className="inline-flex items-center gap-1 px-2 py-1.5 rounded-md border border-gray-300 bg-white text-sm font-medium text-gray-700 hover:bg-gray-50 whitespace-nowrap">
                              <Plus className="w-3.5 h-3.5" /> 新增
                            </button>
                          </div>
                          <div className="flex flex-wrap gap-1.5">
                            {groupItems.length === 0 ? <div className="text-xs text-gray-400">尚未新增{group}自訂上班</div> : groupItems.map(item => (
                              <span key={`${item.group}-${item.code}`} className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-indigo-50 text-indigo-700 text-xs font-bold rounded-md border border-indigo-200">{item.code}<button type="button" onClick={() => removeCustomWorkShift(item.group, item.code)} className="text-red-500 hover:text-red-600"><Minus className="w-3.5 h-3.5" /></button></span>
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                  <div className="text-xs text-gray-500">新增後可視為上班代碼使用，並會依白班／小夜／大夜歸入各班別統計。</div>
                </div>
                <div className="pt-3 border-t border-gray-100 space-y-2">
                  <div className="flex flex-wrap items-center gap-2">
                    <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mr-2">備註欄</label>
                    <button type="button" onClick={addCustomColumn} className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-md border border-gray-300 bg-white text-sm font-medium text-gray-700 hover:bg-gray-50">
                      <Plus className="w-3.5 h-3.5" /> 新增
                    </button>
                    {(customColumns || []).length === 0 ? <div className="text-xs text-gray-400">尚未新增備註欄</div> : (customColumns || []).map(col => (
                      <span key={col} className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-violet-50 text-violet-700 text-xs font-bold rounded-md border border-violet-200">{col}<button type="button" onClick={() => removeCustomColumn(col)} className="text-red-500 hover:text-red-600"><Minus className="w-3.5 h-3.5" /></button></span>
                    ))}
                  </div>
                  <div className="text-xs text-gray-500">新增後會同步出現在主頁右側，作為延伸紀錄欄位。可用來記錄如門診、支援、教學、行政或其他單位自訂資訊。</div>
                </div>
                <div className="pt-3 border-t border-gray-100 space-y-2">
                  <div>
                    <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">匯出用排班規則</label>
                    <textarea
                      value={schedulingRulesText}
                      onChange={(e) => setSchedulingRulesText(e.target.value)}
                      rows={4}
                      placeholder={`請逐行輸入排班規則
例如：
白班每日至少 6 人
小夜不跨白班支援`}
                      className="w-full rounded-lg border border-gray-200 bg-gray-50 px-3 py-2.5 text-sm text-gray-700 focus:outline-none focus:ring-2 focus:ring-indigo-100"
                    />
                  </div>
                  <div className="text-xs text-gray-500">這裡輸入的內容不會同步到主頁顯示，但會在匯出 Word 時顯示於最下方，格式為：排班規則：1.XXX 2.XXX。</div>
                </div>
              </div>
            </SettingRow>
            <SettingRow icon={UserCheck} title="規則補空需求設定" desc="設定平日、週六、週日各班需求，作為規則全月補空與規則指定補空的直接依據。" iconBg="bg-sky-50" iconColor="text-sky-600">
              <div className="space-y-4">
                <div className="grid grid-cols-1 xl:grid-cols-3 gap-3">
                  <div className="rounded-xl border border-gray-200 p-3 bg-gray-50/50 space-y-2.5">
                    <h4 className="font-bold text-gray-800">平日需求</h4>
                    <div className="grid grid-cols-3 gap-2">
                      <div>
                        <label className="text-[11px] font-bold text-gray-400 block mb-1.5">白班</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.white}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, white: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-1.5 text-sm border border-gray-200 rounded-lg bg-white" />
                      </div>
                      <div>
                        <label className="text-[11px] font-bold text-gray-400 block mb-1.5">小夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.evening}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, evening: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-1.5 text-sm border border-gray-200 rounded-lg bg-white" />
                      </div>
                      <div>
                        <label className="text-[11px] font-bold text-gray-400 block mb-1.5">大夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.weekday.night}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, weekday: { ...prev.requiredStaffing.weekday, night: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-1.5 text-sm border border-gray-200 rounded-lg bg-white" />
                      </div>
                    </div>
                  </div>

                  <div className="rounded-xl border border-gray-200 p-3 bg-gray-50/50 space-y-2.5">
                    <h4 className="font-bold text-gray-800">週六需求</h4>
                    <div className="grid grid-cols-3 gap-2">
                      <div>
                        <label className="text-[11px] font-bold text-gray-400 block mb-1.5">白班</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.saturday.white}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, saturday: { ...prev.requiredStaffing.saturday, white: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-1.5 text-sm border border-gray-200 rounded-lg bg-white" />
                      </div>
                      <div>
                        <label className="text-[11px] font-bold text-gray-400 block mb-1.5">小夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.saturday.evening}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, saturday: { ...prev.requiredStaffing.saturday, evening: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-1.5 text-sm border border-gray-200 rounded-lg bg-white" />
                      </div>
                      <div>
                        <label className="text-[11px] font-bold text-gray-400 block mb-1.5">大夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.saturday.night}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, saturday: { ...prev.requiredStaffing.saturday, night: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-1.5 text-sm border border-gray-200 rounded-lg bg-white" />
                      </div>
                    </div>
                  </div>

                  <div className="rounded-xl border border-gray-200 p-3 bg-gray-50/50 space-y-2.5">
                    <h4 className="font-bold text-gray-800">週日需求</h4>
                    <div className="grid grid-cols-3 gap-2">
                      <div>
                        <label className="text-[11px] font-bold text-gray-400 block mb-1.5">白班</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.sunday.white}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, sunday: { ...prev.requiredStaffing.sunday, white: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-1.5 text-sm border border-gray-200 rounded-lg bg-white" />
                      </div>
                      <div>
                        <label className="text-[11px] font-bold text-gray-400 block mb-1.5">小夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.sunday.evening}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, sunday: { ...prev.requiredStaffing.sunday, evening: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-1.5 text-sm border border-gray-200 rounded-lg bg-white" />
                      </div>
                      <div>
                        <label className="text-[11px] font-bold text-gray-400 block mb-1.5">大夜</label>
                        <input type="number" min="0" value={staffingConfig.requiredStaffing.sunday.night}
                          onChange={(e) => setStaffingConfig(prev => ({ ...prev, requiredStaffing: { ...prev.requiredStaffing, sunday: { ...prev.requiredStaffing.sunday, night: parseInt(e.target.value, 10) || 0 } } }))}
                          className="w-full px-2 py-1.5 text-sm border border-gray-200 rounded-lg bg-white" />
                      </div>
                    </div>
                  </div>
                </div>

                <div className="rounded-lg border border-gray-200 bg-white px-3 py-2.5 text-sm text-gray-600 space-y-1">
                  <div className="font-semibold text-gray-800">目前補空依據</div>
                  <div>平日：白班 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.weekday.white}</span> 人、小夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.weekday.evening}</span> 人、大夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.weekday.night}</span> 人</div>
                  <div>週六：白班 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.saturday.white}</span> 人、小夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.saturday.evening}</span> 人、大夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.saturday.night}</span> 人</div>
                  <div>週日：白班 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.sunday.white}</span> 人、小夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.sunday.evening}</span> 人、大夜 <span className="font-bold text-sky-700">{staffingConfig.requiredStaffing.sunday.night}</span> 人</div>
                </div>
              </div>
            </SettingRow>
            <SettingRow icon={Calendar} title="假期新增" desc="使用西曆年月日新增自訂假期，並可個別刪除。">
              <div className="space-y-3.5">
                <div>
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">西曆年月日</label>
                  <div className="grid grid-cols-1 md:grid-cols-4 gap-2.5 items-start">
                    <input type="text" inputMode="numeric" placeholder="年" value={holidayInput.year} onChange={(e) => updateHolidayInput('year', e.target.value)} className={getHolidayInputClassName('year')} />
                    <input type="text" inputMode="numeric" placeholder="月" value={holidayInput.month} onChange={(e) => updateHolidayInput('month', e.target.value)} className={getHolidayInputClassName('month')} />
                    <input type="text" inputMode="numeric" placeholder="日" value={holidayInput.day} onChange={(e) => updateHolidayInput('day', e.target.value)} className={getHolidayInputClassName('day')} />
                    <button onClick={addCustomHoliday} className="inline-flex items-center justify-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-gray-300 bg-white text-sm font-medium text-gray-700 hover:bg-gray-50"><Plus className="w-4 h-4" /> 新增假期</button>
                  </div>
                </div>
                <div className="pt-2.5 border-t border-gray-100">
                  <label className="text-xs font-semibold text-gray-400 uppercase tracking-wider block mb-2">已新增假期</label>
                  <div className="space-y-1.5 max-h-44 overflow-y-auto pr-1">
                    {customHolidays.length === 0 ? (
                      <div className="text-xs text-gray-400 p-3 bg-gray-50 border border-dashed border-gray-300 rounded-lg text-center">尚未新增自訂假期</div>
                    ) : customHolidays.map(dateStr => (
                      <div key={dateStr} className="flex items-center justify-between px-3 py-2 bg-gray-50 border border-gray-200 rounded-lg">
                        <span className="text-sm text-gray-700 font-medium">{dateStr}</span>
                        <button onClick={() => removeCustomHoliday(dateStr)} className="w-7 h-7 flex items-center justify-center rounded-full border border-red-200 text-red-500 hover:bg-red-50 font-bold">-</button>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </SettingRow>
          </div>
        </section>
      </main>
      <footer className="max-w-7xl mx-auto px-8 py-12 border-t border-gray-200 flex flex-col md:flex-row justify-between items-center gap-4"><div className="flex items-center gap-4"><div className="flex items-center gap-2"><div className="w-3 h-3 bg-green-500 rounded-full animate-pulse"></div><span className="text-sm font-semibold text-gray-700">系統核心版本: v2.5.0-PRO</span></div><span className="text-gray-300">|</span><span className="text-sm text-gray-500">規則式補班引擎開發測試版</span></div><div className="text-sm text-gray-400">© 2024 Intelligent Scheduling System PRO. All rights reserved.</div></footer>
    </div>
  );
}

function EntryView({ changeScreen, goToLatestHistory, onImportScheduleFiles, onImportPreScheduleFiles, hasActiveDraft, activeDraftMeta, restoreActiveDraft, discardActiveDraft }) {
  const scheduleImportInputRef = useRef(null);
  const preScheduleImportInputRef = useRef(null);

  const handleImportButtonClick = (mode = 'schedule') => {
    if (mode === 'preSchedule') preScheduleImportInputRef.current?.click();
    else scheduleImportInputRef.current?.click();
  };

  const handleImportInputChange = async (event, mode = 'schedule') => {
    const files = Array.from(event.target.files || []);
    event.target.value = '';
    if (files.length === 0) return;

    const invalidFile = files.find((file) => {
      const fileName = String(file.name || '').toLowerCase();
      return !fileName.endsWith('.xlsx') && !fileName.endsWith('.xls');
    });
    if (invalidFile) {
      window.alert('目前僅支援 Excel 檔案（.xlsx / .xls）');
      return;
    }

    try {
      if (mode === 'preSchedule') await onImportPreScheduleFiles?.(files);
      else await onImportScheduleFiles?.(files);
    } catch (error) {
      console.error('匯入檔案失敗:', error);
      window.alert(error?.message || '匯入檔案失敗，請確認是否使用系統範本。');
    }
  };

  const handleDownloadImportTemplate = async () => {
    try {
      const ExcelJS = await loadExcelJS();
      const workbook = new ExcelJS.Workbook();
      const dataSheet = workbook.addWorksheet('班表匯入範本');
      const guideSheet = workbook.addWorksheet('填寫說明');

      const templateMonth = new Date().getMonth() + 1;
      const dayHeaders = Array.from({ length: 31 }, (_, i) => `${i + 1}日`);
      const headers = ['姓名', '班別群組', ...dayHeaders];

      const templateTheme = {
        titleBg: '#EFF6FF',
        titleFont: '#1F2937',
        shiftBg: '#F8FAFC',
        shiftFont: '#1E293B',
        nameBg: '#FFFFFF',
        nameFont: '#1E293B',
        dayBg: '#F8FAFC'
      };

      dataSheet.addRow([`${templateMonth}月班表匯入範本`]);
      dataSheet.mergeCells(1, 1, 1, headers.length);
      const titleCell = dataSheet.getCell(1, 1);
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(templateTheme.titleBg, '#EFF6FF') } };
      titleCell.font = { bold: true, size: 14, color: { argb: hexToExcelArgb(templateTheme.titleFont, '#1F2937') } };
      dataSheet.getRow(1).height = 24;

      const headerRow = dataSheet.addRow(headers);
      headerRow.height = 24;
      headerRow.eachCell((cell, colNumber) => {
        cell.font = { bold: true, size: 10, color: { argb: hexToExcelArgb(templateTheme.titleFont, '#1F2937') } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' },
          bottom: { style: 'thin' }, right: { style: 'thin' }
        };

        if (colNumber === 1) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(templateTheme.shiftBg, '#F8FAFC') } };
          cell.font = { ...(cell.font || {}), color: { argb: hexToExcelArgb(templateTheme.shiftFont, '#1E293B') } };
        } else if (colNumber === 2) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(templateTheme.nameBg, '#FFFFFF') } };
          cell.font = { ...(cell.font || {}), color: { argb: hexToExcelArgb(templateTheme.nameFont, '#1E293B') } };
        } else {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(templateTheme.dayBg, '#F8FAFC') } };
        }
      });

      const sampleRows = [
        ['王小美', '白班'],
        ['李小芳', '小夜'],
        ['陳小君', '大夜']
      ];
      sampleRows.forEach((rowData) => {
        const row = dataSheet.addRow(rowData);
        row.eachCell((cell) => {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' },
            bottom: { style: 'thin' }, right: { style: 'thin' }
          };
        });
      });

      dataSheet.columns = [
        { width: 14 },
        { width: 12 },
        ...Array.from({ length: 31 }, () => ({ width: 6 }))
      ];
      dataSheet.views = [{ state: 'frozen', ySplit: 2, xSplit: 2 }];

      const guideRows = [
        ['排班匯入範本說明'],
        ['1. 請依此範本填寫，避免欄位遺漏或格式不一致。'],
        ['2. 必填欄位為：姓名、班別群組。'],
        ['3. 班別群組請填：白班 / 小夜 / 大夜。'],
        ['4. 日期欄可填班別代碼或假別代碼，例如：D、E、N、off、例、休。'],
        ['5. 若當月不足31天，超出日期欄可留白。'],
        ['6. 匯入功能完成後，建議保留此範本格式，不要自行增刪欄位。']
      ];
      guideRows.forEach(r => guideSheet.addRow(r));
      guideSheet.getCell('A1').font = { bold: true, size: 13 };
      guideSheet.getColumn(1).width = 90;

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = '排班匯入範本.xlsx';
      a.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error('下載匯入範本失敗:', error);
      window.alert('下載匯入範本失敗，請稍後再試。');
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 font-sans text-slate-800">
      <div className="mb-10 text-center">
        <div className="flex items-center justify-center gap-3 mb-4">
          <div className="bg-blue-600 p-2.5 rounded-xl shadow-md shadow-blue-200/50">
            <CalendarDays className="w-8 h-8 text-white" />
          </div>
          <div className="text-left">
            <h1 className="text-2xl font-bold tracking-tight text-slate-900">智慧排班系統</h1>
            <p className="text-slate-500 text-sm font-medium tracking-wide mt-1">護理排班管理平台</p>
          </div>
        </div>
      </div>

      <div className="w-full max-w-[440px] bg-white rounded-2xl shadow-[0_12px_40px_rgba(0,0,0,0.03)] border border-slate-200/60 overflow-hidden">
        <div className="p-10">
          <div className="mb-8 text-center">
            <h2 className="text-xl font-bold text-slate-900 mb-2">系統入口</h2>
            <p className="text-sm text-slate-500 leading-relaxed">
              請選擇要進入的功能。
            </p>
          </div>

          <input
            ref={scheduleImportInputRef}
            type="file"
            multiple
            accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
            className="hidden"
            onChange={(event) => handleImportInputChange(event, 'schedule')}
          />
          <input
            ref={preScheduleImportInputRef}
            type="file"
            multiple
            accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
            className="hidden"
            onChange={(event) => handleImportInputChange(event, 'preSchedule')}
          />

          <div className="space-y-4">
            <button
              type="button"
              onClick={() => changeScreen('schedule')}
              className="w-full flex justify-center items-center gap-2 py-4 px-4 rounded-xl text-sm font-bold text-white bg-blue-600 hover:bg-blue-700"
            >
              <ShieldCheck className="w-4.5 h-4.5" />
              進入排班系統
            </button>

            <button
              type="button"
              onClick={() => handleImportButtonClick('schedule')}
              className="w-full flex justify-center items-center gap-2 py-3.5 px-4 border border-slate-200 rounded-xl shadow-sm text-sm font-bold text-slate-700 bg-white hover:bg-slate-50"
            >
              <Database className="w-4 h-4 text-slate-500" />
              匯入檔案
            </button>

            <button
              type="button"
              onClick={() => handleImportButtonClick('preSchedule')}
              className="w-full flex justify-center items-center gap-2 py-3.5 px-4 border border-slate-200 rounded-xl shadow-sm text-sm font-bold text-slate-700 bg-white hover:bg-slate-50"
            >
              <CalendarDays className="w-4 h-4 text-slate-500" />
              匯入預班表
            </button>

            <button
              type="button"
              onClick={handleDownloadImportTemplate}
              className="w-full flex justify-center items-center gap-2 py-3.5 px-4 border border-slate-200 rounded-xl shadow-sm text-sm font-bold text-slate-700 bg-white hover:bg-slate-50"
            >
              <Download className="w-4 h-4 text-slate-500" />
              下載匯入範本
            </button>

            <button
              type="button"
              onClick={() => window.alert('建議定期匯出 Excel / Word 或後續備份檔，以保留班表資料。')}
              className="w-full flex justify-center items-center gap-2 py-3 px-4 rounded-xl text-xs font-bold text-amber-700 bg-amber-50 border border-amber-200 hover:bg-amber-100"
            >
              <Info className="w-4 h-4" />
              建議定期備份
            </button>
          </div>
        </div>
      </div>

      <div className="mt-12 text-center">
        <p className="text-xs text-slate-400 font-medium tracking-[0.1em] uppercase">&copy; {new Date().getFullYear()} 智能排班系統 PRO. ALL RIGHTS RESERVED.</p>
      </div>
    </div>
  );
}

export default function App() {
  const [screen, setScreen] = useState('entry');
  const [colors, setColors] = useState({ weekend: '#dcfce7', holiday: '#fca5a5' });
  const [customHolidays, setCustomHolidays] = useState([]);
  const [specialWorkdays, setSpecialWorkdays] = useState([]);
  const [medicalCalendarAdjustments, setMedicalCalendarAdjustments] = useState({ holidays: [], workdays: [] });
  const [uiSettings, setUiSettings] = useState({
    pageBackgroundColor: '#f8fafc',
    tableFontSize: 'medium',
    tableFontColor: '#1f2937',
    shiftColumnFontSize: 'medium',
    shiftColumnFontColor: '#1e293b',
    nameDateColumnFontSize: 'medium',
    nameDateColumnFontColor: '#1e293b',
    shiftColumnBgColor: '#ffffff',
    nameDateColumnBgColor: '#ffffff',
    tableDensity: 'standard',
    showStats: true,
    themePreset: 'custom',
    showRightStats: true,
    showLeaveStats: true,
    showBottomStats: true,
    showBlueDots: true,
    showShiftLabels: true,
    defaultAutoLeaveCode: 'off',
    selectionMode: 'dot',
    shiftColumnWidthMode: 'standard',
    nameDateColumnWidthMode: 'standard',
    dayColumnWidthMode: 'standard',
    cellHeightMode: 'standard',
    demandOverColor: '#fde68a',
    groupSummaryRowBgColor: '#fef3c7'
  });
  const [staffingConfig, setStaffingConfig] = useState({
    hospitalLevel: 'regional',
    totalBeds: 60,
    totalNurses: 20,
    requiredStaffing: normalizeRequiredStaffingConfig()
  });
  const [customLeaveCodes, setCustomLeaveCodes] = useState([]);
  const [customWorkShifts, setCustomWorkShifts] = useState([]);
  const [customColumns, setCustomColumns] = useState([]);
  const [customColumnValues, setCustomColumnValues] = useState({});
  const [schedulingRulesText, setSchedulingRulesText] = useState('');
  const [loadLatestOnEnter, setLoadLatestOnEnter] = useState(false);
  const [importedSchedulePayload, setImportedSchedulePayload] = useState(null);
  const [importedPreSchedulePayload, setImportedPreSchedulePayload] = useState(null);
  const [monthlySchedules, setMonthlySchedules] = useState({});
  const [preScheduleMonthlySchedules, setPreScheduleMonthlySchedules] = useState({});
  const [pendingOpenMonthKey, setPendingOpenMonthKey] = useState('');
  const [year, setYear] = useState(2025);
  const [month, setMonth] = useState(3);
  const [staffs, setStaffs] = useState(() => createBlankMonthState(2025, 3).staffs);
  const [schedule, setSchedule] = useState(() => createBlankMonthState(2025, 3).schedule);
  const [hasActiveDraft, setHasActiveDraft] = useState(false);
  const [activeDraftMeta, setActiveDraftMeta] = useState(null);
  const activeDraftHydratedRef = useRef(false);
  const activeDraftSaveReadyRef = useRef(false);
  const draftImportInputRef = useRef(null);

  useEffect(() => {
    setCustomShiftDefsRegistry(customWorkShifts);
  }, [customWorkShifts]);

  useEffect(() => {
    const initialWorkspace = createInitialWorkspaceState();
    applyWorkspaceState(initialWorkspace);
  }, []);

  const formatDraftSavedAt = (isoString) => {
    if (!isoString) return '';
    const date = new Date(isoString);
    if (Number.isNaN(date.getTime())) return '';
    return date.toLocaleString();
  };

  const createDefaultSettingsState = () => ({
    colors: { weekend: '#dcfce7', holiday: '#fca5a5' },
    customHolidays: [],
    specialWorkdays: [],
    medicalCalendarAdjustments: { holidays: [], workdays: [] },
    uiSettings: {
      pageBackgroundColor: '#f8fafc',
      tableFontSize: 'medium',
      tableFontColor: '#1f2937',
      shiftColumnFontSize: 'medium',
      shiftColumnFontColor: '#1e293b',
      nameDateColumnFontSize: 'medium',
      nameDateColumnFontColor: '#1e293b',
      shiftColumnBgColor: '#ffffff',
      nameDateColumnBgColor: '#ffffff',
      tableDensity: 'standard',
      showStats: true,
      themePreset: 'custom',
      showRightStats: true,
      showLeaveStats: true,
      showBottomStats: true,
      showBlueDots: true,
      showShiftLabels: true,
      defaultAutoLeaveCode: 'off',
      selectionMode: 'dot',
      shiftColumnWidthMode: 'standard',
      nameDateColumnWidthMode: 'standard',
      dayColumnWidthMode: 'standard',
      cellHeightMode: 'standard',
      demandOverColor: '#fde68a',
      groupSummaryRowBgColor: '#fef3c7',
      warningTintColor: '#f59e0b',
      warningTextColor: '#92400e',
      infoTintColor: '#38bdf8',
      infoTextColor: '#075985',
      dangerTintColor: '#ef4444',
      dangerTextColor: '#9f1239'
    },
    staffingConfig: {
      hospitalLevel: 'regional',
      totalBeds: 60,
      totalNurses: 20,
      requiredStaffing: normalizeRequiredStaffingConfig()
    },
    customLeaveCodes: [],
    customWorkShifts: [],
    customColumns: [],
    schedulingRulesText: ''
  });

  const sanitizeForLocalStorage = (value) => {
    const seen = new WeakSet();

    const walk = (input) => {
      if (input === null || input === undefined) return input;
      if (typeof input === 'bigint') return Number(input);
      if (typeof input === 'function') return undefined;
      if (typeof input !== 'object') return input;
      if (input instanceof Date) return input.toISOString();
      if (seen.has(input)) return undefined;
      seen.add(input);

      if (Array.isArray(input)) {
        return input.map((item) => walk(item)).filter((item) => item !== undefined);
      }

      const next = {};
      Object.entries(input).forEach(([key, val]) => {
        const walked = walk(val);
        if (walked !== undefined) next[key] = walked;
      });
      return next;
    };

    return walk(value);
  };

  const buildLocalSettingsPayload = () => sanitizeForLocalStorage({
    colors: { ...(colors || {}) },
    customHolidays: Array.isArray(customHolidays) ? [...customHolidays] : [],
    specialWorkdays: Array.isArray(specialWorkdays) ? [...specialWorkdays] : [],
    medicalCalendarAdjustments: {
      holidays: Array.isArray(medicalCalendarAdjustments?.holidays) ? [...medicalCalendarAdjustments.holidays] : [],
      workdays: Array.isArray(medicalCalendarAdjustments?.workdays) ? [...medicalCalendarAdjustments.workdays] : []
    },
    uiSettings: { ...(uiSettings || {}) },
    staffingConfig: {
      hospitalLevel: staffingConfig?.hospitalLevel || 'regional',
      totalBeds: Number(staffingConfig?.totalBeds) || 0,
      totalNurses: Number(staffingConfig?.totalNurses) || 0,
      requiredStaffing: normalizeRequiredStaffingConfig(staffingConfig?.requiredStaffing)
    },
    customLeaveCodes: Array.isArray(customLeaveCodes) ? [...customLeaveCodes] : [],
    customWorkShifts: Array.isArray(customWorkShifts)
      ? customWorkShifts.map((item) => ({
          code: String(item?.code || '').trim(),
          label: String(item?.label || '').trim(),
          group: String(item?.group || '').trim()
        }))
      : [],
    customColumns: Array.isArray(customColumns)
      ? customColumns.map((item) => ({
          key: String(item?.key || '').trim(),
          label: String(item?.label || '').trim()
        }))
      : [],
    schedulingRulesText: typeof schedulingRulesText === 'string' ? schedulingRulesText : ''
  });

  const readLocalSettingsPayload = () => {
    const defaults = createDefaultSettingsState();
    try {
      const stored = localStorage.getItem(LOCAL_SETTINGS_KEY);
      if (!stored) return defaults;
      const parsed = JSON.parse(stored);
      return {
        colors: parsed?.colors || defaults.colors,
        customHolidays: Array.isArray(parsed?.customHolidays) ? parsed.customHolidays : defaults.customHolidays,
        specialWorkdays: Array.isArray(parsed?.specialWorkdays) ? parsed.specialWorkdays : defaults.specialWorkdays,
        medicalCalendarAdjustments: parsed?.medicalCalendarAdjustments || defaults.medicalCalendarAdjustments,
        uiSettings: { ...defaults.uiSettings, ...(parsed?.uiSettings || {}) },
        staffingConfig: {
          hospitalLevel: parsed?.staffingConfig?.hospitalLevel || defaults.staffingConfig.hospitalLevel,
          totalBeds: Number(parsed?.staffingConfig?.totalBeds) || defaults.staffingConfig.totalBeds,
          totalNurses: Number(parsed?.staffingConfig?.totalNurses) || defaults.staffingConfig.totalNurses,
          requiredStaffing: normalizeRequiredStaffingConfig(parsed?.staffingConfig?.requiredStaffing)
        },
        customLeaveCodes: Array.isArray(parsed?.customLeaveCodes) ? parsed.customLeaveCodes : defaults.customLeaveCodes,
        customWorkShifts: Array.isArray(parsed?.customWorkShifts) ? parsed.customWorkShifts : defaults.customWorkShifts,
        customColumns: Array.isArray(parsed?.customColumns) ? parsed.customColumns : defaults.customColumns,
        schedulingRulesText: typeof parsed?.schedulingRulesText === 'string' ? parsed.schedulingRulesText : defaults.schedulingRulesText
      };
    } catch (error) {
      console.error('讀取本機預設設定失敗', error);
      return defaults;
    }
  };

  const applyLocalSettingsPayload = (payload = {}) => {
    const normalized = { ...createDefaultSettingsState(), ...readLocalSettingsPayload(), ...payload };
    setColors(normalized.colors || createDefaultSettingsState().colors);
    setCustomHolidays(Array.isArray(normalized.customHolidays) ? normalized.customHolidays : []);
    setSpecialWorkdays(Array.isArray(normalized.specialWorkdays) ? normalized.specialWorkdays : []);
    setMedicalCalendarAdjustments(normalized.medicalCalendarAdjustments || { holidays: [], workdays: [] });
    setUiSettings(prev => ({ ...createDefaultSettingsState().uiSettings, ...(normalized.uiSettings || {}) }));
    setStaffingConfig({
      hospitalLevel: normalized.staffingConfig?.hospitalLevel || 'regional',
      totalBeds: Number(normalized.staffingConfig?.totalBeds) || 60,
      totalNurses: Number(normalized.staffingConfig?.totalNurses) || 20,
      requiredStaffing: normalizeRequiredStaffingConfig(normalized.staffingConfig?.requiredStaffing)
    });
    setCustomLeaveCodes(Array.isArray(normalized.customLeaveCodes) ? normalized.customLeaveCodes : []);
    setCustomWorkShifts(Array.isArray(normalized.customWorkShifts) ? normalized.customWorkShifts : []);
    setCustomColumns(Array.isArray(normalized.customColumns) ? normalized.customColumns : []);
    setSchedulingRulesText(typeof normalized.schedulingRulesText === 'string' ? normalized.schedulingRulesText : '');
  };

  const saveLocalSettingsPayload = () => {
    try {
      const payload = buildLocalSettingsPayload();
      const serialized = JSON.stringify({
        savedAt: new Date().toISOString(),
        ...payload
      });
      localStorage.setItem(LOCAL_SETTINGS_KEY, serialized);
      return true;
    } catch (error) {
      console.error('儲存本機預設設定失敗', error);
      return false;
    }
  };

  const createInitialWorkspaceState = () => {
    const blankYear = 2025;
    const blankMonth = 3;
    const blankMonthState = createBlankMonthState(blankYear, blankMonth);
    const savedSettings = readLocalSettingsPayload();
    return {
      ...savedSettings,
      customColumnValues: {},
      monthlySchedules: {},
      preScheduleMonthlySchedules: {},
      year: blankYear,
      month: blankMonth,
      staffs: blankMonthState.staffs,
      schedule: blankMonthState.schedule
    };
  };

  const buildWorkspaceState = () => ({
    colors,
    customHolidays,
    specialWorkdays,
    medicalCalendarAdjustments,
    uiSettings,
    staffingConfig,
    customLeaveCodes,
    customWorkShifts,
    customColumns,
    customColumnValues,
    schedulingRulesText,
    monthlySchedules,
    preScheduleMonthlySchedules,
    year,
    month,
    staffs,
    schedule
  });

  const applyWorkspaceState = (state = {}) => {
    setColors(state.colors || { weekend: '#dcfce7', holiday: '#fca5a5' });
    setCustomHolidays(Array.isArray(state.customHolidays) ? state.customHolidays : []);
    setSpecialWorkdays(Array.isArray(state.specialWorkdays) ? state.specialWorkdays : []);
    setMedicalCalendarAdjustments(state.medicalCalendarAdjustments || { holidays: [], workdays: [] });
    setUiSettings({ ...createDefaultSettingsState().uiSettings, ...(state.uiSettings || {}) });
    setStaffingConfig({
      hospitalLevel: state.staffingConfig?.hospitalLevel || 'regional',
      totalBeds: Number(state.staffingConfig?.totalBeds) || 60,
      totalNurses: Number(state.staffingConfig?.totalNurses) || 20,
      requiredStaffing: normalizeRequiredStaffingConfig(state.staffingConfig?.requiredStaffing)
    });
    setCustomLeaveCodes(Array.isArray(state.customLeaveCodes) ? state.customLeaveCodes : []);
    setCustomWorkShifts(Array.isArray(state.customWorkShifts) ? state.customWorkShifts : []);
    setCustomColumns(Array.isArray(state.customColumns) ? state.customColumns : []);
    setCustomColumnValues(state.customColumnValues || {});
    setSchedulingRulesText(typeof state.schedulingRulesText === 'string' ? state.schedulingRulesText : '');
    setMonthlySchedules(state.monthlySchedules || {});
    setPreScheduleMonthlySchedules(state.preScheduleMonthlySchedules || {});
    setYear(Number(state.year) || 2025);
    setMonth(Number(state.month) || 3);
    setStaffs(normalizeStaffGroup(state.staffs || createBlankMonthState(Number(state.year) || 2025, Number(state.month) || 3).staffs));
    setSchedule(state.schedule || createBlankMonthState(Number(state.year) || 2025, Number(state.month) || 3).schedule);
    setImportedSchedulePayload(null);
    setImportedPreSchedulePayload(null);
    setPendingOpenMonthKey('');
    setLoadLatestOnEnter(false);
  };

  useEffect(() => {
    try {
      const storedDraft = localStorage.getItem(ACTIVE_DRAFT_KEY);
      if (!storedDraft) {
        activeDraftHydratedRef.current = true;
        return;
      }
      const parsed = JSON.parse(storedDraft);
      if (parsed?.state) {
        setHasActiveDraft(true);
        setActiveDraftMeta({
          savedAt: parsed.savedAt || '',
          savedAtText: formatDraftSavedAt(parsed.savedAt),
          year: parsed.state.year,
          month: parsed.state.month
        });
      }
    } catch (error) {
      console.error('讀取自動暫存失敗', error);
    } finally {
      activeDraftHydratedRef.current = true;
    }
  }, []);

  useEffect(() => {
    if (!activeDraftHydratedRef.current) return;
    if (screen !== 'schedule') return;
    if (!activeDraftSaveReadyRef.current) {
      activeDraftSaveReadyRef.current = true;
      return;
    }

    const savedAt = new Date().toISOString();
    const payload = {
      savedAt,
      state: {
        colors,
        customHolidays,
        specialWorkdays,
        medicalCalendarAdjustments,
        uiSettings,
        staffingConfig,
        customLeaveCodes,
        customWorkShifts,
        customColumns,
        customColumnValues,
        schedulingRulesText,
        monthlySchedules,
        preScheduleMonthlySchedules,
        year,
        month,
        staffs,
        schedule
      }
    };

    try {
      localStorage.setItem(ACTIVE_DRAFT_KEY, JSON.stringify(payload));
      setHasActiveDraft(true);
      setActiveDraftMeta({ savedAt, savedAtText: formatDraftSavedAt(savedAt), year, month });
    } catch (error) {
      console.error('寫入自動暫存失敗', error);
    }
  }, [screen, colors, customHolidays, specialWorkdays, medicalCalendarAdjustments, uiSettings, staffingConfig, customLeaveCodes, customWorkShifts, customColumns, customColumnValues, schedulingRulesText, monthlySchedules, preScheduleMonthlySchedules, year, month, staffs, schedule]);

  const restoreActiveDraft = () => {
    try {
      const storedDraft = localStorage.getItem(ACTIVE_DRAFT_KEY);
      if (!storedDraft) return;
      const parsed = JSON.parse(storedDraft);
      if (!parsed?.state) return;
      applyWorkspaceState(parsed.state);
      setHasActiveDraft(true);
      setActiveDraftMeta({
        savedAt: parsed.savedAt || '',
        savedAtText: formatDraftSavedAt(parsed.savedAt),
        year: parsed.state.year,
        month: parsed.state.month
      });
      setScreen('schedule');
    } catch (error) {
      console.error('恢復自動暫存失敗', error);
      window.alert('恢復未完成進度失敗，請稍後再試。');
    }
  };

  const discardActiveDraft = () => {
    if (!window.confirm('確定要捨棄上次未完成進度嗎？')) return;
    localStorage.removeItem(ACTIVE_DRAFT_KEY);
    setHasActiveDraft(false);
    setActiveDraftMeta(null);
  };

  const handleSaveSettings = () => {
    setCustomShiftDefsRegistry(customWorkShifts);
    const currentMonthKey = buildMonthKey(year, month);
    const normalizedCurrentSchedule = reconcileScheduleDataMap(schedule, customLeaveCodes);
    const normalizedMonthlySchedules = reconcileMonthStateCollections({
      ...(monthlySchedules || {}),
      [currentMonthKey]: {
        ...(monthlySchedules?.[currentMonthKey] || {}),
        year,
        month,
        staffs: normalizeStaffGroup(staffs),
        scheduleData: normalizedCurrentSchedule,
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
    }, customLeaveCodes);
    const normalizedPreScheduleMonthlySchedules = reconcileMonthStateCollections(preScheduleMonthlySchedules, customLeaveCodes);

    setSchedule(normalizedCurrentSchedule);
    setMonthlySchedules(normalizedMonthlySchedules);
    setPreScheduleMonthlySchedules(normalizedPreScheduleMonthlySchedules);

    const saved = saveLocalSettingsPayload();
    if (!saved) {
      window.alert('儲存設定失敗，請稍後再試。');
      return;
    }
    setScreen('schedule');
  };

  const handleDownloadDraftFile = () => {
    try {
      const exportedAt = new Date().toISOString();
      const payload = {
        type: 'schedule-draft',
        version: '1.6.0',
        exportedAt,
        state: buildWorkspaceState()
      };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `排班工作檔_${year}年${String(month).padStart(2, '0')}月.json`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error('下載工作檔失敗', error);
      window.alert('下載工作檔失敗，請稍後再試。');
    }
  };

  const handleImportDraftFileClick = () => {
    if (draftImportInputRef.current) draftImportInputRef.current.click();
  };

  const handleImportDraftFileChange = async (event) => {
    const file = event.target?.files?.[0];
    if (!file) return;

    try {
      const rawText = await file.text();
      const parsed = JSON.parse(rawText);
      const importedState = parsed?.state || parsed;
      if (!importedState || typeof importedState !== 'object') {
        throw new Error('暫存檔格式不正確');
      }

      applyWorkspaceState(importedState);
      const savedAt = new Date().toISOString();
      const payload = { savedAt, state: importedState };
      localStorage.setItem(ACTIVE_DRAFT_KEY, JSON.stringify(payload));
      setHasActiveDraft(true);
      setActiveDraftMeta({
        savedAt,
        savedAtText: formatDraftSavedAt(savedAt),
        year: importedState.year,
        month: importedState.month
      });
      setScreen('schedule');
      window.alert('匯入工作檔成功，已載入目前工作內容。');
    } catch (error) {
      console.error('匯入工作檔失敗', error);
      window.alert('匯入工作檔失敗，請確認檔案格式是否正確。');
    } finally {
      if (event.target) event.target.value = '';
    }
  };

  const handleImportFiles = async (files) => {
    const imported = await parseImportedExcelFiles(files, new Date().getFullYear(), {
      customLeaveCodes,
      importMode: 'schedule',
      existingMonthlySchedules: monthlySchedules
    });
    if (Array.isArray(imported.unknownCodes) && imported.unknownCodes.length > 0) {
      window.alert(`偵測到尚未建立的代碼：${imported.unknownCodes.join('、')}\n\n請先到系統設定補上代碼；本次匯入已先保留這些代碼。`);
    }
    setImportedSchedulePayload(imported);
    setPendingOpenMonthKey(imported.firstMonthKey || '');
    setLoadLatestOnEnter(false);
    setScreen('schedule');
  };

  const handleImportPreScheduleFiles = async (files) => {
    const imported = await parseImportedExcelFiles(files, new Date().getFullYear(), {
      customLeaveCodes,
      importMode: 'preSchedule',
      existingMonthlySchedules: monthlySchedules,
      existingImportedMonthStates: preScheduleMonthlySchedules
    });
    if (Array.isArray(imported.unknownCodes) && imported.unknownCodes.length > 0) {
      window.alert(`偵測到尚未建立的代碼：${imported.unknownCodes.join('、')}\n\n請先到系統設定補上代碼；本次匯入已先保留這些代碼。`);
    }
    setPreScheduleMonthlySchedules(prev => {
      const next = { ...(prev || {}) };
      Object.entries(imported.monthlySchedules || {}).forEach(([monthKey, monthState]) => {
        next[monthKey] = mergeImportedMonthStates(next[monthKey], monthState);
      });
      return next;
    });
    setImportedPreSchedulePayload(imported);
    setPendingOpenMonthKey(imported.firstMonthKey || '');
    setLoadLatestOnEnter(false);
    setScreen('schedule');
  };

  const goToSchedule = () => {
    setLoadLatestOnEnter(false);
    setPendingOpenMonthKey('');
    setScreen('schedule');
  };

  const returnToEntry = () => {
    applyWorkspaceState(createInitialWorkspaceState());
    setLoadLatestOnEnter(false);
    setPendingOpenMonthKey('');
    activeDraftSaveReadyRef.current = false;
    setScreen('entry');
  };

  const goToLatestHistory = () => {
    setLoadLatestOnEnter(true);
    setPendingOpenMonthKey('');
    setScreen('schedule');
  };

  if (screen === 'schedule') {
    return (
      <ScheduleView
        changeScreen={(target) => {
          if (target === 'entry') returnToEntry();
          else setScreen(target);
        }}
        colors={colors}
        setColors={setColors}
        customHolidays={customHolidays}
        setCustomHolidays={setCustomHolidays}
        specialWorkdays={specialWorkdays}
        setSpecialWorkdays={setSpecialWorkdays}
        medicalCalendarAdjustments={medicalCalendarAdjustments}
        setMedicalCalendarAdjustments={setMedicalCalendarAdjustments}
        staffingConfig={staffingConfig}
        setStaffingConfig={setStaffingConfig}
        uiSettings={uiSettings}
        setUiSettings={setUiSettings}
        customLeaveCodes={customLeaveCodes}
        setCustomLeaveCodes={setCustomLeaveCodes}
        customWorkShifts={customWorkShifts}
        setCustomWorkShifts={setCustomWorkShifts}
        customColumns={customColumns}
        setCustomColumns={setCustomColumns}
        customColumnValues={customColumnValues}
        setCustomColumnValues={setCustomColumnValues}
        schedulingRulesText={schedulingRulesText}
        setSchedulingRulesText={setSchedulingRulesText}
        loadLatestOnEnter={loadLatestOnEnter}
        onLatestLoaded={() => setLoadLatestOnEnter(false)}
        importedSchedulePayload={importedSchedulePayload}
        onImportedScheduleApplied={() => setImportedSchedulePayload(null)}
        monthlySchedules={monthlySchedules}
        setMonthlySchedules={setMonthlySchedules}
        preScheduleMonthlySchedules={preScheduleMonthlySchedules}
        setPreScheduleMonthlySchedules={setPreScheduleMonthlySchedules}
        importedPreSchedulePayload={importedPreSchedulePayload}
        onImportedPreScheduleApplied={() => setImportedPreSchedulePayload(null)}
        pendingOpenMonthKey={pendingOpenMonthKey}
        onPendingOpenHandled={() => setPendingOpenMonthKey('')}
        year={year}
        setYear={setYear}
        month={month}
        setMonth={setMonth}
        staffs={staffs}
        setStaffs={setStaffs}
        schedule={schedule}
        setSchedule={setSchedule}
        onDownloadDraftFile={handleDownloadDraftFile}
        onImportDraftFileClick={handleImportDraftFileClick}
        draftImportInputRef={draftImportInputRef}
        onImportDraftFileChange={handleImportDraftFileChange}
      />
    );
  }

  if (screen === 'settings') {
    return (
      <SettingsView
        changeScreen={(target) => {
          if (target === 'entry') returnToEntry();
          else setScreen(target);
        }}
        colors={colors}
        setColors={setColors}
        customHolidays={customHolidays}
        setCustomHolidays={setCustomHolidays}
        specialWorkdays={specialWorkdays}
        setSpecialWorkdays={setSpecialWorkdays}
        medicalCalendarAdjustments={medicalCalendarAdjustments}
        setMedicalCalendarAdjustments={setMedicalCalendarAdjustments}
        staffingConfig={staffingConfig}
        setStaffingConfig={setStaffingConfig}
        uiSettings={uiSettings}
        setUiSettings={setUiSettings}
        customLeaveCodes={customLeaveCodes}
        setCustomLeaveCodes={setCustomLeaveCodes}
        customWorkShifts={customWorkShifts}
        setCustomWorkShifts={setCustomWorkShifts}
        customColumns={customColumns}
        setCustomColumns={setCustomColumns}
        schedulingRulesText={schedulingRulesText}
        setSchedulingRulesText={setSchedulingRulesText}
        onSaveSettings={handleSaveSettings}
      />
    );
  }

  return (
    <EntryView
      changeScreen={(target) => {
        if (target === 'schedule') goToSchedule();
        else setScreen(target);
      }}
      goToLatestHistory={goToLatestHistory}
      onImportScheduleFiles={handleImportFiles}
      onImportPreScheduleFiles={handleImportPreScheduleFiles}
      hasActiveDraft={hasActiveDraft}
      activeDraftMeta={activeDraftMeta}
      restoreActiveDraft={restoreActiveDraft}
      discardActiveDraft={discardActiveDraft}
    />
  );
}
