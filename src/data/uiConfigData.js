export const UI_FONT_SIZE_OPTIONS = {
  small: { label: '小', className: 'text-xs', shiftLabelSize: '1.45rem', shiftCellLabelSize: '1.55rem' },
  medium: { label: '標準', className: 'text-sm', shiftLabelSize: '1.95rem', shiftCellLabelSize: '2.05rem' },
  large: { label: '大', className: 'text-base', shiftLabelSize: '2.45rem', shiftCellLabelSize: '2.55rem' }
};

export const getUiFontSizeClass = (sizeKey = 'medium') => UI_FONT_SIZE_OPTIONS[sizeKey]?.className || UI_FONT_SIZE_OPTIONS.medium.className;
export const getShiftLabelFontSize = (sizeKey = 'medium') => UI_FONT_SIZE_OPTIONS[sizeKey]?.shiftLabelSize || UI_FONT_SIZE_OPTIONS.medium.shiftLabelSize;
export const getShiftCellLabelFontSize = (sizeKey = 'medium') => UI_FONT_SIZE_OPTIONS[sizeKey]?.shiftCellLabelSize || UI_FONT_SIZE_OPTIONS.medium.shiftCellLabelSize;

export const UI_DENSITY_OPTIONS = {
  compact: {
    shiftWidth: 58,
    nameWidth: 84,
    dayMinWidth: 32,
    dayHeaderClass: 'px-0.5 py-1 text-[11px]',
    statHeaderClass: 'p-1.5',
    leaveHeaderClass: 'p-1',
    cellHeightClass: 'h-8',
    nameCellPaddingClass: 'px-0.5 py-0.5',
    footCellPaddingClass: 'p-1.5',
    groupLabelClass: '',
    selectorDotClass: 'w-1.5 h-1.5',
    rowMinHeight: 72
  },
  standard: {
    shiftWidth: 68,
    nameWidth: 84,
    dayMinWidth: 52,
    dayHeaderClass: 'px-1.5 py-2 text-xs',
    statHeaderClass: 'px-1 py-1',
    leaveHeaderClass: 'px-1 py-1.5',
    cellHeightClass: 'h-9',
    nameCellPaddingClass: 'px-1 py-1',
    footCellPaddingClass: 'px-1 py-1',
    groupLabelClass: '',
    selectorDotClass: 'w-2 h-2',
    rowMinHeight: 72
  },
  relaxed: {
    shiftWidth: 100,
    nameWidth: 156,
    dayMinWidth: 68,
    dayHeaderClass: 'px-1 py-1.5 text-sm',
    statHeaderClass: 'p-4',
    leaveHeaderClass: 'p-2',
    cellHeightClass: 'h-12',
    nameCellPaddingClass: 'px-2 py-2',
    footCellPaddingClass: 'p-3',
    groupLabelClass: '',
    selectorDotClass: 'w-3 h-3',
    rowMinHeight: 96
  }
};

export const getUiDensityConfig = (densityKey = 'standard') => UI_DENSITY_OPTIONS[densityKey] || UI_DENSITY_OPTIONS.standard;

export const UI_THEME_PRESETS = {
  classic: {
    pageBackgroundColor: '#f8fbff',
    weekendColor: '#dbeafe',
    holidayColor: '#fca5a5',
    tableFontColor: '#1f3b5b',
    shiftColumnBgColor: '#ffffff',
    nameDateColumnBgColor: '#f8fbff',
    shiftColumnFontColor: '#1d4ed8',
    nameDateColumnFontColor: '#1f3b5b',
    demandOverColor: '#fde68a',
    groupSummaryRowBgColor: '#dbeafe',
    warningTintColor: '#60a5fa',
    warningTextColor: '#1d4ed8',
    infoTintColor: '#38bdf8',
    infoTextColor: '#075985',
    dangerTintColor: '#ef4444',
    dangerTextColor: '#991b1b'
  },
  soft: {
    pageBackgroundColor: '#f5faf7',
    weekendColor: '#dcfce7',
    holidayColor: '#f9a8d4',
    tableFontColor: '#334155',
    shiftColumnBgColor: '#fbfefc',
    nameDateColumnBgColor: '#f8fcfa',
    shiftColumnFontColor: '#3f6212',
    nameDateColumnFontColor: '#334155',
    demandOverColor: '#d9f99d',
    groupSummaryRowBgColor: '#ecfccb',
    warningTintColor: '#84cc16',
    warningTextColor: '#3f6212',
    infoTintColor: '#5eead4',
    infoTextColor: '#115e59',
    dangerTintColor: '#fb7185',
    dangerTextColor: '#9f1239'
  },
  warm: {
    pageBackgroundColor: '#fff9f2',
    weekendColor: '#ffedd5',
    holidayColor: '#fdba74',
    tableFontColor: '#7c2d12',
    shiftColumnBgColor: '#fffdf9',
    nameDateColumnBgColor: '#fffbf5',
    shiftColumnFontColor: '#c2410c',
    nameDateColumnFontColor: '#7c2d12',
    demandOverColor: '#fed7aa',
    groupSummaryRowBgColor: '#ffedd5',
    warningTintColor: '#fb923c',
    warningTextColor: '#9a3412',
    infoTintColor: '#fdba74',
    infoTextColor: '#9a3412',
    dangerTintColor: '#ef4444',
    dangerTextColor: '#991b1b'
  },
  dark: {
    pageBackgroundColor: '#0f172a',
    weekendColor: '#1e293b',
    holidayColor: '#7f1d1d',
    tableFontColor: '#e2e8f0',
    shiftColumnBgColor: '#111827',
    nameDateColumnBgColor: '#172033',
    shiftColumnFontColor: '#f8fafc',
    nameDateColumnFontColor: '#e2e8f0',
    demandOverColor: '#78350f',
    groupSummaryRowBgColor: '#334155',
    warningTintColor: '#fbbf24',
    warningTextColor: '#fef3c7',
    infoTintColor: '#38bdf8',
    infoTextColor: '#e0f2fe',
    dangerTintColor: '#fb7185',
    dangerTextColor: '#ffe4e6'
  },
  sky: {
    pageBackgroundColor: '#f2f8ff',
    weekendColor: '#cfe8ff',
    holidayColor: '#fca5a5',
    tableFontColor: '#0f3d66',
    shiftColumnBgColor: '#ffffff',
    nameDateColumnBgColor: '#f7fbff',
    shiftColumnFontColor: '#0369a1',
    nameDateColumnFontColor: '#0f3d66',
    demandOverColor: '#bae6fd',
    groupSummaryRowBgColor: '#dbeafe',
    warningTintColor: '#0ea5e9',
    warningTextColor: '#075985',
    infoTintColor: '#7dd3fc',
    infoTextColor: '#075985',
    dangerTintColor: '#f87171',
    dangerTextColor: '#991b1b'
  },
  lavender: {
    pageBackgroundColor: '#f5f5f7',
    weekendColor: '#e5e7eb',
    holidayColor: '#d1d5db',
    tableFontColor: '#111827',
    shiftColumnBgColor: '#ffffff',
    nameDateColumnBgColor: '#fafafa',
    shiftColumnFontColor: '#374151',
    nameDateColumnFontColor: '#111827',
    demandOverColor: '#d1d5db',
    groupSummaryRowBgColor: '#e5e7eb',
    warningTintColor: '#9ca3af',
    warningTextColor: '#374151',
    infoTintColor: '#94a3b8',
    infoTextColor: '#334155',
    dangerTintColor: '#ef4444',
    dangerTextColor: '#991b1b'
  },
  forest: {
    pageBackgroundColor: '#f3faf5',
    weekendColor: '#d1fae5',
    holidayColor: '#fca5a5',
    tableFontColor: '#14532d',
    shiftColumnBgColor: '#ffffff',
    nameDateColumnBgColor: '#f7fcf8',
    shiftColumnFontColor: '#166534',
    nameDateColumnFontColor: '#14532d',
    demandOverColor: '#86efac',
    groupSummaryRowBgColor: '#d1fae5',
    warningTintColor: '#22c55e',
    warningTextColor: '#166534',
    infoTintColor: '#34d399',
    infoTextColor: '#065f46',
    dangerTintColor: '#ef4444',
    dangerTextColor: '#991b1b'
  },
  sakura: {
    pageBackgroundColor: '#fff7fb',
    weekendColor: '#fce7f3',
    holidayColor: '#fb7185',
    tableFontColor: '#9d174d',
    shiftColumnBgColor: '#ffffff',
    nameDateColumnBgColor: '#fffafd',
    shiftColumnFontColor: '#be185d',
    nameDateColumnFontColor: '#9d174d',
    demandOverColor: '#fbcfe8',
    groupSummaryRowBgColor: '#fce7f3',
    warningTintColor: '#f472b6',
    warningTextColor: '#be185d',
    infoTintColor: '#f9a8d4',
    infoTextColor: '#9d174d',
    dangerTintColor: '#ef4444',
    dangerTextColor: '#991b1b'
  },
  sand: {
    pageBackgroundColor: '#fff8ef',
    weekendColor: '#fde68a',
    holidayColor: '#fb923c',
    tableFontColor: '#78350f',
    shiftColumnBgColor: '#fffdf8',
    nameDateColumnBgColor: '#fffbf3',
    shiftColumnFontColor: '#92400e',
    nameDateColumnFontColor: '#78350f',
    demandOverColor: '#fdba74',
    groupSummaryRowBgColor: '#fef3c7',
    warningTintColor: '#f59e0b',
    warningTextColor: '#92400e',
    infoTintColor: '#fbbf24',
    infoTextColor: '#92400e',
    dangerTintColor: '#dc2626',
    dangerTextColor: '#991b1b'
  }
};

export const WIDTH_ADJUST_MAP = { narrow: -12, standard: 0, wide: 12 };
export const HEIGHT_ADJUST_MAP = { compact: -4, standard: 0, roomy: 4 };

export const getAdjustedDensityConfig = (baseConfig, uiSettings = {}) => {
  const shiftAdjust = WIDTH_ADJUST_MAP[uiSettings.shiftColumnWidthMode || 'standard'] || 0;
  const nameAdjust = WIDTH_ADJUST_MAP[uiSettings.nameDateColumnWidthMode || 'standard'] || 0;
  const dayAdjust = WIDTH_ADJUST_MAP[uiSettings.dayColumnWidthMode || 'standard'] || 0;
  const heightAdjust = HEIGHT_ADJUST_MAP[uiSettings.cellHeightMode || 'standard'] || 0;
  const dotClassMap = { compact: 'w-1.5 h-1.5', standard: 'w-2 h-2', roomy: 'w-2.5 h-2.5' };
  return {
    ...baseConfig,
    shiftWidth: Math.max(48, baseConfig.shiftWidth + shiftAdjust),
    nameWidth: Math.max(76, baseConfig.nameWidth + nameAdjust),
    dayMinWidth: Math.max(28, baseConfig.dayMinWidth + dayAdjust),
    rowMinHeight: Math.max(72, (baseConfig.rowMinHeight || 80) + heightAdjust * 4),
    selectorDotClass: dotClassMap[uiSettings.cellHeightMode || 'standard'] || baseConfig.selectorDotClass
  };
};
