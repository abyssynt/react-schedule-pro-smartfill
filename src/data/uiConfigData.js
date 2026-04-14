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
    nameWidth: 98,
    dayMinWidth: 32,
    dayHeaderClass: 'px-0.5 py-1 text-[11px]',
    statHeaderClass: 'p-1.5',
    leaveHeaderClass: 'p-1',
    cellHeightClass: 'h-8',
    nameCellPaddingClass: 'px-1.5 py-1',
    footCellPaddingClass: 'p-1.5',
    groupLabelClass: '',
    selectorDotClass: 'w-1.5 h-1.5',
    rowMinHeight: 72
  },
  standard: {
    shiftWidth: 76,
    nameWidth: 122,
    dayMinWidth: 42,
    dayHeaderClass: 'px-1.5 py-2 text-xs',
    statHeaderClass: 'p-3',
    leaveHeaderClass: 'p-1.5',
    cellHeightClass: 'h-10',
    nameCellPaddingClass: 'px-2 py-2',
    footCellPaddingClass: 'p-2.5',
    groupLabelClass: '',
    selectorDotClass: 'w-2 h-2',
    rowMinHeight: 80
  },
  relaxed: {
    shiftWidth: 100,
    nameWidth: 156,
    dayMinWidth: 56,
    dayHeaderClass: 'px-2 py-2.5 text-sm',
    statHeaderClass: 'p-4',
    leaveHeaderClass: 'p-2',
    cellHeightClass: 'h-12',
    nameCellPaddingClass: 'px-3 py-2.5',
    footCellPaddingClass: 'p-3',
    groupLabelClass: '',
    selectorDotClass: 'w-3 h-3',
    rowMinHeight: 96
  }
};

export const getUiDensityConfig = (densityKey = 'standard') => UI_DENSITY_OPTIONS[densityKey] || UI_DENSITY_OPTIONS.standard;

export const UI_THEME_PRESETS = {
  classic: {
    pageBackgroundColor: '#f8fafc',
    weekendColor: '#dcfce7',
    holidayColor: '#fca5a5',
    tableFontColor: '#1f2937',
    shiftColumnBgColor: '#ffffff',
    nameDateColumnBgColor: '#ffffff',
    shiftColumnFontColor: '#1e293b',
    nameDateColumnFontColor: '#1e293b',
    demandOverColor: '#fde68a'
  },
  soft: {
    pageBackgroundColor: '#f7faf7',
    weekendColor: '#e7f7ec',
    holidayColor: '#f6c7c7',
    tableFontColor: '#334155',
    shiftColumnBgColor: '#f7fbf8',
    nameDateColumnBgColor: '#fcfdfc',
    shiftColumnFontColor: '#365314',
    nameDateColumnFontColor: '#334155',
    demandOverColor: '#fde68a'
  },
  warm: {
    pageBackgroundColor: '#fffaf5',
    weekendColor: '#fef3c7',
    holidayColor: '#fecaca',
    tableFontColor: '#44403c',
    shiftColumnBgColor: '#fff7ed',
    nameDateColumnBgColor: '#fffbeb',
    shiftColumnFontColor: '#7c2d12',
    nameDateColumnFontColor: '#44403c',
    demandOverColor: '#fdba74'
  },
  dark: {
    pageBackgroundColor: '#0f172a',
    weekendColor: '#334155',
    holidayColor: '#7f1d1d',
    tableFontColor: '#e2e8f0',
    shiftColumnBgColor: '#1e293b',
    nameDateColumnBgColor: '#172033',
    shiftColumnFontColor: '#f8fafc',
    nameDateColumnFontColor: '#e2e8f0',
    demandOverColor: '#78350f'
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
    nameWidth: Math.max(90, baseConfig.nameWidth + nameAdjust),
    dayMinWidth: Math.max(28, baseConfig.dayMinWidth + dayAdjust),
    rowMinHeight: Math.max(72, (baseConfig.rowMinHeight || 80) + heightAdjust * 4),
    selectorDotClass: dotClassMap[uiSettings.cellHeightMode || 'standard'] || baseConfig.selectorDotClass
  };
};
