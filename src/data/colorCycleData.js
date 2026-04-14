export const clampColorChannel = (value) => Math.max(0, Math.min(255, Math.round(value)));

export const normalizeHexColor = (hex, fallback = '#000000') => {
  const raw = String(hex || '').trim();
  if (/^#[0-9a-fA-F]{6}$/.test(raw)) return raw;
  if (/^#[0-9a-fA-F]{3}$/.test(raw)) {
    return `#${raw[1]}${raw[1]}${raw[2]}${raw[2]}${raw[3]}${raw[3]}`;
  }
  return fallback;
};

export const hexToRgbObject = (hex, fallback = '#000000') => {
  const normalized = normalizeHexColor(hex, fallback).replace('#', '');
  return {
    r: parseInt(normalized.slice(0, 2), 16),
    g: parseInt(normalized.slice(2, 4), 16),
    b: parseInt(normalized.slice(4, 6), 16)
  };
};

export const rgbObjectToHex = ({ r, g, b }) => {
  const toHex = (value) => clampColorChannel(value).toString(16).padStart(2, '0');
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
};

export const blendHexColors = (baseHex, mixHex, mixRatio = 0.5) => {
  const ratio = Math.max(0, Math.min(1, Number(mixRatio) || 0));
  const base = hexToRgbObject(baseHex, '#ffffff');
  const mix = hexToRgbObject(mixHex, '#ffffff');
  return rgbObjectToHex({
    r: base.r * (1 - ratio) + mix.r * ratio,
    g: base.g * (1 - ratio) + mix.g * ratio,
    b: base.b * (1 - ratio) + mix.b * ratio
  });
};

export const hexToExcelArgb = (hex, fallback = '#FFFFFF') => {
  return `FF${normalizeHexColor(hex, fallback).replace('#', '').toUpperCase()}`;
};

export const FOUR_WEEK_CYCLE_START = '2026-04-13';
export const FOUR_WEEK_CYCLE_DAYS = 28;
export const RULE_CROSS_MONTH_CONTEXT_DAYS = 7;

export const isFourWeekCycleEndDate = (dateStr, parseDateKey, cycleStart = FOUR_WEEK_CYCLE_START) => {
  if (!dateStr || typeof parseDateKey !== 'function') return false;
  const target = parseDateKey(dateStr);
  const start = parseDateKey(cycleStart);
  target.setHours(0, 0, 0, 0);
  start.setHours(0, 0, 0, 0);
  const diffDays = Math.floor((target.getTime() - start.getTime()) / 86400000);
  const cycleOffset = ((diffDays + 1) % FOUR_WEEK_CYCLE_DAYS + FOUR_WEEK_CYCLE_DAYS) % FOUR_WEEK_CYCLE_DAYS;
  return cycleOffset === 0;
};
