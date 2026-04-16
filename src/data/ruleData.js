import {
  DICT,
  DEFAULT_REQUIRED_STAFFING,
  DEFAULT_SHIFT_BY_GROUP,
  GROUP_TO_DEMAND_KEY,
  HOSPITAL_LEVEL_LABELS,
  HOSPITAL_RATIO_HINTS,
  RULE_FILL_MAIN_SHIFTS
} from './scheduleData.js';
import { BLOCKED_LEAVE_PREFIXES, getAllShiftCodes, getCodePrefix } from './shiftResolverData.js';

export const SMART_RULES = {
  maxConsecutiveWorkDays: 5,
  allowCrossGroupAssignment: false,
  disallowedNextShiftMap: {
    N: ['D', 'E'],
    E: ['N', 'D'],
    D: ['N'],
    '白8-8': ['D', 'N'],
    '夜8-8': ['E', 'N']
  },
  blockedLeavePrefixes: BLOCKED_LEAVE_PREFIXES,
  pregnancyRestrictedShifts: ['N', '夜8-8'],
  fillPriorityWeights: {
    sameShiftCount: 3,
    totalShiftCount: 2,
    sameGroup: 1
  }
};

export const isLeaveCode = (code = '') => SMART_RULES.blockedLeavePrefixes.includes(getCodePrefix(code));
export const isShiftCode = (code = '') => getAllShiftCodes().includes(code);

export const normalizeRequiredStaffingConfig = (requiredStaffing = {}) => {
  const holidayFallback = requiredStaffing?.holiday || {};
  return {
    weekday: {
      white: Number(requiredStaffing?.weekday?.white ?? DEFAULT_REQUIRED_STAFFING.weekday.white) || 0,
      evening: Number(requiredStaffing?.weekday?.evening ?? DEFAULT_REQUIRED_STAFFING.weekday.evening) || 0,
      night: Number(requiredStaffing?.weekday?.night ?? DEFAULT_REQUIRED_STAFFING.weekday.night) || 0,
    },
    saturday: {
      white: Number(requiredStaffing?.saturday?.white ?? holidayFallback?.white ?? DEFAULT_REQUIRED_STAFFING.saturday.white) || 0,
      evening: Number(requiredStaffing?.saturday?.evening ?? holidayFallback?.evening ?? DEFAULT_REQUIRED_STAFFING.saturday.evening) || 0,
      night: Number(requiredStaffing?.saturday?.night ?? holidayFallback?.night ?? DEFAULT_REQUIRED_STAFFING.saturday.night) || 0,
    },
    sunday: {
      white: Number(requiredStaffing?.sunday?.white ?? holidayFallback?.white ?? DEFAULT_REQUIRED_STAFFING.sunday.white) || 0,
      evening: Number(requiredStaffing?.sunday?.evening ?? holidayFallback?.evening ?? DEFAULT_REQUIRED_STAFFING.sunday.evening) || 0,
      night: Number(requiredStaffing?.sunday?.night ?? holidayFallback?.night ?? DEFAULT_REQUIRED_STAFFING.sunday.night) || 0,
    }
  };
};

export const getRequiredStaffingBucketByDay = (day) => {
  if (!day) return 'weekday';
  if (day.isHoliday) return 'sunday';
  const date = typeof day.date === 'string' ? new Date(`${day.date}T00:00:00`) : null;
  const weekDay = date ? date.getDay() : null;
  if (weekDay === 6) return 'saturday';
  if (weekDay === 0) return 'sunday';
  return 'weekday';
};

export {
  DICT,
  DEFAULT_REQUIRED_STAFFING,
  DEFAULT_SHIFT_BY_GROUP,
  GROUP_TO_DEMAND_KEY,
  HOSPITAL_LEVEL_LABELS,
  HOSPITAL_RATIO_HINTS,
  RULE_FILL_MAIN_SHIFTS
};
