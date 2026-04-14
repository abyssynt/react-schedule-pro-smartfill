import { SHIFT_GROUPS } from './scheduleData';

export const buildMonthKey = (year, month) => `${year}-${String(month).padStart(2, '0')}`;

export const normalizeStaffGroup = (staffList = []) => {
  if (!Array.isArray(staffList) || staffList.length === 0) return [];

  const fallbackGroups = [
    '白班', '白班', '白班', '白班', '白班',
    '小夜', '小夜', '小夜', '小夜', '小夜',
    '大夜', '大夜', '大夜', '大夜', '大夜'
  ];

  return staffList.map((staff, index) => ({
    ...staff,
    pregnant: Boolean(staff?.pregnant),
    group: SHIFT_GROUPS.includes(staff?.group) ? staff.group : (fallbackGroups[index] || '白班')
  }));
};

export const createBlankMonthStaffs = (targetYear, targetMonth) => {
  const monthKey = buildMonthKey(targetYear, targetMonth);
  const blankStaffs = Array.from({ length: 15 }, (_, index) => ({
    id: `blank_${monthKey}_${index + 1}`,
    name: '新成員',
    group: SHIFT_GROUPS[Math.floor(index / 5)] || '白班',
    pregnant: false
  }));
  return normalizeStaffGroup(blankStaffs);
};

export const createBlankMonthState = (targetYear, targetMonth) => {
  const blankStaffs = createBlankMonthStaffs(targetYear, targetMonth);
  const blankSchedule = blankStaffs.reduce((acc, staff) => {
    acc[staff.id] = {};
    return acc;
  }, {});

  return {
    staffs: blankStaffs,
    schedule: blankSchedule,
    customColumnValues: {},
    schedulingRulesText: ''
  };
};
