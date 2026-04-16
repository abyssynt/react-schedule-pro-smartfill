export const createHandleRuleBasedAutoSchedule = (deps = {}) => {
  const {
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
  } = deps;

  return async (isPartial = false) => {
    setIsRuleFillLoading(true);
    setRuleFillFeedback(isPartial ? "🧩 系統正在依指定範圍進行規則補空..." : "🧩 系統正在依人力需求執行整月規則補空...");

    try {
      const mergedSchedule = JSON.parse(JSON.stringify(schedule));
      const targetStaffIds = isPartial && ruleFillConfig.selectedStaffs.length > 0
        ? new Set(ruleFillConfig.selectedStaffs)
        : new Set(staffs.map(s => s.id));

      const targetDays = daysInMonth.filter(d => {
        if (!isPartial) return true;
        return d.day >= ruleFillConfig.dateRange.start && d.day <= ruleFillConfig.dateRange.end;
      });

      const normalizedTargetShift = RULE_FILL_MAIN_SHIFTS.includes(ruleFillConfig.targetShift) ? ruleFillConfig.targetShift : '';
      const restrictedGroup = normalizedTargetShift ? getShiftGroupByCode(normalizedTargetShift) : null;
      const summary = { workFilled: 0, leaveFilled: 0, skipped: 0 };
      const touchedRuleFillCellMap = new Map();
      const markRuleFillCellTouched = (staffId, dateStr) => {
        if (!staffId || !dateStr) return;
        touchedRuleFillCellMap.set(makeCellKey(staffId, dateStr), { staffId, dateStr });
      };

      const getScheduleCode = (snapshot, staffRef, dateStr) => {
        return getContextCellCode(staffRef, dateStr, { snapshot });
      };

      const setScheduleCode = (snapshot, staffId, dateStr, value, source = 'auto') => {
        if (!snapshot[staffId]) snapshot[staffId] = {};
        snapshot[staffId][dateStr] = value ? { value, source } : null;
        markRuleFillCellTouched(staffId, dateStr);
      };

      const getDemandType = (day) => getRequiredStaffingBucketByDay(day);
      const getDemandForGroup = (day, group) => {
        const bucket = getDemandType(day);
        const key = GROUP_TO_DEMAND_KEY[group];
        return Number(staffingConfig?.requiredStaffing?.[bucket]?.[key] || 0);
      };

      const getAssignedCountByGroup = (snapshot, dateStr, group) => {
        return staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => {
          const code = getScheduleCode(snapshot, s, dateStr);
          return sum + (getShiftGroupByCode(code) === group ? 1 : 0);
        }, 0);
      };

      const countConsecutiveBeforeFromSnapshot = (snapshot, staffRef, dateStr) => {
        let count = 0;
        let cursor = addDays(parseDateKey(dateStr), -1);
        while (true) {
          const key = formatDateKey(cursor);
          const code = getScheduleCode(snapshot, staffRef, key);
          if (!isShiftCode(code)) break;
          count += 1;
          cursor = addDays(cursor, -1);
        }
        return count;
      };

      const canAssignWithSnapshot = (snapshot, staff, dateStr, shiftCode) => {
        const reasons = [];
        const currentCode = getScheduleCode(snapshot, staff, dateStr);
        const preScheduleCode = getVisiblePreScheduleCode(staff, dateStr);
        const hasBlockedPreScheduleCode = Boolean(preScheduleCode);
        if (currentCode) reasons.push('該格已有排班或休假代碼');
        if (hasBlockedPreScheduleCode) reasons.push('該格已有預班／預假，不可被規則補空覆蓋');
        if (isConfiguredLeaveCode(currentCode)) reasons.push('該格已有休假，不可再排班');
        const staffGroup = staff.group || '白班';
        const shiftGroup = getShiftGroupByCode(shiftCode);
        if (!SMART_RULES.allowCrossGroupAssignment && shiftGroup && staffGroup !== shiftGroup) reasons.push('不可跨群組排班');
        const prevKey = formatDateKey(addDays(parseDateKey(dateStr), -1));
        const prevCode = getScheduleCode(snapshot, staff, prevKey);
        const disallowed = SMART_RULES.disallowedNextShiftMap[prevCode] || [];
        if (disallowed.includes(shiftCode)) reasons.push(`${prevCode} 後不可接 ${shiftCode}`);
        const consecutiveBefore = countConsecutiveBeforeFromSnapshot(snapshot, staff, dateStr);
        if (consecutiveBefore + 1 > SMART_RULES.maxConsecutiveWorkDays) reasons.push(`連續上班不可超過 ${SMART_RULES.maxConsecutiveWorkDays} 天`);
        if (staff.pregnant && SMART_RULES.pregnancyRestrictedShifts.includes(shiftCode)) reasons.push('懷孕標記人員不可排 N / 夜8-8');
        return { allowed: reasons.length === 0, reasons };
      };

      const getWorkCountFromSnapshot = (snapshot, staffId) => {
        const staffRef = getStaffRefFromCurrentMonth(staffId) || staffId;
        return daysInMonth.reduce((sum, d) => sum + (isShiftCode(getScheduleCode(snapshot, staffRef, d.date)) ? 1 : 0), 0);
      };

      const getShiftCountFromSnapshot = (snapshot, staffId, shiftCode) => {
        const staffRef = getStaffRefFromCurrentMonth(staffId) || staffId;
        return daysInMonth.reduce((sum, d) => sum + (getScheduleCode(snapshot, staffRef, d.date) === shiftCode ? 1 : 0), 0);
      };

      const getLeaveCountFromSnapshot = (snapshot, staffId) => {
        const staffRef = getStaffRefFromCurrentMonth(staffId) || staffId;
        return daysInMonth.reduce((sum, d) => sum + (isConfiguredLeaveCode(getScheduleCode(snapshot, staffRef, d.date)) ? 1 : 0), 0);
      };

      const getBlankCountFromSnapshot = (snapshot, staffId) => {
        const staffRef = getStaffRefFromCurrentMonth(staffId) || staffId;
        return daysInMonth.reduce((sum, d) => sum + (!getScheduleCode(snapshot, staffRef, d.date) ? 1 : 0), 0);
      };

      const canStillMeetRequiredLeavesAfterAssign = (snapshot, staffId) => {
        const currentLeaves = getLeaveCountFromSnapshot(snapshot, staffId);
        const remainingBlanks = getBlankCountFromSnapshot(snapshot, staffId);
        const remainingLeavesNeeded = Math.max(0, requiredLeaves - currentLeaves);
        return remainingBlanks >= remainingLeavesNeeded;
      };

      const canStillMeetRequiredLeavesIfAssignShift = (snapshot, staffId) => {
        const currentLeaves = getLeaveCountFromSnapshot(snapshot, staffId);
        const remainingBlanks = getBlankCountFromSnapshot(snapshot, staffId);
        const remainingLeavesNeeded = Math.max(0, requiredLeaves - currentLeaves);
        return (remainingBlanks - 1) >= remainingLeavesNeeded;
      };

      const getRecentWorkPressure = (snapshot, staffId, dateStr, lookback = 3) => {
        const staffRef = getStaffRefFromCurrentMonth(staffId) || staffId;
        let count = 0;
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 0; i < lookback; i += 1) {
          const code = getScheduleCode(snapshot, staffRef, formatDateKey(cursor));
          if (isShiftCode(code)) count += 1;
          cursor = addDays(cursor, -1);
        }
        return count;
      };

      const getRecentLeavePressure = (snapshot, staffId, dateStr, lookback = 4) => {
        const staffRef = getStaffRefFromCurrentMonth(staffId) || staffId;
        let count = 0;
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 0; i < lookback; i += 1) {
          const code = getScheduleCode(snapshot, staffRef, formatDateKey(cursor));
          if (isLeaveCode(code)) count += 1;
          cursor = addDays(cursor, -1);
        }
        return count;
      };

      const getNearbyLeavePressure = (snapshot, staffId, dateStr, radius = 2) => {
        const staffRef = getStaffRefFromCurrentMonth(staffId) || staffId;
        let count = 0;
        const center = parseDateKey(dateStr);
        for (let offset = -radius; offset <= radius; offset += 1) {
          if (offset === 0) continue;
          const key = formatDateKey(addDays(center, offset));
          const code = getScheduleCode(snapshot, staffRef, key);
          if (isLeaveCode(code)) count += 1;
        }
        return count;
      };

      const getDaysSinceLastLeave = (snapshot, staffId, dateStr, maxLookback = 10) => {
        const staffRef = getStaffRefFromCurrentMonth(staffId) || staffId;
        let cursor = addDays(parseDateKey(dateStr), -1);
        for (let i = 1; i <= maxLookback; i += 1) {
          const code = getScheduleCode(snapshot, staffRef, formatDateKey(cursor));
          if (isLeaveCode(code)) return i;
          cursor = addDays(cursor, -1);
        }
        return maxLookback + 1;
      };

      const getGroupLeaveLoad = (snapshot, dateStr, group) => {
        return staffs.filter(s => (s.group || '白班') === group).reduce((sum, s) => {
          const code = getScheduleCode(snapshot, s, dateStr);
          return sum + (isLeaveCode(code) ? 1 : 0);
        }, 0);
      };

      const getConsecutiveLeavePattern = (snapshot, staffId, dateStr) => {
        const staffRef = getStaffRefFromCurrentMonth(staffId) || staffId;
        const prevCode = getScheduleCode(snapshot, staffRef, formatDateKey(addDays(parseDateKey(dateStr), -1)));
        const nextCode = getScheduleCode(snapshot, staffRef, formatDateKey(addDays(parseDateKey(dateStr), 1)));
        const prevIsLeave = isLeaveCode(prevCode);
        const nextIsLeave = isLeaveCode(nextCode);
        return {
          prevIsLeave,
          nextIsLeave,
          adjacentLeaveCount: (prevIsLeave ? 1 : 0) + (nextIsLeave ? 1 : 0)
        };
      };

      const scoreCandidateWithSnapshot = (snapshot, staff, dateStr, shiftCode) => {
        let score = 0;
        score += (999 - getShiftCountFromSnapshot(snapshot, staff.id, shiftCode)) * SMART_RULES.fillPriorityWeights.sameShiftCount;
        score += (999 - getWorkCountFromSnapshot(snapshot, staff.id)) * SMART_RULES.fillPriorityWeights.totalShiftCount;
        if (getShiftGroupByCode(shiftCode) === (staff.group || '白班')) score += 100 * SMART_RULES.fillPriorityWeights.sameGroup;
        score -= getRecentWorkPressure(snapshot, staff.id, dateStr, 3) * 18;
        return score;
      };

      const scoreLeaveCandidateWithSnapshot = (snapshot, staff, dateStr) => {
        const leaveDeficit = Math.max(0, requiredLeaves - getLeaveCountFromSnapshot(snapshot, staff.id));
        const workCount = getWorkCountFromSnapshot(snapshot, staff.id);
        const group = staff.group || '白班';
        const sameDayLeaveLoad = getGroupLeaveLoad(snapshot, dateStr, group);
        const recentLeavePressure = getRecentLeavePressure(snapshot, staff.id, dateStr, 4);
        const nearbyLeavePressure = getNearbyLeavePressure(snapshot, staff.id, dateStr, 2);
        const daysSinceLastLeave = getDaysSinceLastLeave(snapshot, staff.id, dateStr, 10);
        const consecutiveLeavePattern = getConsecutiveLeavePattern(snapshot, staff.id, dateStr);
        let score = 0;
        score += leaveDeficit * 120;
        score += workCount * 5;
        score += getRecentWorkPressure(snapshot, staff.id, dateStr, 3) * 18;
        score += Math.min(daysSinceLastLeave, 10) * 8;
        score -= sameDayLeaveLoad * 30;
        score -= recentLeavePressure * 18;
        score -= nearbyLeavePressure * 28;
        if (consecutiveLeavePattern.adjacentLeaveCount === 1) score += 22;
        if (consecutiveLeavePattern.adjacentLeaveCount >= 2) score -= 12;
        return score;
      };

      for (const day of targetDays) {
        for (const group of SHIFT_GROUPS) {
          if (restrictedGroup && restrictedGroup !== group) continue;

          const shiftCode = normalizedTargetShift && getShiftGroupByCode(normalizedTargetShift) === group
            ? normalizedTargetShift
            : DEFAULT_SHIFT_BY_GROUP[group];

          const demand = getDemandForGroup(day, group);
          const alreadyAssigned = getAssignedCountByGroup(mergedSchedule, day.date, group);
          const needed = Math.max(0, demand - alreadyAssigned);

          const groupStaffs = staffs.filter(s => (s.group || '白班') === group && targetStaffIds.has(s.id));
          const groupStaffIds = new Set(groupStaffs.map(s => s.id));

          // 逐格補主班別，補到需求就停
          for (let slot = 0; slot < needed; slot += 1) {
            const assignableCandidates = groupStaffs
              .filter(staff => !getScheduleCode(mergedSchedule, staff.id, day.date))
              .filter(staff => !getVisiblePreScheduleCode(staff.id, day.date))
              .map(staff => {
                const result = canAssignWithSnapshot(mergedSchedule, staff, day.date, shiftCode);
                const canKeepLeaveTarget = result.allowed ? canStillMeetRequiredLeavesIfAssignShift(mergedSchedule, staff.id) : false;
                return {
                  staff,
                  allowed: result.allowed && canKeepLeaveTarget,
                  score: result.allowed && canKeepLeaveTarget ? scoreCandidateWithSnapshot(mergedSchedule, staff, day.date, shiftCode) : -1
                };
              })
              .filter(item => item.allowed)
              .sort((a, b) => b.score - a.score);

            if (assignableCandidates.length === 0) {
              summary.skipped += 1;
              continue;
            }

            const picked = assignableCandidates[0];
            setScheduleCode(mergedSchedule, picked.staff.id, day.date, shiftCode, 'auto');
            summary.workFilled += 1;
          }

          // 需求已滿後，只替休假不足者補 off；其他空白保留
          const leaveCandidates = groupStaffs
            .filter(staff => !getScheduleCode(mergedSchedule, staff.id, day.date))
            .filter(staff => {
              const preScheduleCode = getVisiblePreScheduleCode(staff.id, day.date);
              return !preScheduleCode;
            })
            .filter(staff => getLeaveCountFromSnapshot(mergedSchedule, staff.id) < requiredLeaves)
            .map(staff => ({ staff, score: scoreLeaveCandidateWithSnapshot(mergedSchedule, staff, day.date) }))
            .sort((a, b) => b.score - a.score);

          if (leaveCandidates.length > 0) {
            const currentLeaveLoad = getGroupLeaveLoad(mergedSchedule, day.date, group);
            const maxLeaveForDay = Math.max(0, groupStaffIds.size - demand);
            if (currentLeaveLoad < maxLeaveForDay) {
              const bestLeaveCandidate = leaveCandidates[0];
              if (bestLeaveCandidate && canStillMeetRequiredLeavesAfterAssign(mergedSchedule, bestLeaveCandidate.staff.id)) {
                setScheduleCode(mergedSchedule, bestLeaveCandidate.staff.id, day.date, defaultAutoLeaveCode, 'auto');
                summary.leaveFilled += 1;
              }
            }
          }
        }
      }

      const touchedRuleFillCells = Array.from(touchedRuleFillCellMap.values());
      const ruleFillChangedEntries = buildEntriesFromSnapshotDiff(mergedSchedule, { onlyCells: touchedRuleFillCells });
      if (ruleFillChangedEntries.length > 0) {
        applyRuleFillEntries(ruleFillChangedEntries, {
          preserveSelection: true,
          selectionCells: ruleFillChangedEntries.map(({ staffId, dateStr }) => ({ staffId, dateStr })),
          activeCell: ruleFillChangedEntries.length > 0 ? { staffId: ruleFillChangedEntries[ruleFillChangedEntries.length - 1].staffId, dateStr: ruleFillChangedEntries[ruleFillChangedEntries.length - 1].dateStr } : null,
          clearAssist: false,
          resetBuffer: true
        });
      }
      saveToHistory(isPartial ? '規則指定補空' : '規則全月補空', mergedSchedule);
      const changedCount = ruleFillChangedEntries.length;
      setRuleFillFeedback(`✅ 補空完成：上班 ${summary.workFilled} 格、休假 ${summary.leaveFilled} 格、未補成功 ${summary.skipped} 格${changedCount > 0 ? `，實際寫入 ${changedCount} 格` : '，沒有可寫入的新格'}`);
    } catch (error) {
      console.error(error);
      setRuleFillFeedback("❌ 規則補空失敗，請檢查設定。");
    } finally {
      setIsRuleFillLoading(false);
    }
  
  };
};
