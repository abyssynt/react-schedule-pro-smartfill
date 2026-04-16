const formatWordDayCellValue = (value = "") => {
  const text = String(value || '').trim();
  if (/^(例|休)[1-4]$/.test(text)) {
    return `<span class="word-numbered-leave">${text}</span>`;
  }
  return text;
};

export const exportToExcelService = async (ctx) => {
  const {
    loadExcelJS, year, month, uiSettings, blendHexColors, colors, hexToExcelArgb, mergedLeaveCodes, customColumns, daysInMonth, requiredLeaves, getExportNumberedValue, customColumnValues, getExportCellPresentation, buildExportStaffStats, groupedStaffs, buildExportDailyStats, applyExcelFourWeekDivider, tableFontColor, setShowExportMenu
  } = ctx;
    const ExcelJS = await loadExcelJS();
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`${year}年${month}月班表`);

    const exportTheme = {
      pageBg: uiSettings?.pageBackgroundColor || '#f8fafc',
      tableFont: uiSettings?.tableFontColor || '#1f2937',
      shiftBg: uiSettings?.shiftColumnBgColor || '#ffffff',
      shiftFont: uiSettings?.shiftColumnFontColor || '#1e293b',
      nameBg: uiSettings?.nameDateColumnBgColor || '#ffffff',
      nameFont: uiSettings?.nameDateColumnFontColor || '#1e293b',
      weekdayHeadBg: blendHexColors(uiSettings?.nameDateColumnBgColor || '#ffffff', '#f1f5f9', 0.7),
      weekendHeadBg: colors.weekend || '#dcfce7',
      holidayHeadBg: colors.holiday || '#fca5a5',
      weekendCellBg: blendHexColors(colors.weekend || '#dcfce7', '#ffffff', 0.35),
      holidayCellBg: blendHexColors(colors.holiday || '#fca5a5', '#ffffff', 0.35),
      monthTitleBg: blendHexColors(uiSettings?.pageBackgroundColor || '#f8fafc', '#ffffff', 0.55),
      summaryBg: uiSettings?.groupSummaryRowBgColor || '#fef3c7',
      leaveRowBg: blendHexColors(uiSettings?.pageBackgroundColor || '#f8fafc', '#ffffff', 0.2)
    };

    const statHeaders = ['上班', '假日休', '總休', ...mergedLeaveCodes, ...(customColumns || [])];
    const totalColumns = 1 + daysInMonth.length + statHeaders.length;
    const lastDateColumn = daysInMonth.length + 1;

    const monthTitleRow = worksheet.addRow([]);
    monthTitleRow.height = 26;

    const titleStartCol = 2;
    const titleEndCol = Math.max(2, lastDateColumn - 2);
    const leaveStartCol = Math.max(titleEndCol + 1, lastDateColumn - 1);
    const leaveEndCol = lastDateColumn;

    if (titleEndCol >= titleStartCol) {
      worksheet.mergeCells(1, titleStartCol, 1, titleEndCol);
      const titleCell = monthTitleRow.getCell(titleStartCol);
      titleCell.value = `${month}月班表`;
      titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.monthTitleBg, '#EFF6FF') } };
      titleCell.font = { bold: true, size: 14, color: { argb: hexToExcelArgb(exportTheme.tableFont, '#1F2937') } };
    }

    if (leaveEndCol >= leaveStartCol) {
      if (leaveEndCol > leaveStartCol) worksheet.mergeCells(1, leaveStartCol, 1, leaveEndCol);
      const leaveCell = monthTitleRow.getCell(leaveStartCol);
      leaveCell.value = `應休${requiredLeaves}天`;
      leaveCell.alignment = { vertical: 'middle', horizontal: 'right' };
      leaveCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.monthTitleBg, '#EFF6FF') } };
      leaveCell.font = { bold: true, size: 11, color: { argb: hexToExcelArgb(exportTheme.tableFont, '#1F2937') } };
    }

    for (let col = 1; col <= totalColumns; col += 1) {
      const cell = monthTitleRow.getCell(col);
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
      if (col === 1 || col > lastDateColumn) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.pageBg, '#FFFFFF') } };
      }
    }

    const headerRow = ['姓名', ...daysInMonth.map(d => `${d.day}\n(${d.weekStr})`), ...statHeaders];
    const header = worksheet.addRow(headerRow);
    header.height = 30;

    header.eachCell((cell, colNumber) => {
      cell.font = { bold: true, size: 10, color: { argb: hexToExcelArgb(exportTheme.tableFont, '#1F2937') } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      const baseBorder = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
      cell.border = baseBorder;

      if (colNumber >= 2 && colNumber <= daysInMonth.length + 1) {
        const d = daysInMonth[colNumber - 2];
        cell.border = applyExcelFourWeekDivider(baseBorder, d.date);
        if (d.isHoliday) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.holidayHeadBg, '#FFCACA') } };
        else if (d.isWeekend) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.weekendHeadBg, '#DCFCE7') } };
        else cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.weekdayHeadBg, '#F1F5F9') } };
      } else if (colNumber === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.shiftBg, '#FFFFFF') } };
        cell.font = { ...(cell.font || {}), color: { argb: hexToExcelArgb(exportTheme.shiftFont, '#1E293B') } };
      } else {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.pageBg, '#F8FAFC') } };
      }
    });

    const makeBaseBorder = () => ({
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' }
    });

    const applyStandardCellStyle = (cell, colNumber, dateObj = null) => {
      const baseBorder = makeBaseBorder();
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.font = { color: { argb: hexToExcelArgb(colNumber === 1 ? exportTheme.nameFont : exportTheme.tableFont, '#1F2937') } };
      cell.border = baseBorder;
      if (dateObj) {
        cell.numFmt = '@';
        cell.border = applyExcelFourWeekDivider(baseBorder, dateObj.date);
        if (dateObj.isHoliday) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.holidayCellBg, '#FFE4E4') } };
        else if (dateObj.isWeekend) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.weekendCellBg, '#F0FDF4') } };
      }
    };

    const addStaffRow = (staff) => {
      const stats = buildExportStaffStats(staff.id);
      const rowData = [
        staff.name,
        ...daysInMonth.map((d) => getExportNumberedValue(staff.id, d.date) || ''),
        stats.work,
        stats.holidayLeave,
        stats.totalLeave,
        ...mergedLeaveCodes.map(l => stats.leaveDetails[l] || ''),
        ...(customColumns || []).map(col => customColumnValues?.[staff.id]?.[col] || '')
      ];
      const row = worksheet.addRow(rowData);
      row.eachCell((cell, colNumber) => {
        const dateObj = (colNumber >= 2 && colNumber <= daysInMonth.length + 1) ? daysInMonth[colNumber - 2] : null;
        applyStandardCellStyle(cell, colNumber, dateObj);
        if (dateObj) {
          const presentation = getExportCellPresentation(staff.id, dateObj);
          if (presentation.backgroundColor) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(presentation.backgroundColor, '#DBEAFE') } };
          }
          cell.font = { ...(cell.font || {}), color: { argb: hexToExcelArgb(presentation.textColor || exportTheme.tableFont, '#1F2937') } };
        }
      });
      return row;
    };

    const addSummaryRow = (summaryKey, includeRightStats = false) => {
      const rowData = [
        '',
        ...daysInMonth.map(d => buildExportDailyStats(d.date)[summaryKey] || ''),
        ...(includeRightStats ? Array(statHeaders.length).fill('') : Array(statHeaders.length).fill(''))
      ];
      const row = worksheet.addRow(rowData);
      row.eachCell((cell, colNumber) => {
        const baseBorder = makeBaseBorder();
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { bold: true, color: { argb: hexToExcelArgb(exportTheme.tableFont, '#1F2937') } };
        cell.border = baseBorder;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.summaryBg, '#FEF3C7') } };
        if (colNumber >= 2 && colNumber <= daysInMonth.length + 1) {
          const d = daysInMonth[colNumber - 2];
          cell.border = applyExcelFourWeekDivider(baseBorder, d.date);
        }
      });
      return row;
    };

    groupedStaffs.forEach(({ group, staffs: groupStaffList }) => {
      groupStaffList.forEach(addStaffRow);
      const summaryKey = group === '白班' ? 'D' : group === '小夜' ? 'E' : 'N';
      addSummaryRow(summaryKey);
    });

    const leaveRowData = [
      '',
      ...daysInMonth.map(d => buildExportDailyStats(d.date).totalLeave || ''),
      ...Array(statHeaders.length).fill('')
    ];
    const leaveRow = worksheet.addRow(leaveRowData);
    leaveRow.eachCell((cell, colNumber) => {
      const baseBorder = makeBaseBorder();
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.font = { bold: true, color: { argb: hexToExcelArgb(exportTheme.tableFont, '#1F2937') } };
      cell.border = baseBorder;
      if (colNumber >= 1 && colNumber <= totalColumns) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToExcelArgb(exportTheme.leaveRowBg, '#FFFFFF') } };
      }
      if (colNumber >= 2 && colNumber <= daysInMonth.length + 1) {
        const d = daysInMonth[colNumber - 2];
        cell.border = applyExcelFourWeekDivider(baseBorder, d.date);
      }
    });

    worksheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }];

    worksheet.getColumn(1).width = 15;
    for (let i = 2; i <= daysInMonth.length + 1; i += 1) worksheet.getColumn(i).width = 5;
    for (let i = daysInMonth.length + 2; i <= totalColumns; i += 1) worksheet.getColumn(i).width = 8;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `排班表_${year}年${month}月.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    setShowExportMenu(false);
  };

export const exportToWordService = (ctx) => {
  const {
    year, month, uiSettings, blendHexColors, colors, daysInMonth, requiredLeaves, schedulingRulesText, SHIFT_GROUPS, staffs, getExportCellPresentation, getExportNumberedValue, buildExportStaffStats, buildExportDailyStats, getWordCycleDividerStyle, tableFontColor, setShowExportMenu
  } = ctx;
    const statHeaders = ['上班', '假日休', '總休'];
    const exportTheme = {
      pageBg: uiSettings?.pageBackgroundColor || '#f8fafc',
      tableFont: uiSettings?.tableFontColor || '#1f2937',
      shiftBg: uiSettings?.shiftColumnBgColor || '#ffffff',
      shiftFont: uiSettings?.shiftColumnFontColor || '#1e293b',
      nameBg: uiSettings?.nameDateColumnBgColor || '#ffffff',
      nameFont: uiSettings?.nameDateColumnFontColor || '#1e293b',
      weekdayHeadBg: blendHexColors(uiSettings?.nameDateColumnBgColor || '#ffffff', '#f1f5f9', 0.7),
      weekendHeadBg: colors.weekend || '#dcfce7',
      holidayHeadBg: colors.holiday || '#fca5a5',
      weekendCellBg: blendHexColors(colors.weekend || '#dcfce7', '#ffffff', 0.35),
      holidayCellBg: blendHexColors(colors.holiday || '#fca5a5', '#ffffff', 0.35),
      statWorkBg: blendHexColors(uiSettings?.pageBackgroundColor || '#f8fafc', '#93c5fd', 0.45),
      statHolidayBg: blendHexColors(uiSettings?.pageBackgroundColor || '#f8fafc', colors.weekend || '#dcfce7', 0.5),
      statTotalBg: blendHexColors(uiSettings?.pageBackgroundColor || '#f8fafc', colors.holiday || '#fca5a5', 0.45)
    };

    const leaveTitleColSpan = Math.min(3, Math.max(1, daysInMonth.length));
    const titleColSpan = Math.max(1, daysInMonth.length - leaveTitleColSpan);
    const statColSpan = statHeaders.length;
    const totalColumns = 1 + daysInMonth.length + statHeaders.length;
    const wordPageWidthPt = 841.9;
    const wordMarginPt = 18;
    const wordUsableWidthPt = wordPageWidthPt - (wordMarginPt * 2);
    const wordNameColWidthPt = 54;
    const wordStatColWidthPt = 30;
    const rawWordDayColWidthPt = (wordUsableWidthPt - wordNameColWidthPt - (statHeaders.length * wordStatColWidthPt)) / Math.max(daysInMonth.length, 1);
    const wordDayColWidthPt = Math.max(18, Math.min(24, rawWordDayColWidthPt));
    const wordTableWidthPt = wordNameColWidthPt + (daysInMonth.length * wordDayColWidthPt) + (statHeaders.length * wordStatColWidthPt);
    const schedulingRuleLines = String(schedulingRulesText || '')
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean);
    const schedulingRulesHtml = schedulingRuleLines.length > 0
      ? `排班規則：<br/>${schedulingRuleLines.map((line, index) => `${index + 1}. ${line}`).join('<br/>')}`
      : '排班規則：';

    const webSummaryRowBg = uiSettings?.groupSummaryRowBgColor || '#fef3c7';
    const summaryRows = [
      { group: '白班', key: 'D', label: '白班上班', bg: webSummaryRowBg },
      { group: '小夜', key: 'E', label: '小夜上班', bg: webSummaryRowBg },
      { group: '大夜', key: 'N', label: '大夜上班', bg: webSummaryRowBg }
    ];

    const groupedExportRowsHtml = SHIFT_GROUPS.map((group) => {
      const groupStaffsForExport = staffs.filter((staff) => (staff.group || '白班') === group);

      const staffRowsHtml = groupStaffsForExport.map((staff) => {
        const stats = buildExportStaffStats(staff.id);
        return `
                <tr>
                  <td class="name-col" style="background:${exportTheme.nameBg}; color:${exportTheme.nameFont}; mso-pattern:auto none;">${staff.name}</td>
                  ${daysInMonth.map(d => {
                    const presentation = getExportCellPresentation(staff.id, d);
                    const value = getExportNumberedValue(staff.id, d.date) || '';
                    const cellClass = d.isHoliday ? 'holiday-cell' : (d.isWeekend ? 'weekend-cell' : '');
                    const cellBg = presentation.hasPreSchedule
                      ? presentation.backgroundColor
                      : (d.isHoliday ? exportTheme.holidayCellBg : (d.isWeekend ? exportTheme.weekendCellBg : exportTheme.pageBg));
                    return `<td class="day-col ${cellClass}" style="background:${cellBg}; color:${presentation.textColor || exportTheme.tableFont}; mso-pattern:auto none;${getWordCycleDividerStyle(d.date)}">${formatWordDayCellValue(value)}</td>`;
                  }).join('')}
                  <td class="stat-col stat-work-cell" style="background:${exportTheme.statWorkBg}; color:${exportTheme.tableFont}; mso-pattern:auto none;">${stats.work || ''}</td>
                  <td class="stat-col stat-holiday-cell" style="background:${exportTheme.statHolidayBg}; color:${exportTheme.tableFont}; mso-pattern:auto none;">${stats.holidayLeave || ''}</td>
                  <td class="stat-col stat-total-cell" style="background:${exportTheme.statTotalBg}; color:${exportTheme.tableFont}; mso-pattern:auto none;">${stats.totalLeave || ''}</td>
                </tr>`;
      }).join('');

      const summaryConfig = summaryRows.find((item) => item.group === group);
      const summaryRowHtml = summaryConfig ? `
                <tr>
                  <td class="name-col summary-label-cell" style="background:${summaryConfig.bg}; color:${exportTheme.nameFont}; mso-pattern:auto none;"></td>
                  ${daysInMonth.map(d => {
                    const count = buildExportDailyStats(d.date)[summaryConfig.key];
                    return `<td class="day-col summary-value-cell" style="background:${summaryConfig.bg}; color:${exportTheme.tableFont}; mso-pattern:auto none;${getWordCycleDividerStyle(d.date)}">${count || ''}</td>`;
                  }).join('')}
                  <td class="stat-col summary-value-cell" style="background:${summaryConfig.bg}; mso-pattern:auto none;"></td>
                  <td class="stat-col summary-value-cell" style="background:${summaryConfig.bg}; mso-pattern:auto none;"></td>
                  <td class="stat-col summary-value-cell" style="background:${summaryConfig.bg}; mso-pattern:auto none;"></td>
                </tr>` : '';

      return `${staffRowsHtml}${summaryRowHtml}`;
    }).join('');

    const html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office"
          xmlns:w="urn:schemas-microsoft-com:office:word"
          xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8">
        <meta name="ProgId" content="Word.Document">
        <meta name="Generator" content="Microsoft Word 15">
        <meta name="Originator" content="Microsoft Word 15">
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>90</w:Zoom>
            <w:DoNotOptimizeForBrowser/>
          </w:WordDocument>
        </xml>
        <style>
          @page WordSection1 {
            size: 841.9pt 595.3pt;
            mso-page-orientation: landscape;
            margin: ${wordMarginPt}pt ${wordMarginPt}pt ${wordMarginPt}pt ${wordMarginPt}pt;
          }
          div.WordSection1 { page: WordSection1; }
          body {
            font-family: "Microsoft JhengHei", Arial, sans-serif;
            margin: 0;
            padding: 0;
            color: ${exportTheme.tableFont};
            background: ${exportTheme.pageBg};
          }
          table {
            border-collapse: collapse;
            table-layout: fixed;
            width: ${wordTableWidthPt}pt;
            max-width: ${wordTableWidthPt}pt;
            margin: 0 auto;
            font-size: 9pt;
          }
          th, td {
            border: 1px solid #000;
            padding: 2px 3px;
            text-align: center;
            vertical-align: middle;
            word-break: break-all;
          }
          .month-row td {
            height: 24pt;
            font-weight: 700;
            background: ${exportTheme.pageBg};
            border-top: 0;
            border-left: 0;
            border-right: 0;
            border-bottom: 1px solid #000;
          }
          .month-name-spacer,
          .month-stat-spacer {
            color: transparent;
          }
          .month-title-zone,
          .month-leave-zone {
            height: 24pt;
            padding: 0 6pt;
          }
          .month-title-wrap,
          .month-leave-wrap {
            position: relative;
            width: 100%;
            height: 24pt;
          }
          .month-title {
            display: block;
            width: 100%;
            text-align: center;
            font-size: 14pt;
            font-weight: 700;
            line-height: 24pt;
            white-space: nowrap;
          }
          .month-leave-inline {
            display: block;
            width: 100%;
            text-align: right;
            font-size: 10.5pt;
            font-weight: 700;
            line-height: 24pt;
            white-space: nowrap;
          }
          .name-col {
            width: ${wordNameColWidthPt}pt;
            min-width: ${wordNameColWidthPt}pt;
            font-weight: 700;
          }
          .day-col {
            width: ${wordDayColWidthPt}pt;
            min-width: ${wordDayColWidthPt}pt;
            white-space: nowrap;
          }
          .word-numbered-leave {
            display: inline-block;
            white-space: nowrap;
            word-break: keep-all;
            font-size: 8pt;
            line-height: 1;
            letter-spacing: -0.1pt;
          }
          .stat-col {
            width: ${wordStatColWidthPt}pt;
            min-width: ${wordStatColWidthPt}pt;
          }
          .header-cell {
            font-weight: 700;
            line-height: 1.1;
          }
          .weekday-head { background-color: ${exportTheme.weekdayHeadBg}; }
          .holiday-head { background-color: ${exportTheme.holidayHeadBg}; }
          .weekend-head { background-color: ${exportTheme.weekendHeadBg}; }
          .holiday-cell { background-color: ${exportTheme.holidayCellBg}; }
          .weekend-cell { background-color: ${exportTheme.weekendCellBg}; }
          .stat-work-head { background-color: ${exportTheme.statWorkBg}; color: ${exportTheme.tableFont}; }
          .stat-holiday-head { background-color: ${exportTheme.statHolidayBg}; color: ${exportTheme.tableFont}; }
          .stat-total-head { background-color: ${exportTheme.statTotalBg}; color: ${exportTheme.tableFont}; }
          .stat-work-cell { background-color: ${exportTheme.statWorkBg}; }
          .stat-holiday-cell { background-color: ${exportTheme.statHolidayBg}; }
          .stat-total-cell { background-color: ${exportTheme.statTotalBg}; }
          .summary-label-cell, .summary-value-cell {
            font-weight: 700;
          }
          .rules-row td {
            padding: 8pt 10pt;
            text-align: left;
            vertical-align: top;
            line-height: 1.7;
            font-size: 10pt;
            background: ${exportTheme.pageBg};
          }
        </style>
      </head>
      <body>
        <div class="WordSection1">
          <table>
            <thead>
              <tr class="month-row">
                <td class="name-col month-name-spacer"></td>
                <td class="month-title-zone" colspan="${titleColSpan}">
                  <div class="month-title-wrap">
                    <span class="month-title">${month}月班表</span>
                  </div>
                </td>
                <td class="month-leave-zone" colspan="${leaveTitleColSpan}">
                  <div class="month-leave-wrap">
                    <span class="month-leave-inline">應休${requiredLeaves}天</span>
                  </div>
                </td>
                <td class="month-stat-spacer" colspan="${statColSpan}"></td>
              </tr>
              <tr>
                <th class="name-col header-cell" style="background:${exportTheme.nameBg}; color:${exportTheme.nameFont}; mso-pattern:auto none;">姓名</th>
                ${daysInMonth.map(d => {
                  const headClass = d.isHoliday ? 'holiday-head' : (d.isWeekend ? 'weekend-head' : 'weekday-head');
                  const headBg = d.isHoliday ? exportTheme.holidayHeadBg : (d.isWeekend ? exportTheme.weekendHeadBg : exportTheme.weekdayHeadBg);
                  return `<th class="day-col header-cell ${headClass}" style="background:${headBg}; mso-pattern:auto none;${getWordCycleDividerStyle(d.date)}">${d.day}<br/>(${d.weekStr})</th>`;
                }).join('')}
                <th class="stat-col header-cell stat-work-head">上班</th>
                <th class="stat-col header-cell stat-holiday-head">假日休</th>
                <th class="stat-col header-cell stat-total-head">總休</th>
              </tr>
            </thead>
            <tbody>
              ${groupedExportRowsHtml}
            </tbody>
            <tfoot>
              <tr class="rules-row">
                <td colspan="${totalColumns}">${schedulingRulesHtml}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </body>
    </html>`;

    const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `列印班表_${year}年${month}月.doc`;
    a.click();
    URL.revokeObjectURL(url);
    setShowExportMenu(false);
  };
