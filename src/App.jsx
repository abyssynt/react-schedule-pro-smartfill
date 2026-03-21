import React, { useState, useMemo, useEffect } from 'react';
import { 
  Plus, Minus, AlertCircle, Settings, Sparkles, Loader2, 
  ArrowUp, ArrowDown, Save, History as Clock, Download, 
  FileSpreadsheet, FileText, X, Check, Calendar 
} from 'lucide-react';

// ==========================================
// 1. 系統代碼字典
// ==========================================
const DICT = {
  SHIFTS: ['D', 'E', 'N', '8-8'],
  LEAVES: ['off', '例', '休', '特', '補', '國', '喪', '婚', '產', '病', '事', '陪產']
};

const SHIFT_GROUPS = ['白班', '小夜', '大夜'];

const HOLIDAY_MAP = {
  2024: ['2024-01-01', '2024-02-08', '2024-02-09', '2024-02-10', '2024-02-11', '2024-02-12', '2024-02-13', '2024-02-14', '2024-02-28', '2024-04-04', '2024-04-05', '2024-06-10', '2024-09-17', '2024-10-10'],
  2025: ['2025-01-01', '2025-01-27', '2025-01-28', '2025-01-29', '2025-01-30', '2025-01-31', '2025-02-28', '2025-04-03', '2025-04-04', '2025-05-31', '2025-10-06', '2025-10-10'],
  2026: ['2026-01-01', '2026-02-16', '2026-02-17', '2026-02-18', '2026-02-19', '2026-02-20', '2026-02-28', '2026-04-04', '2026-04-05', '2026-06-19', '2026-09-25', '2026-10-10']
};

const apiKey = ""; 
const STORAGE_KEY = 'schedule_app_history';

// 外部套件載入：ExcelJS 用於高品質 Excel 樣式輸出
const loadExcelJS = () => {
  return new Promise((resolve) => {
    if (window.ExcelJS) return resolve(window.ExcelJS);
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js";
    script.onload = () => resolve(window.ExcelJS);
    document.head.appendChild(script);
  });
};

const normalizeStaffGroup = (staffList = []) => {
  if (!Array.isArray(staffList) || staffList.length === 0) return [];

  const fallbackGroups = ['白班', '白班', '白班', '白班', '白班', '小夜', '小夜', '小夜', '小夜', '小夜', '大夜', '大夜', '大夜', '大夜', '大夜'];

  return staffList.map((staff, index) => ({
    ...staff,
    group: SHIFT_GROUPS.includes(staff.group) ? staff.group : (fallbackGroups[index] || '白班')
  }));
};

export default function App() {
  // ==========================================
  // 2. 核心 State 定義
  // ==========================================
  const [year, setYear] = useState(2025);
  const [month, setMonth] = useState(3);
  const [holidaysText, setHolidaysText] = useState(''); 
  const [isAiLoading, setIsAiLoading] = useState(false);
  
  const [colors, setColors] = useState({ weekend: '#dcfce7', holiday: '#fca5a5' });

  const [staffs, setStaffs] = useState(normalizeStaffGroup([
    { id: 's1', name: '新成員' }, { id: 's2', name: '新成員' }, { id: 's3', name: '新成員' }, { id: 's4', name: '新成員' }, { id: 's5', name: '新成員' },
    { id: 's6', name: '新成員' }, { id: 's7', name: '新成員' }, { id: 's8', name: '新成員' }, { id: 's9', name: '新成員' }, { id: 's10', name: '新成員' },
    { id: 's11', name: '新成員' }, { id: 's12', name: '新成員' }, { id: 's13', name: '新成員' }, { id: 's14', name: '新成員' }, { id: 's15', name: '新成員' }
  ]));
  const [schedule, setSchedule] = useState({ 's1': {}, 's2': {}, 's3': {}, 's4': {}, 's5': {}, 's6': {}, 's7': {}, 's8': {}, 's9': {}, 's10': {}, 's11': {}, 's12': {}, 's13': {}, 's14': {}, 's15': {} });

  const [showSettings, setShowSettings] = useState(false);
  const [showAiControl, setShowAiControl] = useState(false); 
  const [aiFeedback, setAiFeedback] = useState("");

  const [showHistoryModal, setShowHistoryModal] = useState(false);
  const [historyList, setHistoryList] = useState([]);
  const [showDraftPrompt, setShowDraftPrompt] = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);

  // AI 指定排班設定
  const [aiConfig, setAiConfig] = useState({
    selectedStaffs: [], 
    dateRange: { start: 1, end: 31 },
    targetShift: ''
  });

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
          setShowDraftPrompt(true);
        }
      } catch (e) { console.error("歷史紀錄解析失敗"); }
    }
  }, []);

  useEffect(() => {
    const autoHolidays = HOLIDAY_MAP[year] || null;
    if (autoHolidays) {
      setHolidaysText(autoHolidays.join(', '));
    }
  }, [year]);

  const holidays = useMemo(() => holidaysText.split(',').map(s => s.trim()).filter(Boolean), [holidaysText]);

  const daysInMonth = useMemo(() => {
    const days = [];
    const daysCount = new Date(year, month, 0).getDate();
    const weekNames = ['日', '一', '二', '三', '四', '五', '六'];
    for (let i = 1; i <= daysCount; i++) {
      const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
      const weekNum = new Date(year, month - 1, i).getDay();
      days.push({
        day: i, date: dateStr, weekStr: weekNames[weekNum],
        isWeekend: weekNum === 0 || weekNum === 6,
        isHoliday: holidays.includes(dateStr)
      });
    }
    return days;
  }, [year, month, holidays]);

  const requiredLeaves = useMemo(() => daysInMonth.filter(d => d.isWeekend || d.isHoliday).length, [daysInMonth]);

  // ==========================================
  // 4. Excel 匯出 (ExcelJS 實現)
  // ==========================================
  const exportToExcel = async () => {
    setAiFeedback("📊 正在產生高品質 Excel 報表...");
    const ExcelJS = await loadExcelJS();
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`${year}年${month}月班表`);

    const headerRow = ['班別', '日期/姓名', ...daysInMonth.map(d => `${d.day}\n(${d.weekStr})`), '上班', '假日休', '總休', ...DICT.LEAVES];
    const header = worksheet.addRow(headerRow);
    header.height = 30;

    header.eachCell((cell, colNumber) => {
      cell.font = { bold: true, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
      if (colNumber > 1 && colNumber <= daysInMonth.length + 1) {
        const d = daysInMonth[colNumber - 2];
        if (d.isHoliday) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCACA' } };
        else if (d.isWeekend) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
      }
    });

    staffs.forEach(staff => {
      const stats = getStaffStats(staff.id);
      const rowData = [
        staff.group,
        staff.name,
        ...daysInMonth.map(d => {
          const cellData = schedule[staff.id]?.[d.date];
          return typeof cellData === 'object' ? (cellData?.value || '') : (cellData || '');
        }),
        stats.work, stats.holidayLeave, stats.totalLeave,
        ...DICT.LEAVES.map(l => stats.leaveDetails[l] || '')
      ];
      const row = worksheet.addRow(rowData);
      row.eachCell((cell, colNumber) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        if (colNumber > 1 && colNumber <= daysInMonth.length + 1) {
          cell.numFmt = '@';
          const d = daysInMonth[colNumber - 2];
          if (d.isHoliday) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE4E4' } };
          else if (d.isWeekend) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FDF4' } };
        }
      });
    });

    ['D', 'E', 'N', 'totalLeave'].forEach(rowKey => {
      const label = rowKey === 'totalLeave' ? '當日休假' : `${rowKey} 班人數`;
      const rowData = [label, ...daysInMonth.map(d => getDailyStats(d.date)[rowKey] || '')];
      const row = worksheet.addRow(rowData);
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
      });
    });

    worksheet.getColumn(1).width = 10;
    worksheet.getColumn(2).width = 15;
    for(let i=3; i <= daysInMonth.length + 2; i++) worksheet.getColumn(i).width = 5;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `排班表_${year}年${month}月.xlsx`;
    a.click();
    setShowExportMenu(false);
    setAiFeedback("✅ Excel 導出成功！");
  };

  const exportToWord = () => {
    let html = `
    <html><head><meta charset="utf-8"><style>@page { size: landscape; margin: 1cm; } body { font-family: sans-serif; } table { border-collapse: collapse; width: 100%; font-size: 9pt; } th, td { border: 1px solid #000; padding: 4px; text-align: center; } .holiday { background-color: ${colors.holiday}; } .weekend { background-color: ${colors.weekend}; }</style></head>
    <body><h2 style="text-align:center;">${year}年${month}月 班表</h2><table><thead><tr><th>班別</th><th>姓名</th>${daysInMonth.map(d => `<th class="${d.isHoliday ? 'holiday' : (d.isWeekend ? 'weekend' : '')}">${d.day}<br/>(${d.weekStr})</th>`).join('')}<th>上班</th><th>總休</th></tr></thead>
    <tbody>${staffs.map(staff => {
      const stats = getStaffStats(staff.id);
      return `<tr><td>${staff.group}</td><td>${staff.name}</td>${daysInMonth.map(d => {
        const cellData = schedule[staff.id]?.[d.date];
        return `<td>${typeof cellData === 'object' ? (cellData?.value || '') : (cellData || '')}</td>`;
      }).join('')}<td>${stats.work}</td><td>${stats.totalLeave}</td></tr>`;
    }).join('')}</tbody></table></body></html>`;
    const blob = new Blob([html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `列印班表_${year}年${month}月.doc`;
    a.click();
    setShowExportMenu(false);
  };

  // ==========================================
  // 5. AI 指定排班功能
  // ==========================================
  const handleAiAutoSchedule = async (isPartial = false) => {
    setIsAiLoading(true);
    setAiFeedback(isPartial ? "✨ AI 正在根據指定區域填寫班表..." : "✨ AI 正在規劃全月空缺...");
    
    const systemPrompt = `你是一個專業護理排班專家。請為未排班的日期填寫代碼。可用班別: ${DICT.SHIFTS.join(', ')}，可用休假: ${DICT.LEAVES.join(', ')}。${isPartial && aiConfig.targetShift ? `優先填寫指定班別: ${aiConfig.targetShift}` : ''} 格式要求: {"schedule": {"staffId": {"YYYY-MM-DD": "CODE"}}}`;
    
    const currentScheduleForAi = {};
    staffs.forEach(s => {
      currentScheduleForAi[s.id] = {};
      Object.keys(schedule[s.id] || {}).forEach(date => {
        const cell = schedule[s.id][date];
        if(cell && (cell.value || typeof cell === 'string')) currentScheduleForAi[s.id][date] = cell.value || cell;
      });
    });

    const userPrompt = `年: ${year}, 月: ${month}, 人員對象: ${JSON.stringify(isPartial ? aiConfig.selectedStaffs : staffs.map(s=>s.id))}, 日期區間: ${isPartial ? `${aiConfig.dateRange.start}號到${aiConfig.dateRange.end}號` : '全月'}, 既有排班數據: ${JSON.stringify(currentScheduleForAi)}。注意：絕對不可覆蓋已有的手動排班內容。`;

    try {
      const result = await callGemini(userPrompt, systemPrompt);
      if (result.schedule) {
        const mergedSchedule = { ...schedule };
        Object.keys(result.schedule).forEach(staffId => {
          if (isPartial && !aiConfig.selectedStaffs.includes(staffId)) return;
          if (!mergedSchedule[staffId]) mergedSchedule[staffId] = {};
          Object.keys(result.schedule[staffId]).forEach(dateStr => {
            const dayNum = parseInt(dateStr.split('-')[2]);
            if (isPartial && (dayNum < aiConfig.dateRange.start || dayNum > aiConfig.dateRange.end)) return;
            const aiCode = result.schedule[staffId][dateStr];
            const existingCell = mergedSchedule[staffId][dateStr];
            if (existingCell && existingCell.source === 'manual') return; 

            if (aiCode) {
              let finalValue = aiCode;
              if(isPartial && aiConfig.targetShift && !DICT.LEAVES.includes(aiCode)) finalValue = aiConfig.targetShift;
              mergedSchedule[staffId][dateStr] = { value: finalValue, source: 'ai' };
            }
          });
        });
        setSchedule(mergedSchedule);
        saveToHistory(isPartial ? 'AI區域排班' : 'AI全月排班', mergedSchedule);
        setAiFeedback(`✅ ${isPartial ? '區域' : '全月'}補空完成！`);
      }
    } catch (error) { setAiFeedback("❌ AI 排班失敗，請檢查網路。"); }
    finally { setIsAiLoading(false); }
  };

  const callGemini = async (prompt, systemInstruction = "") => {
    let delay = 1000;
    for (let i = 0; i < 5; i++) {
      try {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=${apiKey}`, {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            systemInstruction: systemInstruction ? { parts: [{ text: systemInstruction }] } : undefined,
            generationConfig: { responseMimeType: "application/json" }
          })
        });
        if (!response.ok) throw new Error('API Error');
        const data = await response.json();
        return JSON.parse(data.candidates[0].content.parts[0].text);
      } catch (err) {
        if (i === 4) throw err;
        await new Promise(resolve => setTimeout(resolve, delay));
        delay *= 2;
      }
    }
  };

  // ==========================================
  // 6. 輔助統計與操作
  // ==========================================
  const getStaffStats = (staffId) => {
    let stats = { work: 0, holidayLeave: 0, totalLeave: 0, leaveDetails: Object.fromEntries(DICT.LEAVES.map(l => [l, 0])) };
    const mySchedule = schedule[staffId] || {};
    daysInMonth.forEach(d => {
      const cellData = mySchedule[d.date];
      let code = typeof cellData === 'object' && cellData !== null ? cellData.value : cellData;
      if (!code) return;
      if (DICT.SHIFTS.includes(code)) stats.work += 1;
      if (DICT.LEAVES.includes(code)) {
        stats.totalLeave += 1;
        if(stats.leaveDetails[code] !== undefined) stats.leaveDetails[code] += 1;
        if (d.isWeekend || d.isHoliday) stats.holidayLeave += 1;
      }
    });
    return stats;
  };

  const getDailyStats = (dateStr) => {
    let stats = { 'D': 0, 'E': 0, 'N': 0, '8-8': 0, totalLeave: 0 };
    staffs.forEach(staff => {
      const cellData = schedule[staff.id]?.[dateStr];
      let code = typeof cellData === 'object' && cellData !== null ? cellData.value : cellData;
      if (!code) return;
      if (DICT.SHIFTS.includes(code)) stats[code] += 1;
      else if (DICT.LEAVES.includes(code)) stats.totalLeave += 1;
    });
    return stats;
  };

  const saveToHistory = (label, currentSchedule = schedule) => {
    const newRecord = {
      id: Date.now(), label, timestamp: new Date().toLocaleString(),
      state: { year, month, holidaysText, staffs, schedule: currentSchedule, colors }
    };
    setHistoryList(prev => {
      const updated = [newRecord, ...prev].slice(0, 10);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
      return updated;
    });
  };

  const loadHistory = (record) => {
    const { state } = record;
    setYear(state.year); setMonth(state.month); setHolidaysText(state.holidaysText);
    setStaffs(normalizeStaffGroup(state.staffs)); setSchedule(state.schedule);
    if(state.colors) setColors(state.colors);
    setShowHistoryModal(false); setShowDraftPrompt(false);
  };

  const clearHistory = () => {
    if(window.confirm("確定要清空所有歷史紀錄嗎？")) {
      localStorage.removeItem(STORAGE_KEY);
      setHistoryList([]);
    }
  };

  const handleCellChange = (staffId, dateStr, value) => {
    setSchedule(prev => ({
      ...prev, [staffId]: { ...prev[staffId], [dateStr]: value ? { value, source: 'manual' } : null }
    }));
  };

  const addStaff = (group = '白班') => {
    const newId = 's' + Date.now();
    setStaffs(prev => [...prev, { id: newId, name: '新成員', group }]);
    setSchedule(prev => ({ ...prev, [newId]: {} }));
  };

  const removeStaff = (id) => {
    if(window.confirm("確定刪除此人員？")) {
      setStaffs(staffs.filter(s => s.id !== id));
      const newSchedule = { ...schedule }; delete newSchedule[id]; setSchedule(newSchedule);
    }
  };

  const moveStaffInGroup = (staffId, direction) => {
    const newStaffs = [...staffs];
    const currentIndex = newStaffs.findIndex(s => s.id === staffId);
    if (currentIndex === -1) return;

    const currentGroup = newStaffs[currentIndex].group;
    const groupIndexes = newStaffs
      .map((staff, index) => ({ staff, index }))
      .filter(item => item.staff.group === currentGroup)
      .map(item => item.index);

    const currentGroupPos = groupIndexes.indexOf(currentIndex);
    const targetGroupPos = direction === 'up' ? currentGroupPos - 1 : currentGroupPos + 1;
    if (targetGroupPos < 0 || targetGroupPos >= groupIndexes.length) return;

    const targetIndex = groupIndexes[targetGroupPos];
    [newStaffs[currentIndex], newStaffs[targetIndex]] = [newStaffs[targetIndex], newStaffs[currentIndex]];
    setStaffs(newStaffs);
  };

  const groupedStaffs = useMemo(() => {
    return SHIFT_GROUPS.map(group => ({
      group,
      staffs: staffs.filter(staff => (staff.group || '白班') === group)
    }));
  }, [staffs]);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 p-4 font-sans overflow-x-hidden relative">
      <style>{`
        @keyframes pulse-once { 0% { transform: translateY(-10px); opacity: 0; } 100% { transform: translateY(0); opacity: 1; } }
        @keyframes fade-in-down { 0% { opacity: 0; transform: translateY(-5px); } 100% { opacity: 1; transform: translateY(0); } }
        .animate-pulse-once { animation: pulse-once 0.5s ease-out forwards; }
        .animate-fade-in-down { animation: fade-in-down 0.3s ease-out forwards; }
      `}</style>
      
      {showDraftPrompt && (
        <div className="max-w-[95vw] mx-auto mb-4 bg-amber-50 border border-amber-200 text-amber-800 p-4 rounded-xl flex items-center justify-between shadow-sm animate-fade-in-down">
          <div className="flex items-center gap-2"><Clock size={18} className="text-amber-600"/><span className="text-sm font-bold">偵測到先前暫存紀錄。</span></div>
          <div className="flex gap-2">
            <button onClick={() => loadHistory(historyList[0])} className="text-sm bg-amber-600 text-white px-3 py-1.5 rounded-lg hover:bg-amber-700 transition font-bold">載入最新</button>
            <button onClick={() => setShowDraftPrompt(false)} className="text-sm text-amber-700 hover:bg-amber-100 px-3 py-1.5 rounded-lg transition">忽略</button>
          </div>
        </div>
      )}

      <div className="max-w-[95vw] mx-auto mb-6">
        <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-200 flex flex-col xl:flex-row xl:items-center justify-between gap-4">
          <div>
            <h1 className="text-2xl font-black text-slate-800 flex items-center gap-2">智能排班系統 PRO｜智慧排班開發版 <span className="text-blue-500 text-sm font-normal px-2 py-1 bg-blue-50 rounded-lg border border-blue-100">PRO v1.6.0</span></h1>
            <p className="text-slate-500 text-xs mt-1 italic">開發版開發使用</p>
          </div>
          
          <div className="flex flex-wrap items-center gap-2">
            <button onClick={() => saveToHistory('手動暫存')} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-2 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm"><Save size={16} /> 暫存</button>
            <button onClick={() => setShowHistoryModal(true)} className="flex items-center gap-1.5 bg-slate-100 text-slate-700 border border-slate-300 px-3 py-2 rounded-xl font-bold hover:bg-slate-200 transition-all text-sm"><Clock size={16} /> 歷史</button>

            <div className="relative">
              <button onClick={() => setShowExportMenu(!showExportMenu)} className="flex items-center gap-1.5 bg-slate-800 text-white px-3 py-2 rounded-xl font-bold hover:bg-slate-900 transition-all text-sm"><Download size={16} /> 匯出</button>
              {showExportMenu && (
                <div className="absolute right-0 mt-2 w-48 bg-white border border-slate-200 rounded-xl shadow-xl z-50 overflow-hidden animate-fade-in-down">
                  <button onClick={exportToExcel} className="w-full text-left px-4 py-3 text-sm font-bold text-slate-700 hover:bg-green-50 hover:text-green-700 flex items-center gap-2 transition-colors border-b"><FileSpreadsheet size={16}/> 高品質 Excel</button>
                  <button onClick={exportToWord} className="w-full text-left px-4 py-3 text-sm font-bold text-slate-700 hover:bg-blue-50 hover:text-blue-700 flex items-center gap-2 transition-colors"><FileText size={16}/> 橫向 Word (列印)</button>
                </div>
              )}
            </div>

            <div className="w-px h-8 bg-slate-200 mx-2 hidden sm:block"></div>

            <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200">
                <button onClick={() => handleAiAutoSchedule(false)} disabled={isAiLoading} className="flex items-center gap-2 bg-white text-blue-600 px-3 py-2 rounded-lg font-bold hover:bg-blue-50 transition-all disabled:opacity-50 text-xs">
                  {isAiLoading ? <Loader2 className="animate-spin" size={14}/> : <Sparkles size={14} />} 全月補空
                </button>
                <button onClick={() => setShowAiControl(!showAiControl)} className={`flex items-center gap-2 px-3 py-2 rounded-lg font-bold transition-all text-xs ${showAiControl ? 'bg-blue-600 text-white shadow-inner' : 'text-slate-600 hover:bg-slate-200'}`}>
                    <Calendar size={14} /> 指定排班
                </button>
            </div>
          </div>
        </div>
        {aiFeedback && <div className="mt-4 bg-indigo-50 border border-indigo-100 p-4 rounded-xl text-indigo-900 text-sm animate-pulse-once flex items-center gap-2"><Check size={16} className="text-green-600"/>{aiFeedback}</div>}
      </div>

      {showAiControl && (
        <div className="max-w-[95vw] mx-auto mb-6 bg-blue-50 border border-blue-200 p-6 rounded-2xl shadow-sm animate-fade-in-down">
            <h3 className="font-black text-blue-900 mb-4 flex items-center gap-2"><Sparkles size={18}/> 指定區域排班設定</h3>
            <div className="grid lg:grid-cols-4 gap-6">
                <div>
                    <label className="block text-xs font-bold text-blue-700 mb-2 uppercase">1. 選擇人員</label>
                    <div className="flex flex-wrap gap-2">
                        {staffs.map(s => (
                            <button key={s.id} 
                                onClick={() => {
                                    const next = aiConfig.selectedStaffs.includes(s.id) ? aiConfig.selectedStaffs.filter(id => id !== s.id) : [...aiConfig.selectedStaffs, s.id];
                                    setAiConfig({...aiConfig, selectedStaffs: next});
                                }}
                                className={`px-3 py-1.5 rounded-lg text-xs font-bold border transition-all ${aiConfig.selectedStaffs.includes(s.id) ? 'bg-blue-600 border-blue-600 text-white shadow-md' : 'bg-white border-blue-200 text-blue-600 hover:bg-blue-100'}`}
                            > {s.name}（{s.group || '白班'}） </button>
                        ))}
                    </div>
                </div>
                <div>
                    <label className="block text-xs font-bold text-blue-700 mb-2 uppercase">2. 日期範圍 ({aiConfig.dateRange.start} ~ {aiConfig.dateRange.end} 號)</label>
                    <div className="flex items-center gap-2">
                        <input type="number" min="1" max="31" value={aiConfig.dateRange.start} onChange={(e)=>setAiConfig({...aiConfig, dateRange: {...aiConfig.dateRange, start: parseInt(e.target.value) || 1}})} className="w-full border-blue-200 border p-2 rounded-lg text-sm text-center font-bold"/>
                        <span>至</span>
                        <input type="number" min="1" max="31" value={aiConfig.dateRange.end} onChange={(e)=>setAiConfig({...aiConfig, dateRange: {...aiConfig.dateRange, end: parseInt(e.target.value) || 31}})} className="w-full border-blue-200 border p-2 rounded-lg text-sm text-center font-bold"/>
                    </div>
                </div>
                <div>
                    <label className="block text-xs font-bold text-blue-700 mb-2 uppercase">3. 指定班別 (非必填)</label>
                    <select value={aiConfig.targetShift} onChange={(e)=>setAiConfig({...aiConfig, targetShift: e.target.value})} className="w-full border-blue-200 border p-2 rounded-lg text-sm font-bold bg-white">
                        <option value="">由 AI 自由規劃</option>
                        {DICT.SHIFTS.map(s => <option key={s} value={s}>{s} 班</option>)}
                        <option value="off">休假 (off)</option>
                    </select>
                </div>
                <div className="flex items-end">
                    <button disabled={isAiLoading || aiConfig.selectedStaffs.length === 0} onClick={() => handleAiAutoSchedule(true)} className="w-full bg-blue-600 hover:bg-blue-700 text-white font-black py-2 rounded-xl transition-all shadow-lg active:scale-95 disabled:opacity-50 disabled:grayscale flex items-center justify-center gap-2">
                        {isAiLoading ? <Loader2 className="animate-spin" size={18}/> : <Check size={18}/>} 套用並補空
                    </button>
                </div>
            </div>
        </div>
      )}

      <div className="max-w-[95vw] mx-auto mb-6 grid grid-cols-1 lg:grid-cols-12 gap-4">
        <div className="lg:col-span-4 bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex items-center gap-4">
          <input type="number" value={year} onChange={(e) => setYear(Number(e.target.value))} className="w-24 border rounded-lg p-2 text-center font-bold" />
          <select value={month} onChange={(e) => setMonth(Number(e.target.value))} className="w-20 border rounded-lg p-2 text-center font-bold">
            {[...Array(12).keys()].map(m => <option key={m+1} value={m+1}>{m+1}月</option>)}
          </select>
        </div>
        <div className="lg:col-span-3 bg-blue-600 p-4 rounded-xl shadow-md text-white flex flex-col justify-center">
          <span className="text-xs opacity-80 uppercase tracking-wider">本月應休天數</span>
          <span className="text-2xl font-black">{requiredLeaves} <small className="text-sm font-normal">DAYS</small></span>
        </div>
        <div className="lg:col-span-5 flex items-center justify-end gap-2">
          <button onClick={() => setShowSettings(!showSettings)} className="bg-white border p-3 rounded-xl hover:bg-slate-50 transition-colors"><Settings size={20} className="text-slate-600" /></button>
        </div>
      </div>

      {showSettings && (
        <div className="max-w-[95vw] mx-auto mb-6 bg-white p-6 rounded-2xl border-2 border-dashed border-slate-200 animate-fade-in-down">
          <h3 className="font-bold mb-4 flex items-center gap-2 text-slate-700"><Settings size={18}/> 條件設置</h3>
          <div className="grid md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-bold text-slate-600 mb-2">國定假日清單:</label>
              <textarea value={holidaysText} onChange={(e) => setHolidaysText(e.target.value)} className="w-full border rounded-xl p-3 text-sm focus:ring-4 focus:ring-blue-100 transition-all outline-none" rows="3" />
            </div>
            <div className="flex gap-4">
              <div><label className="block text-sm font-bold text-slate-600 mb-2">週末底色</label><input type="color" value={colors.weekend} onChange={(e) => setColors({...colors, weekend: e.target.value})} className="h-10 w-20 rounded cursor-pointer" /></div>
              <div><label className="block text-sm font-bold text-slate-600 mb-2">假日底色</label><input type="color" value={colors.holiday} onChange={(e) => setColors({...colors, holiday: e.target.value})} className="h-10 w-20 rounded cursor-pointer" /></div>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-[95vw] mx-auto bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto pb-4">
          <table className="w-max min-w-full border-collapse">
            <thead>
              <tr className="bg-slate-100 border-b-2 border-slate-200">
                <th className="sticky left-0 bg-slate-200 z-30 p-4 border-r font-black text-slate-700 w-24 min-w-[96px]">班別</th>
                <th className="sticky left-[96px] bg-slate-200 z-30 p-4 border-r font-black text-slate-700 w-36 min-w-[144px]">日期/姓名</th>
                {daysInMonth.map(d => (
                  <th key={d.day} className="p-2 border-r min-w-[48px] text-center" style={{ backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent') }}>
                    <div className="text-[10px] opacity-60 uppercase">{d.weekStr}</div>
                    <div className="text-lg font-black">{d.day}</div>
                  </th>
                ))}
                <th className="p-4 border-r min-w-[60px] bg-blue-50 text-blue-700 font-bold">上班</th>
                <th className="p-4 border-r min-w-[60px] bg-green-50 text-green-700 font-bold">假日休</th>
                <th className="p-4 border-r min-w-[60px] bg-red-50 text-red-700 font-bold">總休</th>
                {DICT.LEAVES.map(l => <th key={l} className="p-2 border-r min-w-[40px] bg-slate-50 text-[10px] uppercase text-slate-500 font-bold">{l}</th>)}
                <th className="p-4 bg-slate-100 w-24">操作</th>
              </tr>
            </thead>
            <tbody>
              {groupedStaffs.map(({ group, staffs: groupStaffList }) => (
                <React.Fragment key={group}>
                  {groupStaffList.map((staff, index) => {
                    const stats = getStaffStats(staff.id);
                    const groupCount = groupStaffList.length + 1;
                    const groupIndex = groupStaffList.findIndex(s => s.id === staff.id);

                    return (
                      <tr key={staff.id} className="border-b border-slate-100 hover:bg-slate-50/50 transition-colors">
                        {index === 0 && (
                          <td
                            rowSpan={groupCount}
                            className="sticky left-0 bg-white z-20 border-r p-3 text-center shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)]"
                          >
                            <div className="flex items-center justify-center h-full min-h-[80px]">
                              <span className="text-3xl font-black text-slate-800 leading-tight tracking-[0.2em] [writing-mode:vertical-rl]">
                                {group}
                              </span>
                            </div>
                          </td>
                        )}

                        <td className="sticky left-[96px] bg-white z-10 border-r p-2 shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)]">
                          <input
                            type="text"
                            value={staff.name}
                            onChange={(e) => {
                              const next = [...staffs];
                              const currentIndex = next.findIndex(s => s.id === staff.id);
                              if (currentIndex !== -1) next[currentIndex].name = e.target.value;
                              setStaffs(next);
                            }}
                            className="w-full text-center py-2 font-bold text-slate-700 border-none rounded-lg focus:ring-2 focus:ring-blue-400 bg-transparent"
                          />
                        </td>

                        {daysInMonth.map(d => {
                          const cellData = schedule[staff.id]?.[d.date];
                          const val = typeof cellData === 'object' && cellData !== null ? (cellData?.value || '') : (cellData || '');
                          return (
                            <td key={d.date} className="border-r p-0" style={{ backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent'), opacity: d.isHoliday || d.isWeekend ? 0.9 : 1 }}>
                              <select value={val} onChange={(e) => handleCellChange(staff.id, d.date, e.target.value)} className={`w-full h-10 text-center bg-transparent border-none cursor-pointer text-sm font-bold appearance-none hover:bg-black/5 ${DICT.LEAVES.includes(val) ? 'text-red-500' : 'text-slate-800'}`}>
                                <option value=""></option>
                                <optgroup label="上班">{DICT.SHIFTS.map(s => <option key={s} value={s}>{s}</option>)}</optgroup>
                                <optgroup label="休假">{DICT.LEAVES.map(l => <option key={l} value={l}>{l}</option>)}</optgroup>
                              </select>
                            </td>
                          );
                        })}

                        <td className="border-r text-center font-black text-blue-600 bg-blue-50/30">{stats.work}</td>
                        <td className="border-r text-center font-black text-green-600 bg-green-50/30">{stats.holidayLeave}</td>
                        <td className="border-r text-center font-black text-red-600 bg-red-50/30">{stats.totalLeave}</td>
                        {DICT.LEAVES.map(l => <td key={l} className="border-r text-center text-[10px] text-slate-500 bg-slate-50/20">{stats.leaveDetails[l] || ''}</td>)}
                        <td className="p-2">
                          <div className="flex justify-center gap-1">
                            <button onClick={() => moveStaffInGroup(staff.id, 'up')} disabled={groupIndex === 0} className="text-slate-400 hover:text-blue-500 disabled:opacity-10">
                              <ArrowUp size={14}/>
                            </button>
                            <button onClick={() => moveStaffInGroup(staff.id, 'down')} disabled={groupIndex === groupStaffList.length - 1} className="text-slate-400 hover:text-blue-500 disabled:opacity-10">
                              <ArrowDown size={14}/>
                            </button>
                            <button onClick={() => removeStaff(staff.id)} className="text-slate-400 hover:text-red-500">
                              <Minus size={14}/>
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}

                  <tr className="border-b border-slate-200 bg-slate-50/70">
                    <td className="sticky left-[96px] bg-white z-10 border-r p-2 shadow-[4px_0_10px_-5px_rgba(0,0,0,0.1)]">
                      <button
                        onClick={() => addStaff(group)}
                        className="bg-blue-600 text-white px-3 py-2 rounded-xl hover:bg-blue-700 transition-colors flex items-center gap-2 font-bold text-sm mx-auto"
                      >
                        <Plus size={16}/> 新增人員
                      </button>
                    </td>

                    {daysInMonth.map(d => (
                      <td
                        key={d.date}
                        className="border-r"
                        style={{ backgroundColor: d.isHoliday ? colors.holiday : (d.isWeekend ? colors.weekend : 'transparent'), opacity: d.isHoliday || d.isWeekend ? 0.9 : 1 }}
                      ></td>
                    ))}

                    <td colSpan={3 + DICT.LEAVES.length + 1}></td>
                  </tr>
                </React.Fragment>
              ))}
            </tbody>
            <tfoot className="bg-slate-100 border-t-2 border-slate-200">
              {['D', 'E', 'N', 'totalLeave'].map((rowKey) => (
                <tr key={rowKey}>
                  <td className="sticky left-0 bg-slate-200 z-10 border-r p-3 w-24 min-w-[96px]"></td>
                  <td className="sticky left-[96px] bg-slate-200 z-10 border-r p-3 text-right text-xs font-bold text-slate-600 min-w-[144px]">
                    {rowKey === 'totalLeave' ? '當日休假' : `${rowKey} 班人數`}
                  </td>
                  {daysInMonth.map(d => {
                    const count = getDailyStats(d.date)[rowKey];
                    return <td key={d.date} className="border-r p-2 text-center text-sm font-black text-slate-700">{count || ''}</td>;
                  })}
                  <td colSpan={3 + DICT.LEAVES.length + 1}></td>
                </tr>
              ))}
            </tfoot>
          </table>
        </div>
      </div>

      {showHistoryModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-lg overflow-hidden shadow-2xl animate-pulse-once">
            <div className="p-6 border-b flex justify-between items-center bg-slate-50">
              <h3 className="font-black text-slate-800 flex items-center gap-2"><Clock /> 歷史存檔紀錄</h3>
              <button onClick={() => setShowHistoryModal(false)} className="p-2 hover:bg-slate-200 rounded-full transition-colors"><X/></button>
            </div>
            <div className="max-h-[60vh] overflow-y-auto p-4 space-y-3">
              {historyList.length === 0 ? <p className="text-center py-10 text-slate-400 font-bold">目前尚無存檔紀錄</p> : historyList.map(record => (
                <div key={record.id} className="flex items-center justify-between p-4 border rounded-2xl hover:border-blue-400 hover:bg-blue-50/30 transition-all group">
                  <div>
                    <div className="font-black text-slate-800">{record.label} <span className="text-xs font-normal text-slate-400 ml-2">{record.state.year}/{record.state.month}</span></div>
                    <div className="text-xs text-slate-500">{record.timestamp}</div>
                  </div>
                  <button onClick={() => loadHistory(record)} className="bg-slate-100 text-slate-700 px-4 py-2 rounded-xl text-sm font-bold hover:bg-blue-600 hover:text-white transition-all">載入</button>
                </div>
              ))}
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
