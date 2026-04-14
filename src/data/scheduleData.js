import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'

ReactDOM.createRoot(document.getElementById('root')).render(<App />)
export const DICT = {
  SHIFTS: ['D', 'E', 'N', '白8-8', '夜8-8', '8-12', '12-16'],
  LEAVES: ['off', '例', '休', '特', '補', '國', '喪', '婚', '產', '病', '事', '陪產', 'AM', 'PM']
};

export const SHIFT_GROUPS = ['白班', '小夜', '大夜'];

export const GROUP_TO_DEMAND_KEY = {
  '白班': 'white',
  '小夜': 'evening',
  '大夜': 'night'
};

export const DEFAULT_REQUIRED_STAFFING = {
  weekday: { white: 6, evening: 3, night: 2 },
  saturday: { white: 4, evening: 2, night: 2 },
  sunday: { white: 4, evening: 2, night: 2 }
};

export const DEFAULT_SHIFT_BY_GROUP = {
  '白班': 'D',
  '小夜': 'E',
  '大夜': 'N'
};

export const RULE_FILL_MAIN_SHIFTS = ['D', 'E', 'N'];

export const HOSPITAL_LEVEL_LABELS = {
  medical: '醫學中心',
  regional: '區域醫院',
  local: '地區醫院'
};

export const HOSPITAL_RATIO_HINTS = {
  medical: { white: '1:6', evening: '1:9', night: '1:11' },
  regional: { white: '1:7', evening: '1:11', night: '1:13' },
  local: { white: '1:10', evening: '1:13', night: '1:15' }
};

export const STORAGE_KEYS = {
  HISTORY: 'schedule_app_history',
  ACTIVE_DRAFT: 'schedule_app_active_draft',
  LOCAL_SETTINGS: 'schedule_app_local_settings'
};