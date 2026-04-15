export const loadHistoryListFromStorage = (historyKey) => {
  try {
    const stored = localStorage.getItem(historyKey);
    if (!stored) return [];
    const parsed = JSON.parse(stored);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    console.error('本機暫存紀錄解析失敗', error);
    return [];
  }
};

export const saveHistoryListToStorage = (historyKey, historyList = []) => {
  try {
    localStorage.setItem(historyKey, JSON.stringify(historyList || []));
    return true;
  } catch (error) {
    console.error('本機暫存紀錄寫入失敗', error);
    return false;
  }
};

export const clearHistoryStorage = (historyKey) => {
  try {
    localStorage.removeItem(historyKey);
    return true;
  } catch (error) {
    console.error('本機暫存紀錄清除失敗', error);
    return false;
  }
};

export const loadLocalSettingsBlob = (storageKey) => {
  try {
    const stored = localStorage.getItem(storageKey);
    if (!stored) return null;
    const parsed = JSON.parse(stored);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (error) {
    console.error('本機設定解析失敗', error);
    return null;
  }
};
