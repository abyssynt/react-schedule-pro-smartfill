export const loadExcelJS = () => {
  return new Promise((resolve) => {
    if (window.ExcelJS) return resolve(window.ExcelJS);
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js";
    script.onload = () => resolve(window.ExcelJS);
    document.head.appendChild(script);
  });
};

export const loadSheetJS = () => {
  return new Promise((resolve, reject) => {
    if (window.XLSX) return resolve(window.XLSX);
    const existing = document.querySelector('script[data-sheetjs-loader="true"]');
    if (existing) {
      existing.addEventListener('load', () => resolve(window.XLSX), { once: true });
      existing.addEventListener('error', () => reject(new Error('SheetJS 載入失敗')), { once: true });
      return;
    }
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    script.dataset.sheetjsLoader = 'true';
    script.onload = () => resolve(window.XLSX);
    script.onerror = () => reject(new Error('SheetJS 載入失敗'));
    document.head.appendChild(script);
  });
};
