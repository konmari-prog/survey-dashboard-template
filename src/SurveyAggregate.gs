/**
 * アンケートダッシュボード テンプレート - ダッシュボード用データ提供
 *
 * 複数フォームシートを統合して正規化し、ダッシュボードに返す。
 * 列マッピングはSetup.gsで自動検出した設定に基づく。
 */

// === 正規化後の共通キー一覧 ===
const NORM_KEYS_ = [
  'timestamp', 'email', 'memberType', 'session',
  'joinMethod', 'satisfaction', 'understanding',
  'impression', 'nextAction', 'output', 'question',
];

// === ダッシュボード用API ===
function getDashboardData() {
  const config = getConfig_();
  if (!config) return { rows: [], columnMap: {}, freeTextColumns: [], projectName: '', primaryColor: '#4A90D9', gradient: ['#4A90D9', '#357ABD'] };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = getMergedSurveyRows_(ss, config);

  return {
    rows: result.rows,
    columnMap: result.columnMap,         // { satisfaction: 5, understanding: 6, ... } 0始まり
    freeTextColumns: result.freeTextCols, // [{ name: 'サポート要望', colIndex: 11 }, ...]
    projectName: config.projectName,
    primaryColor: config.primaryColor,
    gradient: config.gradient || [config.primaryColor, config.primaryColor],
  };
}

/**
 * 全フォームシートのデータを統合して正規化する
 * 正規化列: NORM_KEYS_ の順 + 自由記述列
 */
function getMergedSurveyRows_(ss, config) {
  // 自由記述列のユニーク名リスト
  const freeTextNames = [];
  const freeTextNameSet = new Set();
  (config.freeTextColumns || []).forEach(f => {
    if (!freeTextNameSet.has(f.name)) {
      freeTextNames.push(f.name);
      freeTextNameSet.add(f.name);
    }
  });

  const totalCols = NORM_KEYS_.length + freeTextNames.length;

  // columnMap: 正規化後のインデックス（0始まり）
  const columnMap = {};
  NORM_KEYS_.forEach((key, idx) => { columnMap[key] = idx; });
  freeTextNames.forEach((name, idx) => { columnMap['free_' + idx] = NORM_KEYS_.length + idx; });

  const allRows = [];

  config.sheets.forEach(sheetConfig => {
    const sheet = ss.getSheetByName(sheetConfig.name);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();

    // このシートの自由記述列のマッピング
    const sheetFreeText = (config.freeTextColumns || []).filter(f => f.sheet === sheetConfig.name);
    const freeTextMap = {}; // { 列番号(1始まり): freeTextNames内のindex }
    sheetFreeText.forEach(f => {
      const nameIdx = freeTextNames.indexOf(f.name);
      if (nameIdx >= 0) freeTextMap[f.colIndex] = nameIdx;
    });

    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const normalized = new Array(totalCols).fill('');

      // NORM_KEYS_ の各キーをマッピング
      NORM_KEYS_.forEach((key, normIdx) => {
        const srcCol = sheetConfig.columns[key]; // 1始まり or undefined
        if (srcCol && srcCol <= r.length) {
          let val = r[srcCol - 1];
          // 日付はISO文字列に変換
          if (val instanceof Date) {
            val = Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
          }
          normalized[normIdx] = val;
        }
      });

      // 自由記述列をマッピング
      for (const [colStr, nameIdx] of Object.entries(freeTextMap)) {
        const col = parseInt(colStr);
        if (col <= r.length) {
          normalized[NORM_KEYS_.length + nameIdx] = r[col - 1] || '';
        }
      }

      allRows.push(normalized);
    }
  });

  // freeTextCols: ダッシュボードに返す自由記述情報
  const freeTextCols = freeTextNames.map((name, idx) => ({
    name: name,
    colIndex: NORM_KEYS_.length + idx, // 0始まり
  }));

  return { rows: allRows, columnMap: columnMap, freeTextCols: freeTextCols };
}
