/**
 * アンケートダッシュボード テンプレート - Webアプリ
 *
 * ページルーティング＆Q&AデータAPI
 * 設定からプロジェクト名・カラーを動的に注入
 */

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || '';
  const config = getConfig_();
  const projectName = config ? config.projectName : 'アンケート';

  if (page === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('dashboard')
      .setTitle(projectName + ' ダッシュボード')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'guide') {
    return HtmlService.createHtmlOutputFromFile('guide')
      .setTitle('セットアップガイド')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // デフォルト: Q&A質問回答集
  const template = HtmlService.createTemplateFromFile('index');
  template.projectName = projectName;
  template.primaryColor = config ? config.primaryColor : '#4A90D9';
  template.gradient = config ? config.gradient : ['#4A90D9', '#357ABD'];

  return template.evaluate()
    .setTitle(projectName + ' Q&A')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// === Q&Aデータをセッションごとにグルーピングして返す ===
function getQAData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig_();
  const qaSheetName = config ? config.qaSheet : 'Q&A';
  const qaSheet = ss.getSheetByName(qaSheetName);
  if (!qaSheet) return [];

  const data = qaSheet.getDataRange().getValues();
  const grouped = {};

  for (let i = 1; i < data.length; i++) {
    const no = data[i][QA_COL.NO - 1];
    const session = String(data[i][QA_COL.SESSION - 1]).trim() || '未分類';
    const date = data[i][QA_COL.DATE - 1];
    const question = String(data[i][QA_COL.QUESTION - 1]).trim();
    const answer = String(data[i][QA_COL.ANSWER - 1]).trim();

    if (!question) continue;
    if (!answer) continue; // 未回答はWebアプリに表示しない

    let dateStr = '';
    if (date) {
      if (date instanceof Date) {
        dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年M月d日');
      } else {
        dateStr = String(date);
      }
    }

    if (!grouped[session]) grouped[session] = [];
    grouped[session].push({ no, date: dateStr, question, answer });
  }

  // 新しいセッションが上に来るよう逆順
  return Object.keys(grouped).reverse().map(session => ({
    session: session,
    items: grouped[session],
  }));
}
