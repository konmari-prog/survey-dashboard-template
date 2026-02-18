/**
 * アンケートダッシュボード テンプレート - 設定ユーティリティ
 *
 * Setup.gs の runSetup() で自動生成された設定を読み込む。
 * 手動での編集は不要。再設定したい場合はメニューから「初期セットアップ」を再実行。
 */

// === セットアップ済み設定の読み込み ===
function getConfig_() {
  const json = PropertiesService.getScriptProperties().getProperty('SURVEY_CONFIG');
  if (!json) return null;
  try {
    return JSON.parse(json);
  } catch (e) {
    Logger.log('設定の読み込みに失敗: ' + e.message);
    return null;
  }
}

// === 設定の保存 ===
function saveConfig_(config) {
  PropertiesService.getScriptProperties().setProperty('SURVEY_CONFIG', JSON.stringify(config));
}

// === Q&Aシートの列番号（固定：自動作成するため） ===
const QA_COL = {
  NO: 1,
  SESSION: 2,
  DATE: 3,
  QUESTION: 4,
  ANSWER: 5,
};

// === カラープリセット ===
const COLOR_PRESETS = {
  pink:    { primary: '#F490A1', gradient: ['#F490A1', '#E87A8E'], name: 'ピンク' },
  blue:    { primary: '#4A90D9', gradient: ['#4A90D9', '#357ABD'], name: 'ブルー' },
  green:   { primary: '#27AE60', gradient: ['#27AE60', '#1E8449'], name: 'グリーン' },
  orange:  { primary: '#F39C12', gradient: ['#F39C12', '#E67E22'], name: 'オレンジ' },
  shiftai: { primary: '#ff5757', gradient: ['#ff5757', '#8c52ff'], name: 'SHIFT AI' },
};

// === WebアプリURL（スクリプトプロパティから取得） ===
function getWebAppUrl_() {
  return PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || '';
}
