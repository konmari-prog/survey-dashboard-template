/**
 * アンケートダッシュボード テンプレート - メイン処理
 *
 * メニュー、フォーム回答の取り込み、onFormSubmitトリガー
 */

// === UI安全呼び出し ===
function showAlert_(message) {
  try { SpreadsheetApp.getUi().alert(message); }
  catch (e) { Logger.log(message); }
}

// === メニュー追加 ===
function onOpen() {
  // カラー選択後に保留されたトリガー登録を実行
  completePendingSetup_();

  const menu = SpreadsheetApp.getUi().createMenu('アンケート管理');

  if (!isSetupDone_()) {
    menu.addItem('初期セットアップ', 'runSetup');
  } else {
    menu.addItem('フォーム回答を取り込む', 'syncFormResponses');
    menu.addSeparator();

    // --- Webページを開く ---
    const webAppUrl = getWebAppUrl_();
    if (webAppUrl) {
      menu.addItem('ダッシュボードを開く', 'openDashboard');
      menu.addItem('Q&A質問回答集を開く', 'openQA');
    }
    menu.addSeparator();

    // --- 設定 ---
    menu.addItem('トリガーを設定', 'manualSetupTrigger');
    menu.addItem('WebアプリURL設定', 'promptWebAppUrl_');
    menu.addItem('セットアップをやり直す', 'resetSetup');
  }

  menu.addSeparator();
  menu.addItem('セットアップガイドを開く', 'openGuide');
  menu.addItem('デモフォームを作成', 'createDemoForms');

  menu.addToUi();
}

// === 新しいタブでURLを開くヘルパー ===
function openUrl_(url, label) {
  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '","_blank");google.script.host.close();</script>'
  ).setWidth(200).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, label + 'を開いています...');
}

// === ダッシュボードを開く ===
function openDashboard() {
  const url = getWebAppUrl_();
  if (!url) { showAlert_('WebアプリURLが未設定です。\nメニューの「WebアプリURL設定」から登録してください。'); return; }
  openUrl_(url + '?page=dashboard', 'ダッシュボード');
}

// === Q&A質問回答集を開く ===
function openQA() {
  const url = getWebAppUrl_();
  if (!url) { showAlert_('WebアプリURLが未設定です。\nメニューの「WebアプリURL設定」から登録してください。'); return; }
  openUrl_(url, 'Q&A質問回答集');
}

// === セットアップガイドを開く ===
function openGuide() {
  const url = getWebAppUrl_();
  if (url) {
    openUrl_(url + '?page=guide', 'セットアップガイド');
  } else {
    // WebApp未デプロイ時はダイアログで直接表示
    const html = HtmlService.createHtmlOutputFromFile('guide')
      .setTitle('セットアップガイド')
      .setWidth(800)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'セットアップガイド');
  }
}

// === フォーム回答をQ&Aシートに取り込む ===
function syncFormResponses() {
  const config = getConfig_();
  if (!config) {
    showAlert_('セットアップが完了していません。\nメニューから「初期セットアップ」を実行してください。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let qaSheet = ss.getSheetByName(config.qaSheet);
  if (!qaSheet) {
    showAlert_('Q&Aシートが見つかりません。セットアップを再実行してください。');
    return;
  }

  // 既存の質問を取得（重複防止）
  const qaData = qaSheet.getDataRange().getValues();
  const existingQuestions = new Set();
  for (let i = 1; i < qaData.length; i++) {
    existingQuestions.add(String(qaData[i][QA_COL.QUESTION - 1]).trim());
  }

  let qaCount = 0;
  let nextQaNo = qaData.length;

  // 各フォームシートから質問を取り込む
  config.sheets.forEach(sheetConfig => {
    const sheet = ss.getSheetByName(sheetConfig.name);
    if (!sheet) return;

    const questionCol = sheetConfig.columns.question;
    const sessionCol = sheetConfig.columns.session;
    if (!questionCol) return; // 質問列がなければスキップ

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const question = String(data[i][questionCol - 1] || '').trim();
      const session = sessionCol ? (data[i][sessionCol - 1] || '') : '';

      if (question && !existingQuestions.has(question)) {
        qaSheet.appendRow([nextQaNo + 1, session, '', question, '']);
        existingQuestions.add(question);
        qaCount++;
        nextQaNo++;
      }
    }
  });

  showAlert_('取り込み完了!\n\nQ&A質問: ' + qaCount + '件追加');
}

// === トリガー手動設定（メニューから実行） ===
function manualSetupTrigger() {
  setupFormTrigger_();
  PropertiesService.getScriptProperties().deleteProperty('TRIGGER_PENDING_');
  showAlert_('フォーム送信トリガーを設定しました!\n\nフォームから回答が送信されると、Q&Aシートに質問が自動で取り込まれます。');
}

// === セットアップ保留タスクの完了（onOpen時に実行） ===
function completePendingSetup_() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('TRIGGER_PENDING_') !== 'true') return;

    // トリガー登録
    setupFormTrigger_();
    props.deleteProperty('TRIGGER_PENDING_');
    Logger.log('保留中のトリガー登録を完了しました');

    // セットアップ完了メッセージ
    const config = getConfig_();
    if (config) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'セットアップ完了!',
        'プロジェクト名: ' + config.projectName + '\n' +
        'テーマカラー: ' + (COLOR_PRESETS[config.colorKey] || {}).name + '\n' +
        'フォームシート: ' + config.sheets.length + '件\n' +
        'Q&Aシート: 作成済み\n' +
        'フォーム送信トリガー: 設定済み\n\n' +
        '次に Webアプリをデプロイして URL を設定してください。\n' +
        'メニュー「アンケート管理」>「WebアプリURL設定」から入力できます。',
        ui.ButtonSet.OK
      );
    }
  } catch (e) {
    Logger.log('completePendingSetup_ エラー: ' + e.message);
  }
}

// === フォーム送信時の自動取り込み（トリガーで実行） ===
function onFormSubmit(e) {
  Logger.log('=== onFormSubmit 発火 ===');
  if (!e || !e.values) {
    Logger.log('イベントデータなし。終了');
    return;
  }

  const config = getConfig_();
  if (!config) {
    Logger.log('config が見つかりません。セットアップ未完了');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let qaSheet = ss.getSheetByName(config.qaSheet || 'Q&A');
  if (!qaSheet) {
    Logger.log('Q&Aシートが見つかりません');
    return;
  }

  const row = e.values;
  Logger.log('受信データ: ' + JSON.stringify(row));
  Logger.log('列数: ' + row.length);
  Logger.log('config.sheets: ' + JSON.stringify(config.sheets));

  // どのフォームシートから来たか判定（列数でマッチ）
  let matchedSheet = null;
  for (const sheetConfig of config.sheets) {
    Logger.log('シート比較: ' + sheetConfig.name + ' columnCount=' + sheetConfig.columnCount + ' vs row.length=' + row.length);
    if (row.length <= sheetConfig.columnCount + 1) {
      matchedSheet = sheetConfig;
      break;
    }
  }
  if (!matchedSheet && config.sheets.length > 0) {
    matchedSheet = config.sheets[0];
    Logger.log('列数マッチなし。デフォルトシート使用: ' + matchedSheet.name);
  }
  if (!matchedSheet) {
    Logger.log('マッチするシート設定がありません。終了');
    return;
  }

  Logger.log('マッチしたシート: ' + matchedSheet.name);
  Logger.log('columns: ' + JSON.stringify(matchedSheet.columns));

  const questionCol = matchedSheet.columns.question;
  const sessionCol = matchedSheet.columns.session;
  Logger.log('questionCol=' + questionCol + ', sessionCol=' + sessionCol);

  if (!questionCol) {
    Logger.log('question列がマッピングされていません。終了');
    return;
  }

  const question = String(row[questionCol - 1] || '').trim();
  const session = sessionCol ? (row[sessionCol - 1] || '') : '';
  Logger.log('質問: "' + question + '", セッション: "' + session + '"');

  if (question) {
    const qaData = qaSheet.getDataRange().getValues();
    const nextNo = qaData.length;
    qaSheet.appendRow([nextNo, session, '', question, '']);
    Logger.log('Q&Aシートに追加: No.' + nextNo);
  } else {
    Logger.log('質問が空のためスキップ');
  }
}
