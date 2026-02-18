/**
 * アンケートダッシュボード テンプレート - 初期セットアップ
 *
 * メニューから「初期セットアップ」を実行すると:
 * 1. フォーム回答シートを自動検出
 * 2. ヘッダー行から列を自動マッピング
 * 3. プロジェクト名を入力
 * 4. カラー選択（HTMLダイアログ）
 * 5. Q&Aシート作成 → トリガー登録 → URL設定
 */

// === ヘッダーマッチ用キーワード定義 ===
const HEADER_KEYWORDS_ = {
  timestamp:     ['タイムスタンプ', 'timestamp'],
  email:         ['メール', 'email', 'mail'],
  memberType:    ['会員', '区分', '所属', 'メンバー', 'member', '属性'],
  session:       ['ウェビナー', '回', 'セッション', '講座', '研修', 'session', '開催'],
  joinMethod:    ['参加方法', '参加', '視聴', 'join'],
  satisfaction:  ['満足度', '満足', 'satisfaction'],
  understanding: ['理解度', '理解', 'understanding'],
  impression:    ['印象', '感想', '学び', 'impression'],
  nextAction:    ['やること', 'アクション', '実践', '行動', 'action', 'next'],
  output:        ['アウトプット', '実行', '成果物', 'output'],
  question:      ['質問', 'question', 'Q&A'],
};

// === メインのセットアップ関数 ===
function runSetup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Step 1: フォーム回答シートを自動検出 ---
  const allSheets = ss.getSheets();
  const detectedSheets = [];

  allSheets.forEach(sheet => {
    const name = sheet.getName();
    if (name === 'Q&A') return;

    const lastCol = sheet.getLastColumn();
    if (lastCol < 3) return;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const firstHeader = String(headers[0] || '').trim();

    if (firstHeader === 'タイムスタンプ' || firstHeader.toLowerCase() === 'timestamp') {
      detectedSheets.push({ name: name, headers: headers });
    }
  });

  if (detectedSheets.length === 0) {
    ui.alert('フォーム回答シートが見つかりませんでした。\n\nGoogleフォームを紐付けてから再実行してください。\n（1列目のヘッダーが「タイムスタンプ」のシートを検出します）');
    return;
  }

  const sheetNames = detectedSheets.map(s => '  - ' + s.name + '（' + s.headers.length + '列）').join('\n');
  const confirmDetect = ui.alert(
    'フォーム回答シートを検出しました',
    detectedSheets.length + '件のシートが見つかりました:\n\n' + sheetNames + '\n\nこのまま設定を続けますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirmDetect !== ui.Button.YES) return;

  // --- Step 2: ヘッダーから列を自動マッピング ---
  const sheets = [];
  const allFreeTextCols = [];

  detectedSheets.forEach(sheetInfo => {
    const columns = {};
    const freeTextCols = [];
    const matchedIndices = new Set();

    sheetInfo.headers.forEach((header, idx) => {
      const h = String(header || '').trim();
      if (!h) return;
      const colNum = idx + 1;
      let matched = false;

      for (const [key, keywords] of Object.entries(HEADER_KEYWORDS_)) {
        for (const kw of keywords) {
          if (h.includes(kw) || h.toLowerCase().includes(kw.toLowerCase())) {
            if (!columns[key]) {
              columns[key] = colNum;
              matchedIndices.add(idx);
              matched = true;
            }
            break;
          }
        }
        if (matched) break;
      }
    });

    sheetInfo.headers.forEach((header, idx) => {
      const h = String(header || '').trim();
      if (!h) return;
      if (!matchedIndices.has(idx) && idx > 0) {
        freeTextCols.push({ name: h, colIndex: idx + 1, sheet: sheetInfo.name });
      }
    });

    sheets.push({
      name: sheetInfo.name,
      columnCount: sheetInfo.headers.length,
      columns: columns,
    });

    freeTextCols.forEach(f => allFreeTextCols.push(f));
  });

  // マッピング結果をサマリー表示
  const summaryLines = [];
  sheets.forEach(s => {
    summaryLines.push('【' + s.name + '】');
    for (const [key, col] of Object.entries(s.columns)) {
      summaryLines.push('  ' + key + ' → ' + col + '列目');
    }
  });
  if (allFreeTextCols.length > 0) {
    summaryLines.push('');
    summaryLines.push('【自由記述列（受講者の声に表示）】');
    allFreeTextCols.forEach(f => {
      summaryLines.push('  ' + f.name + '（' + f.sheet + ' ' + f.colIndex + '列目）');
    });
  }

  const confirmMapping = ui.alert(
    '列の自動マッピング結果',
    summaryLines.join('\n') + '\n\nこの設定で続けますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirmMapping !== ui.Button.YES) return;

  // --- Step 3: プロジェクト名の入力 ---
  const nameResult = ui.prompt(
    'プロジェクト名を入力してください',
    'ダッシュボードのヘッダーに表示されます（例: おうちAIラボ Season2）',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResult.getSelectedButton() !== ui.Button.OK) return;
  const projectName = nameResult.getResponseText().trim() || 'アンケートダッシュボード';

  // --- Step 1-3 の結果を一時保存（カラー選択ダイアログは非同期のため） ---
  const pending = {
    projectName: projectName,
    sheets: sheets,
    freeTextColumns: allFreeTextCols,
  };
  PropertiesService.getScriptProperties().setProperty('SETUP_PENDING_', JSON.stringify(pending));
  Logger.log('SETUP_PENDING_ を保存しました: ' + JSON.stringify(pending).substring(0, 200));

  // --- Step 4: カラー選択ダイアログを表示 ---
  const html = HtmlService.createHtmlOutputFromFile('colorpicker')
    .setWidth(340)
    .setHeight(420);
  ui.showModalDialog(html, 'テーマカラーを選択');
}

// === カラー選択の受け取り（colorpicker.html から呼ばれる） ===
function receiveColorChoice(colorKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 一時保存データを取得
  const props = PropertiesService.getScriptProperties();
  const pendingJson = props.getProperty('SETUP_PENDING_');
  Logger.log('SETUP_PENDING_ = ' + pendingJson);
  if (!pendingJson) {
    // 一時データがない場合、最低限のデフォルト設定で続行を試みる
    Logger.log('SETUP_PENDING_ が見つかりません。runSetup() を先に実行してください。');
    throw new Error('セットアップデータが見つかりません。メニューから「初期セットアップ」を再実行してください。');
  }
  const pending = JSON.parse(pendingJson);
  props.deleteProperty('SETUP_PENDING_'); // 一時データをクリーンアップ

  const selectedColor = COLOR_PRESETS[colorKey] || COLOR_PRESETS['blue'];

  // --- Step 5: Q&Aシートを作成 ---
  let qaSheet = ss.getSheetByName('Q&A');
  if (!qaSheet) {
    qaSheet = ss.insertSheet('Q&A');
    qaSheet.getRange(1, 1, 1, 5).setValues([
      ['No.', 'セッション名', '公開日', 'ご質問', '回答欄']
    ]);
    qaSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    qaSheet.setFrozenRows(1);
    qaSheet.setColumnWidth(2, 250);
    qaSheet.setColumnWidth(4, 400);
    qaSheet.setColumnWidth(5, 400);
  }

  // --- Step 6: 設定を保存 ---
  const config = {
    projectName: pending.projectName,
    colorKey: colorKey,
    primaryColor: selectedColor.primary,
    gradient: selectedColor.gradient,
    sheets: pending.sheets,
    freeTextColumns: pending.freeTextColumns,
    qaSheet: 'Q&A',
    setupDate: new Date().toISOString(),
  };

  saveConfig_(config);

  // --- Step 7: onFormSubmitトリガーを自動登録 ---
  // ※ HTMLダイアログ経由では ScriptApp の権限が制限されるため
  //    トリガー登録はフラグを立てて、次回メニュー操作時に実行する
  props.setProperty('TRIGGER_PENDING_', 'true');
  Logger.log('トリガー登録を保留しました（次回メニュー操作時に実行）');

  // 完了メッセージ（google.script.run 経由なので ui.alert は使えない）
  // → colorpicker.html の successHandler でダイアログが閉じる
  Logger.log('セットアップ設定保存完了: ' + pending.projectName + ' / ' + selectedColor.name);
}

// === onFormSubmitトリガーを自動登録（既存があれば二重登録しない） ===
function setupFormTrigger_() {
  const existing = ScriptApp.getProjectTriggers();
  const hasFormTrigger = existing.some(t =>
    t.getHandlerFunction() === 'onFormSubmit' &&
    t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT
  );

  if (!hasFormTrigger) {
    ScriptApp.newTrigger('onFormSubmit')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onFormSubmit()
      .create();
  }
}

// === デプロイURL入力プロンプト ===
function promptWebAppUrl_() {
  const ui = SpreadsheetApp.getUi();
  const currentUrl = getWebAppUrl_();

  const result = ui.prompt(
    'WebアプリのURLを入力',
    'Apps Script → デプロイ → 新しいデプロイ でURLを取得し、貼り付けてください。\n\n' +
    '※ まだデプロイしていない場合は「キャンセル」して、デプロイ後にメニューから\n' +
    '「WebアプリURL設定」で入力できます。' +
    (currentUrl ? '\n\n現在のURL: ' + currentUrl : ''),
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const url = result.getResponseText().trim();
  if (!url) return;

  PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', url);
  ui.alert('URL設定完了!', 'メニューからダッシュボードやQ&Aを直接開けます。\n\nQ&A質問回答集: ' + url + '\nダッシュボード: ' + url + '?page=dashboard', ui.ButtonSet.OK);
}

// === セットアップ状態の確認 ===
function isSetupDone_() {
  return getConfig_() !== null;
}

// === セットアップのリセット ===
function resetSetup() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '設定をリセット',
    '現在の設定を削除して、再セットアップできるようにします。\nQ&Aシートのデータは保持されます。\n\nよろしいですか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  PropertiesService.getScriptProperties().deleteProperty('SURVEY_CONFIG');
  ui.alert('設定をリセットしました。\nメニューから「初期セットアップ」を再実行してください。');
}
