/**
 * デモ用 Googleフォーム＋ダミー回答 生成スクリプト
 *
 * メニュー「アンケート管理」>「デモフォームを作成」で実行
 * 3つのGoogleフォームを作成 → スプレッドシートに自動紐付け → ダミー回答を投入
 * すべて自動。実行後「初期セットアップ」するだけでダッシュボードが動く
 */

function createDemoForms() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'デモフォームを作成',
    '3つのGoogleフォーム + ダミー回答を自動作成します。\n\n' +
    '  1. 企業研修アンケート（8件）\n' +
    '  2. セミナーアンケート（7件）\n' +
    '  3. ワークショップアンケート（6件）\n\n' +
    'フォーム作成 → スプレッドシート紐付け → ダミー回答投入\n' +
    'すべて自動で行います。\n\n' +
    '作成しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ===== フォーム1: 企業研修風 =====
  ui.alert('進捗', '1/3 企業研修アンケートを作成中...', ui.ButtonSet.OK);
  const form1 = createAndLinkForm_('企業研修アンケート', ss, [
    { title: '所属',           type: 'list',     options: ['営業部', '人事部', '開発部', 'マーケ部', '総務部'] },
    { title: '研修名',         type: 'list',     options: ['第1回 AI基礎', '第2回 プロンプト実践', '第3回 業務活用'] },
    { title: '満足度',         type: 'scale',    low: 1, high: 5, lowLabel: '不満', highLabel: '大満足' },
    { title: '理解度',         type: 'list',     options: ['よく理解できた', 'だいたい理解できた', 'まあまあ', 'あまり理解できなかった'] },
    { title: '感想',           type: 'paragraph' },
    { title: '実践したいこと', type: 'paragraph' },
    { title: 'アウトプット状況', type: 'list',   options: ['試してみた', 'まだ'] },
    { title: '質問',           type: 'paragraph' },
  ]);

  submitDemoResponses_(form1, [
    ['営業部',   '第1回 AI基礎',       5, 'よく理解できた',   '身近な事例が多くて分かりやすかった', 'ChatGPTで議事録を要約する',   '試してみた', '社内データを使っても大丈夫ですか？'],
    ['人事部',   '第1回 AI基礎',       4, 'だいたい理解できた', 'AIの歴史パートが面白かった',       '採用メールの下書きをAIで作る', 'まだ',     ''],
    ['開発部',   '第1回 AI基礎',       5, 'よく理解できた',   '技術的な裏側も知れてよかった',      'コードレビューにCopilot導入',  '試してみた', 'APIの料金体系を詳しく知りたいです'],
    ['マーケ部', '第1回 AI基礎',       3, 'まあまあ',        '専門用語が多くて少し難しかった',     'SNS投稿文のたたき台をAIで',   'まだ',     'もう少しゆっくり進めてほしいです'],
    ['営業部',   '第1回 AI基礎',       4, 'だいたい理解できた', '実演が特に参考になった',           'お客様への提案書作成に活用',   'まだ',     ''],
    ['人事部',   '第2回 プロンプト実践', 5, 'よく理解できた',   'プロンプトの型が実用的',           '面接質問リストの生成',        '試してみた', '長文を要約するときのコツは？'],
    ['開発部',   '第2回 プロンプト実践', 5, 'よく理解できた',   'ハンズオンで手を動かせてよかった',   'テストケース作成の自動化',     '試してみた', ''],
    ['マーケ部', '第2回 プロンプト実践', 4, 'だいたい理解できた', '前回より格段にわかりやすかった',     'ペルソナ分析にAIを使う',      'まだ',     'マーケ向けのプロンプト集がほしい'],
  ]);

  // ===== フォーム2: セミナー風 =====
  ui.alert('進捗', '2/3 セミナーアンケートを作成中...', ui.ButtonSet.OK);
  const form2 = createAndLinkForm_('セミナーアンケート', ss, [
    { title: '会員区分',           type: 'list',     options: ['一般会員', 'プレミアム', '無料会員'] },
    { title: 'ウェビナー回',       type: 'list',     options: ['Vol.1 はじめてのAI', 'Vol.2 画像生成AI', 'Vol.3 自動化'] },
    { title: '全体の満足度',       type: 'scale',    low: 1, high: 5, lowLabel: '不満', highLabel: '大満足' },
    { title: '内容の理解度',       type: 'list',     options: ['よく理解できた', 'だいたい理解できた', 'まあまあ', 'あまり理解できなかった'] },
    { title: '学びになったこと',   type: 'paragraph' },
    { title: 'アクションプラン',   type: 'paragraph' },
    { title: '質問・気になること', type: 'paragraph' },
  ]);

  submitDemoResponses_(form2, [
    ['一般会員',   'Vol.1 はじめてのAI', 5, 'よく理解できた',   'AIの全体像がつかめた',            'まずChatGPTに登録する',        '無料と有料の違いは？'],
    ['プレミアム', 'Vol.1 はじめてのAI', 4, 'だいたい理解できた', '事例紹介が参考になった',           'チームに共有する',              ''],
    ['一般会員',   'Vol.1 はじめてのAI', 4, 'だいたい理解できた', '思ったより簡単に使えそう',          '日報の自動化を試す',             'セキュリティが心配です'],
    ['無料会員',   'Vol.1 はじめてのAI', 3, 'まあまあ',         '概要はわかったが深掘りしたい',       '本を1冊読む',                  'おすすめの書籍はありますか？'],
    ['プレミアム', 'Vol.2 画像生成AI',   5, 'よく理解できた',   'Midjourneyの可能性に驚いた',       'SNS用のサムネ作成に活用',        '商用利用のライセンスは？'],
    ['一般会員',   'Vol.2 画像生成AI',   5, 'よく理解できた',   '実際に作れて楽しかった',            'ブログのアイキャッチを自作',       ''],
    ['無料会員',   'Vol.2 画像生成AI',   4, 'だいたい理解できた', 'プロンプトの書き方が勉強になった',    'まずは無料ツールで練習',          'DALLEとMidjourneyどちらがおすすめ？'],
  ]);

  // ===== フォーム3: ワークショップ風 =====
  ui.alert('進捗', '3/3 ワークショップアンケートを作成中...', ui.ButtonSet.OK);
  const form3 = createAndLinkForm_('ワークショップアンケート', ss, [
    { title: 'メンバー種別',       type: 'list',     options: ['受講生', 'メンター', 'TA'] },
    { title: '講座',               type: 'list',     options: ['Day1 環境構築', 'Day2 API連携', 'Day3 デプロイ'] },
    { title: '満足',               type: 'scale',    low: 1, high: 5, lowLabel: '不満', highLabel: '大満足' },
    { title: '理解できた度合い',   type: 'list',     options: ['バッチリ', 'だいたいOK', 'ちょっと難しい', 'わからなかった'] },
    { title: '印象に残ったこと',   type: 'paragraph' },
    { title: 'やること',           type: 'paragraph' },
    { title: '成果物',             type: 'list',     options: ['環境構築完了', '試してみた', 'まだ'] },
    { title: 'Q&A',                type: 'paragraph' },
  ]);

  submitDemoResponses_(form3, [
    ['受講生',   'Day1 環境構築', 5, 'バッチリ',       'セットアップが思ったより簡単だった',  'VS Codeの拡張機能を入れる',        '環境構築完了', ''],
    ['メンター', 'Day1 環境構築', 4, 'だいたいOK',     '初心者向けの説明が丁寧',           'メンター用のチェックリスト作成',     '環境構築完了', 'Windowsの場合の注意点は？'],
    ['受講生',   'Day1 環境構築', 4, 'だいたいOK',     'ターミナルに触れたのが新鮮',        'コマンドを練習する',              'まだ',        ''],
    ['受講生',   'Day1 環境構築', 3, 'ちょっと難しい',   'エラーが出て焦ったが解決できた',     '復習する',                      'まだ',        'エラーが出たときの対処法まとめがほしい'],
    ['受講生',   'Day2 API連携', 5, 'バッチリ',       'APIが動いた瞬間が感動',           'ポートフォリオに組み込む',          '試してみた',   ''],
    ['メンター', 'Day2 API連携', 5, 'バッチリ',       '実装パートが充実していた',          '自分のプロジェクトにも応用',         '試してみた',   'レートリミットの対策は？'],
  ]);

  // 完了メッセージ
  ui.alert(
    'デモ作成完了!',
    '3つのGoogleフォーム + ダミー回答を作成しました!\n\n' +
    '  1. 企業研修アンケート（8件 / ヘッダー: 所属, 研修名, 満足度...）\n' +
    '  2. セミナーアンケート（7件 / ヘッダー: 会員区分, ウェビナー回, 全体の満足度...）\n' +
    '  3. ワークショップアンケート（6件 / ヘッダー: メンバー種別, 講座, 満足...）\n\n' +
    'スプレッドシートに回答シートが3つ作成されています。\n' +
    'ヘッダー名がすべて違いますが、自動検出されます!\n\n' +
    '→ 次は「初期セットアップ」を実行してください。',
    ui.ButtonSet.OK
  );

  Logger.log('=== デモフォームURL ===');
  Logger.log('1. 企業研修: ' + form1.getEditUrl());
  Logger.log('2. セミナー: ' + form2.getEditUrl());
  Logger.log('3. ワークショップ: ' + form3.getEditUrl());
}


// ====================================================================
// ヘルパー関数
// ====================================================================

/**
 * Googleフォームを作成してスプレッドシートに自動紐付け
 */
function createAndLinkForm_(title, ss, questions) {
  const form = FormApp.create(title);
  form.setCollectEmail(false);
  form.setConfirmationMessage('回答ありがとうございました！');

  // スプレッドシートに紐付け
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  questions.forEach(q => {
    switch (q.type) {
      case 'list':
        form.addListItem()
          .setTitle(q.title)
          .setChoiceValues(q.options);
        break;

      case 'scale':
        form.addScaleItem()
          .setTitle(q.title)
          .setBounds(q.low, q.high)
          .setLabels(q.lowLabel, q.highLabel);
        break;

      case 'paragraph':
        form.addParagraphTextItem()
          .setTitle(q.title);
        break;
    }
  });

  // 紐付け後、回答シートが作成されるまで待つ
  SpreadsheetApp.flush();
  Utilities.sleep(3000);

  return form;
}

/**
 * フォームにダミー回答をプログラムで投入
 */
function submitDemoResponses_(form, responses) {
  const items = form.getItems();

  responses.forEach(row => {
    const formResponse = form.createResponse();

    items.forEach((item, idx) => {
      const value = row[idx];
      if (value === '' || value === undefined || value === null) return;

      const type = item.getType();

      if (type === FormApp.ItemType.LIST) {
        formResponse.withItemResponse(
          item.asListItem().createResponse(String(value))
        );
      } else if (type === FormApp.ItemType.SCALE) {
        formResponse.withItemResponse(
          item.asScaleItem().createResponse(Number(value))
        );
      } else if (type === FormApp.ItemType.PARAGRAPH_TEXT) {
        formResponse.withItemResponse(
          item.asParagraphTextItem().createResponse(String(value))
        );
      }
    });

    formResponse.submit();
    Utilities.sleep(500); // タイムスタンプをずらす
  });
}
