const ReportParser = {
  /**
   * フォーム回答を解析し、所属別のシートオブジェクトや通知用URLを特定してconfigへ格納する
   * @param {Config} config - 共通データオブジェクト
   */
  analyzeResponses(config) {
    const mappedRes = {};

    for (const item of config.itemResponses) {
      // Googleフォームのタイトルに対応したレスポンスをオブジェクトで使用できるようにする
      mappedRes[item.getItem().getTitle()] = item.getResponse();
    }

    // -----------------------------------------------------------------
    // 【項目追加エリア】フォームに質問を追加した場合は、以下に1行追加します。
    // 例：config.newField = mappedRes[CONFIG.QUESTIONS.NEW_FIELD];
    // -----------------------------------------------------------------
    // 一覧からそれぞれの回答を取得
    // 名前を取得
    config.name = mappedRes[CONFIG.QUESTIONS.NAME] || "不明";
    // 出勤、退勤のどちらかを取得
    config.situation = mappedRes[CONFIG.QUESTIONS.SITUATION] || "";
    // -----------------------------------------------------------------

    // 現在の日を取得
    config.d = config.now.getDate();
    // 現在時刻を取得
    config.time = Utilities.formatDate(config.now, "Asia/Tokyo", "HH:mm");
  },


  /**
   * 実績票特有のタイトル行の項目名の列のインデックス(列番号 -1)を取得 (内部関数)
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - タイトル行の項目名の列のインデックスを取得するシート
   * @param {number} headerRow - ヘッダーの列番号
   * @param {number} lastCol - 最終列の番号
   */
  getRecordColumnMap(sheet, headerRow, lastCol) {
    // 見出しとなっている行の値を合算して1つの見出しとして扱う
    const headerData = sheet.getRange(headerRow, 1, CONFIG.TABLE.RECORD.ROWS_PER_TITLE, lastCol).getValues();

    const map = {};

    // 各列の値をループして、CONFIGの名前が含まれているかチェック
    for (let colIndex = 0; colIndex < lastCol; colIndex++) {
      let combinedHeader = "";
      for (let i = 0; i < CONFIG.TABLE.RECORD.ROWS_PER_TITLE; i++) {
        // 各行の同じ列の文字を足していく
        combinedHeader += headerData[i][colIndex];
      }

      // 統合されたタイトルの改行や空白をなくす
      combinedHeader = combinedHeader.replace(/\n|\s/g, "");

      for (const [key, value] of Object.entries(CONFIG.HEADER.RECORD)) {
        // タイトル名が、連結した見出しの中に含まれているか確認
        // 該当のタイトルがあれば、そのタイトルの列が何番目かインデックス番号にしてオブジェクトにする
        if (combinedHeader.includes(value)) {
          map[key] = colIndex;
        }
      }
    }
    return map;
  }
};