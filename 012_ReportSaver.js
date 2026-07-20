const ReportSaver = {
  /**
   * 開始時刻を実績票に反映させる
   * @param {Config} config
   */
  startTime(config) {
    // 日付を配列で取得する関数を呼び出す
    const dateValues = this._getDateValues(config);

    for (let i = 0; i < dateValues.length; i++) {
      // 日付を取得
      const day = Number(dateValues[i]);

      if (day === config.d) {
        // 開始時間を入れる
        config.sheet.getRange(config.dataStartRow + i, config.colMap.START_TIME + 1).setValue(config.time);
        return;
      }
    }
  },


  /**
   * 終了時刻実績票へ反映
   * @param {Config} config
   */
  endTime(config) {
     // 日付を配列で取得する関数を呼び出す
    const dateValues = this._getDateValues(config);

    for (let i = 0; i < dateValues.length; i++) {
      // 日付を取得
      const day = Number(dateValues[i]);

      if (day === config.d) {
        // 開始時間を入れる
        config.sheet.getRange(config.dataStartRow + i, config.colMap.END_TIME + 1).setValue(config.time);
        return;
      }
    }
  },


  /**
   * 日付データを取得する関数 (内部関数)
   * @param {Config} config
   * @return {Object} - 日付の配列
   */
  _getDateValues(config) {
    // 実績票を取得
    const ss = SpreadsheetApp.openById(config.props[CONFIG.PROPS.RECORD_SS_ID]);

    // タイムカードを送信した利用者のシートを取得
    config.sheet = ss.getSheetByName(config.name);

    const lastCol = config.sheet.getLastColumn();

    // ヘッダーの行を取得
    const headerRow = config.sheet.createTextFinder(CONFIG.LABELS.RECORD.ACHIEVEMENTS).matchEntireCell(true).findNext().getRow();

    config.dataStartRow = headerRow + CONFIG.TABLE.RECORD.ROWS_PER_TITLE;

    // シートの見出し(タイトル)の位置を取得
    config.colMap = ReportParser.getRecordColumnMap(config.sheet, headerRow, lastCol);

    // 合計欄の行を取得
    const totalRow = config.sheet.createTextFinder(CONFIG.LABELS.RECORD.TOTAL).matchEntireCell(true).findNext().getRow();

    // 合計欄を起点として表の最終行を取得
    let tableLastRow = config.sheet.getRange(totalRow, config.colMap.DATE + 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

    if (tableLastRow <= config.dataStartRow) {
      // 表の最終行が表の最初の行より上の場合は、合計欄の行の前の行を表の最終行にする
      tableLastRow = totalRow - 1;
    }

    // 日付を配列で取得
    const dateValues = config.sheet.getRange(config.dataStartRow, config.colMap.DATE + 1, tableLastRow - (config.dataStartRow - 1), 1).getValues().flat();

    return dateValues;
  }
};