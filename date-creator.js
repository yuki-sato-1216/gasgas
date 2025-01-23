// GAS: MEMO
// - GASではオーバーロードメソッドは提供していないっぽいので、メソッド名は一意にすること

// MEMO
// - getRange(row, column)
//   - 指定した行と列のセルを取得
//     - A1セルを取得: sheet.getRange(1, 1);
// 
// - getRange(row, column, numRows, numColumns)
//   - 指定した開始行、開始列、行数、列数に基づいて範囲を取得
//     - A1からB3までの範囲を取得: sheet.getRange(1, 1, 3, 2);
// 
// - getRange(start, column, numRows)
//   - 指定した開始行、列、行数に基づいて行全 or 列全体 を取得
//     - 1行目全体を取得: sheet.getRange(1, 1, 1, sheet.getLastColumn());
//     - 1列目全体を取得: sheet.getRange(1, 1, sheet.getLastRow(), 1);
// 
// - createTextFinder(findText)(row, column)
//   - 指定したテキストを検索し、検索対象の範囲を取得
//     - var textFinder = sheet.createTextFinder("keyword");
//     - var foundCell = textFinder.findNext();


// 掲載期間計算関数
function calcDateToPublish() {
  const DAY_OF_WEEK = ["日", "月", "火", "水", "木", "金", "土"];

  // 掲載パターンenumもどき
  const PublishPattern = Object.freeze({
    Tuesday_1: Symbol("1"),
    Tuesday_2: Symbol("2"),
    Tuesday_3: Symbol("3"),
    Tuesday_4: Symbol("4"),
    Tuesday_1_3: Symbol("5"),
    Tuesday_2_4: Symbol("6"),
    EveryTuesday: Symbol("7"),
    Friday_1: Symbol("8"),
    Friday_2: Symbol("9"),
    Friday_3: Symbol("10"),
    Friday_4: Symbol("11"),
    Friday_1_3: Symbol("12"),
    Friday_2_4: Symbol("13"),
    EveryFriday: Symbol("14")
  });

  const ONE_MONTH = 1;
  // 掲載終了日をセットする最初の行（4話）
  const ROW_FOR_SETTING_END_DATE = 5
  // 掲載開始日の列
  const COLUMN_FOR_START_DATE = 2

  const INPUT_RANGE = "A2:E2";
  // Input
  const sheet4 = SpreadsheetApp.getActive().getSheetByName("シート4");
  // Output
  const sheet5 = SpreadsheetApp.getActive().getSheetByName("シート5");

  const inputDatas = sheet4
    .getRange(INPUT_RANGE)
    .getValues()
    .flat();
  // console.log("inputDatas is", inputDatas);

  // 作品ID（数値: 1～9999の値）
  const workId = typeof inputDatas[0] === "number" ? inputDatas[0] : null;
  // 作品タイトル（文字列）
  const workTitle = typeof inputDatas[1] === "string" ? inputDatas[1] : "";
  // 作品の話数（数値: 1～の値）
  const numberOfEpisodes = typeof inputDatas[2] === "number" ? inputDatas[2] : null;
  // 掲載開始日（スプレッドシートのデータ入力規則にある「有効な日付」）
  const publishStartDate = new Date(inputDatas[3]);
  // 連載パターン（文字列: スプレッドシートのデータ入力規則にある「プルダウン」にて設定）
  const patternString = typeof inputDatas[4] === "string" ? inputDatas[4] : "";
  const matchPattern = patternString.match(/^(\d+)_/);
  const patternValue = matchPattern ? parseInt(matchPattern[1], 10) : null;
  const pattern = createPublishPattern(patternValue);

  console.log("Input: workId:", workId);
  console.log("Input: workTitle:", workTitle);
  console.log("Input: numberOfEpisodes:", numberOfEpisodes);
  console.log("Input: publishStartDate:", publishStartDate);
  console.log("Input: publishPattern:", patternString);

  // 既存データがある場合は削除
  const existingDataRange = sheet5.getRange(2, 1, sheet5.getLastRow(), sheet5.getLastColumn());
  if (existingDataRange.getNumRows() > 0) {
    existingDataRange.clearContent();
  }

  // console.log("パターンごとの日付計算 - START");
  calculateDateWithEachPattern(pattern, sheet5, numberOfEpisodes, publishStartDate);
  // console.log("パターンごとの日付計算 - END");

  function calculateDateWithEachPattern(pattern, sheet5, numberOfEpisodes, publishStartDate) {
    const nthWeek = getNthWeekFromPattern(pattern);
    const patternWeekDay = getWeekDayFromPattern(pattern);
    const publishStartDateData = getNthWeekAndWeekDayFromDate(publishStartDate);
    const firstDateOfWeekdayOfMonth = getFirstWeekdayOfMonth(publishStartDate, patternWeekDay, nthWeek);
    const firstDateOfWeekdayOfMonthData = getNthWeekAndWeekDayFromDate(firstDateOfWeekdayOfMonth);
    const secondDateOfWeekdayOfMonth = getFirstWeekdayOfMonth(publishStartDate, patternWeekDay, getSecondNthWeekFromPattern(pattern));
    const secondDateOfWeekdayOfMonthData = getNthWeekAndWeekDayFromDate(secondDateOfWeekdayOfMonth);
    console.log("[calculateDateWithEachPattern]: firstDateOfWeekdayOfMonth", firstDateOfWeekdayOfMonth);
    console.log("[calculateDateWithEachPattern]: secondDateOfWeekdayOfMonth", secondDateOfWeekdayOfMonth);

    const isSameDateWithFirstNthWeekDate = publishStartDate.getTime() === firstDateOfWeekdayOfMonth.getTime();
    const isSameDateWithSecondNthWeekDate = publishStartDate.getTime() === firstDateOfWeekdayOfMonth.getTime();
    console.log("[calculateDateWithEachPattern]: isSameDateWithFirstNthWeekDate", isSameDateWithFirstNthWeekDate);
    console.log("[calculateDateWithEachPattern]: isSameDateWithSecondNthWeekDate", isSameDateWithSecondNthWeekDate);
    const isSameDateWithPattern = (isSameDateWithFirstNthWeekDate || isSameDateWithSecondNthWeekDate);
    const isLessThanOrEqualToFirstNthWeek = publishStartDateData.nthWeek < firstDateOfWeekdayOfMonthData.nthWeek;
    const isGreaterThanSecondNthWeek = publishStartDateData.nthWeek > secondDateOfWeekdayOfMonthData.nthWeek;
    const isMoveUp = publishStartDate.getTime() < firstDateOfWeekdayOfMonth.getTime();
    console.log("[calculateDateWithEachPattern]: isSameDateWithPattern[%s] :: isLessThanOrEqualToFirstNthWeek[%s] :: isMoveUp[%s]", isSameDateWithPattern, isLessThanOrEqualToFirstNthWeek, isMoveUp);
    console.log(
      "[calculateDateWithEachPattern]: nthWeek:[%s] - patternWeekDay:[%s]",
      nthWeek,
      patternWeekDay
    );

    let switchableNthWeek = -1;

    // パターンごとにループ
    for (let episode = 1; episode <= numberOfEpisodes; episode++) {
      const isLastEpisode = episode === numberOfEpisodes;
      switch (pattern) {
        case PublishPattern.Tuesday_1:
        case PublishPattern.Tuesday_2:
        case PublishPattern.Tuesday_3:
        case PublishPattern.Tuesday_4:
        case PublishPattern.Friday_1:
        case PublishPattern.Friday_2:
        case PublishPattern.Friday_3:
        case PublishPattern.Friday_4:
          handleSingleNthWeekOfMonth(
            sheet5,
            episode,
            publishStartDate,
            nthWeek,
            patternWeekDay,
            isLastEpisode
          );
          break;

        case PublishPattern.EveryTuesday:
        case PublishPattern.EveryFriday:
          handleEveryWeek(
            sheet5,
            episode,
            isMoveUp,
            publishStartDate,
            patternWeekDay,
            isLastEpisode
          );
          break;

        case PublishPattern.Tuesday_1_3:
        case PublishPattern.Friday_1_3:
          if (episode > 3) {
            if ((episode === 4 || episode === 5) && isSameDateWithPattern) {
              // 指定パターンの週のどちらかに一致している場合、対象の日付が第何週かを計算する
              if (episode === 4 && isSameDateWithFirstNthWeekDate) {
                switchableNthWeek = 1;
              } else if (episode === 5 && isSameDateWithSecondNthWeekDate) {
                switchableNthWeek = 3;
              } else {
                console.log("[calculateDateWithEachPattern]: else - not exec calculateWeekNumber: switchableNthWeek[%s]", switchableNthWeek);
              }
              console.log("[calculateDateWithEachPattern]: 1(2) or 3(4) switchableNthWeek", switchableNthWeek);
            } else if (episode === 5 && (isLessThanOrEqualToFirstNthWeek || isGreaterThanSecondNthWeek)) {
              switchableNthWeek = 1;
              console.log("[calculateDateWithEachPattern]: (isLessThanOrEqualToFirstNthWeek || isGreaterThanSecondNthWeek)[%s]", (isLessThanOrEqualToFirstNthWeek || isGreaterThanSecondNthWeek));
            } else {
              switchableNthWeek = getNthWeekFromPatternAndNthWeek(pattern, switchableNthWeek);
            }
            console.log("[calculateDateWithEachPattern]: isGreaterThanSecondNthWeek[%s]: episode[%s]: switchableNthWeek[%s]", isGreaterThanSecondNthWeek, episode, switchableNthWeek);
          }
          handleDoubleNthWeekOfMonth(
            sheet5,
            episode,
            isMoveUp,
            publishStartDate,
            switchableNthWeek,
            patternWeekDay,
            isLastEpisode
          );
          break;

        case PublishPattern.Tuesday_2_4:
        case PublishPattern.Friday_2_4:
          if (episode > 3) {
            if ((episode === 4 || episode === 5) && isSameDateWithPattern) {
              // 指定パターンの週のどちらかに一致している場合、対象の日付が第何週かを計算する
              if (episode === 4 && isSameDateWithFirstNthWeekDate) {
                switchableNthWeek = 2;
              } else if (episode === 5 && isSameDateWithSecondNthWeekDate) {
                switchableNthWeek = 4;
              } else {
                console.log("[calculateDateWithEachPattern]: else - not exec calculateWeekNumber: switchableNthWeek[%s]", switchableNthWeek);
              }
            } else if (episode === 5 && (isLessThanOrEqualToFirstNthWeek || isGreaterThanSecondNthWeek)) {
              switchableNthWeek = 2;
              console.log("[calculateDateWithEachPattern]: (isLessThanOrEqualToFirstNthWeek || isGreaterThanSecondNthWeek)[%s]", (isLessThanOrEqualToFirstNthWeek || isGreaterThanSecondNthWeek));
            } else {
              switchableNthWeek = getNthWeekFromPatternAndNthWeek(pattern, switchableNthWeek);
            }
            console.log("[calculateDateWithEachPattern]: isGreaterThanSecondNthWeek[%s]: episode[%s]: switchableNthWeek[%s]", isGreaterThanSecondNthWeek, episode, switchableNthWeek);
          }
          handleDoubleNthWeekOfMonth(
            sheet5,
            episode,
            isMoveUp,
            publishStartDate,
            switchableNthWeek,
            patternWeekDay,
            isLastEpisode
          );
          break;

        default:
          console.log("Not implemented yet ... the pattern");
          break;
      }
    }
  }

  function getNthWeekAndWeekDayFromDate(date) {
    // 引数の日付が有効な日付かどうかを確認
    if (!isValidDate(date)) {
      return 'Invalid date';
    }

    // 月初の日付を取得
    const firstDateOfMonth = new Date(date.getFullYear(), date.getMonth(), 1);

    // 月初の日曜日を取得
    const firstSundayOfMonth = new Date(firstDateOfMonth);
    while (firstSundayOfMonth.getDay() !== 0) {
      firstSundayOfMonth.setDate(firstSundayOfMonth.getDate() - 1);
    }

    // 引数の日付と月初の曜日の差分を計算
    const daysDiff = Math.floor((date - firstSundayOfMonth) / (1000 * 60 * 60 * 24));

    // 第〇週を取得
    const nthWeek = Math.floor(daysDiff / 7) + 1;
    //console.log("[getNthWeekAndWeekDayFromDate]: nthWeek:", nthWeek);

    // 曜日を取得
    const weekDay = date.getDay();

    console.log("[getNthWeekAndWeekDayFromDate]: TargetDate:", { nthWeek, weekDay });
    return {
      nthWeek,
      weekDay,
    };
  }

  // 指定した最初の曜日の日付を取得する関数
  function getFirstWeekdayOfMonth(date, weekday, nthWeek) {
    // 引数の日付が有効な日付かどうかを確認
    if (!isValidDate(date)) {
      return 'Invalid date';
    }

    // 月初の日付を取得
    const firstDateOfMonth = new Date(date.getFullYear(), date.getMonth(), 1);

    // 指定した週の最初の日付を取得
    const firstDateOfNthWeek = new Date(firstDateOfMonth);
    firstDateOfNthWeek.setDate(firstDateOfNthWeek.getDate() + (nthWeek - 1) * 7);
    console.log("[getFirstWeekdayOfMonth]: firstDateOfNthWeek:", firstDateOfNthWeek);

    // 最初の指定した曜日の日付を取得
    const firstWeekdayDate = new Date(firstDateOfNthWeek);
    while (firstWeekdayDate.getDay() !== weekday) {
      firstWeekdayDate.setDate(firstWeekdayDate.getDate() + 1);
    }

    console.log("[getFirstWeekdayOfMonth]: firstWeekdayDate:", firstWeekdayDate);
    return firstWeekdayDate;
  }

  // 日付が有効かどうかチェックする関数
  function isValidDate(date) {
    return date instanceof Date && !isNaN(date.getTime());
  }

  // 最終行の掲載開始日を取得する関数（Dateへキャストして利用する）
  function getLastStartDateFromSheet(targetSheet) {
    //console.log("[getLastStartDateFromSheet]: date:", targetSheet.getRange(targetSheet.getLastRow(), 2).getValue());
    return targetSheet.getRange(targetSheet.getLastRow(), 2).getValue();
  }

  // 数値からenumの値を取得する関数
  function createPublishPattern(number) {
    const patternName = Object.keys(PublishPattern)[number - 1];
    return PublishPattern[patternName];
  }

  // 掲載パターンに沿う2つ目の週を取得する関数
  function getSecondNthWeekFromPattern(pattern) {
    switch (pattern) {
      case PublishPattern.Tuesday_1_3:
      case PublishPattern.Friday_1_3:
        return 3;
      case PublishPattern.Tuesday_2_4:
      case PublishPattern.Friday_2_4:
        return 4;
      default:
        return -1;
    }
  }

  // 掲載パターンに沿う週を取得する関数
  // - 月2回のパターンは、最初の週を返す
  function getNthWeekFromPattern(pattern) {
    switch (pattern) {
      case PublishPattern.Tuesday_1:
      case PublishPattern.Friday_1:
      case PublishPattern.Tuesday_1_3:
      case PublishPattern.Friday_1_3:
        return 1;
      case PublishPattern.Tuesday_2:
      case PublishPattern.Friday_2:
      case PublishPattern.Tuesday_2_4:
      case PublishPattern.Friday_2_4:
        return 2;
      case PublishPattern.Tuesday_3:
      case PublishPattern.Friday_3:
        return 3;
      case PublishPattern.Tuesday_4:
      case PublishPattern.Friday_4:
        return 4;
      default:
        // EveryTuesday, EveryFridayは無視
        return -1;
    }
  }

  // 掲載パターンに沿う週を取得する関数（nthWeek指定あり）
  function getNthWeekFromPatternAndNthWeek(pattern, nthWeek) {
    switch (pattern) {
      case PublishPattern.Tuesday_1_3:
      case PublishPattern.Friday_1_3:
        return nthWeek === 1 ? 3 : 1;
      case PublishPattern.Tuesday_2_4:
      case PublishPattern.Friday_2_4:
        return nthWeek === 2 ? 4 : 2;
      default:
        console.log("[getNthWeekFromPatternAndNthWeek(pattern, nthWeek)]: default PublishPattern - nthWeek:", nthWeek);
        return 0;
    }
  }

  // 掲載パターンに沿う曜日（数値:0～6）を取得する関数
  function getWeekDayFromPattern(pattern) {
    switch (pattern) {
      case PublishPattern.Tuesday_1:
      case PublishPattern.Tuesday_2:
      case PublishPattern.Tuesday_3:
      case PublishPattern.Tuesday_4:
      case PublishPattern.Tuesday_1_3:
      case PublishPattern.Tuesday_2_4:
      case PublishPattern.EveryTuesday:
        return 2;
      case PublishPattern.Friday_1:
      case PublishPattern.Friday_2:
      case PublishPattern.Friday_3:
      case PublishPattern.Friday_4:
      case PublishPattern.Friday_1_3:
      case PublishPattern.Friday_2_4:
      case PublishPattern.EveryFriday:
        return 5;
      default:
        console.log("[getWeekDayFromPattern]: default PublishPattern");
        return 0;
    }
  }

  // 掲載開始日の前倒しを考慮し、開始日の週の指定した曜日の日付を取得する関数
  function getDateFromMoveUpStartDate(startDate, weekday) {
    // 指定した曜日の日付になるまで日付を加算する
    const nextWeekdayDate = new Date(startDate);
    while (nextWeekdayDate.getDay() !== weekday) {
      nextWeekdayDate.setDate(nextWeekdayDate.getDate() + 1);
    }
    console.log("[getDateFromMoveUpStartDate]: nextWeekdayDate", nextWeekdayDate);
    return nextWeekdayDate;
  }

  // 第〇週の〇曜日を取得する関数
  function getNthWeekdayInMonth(year, month, nthWeek, targetWeekday) {
    const firstDayOfMonth = new Date(year, month, 1);
    const dayOfWeekOfFirstDay = firstDayOfMonth.getDay();
    const offset = (targetWeekday - dayOfWeekOfFirstDay + 7) % 7;
    const firstTargetWeekday = new Date(year, month, 1 + offset + (nthWeek - 1) * 7);

    return firstTargetWeekday;
  }

  // 掲載終了日を計算する関数
  function calculateDate(isEnd, startDate, monthsToAdd, nthWeek, targetWeekday) {
    const clonedDate = new Date(startDate.getTime());
    // console.log("[calculateDate]: START);

    // 月を足す
    clonedDate.setMonth(clonedDate.getMonth() + monthsToAdd);

    // 第〇週の〇曜日を取得
    const targetDate = getNthWeekdayInMonth(clonedDate.getFullYear(), clonedDate.getMonth(), nthWeek, targetWeekday);
    // console.log("[calculateDate]: getNthWeekdayInMonth >>> targetDate", targetDate);

    if (isEnd) {
      // 掲載終了日は、開始日から指定された月数（monthsToAdd）後の第〇週の〇曜日として再計算
      const endDate = getNthWeekdayInMonth(targetDate.getFullYear(), targetDate.getMonth(), nthWeek, targetWeekday);
      // console.log("[calculateDate]: getNthWeekdayInMonth >>> endDate", endDate);
      return endDate;
    }

    return targetDate;
  }

  // 週の値を確認し、追加する月数を取得する
  function getMonthToAddFromNthWeek(nthWeek) {
    return (nthWeek === 1 || nthWeek === 2) ? 1 : 0;
  }

  // 毎週の日付を計算する関数
  function calculateNextWeekDate(startDate) {
    const clonedDate = new Date(startDate.getTime());
    // console.log("[calculateNextWeekDate]: clonedDate", clonedDate);
    clonedDate.setDate(clonedDate.getDate() + 7);
    return clonedDate;
  }

  // 日付を指定のフォーマットに整形する関数（yyyy/MM/dd HH:mm:ss）
  function formatDateWithTime(date) {
    const clonedDate = new Date(date.getTime());
    // console.log('[formatDateWithTime]: Before calculation: ', clonedDate);

    clonedDate.setMonth(clonedDate.getMonth());

    const year = clonedDate.getFullYear();
    const month = (clonedDate.getMonth() + 1).toString().padStart(2, '0');
    const day = clonedDate.getDate().toString().padStart(2, '0');

    const formattedDate = `${year}-${month}-${day} 12:00:00`;

    // 不正な日付
    if (isNaN(clonedDate.getTime())) {
      return '[formatDateWithTime]: Invalid Date';
    }
    // console.log('[formatDateWithTime]: After formattedDate: ', formattedDate);

    return formattedDate;
  }

  function getPreviousStartDate(targetSheet, tmpRowData) {
    //console.log("[getPreviousStartDate]: tmpRowData:", tmpRowData);
    return new Date(
      tmpRowData !== undefined
        ? tmpRowData
        : getLastStartDateFromSheet(targetSheet)
    );
  }

  // 4話以降の掲載終了日をセットする関数
  //   - targetSheet: シート
  //   - additionalTmpRowDatas: 追加3話分の日付リスト
  function setEndDateToAllEpisodes(
    targetSheet,
    additionalTmpRowDatas
  ) {
    // 終了日セット用日付データ作成関数
    // - 行データ取得
    // - 行データの値を取得
    // - 一次元配列に修正
    // - 何故か空文字が入るので、念のためフィルターする
    // - Dateにした後フォーマットする
    // - 前の処理で作成した終了日調整用の3話分を追加
    const allStartDateList = sheet5
      .getRange(
        ROW_FOR_SETTING_END_DATE,
        COLUMN_FOR_START_DATE,
        targetSheet.getLastRow()
      )
      .getValues()
      .flat()
      .filter(value => value !== '' && value !== undefined)
      .map(rowData => formatDateWithTime(new Date(rowData)))
      .concat(additionalTmpRowDatas);

    let endDateRowData;
    for (let i = 0; i < allStartDateList.length; i++) {
      const baseEndDate = allStartDateList[i + 3];
      // 終了日のみをセット
      const targetRange = targetSheet.getRange(ROW_FOR_SETTING_END_DATE + i, 3);
      endDateRowData = [[baseEndDate]];
      targetRange.setValues(endDateRowData);
    }
  }

  // パターン: 毎月第〇 〇曜日更新
  // - nthWeek: 1～4
  // - targetWeekday: 0～6で指定（火2 or 金5）
  function handleSingleNthWeekOfMonth(
    sheet5,
    episode,
    publishStartDate,
    nthWeek,
    targetWeekday,
    isLastEpisode
  ) {
    console.log("[handleSingleNthWeekOfMonth]: パターン: 毎月第[%s]-[%s]曜日更新 - 第[%s]話 - 処理中...", nthWeek, DAY_OF_WEEK[targetWeekday], episode);
    // Output
    let rowData;

    if (episode > 4) {
      const previousStartDate = new Date(getLastStartDateFromSheet(sheet5));
      const newStartDate = calculateDate(false, previousStartDate, ONE_MONTH, nthWeek, targetWeekday);
      rowData = [
        [
          `${episode}話`,
          formatDateWithTime(newStartDate),
          ""
        ],
      ];
    } else {
      // 1～3話は掲載開始日そのまま
      rowData = [
        [
          `${episode}話`,
          formatDateWithTime(publishStartDate),
          episode < 4 ? "なし" : "",
        ],
      ];
    }

    const targetRange = sheet5.getRange(sheet5.getLastRow() + 1, 1, 1, 3);
    targetRange.setValues(rowData);

    if (!isLastEpisode) { return; }

    /* ========= 終了日調整用の3話を追加する処理 ========= */
    {
      // 最後話数から+3話分を作成しておく
      // pushするためにArrayで初期化
      let additionalTmpRowDatas = Array();
      let tmpRowData;
      for (let i = 1; i <= 3; i++) {
        const previousStartDate = getPreviousStartDate(sheet5, tmpRowData);
        const newStartDate = formatDateWithTime(
          calculateDate(false, previousStartDate, ONE_MONTH, nthWeek, targetWeekday)
        );
        additionalTmpRowDatas.push(newStartDate);
        tmpRowData = additionalTmpRowDatas[i - 1];
      }

      if (sheet5.getLastRow() < 2) { return; }
      console.log("[handleSingleNthWeekOfMonth]: Setup End Date - START");

      // 終了日セット用日付データ作成
      setEndDateToAllEpisodes(sheet5, additionalTmpRowDatas);

      console.log("[handleSingleNthWeekOfMonth]: Setup End Date - END");
    }
  }

  // パターン: 毎月2回更新（指定曜日）
  // - nthWeek: 1～4
  // - targetWeekday: 0～6で指定（火2 or 金5）
  function handleDoubleNthWeekOfMonth(
    sheet5,
    episode,
    isMoveUp,
    publishStartDate,
    nthWeek,
    targetWeekday,
    isLastEpisode
  ) {
    console.log("[handleDoubleNthWeekOfMonth]: パターン: 毎月第[%s]-[%s]曜日更新 - 第[%s]話 - 処理中...", nthWeek, DAY_OF_WEEK[targetWeekday], episode);
    // Output
    let rowData;

    if (episode > 4) {
      const previousStartDate = new Date(getLastStartDateFromSheet(sheet5));
      const monthToAdd = getMonthToAddFromNthWeek(nthWeek);
      console.log("[handleDoubleNthWeekOfMonth]: previousStartDate[%s]:", previousStartDate);
      console.log("[handleDoubleNthWeekOfMonth]: monthToAdd[%s]:", monthToAdd);
      let newStartDate;
      if (episode === 5 && isMoveUp) {
        newStartDate = getFirstWeekdayOfMonth(publishStartDate, targetWeekday, nthWeek);
      } else {
        newStartDate = calculateDate(false, previousStartDate, monthToAdd, nthWeek, targetWeekday);
      }
      rowData = [
        [
          `${episode}話`,
          formatDateWithTime(newStartDate),
          ""
        ],
      ];
    } else {
      // 1～3話は掲載開始日そのまま
      rowData = [
        [
          `${episode}話`,
          formatDateWithTime(publishStartDate),
          episode < 4 ? "なし" : "",
        ],
      ];
    }

    const targetRange = sheet5.getRange(sheet5.getLastRow() + 1, 1, 1, 3);
    targetRange.setValues(rowData);

    if (!isLastEpisode) { return; }

    /* ========= 終了日調整用の3話を追加する処理 ========= */
    {
      // 最後話数から+3話分を作成しておく
      // pushするためにArrayで初期化
      let additionalTmpRowDatas = Array();
      let tmpRowData;
      let additionalNthWeek = getNthWeekFromPatternAndNthWeek(pattern, nthWeek);
      for (let i = 1; i <= 3; i++) {
        const previousStartDate = getPreviousStartDate(sheet5, tmpRowData);

        // 1、2週目の場合、翌月へ進める
        const monthToAdd = getMonthToAddFromNthWeek(additionalNthWeek);
        const newStartDate = formatDateWithTime(
          calculateDate(false, previousStartDate, monthToAdd, additionalNthWeek, targetWeekday)
        );
        additionalTmpRowDatas.push(newStartDate);
        tmpRowData = additionalTmpRowDatas[i - 1];

        additionalNthWeek = getNthWeekFromPatternAndNthWeek(pattern, additionalNthWeek);
      }

      if (sheet5.getLastRow() < 2) { return; }
      console.log("[handleDoubleNthWeekOfMonth]: Setup End Date - START");

      // 終了日セット用日付データ作成
      setEndDateToAllEpisodes(sheet5, additionalTmpRowDatas);

      console.log("[handleDoubleNthWeekOfMonth]: Setup End Date - END");
    }
  }

  function handleEveryWeek(
    sheet5,
    episode,
    isMoveUp,
    publishStartDate,
    targetWeekday,
    isLastEpisode
  ) {
    console.log("[handleEveryWeek]: パターン: 毎週-[%s]曜日更新 - 第[%s]話 - 処理中...", DAY_OF_WEEK[targetWeekday], episode);
    // Output
    let rowData;

    if (episode > 4) {
      const previousStartDate = new Date(getLastStartDateFromSheet(sheet5));
      let newStartDate = calculateNextWeekDate(previousStartDate);
      if (episode === 5 && isMoveUp) {
        newStartDate = getDateFromMoveUpStartDate(publishStartDate, targetWeekday);
      } else {
        newStartDate = calculateNextWeekDate(previousStartDate);
      }
      rowData = [
        [
          `${episode}話`,
          formatDateWithTime(newStartDate),
          ""
        ],
      ];
    } else {
      // 1～3話は掲載開始日そのまま
      rowData = [
        [
          `${episode}話`,
          formatDateWithTime(publishStartDate),
          episode < 4 ? "なし" : "",
        ],
      ];
    }

    const targetRange = sheet5.getRange(sheet5.getLastRow() + 1, 1, 1, 3);
    targetRange.setValues(rowData);

    if (!isLastEpisode) { return; }

    /* ========= 終了日調整用の3話を追加する処理 ========= */
    {
      // 最後話数から+3話分を作成しておく
      // pushするためにArrayで初期化
      let additionalTmpRowDatas = Array();
      let tmpRowData;
      for (let i = 1; i <= 3; i++) {
        const previousStartDate = getPreviousStartDate(sheet5, tmpRowData);
        const newStartDate = formatDateWithTime(
          calculateNextWeekDate(previousStartDate)
        );
        additionalTmpRowDatas.push(newStartDate);
        tmpRowData = additionalTmpRowDatas[i - 1];
      }

      if (sheet5.getLastRow() < 2) { return; }
      console.log("[handleDoubleNthWeekOfMonth]: Setup End Date - START");

      // 終了日セット用日付データ作成
      setEndDateToAllEpisodes(sheet5, additionalTmpRowDatas);

      console.log("[handleDoubleNthWeekOfMonth]: Setup End Date - END");
    }
  }
}