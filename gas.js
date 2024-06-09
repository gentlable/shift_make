function createShift() {
  // 対象の年と月を取得
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var creationSheet = ss.getSheetByName('作成')
  var year = creationSheet.getRange('B2').getValue()
  var month = creationSheet.getRange('C2').getValue()

  // シフト作成対象のシートを取得
  var targetSheet = ss.getSheetByName(month + '月')
  if (!targetSheet) {
    throw new Error(month + '月シートが見つかりません。')
  }

  // 対象月の確定済チェック
  if (targetSheet.getRange('F1').getValue() === '確定済') {
    Logger.log(month + '月のシフトは確定済のため、処理を中止します。')
    return
  }

  // 前月のシートを取得
  var previousMonth = getPreviousMonth(month)
  var previousMonthSheet = ss.getSheetByName(previousMonth + '月')
  if (!previousMonthSheet) {
    throw new Error(previousMonth + '月シートが見つかりません。')
  }

  // 前月の最終日のシフトを取得して転記
  var lastRow = previousMonthSheet.getLastRow()
  var lastDayShift = {
    day: previousMonthSheet.getRange('C' + lastRow).getValue(),
    night: previousMonthSheet.getRange('D' + lastRow).getValue(),
  }
  targetSheet.getRange('C2').setValue(lastDayShift.day)
  targetSheet.getRange('D2').setValue(lastDayShift.night)

  // 日付と曜日を入力
  initializeMonthSheet(targetSheet, year, month)

  // 対象月のシートを初期化（C列とD列のデータをクリア）
  targetSheet.getRange('C3:D' + targetSheet.getLastRow()).clearContent()

  var dateRange = targetSheet.getRange('A3:A' + targetSheet.getLastRow())
  var dates = dateRange.getValues()

  // 当番回数シートを取得
  var dutyCountSheet = ss.getSheetByName('当番回数')
  if (!dutyCountSheet) {
    throw new Error('当番回数シートが見つかりません。')
  }

  // 祝日一覧を取得
  var holidaySheet = ss.getSheetByName('祝日一覧')
  if (!holidaySheet) {
    throw new Error('祝日一覧シートが見つかりません。')
  }
  var holidays = getHolidays(holidaySheet)

  // 担当者のリストを取得
  var memberList = dutyCountSheet
    .getRange('A2:A' + dutyCountSheet.getLastRow())
    .getValues()
    .flat()

  // 休み希望シートを取得
  var vacationSheet = ss.getSheetByName('休み希望')
  if (!vacationSheet) {
    throw new Error('休み希望シートが見つかりません。')
  }
  var vacationRequests = getVacationRequests(vacationSheet, month)

  // 前月の最終日のシフトを取得
  var previousDayDuty = {
    day: lastDayShift.day,
    night: lastDayShift.night,
  }

  // シフト作成ロジック
  for (var i = 0; i < dates.length; i++) {
    var date = new Date(dates[i][0])
    var dayOfWeek = date.getDay()
    var isHoliday = holidays.includes(
      Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd')
    )

    // 背景色の設定
    if (dayOfWeek === 0 || isHoliday) {
      // 日曜日または祝日
      targetSheet
        .getRange('A' + (i + 3) + ':D' + (i + 3))
        .setBackground('#FFE6E6') // 薄い桃色
    } else if (dayOfWeek === 6) {
      // 土曜日
      targetSheet
        .getRange('A' + (i + 3) + ':D' + (i + 3))
        .setBackground('#E6F2FF') // 薄い水色
    } else {
      // 平日
      targetSheet.getRange('A' + (i + 3) + ':D' + (i + 3)).setBackground(null) // 背景色をリセット
    }

    // 日直の割り当て（土日祝のみ）
    if (dayOfWeek === 0 || dayOfWeek === 6 || isHoliday) {
      var dayShiftIndex = getNextDutyIndex(
        dutyCountSheet,
        memberList,
        previousDayDuty,
        date,
        vacationRequests
      )
      if (dayShiftIndex === -1) {
        throw new Error('適切な日直担当者が見つかりません。')
      }
      targetSheet.getRange('C' + (i + 3)).setValue(memberList[dayShiftIndex])
      previousDayDuty.day = memberList[dayShiftIndex]
    } else {
      targetSheet.getRange('C' + (i + 3)).setValue('')
    }

    // 当直の割り当て（毎日）
    var nightShiftIndex = getNextDutyIndex(
      dutyCountSheet,
      memberList,
      previousDayDuty,
      date,
      vacationRequests
    )
    if (nightShiftIndex === -1) {
      throw new Error('適切な当直担当者が見つかりません。')
    }
    targetSheet.getRange('D' + (i + 3)).setValue(memberList[nightShiftIndex])
    previousDayDuty.night = memberList[nightShiftIndex]
  }
}

/**
 * 前月を取得する関数
 * @param {string} month - 現在の月（例: "7月"）
 * @return {string} 前月（例: "6月"）
 */
function getPreviousMonth(month) {
  var months = [
    '1月',
    '2月',
    '3月',
    '4月',
    '5月',
    '6月',
    '7月',
    '8月',
    '9月',
    '10月',
    '11月',
    '12月',
  ]
  var index = months.indexOf(month + '月')
  if (index === -1) {
    throw new Error('月のフォーマットが正しくありません。')
  }
  return months[(index + 11) % 12].replace('月', '') // 月を削除
}

/**
 * 対象月のシートに日付と曜日を初期化する関数
 * @param {Sheet} sheet - 対象月のシート
 * @param {number} year - 年
 * @param {string} month - 月（例: "7月"）
 */
function initializeMonthSheet(sheet, year, month) {
  var startDate = new Date(year, month - 1, 1)
  var endDate = new Date(year, month, 0) // 月の最終日
  var rowIndex = 3 // シフトは3行目から開始

  for (
    var date = startDate;
    date <= endDate;
    date.setDate(date.getDate() + 1)
  ) {
    sheet.getRange('A' + rowIndex).setValue(new Date(date)) // 日付を設定
    sheet
      .getRange('B' + rowIndex)
      .setValue(['日', '月', '火', '水', '木', '金', '土'][date.getDay()]) // 曜日を設定
    rowIndex++
  }
}

/**
 * 祝日一覧を取得する関数
 * @param {Sheet} holidaySheet - 祝日一覧シート
 * @return {Array} 祝日の配列
 */
function getHolidays(holidaySheet) {
  var holidays = holidaySheet
    .getRange('A2:A' + holidaySheet.getLastRow())
    .getValues()
    .flat()
  return holidays.map(function (date) {
    return Utilities.formatDate(
      new Date(date),
      Session.getScriptTimeZone(),
      'yyyy/MM/dd'
    )
  })
}

/**
 * 休み希望を取得する関数
 * @param {Sheet} vacationSheet - 休み希望シート
 * @param {string} month - 月（例: "7月"）
 * @return {Object} 休み希望のオブジェクト
 */
function getVacationRequests(vacationSheet, month) {
  var requests = {}
  var monthColumn = getMonthColumn(month)
  var values = vacationSheet
    .getRange('A2:A' + vacationSheet.getLastRow())
    .getValues()
  var vacationData = vacationSheet
    .getRange('B2:L' + vacationSheet.getLastRow())
    .getValues()

  for (var i = 0; i < values.length; i++) {
    var name = values[i][0]
    var requestDays = vacationData[i][monthColumn]
    if (requestDays) {
      requests[name] = requestDays.split(/[\s,、]+/).map(function (day) {
        return parseInt(day, 10)
      })
    }
  }
  return requests
}

/**
 * 月のインデックスを取得する関数
 * @param {string} month - 月（例: "7月"）
 * @return {number} 月のインデックス
 */
function getMonthColumn(month) {
  var months = [
    '4月',
    '5月',
    '6月',
    '7月',
    '8月',
    '9月',
    '10月',
    '11月',
    '12月',
    '1月',
    '2月',
    '3月',
  ]
  return months.indexOf(month + '月')
}

/**
 * 次の当番を決定する関数
 * @param {Sheet} dutyCountSheet - 当番回数シート
 * @param {Array} memberList - メンバー一覧配列
 * @param {Object} previousDayDuty - 前日の当番（日直と当直のオブジェクト）
 * @param {Date} date - 対象の日付
 * @param {Object} vacationRequests - 休み希望のオブジェクト
 * @return {number} 次の当番のインデックス、条件に合うメンバーがいない場合は-1を返す
 */
function getNextDutyIndex(
  dutyCountSheet,
  memberList,
  previousDayDuty,
  date,
  vacationRequests
) {
  var counts = memberList.map(function (name) {
    var range = dutyCountSheet.getRange('B' + (memberList.indexOf(name) + 2))
    return range.getValue()
  })

  var minCount = Math.min.apply(null, counts)
  var candidates = []

  for (var i = 0; i < memberList.length; i++) {
    var name = memberList[i]
    if (
      counts[i] === minCount &&
      name !== previousDayDuty.day &&
      name !== previousDayDuty.night &&
      !isOnVacation(name, date, vacationRequests)
    ) {
      candidates.push(i)
    }
  }

  if (candidates.length === 0) {
    for (var j = 0; j < memberList.length; j++) {
      var name = memberList[j]
      if (
        name !== previousDayDuty.day &&
        name !== previousDayDuty.night &&
        !isOnVacation(name, date, vacationRequests)
      ) {
        candidates.push(j)
      }
    }
  }

  if (candidates.length === 0) {
    return -1 // 条件に合うメンバーがいない場合
  }

  return candidates[Math.floor(Math.random() * candidates.length)]
}

/**
 * 休み希望を確認する関数
 * @param {string} name - 担当者の名前
 * @param {Date} date - 対象の日付
 * @param {Object} vacationRequests - 休み希望のオブジェクト
 * @return {boolean} 休み希望かどうか
 */
function isOnVacation(name, date, vacationRequests) {
  if (vacationRequests[name]) {
    var day = date.getDate()
    if (vacationRequests[name].includes(day))
      return vacationRequests[name].includes(day)
  }
  return false
}
