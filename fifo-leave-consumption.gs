/**
 * 共通エンジン：FIFOで計算だけを行う
 * *※運用事故（土日跨ぎ・複数チケット更新ミス）を防ぐため、1日単位の申請を前提とした仕様
 * @param {string} empId - 従業員のID
 * @param {Date} targetDate - 有給を消費する対象の日付
 * @param {number} consumeDays - 消費する有給の日数
 * @returns {Array<Object>} 消化対象となる有給付与データの更新リスト
 * @property {number} row - スプレッドシート上の行番号（更新用）
 * @property {number} newUsedDays - 更新後の累計消費済日数
 * @throws {Error} 有給残高が不足している場合にエラーをスローします
 */
function calculateFIFO(empId, targetDate, consumeDays) {
  targetDate = new Date(targetDate);
  //有効な付与チケットを2枚持ってくる//
  const validGrants = getValidGrants(empId, targetDate); //引数は12行目で受け取ったものと同じ//
  //calculateFIFOの第三引数を"remainingNeed"に代入//
  let remainingNeed = consumeDays;
  let updates = [];

  //getValidGrantsで取得したチケットに関してfifo処理//
  for (const grantChunk of validGrants) {
    if (remainingNeed <= 0) break;

    const amountToConsume = Math.min(grantChunk.leftDays, remainingNeed);

    // 累計消費日数に足すときも、小数点の誤差が乗っからないように丸める
    const updatedUsedDays = Math.round((grantChunk.gUsedDays + amountToConsume) * 100) / 100;

    // 更新内容をリスト化（既存のインデックスをそのまま利用）
    updates.push({
      row: grantChunk.sheetRowIndex,
      newUsedDays: grantChunk.gUsedDays + amountToConsume
    });

    remainingNeed = Math.round((remainingNeed - amountToConsume) * 100) / 100;
  }

  if (remainingNeed > 0) throw new Error("有給残高が不足しています");
  //"calculateFIFO"は"updates"を返す//
  return updates;
}
