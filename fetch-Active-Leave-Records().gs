/**
 * 対象社員の現在有効な有給データを「付与履歴」から抽出して古い順にソートして返す
 * ※JSDoc-各変数の内容を明確にする注釈
 * @param {string} employeeId - 社員ID
 * @param {Date} targetDate - 判定の基準日（申請日や照会当日）
 * @return {Array} 有効な有給データの配列
 */
//同様の形で変数を入れる窓口関数//
function getValidGrants(employeeId, targetDate) {
  targetDate = new Date(targetDate);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grantSheet = ss.getSheetByName("付与履歴");
  if (!grantSheet) return [];

  const grantRows = grantSheet.getDataRange().getValues().slice(1); // 見出し列を除外
  //付与履歴:付与ID|社員ID|付与日|有効期限|付与日数|消費済日数//

  // 1. まず条件に合うものだけを「選別」して、同時に「扱いやすい形」に変換する
  return grantRows
    .map((row, index) => {
      // row[2]などの「謎の数字」を、名前付きのデータに変換

      // ★ 1. 計算用に、ここで一度変数として定義する
      const gTotalDays = Number(row[4]);
      const gUsedDays = Number(row[5]);

      return {
        sheetRowIndex: index + 2,
        empId: row[1],
        grantDate: new Date(row[2]),
        expireDate: new Date(row[3]),
        gTotalDays: gTotalDays, // ★ 2. 上で定義した変数を入れる
        gUsedDays: gUsedDays,   // ★ 3. 同上
        leftDays: gTotalDays - gUsedDays // ★ 4. エラーにならずに計算できる！
      };
    })
    
    .filter(g => {
      // 2. ここで条件判定（読みやすい！）
      const isTargetEmployee = (g.empId === employeeId);//employeeIdは関数に入れる値//
      const isWithinPeriod = (targetDate >= g.grantDate && targetDate <= g.expireDate);
      const hasBalance = (g.leftDays > 0.01);//拡張用お守り"0"→"0.01"//
      //returnの後の条件を満たす配列を返してくれる//
      return isTargetEmployee && isWithinPeriod && hasBalance;
    })
    .sort((a, b) => a.grantDate - b.grantDate); // 最後に古い順にソート
}
