function nightlyLeaveGrantBatch() {
  const CONFIG = {
    ADMIN_EMAIL: "admin@example.com",
    TIME_ZONE: "Asia/Tokyo"
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const grantRecords = [];

  // データの読み取り部分はロック外でもOK（重い処理でなければ）
  const empSheet = ss.getSheetByName("社員マスタ");
  const empRawData = empSheet.getDataRange().getValues().slice(1);
  const employeeData = empRawData.map(row => ({
    id: row[0],
    name: row[1],
    hireDate: row[2],
    status: row[3]
  }));

  const ruleSheet = ss.getSheetByName("付与ルールマスタ");
  const ruleRawData = ruleSheet.getDataRange().getValues().slice(1);
  const ruleMaster = ruleRawData.map(row => ({
    months: Number(row[0]),
    grantDays: Number(row[1])
  }));

  employeeData.forEach(employee => {
    const employeeId = employee.id;
    const hireDate = new Date(employee.hireDate);

    if (employee.status !== '在籍') return;

    const monthsDiff = (today.getFullYear() - hireDate.getFullYear()) * 12 + (today.getMonth() - hireDate.getMonth());
    const isAnniversary = today.getDate() === hireDate.getDate();

    if (isAnniversary) {
      const rule = ruleMaster.find(r => r.months === monthsDiff);

      if (rule) {
        const expireDate = new Date(today.getFullYear() + 2, today.getMonth(), today.getDate() - 1);
        const strToday = Utilities.formatDate(today, CONFIG.TIME_ZONE, "yyyy/MM/dd");
        const strExpire = Utilities.formatDate(expireDate, CONFIG.TIME_ZONE, "yyyy/MM/dd");
        const formattedDate = Utilities.formatDate(today, CONFIG.TIME_ZONE, "yyyyMMdd");
        const GrantId = `${employeeId}_${formattedDate}`;

        grantRecords.push([GrantId, employeeId, strToday, strExpire, rule.grantDays, 0]);
      }
    }
  });

  // 対象者がいないならここで終了
  if (grantRecords.length === 0) return;

  // --- 💡 書き込みの瞬間だけロックを取得する ---
  const lock = LockService.getScriptLock();
  try {
    const hasLock = lock.tryLock(10000);
    if (!hasLock) {
      console.warn("他のプロセスが実行中のため、自動付与バッチはスキップされました。");
      return;
    }

    try {
      const grantSheet = ss.getSheetByName("付与履歴");
      const lastRow = grantSheet.getLastRow();
      grantSheet.getRange(lastRow + 1, 1, grantRecords.length, grantRecords[0].length).setValues(grantRecords);
      console.log(`${grantRecords.length}件の有休を自動付与しました。`);
    } catch (e) {
      const subject = "【要対応】有休自動付与処理でエラー発生";
      const body = `自動付与処理中にエラーが発生しました。\n\n詳細: ${e.message}\n\n至急、スプレッドシートの「付与履歴」シートを確認してください。`;

      MailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, body);
      console.error(`付与エラー: ${e.message}`);
      throw e; // 書き込み失敗時はしっかりエラーを投げる
    }

  } catch (outerErr) {
    // ロック取得失敗や、書き込みエラーをここで最終キャッチ
    throw outerErr;
  } finally {
    // どんな結末であれ、絶対にロックを解放する
    lock.releaseLock();
  }
}
