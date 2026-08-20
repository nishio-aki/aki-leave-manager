/**
 * 申請受付用（フォーム送信で発火）
 * 有給申請受付～ステータス"承認待ち"で貼り付け～上司へ承認フォーム送信迄
 */
function onFormSubmitHandler(e) {

  const sheetName = e.range.getSheet().getName();
  if (sheetName !== "有休申請") {
    return;
  }
  
  const CONFIG = {
    FORM_BASE_URL: "https://docs.google.com/forms/d/e/1FAIpQLSeasrMXfIZgIRQwCAk04w8YJJIbdEKgD-D46UGfaPe-P-g8Bg/viewform",
    BOSS_MAIL: "boss@gmail.com",
    ADMIN_EMAIL: "admin@gmail.com"
  };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("社員マスタ");

  const empId = e.namedValues['社員ID'][0];
  const targetDate = e.namedValues['取得日'][0];
  const type = e.namedValues['取得区分'][0];
  const email = e.namedValues['メールアドレス'] ? e.namedValues['メールアドレス'][0] : null;
  
  let empName = "不明";
  const masterData = masterSheet.getDataRange().getValues();

  for (let i = 0; i < masterData.length; i++) {
    if (String(masterData[i][0]) === String(empId)) {
      empName = masterData[i][1];
      break;
    }
  }

  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyyMMdd-HHmmss");
  const requestId = "G-" + timestamp;

// --- 書き込みの瞬間だけロックをかける ---
  const lock = LockService.getScriptLock();
  let errorToThrow = null;

  try {
    const hasLock = lock.tryLock(10000);
    if (!hasLock) {
      throw new Error("他の申請処理と競合したため、タイムアウトしました。もう一度お試しください。");
    }

    const logSheet = ss.getSheetByName("有休申請");
    
    // 💡 【スマートな修正】
    // フォームが自動で追加した「一番下の行」の行番号を取得する
    const lastRow = logSheet.getLastRow();

    // その行の「6列目（申請ID）」と「7列目（ステータス）」に、値を追加で書き込む
    logSheet.getRange(lastRow, 6).setValue(requestId);
    logSheet.getRange(lastRow, 7).setValue("承認待ち");

  } catch (err) {
    const subject = "【要対応】有休申請の受付処理でエラー発生";
    MailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, `詳細: ${err.message}`);
    console.error(`受付エラー: ${err.message}`);

    errorToThrow = err; // 💡 エラーを退避

  } finally {
    lock.releaseLock();
  }

  // ロック解放後にエラーがあれば投げる
  if (errorToThrow) {
    throw errorToThrow;
  }

  // 承認フォームへのURL生成 ＆ 上司へメール送信
  const prefilledUrl = CONFIG.FORM_BASE_URL + "?usp=pp_url&entry.384500102=" + requestId;

  MailApp.sendEmail(
    CONFIG.BOSS_MAIL,
    `【有休承認依頼】${empName}さんからの申請`,
    `以下の申請の承認をお願いします。\n\n申請者: ${empName} 様（ID: ${empId}）\n取得日: ${targetDate}（${type}）\n申請ID: ${requestId}\n申請者メール: ${email}\n\n▼こちらのURLから承認、または調整を行ってください:\n${prefilledUrl}`
  );
}
