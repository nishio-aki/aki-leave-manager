/**
 * 申請受付用（フォーム送信で発火）
 * 有給申請受付～ステータス"承認待ち"で貼り付け～上司へ承認フォーム送信迄
 */
function onFormSubmitHandler(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("取得履歴");
  const masterSheet = ss.getSheetByName("社員マスタ"); // ★社員マスタシートを取得

  // フォームの回答内容を取得
  const empId = e.namedValues['社員ID'][0];
  const targetDate = e.namedValues['取得日'][0];
  const type = e.namedValues['取得区分'][0];
  const email = e.userEmail; // 自動収集されたメールアドレス

  // ★社員マスタから社員ID（A列）を検索して名前（B列）を取得する処理
  let empName = "不明"; // 見つからなかった場合の初期値
  const masterData = masterSheet.getDataRange().getValues(); // マスタの全データ取得

  for (let i = 0; i < masterData.length; i++) {
    // A列(インデックス[0])の社員IDと一致するか（文字列として比較）
    if (String(masterData[i][0]) === String(empId)) {
      empName = masterData[i][1]; // B列(インデックス[1])の氏名を取得
      break;
    }
  }

  // 日時取得と重複しない取得IDの生成
  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyyMMdd-HHmmss");
  const requestId = "G-" + timestamp;

  // A:タイムスタンプ, B:社員ID, C:取得日, D:取得区分, E:メールアドレス, F:取得ID, G:ステータス
  logSheet.appendRow([now, empId, targetDate, type, email, requestId, "承認待ち"]);

  // 承認フォームへのURL生成
  const formBaseUrl = "https://docs.google.com/forms/d/e/13qr4ek6SvJogwMbZ9E_HWUT_JOfCcgfNHcwDzQ6id6U/viewform";
  const prefilledUrl = formBaseUrl + "?usp=pp_url&entry.384500102=" + requestId;

  // ★上司へメール送信（件名と本文に名前をプラス！）
  //申請IDは連投対応・検索キー//
  MailApp.sendEmail(
    "boss@example.com",
    `【有給承認依頼】${empName}さんからの申請`,
    `以下の申請の承認をお願いします。

申請者: ${empName} 様（ID: ${empId}）
取得日: ${targetDate}（${type}）
申請ID: ${requestId}
申請者メール: ${email}

▼こちらのURLから承認、または調整を行ってください:
${prefilledUrl}`
  );
}


/**
 * 承認用（承認フォーム送信で発火）
 * （ステータス）：（対応）
 * 承認：申請IDをキーにして当該行のステータスを「承認済み」に更新、消化日数を反映の上、申請者と事務にメール通知
 * 時季変更を打診する：変更を打診する理由と、避けてほしい期間をメールで返送
 */
function onApprovalFormSubmit(e) {
  const requestId = e.namedValues['申請ID'][0];
  const status = e.namedValues['承認判定'][0];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("取得履歴");
  const grantSheet = ss.getSheetByName("付与履歴");

  const logData = logSheet.getDataRange().getValues();
  let targetRow = -1;
  let currentStatus = "";

  for (let i = 0; i < logData.length; i++) {
    if (logData[i][5] === requestId) {
      targetRow = i + 1;
      currentStatus = logData[i][6];
      break;
    }
  }

  if (targetRow === -1) {
    throw new Error("該当する取得IDが見つかりませんでした: " + requestId);
  }

  if (currentStatus === "承認済み" || currentStatus === "調整依頼中") {
    Logger.log(`[警告] 申請ID: ${requestId} はすでに「${currentStatus}」のため、二重処理を防止しました。`);
    return;
  }

  const applicantEmail = logSheet.getRange(targetRow, 5).getValue();
  const empId = logSheet.getRange(targetRow, 2).getValue();
  const rawDate = logSheet.getRange(targetRow, 3).getValue();
  const formattedDate = Utilities.formatDate(new Date(rawDate), "Asia/Tokyo", "yyyy/MM/dd");
  const type = logSheet.getRange(targetRow, 4).getValue();

  if (status === "承認する") {
    try {
      logSheet.getRange(targetRow, 7).setValue("承認済み");

      const date = new Date(rawDate);
      const consumeDays = (type === "全休") ? 1.0 : 0.5;

      const updates = calculateFIFO(empId, date, consumeDays);
      updates.forEach(u => grantSheet.getRange(u.row, 6).setValue(u.newUsedDays));

      if (applicantEmail) {
        MailApp.sendEmail(
          applicantEmail,
          "【承認完了】有給休暇の申請が承認されました",
          `申請者様\n\n以下の有給休暇申請が【承認】されましたのでお知らせいたします。\n\n・取得予定日: ${formattedDate}（${type}）\n・申請ID: ${requestId}\n\n※本メールはシステムからの自動送信です。`
        );
      }

      MailApp.sendEmail(
        "admin+clerk@example.com",
        `【有給確定】${empId}さんの有給申請が承認されました`,
        `事務担当者様\n\n以下の有給申請が承認され、消化処理が正常に完了しました。\n\n・社員ID: ${empId}\n・取得日: ${formattedDate}\n・区分: ${type}\n・申請ID: ${requestId}`
      );
    } catch (err) {
      logSheet.getRange(targetRow, 7).setValue("エラー:" + err.message);
      // ★エラー発生時に事務担当者へ通知メールを送信
      MailApp.sendEmail(
        "admin+clerk@example.com",
        `【要確認・有給エラー】申請ID: ${requestId}`,
        `事務担当者様\n\n有給承認の処理中にエラーが発生しました。内容を確認してください。\n\n・申請ID: ${requestId}\n・社員ID: ${empId}\n・エラー内容: ${err.message}`
      );
      throw err;
    }
  } else if (status === "時季変更を打診する") {
    logSheet.getRange(targetRow, 7).setValue("調整依頼中");

    // 打診フォーム側の質問3〜5の回答を取得（フォームの質問項目名に合わせる）
    const reason = e.namedValues['時季変更を打診する理由'] ? e.namedValues['時季変更を打診する理由'][0] : "記載なし";
    const avoidStart = e.namedValues['有給取得を避けてほしい期間（開始日）'] ? e.namedValues['有給取得を避けてほしい期間（開始日）'][0] : "";
    const avoidEnd = e.namedValues['有給取得を避けてほしい期間（終了日）'] ? e.namedValues['有給取得を避けてほしい期間（終了日）'][0] : "";

    const avoidPeriodText = (avoidStart && avoidEnd) ? `・避けてほしい期間: ${avoidStart} 〜 ${avoidEnd}\n` : "";

    if (applicantEmail) {
      MailApp.sendEmail(
        applicantEmail,
        "【要確認】有給休暇の時季変更についてのご相談",
        `申請者様\n\nお疲れ様です。\n以下の有給休暇申請について、上司より時季変更（日程の再調整）の打診がございました。\n\n・取得予定日: ${formattedDate}（${type}）\n・申請ID: ${requestId}\n${avoidPeriodText}・理由/コメント:\n${reason}\n\n内容をご確認の上、日程の再調整をお願いいたします。\n\n※本メールはシステムからの自動送信です。`
      );
    }
  } else {
    logSheet.getRange(targetRow, 7).setValue("エラー:不明な判定(" + status + ")");
  }
}
