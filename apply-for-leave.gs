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

