// 1. 深夜に走る関数
function nightlyLeaveGrantBatch() {
  // --- 【追加】スクリプト全体のロックを取得 ---
  const lock = LockService.getScriptLock();

  // 10秒間ロックの獲得を試みる（取得できない場合は処理を中断）
  try {
    const hasLock = lock.tryLock(10000);
    if (!hasLock) {
      console.warn("他のプロセスが実行中のため、今回の自動付与バッチはスキップされました。");
      return;
    }
  } catch (e) {
    console.error("ロック取得時にエラーが発生しました: " + e.message);
    return;
  }

  // ★ ここにCONFIGや設定値をまとめる（スッキリ！）
  const CONFIG = {
    ADMIN_EMAIL: "admin@example.com",
    TIME_ZONE: "Asia/Tokyo"
  };
  const today = new Date();
  const grantRecords = [];

  // ----- 【ここを追加・再定義！】 -----
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ①「社員マスタ」シートからデータを取得
  const empSheet = ss.getSheetByName("社員マスタ");
  const empRawData = empSheet.getDataRange().getValues().slice(1);
  // [ [ID, 氏名, 部署, 入社日, ステータス], [...] ] という2次元配列になる

  //13で見出しは切ったが、データを扱いやすいように列ごとに配列をまとめて、名前を付けておく
  const employeeData = empRawData.map(row => {
    return {
      id: row[0],
      name: row[1],
      hireDate: row[2], // 入社日
      status: row[3]   // 在籍ステータス
    };
  });

  // ③「付与ルールマスタ」シートからデータを取得（ruleMasterも定義し直す）
  const ruleSheet = ss.getSheetByName("付与ルールマスタ");
  const ruleRawData = ruleSheet.getDataRange().getValues().slice(1);
  //18行目と同様に配列をまとめて名前を付ける
  const ruleMaster = ruleRawData.map(row => {
    return {
      months: Number(row[0]),   // 経過月数
      grantDays: Number(row[1]) // 付与日数
    };
  });
  // -------------------------------------


  // （中略：社員マスタとルールマスタのデータを取得）
  //for (const employee of employeeData) {　と同じ//
  employeeData.forEach(employee => {
    const employeeId = employee.id;//ここまででようやく社員IDを獲得!⇒employeeId//
    const hireDate = new Date(employee.hireDate);
    //同様に入社日も獲得⇒hireDate//

    // 在籍ステータスのチェック（退職者を弾くロジックがここにあるとベター）
    if (employee.status !== '在籍') return;
    //forEachは毎回ミニ関数を走らせてるので、return戻って次の関数を走らせる//

    // 経過月数の計算
    const monthsDiff = (today.getFullYear() - hireDate.getFullYear()) * 12 + (today.getMonth() - hireDate.getMonth());
    //getFullYear()だとちゃんと４桁の西暦で返してくれる。*getyear()だと1990年との差を返すためわかりづらいし、バグの温床だから×

    const isAnniversary = today.getDate() === hireDate.getDate()
    // 今日がちょうど「入社日と同じ日」の場合のみチェック *ifをネストさせる（日数でカット）することで、ruleの判定処理が軽くなる//
    if (isAnniversary) {
      const rule = ruleMaster.find(r => r.months === monthsDiff);

      if (rule) {
        // 【改善】2年後の前日を計算　*失効日
        const expireDate = new Date(today.getFullYear() + 2, today.getMonth(), today.getDate() - 1);

        // 【改善】スプレッドシート用に日付をフォーマット化
        const strToday = Utilities.formatDate(today, timeZone, "yyyy/MM/dd");
        const strExpire = Utilities.formatDate(expireDate, timeZone, "yyyy/MM/dd");


        const formattedDate = Utilities.formatDate(today, CONFIG.TIME_ZONE, "yyyyMMdd");
        const GrantId = `${employeeId}_${formattedDate}`;

        // [付与ID, 社員ID, 付与日, 有効期限, 付与日数, 消費済日数]
        grantRecords.push([
          GrantId, // 引数に情報を持たせるとID生成が楽になります
          employeeId,
          strToday,
          //61行目のfindで該当のruleの行をストックしたので、.[要素]でlookupできる//
          strExpire,
          rule.grantDays,
          0
        ]);
      }
    }
  });

  // 書き込み処理*エラーはメール通知
  if (grantRecords.length > 0) {
    try {
      const grantSheet = ss.getSheetByName("付与履歴");
      const lastRow = grantSheet.getLastRow();
      grantSheet.getRange(lastRow + 1, 1, grantRecords.length, grantRecords[0].length).setValues(grantRecords);
      console.log(`${grantRecords.length}件の有給を自動付与しました。`);
    }

    catch (e) {
      // 【エラー通知】何か起きた時だけメールが飛ぶ
      const subject = "【要対応】有給自動付与処理でエラー発生";
      const body = `自動付与処理中にエラーが発生しました。\n\n詳細: ${e.message}\n\n至急、スプレッドシートの「付与履歴」シートを確認してください。`;

      MailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, body);
      console.error(`付与エラー: ${e.message}`);
      throw e; // 処理を中断して異常を確定させる
    } finally {
      // メイン処理でエラーが発生した場合でも確実にロックを解放する---
      lock.releaseLock();
    }
  }
}
