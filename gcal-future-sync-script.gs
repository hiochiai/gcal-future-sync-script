// === 設定項目 ===
// スクリプトを初めて使用する際、または設定を変更する際は、以下の項目を確認・編集してください。
// 設定方法は2通りあります:
// 1. このスクリプト内の定数を直接編集する。
// 2. スクリプトエディタの「プロジェクトの設定」>「スクリプト プロパティ」で設定する（推奨）。

// --- スクリプトプロパティを使用する場合のキー名 ---
// SOURCE_CALENDAR_IDS_CSV : 同期元のカレンダーID（複数ある場合はカンマ区切りで入力）
// DESTINATION_CALENDAR_ID : 同期先のカレンダーID
// DAYS_TO_SYNC            : 同期する日数（今日から何日先までか）
// NO_COPY_TAG_OVERRIDE    : (任意) NO_COPY_TAGの値を上書きする場合に設定

// --- 直接編集する場合のデフォルト値 ---
const DEFAULT_SOURCE_CALENDAR_IDS = [
  'YOUR_SOURCE_CALENDAR_ID_1@group.calendar.google.com', // TODO: 1つ目のコピー元カレンダーIDに置き換えてください
  // 'YOUR_SOURCE_CALENDAR_ID_2@group.calendar.google.com'
];
const DEFAULT_DESTINATION_CALENDAR_ID = 'YOUR_DESTINATION_CALENDAR_ID@group.calendar.google.com'; // TODO: コピー先のカレンダーIDに置き換えてください
const DEFAULT_DAYS_TO_SYNC = 35; // 今日から何日先までの予定を同期対象とするか
const DEFAULT_NO_COPY_TAG = '#nocopy'; // このタグがタイトルに含まれる予定はコピーしない

// --- リトライ処理に関する設定 ---
const MAX_RETRIES = 3; // API呼び出しの最大リトライ回数
const RETRY_DELAY_MS = 2000; // リトライ時の待機時間 (ミリ秒)

/**
 * 指定されたアクションをリトライ付きで実行するヘルパー関数。
 * @param {function} action - 実行するアクション (API呼び出しなど)。
 * @param {string} actionDescription - ログ出力用のアクションの説明。
 * @return {any} アクションが成功した場合の戻り値。
 * @throws {Error} 全てのリトライが失敗した場合。
 */
function executeWithRetry(action, actionDescription) {
  for (let i = 0; i <= MAX_RETRIES; i++) {
    try {
      return action();
    } catch (e) {
      // Google Calendar APIのレート制限エラーメッセージの一部をチェック (より堅牢にするにはエラーコードなどがあればそれを使う)
      if (e.message && e.message.includes("You have been creating or deleting too many calendars or calendar events")) {
        if (i < MAX_RETRIES) {
          Logger.log(`警告: ${actionDescription} でレート制限エラーが発生しました。${RETRY_DELAY_MS / 1000}秒後にリトライします... (試行 ${i + 1}/${MAX_RETRIES + 1})`);
          Utilities.sleep(RETRY_DELAY_MS * (i + 1)); // リトライごとに待機時間を増やす (簡易的なエクスポネンシャルバックオフ)
        } else {
          Logger.log(`エラー: ${actionDescription} の最大リトライ回数(${MAX_RETRIES + 1})を超えました。エラー: ${e.toString()}`);
          throw e; // 最終的にエラーをスロー
        }
      } else {
        // レート制限以外のエラーは即座にスロー
        Logger.log(`エラー: ${actionDescription} で予期せぬエラーが発生しました: ${e.toString()}`);
        throw e;
      }
    }
  }
}


/**
 * メインの同期処理関数。
 * 1. デスティネーションカレンダーの今日から指定日数先までの予定を全て削除します。
 * 2. ソースカレンダーから今日以降の予定を取得し、デスティネーションカレンダーにコピーします。
 */
function simpleSyncCalendarEvents() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const sourceCalendarIdsCsv = properties.getProperty('SOURCE_CALENDAR_IDS_CSV');
    const sourceCalendarIds = sourceCalendarIdsCsv
                              ? sourceCalendarIdsCsv.split(',').map(id => id.trim()).filter(id => id)
                              : DEFAULT_SOURCE_CALENDAR_IDS;
    const destinationCalendarId = properties.getProperty('DESTINATION_CALENDAR_ID') || DEFAULT_DESTINATION_CALENDAR_ID;
    const daysToSync = parseInt(properties.getProperty('DAYS_TO_SYNC')) || DEFAULT_DAYS_TO_SYNC;
    const noCopyTag = properties.getProperty('NO_COPY_TAG_OVERRIDE') || DEFAULT_NO_COPY_TAG;

    if (!sourceCalendarIds || sourceCalendarIds.length === 0 || !destinationCalendarId || destinationCalendarId.startsWith('YOUR_')) {
      Logger.log('エラー: 同期元カレンダーID(1つ以上)または同期先カレンダーIDが正しく設定されていません。スクリプト冒頭の定数またはスクリプトプロパティを確認し、実際のIDに置き換えてください。');
      return;
    }
    if (sourceCalendarIds.some(id => id.startsWith('YOUR_'))) {
        Logger.log('警告: 同期元カレンダーIDのいずれかが初期値のままです。実際のIDに置き換えてください。');
    }

    const destinationCalendar = CalendarApp.getCalendarById(destinationCalendarId);
    if (!destinationCalendar) {
      Logger.log(`エラー: 同期先カレンダーが見つかりません: ${destinationCalendarId}`);
      return;
    }

    const today = new Date();
    const startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate()); // 今日の0時0分
    // ソースから取得する期間および削除対象期間の終了日
    const relevantEndDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() + daysToSync);

    Logger.log(`処理開始: ${new Date().toLocaleString()}`);
    Logger.log(`同期元カレンダーID: ${sourceCalendarIds.join(', ')}`);
    Logger.log(`同期先カレンダーID: ${destinationCalendarId}`);
    Logger.log(`同期期間および削除対象期間: 今日 (${startDate.toLocaleDateString()}) から ${daysToSync} 日間 (終了日: ${relevantEndDate.toLocaleDateString()})`);
    Logger.log(`同期除外タグ: "${noCopyTag}"`);

    // 1. デスティネーションカレンダーの今日から指定日数先までの予定を全て削除
    Logger.log(`ステップ1: 同期先カレンダー "${destinationCalendar.getName()}" の今日から ${daysToSync} 日先までの予定を削除します...`);
    const eventsToDelete = destinationCalendar.getEvents(startDate, relevantEndDate);
    let deletedCount = 0;
    if (eventsToDelete.length > 0) {
      for (const event of eventsToDelete) {
        try {
          const actionDescription = `予定 "${event.getTitle()}" (${event.getStartTime()}) の削除`;
          executeWithRetry(() => event.deleteEvent(), actionDescription);
          deletedCount++;
        } catch (e) {
          // executeWithRetry内で最終エラーがログされているので、ここでは追加ログなしでもOK
          // Logger.log(`  リトライ超過エラー: ${actionDescription} に失敗しました。`);
        }
      }
      Logger.log(`  ${deletedCount} 件の予定を同期先カレンダーから削除しました。`);
    } else {
      Logger.log(`  同期先カレンダーの指定期間内に削除対象の予定はありませんでした。`);
    }


    // 2. ソースカレンダーから今日以降の予定を取得し、デスティネーションカレンダーにコピー
    Logger.log(`ステップ2: 同期元カレンダーから予定をコピーします...`);
    let copiedCount = 0;
    let skippedByTagCount = 0;
    let totalSourceEventsFetched = 0;

    for (const sourceCalId of sourceCalendarIds) {
      if (!sourceCalId) continue;
      const sourceCalendar = CalendarApp.getCalendarById(sourceCalId);
      if (!sourceCalendar) {
        Logger.log(`  警告: 同期元カレンダーが見つかりません: ${sourceCalId}。スキップします。`);
        continue;
      }
      Logger.log(`  同期元カレンダー "${sourceCalendar.getName()}" (${sourceCalId}) から予定を取得中 (期間: ${startDate.toLocaleDateString()} - ${relevantEndDate.toLocaleDateString()})...`);
      const sourceEventsRaw = sourceCalendar.getEvents(startDate, relevantEndDate);
      totalSourceEventsFetched += sourceEventsRaw.length;
      Logger.log(`    "${sourceCalendar.getName()}" から ${sourceEventsRaw.length} 件の予定を取得しました。`);

      for (const srcEvent of sourceEventsRaw) {
        if (srcEvent.getTitle().includes(noCopyTag)) {
          skippedByTagCount++;
          continue;
        }

        try {
          const options = {
            description: srcEvent.getDescription() || "",
            location: srcEvent.getLocation() || "",
          };
          if (srcEvent.isAllDayEvent()) {
            const actionDescription = `全日予定 "${srcEvent.getTitle()}" (${srcEvent.getAllDayStartDate()}) の作成`;
            executeWithRetry(() => destinationCalendar.createAllDayEvent(srcEvent.getTitle(), srcEvent.getAllDayStartDate(), srcEvent.getAllDayEndDate(), options), actionDescription);
          } else {
            const actionDescription = `通常予定 "${srcEvent.getTitle()}" (${srcEvent.getStartTime()}) の作成`;
            executeWithRetry(() => destinationCalendar.createEvent(srcEvent.getTitle(), srcEvent.getStartTime(), srcEvent.getEndTime(), options), actionDescription);
          }
          copiedCount++;
        } catch (e) {
          // executeWithRetry内で最終エラーがログされているので、ここでは追加ログなしでもOK
          // Logger.log(`    リトライ超過エラー: 予定 "${srcEvent.getTitle()}" のコピーに失敗しました。`);
        }
      }
    }
    Logger.log(`  全ての同期元カレンダーから合計 ${totalSourceEventsFetched} 件の予定を取得しました。`);
    Logger.log(`  ${copiedCount} 件の予定を同期先カレンダーにコピーしました。`);
    Logger.log(`  ${skippedByTagCount} 件の予定が "${noCopyTag}" タグによりスキップされました。`);
    Logger.log(`処理完了: ${new Date().toLocaleString()}`);

  } catch (error) {
    Logger.log(`致命的なエラーが発生しました: ${error.toString()}\nStack: ${error.stack}`);
  }
}

