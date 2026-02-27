/**
 * 指定した日の既存の予定を取得し、バッファ時間を加味した時間を返す
 */
function getBusySlots(date) {
  // ガード句：dateが空なら今日にする
  const targetDate = (date instanceof Date) ? date : new Date();
  
  const config = getConfig();
  const bufferMs = config.buffer * 60 * 1000;
  
  const startOfDay = new Date(targetDate.getTime());
  startOfDay.setHours(0, 0, 0, 0);
  const endOfDay = new Date(targetDate.getTime());
  endOfDay.setHours(23, 59, 59, 999);

  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(startOfDay, endOfDay);
  
  return events.map(event => ({
    start: new Date(event.getStartTime().getTime() - bufferMs),
    end: new Date(event.getEndTime().getTime() + bufferMs)
  }));
}

/**
 * 予約可能なスロットを計算する
 */
function getAvailableSlots(date) {
  // ガード句：dateが空なら今日にする
  const targetDate = (date instanceof Date) ? date : new Date();
  
  const config = getConfig();
  const busySlots = getBusySlots(targetDate);
  const durationMs = config.slotDuration * 60 * 1000;
  
  const startLimit = new Date(targetDate.getTime());
  const [startH, startM] = config.displayStart.split(':');
  startLimit.setHours(Number(startH), Number(startM), 0, 0);

  const endLimit = new Date(targetDate.getTime());
  const [endH, endM] = config.displayEnd.split(':');
  endLimit.setHours(Number(endH), Number(endM), 0, 0);

  const availableSlots = [];
  let currentPos = startLimit.getTime();

  while (currentPos + durationMs <= endLimit.getTime()) {
    const slotStart = new Date(currentPos);
    const slotEnd = new Date(currentPos + durationMs);

    const isOverlap = busySlots.some(busy => (slotStart < busy.end && slotEnd > busy.start));
    
    if (!isOverlap) {
      availableSlots.push(Utilities.formatDate(slotStart, "JST", "HH:mm"));
    }
    currentPos += durationMs;
  }
  
  return { slots: availableSlots, config: config };
}

/**
 * テスト実行用（必ずこれを選んで実行してね）
 */
function testAvailable() {
  const res = getAvailableSlots(new Date());
  console.log("予約可能枠:", res.slots);
}