function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('日程予約システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAvailableSlotsForFront(dateStr) {
  const date = new Date(dateStr.replace(/-/g, '/'));
  return getAvailableSlots(date);
}

function submitReservation(formData) {
  try {
    console.log("予約リクエスト受信:", formData);
    const config = getConfig();
    const times = formData.times.sort(); 
    
    // 日時の組み立て
    const startTime = new Date(formData.date.replace(/-/g, '/'));
    const [startH, startM] = times[0].split(':');
    startTime.setHours(Number(startH), Number(startM), 0, 0);
    
    const totalDurationMs = times.length * config.slotDuration * 60 * 1000;
    const endTime = new Date(startTime.getTime() + totalDurationMs);

    const calendar = CalendarApp.getDefaultCalendar();
    
    // 予定の作成
    const event = calendar.createEvent(`【予約】${formData.name}様`, startTime, endTime, {
      description: `備考: ${formData.notes}\nメール: ${formData.email}`,
      guests: formData.email,
      sendInvites: true
    });

    console.log("カレンダー登録成功:", event.getId());

    // Google Meet発行 (ここでコケることが多いのでtry-catchで囲む)
    try {
      const resource = {
        conferenceData: { 
          createRequest: { 
            requestId: "req_" + Date.now(), 
            conferenceSolutionKey: { type: "hangoutsMeet" } 
          } 
        }
      };
      Calendar.Events.patch(resource, calendar.getId(), event.getId().split('@')[0], { conferenceDataVersion: 1 });
      console.log("Meet発行成功");
    } catch (meetError) {
      console.warn("Meet発行に失敗しましたが、予約は完了しています:", meetError);
    }

    return "OK";

  } catch (e) {
    console.error("予約実行エラー:", e.message);
    throw new Error("サーバー側でエラーが発生しました: " + e.message);
  }
}