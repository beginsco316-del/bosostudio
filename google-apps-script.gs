function doPost(e) {
  var data = JSON.parse(e.postData.contents || "{}");

  if (data.action === "syncCalendar") {
    syncCalendarList_(data);
    return json_({ ok: true });
  }

  writeAllSheets_(data);
  return json_({ ok: true });
}

function doGet(e) {
  var params = e.parameter || {};

  if (params.action === "syncCalendarReservation") {
    var reservation = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(params.payload || "")).getDataAsString());
    syncCalendarReservation_(reservation, params.calendarId, Number(params.eventDurationMinutes || 60));
    return jsonp_(params.callback, { ok: true });
  }

  return jsonp_(params.callback, readAllSheets_());
}

function writeAllSheets_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  writeSheet_(ss, "Customers", [
    ["id", "name", "phone", "childName", "childInfo", "address", "memo", "createdAt"]
  ], (data.customers || []).map(function(c) {
    return [c.id, c.name, c.phone, c.childName, c.childInfo, c.address, c.memo, c.createdAt];
  }));

  writeSheet_(ss, "Visits", [
    ["id", "customerId", "visitNo", "date", "shootTime", "shootType", "productName", "totalAmount", "deposit", "balance", "balancePaymentMethod", "balancePaymentStaff", "deliveryStatus", "memo", "photoCount", "createdAt"]
  ], (data.visits || []).map(function(v) {
    return [
      v.id, v.customerId, v.visitNo, v.date, v.shootTime, v.shootType, v.productName,
      v.totalAmount, v.deposit, v.balance, v.balancePaymentMethod, v.balancePaymentStaff,
      v.deliveryStatus, v.memo, (v.photos || []).length, v.createdAt
    ];
  }));

  writeSheet_(ss, "Reservations", [
    ["id", "customerId", "customerName", "date", "time", "shootType", "productName", "staff", "status", "memo", "sourceVisitId", "autoFromVisit", "calendarEventId", "createdAt"]
  ], (data.reservations || []).map(function(r) {
    return [
      r.id, r.customerId, r.customerName, r.date, r.time, r.shootType, r.productName,
      r.staff, r.status, r.memo, r.sourceVisitId, r.autoFromVisit, r.calendarEventId, r.createdAt
    ];
  }));
}

function readAllSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    customers: readSheet_(ss, "Customers"),
    visits: readSheet_(ss, "Visits"),
    reservations: readSheet_(ss, "Reservations")
  };
}

function writeSheet_(ss, name, header, rows) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  var values = header.concat(rows || []);
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}

function readSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) return [];

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  var headers = values[0];
  return values.slice(1).filter(function(row) {
    return row.some(function(value) { return value !== ""; });
  }).map(function(row) {
    var item = {};
    headers.forEach(function(header, index) {
      item[header] = row[index];
    });
    return normalizeItem_(item);
  });
}

function normalizeItem_(item) {
  Object.keys(item).forEach(function(key) {
    if (item[key] instanceof Date) item[key] = Utilities.formatDate(item[key], Session.getScriptTimeZone(), "yyyy-MM-dd");
  });
  ["totalAmount", "deposit", "balance", "visitNo"].forEach(function(key) {
    if (item[key] !== "" && item[key] != null) item[key] = Number(item[key]);
  });
  if (item.autoFromVisit === "TRUE" || item.autoFromVisit === true) item.autoFromVisit = true;
  if (item.autoFromVisit === "FALSE" || item.autoFromVisit === false || item.autoFromVisit === "") item.autoFromVisit = false;
  return item;
}

function syncCalendarList_(data) {
  var calendarId = data.calendarId;
  var duration = Number(data.eventDurationMinutes || 60);
  (data.reservations || []).forEach(function(reservation) {
    syncCalendarReservation_(reservation, calendarId, duration);
  });
}

function syncCalendarReservation_(reservation, calendarId, duration) {
  if (!calendarId || !reservation || !reservation.date) return;

  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) throw new Error("Calendar not found");

  var start = makeDateTime_(reservation.date, reservation.time);
  var end = new Date(start.getTime() + duration * 60 * 1000);
  var title = "[보소] " + (reservation.customerName || reservation.customerId || "예약") + " - " + (reservation.shootType || "촬영");
  var marker = "BOSO_RESERVATION_ID:" + (reservation.id || "");
  var description = [
    marker,
    "고객번호: " + (reservation.customerId || ""),
    "연락처: " + (reservation.customerPhone || ""),
    "아이: " + (reservation.childName || ""),
    "상품: " + (reservation.productName || ""),
    "메모: " + (reservation.memo || "")
  ].join("\n");

  var dayStart = makeDateTime_(reservation.date, "00:00");
  var dayEnd = makeDateTime_(reservation.date, "23:59");
  var events = calendar.getEvents(dayStart, dayEnd);
  var existing = events.find(function(event) {
    return (event.getDescription() || "").indexOf(marker) !== -1;
  });

  if (reservation.status === "취소") {
    if (existing) existing.deleteEvent();
    return;
  }

  if (existing) {
    existing.setTitle(title);
    existing.setTime(start, end);
    existing.setDescription(description);
  } else {
    calendar.createEvent(title, start, end, { description: description });
  }
}

function makeDateTime_(dateText, timeText) {
  var parts = String(dateText).split("-");
  var time = String(timeText || "10:00").split(":");
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), Number(time[0] || 10), Number(time[1] || 0));
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(callback, data) {
  var text = JSON.stringify(data);
  if (callback) text = callback + "(" + text + ")";
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JAVASCRIPT);
}
