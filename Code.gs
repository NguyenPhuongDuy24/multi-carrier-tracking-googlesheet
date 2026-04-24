var API_KEY = "your_api_17track"; 

var COL_TRACKING    = 1; // A - Mã vận đơn
var COL_PHONE       = 2; // B - 4 số cuối SĐT
var COL_CARRIER_IN  = 3; // C - Mã hãng nhập vào
var COL_CARRIER_OUT = 4; // D - Hãng thực tế từ 17TRACK
var COL_STATUS      = 5; // E - Trạng thái
var COL_DETAIL      = 6; // F - Chi tiết mới nhất
var COL_UPDATED     = 7; // G - Cập nhật lúc

var STATUS_MAP = {
  "Delivered": "✅ Đã giao hàng",
  "InTransit": "🚚 Đang vận chuyển",
  "PickedUp": "📦 Đã lấy hàng",
  "OutForDelivery": "🛵 Đang giao hàng",
  "NotFound": "❓ Không tìm thấy",
  "Exception": "⚠️ Ngoại lệ",
  "Expired": "⌛ Quá hạn",
  "InfoReceived": "📄 Đã nhận thông tin"
};

// ── Menu ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚚 17TRACK")
    .addItem("Mở bảng điều khiển", "openSidebar")
    .addItem("Cài tự động cập nhật mỗi giờ", "setupTrigger")
    .addItem("Xóa tự động cập nhật", "deleteTrigger")
    .addToUi();
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("🚚 17TRACK Tracker")
    .setWidth(290);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Lấy sheet đầu tiên ───────────────────────────────
function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

// ── Thêm 1 đơn từ Sidebar ────────────────────────────
function addTrackingFromSidebar(trackingNum, phone, carrierCode) {
  try {
    var sheet   = getSheet();
    var lastRow = sheet.getLastRow();

    // Kiểm tra trùng
    if (lastRow >= 2) {
      var existing = sheet.getRange(2, COL_TRACKING, lastRow - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (existing[i][0].toString().trim() === trackingNum) {
          return { success: false, message: "Mã vận đơn đã tồn tại!" };
        }
      }
    }

    var newRow = lastRow + 1;
    sheet.getRange(newRow, COL_TRACKING).setValue(trackingNum);
    sheet.getRange(newRow, COL_PHONE).setValue(phone);
    sheet.getRange(newRow, COL_CARRIER_IN).setValue(carrierCode ? parseInt(carrierCode) : "");
    sheet.getRange(newRow, COL_STATUS).setValue("⏳ Chờ cập nhật");

    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ── Bước 1: Đăng ký ──────────────────────────────────
function registerOnly() {
  try {
    // Kiểm tra API_KEY
    if (!API_KEY || API_KEY.trim() === "") {
      return { success: false, message: "❌ API_KEY chưa được cài đặt! Vui lòng điền key trong Code.gs" };
    }

    var sheet   = getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "Không có dữ liệu trong sheet!" };

    var items = getItemsFromSheet(sheet, lastRow);
    if (items.length === 0) return { success: false, message: "Không có mã vận đơn nào!" };

    var result = apiCall("register", items);
    Logger.log("Register: " + JSON.stringify(result));

    // Kiểm tra response từ API
    if (!result || result.code !== 0) {
      var errMsg = result && result.msg ? result.msg : "Lỗi từ API 17TRACK";
      return { success: false, message: "❌ " + errMsg };
    }

    return { success: true, count: items.length };
  } catch(e) {
    return { success: false, message: "Lỗi: " + e.message };
  }
}

// ── Bước 2: Lấy thông tin ────────────────────────────
/**
 * Hàm lấy thông tin chi tiết từ 17TRACK và cập nhật vào Sheet
 * Đã tối ưu cho cấu trúc JSON mới (track_info, latest_status, latest_event)
 */
function getInfoOnly() {
  try {
    // Kiểm tra API_KEY
    if (!API_KEY || API_KEY.trim() === "") {
      return { success: false, message: "❌ API_KEY chưa được cài đặt! Vui lòng điền key trong Code.gs" };
    }

    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    
    // 1. Kiểm tra dữ liệu đầu vào
    if (lastRow < 2) {
      return { success: false, message: "Bảng tính không có dữ liệu để cập nhật!" };
    }

    var items = getItemsFromSheet(sheet, lastRow);
    if (items.length === 0) {
      return { success: false, message: "Không tìm thấy mã vận đơn nào hợp lệ!" };
    }

    // 2. Gọi API 17TRACK (Xử lý batch 40 mã mỗi lần để tránh quá tải)
    var trackMap = {};
    var BATCH_SIZE = 40;
    var hasApiError = false;
    Logger.log("📤 Gửi " + items.length + " mã tới API...");
    Logger.log("Items mẫu: " + JSON.stringify(items.slice(0, 2)));

    for (var i = 0; i < items.length; i += BATCH_SIZE) {
      var batch = items.slice(i, i + BATCH_SIZE);
      Logger.log("📦 Batch " + (i/BATCH_SIZE + 1) + ": " + batch.map(function(x) { return x.number; }).join(", "));
      var result = apiCall("gettrackinfo", batch);

      // Kiểm tra lỗi API
      if (!result || result.code !== 0) {
        var errMsg = result && result.msg ? result.msg : "Lỗi từ API 17TRACK";
        Logger.log("❌ API Error (batch " + (i/BATCH_SIZE + 1) + "): " + errMsg);
        hasApiError = true;
        continue;
      }

      if (result.data) {
        // API có thể trả về:
        // 1. Mảy: [{ number: "...", track_info: {...} }]
        // 2. Object được wrap: { accepted: [{...}] }
        // 3. Object đơn: { number: "...", track_info: {...} }
        var list = result.data.accepted || result.data;
        
        if (Array.isArray(list)) {
          // Trường hợp 1: Mảy
          Logger.log("📥 Nhận mảy " + list.length + " tracking");
          Logger.log("📋 Data structure mẫu từ API: " + JSON.stringify(list[0]).substring(0, 500));
          list.forEach(function(obj) {
            trackMap[obj.number] = obj;
          });
        } else if (list && typeof list === 'object' && list.number) {
          // Trường hợp 3: Object đơn
          trackMap[list.number] = list;
          Logger.log("✅ API trả về object đơn: " + list.number);
          Logger.log("📋 Data structure: " + JSON.stringify(list).substring(0, 500));
        }
      }
      // Nghỉ 1 giây giữa các batch để tránh bị API chặn
      if (i + BATCH_SIZE < items.length) Utilities.sleep(1000);
    }

    Logger.log("📊 trackMap có " + Object.keys(trackMap).length + " tracking: " + Object.keys(trackMap).join(", "));

    // Kiểm tra nếu không lấy được dữ liệu
    if (Object.keys(trackMap).length === 0) {
      var msg = hasApiError ? "❌ API trả về lỗi! Kiểm tra API_KEY hoặc kết nối mạng." : "❌ Không lấy được dữ liệu từ API!";
      return { success: false, message: msg };
    }

    // 3. Cập nhật dữ liệu vào Sheet
    var now = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm");
    var updatedCount = 0;
    var notFoundCount = 0;
    Logger.log("📋 Bắt đầu scan " + (lastRow - 1) + " dòng trong Sheet...");

    for (var row = 2; row <= lastRow; row++) {
      var trackingNum = sheet.getRange(row, COL_TRACKING).getValue().toString().trim();
      if (!trackingNum) continue;

      Logger.log("🔍 Dòng " + row + ": Tìm kiếm '" + trackingNum + "'...");
      
      if (!trackMap[trackingNum]) {
        notFoundCount++;
        Logger.log("⚠️ Không tìm thấy trong API response: " + trackingNum);
        continue;
      }

      Logger.log("✅ Tìm thấy trong trackMap: " + trackingNum);

      var data = trackMap[trackingNum];
      
      // API 17TRACK trả về format compact với key viết tắt
      // Nếu có track_info, dùng format cũ; nếu không, dùng format mới (track)
      var info = data.track_info || data.track;
      
      if (!info) {
        Logger.log("❌ Không tìm thấy dữ liệu tracking cho: " + trackingNum);
        continue;
      }

      // --- LẤY TRẠNG THÁI (STATUS) ---
      // Format cũ: info.latest_status.status
      // Format mới: info.e (status code)
      var rawStatus = "NotFound";
      if (info.latest_status && info.latest_status.status) {
        // Format cũ
        rawStatus = info.latest_status.status;
      } else if (info.e !== undefined) {
        // Format mới - map status code
        // 0-10 = InTransit, 301 = Delivered, etc.
        var statusCode = info.e;
        if (statusCode === 0) rawStatus = "InfoReceived";
        else if (statusCode >= 1 && statusCode <= 10) rawStatus = "InTransit";
        else if (statusCode >= 11 && statusCode <= 100) rawStatus = "PickedUp";
        else if (statusCode === 101) rawStatus = "OutForDelivery";
        else if (statusCode === 301 || statusCode === 303) rawStatus = "Delivered";
        else if (statusCode >= 400) rawStatus = "Exception";
        Logger.log("📊 Status code: " + statusCode + " -> " + rawStatus);
      }
      var statusText = STATUS_MAP[rawStatus] || rawStatus;

      // --- LẤY CHI TIẾT MỚI NHẤT (DETAIL) ---
      var detailText = "";
      if (info.latest_event) {
        // Format cũ
        var eventTime = info.latest_event.time_iso ? info.latest_event.time_iso.substring(0, 16).replace("T", " ") : "";
        var eventDesc = info.latest_event.description || "";
        detailText = eventTime + " — " + eventDesc;
      } else if (info.z0 && info.z0.a) {
        // Format mới - z0 = latest event
        var eventTime = info.z0.a; // "2026-04-24 14:20"
        var eventDesc = info.z0.z || info.z0.d || "";
        detailText = eventTime + " — " + eventDesc;
        Logger.log("📝 Chi tiết: " + detailText.substring(0, 50));
      }

      // --- GHI DỮ LIỆU XUỐNG CÁC CỘT ---
      sheet.getRange(row, COL_STATUS).setValue(statusText);
      sheet.getRange(row, COL_DETAIL).setValue(detailText);
      sheet.getRange(row, COL_UPDATED).setValue(now);

      // --- ĐỊNH DẠNG MÀU NỀN (COLORING) ---
      var bgColor = null;
      switch (rawStatus) {
        case "Delivered":
          bgColor = "#c6f6d5"; // Xanh lá nhạt (Thành công)
          break;
        case "InTransit":
        case "OutForDelivery":
        case "PickedUp":
          bgColor = "#fefcbf"; // Vàng nhạt (Đang chạy)
          break;
        case "Exception":
        case "Undelivered":
        case "NotFound":
          bgColor = "#fed7d7"; // Đỏ nhạt (Lỗi/Vấn đề)
          break;
        default:
          bgColor = null;
      }
      sheet.getRange(row, COL_STATUS).setBackground(bgColor);

      updatedCount++;
      Logger.log("✅ Cập nhật: " + trackingNum + " -> " + statusText);
    }

    Logger.log("📊 Tổng kết: Cập nhật " + updatedCount + " đơn, không tìm được " + notFoundCount + " đơn");
    return { success: true, count: updatedCount };

  } catch (e) {
    Logger.log("Lỗi getInfoOnly: " + e.message);
    return { success: false, message: "Lỗi hệ thống: " + e.message };
  }
}

// ── Tự động cập nhật (trigger) ───────────────────────
function autoUpdate() {
  // Chỉ cần lấy thông tin, không cần đăng ký lại mỗi lần
  // (Đăng ký chỉ cần khi thêm đơn mới)
  getInfoOnly();
}

// ── Thống kê ─────────────────────────────────────────
function getStats() {
  try {
    var sheet   = getSheet();
    var lastRow = sheet.getLastRow();
    var stats   = { total: 0, delivered: 0, transit: 0, error: 0 };
    if (lastRow < 2) return stats;

    var statuses = sheet.getRange(2, COL_STATUS, lastRow - 1, 1).getValues();
    for (var i = 0; i < statuses.length; i++) {
      var s = statuses[i][0].toString().trim();
      if (!s || s === "" || s === "⏳ Chờ cập nhật") continue;
      stats.total++;
      if (s.indexOf("Đã giao") > -1)                stats.delivered++;
      else if (s.indexOf("Đang vận chuyển") > -1)   stats.transit++;
      else if (s.indexOf("⚠️") > -1 || s.indexOf("Giao thất bại") > -1) stats.error++;
    }
    return stats;
  } catch(e) {
    return { total: 0, delivered: 0, transit: 0, error: 0 };
  }
}

// ── Helper: đọc items từ sheet ───────────────────────
function getItemsFromSheet(sheet, lastRow) {
  var rawValues = sheet.getRange(2, COL_TRACKING, lastRow - 1, 3).getValues();
  var items = [];
  for (var i = 0; i < rawValues.length; i++) {
    var num       = rawValues[i][0].toString().trim();
    var phone     = rawValues[i][1].toString().trim();
    var carrierIn = rawValues[i][2].toString().trim();
    if (!num) continue;

    var item = { number: num };
    if (carrierIn && !isNaN(carrierIn)) {
      item.carrier = parseInt(carrierIn);
    }
    if (phone) item.param = phone;
    items.push(item);
  }
  return items;
}

// ── Helper: gọi API 17TRACK ──────────────────────────
function apiCall(endpoint, payload) {
  try {
    var res = UrlFetchApp.fetch("https://api.17track.net/track/v1/" + endpoint, {
      method: "post",
      headers: {
        "17token": API_KEY,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 30000
    });

    var statusCode = res.getResponseCode();
    var content = res.getContentText();

    if (statusCode !== 200) {
      Logger.log("❌ API Error - Status: " + statusCode + ", Response: " + content);
      return { code: -1, msg: "HTTP Error " + statusCode };
    }

    return JSON.parse(content);
  } catch(e) {
    Logger.log("❌ API Call Error: " + e.message);
    return { code: -1, msg: "Lỗi kết nối: " + e.message };
  }
}

// ── Trigger ──────────────────────────────────────────
function setupTrigger() {
  deleteTrigger();
  ScriptApp.newTrigger("autoUpdate").timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert("✅ Đã cài tự động cập nhật mỗi 1 giờ!");
}

function deleteTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "autoUpdate") ScriptApp.deleteTrigger(t);
  });
}
