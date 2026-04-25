var COL_TRACKING    = 1; // A - Mã vận đơn
var COL_PHONE       = 2; // B - 4 số cuối SĐT
var COL_CARRIER_OUT = 3; // C - Hãng thực tế từ 17TRACK
var COL_STATUS      = 4; // D - Trạng thái
var COL_DETAIL      = 5; // E - Chi tiết mới nhất
var COL_UPDATED     = 6; // F - Cập nhật lúc

var STATUS_MAP = {
  "Delivered": "✅ Đã giao hàng",
  "InTransit": "🚚 Đang vận chuyển",
  "PickedUp": "📦 Đã lấy hàng",
  "OutForDelivery": "🛵 Đang giao hàng",
  "NotFound": "❓ Không tìm thấy",
  "Exception": "⚠️ Ngoại lệ",
  "Undelivered": "⚠️ Giao thất bại",
  "Expired": "⌛ Quá hạn",
  "InfoReceived": "📄 Đã nhận thông tin"
};

// ── Quản lý API Key (Lưu ngầm) ───────────────────────
function getApiKey() {
  // Trả về API Key lưu trong tài khoản Google của người dùng
  return PropertiesService.getUserProperties().getProperty("17TRACK_API_KEY") || "";
}

function saveApiKey(key) {
  // Lưu API Key
  PropertiesService.getUserProperties().setProperty("17TRACK_API_KEY", key.trim());
  return { success: true, message: "Đã lưu API Key thành công!" };
}

// ── Menu ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚚 17TRACK")
    .addItem("Mở bảng điều khiển", "openSidebar")
    .addItem("⚙️ Khởi tạo Bảng & Cài đặt", "initSheet")
    .addItem("📊 Cập nhật thống kê", "updateSheetStats")
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

// ── Khởi tạo Bảng & Cài đặt (Gộp 2 hàm thành 1) ───────
function initSheet() {
  var sheet = getSheet();
  var headers = ["Mã vận đơn", "4 số ĐT cuối (sf)", "Đơn vị vận chuyển", "Trạng thái", "Chi tiết", "Cập nhật lúc"];
  var maxRows = Math.max(sheet.getMaxRows(), 100);
  
  // Bước 1: Khởi tạo tiêu đề nếu chưa có
  var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var needsInit = false;
  for (var i = 0; i < headers.length; i++) {
    if (currentHeaders[i] !== headers[i]) {
      needsInit = true;
      break;
    }
  }
  
  if (needsInit) {
    var range = sheet.getRange(1, 1, 1, headers.length);
    range.setValues([headers]);
    range.setFontWeight("bold").setBackground("#4a5568").setFontColor("#ffffff").setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
  
  // Bước 2: Tạo Dropdown cho cột Hãng vận chuyển (Cột C)
  var list = ["Tự nhận diện"].concat(getCarrierList().map(function(c) { return c.name; }));
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).build();
  sheet.getRange(2, COL_CARRIER_OUT, maxRows - 1, 1).setDataValidation(rule);

  // Bước 2.5: Căn giữa & xuống dòng cho toàn bộ dữ liệu
  sheet.getRange(2, 1, maxRows - 1, headers.length)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  // Bước 3: Cài đặt Trigger OnEdit
  var triggers = ScriptApp.getProjectTriggers();
  var hasTrigger = false;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onSheetEdit") {
      hasTrigger = true; break;
    }
  }
  if (!hasTrigger) {
    ScriptApp.newTrigger("onSheetEdit").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Đã khởi tạo bảng, tạo Dropdown & bật tự tra cứu!", "Hoàn tất");
  updateSheetStats();
}

// ── Xử lý sự kiện khi chọn Dropdown ở Cột D ──────────
function onSheetEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Nếu người dùng thao tác ở cột Hãng vận chuyển (Cột 3) và từ dòng 2 trở đi
  if (sheet.getIndex() === 1 && col === COL_CARRIER_OUT && row >= 2) {
    var carrierVal = e.value;
    if (!carrierVal) return; // Bỏ qua nếu xóa ô

    var trackingNum = sheet.getRange(row, COL_TRACKING).getValue().toString().trim();
    var phone = sheet.getRange(row, COL_PHONE).getValue().toString().trim();

    if (!trackingNum) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Vui lòng nhập Mã vận đơn ở cột A trước khi chọn hãng!", "⚠️ Lưu ý");
      return;
    }

    var carrierCode = (carrierVal === "Tự nhận diện") ? "" : carrierVal.split("—").pop().trim();
    SpreadsheetApp.getActiveSpreadsheet().toast("Đang tự động đăng ký và lấy dữ liệu...", "⏳ Đang tra cứu mã " + trackingNum);
    
    var res = addAndTrackSingle(trackingNum, phone, carrierCode);
    if (res && res.success) SpreadsheetApp.getActiveSpreadsheet().toast(res.message, "✅ Hoàn tất");
    else if (res) SpreadsheetApp.getActiveSpreadsheet().toast(res.message, "❌ Lỗi");
  }
}

// ── Thêm & Tra cứu nhanh 1 đơn ───────────────────────
function addAndTrackSingle(trackingNum, phone, carrierCode) {
  try {
    var API_KEY = getApiKey();
    if (!API_KEY || API_KEY.trim() === "") {
      return { success: false, message: "❌ Chưa cài đặt API Key!" };
    }

    var sheet   = getSheet();
    var lastRow = sheet.getLastRow();

    // Tìm dòng nếu đã tồn tại, nếu chưa thì thêm mới
    var rowIndex = -1;
    if (lastRow >= 2) {
      var existing = sheet.getRange(2, COL_TRACKING, lastRow - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (existing[i][0].toString().trim() === trackingNum) {
          rowIndex = i + 2;
          break;
        }
      }
    }

    if (rowIndex === -1) {
      rowIndex = lastRow + 1;
      sheet.getRange(rowIndex, COL_TRACKING).setValue(trackingNum);
    }

    // Cập nhật SĐT
    sheet.getRange(rowIndex, COL_PHONE).setValue(phone);
    sheet.getRange(rowIndex, COL_STATUS)
         .setValue("⏳ Đang tra cứu...")
         .setBackground("#e2e8f0")
         .setFontColor("#2d3748")
         .setFontWeight("bold");

    SpreadsheetApp.flush(); // Ép cập nhật giao diện Sheet ngay lập tức

    var item = { number: trackingNum };
    if (carrierCode && !isNaN(carrierCode)) item.carrier = parseInt(carrierCode);
    if (phone) item.param = phone;

    // Bước 1: Đăng ký
    apiCall("register", [item]);

    var data = null;
    var rawStatus = "NotFound";

    // Bước 2: Lặp để chờ API 17TRACK xử lý (Tối đa 4 lần, tổng 10 giây)
    for (var r = 0; r < 4; r++) {
      Utilities.sleep(2500); // Mỗi lần chờ 2.5 giây

      var infoRes = apiCall("gettrackinfo", [{"number": trackingNum, "carrier": item.carrier}]);
      if (infoRes && infoRes.code === 0 && infoRes.data) {
        var list = infoRes.data.accepted || infoRes.data;
        if (Array.isArray(list) && list.length > 0) data = list[0];
        else if (list && list.number) data = list;

        if (data && (data.track_info || data.track)) {
          var info = data.track_info || data.track;
          rawStatus = "NotFound";
          if (info.latest_status && info.latest_status.status) {
            rawStatus = info.latest_status.status;
          } else if (info.e !== undefined) {
            var statusCode = parseInt(info.e);
            if (statusCode === 10) rawStatus = "InTransit";
            else if (statusCode === 20) rawStatus = "Expired";
            else if (statusCode === 30) rawStatus = "PickedUp";
            else if (statusCode === 35) rawStatus = "Undelivered";
            else if (statusCode === 40) rawStatus = "Delivered";
            else if (statusCode === 50) rawStatus = "Exception";
          }
          
          if (rawStatus !== "NotFound") {
            break; // Đã lấy được trạng thái thật, thoát vòng lặp ngay lập tức!
          }
        }
      }
    }

    if (!data || (!data.track_info && !data.track)) {
      // Nếu sau 10s vẫn chưa có, đổi màu và chữ để báo hiệu
      sheet.getRange(rowIndex, COL_STATUS).setValue("📄 Chờ cập nhật").setBackground("#e2e8f0").setFontColor("#2d3748");
      return { success: true, message: "✅ Đã đăng ký! Đang chờ hãng cập nhật." };
    }

    var info = data.track_info || data.track;
    
    var statusText = STATUS_MAP[rawStatus] || rawStatus;
    var detailText = "";
    if (info.latest_event) {
      var eventTime = info.latest_event.time_iso ? info.latest_event.time_iso.substring(0, 16).replace("T", " ") : "";
      detailText = eventTime + " — " + (info.latest_event.description || "");
    } else if (info.z0 && info.z0.a) {
      detailText = info.z0.a + " — " + (info.z0.z || info.z0.d || "");
    }

    var carrierName = "";
    if (info.tracking && info.tracking.providers && info.tracking.providers.length > 0 && info.tracking.providers[0].provider) {
      carrierName = info.tracking.providers[0].provider.name;
    }

    var now = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm");

    sheet.getRange(rowIndex, COL_STATUS).setValue(statusText);
    sheet.getRange(rowIndex, COL_DETAIL).setValue(detailText);
    sheet.getRange(rowIndex, COL_UPDATED).setValue(now);

    var bgColor = null;
    var fontColor = null;
    switch (rawStatus) {
      case "Delivered": 
        bgColor = "#c6f6d5"; fontColor = "#22543d"; break;
      case "InTransit":
      case "OutForDelivery":
      case "PickedUp":  
        bgColor = "#fefcbf"; fontColor = "#744210"; break;
      case "Exception":
      case "Undelivered":
      case "NotFound":  
        bgColor = "#fed7d7"; fontColor = "#742a2a"; break;
      default:
        bgColor = "#ffffff"; fontColor = "#000000";
    }
    
    sheet.getRange(rowIndex, COL_STATUS).setBackground(bgColor).setFontColor(fontColor).setFontWeight("bold");

    updateSheetStats();
    return { success: true, message: "✅ Đã tra cứu và cập nhật trạng thái!" };
  } catch (e) {
    return { success: false, message: "❌ Lỗi: " + e.message };
  }
}

// ── Bước 1: Đăng ký ──────────────────────────────────
function registerOnly() {
  try {
    var API_KEY = getApiKey();
    if (!API_KEY || API_KEY.trim() === "") {
      return { success: false, message: "❌ API Key chưa được cài đặt! Vui lòng nhập ở phần Cài đặt." };
    }

    var sheet   = getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "Không có dữ liệu trong sheet!" };

    var items = getItemsFromSheet(sheet, lastRow);
    if (items.length === 0) return { success: false, message: "Không có mã vận đơn nào!" };

    var result = apiCall("register", items);
    Logger.log("Register: " + JSON.stringify(result));

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
function getInfoOnly() {
  try {
    var API_KEY = getApiKey();
    if (!API_KEY || API_KEY.trim() === "") {
      return { success: false, message: "❌ API Key chưa được cài đặt! Vui lòng nhập ở phần Cài đặt." };
    }

    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return { success: false, message: "Bảng tính không có dữ liệu để cập nhật!" };
    }

    var items = getItemsFromSheet(sheet, lastRow);
    if (items.length === 0) {
      return { success: false, message: "Không tìm thấy mã vận đơn nào hợp lệ!" };
    }

    var trackMap = {};
    var BATCH_SIZE = 40;
    var hasApiError = false;
    Logger.log("📤 Gửi " + items.length + " mã tới API...");

    for (var i = 0; i < items.length; i += BATCH_SIZE) {
      var batch = items.slice(i, i + BATCH_SIZE);
      var result = apiCall("gettrackinfo", batch);

      if (!result || result.code !== 0) {
        hasApiError = true;
        continue;
      }

      if (result.data) {
        var list = result.data.accepted || result.data;
        if (Array.isArray(list)) {
          list.forEach(function(obj) { trackMap[obj.number] = obj; });
        } else if (list && typeof list === 'object' && list.number) {
          trackMap[list.number] = list;
        }
      }
      if (i + BATCH_SIZE < items.length) Utilities.sleep(1000);
    }

    if (Object.keys(trackMap).length === 0) {
      var msg = hasApiError ? "❌ API trả về lỗi! Kiểm tra lại API Key." : "❌ Không lấy được dữ liệu từ API!";
      return { success: false, message: msg };
    }

    var now = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm");
    var updatedCount = 0;

    for (var row = 2; row <= lastRow; row++) {
      var trackingNum = sheet.getRange(row, COL_TRACKING).getValue().toString().trim();
      if (!trackingNum || !trackMap[trackingNum]) continue;

      var data = trackMap[trackingNum];
      var info = data.track_info || data.track;
      if (!info) continue;

      var rawStatus = "NotFound";
      if (info.latest_status && info.latest_status.status) {
        rawStatus = info.latest_status.status;
      } else if (info.e !== undefined) {
        var statusCode = parseInt(info.e);
        if (statusCode === 0) rawStatus = "NotFound";
        else if (statusCode === 10) rawStatus = "InTransit";
        else if (statusCode === 20) rawStatus = "Expired";
        else if (statusCode === 30) rawStatus = "PickedUp";
        else if (statusCode === 35) rawStatus = "Undelivered";
        else if (statusCode === 40) rawStatus = "Delivered";
        else if (statusCode === 50) rawStatus = "Exception";
      }
      var statusText = STATUS_MAP[rawStatus] || rawStatus;

      var detailText = "";
      if (info.latest_event) {
        var eventTime = info.latest_event.time_iso ? info.latest_event.time_iso.substring(0, 16).replace("T", " ") : "";
        detailText = eventTime + " — " + (info.latest_event.description || "");
      } else if (info.z0 && info.z0.a) {
        detailText = info.z0.a + " — " + (info.z0.z || info.z0.d || "");
      }
      
      // Trích xuất tên Hãng vận chuyển từ API
      var carrierName = "";
      if (info.tracking && info.tracking.providers && info.tracking.providers.length > 0 && info.tracking.providers[0].provider) {
        carrierName = info.tracking.providers[0].provider.name;
      }

      sheet.getRange(row, COL_STATUS).setValue(statusText);
      sheet.getRange(row, COL_DETAIL).setValue(detailText);
      sheet.getRange(row, COL_UPDATED).setValue(now);

      var bgColor = null;
      var fontColor = null;
      switch (rawStatus) {
        case "Delivered": 
          bgColor = "#c6f6d5"; fontColor = "#22543d"; 
          break;
        case "InTransit":
        case "OutForDelivery":
        case "PickedUp":  
          bgColor = "#fefcbf"; fontColor = "#744210"; 
          break;
        case "Exception":
        case "Undelivered":
        case "NotFound":  
          bgColor = "#fed7d7"; fontColor = "#742a2a"; 
          break;
        default:
          bgColor = "#ffffff"; fontColor = "#000000";
      }
      
      sheet.getRange(row, COL_STATUS)
           .setBackground(bgColor)
           .setFontColor(fontColor)
           .setFontWeight("bold");

      updatedCount++;
    }

    updateSheetStats();
    return { success: true, count: updatedCount };

  } catch (e) {
    return { success: false, message: "Lỗi hệ thống: " + e.message };
  }
}

// ── Tự động cập nhật (trigger) ───────────────────────
function autoUpdate() {
  getInfoOnly();
}

// ── Thống kê ─────────────────────────────────────────
function updateSheetStats() {
  try {
    var sheet   = getSheet();
    var lastRow = sheet.getLastRow();
    var stats   = { total: 0, delivered: 0, transit: 0, error: 0 };
    
    if (lastRow >= 2) {
      var statuses = sheet.getRange(2, COL_STATUS, lastRow - 1, 1).getValues();
      for (var i = 0; i < statuses.length; i++) {
        var s = statuses[i][0].toString().trim();
        if (!s || s === "") continue;
        
        stats.total++;
        var lowerS = s.toLowerCase();
        
        // Đếm chuẩn 100% dựa trên icon được map trực tiếp từ "status" của API
        if (s.indexOf("✅") > -1 || lowerS === "delivered") {
          stats.delivered++;
        } else if (s.indexOf("⚠️") > -1 || s.indexOf("❓") > -1 || s.indexOf("⌛") > -1 || lowerS === "exception" || lowerS === "notfound" || lowerS === "expired") {
          stats.error++;
        } else if (s.indexOf("⏳") > -1 || lowerS.indexOf("chờ cập nhật") > -1) {
          // Bỏ qua đơn mới thêm (chỉ đếm vào tổng)
        } else {
          stats.transit++;
        }
      }
    }
    
    // Ghi thống kê ra sheet (cột I và J)
    sheet.getRange("I1:J1").merge().setValue("📊 THỐNG KÊ")
         .setBackground("#4a5568").setFontColor("#ffffff")
         .setFontWeight("bold").setHorizontalAlignment("center");
         
    var data = [
      ["Tổng đơn", stats.total],
      ["Đã giao", stats.delivered],
      ["Đang đi", stats.transit],
      ["Có vấn đề", stats.error]
    ];
    
    sheet.getRange("I2:J5").setValues(data).setFontWeight("bold");
    
    // Tô màu cho từng ô thống kê tương ứng
    sheet.getRange("I2:J2").setBackground("#e2e8f0").setFontColor("#2d3748"); // Tổng đơn
    sheet.getRange("I3:J3").setBackground("#c6f6d5").setFontColor("#22543d"); // Đã giao
    sheet.getRange("I4:J4").setBackground("#fefcbf").setFontColor("#744210"); // Đang đi
    sheet.getRange("I5:J5").setBackground("#fed7d7").setFontColor("#742a2a"); // Có vấn đề
    
  } catch(e) {
    Logger.log("Lỗi cập nhật thống kê: " + e.message);
  }
}

// ── Helper ───────────────────────────────────────────
function getItemsFromSheet(sheet, lastRow) {
  var rawValues = sheet.getRange(2, COL_TRACKING, lastRow - 1, 2).getValues();
  var items = [];
  for (var i = 0; i < rawValues.length; i++) {
    var num   = rawValues[i][0].toString().trim();
    var phone = rawValues[i][1].toString().trim();
    if (!num) continue;

    var item = { number: num };
    if (phone) item.param = phone;
    items.push(item);
  }
  return items;
}

function apiCall(endpoint, payload) {
  try {
    var API_KEY = getApiKey();
    var res = UrlFetchApp.fetch("https://api.17track.net/track/v2.4/" + endpoint, {
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

    if (statusCode !== 200) return { code: -1, msg: "HTTP Error " + statusCode };
    return JSON.parse(content);
  } catch(e) {
    return { code: -1, msg: "Lỗi kết nối: " + e.message };
  }
}

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

// ── Lấy danh sách hãng vận chuyển (VN & CN) ───────────
function getCarrierList() {
  var list = [
    // --- Việt Nam ---
    { code: "100593", name: "GHN (Giao Hàng Nhanh)" },
    { code: "100611", name: "Viettel Post" },
    { code: "100456", name: "J&T Express (VN)" },
    { code: "100538", name: "ShopeeExpress (VN)" },
    { code: "22041",  name: "VietNam Post" },
    { code: "22043",  name: "VietNam EMS" },
    { code: "100129", name: "Ninjavan (VN)" },
    { code: "100911", name: "Nhất Tín Logistics" },
    { code: "100997", name: "Lazada Logistics (VN)" },
    { code: "101113", name: "247Express" },
    { code: "101194", name: "BEST Inc (VN)" },
    // --- Trung Quốc ---
    { code: "190766", name: "SF Express (顺丰)" },
    { code: "3011",   name: "China Post" },
    { code: "3013",   name: "China EMS" },
    { code: "190157", name: "YTO Express" },
    { code: "190455", name: "ZTO Express" },
    { code: "190324", name: "STO Express" },
    { code: "191197", name: "Yunda Express" },
    { code: "190271", name: "Cainiao" },
    { code: "191121", name: "JD Logistics" },
    { code: "190174", name: "Deppon" },
    { code: "190012", name: "YANWEN" },
    { code: "190008", name: "YunExpress" },
    { code: "190094", name: "4PX" }
  ];
  return list;
}
