const SPREADSHEET_ID = "1qndWAhAdmWOc8AKE9mtn6utsrXf0wVKYcLX6muUbsec";
const ORDER_SHEET = "Đơn hàng";
const VOUCHER_SHEET = "Voucher";

function doGet(e) {
  try {
    ensureVoucherSheet();

    const action = (e && e.parameter && e.parameter.action) || "health";
    let result;

    if (action === "validateVoucher") {
      const code = String(e.parameter.code || "").trim().toUpperCase();
      const subtotal = Number(e.parameter.subtotal || 0);
      result = validateVoucher(code, subtotal);
    } else {
      result = {
        success: true,
        message: "OLALA Orders API đang hoạt động"
      };
    }

    return outputResult(result, e && e.parameter && e.parameter.callback);
  } catch (error) {
    console.error(error.stack || error);
    return outputResult({
      success: false,
      message: String(error.message || error)
    }, e && e.parameter && e.parameter.callback);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Không nhận được dữ liệu đơn hàng");
    }

    const data = JSON.parse(e.postData.contents);
    const customer = data.customer || {};
    const items = data.items || [];

    if (!customer.name || !customer.phone || !customer.address) {
      throw new Error("Thiếu thông tin khách hàng");
    }

    if (!items.length) {
      throw new Error("Đơn hàng không có sản phẩm");
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ensureOrderSheet(spreadsheet);
    ensureVoucherSheet();

    const subtotal = items.reduce(function (sum, item) {
      return sum + Number(item.unitPrice || 0) * Number(item.quantity || 0);
    }, 0);

    const shipping = subtotal > 0 && subtotal < 150000 ? 15000 : 0;
    const voucherResult = data.voucherCode
      ? validateVoucher(String(data.voucherCode).toUpperCase(), subtotal)
      : { success: false, discount: 0, percent: 0, code: "" };

    const discount = voucherResult.success
      ? Number(voucherResult.discount || 0)
      : 0;

    const total = Math.max(0, subtotal + shipping - discount);
    const now = new Date();
    const orderId =
      "OLA-" +
      Utilities.formatDate(now, "Asia/Ho_Chi_Minh", "yyyyMMdd-HHmmss");

    const itemText = items.map(function (item) {
      let text =
        item.name +
        " x" +
        item.quantity +
        " | Size " +
        (item.size || "M") +
        " | " +
        (item.sugar || "50%") +
        " đường | Đá " +
        (item.ice || "Vừa");

      if (item.toppings && item.toppings.length) {
        text += " | Topping: " + item.toppings.join(", ");
      }

      return text;
    }).join("\n");

    sheet.appendRow([
      orderId,
      now,
      customer.name,
      customer.phone,
      customer.address,
      customer.note || "",
      itemText,
      subtotal,
      shipping,
      total,
      data.payment === "online"
        ? "Thanh toán trực tuyến"
        : "Thanh toán khi nhận hàng",
      "Đơn mới",
      voucherResult.success ? voucherResult.code : "",
      voucherResult.success ? voucherResult.percent : 0,
      discount
    ]);

    SpreadsheetApp.flush();

    return outputResult({
      success: true,
      orderId: orderId,
      subtotal: subtotal,
      shipping: shipping,
      discount: discount,
      total: total
    });
  } catch (error) {
    console.error(error.stack || error);

    return outputResult({
      success: false,
      message: String(error.message || error)
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {}
  }
}

function ensureOrderSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(ORDER_SHEET);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(ORDER_SHEET);
  }

  const headers = [
    "Mã đơn",
    "Thời gian",
    "Họ và tên",
    "Số điện thoại",
    "Địa chỉ",
    "Ghi chú",
    "Chi tiết đơn hàng",
    "Tạm tính",
    "Phí giao hàng",
    "Tổng thanh toán",
    "Phương thức thanh toán",
    "Trạng thái",
    "Mã voucher",
    "Giảm (%)",
    "Số tiền giảm"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#173b2a")
    .setFontColor("#ffffff");

  sheet.setFrozenRows(1);
  return sheet;
}

function ensureVoucherSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(VOUCHER_SHEET);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(VOUCHER_SHEET);
    sheet.appendRow([
      "Mã voucher",
      "Giảm (%)",
      "Đơn tối thiểu",
      "Giảm tối đa",
      "Ngày hết hạn",
      "Trạng thái"
    ]);

    sheet.appendRow([
      "WELCOME10",
      10,
      100000,
      50000,
      new Date(2026, 11, 31),
      "Bật"
    ]);

    sheet
      .getRange(1, 1, 1, 6)
      .setFontWeight("bold")
      .setBackground("#ff7a3d")
      .setFontColor("#ffffff");

    sheet.setFrozenRows(1);
  }

  return sheet;
}

function validateVoucher(code, subtotal) {
  if (!code) {
    return {
      success: false,
      message: "Vui lòng nhập mã voucher"
    };
  }

  const sheet = ensureVoucherSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      success: false,
      message: "Voucher không tồn tại"
    };
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  let voucher = null;

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toUpperCase() === code) {
      voucher = rows[i];
      break;
    }
  }

  if (!voucher) {
    return {
      success: false,
      message: "Voucher không tồn tại"
    };
  }

  const percent = Number(voucher[1] || 0);
  const minOrder = Number(voucher[2] || 0);
  const maxDiscount = Number(voucher[3] || 0);
  const expiry = voucher[4];
  const status = String(voucher[5] || "").trim().toLowerCase();
  const enabled = ["bật", "bat", "active", "true", "1", "yes"].indexOf(status) >= 0;

  if (!enabled) {
    return {
      success: false,
      message: "Voucher đang tạm khóa"
    };
  }

  if (percent <= 0 || percent > 100) {
    return {
      success: false,
      message: "Mức giảm voucher không hợp lệ"
    };
  }

  if (expiry) {
    const expiryDate = expiry instanceof Date ? expiry : new Date(expiry);
    expiryDate.setHours(23, 59, 59, 999);

    if (!isNaN(expiryDate.getTime()) && new Date() > expiryDate) {
      return {
        success: false,
        message: "Voucher đã hết hạn"
      };
    }
  }

  if (subtotal < minOrder) {
    return {
      success: false,
      message: "Đơn hàng cần tối thiểu " +
        Utilities.formatString("%,.0f", minOrder) +
        "đ"
    };
  }

  let discount = Math.round(subtotal * percent / 100);

  if (maxDiscount > 0) {
    discount = Math.min(discount, maxDiscount);
  }

  return {
    success: true,
    code: code,
    percent: percent,
    minOrder: minOrder,
    maxDiscount: maxDiscount,
    discount: discount
  };
}

function outputResult(data, callback) {
  if (callback) {
    const safeCallback = String(callback).replace(/[^a-zA-Z0-9_$]/g, "");
    return ContentService
      .createTextOutput(safeCallback + "(" + JSON.stringify(data) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
