/**
 * THE CAKE LAB — TIÊU CHUẨN SẢN PHẨM
 * Backend Apps Script: ghi dữ liệu vào Google Sheet + lưu ảnh vào Google Drive
 *
 * CÁCH CÀI ĐẶT:
 * 1. Mở Google Sheet: https://docs.google.com/spreadsheets/d/1O8YO4NMOYkuFqImi8f3BU1lssYFbfzNZh3OWq8mvhV0/edit
 * 2. Menu: Extensions (Tiện ích mở rộng) > Apps Script
 * 3. Xoá hết code mẫu, dán toàn bộ nội dung file này vào
 * 4. Bấm Save (biểu tượng đĩa mềm, hoặc Ctrl+S)
 * 5. Ở thanh công cụ trên cùng, chọn hàm "setup" trong ô dropdown (thay vì "doGet"),
 *    rồi bấm nút ▶ Run (Chạy)
 *    - Lần đầu chạy, Google sẽ hiện màn hình xin cấp quyền:
 *      chọn tài khoản của bạn > "Advanced" (Nâng cao) > "Go to ... (unsafe)" > "Allow" (Cho phép)
 *    - Sau khi chạy xong, mở "Executions" (Nhật ký thực thi) hoặc bấm Ctrl+Enter để xem log,
 *      xác nhận thấy dòng "Đã sẵn sàng" — nghĩa là sheet + thư mục Drive đã được tạo thành công
 * 6. Bấm Deploy > New deployment (nếu đây là lần đầu deploy)
 *    - Chọn loại: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7. Bấm Deploy, cấp quyền (Authorize) lần nữa nếu được hỏi
 * 8. Copy "Web app URL" (dạng https://script.google.com/macros/s/xxxxx/exec)
 *    — dán vào biến APPS_SCRIPT_URL trong file index.html/app.html
 *
 * LƯU Ý: Nếu đây là lần cập nhật code cho một deployment đã có sẵn (đã có URL cũ đang dùng),
 * hãy dùng "Manage deployments" (Quản lý bản triển khai) > biểu tượng bút chì (Edit) >
 * chọn "New version" > Deploy, để giữ nguyên URL cũ, khỏi phải sửa lại HTML.
 */

var SPREADSHEET_ID = '1O8YO4NMOYkuFqImi8f3BU1lssYFbfzNZh3OWq8mvhV0';
var SHEET_NAME = 'TieuChuanSanPham';
var DRIVE_FOLDER_NAME = 'TheCakeLab_TieuChuanSanPham_Images';

var HEADERS = [
  'STT', 'Tên Sản Phẩm', 'Nhóm Sản Phẩm', 'Quy Cách Đóng', 'Mô Tả Sản Phẩm',
  'Kích Thước', 'Hạn Sử Dụng', 'Cách Bảo Quản',
  'Hình Ảnh Tổng Thể', 'Hình Ảnh Mặt Trước', 'Hình Ảnh Mặt Ngang', 'Ngày Tạo'
];

/**
 * Chạy hàm này TRƯỚC TIÊN (thủ công, 1 lần) để:
 * - Xin cấp quyền truy cập Sheet + Drive
 * - Tự tạo sheet con "TieuChuanSanPham" (nếu chưa có) và thư mục Drive chứa ảnh
 * Chọn "setup" ở dropdown trên thanh công cụ Apps Script rồi bấm ▶ Run.
 */
function setup() {
  var sheet = getOrCreateSheet_();
  var folder = getOrCreateFolder_();
  Logger.log('Đã sẵn sàng.');
  Logger.log('Sheet: ' + sheet.getParent().getUrl() + ' (tab: ' + sheet.getName() + ')');
  Logger.log('Thư mục ảnh trên Drive: ' + folder.getUrl());
}

function getSpreadsheet_() {
  // openById hoạt động cả khi script được gắn trực tiếp vào Sheet (bound)
  // lẫn khi chạy như một script độc lập (standalone).
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreateSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateFolder_() {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function saveImage_(imgObj, folder) {
  if (!imgObj || !imgObj.data) return '';
  var bytes = Utilities.base64Decode(imgObj.data);
  var blob = Utilities.newBlob(bytes, imgObj.mimeType || 'image/jpeg', imgObj.filename || 'anh.jpg');
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Dùng định dạng lh3.googleusercontent.com thay vì "uc?export=view":
  // link "uc?export=view" hay bị Google chặn hiển thị trực tiếp (hiện trang cảnh báo virus-scan)
  // khi nhúng làm <img src>, đặc biệt trên di động — gây ra tình trạng "không xem được ảnh".
  return 'https://lh3.googleusercontent.com/d/' + file.getId() + '=w1600';
}

// Lấy Drive file ID từ 1 link ảnh đã lưu (dạng lh3.googleusercontent.com/d/xxxx=w1600 hoặc link cũ)
function extractDriveFileId_(url) {
  if (!url) return null;
  var m = String(url).match(/\/d\/([-\w]{20,})/);
  return m ? m[1] : null;
}

// Chuyển 1 file ảnh trên Drive vào thùng rác — dùng khi xoá sản phẩm hoặc thay ảnh mới.
// Bọc trong try/catch vì file có thể đã bị xoá tay từ trước, không nên làm hỏng cả request.
function trashDriveFile_(url) {
  var id = extractDriveFileId_(url);
  if (!id) return;
  try {
    DriveApp.getFileById(id).setTrashed(true);
  } catch (err) {
    // bỏ qua: file có thể đã không còn tồn tại
  }
}

// Quyết định giá trị link ảnh cuối cùng khi update 1 sản phẩm:
// - newVal là object {data,...} (ảnh mới người dùng vừa chọn) -> lưu ảnh mới, xoá ảnh cũ trên Drive
// - newVal là string (client trả lại link ảnh cũ, không đổi) -> giữ nguyên
// - newVal là null/'' (người dùng bấm xoá ảnh khỏi ô) -> xoá ảnh cũ trên Drive, để trống
function resolveImageField_(newVal, oldUrl, folder) {
  if (newVal && typeof newVal === 'object' && newVal.data) {
    if (oldUrl) trashDriveFile_(oldUrl);
    return saveImage_(newVal, folder);
  }
  if (typeof newVal === 'string') {
    return newVal;
  }
  if (oldUrl) trashDriveFile_(oldUrl);
  return '';
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * (Tuỳ chọn) Chạy 1 LẦN nếu bạn đã có sẵn sản phẩm với ảnh bị lỗi "không xem được"
 * (do dùng link cũ dạng uc?export=view). Hàm này quét cả sheet, tìm các ô link ảnh
 * dạng cũ và tự động đổi sang định dạng mới (lh3.googleusercontent.com) — không cần
 * tải lại ảnh, không tạo file Drive mới, chỉ sửa lại đường link trong Sheet.
 * Chọn "fixExistingImageLinks" ở dropdown rồi bấm ▶ Run.
 */
function fixExistingImageLinks() {
  var sheet = getOrCreateSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Chưa có dữ liệu để sửa.'); return; }

  var range = sheet.getRange(2, 9, lastRow - 1, 3); // cột I, J, K
  var values = range.getValues();
  var changed = 0;

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < 3; j++) {
      var url = values[i][j];
      if (!url) continue;
      var match = String(url).match(/[-\w]{20,}/); // lấy Drive file ID trong link cũ
      if (match && url.indexOf('lh3.googleusercontent.com') === -1) {
        values[i][j] = 'https://lh3.googleusercontent.com/d/' + match[0] + '=w1600';
        changed++;
      }
    }
  }

  range.setValues(values);
  Logger.log('Đã sửa ' + changed + ' link ảnh.');
}

// Tìm số thứ tự dòng trong sheet (1-based, kể cả header) ứng với 1 giá trị STT.
// Dò theo giá trị cột STT thay vì suy ra vị trí từ công thức, vì sau khi xoá 1 dòng,
// STT của các dòng còn lại không còn khớp với vị trí dòng trong sheet nữa.
function findRowBySTT_(sheet, stt) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var sttColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < sttColumn.length; i++) {
    if (Number(sttColumn[i][0]) === Number(stt)) return i + 2;
  }
  return -1;
}

function doGet(e) {
  try {
    var sheet = getOrCreateSheet_();
    var lastRow = sheet.getLastRow();
    var data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues() : [];

    var rows = data.map(function (r) {
      return {
        stt: r[0],
        tenSanPham: r[1],
        nhomSanPham: r[2],
        quyCachDong: r[3],
        moTaSanPham: r[4],
        kichThuoc: r[5],
        hanSuDung: r[6],
        cachBaoQuan: r[7],
        hinhAnhTongThe: r[8],
        hinhAnhMatTruoc: r[9],
        hinhAnhMatNgang: r[10],
        ngayTao: r[11] instanceof Date ? r[11].toISOString() : r[11]
      };
    });

    return jsonOut_({ success: true, data: rows });
  } catch (err) {
    return jsonOut_({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action || 'add';

    if (action === 'update') return handleUpdate_(payload);
    if (action === 'delete') return handleDelete_(payload);
    return handleAdd_(payload);
  } catch (err) {
    return jsonOut_({ success: false, error: err.toString() });
  }
}

function handleAdd_(payload) {
  var sheet = getOrCreateSheet_();
  var folder = getOrCreateFolder_();

  var urlTongThe = saveImage_(payload.hinhAnhTongThe, folder);
  var urlMatTruoc = saveImage_(payload.hinhAnhMatTruoc, folder);
  var urlMatNgang = saveImage_(payload.hinhAnhMatNgang, folder);

  var stt = sheet.getLastRow(); // header chiếm dòng 1, nên STT = lastRow hiện tại (trước khi thêm)

  sheet.appendRow([
    stt,
    payload.tenSanPham || '',
    payload.nhomSanPham || '',
    payload.quyCachDong || '',
    payload.moTaSanPham || '',
    payload.kichThuoc || '',
    payload.hanSuDung || '',
    payload.cachBaoQuan || '',
    urlTongThe,
    urlMatTruoc,
    urlMatNgang,
    new Date()
  ]);

  return jsonOut_({ success: true, stt: stt });
}

function handleUpdate_(payload) {
  var sheet = getOrCreateSheet_();
  var folder = getOrCreateFolder_();

  var rowIndex = findRowBySTT_(sheet, payload.stt);
  if (rowIndex === -1) {
    return jsonOut_({ success: false, error: 'Không tìm thấy sản phẩm với STT ' + payload.stt });
  }

  var existing = sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];

  var urlTongThe = resolveImageField_(payload.hinhAnhTongThe, existing[8], folder);
  var urlMatTruoc = resolveImageField_(payload.hinhAnhMatTruoc, existing[9], folder);
  var urlMatNgang = resolveImageField_(payload.hinhAnhMatNgang, existing[10], folder);

  sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([[
    existing[0], // STT giữ nguyên
    payload.tenSanPham || '',
    payload.nhomSanPham || '',
    payload.quyCachDong || '',
    payload.moTaSanPham || '',
    payload.kichThuoc || '',
    payload.hanSuDung || '',
    payload.cachBaoQuan || '',
    urlTongThe,
    urlMatTruoc,
    urlMatNgang,
    existing[11] // Ngày Tạo giữ nguyên
  ]]);

  return jsonOut_({ success: true, stt: existing[0] });
}

function handleDelete_(payload) {
  var sheet = getOrCreateSheet_();

  var rowIndex = findRowBySTT_(sheet, payload.stt);
  if (rowIndex === -1) {
    return jsonOut_({ success: false, error: 'Không tìm thấy sản phẩm với STT ' + payload.stt });
  }

  var existing = sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];
  trashDriveFile_(existing[8]);
  trashDriveFile_(existing[9]);
  trashDriveFile_(existing[10]);

  sheet.deleteRow(rowIndex);

  return jsonOut_({ success: true });
}

/**
 * (Tuỳ chọn) Chạy thử hàm này thủ công để kiểm tra doPost hoạt động đúng
 * mà không cần deploy trước — sẽ thêm 1 dòng test vào sheet.
 * Chọn "testThemSanPhamMau" ở dropdown rồi bấm ▶ Run để thử.
 */
function testThemSanPhamMau() {
  var fakeEvent = {
    postData: {
      contents: JSON.stringify({
        tenSanPham: 'Bánh sinh nhật Xoài Cheese (test)',
        nhomSanPham: 'Bánh Sinh Nhật',
        quyCachDong: 'Hộp giấy, 1 cái/hộp',
        moTaSanPham: 'Thành phần, hương vị, màu sắc... (dữ liệu test)',
        kichThuoc: 'Đường kính: 20cm, Chiều cao: 10cm',
        hanSuDung: '3 ngày kể từ ngày sản xuất',
        cachBaoQuan: 'Bảo quản ngăn mát 2-5°C',
        hinhAnhTongThe: null,
        hinhAnhMatTruoc: null,
        hinhAnhMatNgang: null
      })
    }
  };
  var result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
