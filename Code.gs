/**
 * MÃ GOOGLE APPS SCRIPT (Code.gs) - ĐỀ KIỂM TRA 50 CÂU (100 ĐIỂM)
 * Bài 22: Đại Lượng Tỉ Lệ Thuận - Toán 7 (Kết Nối Tri Thức)
 * Trường THCS và THPT Long Bình – An Phú, An Giang
 *
 * HƯỚNG DẪN CÀI ĐẶT:
 * 1. Mở một file Google Sheets (Trang tính) mới -> Đặt tên tùy ý.
 * 2. Vào menu: Tiện ích mở rộng (Extensions) -> Apps Script.
 * 3. Xóa mã mặc định và dán toàn bộ mã trong file Code.gs này vào -> Bấm Lưu (Ctrl + S).
 * 4. Bấm nút "Triển khai" (Deploy) -> "Tùy chọn triển khai mới" (New deployment).
 * 5. Chọn loại: "Ứng dụng web" (Web app):
 *    - Thực thi dưới dạng (Execute as): Tôi (Me)
 *    - Ai có quyền truy cập (Who has access): Bất kỳ ai (Anyone)
 * 6. Bấm "Triển khai" (Deploy) -> Cấp quyền truy cập -> Sao chép URL Ứng dụng web (Web App URL).
 * 7. Dán URL vừa sao chép vào biến GOOGLE_SCRIPT_URL trong file index.html.
 */

const SHEET_NAME = 'KetQuaHocTap_Bai22_100Diem';

/**
 * Xử lý yêu cầu GET:
 * - Kiểm tra kết nối hoặc lấy danh sách bảng điểm/xếp hạng.
 */
function doGet(e) {
  try {
    const sheet = getOrCreateSheet_();
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'status';

    if (action === 'list') {
      const data = sheet.getDataRange().getValues();
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success', rows: data }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Kết nối Google Apps Script (Bài 22: Đại lượng tỉ lệ thuận - 50 câu / 100 điểm) thành công!'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Xử lý yêu cầu POST:
 * Nhận đầy đủ thông tin học sinh, điểm tổng từng loại trắc nghiệm, điểm tổng tự luận,
 * điểm tổng của học sinh (thang 100 điểm / 50 câu) và xếp hạng từ index.html.
 */
function doPost(e) {
  try {
    const sheet = getOrCreateSheet_();
    let payload = {};

    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
    const studentName = payload.studentName || 'Chưa rõ';
    const studentClass = payload.studentClass || 'Chưa rõ';

    // Điểm từng loại Trắc nghiệm (mỗi loại 10 câu x 2đ = 20đ)
    const scoreMCQ = Number(payload.scoreMCQ) || 0;
    const scoreTF = Number(payload.scoreTF) || 0;
    const scoreMatch = Number(payload.scoreMatch) || 0;
    const scoreShort = Number(payload.scoreShort) || 0;
    const totalQuizScore = payload.totalQuizScore !== undefined
      ? Number(payload.totalQuizScore)
      : (scoreMCQ + scoreTF + scoreMatch + scoreShort);

    // Điểm Tự luận (10 câu x 2đ = 20đ)
    const scoreEssay = Number(payload.scoreEssay) || 0;

    // Điểm tổng của học sinh (50 câu x 2đ = 100đ)
    const totalScore = payload.totalScore !== undefined
      ? Number(payload.totalScore)
      : (totalQuizScore + scoreEssay);

    const maxScore = payload.maxScore || 100;
    const totalQuestions = payload.totalQuestions || 50;
    const correctCount = payload.correctCount !== undefined
      ? `${payload.correctCount}/${totalQuestions}`
      : `${Math.round(totalScore / 2)}/${totalQuestions}`;

    const rank = payload.rank || classifyRank100_(totalScore);
    const tasksDone = payload.tasksDone !== undefined ? payload.tasksDone : 0;
    const totalTasks = payload.totalTasks || 7;
    const progressPct = payload.progressPct !== undefined ? payload.progressPct + '%' : '0%';
    const taskDetails = payload.taskDetails || '';

    sheet.appendRow([
      timestamp,
      studentName,
      studentClass,
      `${scoreMCQ}/20`,
      `${scoreTF}/20`,
      `${scoreMatch}/20`,
      `${scoreShort}/20`,
      `${totalQuizScore}/80`,
      `${scoreEssay}/20`,
      `${totalScore}/${maxScore}`,
      correctCount,
      rank,
      `${tasksDone}/${totalTasks} (${progressPct})`,
      taskDetails
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Đã lưu bảng điểm 100đ của học sinh ' + studentName + ' (' + studentClass + ') thành công!'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Tự động tạo Sheet và định dạng hàng tiêu đề (14 cột chi tiết) nếu chưa tồn tại.
 */
function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    const headers = [
      'Thời gian nộp bài',
      'Họ và tên học sinh',
      'Lớp',
      'I. TN Nhiều lựa chọn (20đ)',
      'II. TN Đúng/Sai (20đ)',
      'III. TN Ghép cột (20đ)',
      'IV. TN Trả lời ngắn (20đ)',
      'Tổng điểm Trắc nghiệm (80đ)',
      'V. Tổng điểm Tự luận (20đ)',
      'TỔNG ĐIỂM HỌC SINH (100đ)',
      'Số câu đúng (/50)',
      'Xếp hạng',
      'Tiến trình bài học',
      'Chi tiết nhiệm vụ hoàn thành'
    ];
    sheet.appendRow(headers);

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4f46e5');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // Độ rộng cột tự động dễ nhìn
    const widths = [155, 190, 80, 165, 145, 145, 165, 180, 170, 185, 125, 190, 150, 260];
    widths.forEach((w, idx) => sheet.setColumnWidth(idx + 1, w));
  }

  return sheet;
}

/**
 * Xếp hạng học tập dựa trên Tổng điểm (thang 100 điểm).
 */
function classifyRank100_(score) {
  const s = Number(score) || 0;
  if (s >= 90) return 'Xuất Sắc (Bậc Thầy Tỉ Lệ Thuận)';
  if (s >= 80) return 'Giỏi (Nắm Rất Vững Kiến Thức)';
  if (s >= 65) return 'Khá (Hiểu Bản Chất Toán 7)';
  if (s >= 50) return 'Đạt Yêu Cầu (Cần Rèn Thêm)';
  return 'Cần Cố Gắng (Ôn Lại Lý Thuyết)';
}
