function createOrUpdateSheetLuong() {
  var ui = SpreadsheetApp.getUi(); // Lấy đối tượng UI

  // sheet cấu hình
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Hiển thị hộp thoại xác nhận
  var response = ui.alert("Xác nhận", "Bạn có muốn tạo bảng lương cho tháng " + sheet.getRange("A1").getValue() + "?", ui.ButtonSet.YES_NO);

  // Kiểm tra phản hồi của người dùng
  if (response == ui.Button.YES) {
    var lastColumn = sheet.getLastColumn();
    if (lastColumn > 17) {
      var rangeUnMerge = sheet.getRange(3, 15, 2, lastColumn);
      rangeUnMerge.breakApart();
      rangeUnMerge.clearContent().clearFormat().setBackground("#ffffff").setBorder(false, false, false, false, false, false).setFontFamily("Arial").setFontSize(12);
    }

    sheet.setColumnWidth(15, 120)
      .setColumnWidth(16, 150)
      .setColumnWidth(17, 130)
      .setColumnWidth(18, 130)
      .setColumnWidth(19, 150)
      .setColumnWidth(20, 150)
      .setColumnWidth(21, 150)
      .setColumnWidth(22, 150)
      .setColumnWidth(23, 150)
      .setColumnWidth(24, 150)
      .setColumnWidth(25, 150)
      .setColumnWidth(26, 150);
    const cellTitle = sheet.getRange("o1");
    let title = "Lương " + sheet.getRange("A1").getValue();
    cellTitle.setValue(title);
    cellTitle.setFontWeight("bold");
    cellTitle.setFontSize(17).setHorizontalAlignment("left");
    var values = [["Ngày", "Tên khách", "Tiền bill", "Tổng bill ngày"]];

    const headerCommons = sheet.getRange(3, 15, 2, 4);
    for (var col = 15; col <= 18; col++) {
      // Gộp ô tại hàng 3 và hàng 4 cho mỗi cột
      sheet.getRange(3, col, 2, 1).merge(); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
      sheet.getRange(3, col).setValue(values[0][col - 15]);
      sheet.getRange(3, col).setVerticalAlignment("middle").setHorizontalAlignment("center");
    }
    headerCommons.setFontWeight("bold");
    headerCommons.setFontSize(12);
    headerCommons.setBackground("#d3d3d3");
    headerCommons.setBorder(true, true, true, true, true, true); // Đặt đường viền cho các cạnh trên, dưới, trái, phải

    let lastRowThoChinh = sheet.getRange(sheet.getMaxRows(), 10).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    let lastRowThoPhu = sheet.getRange(sheet.getMaxRows(), 12).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

    let colThoChinh = sheet.getRange("J4:J" + lastRowThoChinh);
    let thoChinh = colThoChinh.getValues();

    let colThoPhu = sheet.getRange("L4:L" + lastRowThoPhu);
    let thoPhu = colThoPhu.getValues();

    var startCellTho = sheet.getRange("S4"); // Lấy ô E4
    let coutTho = 0;

    for (let i = 0; i < thoChinh.length; i++) {
      startCellTho.offset(0, coutTho).setValue(thoChinh[i])
        .setFontWeight("bold")
        .setFontSize(12)
        .setBackground(colThoChinh.getCell(i + 1, 1).getBackground())
        .setBorder(true, true, true, true, true, true)
        .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang
      coutTho += 1;
    }
    sheet.getRange(3, 19, 1, thoChinh.length).merge(); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
    sheet.getRange(3, 19).setValue("Thợ chính")
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground("#d3d3d3")
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang

    for (let i = 0; i < thoPhu.length; i++) {
      startCellTho.offset(0, coutTho).setValue(thoPhu[i])
        .setFontWeight("bold")
        .setFontSize(12)
        .setBackground(colThoPhu.getCell(i + 1, 1).getBackground())
        .setBorder(true, true, true, true, true, true)
        .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang
      coutTho += 1;
    }
    sheet.getRange(3, 19 + thoChinh.length, 1, thoPhu.length).merge(); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
    sheet.getRange(3, 19 + thoChinh.length).setValue("Thợ phụ")
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground("#d3d3d3")
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang
    sheet.getRange(3, 19 + thoChinh.length + thoPhu.length, 2, 1).merge()
      .setValue("Ngày sửa")
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground("#d3d3d3")
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
   // ui.alert("Tạo bảng lương thành công.");
  } else {
  }

}

function tinhLuong() {
  var ui = SpreadsheetApp.getUi(); // Lấy đối tượng UI

  let dongBatDauLuong = 5;
  let dongBatDauDT = 5;
  let startColTho = 19;
  let tongDoanhThu = 0;
  let cotBatDauLuong = 15;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRowSheet = sheet.getLastRow(); // Lấy dòng cuối cùng có dữ liệu của toàn bộ sheet

  // Hiển thị hộp thoại xác nhận
  var response = ui.alert("Xác nhận", "Bạn có muốn tính lương cho tháng " + sheet.getRange("A1").getValue() + "?", ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {

    // clear data + format
    var dataGetLastRowLuong = sheet.getRange("Q5:Q" + lastRowSheet).getValues(); // Lấy dữ liệu cột Q
    var lastRowLuong = 4;

    for (var i = 0; i < dataGetLastRowLuong.length; i++) {
      if (dataGetLastRowLuong[i] && dataGetLastRowLuong[i][0] !== "") { 
        lastRowLuong = lastRowLuong + 1; // Đếm số dòng có dữ liệu
      }
    }

    var dataGetLastRow = sheet.getRange("A4:A" + lastRowSheet).getValues(); // Lấy dữ liệu cột A
    var lastRow = dongBatDauDT - 1;
    for (var i = 0; i < dataGetLastRow.length; i++) {
      if (dataGetLastRow[i] && dataGetLastRow[i][0] !== "") { 
        lastRow = lastRow + 1; // Đếm số dòng có dữ liệu
      }
    }

    Logger.log('Dong cuoi cung ' + lastRow)

    if (lastRowLuong > 8) { // nếu đã fill dữ liệu lương
      // clear dòng tổng kết
      sheet.getRange(lastRowLuong, 15, 2, 20).clearContent().clearFormat().setBackground("#ffffff").setBorder(false, false, false, false, false, false).setFontFamily("Arial").setFontSize(12).setHorizontalAlignment("center");
      // clear merge của cột ngày
      sheet.getRange(dongBatDauLuong, 15, lastRowLuong, 1).breakApart();
      sheet.getRange(dongBatDauLuong, 18, lastRowLuong, 1).breakApart();
    }

    // Doanh thu - Ngày
    const columnDateDT = sheet.getRange("A" + dongBatDauDT + ":A" + lastRow);
    const dateDT = columnDateDT.getValues();
    // Doanh thu - tên khách
    const columnCustomerDT = sheet.getRange("B" + dongBatDauDT + ":B" + lastRow);
    const customerDT = columnCustomerDT.getValues();
    // Doanh thu - tiền bill
    const columnBillDT = sheet.getRange("C" + dongBatDauDT + ":C" + lastRow);
    const billDT = columnBillDT.getValues();
    // Doanh thu - thợ chính
    const columnThoChinhDT = sheet.getRange("E" + dongBatDauDT + ":E" + lastRow);
    const thoChinhDT = columnThoChinhDT.getValues();
    // Doanh thu - thợ phụ
    const columnThoPhuDT = sheet.getRange("F" + dongBatDauDT + ":F" + lastRow);
    const thoPhuDT = columnThoPhuDT.getValues();

    // Doanh thu - trạng thái tính lương
    const columnStatusDT = sheet.getRange("H" + dongBatDauDT + ":H" + lastRow);
    const statusDT = columnStatusDT.getValues();

    // Lương - ngày
    const columnDateL = sheet.getRange("O" + dongBatDauLuong + ":O" + lastRow + 1);
    const dateL = columnDateL.getValues();
    // Lương - tên khách
    const columnCustomerL = sheet.getRange("P" + dongBatDauLuong + ":P" + lastRow + 1);
    const customerL = columnCustomerL.getValues();
    // Lương - tiền bill
    const columnBillL = sheet.getRange("Q" + dongBatDauLuong + ":Q" + lastRow + 1);
    columnBillL.setNumberFormat("#,##0");
    const billL = columnBillL.getValues();

    // Danh sách thợ
    let listThoChinh = new Array();
    let listThoPhu = new Array();

    let lastRowThoChinh = sheet.getRange(sheet.getMaxRows(), 10).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    let lastRowThoPhu = sheet.getRange(sheet.getMaxRows(), 12).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

    let colThoChinh = sheet.getRange("J4:J" + lastRowThoChinh);
    let thoChinh = colThoChinh.getValues();
    let colLuongThoChinh = sheet.getRange("K4:K" + lastRowThoChinh);
    let luongThoChinh = colLuongThoChinh.getValues();

    let colThoPhu = sheet.getRange("L4:L" + lastRowThoPhu);
    let thoPhu = colThoPhu.getValues();
    let colLuongThoPhu = sheet.getRange("M4:M" + lastRowThoPhu);
    let luongThoPhu = colLuongThoPhu.getValues();

    var headerThoChinh = sheet.getRange(4, 19, 1, thoChinh.length).getValues()[0];
    var headerThoPhu = sheet.getRange(4, 19 + thoChinh.length, 1, thoPhu.length).getValues()[0];

    // start lấy thông tin thợ chính
    for (let i = 0; i < thoChinh.length; i++) {
      listThoChinh.push({
        name: thoChinh[i][0],
        luong: luongThoChinh[i][0],
        index: headerThoChinh.indexOf(thoChinh[i][0]) + startColTho,
        color: colThoChinh.getCell(i + 1, 1).getBackground(),
        tongLuong: 0
      })
    }
    // end lấy thông tin thợ chính
    // start lấy thông tin thợ phụ
    for (let i = 0; i < thoPhu.length; i++) {
      listThoPhu.push({
        name: thoPhu[i][0],
        luong: luongThoPhu[i][0],
        index: headerThoPhu.indexOf(thoPhu[i][0]) + startColTho + thoChinh.length,
        color: colThoPhu.getCell(i + 1, 1).getBackground(),
        tongLuong: 0
      })
    }
    // end lấy thông tin thợ phụ
    //start fill data
    for (let i = 0; i < dateDT.length; i++) {
      if (dateDT[i] != undefined && dateDT[i] != "" && statusDT[i][0] == 0) {

        // reset value dòng
        sheet.getRange(dongBatDauLuong + i, 15, 1, 20).clearContent().clearFormat().setBackground("#ffffff").setBorder(false, false, false, false, false, false).setFontFamily("Arial").setFontSize(12);

        dateL[i] = dateDT[i];
        columnDateL.getCell(i + 1, 1).setFontWeight("bold");
        customerL[i] = customerDT[i];
        billDT[i][0] = billDT[i][0] * 1000;
        billL[i] = billDT[i];

        const curentThoChinh = listThoChinh.find(item => item.name == thoChinhDT[i]);
        if (curentThoChinh) {
          sheet.getRange(dongBatDauLuong + i, curentThoChinh.index)
            .setValue(billDT[i] / 100 * curentThoChinh.luong)
            .setFontSize("12")
            .setBackground(curentThoChinh.color)
            .setNumberFormat("#,##0");
        }

        const curentThoPhu = listThoPhu.find(item => item.name == thoPhuDT[i]);
        if (curentThoPhu) {
          if (customerDT[i][0].toLowerCase() == "bsp" || customerDT[i][0].toLowerCase() == "gội" || customerDT[i][0].toLowerCase() == "cắt") {
            sheet.getRange(dongBatDauLuong + i, curentThoPhu.index)
              .setValue(billDT[i] / 100 * 20)
              .setFontSize("12")
              .setBackground(curentThoPhu.color)
              .setNumberFormat("#,##0");
          } else {
            sheet.getRange(dongBatDauLuong + i, curentThoPhu.index)
              .setValue(billDT[i] / 100 * curentThoPhu.luong)
              .setFontSize("12")
              .setBackground(curentThoPhu.color)
              .setNumberFormat("#,##0");
          }
        }
        var currentDate = new Date(); // Lấy ngày hiện tại
        sheet.getRange(dongBatDauLuong + i, startColTho + listThoPhu.length + listThoChinh.length).setValue(Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "HH:MM dd/MM/yyyy")).setFontSize("12");
        statusDT[i] = [0];
      }
    }

    // end fill data
    // set data
    columnDateL.setValues(dateL).setFontSize("12");
    columnCustomerL.setValues(customerL).setFontSize("12");
    columnBillL.setValues(billL).setFontSize("12");
    columnStatusDT.setValues(statusDT);
    // end set data

    // start tinh tong luong
    for (let i = 0; i < listThoChinh.length; i++) {
      listThoChinh[i].tongLuong = 0;
      var listLuong = sheet.getRange(5, listThoChinh[i].index, lastRow, 1).getValues();
      if (listLuong) {
        for (let l = 0; l < listLuong.length; l++) {
          if (listLuong[l]) {
            listThoChinh[i].tongLuong += Number(listLuong[l]);
          }
        }
      }
    }
    for (let i = 0; i < listThoPhu.length; i++) {
      listThoPhu[i].tongLuong = 0;
      var listLuong = sheet.getRange(5, listThoPhu[i].index, lastRow, 1).getValues();
      if (listLuong) {
        for (let l = 0; l < listLuong.length; l++) {
          if (listLuong[l]) {
            listThoPhu[i].tongLuong += Number(listLuong[l]);
          }
        }
      }
    }
    var listDoanhThu = sheet.getRange(5, 17, lastRow, 1).getValues();
    if (listDoanhThu) {
      for (let l = 0; l < listDoanhThu.length; l++) {
        if (listDoanhThu[l]) {
          tongDoanhThu += Number(listDoanhThu[l]);
        }
      }
    }
    // end tinh tong luong
    // ô Tổng
    sheet.getRange(lastRow + 1, 15, 2, 2).merge()
      .setValue("Tổng").setFontSize("14")
      .setFontWeight("bold");

    // ô Tổng lương số
    sheet.getRange(lastRow + 1, 17, 2, 1).merge()
      .setValue(tongDoanhThu).setFontSize("14")
      .setFontWeight("bold")
      .setNumberFormat("#,##0");
    // ô lương chữ
    sheet.getRange(lastRow + 1, 18, 2, 1).merge()
      .setValue("Lương").setFontSize("14")
      .setFontWeight("bold");

    // in tổng lương thợ chính
    for (let i = 0; i < listThoChinh.length; i++) {
      sheet.getRange(lastRow + 1, listThoChinh[i].index, 2, 1).merge()
        .setValue(listThoChinh[i].tongLuong).setFontSize("14")
        .setFontWeight("bold")
        .setBackground(listThoChinh[i].color)
        .setNumberFormat("#,##0");
    }
    // in tổng lương thợ phụ
    for (let i = 0; i < listThoPhu.length; i++) {
      sheet.getRange(lastRow + 1, listThoPhu[i].index, 2, 1).merge()
        .setValue(listThoPhu[i].tongLuong).setFontSize("14")
        .setFontWeight("bold")
        .setBackground(listThoPhu[i].color)
        .setNumberFormat("#,##0");
    }

    // ô trống cuối
    sheet.getRange(lastRow + 1, +startColTho + listThoPhu.length + listThoChinh.length, 2, 1).merge()
      .setFontWeight("bold");

    // start merge cột ngày giống nhau
    var startRow = 5; // Dòng bắt đầu từ A5

    // Lấy tất cả giá trị trong cột O từ dòng 5 trở đi
    var data = sheet.getRange("O" + dongBatDauLuong + ":O" + (lastRow)).getValues();

    var startMergeRow = startRow;  // Dòng bắt đầu merge
    let curentDateCheck = data[0][0]
    let coutSame = 1;
    let currentColor = 1;
    let hangColor = "";
    Logger.log('So data quét' + data.length)
    for (var i = 1; i <= data.length; i++) {
      let curentDateCheckStr = curentDateCheck ? Utilities.formatDate(curentDateCheck, Session.getScriptTimeZone(), "dd/MM/yyyy") : "DONE";
      let dataStr = data[i] && data[i][0] ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "dd/MM/yyyy") : "DONE";

      if (curentDateCheckStr == dataStr || dataStr == 'DONE') {
        coutSame += 1;
        if (i == data.length - 1 && coutSame != 1) {
          let valueOld = sheet.getRange(startMergeRow, 15, coutSame, 1).getValue();
          sheet.getRange(startMergeRow, 15, coutSame, 1).clearContent().merge().setValue(valueOld).setFontWeight("bold");
          var data2 = sheet.getRange(startMergeRow, 17, coutSame, 1).getValues(); 
          var sum2 = data2.reduce((acc, val) => acc + (val[0] || 0), 0);
          sheet.getRange(startMergeRow, 18, coutSame, 1).clearContent().merge().setValue(sum2).setNumberFormat("#,##0");
          if (currentColor == 1) {
            hangColor = "#d9d2e9";
            currentColor = 2;
          } else {
            hangColor = "#ffffff";
            currentColor = 1;
          }
          sheet.getRange(startMergeRow, 15, coutSame, 4).setBackground(hangColor);
          startMergeRow += coutSame;
        }
      } else {
        if (coutSame != 1) {
          let valueOld = sheet.getRange(startMergeRow, 15, coutSame, 1).getValue();
          sheet.getRange(startMergeRow, 15, coutSame, 1).clearContent().merge().setValue(valueOld).setFontWeight("bold");
          var data1 = sheet.getRange(startMergeRow, 17, coutSame, 1).getValues(); 
          var sum1 = data1.reduce((acc, val) => acc + (val[0] || 0), 0);
          sheet.getRange(startMergeRow, 18, coutSame, 1).clearContent().merge().setValue(sum1).setNumberFormat("#,##0");
          if (currentColor == 1) {
            hangColor = "#d9d2e9";
            currentColor = 2;
          } else {
            hangColor = "#ffffff";
            currentColor = 1;
          }
          sheet.getRange(startMergeRow, 15, coutSame, 4).setBackground(hangColor);
          startMergeRow += coutSame;
        } else {
          if (currentColor == 1) {
            hangColor = "#d9d2e9";
            currentColor = 2;
          } else {
            hangColor = "#ffffff";
            currentColor = 1;
          }
          sheet.getRange(startMergeRow, 15, coutSame, 4).setBackground(hangColor);
          startMergeRow += 1;
        }
        coutSame = 1;
        curentDateCheck = data[i][0];
       }
    }
        // Gộp tất cả ô lại và set viền cùng lúc
    sheet.getRange(dongBatDauLuong, cotBatDauLuong, lastRow - dongBatDauLuong + 3, 5 + listThoChinh.length + listThoPhu.length).setBorder(true, true, true, true, true, true).setHorizontalAlignment("center").setVerticalAlignment("middle"); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
    // end merge cột ngày giống nhau
  //  ui.alert("Tính lương đã xong.");

  }
}

function exportSheetToPdfAndEmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Chọn sheet hiện tại
  var sheetId = sheet.getSheetId(); // Lấy ID của sheet

  // Lấy ID của file Google Sheets
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  // Thiết lập các tham số cho PDF
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' +
            'format=pdf&' + // Định dạng PDF
            'size=A4&' + // Kích thước giấy
            'portrait=true&' + // In dọc
            'fitw=true&' + // Phù hợp với chiều rộng
            'sheetnames=false&' + // Không hiển thị tên sheet
            'printtitle=false&' + // Không hiển thị tiêu đề
            'pagenumbers=true&' + // Hiển thị số trang
            'gridlines=false&' + // Ẩn đường lưới
            'fzr=false&' + // Không đông cột
            'gid=' + sheetId; // ID của sheet

  // Tạo request cho file PDF
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  // Tạo tệp PDF từ dữ liệu trả về
  var pdfBlob = response.getBlob().setName(sheet.getName() + ".pdf");

  // Gửi email với file PDF đính kèm
  var email = "mrthanh260801@gmail.com"; // Thay bằng email người nhận
  var subject = "Báo cáo PDF từ Google Sheets";
  var body = "Đây là báo cáo của bạn dưới dạng file PDF.";
  MailApp.sendEmail(email, subject, body, {
    attachments: [pdfBlob]
  });

  Logger.log("Email đã được gửi với file PDF đính kèm.");
}

function createTable() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("T12");
  
  // Xóa dữ liệu cũ
  sheet.clear();

  // Dữ liệu cho bảng
  var data = [
    ["STT", "Họ và Tên", "Tuổi", "Địa Chỉ"], // Tiêu đề
    [1, "Nguyễn Văn A", 25, "Hà Nội"],
    [2, "Trần Thị B", 30, "TP. Hồ Chí Minh"],
    [3, "Lê Văn C", 28, "Đà Nẵng"]
  ];

  // Ghi dữ liệu vào bảng
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // Định dạng bảng
  var range = sheet.getRange(1, 1, data.length, data[0].length);
  range.setFontWeight("bold").setBackground("#f0f0f0"); // Định dạng hàng tiêu đề
}

function refreshZaloToken() {
  var refreshToken = "YbsBSeTmWZ2FL8jtX7Y0DujTZts7ReDsqJxwKerfWrNa7ULaYaQASyK2zbdASh0UdHRC8P0-kZhdIiq0g0F1AknAa0lW0-PDj6dX2iGNl26ORCXXrn2ZUCTPta2-5zW6n4-z7C5an06q0Qv4wbth1wqncHhn7v8HuKlaLgmowXRT5uLSYbBxLzOA_16PTROwoYRiUvfZbq3G3-5fYZ_jKDr0XrNH1EnVlt_gPTKHhrYHMlTWvt62VemvzLVpUub0g1VmSlDZzqkiB8nnq7J2GRWYY4VsNjPLkdUPKRKEqMFANAHYfZ7iUzPdfsBQJT5aoW6THyLxo4MiCgqFt5VhBDWVcHQqP-uHl1k11lTpy38z7JzpWsw5EG"; // Token cũ của bạn
  var appId = "1800615964614328091";
  var appSecret = "846bSSQK6WcOh14539CJ";

  var url = "https://oauth.zaloapp.com/v4/access_token";
  var payload = {
    app_id: appId,
    refresh_token: refreshToken,
    grant_type: "refresh_token"
  };

  var headers = {
    "secret_key": appSecret, // Đưa app_secret vào header
    "Content-Type": "application/x-www-form-urlencoded"
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": payload
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  if (json.access_token) {
      var accessToken = json.access_token; // Access Token của tài khoản cá nhân
      var message = "Hello từ Zalo API cá nhân! 😃";

      var url = "https://graph.zalo.me/v2.0/me/message";
      var headersMessage = {
        "access_token": accessToken,
        "Content-Type": "application/json"
      };

      var payloadMessage = JSON.stringify({
        "message": message
      });

      var options = {
        "method": "post",
        "headers": headersMessage,
        "payload": payloadMessage
      };

      var responseMessage = UrlFetchApp.fetch(url, options);
      Logger.log(responseMessage.getContentText());
    return json.access_token;
  } else {
    Logger.log("Lỗi làm mới Access Token: " + response.getContentText());
    return null;
  }
}

function sendZaloPersonalMessage() {
  var accessToken = "OzWP1N73hXjcqHa0Siki76BO7GfntSbwPBWR1XVAtXXchZ9R5BNK9rZu7cemfFO2M_yWNmkwc1mRtrXVKgsV0mhtRbLTh8iD4iHQQr-dWY1Ez6Ty0RpJE7gW72zksDSpEOSpCL-QsNT5_MqgAeoUJZ_6TrOachSfG_juQWMPhmnTxKr51hwL47oi85vjvjSs3QKXVNpDaW07eqL75-gK8L-BPsiRzOW50EevOnIbo0r5sZ4I48RpQKRt0GanaiTpAj0n7INDsGfVYJPL4Stq4soY6reVkj4vQD0gBYlD_dT2ZJ1X5FR0Vr2uEJz6ri1MBQLt8sF7k7yKdqauSjlmSWURG2WIyvunHlLcL26mqZnW_JnTGteYAu4lTTsf4W"; // Access Token của tài khoản cá nhân
  var message = "Hello từ Zalo API cá nhân! 😃";

  var url = "https://graph.zalo.me/v2.0/me/message";
  var headers = {
    "access_token": accessToken,
    "Content-Type": "application/json"
  };

  var payload = JSON.stringify({
    "message": message
  });

  var options = {
    "method": "post",
    "headers": headers,
    "payload": payload
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}

function sendZaloZNSMessage() {
  var accessToken = "zxFyGHb-k13okln46dQw5_Qp-4uPOCr1jixhU3fOrr6dyf9h2pYfHENrYqTMQwnApDokMbuKwsFklTjF0dBYJEYyn44ESVLkgjRNVJXEytos-_ji7MNROlhg-J1-KhaAw_M585ybeoVljEDp22AZNe6p_4yj9hPSmv_tKdfMpmRExErNA6NOJgNFcKCq1ALwqfYzRdKf_3dAdyTkSrk4EPtli6i7T_18exon6oaHba-_YuDVAXl6QB-Yv6Sl7DfTWzISN15GgMg0w9uwFr-yRwhJga1CHfTIyVFNTMe_o2FTbO8eH56W7EtYdXrI9fHFueIsUtG-a6ByYensSJpAHfgrocq64Uq-uAJe56SlsKtJXUz3EZlKMednpGenE_GweAhq4ZjYj0Er_P4H7L3Q6ghrmWeICUWYPJdU_N9W7M2_6G"; // Token của Zalo OA
  var phoneNumber = "84901234567"; // Số điện thoại của người nhận (bắt đầu bằng 84 thay vì 0)
  var templateId = "YOUR_TEMPLATE_ID"; // ID của mẫu tin nhắn đã đăng ký với Zalo
  var messageParams = { "name": "Nguyễn Văn A", "order_id": "12345" }; // Dữ liệu thay thế trong template

  var url = "https://business.openapi.zalo.me/message/template";
  
  var headers = {
    "access_token": accessToken,
    "Content-Type": "application/json"
  };

  var payload = JSON.stringify({
    "phone": phoneNumber,
    "template_id": templateId,
    "template_data": messageParams
  });

  var options = {
    "method": "post",
    "headers": headers,
    "payload": payload
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}



