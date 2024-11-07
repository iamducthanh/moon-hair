function createOrUpdateSheetLuong() {
  // sheet cấu hình
  const cauHinh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cấu hình");
  const tenThang = cauHinh.getRange("C1").getValue();

  const sheetDoanhThu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Doanh Thu T" + tenThang);
  const sheetLuong = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lương T" + tenThang);
  sheetLuong.getDataRange().setVerticalAlignment("middle");

  sheetLuong.setColumnWidth(1, 120)
    .setColumnWidth(2, 150)
    .setColumnWidth(3, 130)
    .setColumnWidth(4, 130)
    .setColumnWidth(5, 150)
    .setColumnWidth(6, 150)
    .setColumnWidth(7, 150)
    .setColumnWidth(8, 150)
    .setColumnWidth(9, 150)
    .setColumnWidth(10, 150)
    .setColumnWidth(11, 150);
  const cellTitle = sheetLuong.getRange("A1");
  let title = "Lương tháng " + tenThang;
  cellTitle.setValue(title);
  cellTitle.setFontWeight("bold");
  cellTitle.setFontSize(15);
  var values = [["Ngày", "Tên khách", "Tiền bill", "Phương thức"]];

  const headerCommons = sheetLuong.getRange(3, 1, 2, 4);
  for (var col = 1; col <= 4; col++) {
    // Gộp ô tại hàng 3 và hàng 4 cho mỗi cột
    sheetLuong.getRange(3, col, 2, 1).merge(); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
    sheetLuong.getRange(3, col).setValue(values[0][col - 1]);
    sheetLuong.getRange(3, col).setVerticalAlignment("middle");
  }
  headerCommons.setFontWeight("bold");
  headerCommons.setFontSize(12);
  headerCommons.setBackground("#d3d3d3");
  headerCommons.setBorder(true, true, true, true, true, true); // Đặt đường viền cho các cạnh trên, dưới, trái, phải

  let lastRowThoChinh = sheetDoanhThu.getRange(sheetDoanhThu.getMaxRows(), 10).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  let lastRowThoPhu = sheetDoanhThu.getRange(sheetDoanhThu.getMaxRows(), 12).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  let colThoChinh = sheetDoanhThu.getRange("J4:J" + lastRowThoChinh);
  let thoChinh = colThoChinh.getValues();

  let colThoPhu = sheetDoanhThu.getRange("L4:L" + lastRowThoPhu);
  let thoPhu = colThoPhu.getValues();

  var startCellTho = sheetLuong.getRange("E4"); // Lấy ô E4
  let coutTho = 0;

  for (let i = 0; i < thoChinh.length; i++) {
    startCellTho.offset(0, coutTho).setValue(thoChinh[i])
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground("#d3d3d3")
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang
    coutTho += 1;
  }
  sheetLuong.getRange(3, 5, 1, thoChinh.length).merge(); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
  sheetLuong.getRange(3, 5).setValue("Thợ chính")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground("#d3d3d3")
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang
  for (let i = 0; i < thoPhu.length; i++) {
    startCellTho.offset(0, coutTho).setValue(thoPhu[i])
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground("#d3d3d3")
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang
    coutTho += 1;

  }
  sheetLuong.getRange(3, 5 + thoChinh.length, 1, thoPhu.length).merge(); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
  sheetLuong.getRange(3, 5 + thoChinh.length).setValue("Thợ phụ")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground("#d3d3d3")
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang

  sheetLuong.getRange(3, 5 + thoChinh.length + thoPhu.length, 2, 1).merge()
    .setValue("Ngày sửa")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground("#d3d3d3")
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment("right");

}

function tinhLuong() {
  let dongBatDauLuong = 5;
  let dongBatDauDT = 4;
  let startColTho = 5;
  let tongDoanhThu = 0;

  // sheet cấu hình
  const cauHinh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cấu hình");
  const tenThang = cauHinh.getRange("C1").getValue();

  const sheetDoanhThu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Doanh Thu T" + tenThang);
  const sheetLuong = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lương T" + tenThang);

  var lastRow = sheetDoanhThu.getLastRow();
  // Doanh thu - Ngày
  const columnDateDT = sheetDoanhThu.getRange("A" + dongBatDauDT + ":A" + lastRow);
  const dateDT = columnDateDT.getValues();
  // Doanh thu - tên khách
  const columnCustomerDT = sheetDoanhThu.getRange("B" + dongBatDauDT + ":B" + lastRow);
  const customerDT = columnCustomerDT.getValues();
  // Doanh thu - tiền bill
  const columnBillDT = sheetDoanhThu.getRange("C" + dongBatDauDT + ":C" + lastRow);
  const billDT = columnBillDT.getValues();
  // Doanh thu - phương thức
  const columnPhuongThucDT = sheetDoanhThu.getRange("D" + dongBatDauDT + ":D" + lastRow);
  const phuongThucDT = columnPhuongThucDT.getValues();
  // Doanh thu - thợ chính
  const columnThoChinhDT = sheetDoanhThu.getRange("E" + dongBatDauDT + ":E" + lastRow);
  const thoChinhDT = columnThoChinhDT.getValues();

  // Doanh thu - thợ phụ
  const columnThoPhuDT = sheetDoanhThu.getRange("F" + dongBatDauDT + ":F" + lastRow);
  const thoPhuDT = columnThoPhuDT.getValues();

  // Lương - ngày
  const columnDateL = sheetLuong.getRange("A" + dongBatDauLuong + ":A" + lastRow + 1);
  const dateL = columnDateL.getValues();
  // Lương - tên khách
  const columnCustomerL = sheetLuong.getRange("B" + dongBatDauLuong + ":B" + lastRow + 1);
  const customerL = columnCustomerL.getValues();
  // Lương - tiền bill
  const columnBillL = sheetLuong.getRange("C" + dongBatDauLuong + ":C" + lastRow + 1);
  columnBillL.setNumberFormat("#,##0 đ");
  const billL = columnBillL.getValues();
  // Lương - phương thức
  const columnPhuongThucL = sheetLuong.getRange("D" + dongBatDauLuong + ":D" + lastRow + 1);
  const phuongThucL = columnPhuongThucL.getValues();


  // Danh sách thợ
  let listThoChinh = new Array();
  let listThoPhu = new Array();

  let lastRowThoChinh = sheetDoanhThu.getRange(sheetDoanhThu.getMaxRows(), 10).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  let lastRowThoPhu = sheetDoanhThu.getRange(sheetDoanhThu.getMaxRows(), 12).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  let colThoChinh = sheetDoanhThu.getRange("J4:J" + lastRowThoChinh);
  let thoChinh = colThoChinh.getValues();
  let colLuongThoChinh = sheetDoanhThu.getRange("K4:K" + lastRowThoChinh);
  let luongThoChinh = colLuongThoChinh.getValues();

  let colThoPhu = sheetDoanhThu.getRange("L4:L" + lastRowThoPhu);
  let thoPhu = colThoPhu.getValues();
  let colLuongThoPhu = sheetDoanhThu.getRange("M4:M" + lastRowThoPhu);
  let luongThoPhu = colLuongThoPhu.getValues();

  var headerThoChinh = sheetLuong.getRange(4, 5, 1, thoChinh.length).getValues()[0];
  var headerThoPhu = sheetLuong.getRange(4, 5 + thoChinh.length, 1, thoPhu.length).getValues()[0];

  for (let i = 0; i < thoChinh.length; i++) {
    listThoChinh.push({
      name: thoChinh[i][0],
      luong: luongThoChinh[i][0],
      index: headerThoChinh.indexOf(thoChinh[i][0]) + startColTho,
      color: colThoChinh.getCell(i + 1, 1).getBackground(),
      tongLuong: 0
    })
  }

  for (let i = 0; i < thoPhu.length; i++) {
    listThoPhu.push({
      name: thoPhu[i][0],
      luong: luongThoPhu[i][0],
      index: headerThoPhu.indexOf(thoPhu[i][0]) + startColTho + thoChinh.length,
      color: colThoPhu.getCell(i + 1, 1).getBackground(),
      tongLuong: 0
    })
  }

  for (let i = 0; i < dateDT.length; i++) {
    if (dateDT[i] != undefined && dateDT[i] != "") {
      // reset value dòng
      sheetLuong.getRange(dongBatDauLuong + i, 1, 1, 50).setValue("");

      dateL[i] = dateDT[i];
      columnDateL.getCell(i + 1, 1).setBorder(true, true, true, true, true, true).setHorizontalAlignment("left");
      customerL[i] = customerDT[i];
      columnCustomerL.getCell(i + 1, 1).setBorder(true, true, true, true, true, true);
      billL[i] = billDT[i];
      tongDoanhThu += Number(billDT[i]);
      columnBillL.getCell(i + 1, 1).setBorder(true, true, true, true, true, true);
      phuongThucL[i] = phuongThucDT[i];
      let colorPT = "";
      if (phuongThucDT[i][0].toLowerCase() == "chuyển khoản") {
        colorPT = "#d4edbc";
      } else {
        colorPT = "#bfe1f6";
      }
      columnPhuongThucL.getCell(i + 1, 1).setBorder(true, true, true, true, true, true)
        .setHorizontalAlignment("right")
        .setBackground(colorPT);
      const curentThoChinh = listThoChinh.find(item => item.name == thoChinhDT[i]);
      if (curentThoChinh) {
        let luong = billDT[i] / 100 * curentThoChinh.luong;
        curentThoChinh.tongLuong += luong;
        sheetLuong.getRange(dongBatDauLuong + i, curentThoChinh.index)
          .setValue(luong)
          .setBorder(true, true, true, true, true, true)
          .setHorizontalAlignment("right")
          .setFontSize("12")
          .setBackground(curentThoChinh.color)
          .setNumberFormat("#,##0 đ");
      }

      const curentThoPhu = listThoPhu.find(item => item.name == thoPhuDT[i]);
      if (curentThoPhu) {
        if (customerDT[i][0].toLowerCase() == "bsp" || customerDT[i][0].toLowerCase() == "gội") {
          let luong = billDT[i] / 100 * 20;
          curentThoPhu.tongLuong += luong;
          sheetLuong.getRange(dongBatDauLuong + i, curentThoPhu.index)
            .setValue(luong)
            .setHorizontalAlignment("right")
            .setFontSize("12")
            .setBackground(curentThoPhu.color)
            .setNumberFormat("#,##0 đ");
        } else {
          let luong = billDT[i] / 100 * curentThoPhu.luong;
          curentThoPhu.tongLuong += luong;
          sheetLuong.getRange(dongBatDauLuong + i, curentThoPhu.index)
            .setValue(luong)
            .setHorizontalAlignment("right")
            .setFontSize("12")
            .setBackground(curentThoPhu.color)
            .setNumberFormat("#,##0 đ");
        }
      }
      var currentDate = new Date(); // Lấy ngày hiện tại
      sheetLuong.getRange(dongBatDauLuong + i, startColTho + listThoPhu.length + listThoChinh.length).setValue(Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "HH:MM dd/MM/yyyy")).setFontSize("12");
    }
  }

  columnDateL.setValues(dateL).setFontSize("12");
  columnCustomerL.setValues(customerL).setFontSize("12");
  columnBillL.setValues(billL).setFontSize("12");
  columnPhuongThucL.setValues(phuongThucL).setFontSize("12");

  sheetLuong.getRange(dongBatDauLuong, startColTho, dateDT.length, listThoPhu.length + listThoChinh.length + 1).setBorder(true, true, true, true, true, true).setHorizontalAlignment("right");

  sheetLuong.getRange(lastRow + 2, 1, 2, 2).merge()
    .setBorder(true, true, true, true, true, true)
    .setValue("Tổng").setFontSize("14")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheetLuong.getRange(lastRow + 2, 3, 2, 1).merge()
    .setBorder(true, true, true, true, true, true)
    .setValue(tongDoanhThu).setFontSize("14")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setNumberFormat("#,##0 đ");

  sheetLuong.getRange(lastRow + 2, 4, 2, 1).merge()
    .setBorder(true, true, true, true, true, true)
    .setValue("Lương").setFontSize("14")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");


  for (let i = 0; i < listThoChinh.length; i++) {
    sheetLuong.getRange(lastRow + 2, listThoChinh[i].index, 2, 1).merge()
      .setBorder(true, true, true, true, true, true)
      .setValue(listThoChinh[i].tongLuong).setFontSize("14")
      .setFontWeight("bold")
      .setHorizontalAlignment("right")
      .setVerticalAlignment("middle")
      .setBackground(listThoChinh[i].color)
      .setNumberFormat("#,##0 đ");
  }
    for (let i = 0; i < listThoPhu.length; i++) {
    sheetLuong.getRange(lastRow + 2, listThoPhu[i].index, 2, 1).merge()
      .setBorder(true, true, true, true, true, true)
      .setValue(listThoPhu[i].tongLuong).setFontSize("14")
      .setFontWeight("bold")
      .setHorizontalAlignment("right")
      .setVerticalAlignment("middle")
      .setBackground(listThoPhu[i].color)
      .setNumberFormat("#,##0 đ");
  }

  Logger.log(tongDoanhThu)


  // var startRow = 5; // Dòng bắt đầu từ A5
  // var column = 1; // Cột A

  // // Lấy tất cả giá trị trong cột A từ dòng 5 trở đi
  // var data = sheetLuong.getRange("A" + dongBatDauLuong + ":A" + lastRow).getValues();
  // var startMergeRow = startRow;  // Dòng bắt đầu merge
  // let curentDateCheck = data[0][0]
  // let coutSame = 1;
  // for (var i = 1; i < data.length; i++) {
  //   Logger.log(curentDateCheck)
  //   if(data[i][0] == curentDateCheck) {
  //     coutSame += 1;
  //   } else {
  //     if (coutSame != 1) {
  //       sheetLuong.getRange(startMergeRow, 1, coutSame, 1).merge(); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
  //       startMergeRow += coutSame;
  //     } else {
  //       startMergeRow += 1;
  //     }
  //     coutSame = 1;
  //     curentDateCheck = data[i][0];
  //   }
  // }

}
