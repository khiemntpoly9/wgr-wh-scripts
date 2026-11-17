/**
 * ===================================================================
 *  Auto_Format_Print – Tự động định dạng toàn bộ sheet và tạo sheet mới để in.
 * ===================================================================
 *
 * Script này tự động định dạng toàn bộ sheet và tạo sheet mới để in dựa trên cấu hình.
 * Mọi thông số đều nằm trong CONFIG → dễ tùy chỉnh, mở rộng.
 *
 */
const FORMAT_PRINT_CONFIG = {
  // Cấu hình font chữ chung cho toàn bộ trang
  globalFont: {
    name: "Times New Roman",
    size: 14,
  },

  // Các cột cần ẩn trên sheet mới (VD: ["I:J"] hoặc ["I:I", "J:K"])
  columnsToHide: ["J:J", "K:K", "S:S"],

  // Ô cố định cần định dạng riêng trên sheet in. Tìm theo Name trước (Name Manager,
  // Scope = sheet gốc — Name được copy theo sang sheet in), chưa có Name thì dùng fallback.
  fixedCellStyles: [
    // Số phiếu
    { name: "pxk_ticket_no", fallback: "D4", bold: true, fontSize: 22 },
  ] as {
    name: string;
    fallback: string;
    bold?: boolean;
    fontSize?: number;
  }[],

  // Cấu hình xử lý cho bảng dữ liệu chính
  dataTable: {
    // Các từ khóa dùng để nhận diện ô tiêu đề cột STT trong cột A.
    // Ô phải chứa TẤT CẢ các từ khóa này (hỗ trợ cả trường hợp xuống dòng Alt+Enter trong ô).
    headerKeywords: ["STT", "序号"],
    startValue: 1, // Giá trị kiểm tra bắt buộc tại ô bắt đầu (ngay dưới headerKeyword)
    highlightColumns: ["C", "F", "H", "I", "P"], // Các cột cần tăng kích thước và in đậm
    highlightFontSize: 22, // Kích cỡ chữ mặc định cho phần highlight
    highlightColumnOverrides: {
      P: { fontSize: 18 }, // Cột P dùng kích cỡ riêng
    } as Record<string, { fontSize: number }>,
    rowHeight: 45, // Chiều cao cho các hàng dữ liệu
    columnsToAutofitAndWrap: ["C", "P"], // Các cột cần autofit và wrap text để chống lỗi #####
  },
};

/**
 * Name phạm vi sheet → Name phạm vi workbook (chỉ khi trỏ vào chính sheet này)
 * → địa chỉ fallback. Ghi tên Name không tìm thấy vào missing.
 */
const formatPrintResolveRange = (
  workbook: ExcelScript.Workbook,
  sheet: ExcelScript.Worksheet,
  ref: { name: string; fallback: string },
  missing: string[],
): ExcelScript.Range => {
  const sheetItem = ref.name ? sheet.getNamedItem(ref.name) : undefined;
  if (sheetItem) return sheetItem.getRange();
  const bookItem = ref.name ? workbook.getNamedItem(ref.name) : undefined;
  if (bookItem) {
    const range = bookItem.getRange();
    if (range.getWorksheet().getName() === sheet.getName()) return range;
  }
  if (ref.name) missing.push(`${ref.name} (→ ${ref.fallback})`);
  return sheet.getRange(ref.fallback);
};

// @ts-ignore: Office Scripts — mỗi file chạy độc lập, không trùng main()
async function main(workbook: ExcelScript.Workbook) {
  const CONFIG = FORMAT_PRINT_CONFIG;

  // Lấy sheet hiện tại
  let currentSheet = workbook.getActiveWorksheet();

  // 0. Kiểm tra và bỏ ẩn cột H nếu đang bị ẩn
  let colH = currentSheet.getRange("H:H");
  colH.setColumnHidden(false);

  // 1. Định dạng toàn bộ sheet thành cấu hình font chung
  let entireRange = currentSheet.getRange();
  let font = entireRange.getFormat().getFont();
  font.setName(CONFIG.globalFont.name);
  font.setSize(CONFIG.globalFont.size);

  // 2. Copy sheet hiện tại thành một sheet mới (đặt ngay sau sheet hiện hành)
  let newSheet = currentSheet.copy(
    ExcelScript.WorksheetPositionType.after,
    currentSheet,
  );

  // 3. Tiến hành ẩn các cột đã được cấu hình
  CONFIG.columnsToHide.forEach((col) => {
    newSheet.getRange(col).setColumnHidden(true);
  });

  // 3.1 Định dạng ô cố định (VD: D4 in đậm, cỡ 22)
  const missingNames: string[] = [];
  (CONFIG.fixedCellStyles || []).forEach((style) => {
    const cellFont = formatPrintResolveRange(
      workbook,
      newSheet,
      style,
      missingNames,
    )
      .getFormat()
      .getFont();
    if (style.fontSize !== undefined) cellFont.setSize(style.fontSize);
    if (style.bold !== undefined) cellFont.setBold(style.bold);
  });
  if (missingNames.length) {
    console.log(
      `Chưa có Name, đang dùng địa chỉ mặc định: ${missingNames.join(", ")}.`,
    );
  }

  // 4. Tìm dòng tiêu đề chứa "STT" + "序号" trong cột A để xác định vùng dữ liệu
  //    Duyệt qua vùng dữ liệu đã dùng (usedRange) để tìm ô chứa tất cả headerKeywords.
  //    Cách này xử lý được cả trường hợp ô chứa xuống dòng (Alt+Enter).
  let usedRange = newSheet.getUsedRange();
  let headerRowIndex = -1; // 0-based

  if (usedRange) {
    let lastUsedRow = usedRange.getRowCount();
    let colARange = newSheet.getRange(`A1:A${lastUsedRow}`);
    let colAValues = colARange.getValues();

    for (let i = 0; i < colAValues.length; i++) {
      let cellValue = String(colAValues[i][0]).toUpperCase();
      let allMatch = CONFIG.dataTable.headerKeywords.every((kw) =>
        cellValue.includes(kw.toUpperCase()),
      );
      if (allMatch) {
        headerRowIndex = i; // 0-based index trong mảng = 0-based row index (vì bắt đầu từ A1)
        break;
      }
    }
  }

  if (headerRowIndex === -1) {
    console.log(
      `Không tìm thấy ô chứa [${CONFIG.dataTable.headerKeywords.join(", ")}] trong cột A. Bỏ qua bước định dạng bảng dữ liệu.`,
    );
    return;
  }

  // Ô bắt đầu dữ liệu = ô ngay bên dưới dòng tiêu đề
  let headerRowNum = headerRowIndex + 1; // 1-based (số dòng trên Excel)
  let startRow = headerRowNum + 1;

  // Xác định dòng cuối dựa trên cột D (Tên hàng) thay vì cột A (STT)
  // vì cột STT có thể thiếu dữ liệu ở một số dòng cuối.
  let dataStartD = newSheet.getRange(`D${startRow}`);
  if (!dataStartD.getValue()) {
    console.log(`❌ Ô D${startRow} trống. Không có dữ liệu để xử lý.`);
    return;
  }
  let dataRangeD = dataStartD.getExtendedRange(
    ExcelScript.KeyboardDirection.down,
  );
  let lastRow = startRow - 1 + dataRangeD.getRowCount();

  {
    // 4.1 Điền công thức STT: =ROW()-headerRow cho tất cả hàng dữ liệu
    for (let row = startRow; row <= lastRow; row++) {
      newSheet.getRange(`A${row}`).setFormula(`=ROW()-${headerRowNum}`);
    }

    // 4.2 Kẻ khung (All Borders) cho toàn bộ bảng (tiêu đề → hàng cuối)
    let fullTableRange = newSheet.getRange(`A${headerRowNum}:Q${lastRow}`);
    const borderIndices = [
      ExcelScript.BorderIndex.edgeTop,
      ExcelScript.BorderIndex.edgeBottom,
      ExcelScript.BorderIndex.edgeLeft,
      ExcelScript.BorderIndex.edgeRight,
      ExcelScript.BorderIndex.insideVertical,
      ExcelScript.BorderIndex.insideHorizontal,
    ];
    borderIndices.forEach((bIndex) => {
      let border = fullTableRange.getFormat().getRangeBorder(bIndex);
      border.setStyle(ExcelScript.BorderLineStyle.continuous);
      border.setWeight(ExcelScript.BorderWeight.thin);
      border.setColor("#000000");
    });

    // 4.2 Áp dụng highlight (kích cỡ, in đậm) cho các cột được cấu hình (VD: Cột C, H, I, P)
    CONFIG.dataTable.highlightColumns.forEach((col) => {
      let highlightRange = newSheet.getRange(
        `${col}${startRow}:${col}${lastRow}`,
      );
      let hFont = highlightRange.getFormat().getFont();
      let overrides = CONFIG.dataTable.highlightColumnOverrides;
      let fontSize =
        overrides && overrides[col]
          ? overrides[col].fontSize
          : CONFIG.dataTable.highlightFontSize;
      hFont.setSize(fontSize);
      hFont.setBold(true);
    });

    // 4.3 Áp dụng autofit và wrap text cho các cột dài (VD: Cột C)
    CONFIG.dataTable.columnsToAutofitAndWrap.forEach((col) => {
      let wrapRange = newSheet.getRange(`${col}${startRow}:${col}${lastRow}`);
      let wrapFormat = wrapRange.getFormat();
      wrapFormat.autofitColumns();
      wrapFormat.setWrapText(true);
    });

    // 4.4 Định dạng toàn bộ hàng từ dòng bắt đầu đến dòng kết thúc: Căn giữa, chỉnh chiều cao
    let rowsRange = newSheet.getRange(`${startRow}:${lastRow}`);
    let rowsFormat = rowsRange.getFormat();

    rowsFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    rowsFormat.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    rowsFormat.setRowHeight(CONFIG.dataTable.rowHeight);

    // 4.5 Chèn công thức cột Q: Số lượng còn lại = F (Xuất đi) - H (Đã xuất) - I (Đóng gói)
    let colQRange = newSheet.getRange(`Q${startRow}:Q${lastRow}`);
    // Điền công thức theo dải ô (nhanh hơn loop setValue)
    colQRange.setFormula(`=F${startRow}-H${startRow}-I${startRow}`);

    // Định dạng cột Q: Font 20, Bold
    let qFont = colQRange.getFormat().getFont();
    qFont.setSize(20);
    qFont.setBold(true);

    // 4.6 Thêm dòng tổng cộng (Summary Row) cho cột F và I
    let summaryRow = lastRow + 2;
    ["F", "I"].forEach((col) => {
      let cell = newSheet.getRange(`${col}${summaryRow}`);

      // Công thức tính tổng
      cell.setFormula(`=SUM(${col}${startRow}:${col}${lastRow})`);

      // Định dạng: Font 22, Bold, Căn giữa
      let format = cell.getFormat();
      let font = format.getFont();
      font.setSize(22);
      font.setBold(true);
      format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

      // Nền vàng (tương tự PXK)
      format.getFill().setColor("#FFFF00");

      // Kẻ khung cho ô tổng cộng
      const summaryBorders = [
        ExcelScript.BorderIndex.edgeTop,
        ExcelScript.BorderIndex.edgeBottom,
        ExcelScript.BorderIndex.edgeLeft,
        ExcelScript.BorderIndex.edgeRight,
      ];
      summaryBorders.forEach((bIndex) => {
        let border = format.getRangeBorder(bIndex);
        border.setStyle(ExcelScript.BorderLineStyle.continuous);
        border.setWeight(ExcelScript.BorderWeight.thin);
        border.setColor("#000000");
      });
    });
  }
}
