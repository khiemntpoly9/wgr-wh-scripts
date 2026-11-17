/**
 * ===================================================================
 *  Auto_PXK – Tự động định dạng Phiếu Xuất Kho
 * ===================================================================
 *
 * Script này tự động format sheet Phiếu Xuất Kho dựa trên cấu hình.
 * Mọi thông số đều nằm trong CONFIG → dễ tùy chỉnh, mở rộng.
 *
 */
// ─── CẤU HÌNH (CONFIG) ─────────────────────────────────────────────
// Chỉ cần chỉnh CONFIG, không cần sửa logic bên dưới.
// ────────────────────────────────────────────────────────────────────
const PXK_CONFIG = {
  // ── Font chung cho toàn bộ sheet ──
  globalFont: {
    name: "Times New Roman",
    size: 14,
  },

  // ── Cấu hình nhận diện bảng dữ liệu ──
  dataTable: {
    /**
     * Từ khóa nhận diện hàng tiêu đề trong cột A.
     * Ô phải chứa TẤT CẢ các từ khóa (hỗ trợ xuống dòng Alt+Enter).
     */
    headerKeywords: ["STT", "序号"],

    /**
     * Danh sách cột dữ liệu và thuộc tính.
     * - `letter`: Tên cột Excel (A, B, C, ...)
     * - `header`: Tiêu đề cột (dùng cho tham khảo / mở rộng sau này)
     * - `align`: Căn ngang cho vùng dữ liệu ("center" | "left" | "right")
     */
    columns: [
      { letter: "A", header: "STT 序号", align: "center" as const },
      {
        letter: "B",
        header: "SỐ HỢP ĐỒNG 合同编号",
        align: "center" as const,
      },
      { letter: "C", header: "MÃ HÀNG 货号", align: "center" as const },
      { letter: "D", header: "TÊN HÀNG 产品名称", align: "left" as const },
      {
        letter: "E",
        header: "KÍCH THƯỚC (cm) 尺寸",
        align: "center" as const,
      },
      {
        letter: "F",
        header: "SỐ LƯỢNG XUẤT ĐI 出货数量",
        align: "center" as const,
      },
      { letter: "G", header: "ĐVT 单位", align: "center" as const },
      {
        letter: "H",
        header: "ĐÃ XUẤT HÀNG 已发货数量",
        align: "center" as const,
      },
      {
        letter: "I",
        header: "SỐ LƯỢNG ĐÓNG GÓI 包装数量",
        align: "center" as const,
      },
      { letter: "J", header: "ĐVT 单位", align: "center" as const },
      { letter: "K", header: "SỐ LÔ 批次号", align: "center" as const },
      { letter: "L", header: "TRỌNG LƯỢNG", align: "center" as const },
      {
        letter: "M",
        header: "TỔNG TRỌNG LƯỢNG (G.W./KG)总毛重",
        align: "center" as const,
      },
      { letter: "N", header: "THỂ TÍCH (M³) 体积", align: "center" as const },
      {
        letter: "O",
        header: "TỔNG THỂ TÍCH (M³) 总体积",
        align: "center" as const,
      },
      { letter: "P", header: "GHI CHÚ 备注", align: "left" as const },
      { letter: "Q", header: "", align: "left" as const }, // Cột Q: gộp với P ở header, data để trống
      // Cột Approve PMC y/n:
      { letter: "R", header: "Approve PMC y/n", align: "center" as const },
    ],

    /**
     * Các vùng cần merge ở hàng tiêu đề.
     * Mỗi phần tử: { from: "O", to: "P" } → merge O:P trên hàng header.
     * Dễ mở rộng: thêm phần tử mới nếu cần merge thêm cột.
     */
    headerMergeRegions: [{ from: "P", to: "Q" }],

    /**
     * Các cột cần autofit chiều rộng.
     * Thêm tên cột vào đây nếu muốn autofit thêm cột khác.
     */
    autofitColumns: ["D"],

    /**
     * Các cột KHÔNG được wrap text (tắt xuống dòng).
     * Mặc định Office Scripts wrap text = false, nhưng khai báo rõ
     * để xử lý trường hợp sheet gốc đã bật wrap.
     */
    noWrapColumns: ["D"],

    /**
     * Các cột cần ẩn.
     * Hỗ trợ cả range ("Q:R") và đơn lẻ ("Q", "R").
     * Để trống [] nếu không cần ẩn cột nào.
     */
    columnsToHide: ["H"],

    /**
     * Cấu hình hàng tiêu đề.
     */
    headerStyle: {
      bold: true,
      horizontalAlignment: "center" as const,
      verticalAlignment: "middle" as const,
    },

    /**
     * Vertical alignment mặc định cho tất cả hàng dữ liệu.
     */
    dataVerticalAlignment: "middle" as const,

    /**
     * Kẻ khung (All Borders) cho toàn bộ bảng (từ tiêu đề đến hết dữ liệu).
     */
    applyAllBorders: true,

    /**
     * Các cột cần tính tổng (SUM) ở cuối bảng (cách dòng dữ liệu cuối 1 hàng).
     */
    summaryColumns: ["F", "I", "M", "O"],

    /**
     * Công thức tính toán cho từng dòng dữ liệu.
     * - `target`: Cột chứa kết quả.
     * - `formula`: Chuỗi công thức, dùng {row} làm placeholder cho số dòng.
     *
     * Ví dụ: L = K * H  →  "=K{row}*H{row}"
     */
    formulaColumns: [
      { target: "M", formula: "=ROUND(L{row}*I{row}, 2)" }, // Tổng trọng lượng = Trọng lượng * Số lượng đóng gói (Làm tròn 2 chữ số)
      { target: "O", formula: "=ROUND(N{row}*I{row}, 2)" }, // Tổng thể tích = Thể tích * Số lượng đóng gói (Làm tròn 2 chữ số)
    ],

    /**
     * Cấu hình tự động tính thể tích từ cột kích thước nếu ô thể tích bị trống.
     * - sourceCol: Cột chứa kích thước (VD: "E" -> "42*24.5*25")
     * - targetCol: Cột điền thể tích (VD: "M")
     * - divisor: Hệ số chia (Mặc định 1000000)
     */
    autoCalculateVolume: {
      enabled: true,
      sourceCol: "E",
      targetCol: "N",
      divisor: 1000000,
    },
    /**
     * Cấu hình tự động điền số lô (Cột J) dựa trên mã hợp đồng (Cột B).
     * - contractCol: Cột chứa mã hợp đồng (VD: "B")
     * - batchCol: Cột chứa số lô (VD: "J")
     * Số lô hợp lệ là chuỗi gồm 10 chữ số. Các mã hợp đồng có chung tiền tố
     * (VD: EUF6001-1, EUF6001 3 -> chung tiền tố EUF6001) sẽ dùng chung số lô.
     */
    autoFillBatchNumber: {
      enabled: true,
      contractCol: "B",
      batchCol: "K",
    },

    /**
     * Cấu hình tự động tra cứu số lô từ sheet riêng.
     * - sheetName: Tên sheet chứa danh sách số lô (mặc định "lot-no")
     * - lookupContractCol: Cột mã hợp đồng trong sheet tra cứu (mặc định "A")
     * - lookupBatchCol: Cột số lô trong sheet tra cứu (mặc định "B")
     * - targetContractCol: Cột mã hợp đồng trên sheet hiện tại (mặc định "B")
     * - targetBatchCol: Cột số lô trên sheet hiện tại (mặc định "K")
     * Hàng đầu tiên của sheet tra cứu là tiêu đề, dữ liệu bắt đầu từ hàng 2.
     * Chỉ điền vào ô số lô đang trống.
     */
    autoLookupBatchNumber: {
      enabled: true,
      sheetName: "lot-no",
      lookupContractCol: "A",
      lookupBatchCol: "B",
      targetContractCol: "B",
      targetBatchCol: "K",
    },

    /**
     * Điền ĐVT còn trống: cột G = "PCS"; cột J theo chiều cao (số thứ 3) ở cột E.
     * Nếu kích thước dạng a*b*1,1 (VD: 44*24*1,1) → J = "PACKAGE", ngược lại → "箱/ THÙNG CARTON".
     */
    autoFillUnitColumns: {
      enabled: true,
      qtyUnitCol: "G",
      qtyUnitDefault: "PCS",
      packUnitCol: "J",
      packUnitCarton: "箱/ THÙNG CARTON",
      packUnitPackage: "PACKAGE",
      dimCol: "E",
      packageHeight: 1.1,
    },

    /**
     * Bù trừ cân nặng chỉ chạy khi sheet hiện tại đúng tên này.
     * Sheet "pre-packing-list" chỉ format, không bù trừ.
     */
    saleSheetName: "sale",

    /**
     * Khi sheet "sale" trống dữ liệu, sao chép từ sheet nguồn (cùng cấu trúc cột).
     * Chỉ chạy trên sheet sale; sau khi sao chép tiếp tục format như bình thường.
     */
    copyFromSourceWhenEmpty: {
      enabled: true,
      sourceSheetName: "pre-packing-list",
      firstCol: "A",
      lastCol: "Q",
    },

    /**
     * Mục đích: Tự động điều chỉnh trọng lượng đơn (Cột K) sao cho tổng trọng lượng (Cột L)
     * khớp với con số đích được ghi ngay dưới ô tính tổng.
     */
    autoBalanceWeight: {
      enabled: true,
      productCol: "C", // Cột mã hàng để hiển thị log
      dimCol: "E", // Cột kích thước (gom nhóm)
      weightCol: "L",
      qtyCol: "I",
      totalWeightCol: "M",
      maxAdjustPerStep: 0.5,
    },

    /**
     * Điều chỉnh thể tích container: sửa kích thước (cột E) để tổng thể tích (cột O)
     * khớp số khối mong muốn. Chỉ chạy trên sheet "sale".
     * - enabled: bật / tắt tính năng.
     * - targetTotalVolume: số khối đích (m³). Để 0 → lấy giá trị ô ngay dưới dòng tổng cột O.
     * - adjustSide: cạnh được sửa — "all" (co giãn đều 3 cạnh) | "length" | "width" | "height".
     * - dimensionDecimals: 0 → mỗi cạnh chỉ tăng hoặc giảm theo số nguyên (cm). Không dùng bước thập phân.
     * - maxChangePercent: thể tích mỗi kích thước không được lệch quá % này so với ban đầu.
     * - wholeNumberTarget: true → chỉ cần khớp phần nguyên (đích 66 → tổng 66.00–66.99 đều đạt).
     * - tolerance: sai số chấp nhận (m³), chỉ dùng khi wholeNumberTarget = false.
     * - packageHeight: kích thước có chiều cao (số thứ 3) bằng giá trị này là hàng PACKAGE
     *   (VD: 2*6*1,1) → giữ nguyên kích thước, không tăng / giảm.
     * Khi bật, cột N của các dòng có kích thước hợp lệ (trừ PACKAGE) được tính lại từ cột E.
     */
    autoAdjustContainerVolume: {
      enabled: false,
      targetTotalVolume: 0,
      productCol: "C",
      dimCol: "E",
      qtyCol: "I",
      unitVolumeCol: "N",
      totalVolumeCol: "O",
      divisor: 1000000,
      adjustSide: "all" as "all" | "length" | "width" | "height",
      dimensionDecimals: 0,
      maxChangePercent: 10,
      wholeNumberTarget: true,
      tolerance: 0.01,
      packageHeight: 1.1,
    },

    /**
     * Cấu hình chuẩn hóa mã hàng (Cột C).
     * Mã bắt đầu bằng số 9: giữ toàn bộ dãy số đầu mã (≥ 6 số, kể cả 7 số), chỉ bỏ hậu tố bắt đầu bằng ký tự không phải số.
     */
    autoCleanProductCode: {
      enabled: true,
      column: "C",
    },

    /**
     * Các cột cần định dạng số (giới hạn tối đa 2 số thập phân).
     * Ví dụ: 55.254 -> 55.25, 252.000 -> 252 (định dạng General).
     */
    numberFormatColumns: ["M", "O", "N"],

    /**
     * Cột luôn hiện đủ 2 chữ số thập phân (4400 → 4400.00).
     * Ghi đè General ở bước định dạng số.
     */
    fixedTwoDecimalColumns: ["M"],
  },
};

// ─── HELPER: Map giá trị string → enum ExcelScript ─────────────────
type HAlign = "center" | "left" | "right";
type VAlign = "top" | "middle" | "bottom";

function toHorizontalAlignment(value: HAlign): ExcelScript.HorizontalAlignment {
  const map: Record<HAlign, ExcelScript.HorizontalAlignment> = {
    center: ExcelScript.HorizontalAlignment.center,
    left: ExcelScript.HorizontalAlignment.left,
    right: ExcelScript.HorizontalAlignment.right,
  };
  return map[value] ?? ExcelScript.HorizontalAlignment.left;
}

function toVerticalAlignment(value: VAlign): ExcelScript.VerticalAlignment {
  const map: Record<VAlign, ExcelScript.VerticalAlignment> = {
    top: ExcelScript.VerticalAlignment.top,
    middle: ExcelScript.VerticalAlignment.center,
    bottom: ExcelScript.VerticalAlignment.bottom,
  };
  return map[value] ?? ExcelScript.VerticalAlignment.center;
}

// ─── HELPER: Tìm hàng tiêu đề chứa tất cả keywords ────────────────
/**
 * Quét cột A trong usedRange để tìm hàng chứa TẤT CẢ headerKeywords.
 * Trả về row index (0-based) hoặc -1 nếu không tìm thấy.
 */
function findHeaderRow(
  sheet: ExcelScript.Worksheet,
  keywords: string[],
): number {
  let usedRange = sheet.getUsedRange();
  if (!usedRange) return -1;

  let lastUsedRow = usedRange.getRowCount();
  let colAValues = sheet.getRange(`A1:A${lastUsedRow}`).getValues();

  for (let i = 0; i < colAValues.length; i++) {
    let cellValue = String(colAValues[i][0]).toUpperCase();
    let allMatch = keywords.every((kw) => cellValue.includes(kw.toUpperCase()));
    if (allMatch) return i; // 0-based
  }
  return -1;
}

/** Lấy chuỗi nhận diện từ hàng đầu sheet mẫu "footer" (cùng workbook). */
function getFooterSignatureFromWorkbook(
  workbook: ExcelScript.Workbook,
): string | null {
  let footerSheet = workbook.getWorksheet("footer");
  if (!footerSheet) return null;
  let fullUsedRange = footerSheet.getUsedRange();
  if (!fullUsedRange) return null;
  let footerTop = fullUsedRange.getIntersection("1:1");
  if (!footerTop) return null;
  let row = footerTop.getValues()[0] as unknown[];
  let parts: string[] = [];
  for (let c of row) {
    let s = String(c ?? "").trim();
    if (s) parts.push(s);
  }
  if (parts.length === 0) return null;
  let best = parts.reduce((a, b) => (a.length >= b.length ? a : b), "");
  if (best.length >= 4) return best;
  let joined = parts.join(" ");
  return joined.length >= 6 ? joined : null;
}

/**
 * Dòng đầu tiên (1-based) từ `fromRow` có bất kỳ ô nào (A→lastCol) chứa chữ ký footer.
 * Quét nhiều cột vì sau lần chạy trước cột A có thể đã bị ghi công thức STT.
 */
function findFooterStartRowFromSignature(
  sheet: ExcelScript.Worksheet,
  fromRow: number,
  toRow: number,
  signature: string,
  lastColLetter: string = "Q",
): number | null {
  let needle = signature.trim().toUpperCase();
  if (needle.length < 4) return null;
  if (needle.length > 40) needle = needle.substring(0, 40);
  let rows = sheet.getRange(`A${fromRow}:${lastColLetter}${toRow}`).getValues();
  for (let i = 0; i < rows.length; i++) {
    for (let j = 0; j < rows[i].length; j++) {
      let cell = String(rows[i][j] ?? "")
        .trim()
        .toUpperCase();
      if (cell.length > 0 && cell.includes(needle)) return fromRow + i;
    }
  }
  return null;
}

/** Dòng tổng cộng đã có (ô SUM) — dùng khi chạy lại script. */
function findExistingSummaryRow(
  sheet: ExcelScript.Worksheet,
  fromRow: number,
  toRow: number,
  summaryColumns: string[],
): number | null {
  if (summaryColumns.length === 0 || toRow < fromRow) return null;
  let col = summaryColumns[0];
  let formulas = sheet
    .getRange(`${col}${fromRow}:${col}${toRow}`)
    .getFormulas();
  for (let i = 0; i < formulas.length; i++) {
    let f = String(formulas[i][0] ?? "").trim();
    if (/^=SUM\(/i.test(f)) return fromRow + i;
  }
  return null;
}

/**
 * Xác định vùng dữ liệu (từ startRow đến dòng cuối cùng có tên hàng ở cột D).
 * Cho phép hàng trống xen giữa. Nếu có sheet "footer", chỉ quét đến trước dòng
 * đầu tiên của footer trên sheet (bỏ dữ liệu dán nhầm sau footer).
 */
function findDataBounds(
  sheet: ExcelScript.Worksheet,
  startRow: number,
  workbook: ExcelScript.Workbook,
  summaryColumns: string[] = [],
  lastColLetter: string = "Q",
): { startRow: number; lastRow: number } | null {
  let usedRange = sheet.getUsedRange();
  if (!usedRange) return null;

  let totalRowsInSheet = usedRange.getRowIndex() + usedRange.getRowCount();
  if (totalRowsInSheet < startRow) return null;

  let scanEndRow = totalRowsInSheet;
  let sig = getFooterSignatureFromWorkbook(workbook);
  if (sig) {
    let footerStart = findFooterStartRowFromSignature(
      sheet,
      startRow,
      totalRowsInSheet,
      sig,
      lastColLetter,
    );
    if (footerStart !== null && footerStart > startRow) {
      scanEndRow = footerStart - 1;
      console.log(
        `📎 Giới hạn bảng dữ liệu trước footer (nhận diện sheet "footer", dòng footer: ${footerStart}).`,
      );
    }
  }

  if (summaryColumns.length > 0) {
    let summaryRow = findExistingSummaryRow(
      sheet,
      startRow,
      scanEndRow,
      summaryColumns,
    );
    if (summaryRow !== null && summaryRow > startRow + 1) {
      let capBeforeSummary = summaryRow - 2;
      if (capBeforeSummary < scanEndRow) {
        scanEndRow = capBeforeSummary;
        console.log(
          `📎 Giới hạn bảng dữ liệu trước dòng tổng (dòng tổng: ${summaryRow}).`,
        );
      }
    }
  }

  if (scanEndRow < startRow) return null;

  let colDValues = sheet.getRange(`D${startRow}:D${scanEndRow}`).getValues();

  // Dòng cuối = hàng xa nhất trong vùng quét mà cột D còn tên hàng.
  let lastDataRowOffset = -1;
  for (let i = 0; i < colDValues.length; i++) {
    if (String(colDValues[i][0]).trim() !== "") {
      lastDataRowOffset = i;
    }
  }

  if (lastDataRowOffset === -1) return null;

  let lastRow = startRow + lastDataRowOffset;
  return { startRow, lastRow };
}

/** Số thứ ba trong chuỗi kích thước (VD: 44*24*1,1 → 1.1). */
function getDimensionThirdValue(dimStr: string): number | null {
  let s = String(dimStr ?? "").trim();
  if (!s) return null;
  let normalized = s.replace(/x|X/g, "*").replace(/,/g, ".");
  let parts = normalized
    .split("*")
    .map((p) => p.trim())
    .filter((p) => p.length > 0);
  if (parts.length < 3) return null;
  let n = Number(parts[2]);
  return Number.isFinite(n) ? n : null;
}

/** Chiều cao (số thứ 3) bằng packageHeight → đơn vị đóng gói PACKAGE. */
function isPackagePackUnit(
  dimStr: string,
  packageHeight: number = 1.1,
): boolean {
  let h = getDimensionThirdValue(dimStr);
  if (h === null) return false;
  return Math.abs(h - packageHeight) < 0.001;
}

function isBlankCellValue(val: unknown): boolean {
  if (val === "" || val === null || val === undefined) return true;
  return String(val).trim() === "";
}

/** Tính dòng đích khi sao chép dữ liệu từ sheet nguồn sang sheet sale (cùng cấu trúc). */
function resolvePxkDataCopyPlan(
  targetHeaderRowNum: number,
  sourceBounds: { startRow: number; lastRow: number } | null,
): { targetStartRow: number; targetLastRow: number } | null {
  if (!sourceBounds) return null;
  let rowCount = sourceBounds.lastRow - sourceBounds.startRow + 1;
  if (rowCount <= 0) return null;
  let targetStartRow = targetHeaderRowNum + 1;
  return {
    targetStartRow,
    targetLastRow: targetStartRow + rowCount - 1,
  };
}

type PxkDataCopyResult = {
  copied: boolean;
  sourceSheetName: string;
  sourceStartRow: number;
  sourceLastRow: number;
  targetStartRow: number;
  targetLastRow: number;
  firstCol: string;
  lastCol: string;
};

/**
 * Sao chép vùng dữ liệu từ sheet nguồn sang sheet đích (chỉ giá trị).
 * Trả về null nếu sheet nguồn không có dữ liệu hoặc không tìm thấy tiêu đề.
 */
function copyPxkDataFromSourceSheet(
  targetSheet: ExcelScript.Worksheet,
  sourceSheet: ExcelScript.Worksheet,
  workbook: ExcelScript.Workbook,
  targetHeaderRowNum: number,
  headerKeywords: string[],
  summaryColumns: string[],
  lastColLetter: string,
  firstCol: string,
  lastCol: string,
): PxkDataCopyResult | null {
  let sourceHeaderIdx = findHeaderRow(sourceSheet, headerKeywords);
  if (sourceHeaderIdx === -1) return null;

  let sourceBounds = findDataBounds(
    sourceSheet,
    sourceHeaderIdx + 2,
    workbook,
    summaryColumns,
    lastColLetter,
  );
  if (!sourceBounds) return null;

  let plan = resolvePxkDataCopyPlan(targetHeaderRowNum, sourceBounds);
  if (!plan) return null;

  let srcRange = sourceSheet.getRange(
    `${firstCol}${sourceBounds.startRow}:${lastCol}${sourceBounds.lastRow}`,
  );
  let tgtRange = targetSheet.getRange(
    `${firstCol}${plan.targetStartRow}:${lastCol}${plan.targetLastRow}`,
  );
  tgtRange.copyFrom(srcRange, ExcelScript.RangeCopyType.values);

  return {
    copied: true,
    sourceSheetName: sourceSheet.getName(),
    sourceStartRow: sourceBounds.startRow,
    sourceLastRow: sourceBounds.lastRow,
    targetStartRow: plan.targetStartRow,
    targetLastRow: plan.targetLastRow,
    firstCol,
    lastCol,
  };
}

// Expose helper functions for Node-based regression tests without affecting Excel execution.
const pxkTestables = {
  toHorizontalAlignment,
  toVerticalAlignment,
  findHeaderRow,
  getFooterSignatureFromWorkbook,
  findFooterStartRowFromSignature,
  findExistingSummaryRow,
  findDataBounds,
  getDimensionThirdValue,
  isPackagePackUnit,
  resolvePxkDataCopyPlan,
};

(
  globalThis as typeof globalThis & {
    __AUTO_PXK_TESTABLES__?: typeof pxkTestables;
  }
).__AUTO_PXK_TESTABLES__ = pxkTestables;

// ─── MAIN ───────────────────────────────────────────────────────────
// @ts-ignore: Office Scripts — mỗi file chạy độc lập, không trùng main()
type PxkDimensionParts = {
  values: number[];
  separator: string;
  decimalComma: boolean;
};

/** Tách kích thước "42*24,5*25" / "42x24.5x25" thành 3 số dương. */
function parsePxkDimensions(raw: unknown): PxkDimensionParts | null {
  let s = String(raw ?? "").trim();
  if (!s) return null;
  let parts = s.split(/\s*[*xX×]\s*/);
  if (parts.length !== 3) return null;
  let values = parts.map((p) => Number(p.trim().replace(/,/g, ".")));
  if (values.some((v) => !isFinite(v) || v <= 0)) return null;
  let sepMatch = s.match(/[*xX×]/);
  return {
    values,
    separator: sepMatch ? sepMatch[0] : "*",
    decimalComma: s.includes(","),
  };
}

function formatPxkDimensions(
  values: number[],
  separator: string,
  decimalComma: boolean,
  decimals: number,
): string {
  return values
    .map((v) => {
      let places = Math.max(0, decimals);
      let rounded = Number(v.toFixed(places));
      // Cạnh không nằm trên lưới làm tròn (ví dụ cạnh không được sửa) giữ nguyên số đã có.
      let text =
        Math.abs(rounded - v) < 1e-9
          ? String(rounded)
          : String(Number(v.toFixed(10)));
      return decimalComma ? text.replace(".", ",") : text;
    })
    .join(separator);
}

/**
 * Sửa kích thước theo nhóm (cùng chuỗi kích thước) để tổng thể tích khớp số khối đích.
 * Tổng mô phỏng đúng công thức Excel: SUM(ROUND(ROUND(dài*rộng*cao/divisor, 3) * SL, 2)).
 */
function autoAdjustPxkContainerVolume(
  sheet: ExcelScript.Worksheet,
  vCfg: typeof PXK_CONFIG.dataTable.autoAdjustContainerVolume,
  startRow: number,
  lastRow: number,
  summaryRow: number,
): void {
  const round2 = (n: number) => Math.round(n * 100) / 100;
  const round3 = (n: number) => Math.round(n * 1000) / 1000;
  const divisor = vCfg.divisor || 1000000;
  const decimals = Math.max(0, vCfg.dimensionDecimals ?? 0);
  const step = Math.pow(10, -decimals);
  const roundDim = (v: number) => Number((Math.round(v / step) * step).toFixed(decimals));
  // Số nguyên: bước kế tiếp đúng 1 cm, kể cả khi kích thước gốc đang là số thập phân (40,5 → 41 hoặc 40).
  const moveSide = (value: number, dir: number, k: number): number => {
    if (decimals !== 0) return roundDim(value + dir * k);
    let delta = (dir < 0 ? -1 : 1) * k;
    if (Math.abs(value - Math.round(value)) < 1e-9) return Math.round(value) + delta;
    if (delta > 0) return Math.floor(value + 1e-9) + delta;
    return Math.ceil(value - 1e-9) + delta;
  };
  const toWholeCm = (value: number, grow: boolean): number => {
    if (Math.abs(value - Math.round(value)) < 1e-9) return Math.round(value);
    let n = grow ? Math.ceil(value - 1e-9) : Math.floor(value + 1e-9);
    return n > 0 ? n : value;
  };
  const maxRatio = Math.max(0, Number(vCfg.maxChangePercent) || 0) / 100;
  const tolerance = Math.max(0.005, Number(vCfg.tolerance) || 0);

  let targetVolume = Number(vCfg.targetTotalVolume) || 0;
  let targetSource = "cấu hình targetTotalVolume";
  let targetCellAddress = `${vCfg.totalVolumeCol}${summaryRow + 1}`;
  if (targetVolume <= 0) {
    targetVolume = Number(sheet.getRange(targetCellAddress).getValue()) || 0;
    targetSource = `ô ${targetCellAddress}`;
    if (targetVolume > 0) {
      let targetFormat = sheet.getRange(targetCellAddress).getFormat();
      targetFormat.getFill().setColor("#FFFF00");
      targetFormat.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      sheet.getRange(`${summaryRow + 1}:${summaryRow + 1}`).setRowHidden(true);
    }
  }
  if (targetVolume <= 0) {
    console.log(
      `ℹ️ Điều chỉnh thể tích container: chưa có số khối đích (cấu hình targetTotalVolume = 0 và ô ${targetCellAddress} trống), bỏ qua.`,
    );
    return;
  }

  let sideIndexMap: Record<string, number> = { length: 0, width: 1, height: 2 };
  let sides =
    vCfg.adjustSide === "all" ? [0, 1, 2] : [sideIndexMap[vCfg.adjustSide] ?? 2];

  let productValues = sheet
    .getRange(`${vCfg.productCol}${startRow}:${vCfg.productCol}${lastRow}`)
    .getValues();
  let dimRange = sheet.getRange(`${vCfg.dimCol}${startRow}:${vCfg.dimCol}${lastRow}`);
  let dimValues = dimRange.getValues();
  let qtyValues = sheet
    .getRange(`${vCfg.qtyCol}${startRow}:${vCfg.qtyCol}${lastRow}`)
    .getValues();
  let unitVolRange = sheet.getRange(
    `${vCfg.unitVolumeCol}${startRow}:${vCfg.unitVolumeCol}${lastRow}`,
  );
  let unitVolValues = unitVolRange.getValues();
  let unitVolFormulas = unitVolRange.getFormulas();

  type VolumeGroup = {
    parts: PxkDimensionParts;
    original: number[];
    current: number[];
    rows: number[];
    qtys: number[];
    total: number;
  };
  let groups = new Map<string, VolumeGroup>();
  let fixedTotal = 0;
  let packageHeight = vCfg.packageHeight ?? 1.1;

  for (let i = 0; i < dimValues.length; i++) {
    let qty = Number(qtyValues[i][0]) || 0;
    let parts = parsePxkDimensions(dimValues[i][0]);
    if (!parts || Math.abs(parts.values[2] - packageHeight) < 0.001) {
      fixedTotal += round2((Number(unitVolValues[i][0]) || 0) * qty);
      continue;
    }
    let key = parts.values.join("*");
    if (!groups.has(key)) {
      groups.set(key, {
        parts,
        original: [...parts.values],
        current: [...parts.values],
        rows: [],
        qtys: [],
        total: 0,
      });
    }
    let group = groups.get(key)!;
    group.rows.push(i);
    group.qtys.push(qty);
  }

  let groupList = Array.from(groups.values());
  if (groupList.length === 0) {
    console.log("ℹ️ Điều chỉnh thể tích container: không có dòng nào có kích thước hợp lệ, bỏ qua.");
    return;
  }

  const unitVolume = (dims: number[]) => round3((dims[0] * dims[1] * dims[2]) / divisor);
  const groupTotal = (g: VolumeGroup, dims: number[]) => {
    let uv = unitVolume(dims);
    return g.qtys.reduce((sum, q) => sum + round2(uv * q), 0);
  };
  const withinLimit = (g: VolumeGroup, dims: number[]) => {
    let base = g.original[0] * g.original[1] * g.original[2];
    let next = dims[0] * dims[1] * dims[2];
    return Math.abs(next / base - 1) <= maxRatio + 1e-9;
  };

  let originalTotal = fixedTotal;
  for (let g of groupList) {
    g.total = groupTotal(g, g.original);
    originalTotal += g.total;
  }
  originalTotal = round2(originalTotal);

  let wholeNumber = vCfg.wholeNumberTarget !== false;
  let rangeLow = wholeNumber ? Math.floor(targetVolume) : targetVolume - tolerance;
  let rangeHigh = wholeNumber ? Math.floor(targetVolume) + 0.99 : targetVolume + tolerance;
  let rangeText = wholeNumber
    ? `${rangeLow}–${round2(rangeHigh)} m³`
    : `${targetVolume} ± ${tolerance} m³`;
  const distanceToRange = (t: number) => {
    let v = round2(t);
    if (v < rangeLow - 1e-9) return round2(rangeLow - v);
    if (v > rangeHigh + 1e-9) return round2(v - rangeHigh);
    return 0;
  };

  if (distanceToRange(originalTotal) === 0) {
    console.log(
      `ℹ️ Điều chỉnh thể tích container: tổng hiện tại ${originalTotal} m³ đã nằm trong khoảng đích ${rangeText}.`,
    );
    return;
  }

  // Nhắm vào mép gần nhất của khoảng đích, chừa một chút để làm tròn không đẩy ra ngoài.
  let margin = Math.min(0.1, (rangeHigh - rangeLow) / 2);
  let aimVolume = originalTotal < rangeLow ? rangeLow + margin : rangeHigh - margin;
  let adjustableTotal = originalTotal - fixedTotal;
  let scale = adjustableTotal > 0 ? (aimVolume - fixedTotal) / adjustableTotal : 0;
  if (scale <= 0) {
    console.log(
      `⚠️ Điều chỉnh thể tích container: không thể đạt ${targetVolume} m³ (phần có kích thước hợp lệ: ${round2(adjustableTotal)} m³, phần cố định: ${round2(fixedTotal)} m³).`,
    );
    return;
  }
  if (Math.abs(scale - 1) > maxRatio + 1e-9) {
    console.log(
      `⚠️ Điều chỉnh thể tích container: cần thay đổi ${round2((scale - 1) * 100)}% thể tích, vượt giới hạn maxChangePercent = ${vCfg.maxChangePercent}%. ` +
        `Tổng hiện tại ${originalTotal} m³, số khối đích ${targetVolume} m³ — không điều chỉnh.`,
    );
    return;
  }

  // Đưa các cạnh được sửa về số cm nguyên theo hướng cần tăng hoặc giảm.
  if (decimals === 0) {
    let grow = originalTotal < rangeLow;
    for (let g of groupList) {
      let snapped = g.original.map((v, idx) =>
        sides.indexOf(idx) !== -1 ? toWholeCm(v, grow) : v,
      );
      if (snapped.every((v) => v > 0) && withinLimit(g, snapped)) {
        g.current = snapped;
        continue;
      }
      let progressive = [...g.original];
      for (let idx of sides) {
        let trial = [...progressive];
        trial[idx] = toWholeCm(g.original[idx], grow);
        if (trial[idx] > 0 && withinLimit(g, trial)) progressive = trial;
      }
      g.current = progressive;
    }
  }

  let total = fixedTotal;
  for (let g of groupList) {
    g.total = groupTotal(g, g.current);
    total += g.total;
  }
  total = round2(total);

  if (distanceToRange(total) !== 0) {
    adjustableTotal = round2(total - fixedTotal);
    scale = adjustableTotal > 0 ? (aimVolume - fixedTotal) / adjustableTotal : 0;
    if (scale > 0 && Math.abs(scale - 1) <= maxRatio + 1e-9) {
      let sideFactor = vCfg.adjustSide === "all" ? Math.pow(scale, 1 / 3) : scale;
      let scaledTotal = fixedTotal;
      for (let g of groupList) {
        let scaled = g.current.map((v, idx) =>
          sides.indexOf(idx) !== -1 ? roundDim(v * sideFactor) : v,
        );
        if (
          withinLimit(g, scaled) &&
          scaled.every((v) => v > 0) &&
          unitVolume(scaled) !== unitVolume(g.original)
        ) {
          g.current = scaled;
        }
        g.total = groupTotal(g, g.current);
        scaledTotal += g.total;
      }
      total = round2(scaledTotal);
    }

    // Thể tích đơn vị làm tròn 3 chữ số nên 1 cm thường chưa đổi thể tích:
    // với mỗi cạnh, tìm số cm nguyên ít nhất làm thể tích đơn vị thay đổi rồi chọn phương án giảm sai số nhiều nhất.
    const maxStepsPerMove = 1000;
    for (let iter = 0; iter < 2000; iter++) {
      let currentErr = distanceToRange(total);
      if (currentErr === 0) break;
      let dir = round2(total) < rangeLow ? step : -step;
      let best: { g: VolumeGroup; dims: number[]; total: number } | null = null;
      let bestErr = currentErr;

      for (let g of groupList) {
        let currentUnit = unitVolume(g.current);
        for (let side of sides) {
          let candidate: number[] | null = null;
          for (let k = 1; k <= maxStepsPerMove; k++) {
            let trial = [...g.current];
            trial[side] = moveSide(g.current[side], dir, k);
            if (trial[side] <= 0 || !withinLimit(g, trial)) break;
            if (unitVolume(trial) !== currentUnit) {
              candidate = trial;
              break;
            }
          }
          if (!candidate) continue;
          let candidateGroupTotal = groupTotal(g, candidate);
          let err = distanceToRange(total - g.total + candidateGroupTotal);
          if (err < bestErr - 1e-9) {
            bestErr = err;
            best = { g, dims: candidate, total: candidateGroupTotal };
          }
        }
      }

      if (!best) break;
      total = total - best.g.total + best.total;
      best.g.current = best.dims;
      best.g.total = best.total;
    }
    total = round2(total);
  }

  let logMessages: string[] = [];
  for (let g of groupList) {
    let changed = g.current.some((v, idx) => v !== g.original[idx]);
    let dimText = changed
      ? formatPxkDimensions(g.current, g.parts.separator, g.parts.decimalComma, decimals)
      : null;
    let originalText = String(dimValues[g.rows[0]][0] ?? "").trim();
    let formula = `=ROUND((${g.current.join("*")})/${divisor}, 3)`;
    for (let i of g.rows) {
      if (dimText !== null) dimValues[i][0] = dimText;
      unitVolFormulas[i][0] = formula;
    }
    if (changed) {
      let productCode = String(productValues[g.rows[0]][0] ?? "").trim();
      logMessages.push(
        `- Mã ${productCode}: ${originalText} ➔ ${dimText} ` +
          `(${unitVolume(g.original)} ➔ ${unitVolume(g.current)} m³/đơn vị, ${g.rows.length} dòng)`,
      );
    }
  }

  dimRange.setValues(dimValues);
  unitVolRange.setFormulas(unitVolFormulas);

  console.log(
    `✅ Đã điều chỉnh thể tích container (số khối đích lấy từ ${targetSource}): ${originalTotal} ➔ ${total} m³, ` +
      `khoảng đích ${rangeText}.`,
  );
  if (logMessages.length > 0) {
    console.log(`📝 Chi tiết điều chỉnh kích thước:\n${logMessages.join("\n")}`);
  }
  if (distanceToRange(total) > 0) {
    console.log(
      `⚠️ Chưa vào khoảng đích ${rangeText} (còn lệch ${distanceToRange(total)} m³) — thử tăng maxChangePercent hoặc đổi adjustSide.`,
    );
  }
}

Object.assign(pxkTestables, { autoAdjustPxkContainerVolume });

async function main(workbook: ExcelScript.Workbook) {
  const CONFIG = PXK_CONFIG;

  let sheet = workbook.getActiveWorksheet();
  let cfg = CONFIG;
  let saleSheetName = (cfg.dataTable.saleSheetName || "sale").trim().toLowerCase();
  let isSaleSheet = sheet.getName().trim().toLowerCase() === saleSheetName;
  if (!isSaleSheet) {
    console.log(
      `ℹ️ Sheet hiện tại là "${sheet.getName()}" — chỉ định dạng. Bù trừ cân nặng / điều chỉnh thể tích chỉ chạy trên sheet "${cfg.dataTable.saleSheetName || "sale"}".`,
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 1: Áp dụng font chung cho toàn bộ sheet
  // ═══════════════════════════════════════════════════════════════════
  let entireFont = sheet.getRange().getFormat().getFont();
  entireFont.setName(cfg.globalFont.name);
  entireFont.setSize(cfg.globalFont.size);
  console.log(
    `✅ Đã áp dụng font: ${cfg.globalFont.name} ${cfg.globalFont.size}pt`,
  );

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 2: Tìm hàng tiêu đề (chứa headerKeywords trong cột A)
  // ═══════════════════════════════════════════════════════════════════
  let headerRowIdx = findHeaderRow(sheet, cfg.dataTable.headerKeywords);
  if (headerRowIdx === -1) {
    console.log(
      `❌ Không tìm thấy hàng tiêu đề chứa [${cfg.dataTable.headerKeywords.join(", ")}] trong cột A.`,
    );
    return;
  }
  let headerRowNum = headerRowIdx + 1; // 1-based (số dòng trên Excel)
  console.log(`📍 Tìm thấy hàng tiêu đề: dòng ${headerRowNum}`);

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 2.5: Sao chép dữ liệu từ pre-packing-list nếu sheet sale trống
  // ═══════════════════════════════════════════════════════════════════
  let lastColLetter =
    cfg.dataTable.columns[cfg.dataTable.columns.length - 1].letter;
  let summaryColumns = cfg.dataTable.summaryColumns ?? [];

  if (
    isSaleSheet &&
    cfg.dataTable.copyFromSourceWhenEmpty &&
    cfg.dataTable.copyFromSourceWhenEmpty.enabled
  ) {
    let saleHasData = findDataBounds(
      sheet,
      headerRowNum + 1,
      workbook,
      summaryColumns,
      lastColLetter,
    );
    if (!saleHasData) {
      let copyCfg = cfg.dataTable.copyFromSourceWhenEmpty;
      let sourceSheetName = copyCfg.sourceSheetName || "pre-packing-list";
      let sourceSheet = workbook.getWorksheet(sourceSheetName);

      if (sourceSheet) {
        let copyResult = copyPxkDataFromSourceSheet(
          sheet,
          sourceSheet,
          workbook,
          headerRowNum,
          cfg.dataTable.headerKeywords,
          summaryColumns,
          lastColLetter,
          copyCfg.firstCol || "A",
          copyCfg.lastCol || "Q",
        );
        if (copyResult) {
          console.log(
            `✅ Sheet "sale" trống — đã sao chép dữ liệu từ "${copyResult.sourceSheetName}" ` +
              `(dòng nguồn ${copyResult.sourceStartRow}→${copyResult.sourceLastRow} ` +
              `→ dòng đích ${copyResult.targetStartRow}→${copyResult.targetLastRow}, ` +
              `cột ${copyResult.firstCol}:${copyResult.lastCol}).`,
          );
        } else {
          console.log(
            `⚠️ Sheet "sale" trống nhưng "${sourceSheetName}" không có dữ liệu bảng để sao chép.`,
          );
        }
      } else {
        console.log(
          `ℹ️ Sheet "sale" trống nhưng không tìm thấy sheet "${sourceSheetName}", bỏ qua sao chép.`,
        );
      }
    }
  }

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 3: Xác định vùng dữ liệu (dòng đầu → dòng cuối)
  // ═══════════════════════════════════════════════════════════════════
  let bounds = findDataBounds(
    sheet,
    headerRowNum + 1,
    workbook,
    summaryColumns,
    lastColLetter,
  );
  if (!bounds) {
    console.log(
      `⚠️ Không tìm thấy dữ liệu sau hàng tiêu đề (dòng ${headerRowNum}).`,
    );
    return;
  }
  let { startRow, lastRow } = bounds;
  console.log(`📊 Vùng dữ liệu: dòng ${startRow} → ${lastRow}`);

  let usedRangeForTail = sheet.getUsedRange();
  let sheetLastRow = usedRangeForTail
    ? usedRangeForTail.getRowIndex() + usedRangeForTail.getRowCount()
    : lastRow;
  let footerSignature = getFooterSignatureFromWorkbook(workbook);
  let existingFooterRow: number | null = null;
  if (footerSignature) {
    existingFooterRow = findFooterStartRowFromSignature(
      sheet,
      startRow,
      sheetLastRow,
      footerSignature,
      lastColLetter,
    );
  }
  let existingSummaryRow = findExistingSummaryRow(
    sheet,
    startRow,
    existingFooterRow !== null ? existingFooterRow - 1 : sheetLastRow,
    summaryColumns,
  );
  let summaryRow =
    existingSummaryRow !== null ? existingSummaryRow : lastRow + 2;

  // Điền công thức STT: =ROW()-headerRow cho toàn bộ hàng dữ liệu (nhanh và chính xác)
  sheet
    .getRange(`A${startRow}:A${lastRow}`)
    .setFormula(`=ROW()-${headerRowNum}`);
  console.log(
    `✅ Đã điền công thức STT (=ROW()-${headerRowNum}) cho dải ô A${startRow}:A${lastRow}.`,
  );

  // Kẻ khung (All Borders) cho toàn bộ bảng (tiêu đề → hàng cuối)
  // Đảm bảo các ô STT mới thêm công thức cũng có border.
  let fullTableRange = sheet.getRange(
    `A${headerRowNum}:${lastColLetter}${lastRow}`,
  );
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
  console.log(
    `✅ Đã kẻ khung (All Borders) cho dải A${headerRowNum}:${lastColLetter}${lastRow}`,
  );

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 4: Định dạng hàng tiêu đề
  // ═══════════════════════════════════════════════════════════════════
  let firstCol = cfg.dataTable.columns[0].letter;
  let lastCol = cfg.dataTable.columns[cfg.dataTable.columns.length - 1].letter;
  let headerRange = sheet.getRange(
    `${firstCol}${headerRowNum}:${lastCol}${headerRowNum}`,
  );
  let headerFormat = headerRange.getFormat();

  // Bold
  headerFormat.getFont().setBold(cfg.dataTable.headerStyle.bold);

  // Horizontal & Vertical alignment
  headerFormat.setHorizontalAlignment(
    toHorizontalAlignment(cfg.dataTable.headerStyle.horizontalAlignment),
  );
  headerFormat.setVerticalAlignment(
    toVerticalAlignment(cfg.dataTable.headerStyle.verticalAlignment),
  );
  console.log("✅ Đã định dạng hàng tiêu đề (bold, center, middle).");

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 5: Merge các vùng tiêu đề được cấu hình
  // ═══════════════════════════════════════════════════════════════════
  cfg.dataTable.headerMergeRegions.forEach((region) => {
    let mergeRange = sheet.getRange(
      `${region.from}${headerRowNum}:${region.to}${headerRowNum}`,
    );
    mergeRange.merge(false); // false = merge all (không phải merge across)
  });
  console.log(
    `✅ Đã merge tiêu đề: ${cfg.dataTable.headerMergeRegions.map((r) => `${r.from}:${r.to}`).join(", ")}`,
  );

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 6: Định dạng vùng dữ liệu — Vertical alignment (middle)
  // ═══════════════════════════════════════════════════════════════════
  let dataRange = sheet.getRange(`${firstCol}${startRow}:${lastCol}${lastRow}`);
  dataRange
    .getFormat()
    .setVerticalAlignment(
      toVerticalAlignment(cfg.dataTable.dataVerticalAlignment),
    );
  console.log("✅ Đã căn middle align cho toàn bộ hàng dữ liệu.");

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 7: Căn ngang (horizontal alignment) theo từng cột
  // ═══════════════════════════════════════════════════════════════════
  cfg.dataTable.columns.forEach((col) => {
    let colRange = sheet.getRange(
      `${col.letter}${startRow}:${col.letter}${lastRow}`,
    );
    colRange
      .getFormat()
      .setHorizontalAlignment(toHorizontalAlignment(col.align));
  });
  console.log("✅ Đã căn ngang cho từng cột theo cấu hình.");

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 8: Autofit & No-wrap cho các cột được cấu hình
  // ═══════════════════════════════════════════════════════════════════
  // 8a. Tắt wrap text trước khi autofit (để autofit tính đúng chiều rộng)
  cfg.dataTable.noWrapColumns.forEach((colLetter) => {
    let colRange = sheet.getRange(
      `${colLetter}${startRow}:${colLetter}${lastRow}`,
    );
    colRange.getFormat().setWrapText(false);
  });

  // 8b. Autofit columns
  cfg.dataTable.autofitColumns.forEach((colLetter) => {
    let colRange = sheet.getRange(
      `${colLetter}${startRow}:${colLetter}${lastRow}`,
    );
    colRange.getFormat().autofitColumns();
  });
  console.log(
    `✅ Đã autofit cột: [${cfg.dataTable.autofitColumns.join(", ")}], no-wrap: [${cfg.dataTable.noWrapColumns.join(", ")}]`,
  );

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 8.5: Chuẩn hóa Mã Hàng (Cột C) - Giữ toàn bộ dãy số đầu mã bắt đầu bằng số 9, bỏ hậu tố
  // ═══════════════════════════════════════════════════════════════════
  if (
    cfg.dataTable.autoCleanProductCode &&
    cfg.dataTable.autoCleanProductCode.enabled
  ) {
    let productCodeCol = cfg.dataTable.autoCleanProductCode.column;
    let productCodeRange = sheet.getRange(
      `${productCodeCol}${startRow}:${productCodeCol}${lastRow}`,
    );
    let productCodeValues = productCodeRange.getValues();
    let cleanedLogs: string[] = [];
    let hasCleaned = false;

    for (let i = 0; i < productCodeValues.length; i++) {
      let originalVal = String(productCodeValues[i][0] || "").trim();
      // Regex: mã ≥ 6 số bắt đầu bằng 9, theo sau là phần dư bắt đầu bằng ký tự không phải số.
      // Cặp A/B (912345A / 912345B) là 2 sản phẩm riêng nên không xóa hậu tố.
      let match = /^9\d{5,}[AB]$/i.test(originalVal)
        ? null
        : originalVal.match(/^(9\d{5,})(\D.*)$/);

      if (match) {
        let cleanedVal = match[1];
        let suffix = match[2]; // Phần dư bị xóa
        productCodeValues[i][0] = cleanedVal;
        cleanedLogs.push(
          `- Dòng ${startRow + i}: ${originalVal} ➔ ${cleanedVal} (Đã xóa "${suffix}")`,
        );
        hasCleaned = true;
      }
    }

    if (hasCleaned) {
      productCodeRange.setValues(productCodeValues);
      console.log(
        `✅ Đã chuẩn hóa Mã Hàng (Cột ${productCodeCol}):\n${cleanedLogs.join("\n")}`,
      );
    }
  }

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 9: Ẩn các cột được cấu hình
  // ═══════════════════════════════════════════════════════════════════
  if (cfg.dataTable.columnsToHide.length > 0) {
    cfg.dataTable.columnsToHide.forEach((colStr) => {
      // Nếu là cột đơn (VD: "H"), chuyển thành định dạng range (VD: "H:H")
      let rangeAddress = colStr.includes(":") ? colStr : `${colStr}:${colStr}`;
      sheet.getRange(rangeAddress).setColumnHidden(true);
    });
    console.log(`✅ Đã ẩn cột: [${cfg.dataTable.columnsToHide.join(", ")}]`);
  }

  // BƯỚC 10: Đã thực hiện kẻ khung ở Bước 3 để đảm bảo STT có border.
  // (Giữ lại check applyAllBorders nếu muốn tùy biến sau này, nhưng logic chính đã nằm trên)

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 11: Tính Thể Tích từ Kích Thước nếu bị trống
  // ═══════════════════════════════════════════════════════════════════
  if (
    cfg.dataTable.autoCalculateVolume &&
    cfg.dataTable.autoCalculateVolume.enabled
  ) {
    let sourceCol = cfg.dataTable.autoCalculateVolume.sourceCol;
    let targetCol = cfg.dataTable.autoCalculateVolume.targetCol;
    let divisor = cfg.dataTable.autoCalculateVolume.divisor;

    let targetRange = sheet.getRange(
      `${targetCol}${startRow}:${targetCol}${lastRow}`,
    );
    let targetValues = targetRange.getValues();
    let sourceRange = sheet.getRange(
      `${sourceCol}${startRow}:${sourceCol}${lastRow}`,
    );
    let sourceValues = sourceRange.getValues();

    let updated = false;
    for (let i = 0; i < targetValues.length; i++) {
      let tVal = targetValues[i][0];
      if (tVal === "" || tVal === null || tVal === undefined) {
        let sVal = String(sourceValues[i][0] || "").trim();
        if (sVal) {
          // Chuẩn hóa chuỗi kích thước: đổi 'x'/'X' thành '*', đổi ',' thành '.'
          let mathExpr = sVal.replace(/x|X/g, "*").replace(/,/g, ".");
          // Tạo công thức (VD: =ROUND((42*24.5*25)/1000000, 3))
          let formula = `=ROUND((${mathExpr})/${divisor}, 3)`;
          sheet.getRange(`${targetCol}${startRow + i}`).setFormula(formula);
          updated = true;
        }
      }
    }
    if (updated) {
      console.log(
        `✅ Đã tự động tính Thể Tích (cột ${targetCol}) từ Kích Thước (cột ${sourceCol}) cho các ô trống.`,
      );
    }
  }

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 11.5: Điền ĐVT cột G (PCS) và cột J (PACKAGE / THÙNG) nếu trống
  // ═══════════════════════════════════════════════════════════════════
  if (
    cfg.dataTable.autoFillUnitColumns &&
    cfg.dataTable.autoFillUnitColumns.enabled
  ) {
    let uCfg = cfg.dataTable.autoFillUnitColumns;
    let gCol = uCfg.qtyUnitCol;
    let jCol = uCfg.packUnitCol;
    let dimCol = uCfg.dimCol;
    let packageHeight = uCfg.packageHeight ?? 1.1;

    let gRange = sheet.getRange(`${gCol}${startRow}:${gCol}${lastRow}`);
    let jRange = sheet.getRange(`${jCol}${startRow}:${jCol}${lastRow}`);
    let dimRange = sheet.getRange(`${dimCol}${startRow}:${dimCol}${lastRow}`);
    let gValues = gRange.getValues();
    let jValues = jRange.getValues();
    let dimValues = dimRange.getValues();

    let filledG = 0;
    let filledJPackage = 0;
    let filledJCarton = 0;

    for (let i = 0; i < gValues.length; i++) {
      if (isBlankCellValue(gValues[i][0])) {
        gValues[i][0] = uCfg.qtyUnitDefault;
        filledG++;
      }

      if (isBlankCellValue(jValues[i][0])) {
        let dimStr = String(dimValues[i][0] ?? "").trim();
        if (isPackagePackUnit(dimStr, packageHeight)) {
          jValues[i][0] = uCfg.packUnitPackage;
          filledJPackage++;
        } else {
          jValues[i][0] = uCfg.packUnitCarton;
          filledJCarton++;
        }
      }
    }

    if (filledG > 0) gRange.setValues(gValues);
    if (filledJPackage + filledJCarton > 0) jRange.setValues(jValues);

    if (filledG > 0 || filledJPackage + filledJCarton > 0) {
      console.log(
        `✅ Đã điền ĐVT: cột ${gCol}="${uCfg.qtyUnitDefault}" (${filledG} ô), cột ${jCol} PACKAGE (${filledJPackage} ô), THÙNG CARTON (${filledJCarton} ô).`,
      );
    }
  }

  // Tập hợp dòng bị đánh dấu "thiếu số lô" từ BƯỚC 11.8 → BƯỚC 12 sẽ bỏ qua
  let skipBatchRows = new Set<number>();

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 11.8: Tra cứu Số Lô từ sheet riêng (lot-no)
  // ═══════════════════════════════════════════════════════════════════
  if (
    cfg.dataTable.autoLookupBatchNumber &&
    cfg.dataTable.autoLookupBatchNumber.enabled
  ) {
    let lkCfg = cfg.dataTable.autoLookupBatchNumber;
    let lotSheet = workbook.getWorksheet(lkCfg.sheetName);

    if (lotSheet) {
      let lotUsedRange = lotSheet.getUsedRange();
      if (lotUsedRange && lotUsedRange.getRowCount() > 1) {
        let lotValues = lotUsedRange.getValues();

        // Xác định chỉ số cột tra cứu (A=0, B=1, ...)
        let lkContractIdx = lkCfg.lookupContractCol.charCodeAt(0) - 65;
        let lkBatchIdx = lkCfg.lookupBatchCol.charCodeAt(0) - 65;

        // Xây map: mã hợp đồng (uppercase, trim) => số lô
        let lotMap = new Map<string, string>();
        for (let i = 1; i < lotValues.length; i++) {
          let contractVal = String(lotValues[i][lkContractIdx] || "")
            .trim()
            .toUpperCase();
          let batchVal = String(lotValues[i][lkBatchIdx] || "").trim();
          if (contractVal && batchVal) {
            lotMap.set(contractVal, batchVal);
          }
        }

        // Đọc cột mã hợp đồng và cột số lô trên sheet hiện tại
        let tgtContractCol = lkCfg.targetContractCol;
        let tgtBatchCol = lkCfg.targetBatchCol;
        let contractRange = sheet.getRange(
          `${tgtContractCol}${startRow}:${tgtContractCol}${lastRow}`,
        );
        let batchRange = sheet.getRange(
          `${tgtBatchCol}${startRow}:${tgtBatchCol}${lastRow}`,
        );
        let contractValues = contractRange.getValues();
        let batchValues = batchRange.getValues();

        let filledCount = 0;
        let missingRows: number[] = []; // Dòng không tìm thấy số lô tương ứng
        for (let i = 0; i < contractValues.length; i++) {
          let bVal = String(batchValues[i][0] || "").trim();
          // Chỉ điền khi ô số lô đang trống
          if (bVal === "" || bVal === "0") {
            let cVal = String(contractValues[i][0] || "")
              .trim()
              .toUpperCase();
            if (cVal) {
              let lotBatch = lotMap.get(cVal);
              if (lotBatch) {
                batchValues[i][0] = lotBatch;
                filledCount++;
              } else {
                missingRows.push(startRow + i);
              }
            }
          }
        }

        // Tô đỏ các ô số lô không tìm thấy giá trị tương ứng từ sheet tra cứu
        // và đánh dấu để BƯỚC 12 không ghi đè
        if (missingRows.length > 0) {
          for (let row of missingRows) {
            sheet
              .getRange(`${tgtBatchCol}${row}`)
              .getFormat()
              .getFill()
              .setColor("#FF0000");
            skipBatchRows.add(row);
          }
          console.log(
            `⚠️ ${missingRows.length} ô Số Lô không tìm thấy trong sheet "${lkCfg.sheetName}" (đã tô đỏ): dòng ${missingRows.join(", ")}.`,
          );
        }

        if (filledCount > 0) {
          batchRange.setValues(batchValues);
          console.log(
            `✅ Đã tra cứu và điền ${filledCount} Số Lô từ sheet "${lkCfg.sheetName}" vào cột ${tgtBatchCol}.`,
          );
        } else {
          console.log(
            `ℹ️ Không có ô Số Lô nào cần điền từ sheet "${lkCfg.sheetName}".`,
          );
        }
      } else {
        console.log(
          `⚠️ Sheet "${lkCfg.sheetName}" không có dữ liệu (cần ít nhất 1 hàng tiêu đề + 1 hàng dữ liệu).`,
        );
      }
    } else {
      console.log(
        `ℹ️ Không tìm thấy sheet "${lkCfg.sheetName}", bỏ qua tra cứu Số Lô.`,
      );
    }
  }

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 12: Điền Số Lô (Cột K) bị thiếu dựa trên Mã Hợp Đồng (Cột B)
  // ═══════════════════════════════════════════════════════════════════
  if (
    cfg.dataTable.autoFillBatchNumber &&
    cfg.dataTable.autoFillBatchNumber.enabled
  ) {
    let contractCol = cfg.dataTable.autoFillBatchNumber.contractCol;
    let batchCol = cfg.dataTable.autoFillBatchNumber.batchCol;

    let contractRange = sheet.getRange(
      `${contractCol}${startRow}:${contractCol}${lastRow}`,
    );
    let contractValues = contractRange.getValues();
    let batchRange = sheet.getRange(
      `${batchCol}${startRow}:${batchCol}${lastRow}`,
    );
    let batchValues = batchRange.getValues();

    // Map: Base Contract ID => Batch Number (10 digits)
    let batchMap = new Map<string, string>();

    // Hàm lấy Base Contract ID (chỉ lấy phần chữ số/chữ cái đứng trước khoảng trắng hoặc gạch ngang)
    const getBaseContractId = (id: string) => {
      let match = id.match(/^([A-Za-z0-9]+)/);
      return match ? match[1].toUpperCase() : id.toUpperCase();
    };

    // Pass 1: Thu thập số lô hợp lệ
    for (let i = 0; i < batchValues.length; i++) {
      let bVal = String(batchValues[i][0] || "").trim();
      let cVal = String(contractValues[i][0] || "").trim();

      // Số lô hợp lệ là chuỗi 10 chữ số
      if (/^\d{10}$/.test(bVal) && cVal) {
        let baseId = getBaseContractId(cVal);
        if (!batchMap.has(baseId)) {
          batchMap.set(baseId, bVal);
        }
      }
    }

    // Pass 2: Cập nhật số lô bị thiếu hoặc sai
    let updatedBatch = false;
    for (let i = 0; i < batchValues.length; i++) {
      let bVal = String(batchValues[i][0] || "").trim();
      let cVal = String(contractValues[i][0] || "").trim();

      if (cVal) {
        let baseId = getBaseContractId(cVal);
        let validBatch = batchMap.get(baseId);

        // Nếu có số lô hợp lệ nhưng ô hiện tại không phải là số lô 10 chữ số (trống, 0, #NaN...)
        if (
          validBatch &&
          !/^\d{10}$/.test(bVal) &&
          !skipBatchRows.has(startRow + i)
        ) {
          sheet.getRange(`${batchCol}${startRow + i}`).setValue(validBatch);
          // Căn giữa cho đẹp
          sheet
            .getRange(`${batchCol}${startRow + i}`)
            .getFormat()
            .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
          updatedBatch = true;
        }
      }
    }

    if (updatedBatch) {
      console.log(
        `✅ Đã tự động điền Số Lô (cột ${batchCol}) bị thiếu từ Mã Hợp Đồng (cột ${contractCol}).`,
      );
    }
  }

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 13: Chèn công thức tính toán cho các cột (formulaColumns)
  // ═══════════════════════════════════════════════════════════════════
  if (cfg.dataTable.formulaColumns && cfg.dataTable.formulaColumns.length > 0) {
    cfg.dataTable.formulaColumns.forEach((fc) => {
      for (let row = startRow; row <= lastRow; row++) {
        let cell = sheet.getRange(`${fc.target}${row}`);
        cell.setFormula(fc.formula.replace(/\{row\}/g, String(row)));
      }
    });
    console.log(
      `✅ Đã chèn công thức cho các cột: [${cfg.dataTable.formulaColumns.map((fc) => fc.target).join(", ")}].`,
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 14: Dòng tổng cộng (Summary Row)
  // ═══════════════════════════════════════════════════════════════════
  if (summaryColumns.length > 0 && lastRow >= startRow) {
    let summaryCreated = existingSummaryRow === null;

    summaryColumns.forEach((col) => {
      let cell = sheet.getRange(`${col}${summaryRow}`);

      // Công thức tính tổng (cập nhật lại khi chạy script lần 2)
      cell.setFormula(`=SUM(${col}${startRow}:${col}${lastRow})`);

      // Định dạng
      let format = cell.getFormat();
      format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
      format.getFill().setColor("#FFFF00"); // Nền vàng

      // Kẻ khung
      const borderIndices = [
        ExcelScript.BorderIndex.edgeTop,
        ExcelScript.BorderIndex.edgeBottom,
        ExcelScript.BorderIndex.edgeLeft,
        ExcelScript.BorderIndex.edgeRight,
      ];
      borderIndices.forEach((bIndex) => {
        let border = format.getRangeBorder(bIndex);
        border.setStyle(ExcelScript.BorderLineStyle.continuous);
        border.setWeight(ExcelScript.BorderWeight.thin);
        border.setColor("#000000");
      });
    });
    console.log(
      summaryCreated
        ? `✅ Đã thêm dòng tổng cộng tại dòng ${summaryRow} cho các cột: [${summaryColumns.join(", ")}].`
        : `✅ Đã cập nhật dòng tổng cộng tại dòng ${summaryRow} cho các cột: [${summaryColumns.join(", ")}].`,
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 14.5: Điều chỉnh thể tích container (sửa kích thước cột E)
  // ═══════════════════════════════════════════════════════════════════
  if (
    isSaleSheet &&
    cfg.dataTable.autoAdjustContainerVolume &&
    cfg.dataTable.autoAdjustContainerVolume.enabled &&
    lastRow >= startRow
  ) {
    autoAdjustPxkContainerVolume(
      sheet,
      cfg.dataTable.autoAdjustContainerVolume,
      startRow,
      lastRow,
      summaryRow,
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 15: Bù trừ số cân nặng (Weight Balancing)
  // ═══════════════════════════════════════════════════════════════════
  if (
    isSaleSheet &&
    cfg.dataTable.autoBalanceWeight &&
    cfg.dataTable.autoBalanceWeight.enabled &&
    lastRow >= startRow
  ) {
    let productCol = cfg.dataTable.autoBalanceWeight.productCol || "C";
    let dimCol = cfg.dataTable.autoBalanceWeight.dimCol || "E";
    let weightCol = cfg.dataTable.autoBalanceWeight.weightCol;
    let qtyCol = cfg.dataTable.autoBalanceWeight.qtyCol;
    let totalCol = cfg.dataTable.autoBalanceWeight.totalWeightCol;

    // Ô ghi giá trị cân nặng đích (ngay dưới dòng tổng)
    let targetCellRow = summaryRow + 1;
    let targetCell = sheet.getRange(`${totalCol}${targetCellRow}`);
    let targetVal = targetCell.getValue();
    let targetWeight = Number(targetVal);

    if (!isNaN(targetWeight) && targetWeight > 0) {
      let productRange = sheet.getRange(
        `${productCol}${startRow}:${productCol}${lastRow}`,
      );
      let dimRange = sheet.getRange(`${dimCol}${startRow}:${dimCol}${lastRow}`);
      let kRange = sheet.getRange(
        `${weightCol}${startRow}:${weightCol}${lastRow}`,
      );
      let hRange = sheet.getRange(`${qtyCol}${startRow}:${qtyCol}${lastRow}`);

      let productValues = productRange.getValues();
      let dimValues = dimRange.getValues();
      let kValues = kRange.getValues() as number[][];
      let hValues = hRange.getValues() as number[][];

      let currentTotal = 0;

      // Nhóm các sản phẩm theo kích thước
      let dimGroups = new Map<
        string,
        {
          dim: string;
          originalK: number;
          k: number;
          sumH: number;
          indices: number[];
        }
      >();

      for (let i = 0; i < kValues.length; i++) {
        let dim = String(dimValues[i][0] || "").trim();
        // Fallback cho dòng không có kích thước: gán một ID duy nhất để xử lý rời
        if (!dim) dim = `_empty_${i}`;

        let k = Number(kValues[i][0]) || 0;
        let h = Number(hValues[i][0]) || 0;

        if (!dimGroups.has(dim)) {
          dimGroups.set(dim, {
            dim: dim,
            originalK: k,
            k: k,
            sumH: 0,
            indices: [],
          });
        }
        let group = dimGroups.get(dim)!;
        group.sumH += h;
        group.indices.push(i);

        // Tổng Excel = SUM(ROUND(L*I, 2)), dùng k đại diện của nhóm
        currentTotal += Math.round(group.k * h * 100) / 100;
      }
      currentTotal = Math.round(currentTotal * 100) / 100;

      const round2 = (n: number) => Math.round(n * 100) / 100;
      const toCents = (n: number) => Math.round(n * 100);
      const lineH = (i: number) => Number(hValues[i][0]) || 0;
      const lineContrib = (k: number, h: number) => round2(k * h);

      let diff = round2(targetWeight - currentTotal);
      if (Math.abs(diff) >= 0.005) {
        let groupsData = Array.from(dimGroups.values());

        // Phân loại thành 2 tập:
        // 1. Nhóm độc lập (kích thước khác biệt, chỉ có 1 sản phẩm)
        // 2. Nhóm chung (có từ 2 sản phẩm trở lên cùng kích thước)
        let uniqueGroups = groupsData.filter((g) => g.indices.length === 1);
        let sharedGroups = groupsData.filter((g) => g.indices.length > 1);

        // Mỗi tập sắp xếp theo tổng số lượng đóng gói (sumH) giảm dần
        uniqueGroups.sort((a, b) => b.sumH - a.sumH);
        sharedGroups.sort((a, b) => b.sumH - a.sumH);

        // Ghép lại: Ưu tiên nhóm độc lập trước, sau đó mới đến nhóm có cùng kích thước
        groupsData = [...uniqueGroups, ...sharedGroups];

        let maxAdjust = cfg.dataTable.autoBalanceWeight.maxAdjustPerStep;
        let maxIterations = 1000;
        let iter = 0;

        while (Math.abs(diff) > 0.01 && iter < maxIterations) {
          let changed = false;

          for (let group of groupsData) {
            if (group.sumH <= 0) continue;

            if (diff > 0.01) {
              let increaseK = Math.min(diff / group.sumH, maxAdjust);
              if (increaseK > 0) {
                group.k += increaseK;
                diff -= increaseK * group.sumH;
                changed = true;
              }
            } else if (diff < -0.01) {
              let decreaseK = Math.min(Math.abs(diff) / group.sumH, maxAdjust);
              // Đảm bảo K không bị giảm xuống dưới 0.1
              if (group.k - decreaseK < 0.1) {
                decreaseK = Math.max(0, group.k - 0.1);
              }
              if (decreaseK > 0) {
                group.k -= decreaseK;
                diff += decreaseK * group.sumH;
                changed = true;
              }
            }

            if (Math.abs(diff) <= 0.01) break;
          }

          if (!changed) break;
          iter++;
        }

        // ── Làm tròn group.k 2 chữ số thập phân ──
        for (let group of groupsData) {
          group.k = round2(group.k);
        }

        let lineK: number[] = kValues.map(() => 0);
        for (let g of groupsData) {
          for (let idx of g.indices) {
            lineK[idx] = g.k;
          }
        }

        const calcSimulatedTotal = (): number => {
          let total = 0;
          for (let i = 0; i < lineK.length; i++) {
            total += lineContrib(lineK[i], lineH(i));
          }
          return round2(total);
        };

        const targetCents = toCents(targetWeight);

        // ── Bù sai lệch do làm tròn: chỉnh cả nhóm trước, rồi từng dòng ±0.01 ──
        for (let adj = 0; adj < 200; adj++) {
          let rem = targetCents - toCents(calcSimulatedTotal());
          if (rem === 0) break;

          let step = rem > 0 ? 0.01 : -0.01;
          let bestGroup: (typeof groupsData)[0] | null = null;
          let bestError = Math.abs(rem);

          for (let g of groupsData) {
            if (step < 0 && g.k <= 0.11) continue;
            let saved = g.k;
            g.k = round2(g.k + step);
            for (let idx of g.indices) lineK[idx] = g.k;
            let testError = Math.abs(targetCents - toCents(calcSimulatedTotal()));
            g.k = saved;
            for (let idx of g.indices) lineK[idx] = saved;
            if (testError < bestError) {
              bestError = testError;
              bestGroup = g;
            }
          }

          if (!bestGroup) break;
          bestGroup.k = round2(bestGroup.k + step);
          for (let idx of bestGroup.indices) lineK[idx] = bestGroup.k;
        }

        // Nếu vẫn lệch (thường ±0.01 / ±0.02): chỉnh từng dòng, ưu tiên số lượng nhỏ
        let lineOrder: number[] = [];
        for (let i = 0; i < lineK.length; i++) lineOrder.push(i);
        lineOrder.sort((a, b) => lineH(a) - lineH(b));

        for (let adj = 0; adj < 500; adj++) {
          let rem = targetCents - toCents(calcSimulatedTotal());
          if (rem === 0) break;

          let step = rem > 0 ? 0.01 : -0.01;
          let bestI = -1;
          let bestError = Math.abs(rem);

          for (let i of lineOrder) {
            if (lineH(i) <= 0) continue;
            if (step < 0 && lineK[i] <= 0.11) continue;
            let saved = lineK[i];
            lineK[i] = round2(saved + step);
            let testError = Math.abs(targetCents - toCents(calcSimulatedTotal()));
            lineK[i] = saved;
            if (testError < bestError) {
              bestError = testError;
              bestI = i;
              if (testError === 0) break;
            }
          }

          if (bestI < 0) break;
          lineK[bestI] = round2(lineK[bestI] + step);
        }

        // Tìm tổ hợp 1–2 dòng (bước ±0.01..±0.20) nếu greedy vẫn lệch
        if (targetCents !== toCents(calcSimulatedTotal())) {
          let baseTotal = calcSimulatedTotal();
          let candidates = lineOrder.filter((i) => lineH(i) > 0).slice(0, 60);
          let found = false;

          const tryDeltas = (
            i: number,
            d1: number,
            j: number | null,
            d2: number,
          ): boolean => {
            let t = baseTotal;
            let k1 = round2(lineK[i] + d1 * 0.01);
            if (k1 < 0.1) return false;
            t = round2(t - lineContrib(lineK[i], lineH(i)) + lineContrib(k1, lineH(i)));
            if (j !== null) {
              let k2 = round2(lineK[j] + d2 * 0.01);
              if (k2 < 0.1) return false;
              t = round2(
                t - lineContrib(lineK[j], lineH(j)) + lineContrib(k2, lineH(j)),
              );
            }
            if (toCents(t) !== targetCents) return false;
            lineK[i] = k1;
            if (j !== null) lineK[j] = round2(lineK[j] + d2 * 0.01);
            return true;
          };

          outer: for (let i of candidates) {
            for (let d1 = -20; d1 <= 20; d1++) {
              if (d1 === 0) continue;
              if (tryDeltas(i, d1, null, 0)) {
                found = true;
                break outer;
              }
            }
          }

          if (!found) {
            outer2: for (let a = 0; a < candidates.length; a++) {
              let i = candidates[a];
              for (let d1 = -8; d1 <= 8; d1++) {
                if (d1 === 0) continue;
                for (let b = a + 1; b < candidates.length; b++) {
                  let j = candidates[b];
                  for (let d2 = -8; d2 <= 8; d2++) {
                    if (d2 === 0) continue;
                    if (tryDeltas(i, d1, j, d2)) {
                      found = true;
                      break outer2;
                    }
                  }
                }
              }
            }
          }
        }

        diff = round2(targetWeight - calcSimulatedTotal());

        let logMessages: string[] = [];
        groupsData.forEach((group) => {
          let change = group.k - group.originalK;
          if (Math.abs(change) > 0.001) {
            let firstIndex = group.indices[0];
            let productCode = String(productValues[firstIndex][0] || "").trim();
            let action = change > 0 ? "Tăng" : "Giảm";
            logMessages.push(
              `- Mã ${productCode} (Kích thước: ${group.dim}): ${action} ${Math.abs(change).toFixed(3)} kg/đơn vị.`,
            );
          }
        });

        for (let i = 0; i < lineK.length; i++) {
          kValues[i][0] = round2(lineK[i]);
        }
        kRange.setValues(kValues);
        console.log(
          `✅ Đã bù trừ cân nặng theo nhóm Kích thước. Cân nặng đích: ${targetWeight}, Sai số còn lại: ${diff}`,
        );
        if (logMessages.length > 0) {
          console.log(
            `📝 Chi tiết điều chỉnh cân nặng:\n${logMessages.join("\n")}`,
          );
        }
      }
    }
    // 4. Định dạng và ẩn hàng chứa cân nặng đích
    targetCellRow = summaryRow + 1;
    targetCell = sheet.getRange(`${totalCol}${targetCellRow}`);
    let targetFormat = targetCell.getFormat();

    // Nền vàng + Căn giữa
    targetFormat.getFill().setColor("#FFFF00");
    targetFormat.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    targetFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

    // Kẻ khung
    // const borderIndices = [
    //   ExcelScript.BorderIndex.edgeTop,
    //   ExcelScript.BorderIndex.edgeBottom,
    //   ExcelScript.BorderIndex.edgeLeft,
    //   ExcelScript.BorderIndex.edgeRight,
    // ];
    // borderIndices.forEach((bIndex) => {
    //   let border = targetFormat.getRangeBorder(bIndex);
    //   border.setStyle(ExcelScript.BorderLineStyle.continuous);
    //   border.setWeight(ExcelScript.BorderWeight.thin);
    //   border.setColor("#000000");
    // });

    // Ẩn hàng
    sheet.getRange(`${targetCellRow}:${targetCellRow}`).setRowHidden(true);
    console.log(
      `✅ Đã định dạng và ẩn hàng ${targetCellRow} (hàng chứa giá trị cân nặng đích).`,
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 16: Định dạng số rút gọn (Bỏ số 0 thừa)
  // ═══════════════════════════════════════════════════════════════════
  if (
    cfg.dataTable.numberFormatColumns &&
    cfg.dataTable.numberFormatColumns.length > 0
  ) {
    cfg.dataTable.numberFormatColumns.forEach((col) => {
      let endFormatRow = Math.max(lastRow, summaryRow);
      let fullRange = sheet.getRange(`${col}${startRow}:${col}${endFormatRow}`);
      let twoDecCols = cfg.dataTable.fixedTwoDecimalColumns || [];
      if (twoDecCols.indexOf(col) !== -1) {
        fullRange.setNumberFormat("0.00");
      } else {
        fullRange.setNumberFormat("General");
      }

      // 2. Xử lý triệt để: Làm tròn giá trị thực tế của các ô chứa dữ liệu cứng
      let dataRange = sheet.getRange(`${col}${startRow}:${col}${lastRow}`);
      let formulas = dataRange.getFormulas();
      let values = dataRange.getValues();
      let changed = false;

      for (let i = 0; i < values.length; i++) {
        let val = values[i][0];
        let fml = String(formulas[i][0] || "").trim();

        // Nếu là một số và không phải là công thức (không bắt đầu bằng "=")
        if (typeof val === "number" && !fml.startsWith("=")) {
          // Cột M (Thể tích) làm tròn 3 số, các cột khác làm tròn 2 số
          let decimals = col === "N" ? 3 : 2;
          let factor = Math.pow(10, decimals);
          let rounded = Math.round(val * factor) / factor;

          if (val !== rounded) {
            values[i][0] = rounded;
            changed = true;
          }
        }
      }

      // Nạp lại dữ liệu cứng đã được làm tròn
      if (changed) {
        dataRange.setValues(values);
      }
    });
    console.log(
      `✅ Đã định dạng số và làm tròn thực tế cho các cột: [${cfg.dataTable.numberFormatColumns.join(", ")}].`,
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // BƯỚC 17: Copy footer từ sheet "footer" (nếu có)
  // ═══════════════════════════════════════════════════════════════════
  let footerSheet = workbook.getWorksheet("footer");
  if (footerSheet) {
    let pasteRow = summaryRow + 3;

    if (existingFooterRow !== null) {
      console.log(
        `ℹ️ Footer đã có tại dòng ${existingFooterRow}, bỏ qua chèn footer trùng.`,
      );
    } else {
      // Lấy vùng đã dùng của hàng 1 sheet footer
      let fullUsedRange = footerSheet.getUsedRange();
      let footerRowRange = fullUsedRange
        ? fullUsedRange.getIntersection("1:1")
        : null;

      if (footerRowRange) {
        let targetRange = sheet
          .getCell(pasteRow - 1, 0)
          .getResizedRange(0, footerRowRange.getColumnCount() - 1);

        // Copy cả giá trị và định dạng để giữ nguyên style của footer
        targetRange.copyFrom(footerRowRange, ExcelScript.RangeCopyType.all);

        console.log(
          `✅ Đã copy footer từ sheet "footer" vào dòng ${pasteRow} (cách 2 dòng từ dòng tổng).`,
        );
      }
    }
  } else {
    console.log(`ℹ️ Không tìm thấy sheet "footer", bỏ qua bước chèn footer.`);
  }

  // ═══════════════════════════════════════════════════════════════════
  // HOÀN TẤT
  // ═══════════════════════════════════════════════════════════════════
  console.log(
    `🎉 Hoàn tất định dạng Phiếu Xuất Kho! (${lastRow - startRow + 1} hàng dữ liệu)`,
  );
}
