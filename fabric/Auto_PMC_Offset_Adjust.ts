/**
 * ===================================================================
 * Auto_PMC_Offset_Adjust – Thiết kế form & tính bù trừ số lượng cấp
 * giữa Kho và PMC (ERP) trên sheet "PMC_ADJUST".
 * ===================================================================
 *
 * - Quy trình 2 bước trên sheet "PMC_ADJUST":
 *   + Bước 1 (Template mode): Nếu chưa có dữ liệu ở cột B/D, tạo form
 *     mẫu sẵn với định dạng mã hàng "0" (tránh E+11) và số lượng giữ 4 số lẻ.
 *   + Bước 2 (Calculate mode): Đọc dữ liệu Kho (cột B:C) và ERP/PMC (cột D:E).
 *     Với mỗi mã hàng bên ERP/PMC có khớp bên Kho, tính chênh lệch
 *     (Kho - Tổng PMC) và bù trừ vào dòng dữ liệu cuối cùng của mã đó ở:
 *     * Cột F: Số lượng PMC sau điều chỉnh
 *     * Cột G: Chênh lệch bù trừ (Adjusted - Original)
 *     * Cột H: Ghi chú ("Đã bù trừ", "Khớp", "Không tìm thấy mã bên Kho")
 */

type PmcCellValue = string | number | boolean;

type PmcKhoEntry = {
  itemCode: string;
  qty: number;
};

type PmcErpEntry = {
  itemCode: string;
  qty: number;
};

type PmcAdjustedRow = {
  itemCode: string;
  originalQty: number;
  adjustedQty: number;
  offsetDiff: number;
  note: string;
  isLastOfItem: boolean;
  matchedKho: boolean;
  isShortageWarning: boolean;
};

type PmcParsedInput = {
  khoEntries: PmcKhoEntry[];
  erpEntries: PmcErpEntry[];
};

const PMC_OFFSET_CONFIG = {
  sheetName: "PMC_ADJUST",
  headerRow: 5,
  startDataRow: 6,
  templateInputRowCount: 15,
};

const PMC_OFFSET_HEADERS = [
  "STT",
  "Mã hàng hóa (Kho)",
  "Số lượng cấp của kho",
  "Mã hàng hóa từ ERP",
  "Số lượng cấp của PMC",
  "Số lượng PMC sau điều chỉnh",
  "Chênh lệch bù trừ",
  "Ghi chú",
];

const PMC_STOP_MARKERS = new Set([
  "TỔNG CỘNG",
  "TONG CONG",
  "TOTAL",
]);

const pmcCleanText = (value: unknown): string => {
  if (typeof value === "string") {
    if (value === "") return "";
    return value
      .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }
  if (value === null || value === undefined) return "";
  return String(value)
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
};

const pmcExpandScientific = (input: string): string | null => {
  const match = input.match(/^([+-]?)(\d+)(?:[.,](\d+))?[eE]([+-]?\d+)$/);
  if (!match) return null;
  const sign = match[1] === "-" ? "-" : "";
  const intPart = match[2];
  const fracPart = match[3] ?? "";
  const exp = Number.parseInt(match[4], 10);
  if (!Number.isFinite(exp) || Math.abs(exp) > 100) return null;

  const digits = `${intPart}${fracPart}`;
  const decimalPos = intPart.length + exp;

  if (decimalPos <= 0) {
    const trimmed = `${"0".repeat(-decimalPos)}${digits}`.replace(/0+$/, "");
    return trimmed ? `${sign}0.${trimmed}` : "0";
  }
  if (decimalPos >= digits.length) {
    const whole = `${digits}${"0".repeat(decimalPos - digits.length)}`.replace(
      /^0+(?=\d)/,
      "",
    );
    return whole === "0" ? "0" : `${sign}${whole}`;
  }
  const whole = digits.slice(0, decimalPos).replace(/^0+(?=\d)/, "");
  const frac = digits.slice(decimalPos).replace(/0+$/, "");
  if (!frac) {
    return whole === "0" ? "0" : `${sign}${whole}`;
  }
  return `${sign}${whole}.${frac}`;
};

const pmcNormalizeItemCode = (value: unknown): string => {
  if (value === null || value === undefined || typeof value === "boolean") {
    return "";
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    if (Number.isSafeInteger(value)) return String(value);
    const str = String(Math.trunc(value));
    if (str.indexOf("e") !== -1 || str.indexOf("E") !== -1) {
      return pmcExpandScientific(str) ?? str;
    }
    return str;
  }
  if (typeof value === "string" && value === "") return "";
  const cleaned = String(value).replace(/[\u00A0\u200B\u200C\u200D\uFEFF\s]+/g, "");
  if (!cleaned) return "";
  if (cleaned.indexOf("e") !== -1 || cleaned.indexOf("E") !== -1) {
    const expanded = pmcExpandScientific(cleaned);
    if (expanded !== null) {
      return expanded;
    }
  }
  if (cleaned.indexOf(".") !== -1 && /^\d+\.0+$/.test(cleaned)) {
    return cleaned.replace(/\.0+$/, "");
  }
  return cleaned;
};

const pmcToCellItemCode = (itemCode: string): string | number => {
  const normalized = pmcNormalizeItemCode(itemCode);
  if (!normalized) return "";
  if (/^\d+$/.test(normalized) && !/^0\d+$/.test(normalized)) {
    const num = Number(normalized);
    if (Number.isSafeInteger(num)) {
      return num;
    }
  }
  return normalized;
};

const pmcParseNumber = (value: unknown): number => {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  if (typeof value === "boolean" || value === null || value === undefined) {
    return 0;
  }
  let s = pmcCleanText(value).replace(/\s+/g, "");
  if (!s) return 0;

  const lastComma = s.lastIndexOf(",");
  const lastDot = s.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    if (lastComma > lastDot) {
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      s = s.replace(/,/g, "");
    }
  } else if (lastComma >= 0) {
    if (/^-?\d{1,3}(,\d{3})+$/.test(s)) {
      s = s.replace(/,/g, "");
    } else {
      s = s.replace(",", ".");
    }
  } else if (lastDot >= 0 && /^-?\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, "");
  }

  const parsed = Number(s);
  return Number.isFinite(parsed) ? parsed : 0;
};

const pmcRoundQty = (n: number): number => Math.round(n * 1e6) / 1e6;

const pmcPad2 = (n: number): string => (n < 10 ? `0${n}` : String(n));

const pmcFormatTimestamp = (offsetHours: number = 7): string => {
  const now = new Date(Date.now() + offsetHours * 3600000);
  return `${pmcPad2(now.getUTCDate())}/${pmcPad2(now.getUTCMonth() + 1)}/${now.getUTCFullYear()} ${pmcPad2(now.getUTCHours())}:${pmcPad2(now.getUTCMinutes())}`;
};

/**
 * Đọc dữ liệu nhập từ bảng A:E (cột 0=A STT, 1=B Mã Kho, 2=C SL Kho, 3=D Mã ERP, 4=E SL PMC).
 * Dừng khi gặp dòng TỔNG CỘNG.
 */
const pmcExtractInputEntries = (rowsAE: unknown[][]): PmcParsedInput => {
  const khoEntries: PmcKhoEntry[] = [];
  const erpEntries: PmcErpEntry[] = [];

  for (let i = 0; i < rowsAE.length; i++) {
    const row = rowsAE[i];
    if (!row) continue;

    const colBUpper = pmcCleanText(row[1]).toUpperCase();
    const colDUpper = pmcCleanText(row[3]).toUpperCase();
    if (PMC_STOP_MARKERS.has(colBUpper) || PMC_STOP_MARKERS.has(colDUpper)) {
      break;
    }

    const khoCode = pmcNormalizeItemCode(row[1]);
    if (khoCode) {
      khoEntries.push({
        itemCode: khoCode,
        qty: pmcRoundQty(pmcParseNumber(row[2])),
      });
    }

    const erpCode = pmcNormalizeItemCode(row[3]);
    if (erpCode) {
      erpEntries.push({
        itemCode: erpCode,
        qty: pmcRoundQty(pmcParseNumber(row[4])),
      });
    }
  }

  return { khoEntries, erpEntries };
};

/**
 * Tính số lượng PMC sau điều chỉnh và chênh lệch bù trừ cho từng dòng ERP.
 * Chênh lệch (Kho - Tổng PMC) được bù trừ toàn bộ vào dòng cuối cùng của mã hàng đó.
 */
const pmcComputeOffsetAdjustments = (
  khoEntries: PmcKhoEntry[],
  erpEntries: PmcErpEntry[],
): PmcAdjustedRow[] => {
  const khoMap = new Map<string, number>();
  for (let i = 0; i < khoEntries.length; i++) {
    const entry = khoEntries[i];
    const prev = khoMap.get(entry.itemCode) ?? 0;
    khoMap.set(entry.itemCode, pmcRoundQty(prev + entry.qty));
  }

  const erpSumMap = new Map<string, number>();
  const lastIndexMap = new Map<string, number>();
  for (let i = 0; i < erpEntries.length; i++) {
    const entry = erpEntries[i];
    const prev = erpSumMap.get(entry.itemCode) ?? 0;
    erpSumMap.set(entry.itemCode, pmcRoundQty(prev + entry.qty));
    lastIndexMap.set(entry.itemCode, i);
  }

  const results: PmcAdjustedRow[] = [];
  for (let i = 0; i < erpEntries.length; i++) {
    const entry = erpEntries[i];
    const isLast = lastIndexMap.get(entry.itemCode) === i;
    const matchedKho = khoMap.has(entry.itemCode);

    if (!matchedKho) {
      results.push({
        itemCode: entry.itemCode,
        originalQty: entry.qty,
        adjustedQty: entry.qty,
        offsetDiff: 0,
        note: "Không tìm thấy mã bên Kho",
        isLastOfItem: isLast,
        matchedKho: false,
        isShortageWarning: false,
      });
      continue;
    }

    const khoQty = khoMap.get(entry.itemCode)!;
    const pmcSum = erpSumMap.get(entry.itemCode) ?? 0;
    const diff = pmcRoundQty(khoQty - pmcSum);

    if (diff < 0) {
      results.push({
        itemCode: entry.itemCode,
        originalQty: entry.qty,
        adjustedQty: entry.qty,
        offsetDiff: 0,
        note: isLast
          ? `Cảnh báo: Kho ít hơn PMC (thiếu ${pmcRoundQty(Math.abs(diff))}) - Không bù`
          : "",
        isLastOfItem: isLast,
        matchedKho: true,
        isShortageWarning: true,
      });
      continue;
    }

    if (!isLast) {
      results.push({
        itemCode: entry.itemCode,
        originalQty: entry.qty,
        adjustedQty: entry.qty,
        offsetDiff: 0,
        note: "",
        isLastOfItem: false,
        matchedKho: true,
        isShortageWarning: false,
      });
      continue;
    }

    const adjustedQty = pmcRoundQty(entry.qty + diff);

    results.push({
      itemCode: entry.itemCode,
      originalQty: entry.qty,
      adjustedQty,
      offsetDiff: diff,
      note: diff === 0 ? "Khớp" : "Đã bù trừ",
      isLastOfItem: true,
      matchedKho: true,
      isShortageWarning: false,
    });
  }

  return results;
};

const pmcBuildEmptyTemplateRows = (
  rowCount: number,
  _startDataRow: number = 0,
): PmcCellValue[][] => {
  const safeCount = Math.max(1, rowCount);
  const rows: PmcCellValue[][] = [];
  for (let i = 0; i < safeCount; i++) {
    rows.push([i + 1, "", "", "", "", "", "", ""]);
  }
  return rows;
};

const pmcBuildCalculatedRows = (
  khoEntries: PmcKhoEntry[],
  adjustedRows: PmcAdjustedRow[],
  startDataRow: number,
): PmcCellValue[][] => {
  const rowCount = Math.max(khoEntries.length, adjustedRows.length);
  const rows: PmcCellValue[][] = [];

  for (let i = 0; i < rowCount; i++) {
    const kho = khoEntries[i];
    const adj = adjustedRows[i];
    rows.push([
      i + 1,
      kho ? pmcToCellItemCode(kho.itemCode) : "",
      kho ? kho.qty : "",
      adj ? pmcToCellItemCode(adj.itemCode) : "",
      adj ? adj.originalQty : "",
      adj ? adj.adjustedQty : "",
      adj ? adj.offsetDiff : "",
      adj ? adj.note : "",
    ]);
  }

  if (rowCount > 0) {
    const endDataRow = startDataRow + rowCount - 1;
    rows.push([
      "",
      "TỔNG CỘNG",
      `=SUM(C${startDataRow}:C${endDataRow})`,
      "",
      `=SUM(E${startDataRow}:E${endDataRow})`,
      `=SUM(F${startDataRow}:F${endDataRow})`,
      `=SUM(G${startDataRow}:G${endDataRow})`,
      "",
    ]);
  }

  return rows;
};

const pmcGetLastDataRow = (sheet: ExcelScript.Worksheet): number => {
  const used = sheet.getUsedRange(true);
  if (!used) return 0;
  return used.getLastRow().getRowIndex() + 1;
};

const PMC_LARGE_TABLE_BORDER_ROW_LIMIT = 1000;

const pmcApplyThinBorders = (
  range: ExcelScript.Range,
  rowCount: number = 0,
): void => {
  const format = range.getFormat();
  const edgeIds = [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
  ];
  if (rowCount <= PMC_LARGE_TABLE_BORDER_ROW_LIMIT) {
    edgeIds.push(
      ExcelScript.BorderIndex.insideHorizontal,
      ExcelScript.BorderIndex.insideVertical,
    );
  }
  for (let i = 0; i < edgeIds.length; i++) {
    const border = format.getRangeBorder(edgeIds[i]);
    border.setStyle(ExcelScript.BorderLineStyle.continuous);
    border.setWeight(ExcelScript.BorderWeight.thin);
    border.setColor("#BFBFBF");
  }
};

// @ts-ignore: Office Scripts
async function main(workbook: ExcelScript.Workbook) {
  const CFG = PMC_OFFSET_CONFIG;
  const headerRow = CFG.headerRow;
  const startDataRow = CFG.startDataRow;

  let sheet = workbook.getWorksheet(CFG.sheetName);
  let parsed: PmcParsedInput = { khoEntries: [], erpEntries: [] };

  if (sheet) {
    const lastRow = pmcGetLastDataRow(sheet);
    if (lastRow >= startDataRow) {
      const rawAE = sheet.getRange(`A${startDataRow}:E${lastRow}`).getValues();
      parsed = pmcExtractInputEntries(rawAE);
    }
  } else {
    sheet = workbook.addWorksheet(CFG.sheetName);
  }

  const isEmptyTemplateMode =
    parsed.khoEntries.length === 0 && parsed.erpEntries.length === 0;

  const used = sheet.getUsedRange(true);
  if (used) {
    used.unmerge();
    used.clear(ExcelScript.ClearApplyTo.all);
  }
  const autoFilter = sheet.getAutoFilter();
  if (autoFilter.getEnabled()) autoFilter.remove();
  sheet.getFreezePanes().unfreeze();

  let adjustedRows: PmcAdjustedRow[] = [];
  let dataRows: PmcCellValue[][] = [];
  let dataCount = 0;

  if (isEmptyTemplateMode) {
    dataCount = CFG.templateInputRowCount;
    dataRows = pmcBuildEmptyTemplateRows(dataCount, startDataRow);
  } else {
    adjustedRows = pmcComputeOffsetAdjustments(
      parsed.khoEntries,
      parsed.erpEntries,
    );
    dataCount = Math.max(parsed.khoEntries.length, adjustedRows.length);
    dataRows = pmcBuildCalculatedRows(
      parsed.khoEntries,
      adjustedRows,
      startDataRow,
    );
  }

  const dataEndRow = startDataRow + dataCount - 1;
  const totalRow = startDataRow + dataCount;
  const finalBottomRow = isEmptyTemplateMode ? dataEndRow : totalRow;

  const adjustedCount = adjustedRows.filter(
    (r) => r.isLastOfItem && r.matchedKho && !r.isShortageWarning && r.offsetDiff !== 0,
  ).length;
  const unmatchedCount = adjustedRows.filter((r) => !r.matchedKho).length;
  const shortageItemSet = new Set<string>();
  for (let i = 0; i < adjustedRows.length; i++) {
    if (adjustedRows[i].isShortageWarning) {
      shortageItemSet.add(adjustedRows[i].itemCode);
    }
  }
  const shortageCount = shortageItemSet.size;

  const subtitleText = isEmptyTemplateMode
    ? `Hướng dẫn: Nhập Mã hàng & SL Kho (cột B–C) và Mã hàng ERP & SL PMC (cột D–E) từ dòng ${startDataRow} trở đi (không giới hạn số dòng), rồi chạy lại script để tính bù trừ.`
    : `Thời gian: ${pmcFormatTimestamp(7)}  |  Mã Kho: ${parsed.khoEntries.length}  |  Dòng ERP/PMC: ${parsed.erpEntries.length}  |  Đã bù trừ: ${adjustedCount} mã${shortageCount > 0 ? `  |  Cảnh báo Kho < PMC: ${shortageCount} mã (không bù)` : ""}${unmatchedCount > 0 ? `  |  Chưa khớp Kho: ${unmatchedCount} dòng` : ""}`;

  const emptyRow8 = (): PmcCellValue[] => ["", "", "", "", "", "", "", ""];

  const sheetMatrix: PmcCellValue[][] = [
    ["BẢNG ĐIỀU CHỈNH BÙ TRỪ SỐ LƯỢNG CẤP KHO - PMC", "", "", "", "", "", "", ""],
    [subtitleText, "", "", "", "", "", "", ""],
    emptyRow8(),
    ["BẢNG DỮ LIỆU ĐỐI CHIẾU & BÙ TRỪ", "", "", "", "", "", "", ""],
    [...PMC_OFFSET_HEADERS],
    ...dataRows,
  ];

  // Apply integer format "0" on item code columns before setValues so long codes
  // like 110703000003 stay exact without scientific notation or text warnings.
  // In template mode, format the entire column from startDataRow downward (no row limit).
  const codeFormatEndRow = isEmptyTemplateMode ? 1048576 : dataEndRow;
  sheet.getRange(`B${startDataRow}:B${codeFormatEndRow}`).setNumberFormat("0");
  sheet.getRange(`D${startDataRow}:D${codeFormatEndRow}`).setNumberFormat("0");

  const fullRange = sheet.getRange(`A1:H${finalBottomRow}`);
  fullRange.setValues(sheetMatrix);
  fullRange.getFormat().getFont().setName("Times New Roman");
  fullRange.getFormat().getFont().setSize(11);
  fullRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);

  const titleRange = sheet.getRange("A1:H1");
  titleRange.merge(false);
  titleRange.getFormat().getFont().setSize(15);
  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setColor("#1F4E78");
  titleRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  titleRange.getFormat().setRowHeight(28);

  const subRange = sheet.getRange("A2:H2");
  subRange.merge(false);
  subRange.getFormat().getFont().setSize(10);
  subRange.getFormat().getFont().setItalic(true);
  if (isEmptyTemplateMode) {
    subRange.getFormat().getFont().setColor("#9C6500");
  }
  subRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  const secRange = sheet.getRange("A4:H4");
  secRange.merge(false);
  secRange.getFormat().getFont().setSize(12);
  secRange.getFormat().getFont().setBold(true);
  secRange.getFormat().getFont().setColor("#1F4E78");

  const tableRange = sheet.getRange(`A${headerRow}:H${finalBottomRow}`);
  tableRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  pmcApplyThinBorders(tableRange, dataRows.length + 1);

  const headerRange = sheet.getRange(`A${headerRow}:H${headerRow}`);
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().getFont().setColor("#FFFFFF");
  headerRange.getFormat().getFill().setColor("#1F4E78");
  headerRange.getFormat().setWrapText(true);
  headerRange.getFormat().setRowHeight(30);

  const qtyFormatEndRow = isEmptyTemplateMode ? 1048576 : totalRow;
  sheet.getRange(`C${startDataRow}:C${qtyFormatEndRow}`).setNumberFormat("#,##0.####");
  sheet
    .getRange(`C${startDataRow}:C${qtyFormatEndRow}`)
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

  const pmcQtyRange = sheet.getRange(`E${startDataRow}:G${qtyFormatEndRow}`);
  pmcQtyRange.setNumberFormat("#,##0.####");
  pmcQtyRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

  if (!isEmptyTemplateMode) {
    const totalRowRange = sheet.getRange(`A${totalRow}:H${totalRow}`);
    totalRowRange.getFormat().getFont().setBold(true);
    totalRowRange.getFormat().getFill().setColor("#D9E1F2");
  }

  if (!isEmptyTemplateMode && adjustedRows.length > 0) {
    const adjColRange = sheet.getRange(
      `F${startDataRow}:F${startDataRow + adjustedRows.length - 1}`,
    );
    adjColRange.getFormat().getFill().setColor("#E2EFDA");

    // Batch contiguous red warning highlights for ERP rows where Kho < PMC
    let erpWarnStart = -1;
    for (let i = 0; i <= adjustedRows.length; i++) {
      const isShortage = i < adjustedRows.length && adjustedRows[i].isShortageWarning;
      if (isShortage && erpWarnStart === -1) {
        erpWarnStart = startDataRow + i;
      } else if (!isShortage && erpWarnStart !== -1) {
        const erpWarnEnd = startDataRow + i - 1;
        const warnRange = sheet.getRange(`D${erpWarnStart}:H${erpWarnEnd}`);
        warnRange.getFormat().getFill().setColor("#FFC7CE");
        warnRange.getFormat().getFont().setColor("#9C0006");
        warnRange.getFormat().getFont().setBold(true);
        erpWarnStart = -1;
      }
    }

    // Batch contiguous red warning highlights for matching Kho cells (B:C) where Kho < PMC
    if (shortageItemSet.size > 0) {
      let khoWarnStart = -1;
      for (let i = 0; i <= parsed.khoEntries.length; i++) {
        const isShortageKho =
          i < parsed.khoEntries.length &&
          shortageItemSet.has(parsed.khoEntries[i].itemCode);
        if (isShortageKho && khoWarnStart === -1) {
          khoWarnStart = startDataRow + i;
        } else if (!isShortageKho && khoWarnStart !== -1) {
          const khoWarnEnd = startDataRow + i - 1;
          const khoWarnRange = sheet.getRange(`B${khoWarnStart}:C${khoWarnEnd}`);
          khoWarnRange.getFormat().getFill().setColor("#FFC7CE");
          khoWarnRange.getFormat().getFont().setColor("#9C0006");
          khoWarnRange.getFormat().getFont().setBold(true);
          khoWarnStart = -1;
        }
      }
    }

    for (let i = 0; i < adjustedRows.length; i++) {
      const r = adjustedRows[i];
      const rowNum = startDataRow + i;
      if (!r.matchedKho) {
        const warnRange = sheet.getRange(`G${rowNum}:H${rowNum}`);
        warnRange.getFormat().getFill().setColor("#FFC7CE");
        warnRange.getFormat().getFont().setColor("#9C0006");
      } else if (!r.isShortageWarning && r.isLastOfItem && r.offsetDiff !== 0) {
        const diffRange = sheet.getRange(`F${rowNum}:H${rowNum}`);
        diffRange.getFormat().getFont().setBold(true);
        diffRange.getFormat().getFill().setColor("#FFF2CC");
        diffRange.getFormat().getFont().setColor("#9C6500");
      }
    }
  }

  sheet.getRange(`A${headerRow}:H${finalBottomRow}`).getFormat().autofitColumns();
  sheet.activate();

  if (isEmptyTemplateMode) {
    console.log(
      `Auto_PMC_Offset_Adjust: Đã tạo form mẫu trống trên "${CFG.sheetName}". Nhập dữ liệu từ dòng ${startDataRow} trở đi (không giới hạn số dòng) rồi chạy lại script.`,
    );
  } else {
    console.log(
      `Auto_PMC_Offset_Adjust hoàn tất: ${parsed.khoEntries.length} mã Kho, ${parsed.erpEntries.length} dòng ERP/PMC, đã bù trừ ${adjustedCount} mã${shortageCount > 0 ? `, cảnh báo Kho < PMC (không bù): ${shortageCount} mã` : ""}.`,
    );
  }
}

const pmcOffsetAdjustTestables = {
  PMC_OFFSET_CONFIG,
  PMC_OFFSET_HEADERS,
  pmcCleanText,
  pmcNormalizeItemCode,
  pmcToCellItemCode,
  pmcParseNumber,
  pmcRoundQty,
  pmcExtractInputEntries,
  pmcComputeOffsetAdjustments,
  pmcBuildEmptyTemplateRows,
  pmcBuildCalculatedRows,
  pmcGetLastDataRow,
};

(
  globalThis as typeof globalThis & {
    __AUTO_PMC_OFFSET_ADJUST_TESTABLES__?: typeof pmcOffsetAdjustTestables;
  }
).__AUTO_PMC_OFFSET_ADJUST_TESTABLES__ = pmcOffsetAdjustTestables;
