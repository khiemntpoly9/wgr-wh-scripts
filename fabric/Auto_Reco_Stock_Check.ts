/**
 * ===================================================================
 * Auto_Reco_Stock_Check – Kiểm tra số lượng tồn còn lại theo mã vật tư
 * từ sheet "RECO_DATA" ra sheet "RECO_STOCK_CHECK".
 * ===================================================================
 *
 * - Nguồn dữ liệu: sheet "RECO_DATA", dữ liệu bắt đầu từ dòng 8 (index 7).
 *   + Cột C: Ngày tháng năm (Date)
 *   + Cột D: Loại giao dịch ("IN" / "OUT")
 *   + Cột F: Mã vật tư (Item Code)
 *   + Cột I: Số lượng (Quantity in meters)
 * - Quy trình 2 bước trên sheet "RECO_STOCK_CHECK":
 *   + Lần chạy đầu (hoặc khi chưa nhập mã vật tư ở cột B): tạo bảng mẫu trống
 *     (định dạng cột B là "0" để tránh cảnh báo Number Stored as Text và
 *     không bị đổi sang ký hiệu khoa học E+11).
 *   + Các lần chạy sau: đọc danh sách mã vật tư người dùng đã nhập ở cột B trên
 *     sheet "RECO_STOCK_CHECK", lọc dữ liệu từ "RECO_DATA" và xuất bảng kiểm tra
 *     tồn còn lại (Tổng Nhập IN, Tổng Xuất OUT, Tồn còn lại, Trạng thái tồn,
 *     Ngày nhập cuối, Ngày xuất cuối) cùng dòng TỔNG CỘNG.
 */

type RecoStockCellValue = string | number | boolean;
type RecoStockTxnType = "IN" | "OUT";
type RecoStockStatus = "Còn hàng" | "Hết hàng" | "Âm kho";

type RecoStockTransaction = {
  rowNumber: number;
  dateIso: string | null;
  type: RecoStockTxnType;
  itemCode: string;
  qtyMeters: number;
};

type RecoStockItemSummary = {
  itemCode: string;
  totalIn: number;
  totalOut: number;
  balance: number;
  status: RecoStockStatus;
  inCount: number;
  outCount: number;
  totalCount: number;
  lastInDate: string | null;
  lastOutDate: string | null;
};

type RecoStockRawColumns = {
  date: unknown[][];
  type: unknown[][];
  itemCode: unknown[][];
  qtyMeters: unknown[][];
};

const RECO_STOCK_CHECK_CONFIG = {
  sourceSheet: "RECO_DATA",
  targetSheet: "RECO_STOCK_CHECK",
  startRow: 8, // 1-based row index (0-based index 7)
  columns: {
    date: "C",
    type: "D",
    itemCode: "F",
    qtyMeters: "I",
  },
  headerRow: 5,
  startDataRow: 6,
  templateInputRowCount: 10,
  timezoneOffsetHours: 7,
};

const RECO_STOCK_CHECK_HEADERS = [
  "STT",
  "Mã vật tư (Item Code)",
  "Tổng Nhập IN (m)",
  "Tổng Xuất OUT (m)",
  "Số lượng tồn còn lại (m)",
  "Trạng thái tồn",
  "Ngày nhập cuối",
  "Ngày xuất cuối",
];

const recoStockCleanText = (value: unknown): string => {
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

const recoStockExpandScientific = (input: string): string | null => {
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

const recoStockNormalizeItemCode = (value: unknown): string => {
  if (value === null || value === undefined || typeof value === "boolean") {
    return "";
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    if (Number.isSafeInteger(value)) return String(value);
    const str = String(Math.trunc(value));
    if (str.indexOf("e") !== -1 || str.indexOf("E") !== -1) {
      return recoStockExpandScientific(str) ?? str;
    }
    return str;
  }
  if (typeof value === "string" && value === "") return "";
  const cleaned = String(value).replace(/[\u00A0\u200B\u200C\u200D\uFEFF\s]+/g, "");
  if (!cleaned) return "";
  if (cleaned.indexOf("e") !== -1 || cleaned.indexOf("E") !== -1) {
    const expanded = recoStockExpandScientific(cleaned);
    if (expanded !== null) {
      return expanded;
    }
  }
  if (cleaned.indexOf(".") !== -1 && /^\d+\.0+$/.test(cleaned)) {
    return cleaned.replace(/\.0+$/, "");
  }
  return cleaned;
};

const recoStockToCellItemCode = (itemCode: string): string | number => {
  const normalized = recoStockNormalizeItemCode(itemCode);
  if (!normalized) return "";
  if (/^\d+$/.test(normalized) && !/^0\d+$/.test(normalized)) {
    const num = Number(normalized);
    if (Number.isSafeInteger(num)) {
      return num;
    }
  }
  return normalized;
};

const RECO_STOCK_STOP_MARKERS = new Set([
  "TỔNG CỘNG",
  "TONG CONG",
  "TOTAL",
  "MÃ VẬT TƯ (ITEM CODE)",
  "MÃ HÀNG (ITEM CODE)",
]);

/**
 * Đọc danh sách mã vật tư người dùng nhập từ cột B trên sheet RECO_STOCK_CHECK
 * (bắt đầu từ dòng dữ liệu đầu tiên cho đến dòng TỔNG CỘNG).
 * Giữ nguyên thứ tự nhập của người dùng và loại bỏ trùng lặp.
 */
const recoStockExtractUserItemCodes = (columnBValues: unknown[][]): string[] => {
  const seen = new Set<string>();
  const codes: string[] = [];

  for (const row of columnBValues) {
    const raw = row?.[0];
    const cleanedUpper = recoStockCleanText(raw).toUpperCase();
    if (RECO_STOCK_STOP_MARKERS.has(cleanedUpper)) {
      break;
    }

    const normalized = recoStockNormalizeItemCode(raw);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    codes.push(normalized);
  }

  return codes;
};

const recoStockNormalizeTxnType = (value: unknown): RecoStockTxnType | null => {
  if (value === "IN") return "IN";
  if (value === "OUT") return "OUT";
  const normalized = recoStockCleanText(value).replace(/\s+/g, "").toUpperCase();
  if (normalized === "IN") return "IN";
  if (normalized === "OUT") return "OUT";
  return null;
};

const recoStockParseNumber = (value: unknown): number => {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  if (typeof value === "boolean" || value === null || value === undefined) {
    return 0;
  }
  let s = recoStockCleanText(value).replace(/\s+/g, "");
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

const recoStockPad2 = (n: number): string => (n < 10 ? `0${n}` : String(n));

const recoStockToIsoDate = (
  year: number,
  month: number,
  day: number,
): string | null => {
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day) ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31
  ) {
    return null;
  }
  const d = new Date(Date.UTC(year, month - 1, day));
  if (
    d.getUTCFullYear() !== year ||
    d.getUTCMonth() + 1 !== month ||
    d.getUTCDate() !== day
  ) {
    return null;
  }
  return `${year}-${recoStockPad2(month)}-${recoStockPad2(day)}`;
};

const recoStockParseDate = (value: unknown): string | null => {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    return recoStockToIsoDate(
      value.getFullYear(),
      value.getMonth() + 1,
      value.getDate(),
    );
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value) || value < 1) return null;
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
    return recoStockToIsoDate(
      d.getUTCFullYear(),
      d.getUTCMonth() + 1,
      d.getUTCDate(),
    );
  }

  const text = recoStockCleanText(value);
  if (!text) return null;

  const isoMatch = text.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})(?:\b|T|\s)/);
  if (isoMatch) {
    return recoStockToIsoDate(
      Number(isoMatch[1]),
      Number(isoMatch[2]),
      Number(isoMatch[3]),
    );
  }

  const dmyMatch = text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})(?:\b|\s)/);
  if (dmyMatch) {
    return recoStockToIsoDate(
      Number(dmyMatch[3]),
      Number(dmyMatch[2]),
      Number(dmyMatch[1]),
    );
  }

  if (/^\d+(\.\d+)?$/.test(text)) {
    return recoStockParseDate(Number(text));
  }
  return null;
};

const RECO_STOCK_EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

const recoStockIsoToSerial = (iso: string): number => {
  const year = Number(iso.slice(0, 4));
  const month = Number(iso.slice(5, 7));
  const day = Number(iso.slice(8, 10));
  return Math.round(
    (Date.UTC(year, month - 1, day) - RECO_STOCK_EXCEL_EPOCH_UTC) / 86400000,
  );
};

const recoStockFormatDateVn = (iso: string): string => {
  const [year, month, day] = iso.split("-");
  return `${day}/${month}/${year}`;
};

const recoStockRoundQty = (n: number): number => Math.round(n * 1e6) / 1e6;

const recoStockFormatTimestamp = (offsetHours: number): string => {
  const now = new Date(Date.now() + offsetHours * 3600000);
  return `${recoStockPad2(now.getUTCDate())}/${recoStockPad2(now.getUTCMonth() + 1)}/${now.getUTCFullYear()} ${recoStockPad2(now.getUTCHours())}:${recoStockPad2(now.getUTCMinutes())}`;
};

const recoStockBuildTargetSet = (targetItemCodes: string[]): Set<string> | null => {
  if (targetItemCodes.length === 0) return null;
  const set = new Set<string>();
  for (let i = 0; i < targetItemCodes.length; i++) {
    const code = recoStockNormalizeItemCode(targetItemCodes[i]);
    if (code) set.add(code);
  }
  return set.size > 0 ? set : null;
};

const recoStockResolveTargetItemCode = (
  rawItem: unknown,
  targetSet: Set<string> | null,
): string => {
  if (
    rawItem === null ||
    rawItem === undefined ||
    rawItem === "" ||
    typeof rawItem === "boolean"
  ) {
    return "";
  }
  if (targetSet) {
    if (typeof rawItem === "string" && targetSet.has(rawItem)) {
      return rawItem;
    }
    if (typeof rawItem === "number" && Number.isSafeInteger(rawItem)) {
      const str = String(rawItem);
      return targetSet.has(str) ? str : "";
    }
  }
  const itemCode = recoStockNormalizeItemCode(rawItem);
  if (!itemCode) return "";
  if (targetSet && !targetSet.has(itemCode)) return "";
  return itemCode;
};

const recoStockParseSheetRows = (
  rows: unknown[][],
  targetItemCodes: string[] = [],
  startRow: number = RECO_STOCK_CHECK_CONFIG.startRow,
): RecoStockTransaction[] => {
  const targetSet = recoStockBuildTargetSet(targetItemCodes);
  const transactions: RecoStockTransaction[] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    const itemCode = recoStockResolveTargetItemCode(row[3], targetSet);
    if (!itemCode) continue;

    const type = recoStockNormalizeTxnType(row[1]);
    if (!type) continue;

    const dateIso = recoStockParseDate(row[0]);
    const qtyMeters = recoStockRoundQty(recoStockParseNumber(row[6]));

    transactions.push({
      rowNumber: startRow + i,
      dateIso,
      type,
      itemCode,
      qtyMeters,
    });
  }

  return transactions;
};

const recoStockParseRows = (
  columns: RecoStockRawColumns,
  targetItemCodes: string[] = [],
  startRow: number = RECO_STOCK_CHECK_CONFIG.startRow,
): RecoStockTransaction[] => {
  const targetSet = recoStockBuildTargetSet(targetItemCodes);
  const rowCount = Math.max(
    columns.date.length,
    columns.type.length,
    columns.itemCode.length,
    columns.qtyMeters.length,
  );

  const transactions: RecoStockTransaction[] = [];
  for (let i = 0; i < rowCount; i++) {
    const itemCode = recoStockResolveTargetItemCode(columns.itemCode[i]?.[0], targetSet);
    if (!itemCode) continue;

    const type = recoStockNormalizeTxnType(columns.type[i]?.[0]);
    if (!type) continue;

    const dateIso = recoStockParseDate(columns.date[i]?.[0]);
    const qtyMeters = recoStockRoundQty(
      recoStockParseNumber(columns.qtyMeters[i]?.[0]),
    );

    transactions.push({
      rowNumber: startRow + i,
      dateIso,
      type,
      itemCode,
      qtyMeters,
    });
  }

  return transactions;
};

type RecoStockSourceDiagnostics = {
  totalRows: number;
  nonEmptyItemRows: number;
  matchedItemRows: number;
  validTxnRows: number;
};

const recoStockDiagnoseSourceRows = (
  rows: unknown[][],
  targetItemCodes: string[] = [],
): RecoStockSourceDiagnostics => {
  const targetSet = recoStockBuildTargetSet(targetItemCodes);
  let nonEmptyItemRows = 0;
  let matchedItemRows = 0;
  let validTxnRows = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    const normalized = recoStockNormalizeItemCode(row[3]);
    if (!normalized) continue;
    nonEmptyItemRows += 1;

    if (targetSet && !targetSet.has(normalized)) continue;
    matchedItemRows += 1;

    if (recoStockNormalizeTxnType(row[1]) !== null) {
      validTxnRows += 1;
    }
  }

  return {
    totalRows: rows.length,
    nonEmptyItemRows,
    matchedItemRows,
    validTxnRows,
  };
};

const recoStockGetLastDataRow = (sheet: ExcelScript.Worksheet): number => {
  const used = sheet.getUsedRange(true);
  if (!used) return 0;
  return used.getLastRow().getRowIndex() + 1;
};

const recoStockResolveStatus = (balance: number): RecoStockStatus => {
  const rounded = recoStockRoundQty(balance);
  if (rounded > 0) return "Còn hàng";
  if (rounded < 0) return "Âm kho";
  return "Hết hàng";
};

const recoStockAggregateByItem = (
  transactions: RecoStockTransaction[],
  targetItemCodes: string[] = [],
): RecoStockItemSummary[] => {
  const map = new Map<string, RecoStockItemSummary>();
  const orderedCodes: string[] = [];

  for (const rawCode of targetItemCodes) {
    const code = recoStockNormalizeItemCode(rawCode);
    if (!code || map.has(code)) continue;
    orderedCodes.push(code);
    map.set(code, {
      itemCode: code,
      totalIn: 0,
      totalOut: 0,
      balance: 0,
      status: "Hết hàng",
      inCount: 0,
      outCount: 0,
      totalCount: 0,
      lastInDate: null,
      lastOutDate: null,
    });
  }

  for (const txn of transactions) {
    let entry = map.get(txn.itemCode);
    if (!entry) {
      orderedCodes.push(txn.itemCode);
      entry = {
        itemCode: txn.itemCode,
        totalIn: 0,
        totalOut: 0,
        balance: 0,
        status: "Hết hàng",
        inCount: 0,
        outCount: 0,
        totalCount: 0,
        lastInDate: null,
        lastOutDate: null,
      };
      map.set(txn.itemCode, entry);
    }

    if (txn.type === "IN") {
      entry.totalIn += txn.qtyMeters;
      entry.inCount += 1;
      if (txn.dateIso && (!entry.lastInDate || txn.dateIso > entry.lastInDate)) {
        entry.lastInDate = txn.dateIso;
      }
    } else {
      entry.totalOut += txn.qtyMeters;
      entry.outCount += 1;
      if (txn.dateIso && (!entry.lastOutDate || txn.dateIso > entry.lastOutDate)) {
        entry.lastOutDate = txn.dateIso;
      }
    }
    entry.totalCount += 1;
  }

  return orderedCodes.map((code) => {
    const entry = map.get(code)!;
    entry.totalIn = recoStockRoundQty(entry.totalIn);
    entry.totalOut = recoStockRoundQty(entry.totalOut);
    entry.balance = recoStockRoundQty(entry.totalIn - entry.totalOut);
    entry.status = recoStockResolveStatus(entry.balance);
    return entry;
  });
};

const recoStockSummarizeStatusCounts = (
  items: RecoStockItemSummary[],
): { inStock: number; outOfStock: number; negative: number } => {
  let inStock = 0;
  let outOfStock = 0;
  let negative = 0;
  for (const item of items) {
    if (item.status === "Còn hàng") inStock += 1;
    else if (item.status === "Âm kho") negative += 1;
    else outOfStock += 1;
  }
  return { inStock, outOfStock, negative };
};

const recoStockBuildRows = (
  items: RecoStockItemSummary[],
  startDataRow: number,
): RecoStockCellValue[][] => {
  const rows: RecoStockCellValue[][] = items.map((item, idx) => {
    const row = startDataRow + idx;
    return [
      idx + 1,
      recoStockToCellItemCode(item.itemCode),
      item.totalIn,
      item.totalOut,
      `=C${row}-D${row}`,
      `=IF(ROUND(E${row},6)>0,"Còn hàng",IF(ROUND(E${row},6)<0,"Âm kho","Hết hàng"))`,
      item.lastInDate ? recoStockIsoToSerial(item.lastInDate) : "",
      item.lastOutDate ? recoStockIsoToSerial(item.lastOutDate) : "",
    ];
  });

  if (items.length > 0) {
    const endDataRow = startDataRow + items.length - 1;
    const counts = recoStockSummarizeStatusCounts(items);
    rows.push([
      "",
      "TỔNG CỘNG",
      `=SUM(C${startDataRow}:C${endDataRow})`,
      `=SUM(D${startDataRow}:D${endDataRow})`,
      `=SUM(E${startDataRow}:E${endDataRow})`,
      `Còn: ${counts.inStock} | Hết: ${counts.outOfStock} | Âm: ${counts.negative}`,
      "",
      "",
    ]);
  }

  return rows;
};

const recoStockBuildEmptyTemplateRows = (
  rowCount: number,
  startDataRow: number,
): RecoStockCellValue[][] => {
  const safeCount = Math.max(1, rowCount);
  const rows: RecoStockCellValue[][] = [];
  for (let i = 0; i < safeCount; i++) {
    rows.push([i + 1, "", "", "", "", "", "", ""]);
  }
  const endDataRow = startDataRow + safeCount - 1;
  rows.push([
    "",
    "TỔNG CỘNG",
    `=SUM(C${startDataRow}:C${endDataRow})`,
    `=SUM(D${startDataRow}:D${endDataRow})`,
    `=SUM(E${startDataRow}:E${endDataRow})`,
    "",
    "",
    "",
  ]);
  return rows;
};

const RECO_STOCK_BORDER_ROW_LIMIT = 1000;

const recoStockApplyThinBorders = (
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
  if (rowCount <= RECO_STOCK_BORDER_ROW_LIMIT) {
    edgeIds.push(
      ExcelScript.BorderIndex.insideHorizontal,
      ExcelScript.BorderIndex.insideVertical,
    );
  }
  for (const edgeId of edgeIds) {
    const border = format.getRangeBorder(edgeId);
    border.setStyle(ExcelScript.BorderLineStyle.continuous);
    border.setWeight(ExcelScript.BorderWeight.thin);
    border.setColor("#BFBFBF");
  }
};

const recoStockApplyStatusRuns = (
  sheet: ExcelScript.Worksheet,
  items: RecoStockItemSummary[],
  startDataRow: number,
  targetStatus: RecoStockStatus,
  fillColor: string,
  fontColor: string,
): void => {
  let runStart = -1;
  for (let i = 0; i <= items.length; i++) {
    const match = i < items.length && items[i].status === targetStatus;
    if (match && runStart === -1) {
      runStart = startDataRow + i;
    } else if (!match && runStart !== -1) {
      const runEnd = startDataRow + i - 1;
      const range = sheet.getRange(`E${runStart}:F${runEnd}`);
      range.getFormat().getFill().setColor(fillColor);
      range.getFormat().getFont().setColor(fontColor);
      runStart = -1;
    }
  }
};

// @ts-ignore: Office Scripts
async function main(workbook: ExcelScript.Workbook) {
  const CFG = RECO_STOCK_CHECK_CONFIG;
  const headerRow = CFG.headerRow;
  const startDataRow = CFG.startDataRow;

  let targetSheet = workbook.getWorksheet(CFG.targetSheet);
  let userItemCodes: string[] = [];

  if (targetSheet) {
    const lastTargetRow = recoStockGetLastDataRow(targetSheet);
    if (lastTargetRow >= startDataRow) {
      const colBValues = targetSheet
        .getRange(`B${startDataRow}:B${lastTargetRow}`)
        .getValues();
      userItemCodes = recoStockExtractUserItemCodes(colBValues);
    }
  } else {
    targetSheet = workbook.addWorksheet(CFG.targetSheet);
  }

  const isEmptyTemplateMode = userItemCodes.length === 0;
  let transactions: RecoStockTransaction[] = [];
  let items: RecoStockItemSummary[] = [];

  if (!isEmptyTemplateMode) {
    console.log(
      `Auto_Reco_Stock_Check: Đọc ${userItemCodes.length} mã vật tư từ "${CFG.targetSheet}": ${userItemCodes.join(", ")}`,
    );

    const sourceSheet = workbook.getWorksheet(CFG.sourceSheet);
    if (!sourceSheet) {
      console.log(`LỖI: Không tìm thấy sheet nguồn "${CFG.sourceSheet}".`);
      return;
    }

    const lastSourceRow = recoStockGetLastDataRow(sourceSheet);
    if (lastSourceRow < CFG.startRow) {
      console.log(
        `CẢNH BÁO: Sheet "${CFG.sourceSheet}" chỉ có dữ liệu đến dòng ${lastSourceRow} (< dòng bắt đầu ${CFG.startRow}).`,
      );
    } else {
      const rawMatrix = sourceSheet
        .getRange(
          `${CFG.columns.date}${CFG.startRow}:${CFG.columns.qtyMeters}${lastSourceRow}`,
        )
        .getValues();

      transactions = recoStockParseSheetRows(rawMatrix, userItemCodes, CFG.startRow);
      console.log(
        `Auto_Reco_Stock_Check: Đọc "${CFG.sourceSheet}" dòng ${CFG.startRow}–${lastSourceRow} (${rawMatrix.length} dòng) -> khớp ${transactions.length} giao dịch.`,
      );
      if (rawMatrix.length > 0 && transactions.length === 0) {
        const diag = recoStockDiagnoseSourceRows(rawMatrix, userItemCodes);
        console.log(
          `Chẩn đoán "${CFG.sourceSheet}": ${diag.nonEmptyItemRows}/${diag.totalRows} dòng có Mã vật tư (cột ${CFG.columns.itemCode}), ${diag.matchedItemRows} dòng khớp mã lọc, ${diag.validTxnRows} dòng có Loại GD hợp lệ IN/OUT (cột ${CFG.columns.type}).`,
        );
      }
    }

    items = recoStockAggregateByItem(transactions, userItemCodes);
  }

  const usedTarget = targetSheet.getUsedRange(true);
  if (usedTarget) {
    usedTarget.unmerge();
    usedTarget.clear(ExcelScript.ClearApplyTo.all);
  }
  const autoFilter = targetSheet.getAutoFilter();
  if (autoFilter.getEnabled()) autoFilter.remove();
  targetSheet.getFreezePanes().unfreeze();

  const dataRows = isEmptyTemplateMode
    ? recoStockBuildEmptyTemplateRows(CFG.templateInputRowCount, startDataRow)
    : recoStockBuildRows(items, startDataRow);
  const dataCount = isEmptyTemplateMode
    ? CFG.templateInputRowCount
    : items.length;
  const endDataRow = startDataRow + dataCount - 1;
  const tableEndRow = startDataRow + dataRows.length - 1;
  const totalRow = startDataRow + dataCount;

  const subtitleText = isEmptyTemplateMode
    ? `Hướng dẫn: Nhập Mã vật tư (Item Code) vào cột B (từ dòng ${startDataRow} đến dòng ${endDataRow}), sau đó chạy lại script để kiểm tra số lượng tồn từ "${CFG.sourceSheet}".`
    : `Thời gian: ${recoStockFormatTimestamp(CFG.timezoneOffsetHours)}  |  Nguồn: ${CFG.sourceSheet} (từ dòng ${CFG.startRow})  |  Số mã kiểm tra: ${items.length}  |  Giao dịch khớp: ${transactions.length}`;

  const emptyRow8 = (): RecoStockCellValue[] => ["", "", "", "", "", "", "", ""];

  const sheetMatrix: RecoStockCellValue[][] = [
    ["KIỂM TRA SỐ LƯỢNG TỒN VẬT TƯ (RECO)", "", "", "", "", "", "", ""],
    [subtitleText, "", "", "", "", "", "", ""],
    emptyRow8(),
    ["DANH SÁCH KIỂM TRA TỒN KHO THEO MÃ VẬT TƯ", "", "", "", "", "", "", ""],
    [...RECO_STOCK_CHECK_HEADERS],
    ...dataRows,
  ];

  // Apply integer format "0" before setValues so numeric Item Codes display all digits
  // without scientific notation and without Excel's "Number Stored as Text" warning.
  targetSheet
    .getRange(`B${startDataRow}:B${endDataRow}`)
    .setNumberFormat("0");

  const fullSheetRange = targetSheet.getRange(`A1:H${tableEndRow}`);
  fullSheetRange.setValues(sheetMatrix);
  fullSheetRange.getFormat().getFont().setName("Times New Roman");
  fullSheetRange.getFormat().getFont().setSize(11);
  fullSheetRange
    .getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);

  const titleRange = targetSheet.getRange("A1:H1");
  titleRange.merge(false);
  titleRange.getFormat().getFont().setSize(15);
  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setColor("#1F4E78");
  titleRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  titleRange.getFormat().setRowHeight(28);

  const subRange = targetSheet.getRange("A2:H2");
  subRange.merge(false);
  subRange.getFormat().getFont().setSize(10);
  subRange.getFormat().getFont().setItalic(true);
  if (isEmptyTemplateMode) {
    subRange.getFormat().getFont().setColor("#9C6500");
  }
  subRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  const secTitleRange = targetSheet.getRange("A4:H4");
  secTitleRange.merge(false);
  secTitleRange.getFormat().getFont().setSize(12);
  secTitleRange.getFormat().getFont().setBold(true);
  secTitleRange.getFormat().getFont().setColor("#1F4E78");

  const tableRange = targetSheet.getRange(`A${headerRow}:H${tableEndRow}`);
  tableRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  recoStockApplyThinBorders(tableRange, dataRows.length + 1);

  const headerRange = targetSheet.getRange(`A${headerRow}:H${headerRow}`);
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().getFont().setColor("#FFFFFF");
  headerRange.getFormat().getFill().setColor("#1F4E78");
  headerRange.getFormat().setWrapText(true);
  headerRange.getFormat().setRowHeight(30);

  const qtyRange = targetSheet.getRange(`C${startDataRow}:E${tableEndRow}`);
  qtyRange.setNumberFormat("#,##0.00");
  qtyRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

  targetSheet
    .getRange(`G${startDataRow}:H${endDataRow}`)
    .setNumberFormat("dd/mm/yyyy");

  const totalRowRange = targetSheet.getRange(`A${totalRow}:H${totalRow}`);
  totalRowRange.getFormat().getFont().setBold(true);
  totalRowRange.getFormat().getFill().setColor("#D9E1F2");

  if (isEmptyTemplateMode) {
    targetSheet.getRange(`A${headerRow}:H${tableEndRow}`).getFormat().autofitColumns();
    targetSheet.activate();
    console.log(
      `Auto_Reco_Stock_Check: Đã tạo bảng mẫu trống trên "${CFG.targetSheet}". Vui lòng nhập Mã vật tư vào cột B (dòng ${startDataRow}–${endDataRow}) rồi chạy lại script.`,
    );
    return;
  }

  if (items.length > 0) {
    const statusAndBalRange = targetSheet.getRange(`E${startDataRow}:F${endDataRow}`);
    statusAndBalRange.getFormat().getFont().setBold(true);

    // Default green ("Còn hàng"), then batch-apply yellow ("Hết hàng") and red ("Âm kho") runs
    statusAndBalRange.getFormat().getFill().setColor("#E2EFDA");
    statusAndBalRange.getFormat().getFont().setColor("#375623");

    recoStockApplyStatusRuns(
      targetSheet,
      items,
      startDataRow,
      "Hết hàng",
      "#FFF2CC",
      "#806000",
    );
    recoStockApplyStatusRuns(
      targetSheet,
      items,
      startDataRow,
      "Âm kho",
      "#FFC7CE",
      "#9C0006",
    );

    autoFilter.apply(targetSheet.getRange(`A${headerRow}:H${endDataRow}`));
  }

  targetSheet.getRange(`A${headerRow}:H${tableEndRow}`).getFormat().autofitColumns();
  targetSheet.activate();

  const summaryLog = items
    .map(
      (s) =>
        `${s.itemCode}: Tồn=${s.balance.toFixed(2)}m [${s.status}] (IN=${s.totalIn.toFixed(2)}, OUT=${s.totalOut.toFixed(2)})`,
    )
    .join(" | ");
  console.log(
    `Auto_Reco_Stock_Check hoàn tất (${transactions.length} giao dịch từ "${CFG.sourceSheet}") -> "${CFG.targetSheet}": ${summaryLog}`,
  );
}

const recoStockCheckTestables = {
  RECO_STOCK_CHECK_CONFIG,
  RECO_STOCK_CHECK_HEADERS,
  recoStockCleanText,
  recoStockNormalizeItemCode,
  recoStockToCellItemCode,
  recoStockExtractUserItemCodes,
  recoStockNormalizeTxnType,
  recoStockParseNumber,
  recoStockParseDate,
  recoStockIsoToSerial,
  recoStockFormatDateVn,
  recoStockRoundQty,
  recoStockParseRows,
  recoStockParseSheetRows,
  recoStockDiagnoseSourceRows,
  recoStockGetLastDataRow,
  recoStockResolveStatus,
  recoStockAggregateByItem,
  recoStockSummarizeStatusCounts,
  recoStockBuildRows,
  recoStockBuildEmptyTemplateRows,
};

(
  globalThis as typeof globalThis & {
    __AUTO_RECO_STOCK_CHECK_TESTABLES__?: typeof recoStockCheckTestables;
  }
).__AUTO_RECO_STOCK_CHECK_TESTABLES__ = recoStockCheckTestables;
