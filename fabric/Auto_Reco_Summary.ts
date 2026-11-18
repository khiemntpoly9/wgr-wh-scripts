/**
 * ===================================================================
 * Auto_Reco_Summary – Tổng hợp Nhập (IN) / Xuất (OUT) / Tồn (Balance)
 * theo mã hàng từ sheet "RECO_DATA" ra sheet "RECO_SUMMARY".
 * ===================================================================
 *
 * - Nguồn dữ liệu: sheet "RECO_DATA", dữ liệu bắt đầu từ dòng 8 (index 7).
 *   + Cột C: Ngày tháng năm (Date)
 *   + Cột D: Loại giao dịch ("IN" / "OUT")
 *   + Cột F: Mã hàng (Item Code)
 *   + Cột I: Số lượng mét (Quantity in meters)
 * - Quy trình 2 bước trên sheet "RECO_SUMMARY":
 *   + Lần chạy đầu (hoặc khi chưa nhập mã hàng ở cột B): tạo bảng mẫu trống
 *     (định dạng cột B là "0" để tránh cảnh báo Number Stored as Text và
 *     không bị đổi sang ký hiệu khoa học E+11).
 *   + Các lần chạy sau: đọc danh sách mã hàng người dùng đã nhập ở cột B trên
 *     sheet "RECO_SUMMARY", lọc dữ liệu từ "RECO_DATA" và xuất:
 *     1. Bảng tổng hợp theo từng mã hàng (có công thức Balance = IN - OUT và dòng TỔNG CỘNG).
 *     2. Bảng chi tiết theo ngày & mã hàng.
 */

type RecoCellValue = string | number | boolean;
type RecoTxnType = "IN" | "OUT";

type RecoTransaction = {
  rowNumber: number;
  dateIso: string | null;
  type: RecoTxnType;
  itemCode: string;
  qtyMeters: number;
};

type RecoItemSummary = {
  itemCode: string;
  totalIn: number;
  totalOut: number;
  balance: number;
  inCount: number;
  outCount: number;
  totalCount: number;
  firstDate: string | null;
  lastDate: string | null;
};

type RecoDailySummary = {
  dateIso: string | null;
  itemCode: string;
  type: RecoTxnType;
  qtyMeters: number;
  runningBalance: number;
  txnCount: number;
};

type RecoRawColumns = {
  date: unknown[][];
  type: unknown[][];
  itemCode: unknown[][];
  qtyMeters: unknown[][];
};

const RECO_SUMMARY_CONFIG = {
  sourceSheet: "RECO_DATA",
  targetSheet: "RECO_SUMMARY",
  startRow: 8, // 1-based row index (0-based index 7)
  columns: {
    date: "C",
    type: "D",
    itemCode: "F",
    qtyMeters: "I",
  },
  summaryHeaderRow: 5,
  summaryStartRow: 6,
  templateInputRowCount: 10,
  includeDailyBreakdown: true,
  timezoneOffsetHours: 7,
};

const RECO_SUMMARY_HEADERS = [
  "STT",
  "Mã hàng (Item Code)",
  "Tổng Nhập IN (m)",
  "Tổng Xuất OUT (m)",
  "Tồn Balance (m)",
  "Số lần IN",
  "Số lần OUT",
  "Tổng giao dịch",
  "Từ ngày",
  "Đến ngày",
];

const RECO_DAILY_HEADERS = [
  "STT",
  "Ngày",
  "Mã hàng (Item Code)",
  "Loại (IN/OUT)",
  "Số mét (m)",
  "Số lượng tồn (m)",
  "Số giao dịch",
];

const recoCleanText = (value: unknown): string => {
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

const recoExpandScientific = (input: string): string | null => {
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

const recoNormalizeItemCode = (value: unknown): string => {
  if (value === null || value === undefined || typeof value === "boolean") {
    return "";
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    if (Number.isSafeInteger(value)) return String(value);
    const str = String(Math.trunc(value));
    if (str.indexOf("e") !== -1 || str.indexOf("E") !== -1) {
      return recoExpandScientific(str) ?? str;
    }
    return str;
  }
  if (typeof value === "string" && value === "") return "";
  const cleaned = String(value).replace(/[\u00A0\u200B\u200C\u200D\uFEFF\s]+/g, "");
  if (!cleaned) return "";
  if (cleaned.indexOf("e") !== -1 || cleaned.indexOf("E") !== -1) {
    const expanded = recoExpandScientific(cleaned);
    if (expanded !== null) {
      return expanded;
    }
  }
  if (cleaned.indexOf(".") !== -1 && /^\d+\.0+$/.test(cleaned)) {
    return cleaned.replace(/\.0+$/, "");
  }
  return cleaned;
};

/**
 * Chuyển mã hàng sang kiểu `number` nếu là số nguyên an toàn (không có số 0 đứng đầu),
 * kết hợp định dạng ô `"0"` để hiển thị đủ chữ số mà không gây cảnh báo "Number Stored as Text".
 * Mã hàng chứa chữ/ký tự đặc biệt được giữ nguyên dạng `string`.
 */
const recoToCellItemCode = (itemCode: string): string | number => {
  const normalized = recoNormalizeItemCode(itemCode);
  if (!normalized) return "";
  if (/^\d+$/.test(normalized) && !/^0\d+$/.test(normalized)) {
    const num = Number(normalized);
    if (Number.isSafeInteger(num)) {
      return num;
    }
  }
  return normalized;
};

const RECO_STOP_MARKERS = new Set([
  "TỔNG CỘNG",
  "TONG CONG",
  "TOTAL",
  "MÃ HÀNG (ITEM CODE)",
  "NGÀY",
]);

/**
 * Đọc danh sách mã hàng người dùng nhập từ cột B trên sheet RECO_SUMMARY
 * (bắt đầu từ dòng dữ liệu đầu tiên của bảng tổng hợp cho đến dòng TỔNG CỘNG hoặc phần chi tiết bên dưới).
 * Giữ nguyên thứ tự nhập của người dùng và loại bỏ trùng lặp.
 */
const recoExtractUserItemCodes = (columnBValues: unknown[][]): string[] => {
  const seen = new Set<string>();
  const codes: string[] = [];

  for (const row of columnBValues) {
    const raw = row?.[0];
    const cleanedUpper = recoCleanText(raw).toUpperCase();
    if (RECO_STOP_MARKERS.has(cleanedUpper)) {
      break;
    }
    if (
      cleanedUpper.startsWith("2. ") ||
      cleanedUpper.startsWith("CHI TIẾT THEO NGÀY")
    ) {
      break;
    }

    const normalized = recoNormalizeItemCode(raw);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    codes.push(normalized);
  }

  return codes;
};

const recoNormalizeTxnType = (value: unknown): RecoTxnType | null => {
  if (value === "IN") return "IN";
  if (value === "OUT") return "OUT";
  const normalized = recoCleanText(value).replace(/\s+/g, "").toUpperCase();
  if (normalized === "IN") return "IN";
  if (normalized === "OUT") return "OUT";
  return null;
};

const recoParseNumber = (value: unknown): number => {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  if (typeof value === "boolean" || value === null || value === undefined) {
    return 0;
  }
  let s = recoCleanText(value).replace(/\s+/g, "");
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

const recoPad2 = (n: number): string => (n < 10 ? `0${n}` : String(n));

const recoToIsoDate = (year: number, month: number, day: number): string | null => {
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
  return `${year}-${recoPad2(month)}-${recoPad2(day)}`;
};

/**
 * Chuyển Excel serial, Date object hoặc text (dd/mm/yyyy, yyyy-mm-dd, dd-mm-yyyy, dd.mm.yyyy)
 * sang chuẩn ISO yyyy-mm-dd.
 */
const recoParseDate = (value: unknown): string | null => {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    return recoToIsoDate(
      value.getFullYear(),
      value.getMonth() + 1,
      value.getDate(),
    );
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value) || value < 1) return null;
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
    return recoToIsoDate(
      d.getUTCFullYear(),
      d.getUTCMonth() + 1,
      d.getUTCDate(),
    );
  }

  const text = recoCleanText(value);
  if (!text) return null;

  const isoMatch = text.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})(?:\b|T|\s)/);
  if (isoMatch) {
    return recoToIsoDate(
      Number(isoMatch[1]),
      Number(isoMatch[2]),
      Number(isoMatch[3]),
    );
  }

  const dmyMatch = text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})(?:\b|\s)/);
  if (dmyMatch) {
    return recoToIsoDate(
      Number(dmyMatch[3]),
      Number(dmyMatch[2]),
      Number(dmyMatch[1]),
    );
  }

  if (/^\d+(\.\d+)?$/.test(text)) {
    return recoParseDate(Number(text));
  }
  return null;
};

const RECO_EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

const recoIsoToSerial = (iso: string): number => {
  const year = Number(iso.slice(0, 4));
  const month = Number(iso.slice(5, 7));
  const day = Number(iso.slice(8, 10));
  return Math.round(
    (Date.UTC(year, month - 1, day) - RECO_EXCEL_EPOCH_UTC) / 86400000,
  );
};

const recoFormatDateVn = (iso: string): string => {
  const [year, month, day] = iso.split("-");
  return `${day}/${month}/${year}`;
};

const recoRoundQty = (n: number): number => Math.round(n * 1e6) / 1e6;

const recoFormatTimestamp = (offsetHours: number): string => {
  const now = new Date(Date.now() + offsetHours * 3600000);
  return `${recoPad2(now.getUTCDate())}/${recoPad2(now.getUTCMonth() + 1)}/${now.getUTCFullYear()} ${recoPad2(now.getUTCHours())}:${recoPad2(now.getUTCMinutes())}`;
};

const recoBuildTargetSet = (targetItemCodes: string[]): Set<string> | null => {
  if (targetItemCodes.length === 0) return null;
  const set = new Set<string>();
  for (let i = 0; i < targetItemCodes.length; i++) {
    const code = recoNormalizeItemCode(targetItemCodes[i]);
    if (code) set.add(code);
  }
  return set.size > 0 ? set : null;
};

const recoResolveTargetItemCode = (
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
  const itemCode = recoNormalizeItemCode(rawItem);
  if (!itemCode) return "";
  if (targetSet && !targetSet.has(itemCode)) return "";
  return itemCode;
};

/**
 * Parse contiguous C:I matrix from RECO_DATA (col 0=C Date, 1=D Type, 3=F Item Code, 6=I Qty).
 * Filters by Item Code FIRST before parsing txnType, date, or quantity.
 */
const recoParseSheetRows = (
  rows: unknown[][],
  targetItemCodes: string[] = [],
  startRow: number = RECO_SUMMARY_CONFIG.startRow,
): RecoTransaction[] => {
  const targetSet = recoBuildTargetSet(targetItemCodes);
  const transactions: RecoTransaction[] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    const itemCode = recoResolveTargetItemCode(row[3], targetSet);
    if (!itemCode) continue;

    const type = recoNormalizeTxnType(row[1]);
    if (!type) continue;

    const dateIso = recoParseDate(row[0]);
    const qtyMeters = recoRoundQty(recoParseNumber(row[6]));

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

type RecoSourceDiagnostics = {
  totalRows: number;
  nonEmptyItemRows: number;
  matchedItemRows: number;
  validTxnRows: number;
};

/**
 * Đếm số dòng có mã hàng, số dòng khớp danh sách lọc và số dòng có loại giao dịch hợp lệ
 * để ghi log chẩn đoán khi không có giao dịch nào được tổng hợp từ RECO_DATA.
 */
const recoDiagnoseSourceRows = (
  rows: unknown[][],
  targetItemCodes: string[] = [],
): RecoSourceDiagnostics => {
  const targetSet = recoBuildTargetSet(targetItemCodes);
  let nonEmptyItemRows = 0;
  let matchedItemRows = 0;
  let validTxnRows = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    const normalized = recoNormalizeItemCode(row[3]);
    if (!normalized) continue;
    nonEmptyItemRows += 1;

    if (targetSet && !targetSet.has(normalized)) continue;
    matchedItemRows += 1;

    if (recoNormalizeTxnType(row[1]) !== null) {
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

/**
 * Lấy chỉ số dòng cuối cùng (1-based) có dữ liệu thực tế trên sheet bằng getUsedRange(true)
 * để tránh bị dừng sớm bởi AutoFilter, dòng ẩn, ô trống ở cột F hoặc ô gộp.
 */
const recoGetLastDataRow = (sheet: ExcelScript.Worksheet): number => {
  const used = sheet.getUsedRange(true);
  if (!used) return 0;
  return used.getLastRow().getRowIndex() + 1;
};

const recoParseRows = (
  columns: RecoRawColumns,
  targetItemCodes: string[] = [],
  startRow: number = RECO_SUMMARY_CONFIG.startRow,
): RecoTransaction[] => {
  const targetSet = recoBuildTargetSet(targetItemCodes);

  const rowCount = Math.max(
    columns.date.length,
    columns.type.length,
    columns.itemCode.length,
    columns.qtyMeters.length,
  );

  const transactions: RecoTransaction[] = [];
  for (let i = 0; i < rowCount; i++) {
    const itemCode = recoResolveTargetItemCode(columns.itemCode[i]?.[0], targetSet);
    if (!itemCode) continue;

    const type = recoNormalizeTxnType(columns.type[i]?.[0]);
    if (!type) continue;

    const dateIso = recoParseDate(columns.date[i]?.[0]);
    const qtyMeters = recoRoundQty(recoParseNumber(columns.qtyMeters[i]?.[0]));

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

const recoAggregateByItem = (
  transactions: RecoTransaction[],
  targetItemCodes: string[] = [],
): RecoItemSummary[] => {
  const map = new Map<string, RecoItemSummary>();
  const orderedCodes: string[] = [];

  for (const rawCode of targetItemCodes) {
    const code = recoNormalizeItemCode(rawCode);
    if (!code || map.has(code)) continue;
    orderedCodes.push(code);
    map.set(code, {
      itemCode: code,
      totalIn: 0,
      totalOut: 0,
      balance: 0,
      inCount: 0,
      outCount: 0,
      totalCount: 0,
      firstDate: null,
      lastDate: null,
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
        inCount: 0,
        outCount: 0,
        totalCount: 0,
        firstDate: null,
        lastDate: null,
      };
      map.set(txn.itemCode, entry);
    }

    if (txn.type === "IN") {
      entry.totalIn += txn.qtyMeters;
      entry.inCount += 1;
    } else {
      entry.totalOut += txn.qtyMeters;
      entry.outCount += 1;
    }
    entry.totalCount += 1;

    if (txn.dateIso) {
      if (!entry.firstDate || txn.dateIso < entry.firstDate) {
        entry.firstDate = txn.dateIso;
      }
      if (!entry.lastDate || txn.dateIso > entry.lastDate) {
        entry.lastDate = txn.dateIso;
      }
    }
  }

  return orderedCodes.map((code) => {
    const entry = map.get(code)!;
    entry.totalIn = recoRoundQty(entry.totalIn);
    entry.totalOut = recoRoundQty(entry.totalOut);
    entry.balance = recoRoundQty(entry.totalIn - entry.totalOut);
    return entry;
  });
};

const recoAggregateByDateAndItem = (
  transactions: RecoTransaction[],
): RecoDailySummary[] => {
  const map = new Map<string, RecoDailySummary>();

  for (const txn of transactions) {
    const key = `${txn.dateIso ?? ""}|${txn.itemCode}|${txn.type}`;
    let entry = map.get(key);
    if (!entry) {
      entry = {
        dateIso: txn.dateIso,
        itemCode: txn.itemCode,
        type: txn.type,
        qtyMeters: 0,
        runningBalance: 0,
        txnCount: 0,
      };
      map.set(key, entry);
    }
    entry.qtyMeters += txn.qtyMeters;
    entry.txnCount += 1;
  }

  const result = Array.from(map.values());
  for (let i = 0; i < result.length; i++) {
    result[i].qtyMeters = recoRoundQty(result[i].qtyMeters);
  }

  result.sort((a, b) => {
    if (a.itemCode !== b.itemCode) {
      return a.itemCode < b.itemCode ? -1 : 1;
    }
    if (a.dateIso && b.dateIso && a.dateIso !== b.dateIso) {
      return a.dateIso < b.dateIso ? -1 : 1;
    }
    if (a.dateIso && !b.dateIso) return -1;
    if (!a.dateIso && b.dateIso) return 1;
    if (a.type !== b.type) {
      return a.type === "IN" ? -1 : 1;
    }
    return 0;
  });

  const runningByItem = new Map<string, number>();
  for (let i = 0; i < result.length; i++) {
    const entry = result[i];
    const prev = runningByItem.get(entry.itemCode) ?? 0;
    const next = recoRoundQty(
      entry.type === "IN" ? prev + entry.qtyMeters : prev - entry.qtyMeters,
    );
    entry.runningBalance = next;
    runningByItem.set(entry.itemCode, next);
  }

  return result;
};

const recoBuildSummaryRows = (
  summaries: RecoItemSummary[],
  startDataRow: number,
): RecoCellValue[][] => {
  const rows: RecoCellValue[][] = summaries.map((item, idx) => {
    const row = startDataRow + idx;
    return [
      idx + 1,
      recoToCellItemCode(item.itemCode),
      item.totalIn,
      item.totalOut,
      `=C${row}-D${row}`,
      item.inCount,
      item.outCount,
      `=F${row}+G${row}`,
      item.firstDate ? recoIsoToSerial(item.firstDate) : "",
      item.lastDate ? recoIsoToSerial(item.lastDate) : "",
    ];
  });

  if (summaries.length > 0) {
    const endDataRow = startDataRow + summaries.length - 1;
    rows.push([
      "",
      "TỔNG CỘNG",
      `=SUM(C${startDataRow}:C${endDataRow})`,
      `=SUM(D${startDataRow}:D${endDataRow})`,
      `=SUM(E${startDataRow}:E${endDataRow})`,
      `=SUM(F${startDataRow}:F${endDataRow})`,
      `=SUM(G${startDataRow}:G${endDataRow})`,
      `=SUM(H${startDataRow}:H${endDataRow})`,
      "",
      "",
    ]);
  }

  return rows;
};

const recoBuildEmptyTemplateRows = (
  rowCount: number,
  _startDataRow: number = 0,
): RecoCellValue[][] => {
  const safeCount = Math.max(1, rowCount);
  const rows: RecoCellValue[][] = [];
  for (let i = 0; i < safeCount; i++) {
    rows.push([i + 1, "", "", "", "", "", "", "", "", ""]);
  }
  return rows;
};

const recoBuildDailyRows = (
  daily: RecoDailySummary[],
  _startDataRow: number = 0,
): RecoCellValue[][] =>
  daily.map((entry, idx) => [
    idx + 1,
    entry.dateIso ? recoIsoToSerial(entry.dateIso) : "",
    recoToCellItemCode(entry.itemCode),
    entry.type,
    entry.qtyMeters,
    entry.runningBalance,
    entry.txnCount,
  ]);

const RECO_LARGE_TABLE_BORDER_ROW_LIMIT = 1000;

const recoApplyThinBorders = (
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
  if (rowCount <= RECO_LARGE_TABLE_BORDER_ROW_LIMIT) {
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

const recoPadRow10 = (row: RecoCellValue[]): RecoCellValue[] => [
  row[0] ?? "",
  row[1] ?? "",
  row[2] ?? "",
  row[3] ?? "",
  row[4] ?? "",
  row[5] ?? "",
  row[6] ?? "",
  row[7] ?? "",
  row[8] ?? "",
  row[9] ?? "",
];

// @ts-ignore: Office Scripts
async function main(workbook: ExcelScript.Workbook) {
  const CFG = RECO_SUMMARY_CONFIG;
  const summaryHeaderRow = CFG.summaryHeaderRow;
  const summaryStartRow = CFG.summaryStartRow;

  let targetSheet = workbook.getWorksheet(CFG.targetSheet);
  let userItemCodes: string[] = [];

  if (targetSheet) {
    const lastTargetRow = recoGetLastDataRow(targetSheet);
    if (lastTargetRow >= summaryStartRow) {
      const colBValues = targetSheet
        .getRange(`B${summaryStartRow}:B${lastTargetRow}`)
        .getValues();
      userItemCodes = recoExtractUserItemCodes(colBValues);
    }
  } else {
    targetSheet = workbook.addWorksheet(CFG.targetSheet);
  }

  const isEmptyTemplateMode = userItemCodes.length === 0;
  let transactions: RecoTransaction[] = [];
  let summaries: RecoItemSummary[] = [];
  let daily: RecoDailySummary[] = [];

  if (!isEmptyTemplateMode) {
    console.log(
      `Auto_Reco_Summary: Đọc ${userItemCodes.length} mã hàng từ "${CFG.targetSheet}": ${userItemCodes.join(", ")}`,
    );

    const sourceSheet = workbook.getWorksheet(CFG.sourceSheet);
    if (!sourceSheet) {
      console.log(`LỖI: Không tìm thấy sheet nguồn "${CFG.sourceSheet}".`);
      return;
    }

    const lastSourceRow = recoGetLastDataRow(sourceSheet);
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

      transactions = recoParseSheetRows(rawMatrix, userItemCodes, CFG.startRow);
      console.log(
        `Auto_Reco_Summary: Đọc "${CFG.sourceSheet}" dòng ${CFG.startRow}–${lastSourceRow} (${rawMatrix.length} dòng) -> khớp ${transactions.length} giao dịch.`,
      );
      if (rawMatrix.length > 0 && transactions.length === 0) {
        const diag = recoDiagnoseSourceRows(rawMatrix, userItemCodes);
        console.log(
          `Chẩn đoán "${CFG.sourceSheet}": ${diag.nonEmptyItemRows}/${diag.totalRows} dòng có Mã hàng (cột ${CFG.columns.itemCode}), ${diag.matchedItemRows} dòng khớp mã lọc, ${diag.validTxnRows} dòng có Loại GD hợp lệ IN/OUT (cột ${CFG.columns.type}).`,
        );
      }
    }

    summaries = recoAggregateByItem(transactions, userItemCodes);
    daily = CFG.includeDailyBreakdown
      ? recoAggregateByDateAndItem(transactions)
      : [];
  }

  const usedTarget = targetSheet.getUsedRange(true);
  if (usedTarget) {
    usedTarget.unmerge();
    usedTarget.clear(ExcelScript.ClearApplyTo.all);
  }
  const autoFilter = targetSheet.getAutoFilter();
  if (autoFilter.getEnabled()) autoFilter.remove();
  targetSheet.getFreezePanes().unfreeze();

  const summaryRows = isEmptyTemplateMode
    ? recoBuildEmptyTemplateRows(CFG.templateInputRowCount, summaryStartRow)
    : recoBuildSummaryRows(summaries, summaryStartRow);
  const summaryDataCount = isEmptyTemplateMode
    ? CFG.templateInputRowCount
    : summaries.length;
  const summaryDataEndRow = summaryStartRow + summaryDataCount - 1;
  const summaryEndRow = summaryStartRow + summaryRows.length - 1;
  const summaryTotalRow = summaryStartRow + summaryDataCount;

  const subtitleText = isEmptyTemplateMode
    ? `Hướng dẫn: Nhập Mã hàng (Item Code) vào cột B (từ dòng ${summaryStartRow} trở đi, không giới hạn số dòng), sau đó chạy lại script để tổng hợp dữ liệu từ "${CFG.sourceSheet}".`
    : `Thời gian: ${recoFormatTimestamp(CFG.timezoneOffsetHours)}  |  Nguồn: ${CFG.sourceSheet} (từ dòng ${CFG.startRow})  |  Mã hàng lọc: ${userItemCodes.join(", ")}  |  Số giao dịch khớp: ${transactions.length}`;

  const emptyRow10 = (): RecoCellValue[] => ["", "", "", "", "", "", "", "", "", ""];

  const sheetMatrix: RecoCellValue[][] = [
    ["BÁO CÁO TỔNG HỢP NHẬP - XUẤT - TỒN VẢI (RECO)", "", "", "", "", "", "", "", "", ""],
    [subtitleText, "", "", "", "", "", "", "", "", ""],
    emptyRow10(),
    ["1. TỔNG HỢP THEO MÃ HÀNG (SUMMARY BY ITEM CODE)", "", "", "", "", "", "", "", "", ""],
    [...RECO_SUMMARY_HEADERS],
    ...summaryRows,
  ];

  const hasDailySection = !isEmptyTemplateMode && CFG.includeDailyBreakdown;
  let sec2TitleRow = 0;
  let dailyHeaderRow = 0;
  let dailyStartRow = 0;
  let dailyEndRow = 0;
  let dailyRows: RecoCellValue[][] = [];
  let finalBottomRow = summaryEndRow;

  if (hasDailySection) {
    sec2TitleRow = summaryEndRow + 2;
    dailyHeaderRow = sec2TitleRow + 1;
    dailyStartRow = dailyHeaderRow + 1;
    dailyRows = recoBuildDailyRows(daily, dailyStartRow);
    dailyEndRow = dailyStartRow + Math.max(dailyRows.length, 1) - 1;
    finalBottomRow = dailyEndRow;

    sheetMatrix.push(emptyRow10());
    sheetMatrix.push([
      "2. CHI TIẾT THEO NGÀY & MÃ HÀNG (DAILY BREAKDOWN)",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
    ]);
    sheetMatrix.push(recoPadRow10(RECO_DAILY_HEADERS));
    if (dailyRows.length > 0) {
      for (let i = 0; i < dailyRows.length; i++) {
        sheetMatrix.push(recoPadRow10(dailyRows[i]));
      }
    } else {
      sheetMatrix.push([
        "",
        "",
        "Không có giao dịch nào cho mã hàng đã chọn.",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
      ]);
    }
  }

  // Apply integer format "0" before setValues so numeric Item Codes display all digits
  // without scientific notation and without Excel's "Number Stored as Text" warning.
  // In template mode, format the entire column from summaryStartRow downward (no row limit).
  const itemCodeFormatRange = isEmptyTemplateMode
    ? `B${summaryStartRow}:B1048576`
    : `B${summaryStartRow}:B${summaryDataEndRow}`;
  targetSheet.getRange(itemCodeFormatRange).setNumberFormat("0");
  if (hasDailySection && dailyRows.length > 0) {
    targetSheet
      .getRange(`C${dailyStartRow}:C${dailyEndRow}`)
      .setNumberFormat("0");
  }

  // Single contiguous write for the entire RECO_SUMMARY sheet (values + formulas).
  const fullSheetRange = targetSheet.getRange(`A1:J${finalBottomRow}`);
  fullSheetRange.setValues(sheetMatrix);
  fullSheetRange.getFormat().getFont().setName("Times New Roman");
  fullSheetRange.getFormat().getFont().setSize(11);
  fullSheetRange
    .getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);

  // Title & Metadata / Instruction
  const titleRange = targetSheet.getRange("A1:J1");
  titleRange.merge(false);
  titleRange.getFormat().getFont().setSize(15);
  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setColor("#1F4E78");
  titleRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  titleRange.getFormat().setRowHeight(28);

  const subRange = targetSheet.getRange("A2:J2");
  subRange.merge(false);
  subRange.getFormat().getFont().setSize(10);
  subRange.getFormat().getFont().setItalic(true);
  if (isEmptyTemplateMode) {
    subRange.getFormat().getFont().setColor("#9C6500");
  }
  subRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Section 1: Summary by Item Code
  const sec1Range = targetSheet.getRange("A4:J4");
  sec1Range.merge(false);
  sec1Range.getFormat().getFont().setSize(12);
  sec1Range.getFormat().getFont().setBold(true);
  sec1Range.getFormat().getFont().setColor("#1F4E78");

  const summaryTableRange = targetSheet.getRange(
    `A${summaryHeaderRow}:J${summaryEndRow}`,
  );
  summaryTableRange
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  recoApplyThinBorders(summaryTableRange, summaryRows.length + 1);

  const summaryHeaderRange = targetSheet.getRange(
    `A${summaryHeaderRow}:J${summaryHeaderRow}`,
  );
  summaryHeaderRange.getFormat().getFont().setBold(true);
  summaryHeaderRange.getFormat().getFont().setColor("#FFFFFF");
  summaryHeaderRange.getFormat().getFill().setColor("#1F4E78");
  summaryHeaderRange.getFormat().setWrapText(true);
  summaryHeaderRange.getFormat().setRowHeight(30);

  const qtyFormatEndRow = isEmptyTemplateMode ? 1048576 : summaryEndRow;
  const dateFormatEndRow = isEmptyTemplateMode ? 1048576 : summaryDataEndRow;
  const summaryQtyRange = targetSheet.getRange(`C${summaryStartRow}:E${qtyFormatEndRow}`);
  summaryQtyRange.setNumberFormat("#,##0.00");
  summaryQtyRange
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  targetSheet
    .getRange(`F${summaryStartRow}:H${qtyFormatEndRow}`)
    .setNumberFormat("#,##0");
  targetSheet
    .getRange(`I${summaryStartRow}:J${dateFormatEndRow}`)
    .setNumberFormat("dd/mm/yyyy");

  if (isEmptyTemplateMode) {
    targetSheet.getRange(`A${summaryHeaderRow}:J${summaryEndRow}`).getFormat().autofitColumns();
    targetSheet.activate();
    console.log(
      `Auto_Reco_Summary: Đã tạo bảng mẫu trống trên "${CFG.targetSheet}". Vui lòng nhập Mã hàng (Item Code) vào cột B (từ dòng ${summaryStartRow} trở đi, không giới hạn số dòng) rồi chạy lại script.`,
    );
    return;
  }

  const totalRowRange = targetSheet.getRange(
    `A${summaryTotalRow}:J${summaryTotalRow}`,
  );
  totalRowRange.getFormat().getFont().setBold(true);
  totalRowRange.getFormat().getFill().setColor("#D9E1F2");

  // Batch-format Balance column: apply positive style once, then batch contiguous negative runs.
  if (summaries.length > 0) {
    const balRange = targetSheet.getRange(`E${summaryStartRow}:E${summaryDataEndRow}`);
    balRange.getFormat().getFont().setBold(true);
    balRange.getFormat().getFill().setColor("#E2EFDA");
    balRange.getFormat().getFont().setColor("#375623");

    let negRunStart = -1;
    for (let i = 0; i <= summaries.length; i++) {
      const isNeg = i < summaries.length && summaries[i].balance < 0;
      if (isNeg && negRunStart === -1) {
        negRunStart = summaryStartRow + i;
      } else if (!isNeg && negRunStart !== -1) {
        const negRunEnd = summaryStartRow + i - 1;
        const negRange = targetSheet.getRange(`E${negRunStart}:E${negRunEnd}`);
        negRange.getFormat().getFill().setColor("#FFC7CE");
        negRange.getFormat().getFont().setColor("#9C0006");
        negRunStart = -1;
      }
    }
  }

  // Section 2: Daily Breakdown formatting
  if (hasDailySection) {
    const sec2Range = targetSheet.getRange(`A${sec2TitleRow}:G${sec2TitleRow}`);
    sec2Range.merge(false);
    sec2Range.getFormat().getFont().setSize(12);
    sec2Range.getFormat().getFont().setBold(true);
    sec2Range.getFormat().getFont().setColor("#1F4E78");

    const dailyTableRange = targetSheet.getRange(
      `A${dailyHeaderRow}:G${dailyEndRow}`,
    );
    dailyTableRange
      .getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    recoApplyThinBorders(dailyTableRange, dailyRows.length + 1);

    const dailyHeaderRange = targetSheet.getRange(
      `A${dailyHeaderRow}:G${dailyHeaderRow}`,
    );
    dailyHeaderRange.getFormat().getFont().setBold(true);
    dailyHeaderRange.getFormat().getFont().setColor("#FFFFFF");
    dailyHeaderRange.getFormat().getFill().setColor("#305496");
    dailyHeaderRange.getFormat().setWrapText(true);
    dailyHeaderRange.getFormat().setRowHeight(28);

    autoFilter.apply(dailyRows.length > 0 ? dailyTableRange : dailyHeaderRange);

    if (dailyRows.length > 0) {
      targetSheet
        .getRange(`B${dailyStartRow}:B${dailyEndRow}`)
        .setNumberFormat("dd/mm/yyyy");
      const dailyQtyRange = targetSheet.getRange(`E${dailyStartRow}:F${dailyEndRow}`);
      dailyQtyRange.setNumberFormat("#,##0.00");
      dailyQtyRange
        .getFormat()
        .setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
      targetSheet
        .getRange(`G${dailyStartRow}:G${dailyEndRow}`)
        .setNumberFormat("#,##0");

      const dailyBalRange = targetSheet.getRange(`F${dailyStartRow}:F${dailyEndRow}`);
      dailyBalRange.getFormat().getFont().setBold(true);
      dailyBalRange.getFormat().getFill().setColor("#E2EFDA");
      dailyBalRange.getFormat().getFont().setColor("#375623");

      let dailyNegRunStart = -1;
      for (let i = 0; i <= daily.length; i++) {
        const isNeg = i < daily.length && daily[i].runningBalance < 0;
        if (isNeg && dailyNegRunStart === -1) {
          dailyNegRunStart = dailyStartRow + i;
        } else if (!isNeg && dailyNegRunStart !== -1) {
          const dailyNegRunEnd = dailyStartRow + i - 1;
          const dailyNegRange = targetSheet.getRange(
            `F${dailyNegRunStart}:F${dailyNegRunEnd}`,
          );
          dailyNegRange.getFormat().getFill().setColor("#FFC7CE");
          dailyNegRange.getFormat().getFont().setColor("#9C0006");
          dailyNegRunStart = -1;
        }
      }
    }
  }

  const autofitEndRow = Math.min(finalBottomRow, summaryEndRow + 50);
  targetSheet.getRange(`A${summaryHeaderRow}:J${autofitEndRow}`).getFormat().autofitColumns();
  targetSheet.activate();

  const summaryLog = summaries
    .map(
      (s) =>
        `${s.itemCode}: IN=${s.totalIn.toFixed(2)}m, OUT=${s.totalOut.toFixed(2)}m, Balance=${s.balance.toFixed(2)}m (${s.totalCount} txns)`,
    )
    .join(" | ");
  console.log(
    `Auto_Reco_Summary hoàn tất (${transactions.length} giao dịch từ "${CFG.sourceSheet}") -> "${CFG.targetSheet}": ${summaryLog}`,
  );
}

const recoSummaryTestables = {
  RECO_SUMMARY_CONFIG,
  RECO_SUMMARY_HEADERS,
  RECO_DAILY_HEADERS,
  recoCleanText,
  recoNormalizeItemCode,
  recoToCellItemCode,
  recoExtractUserItemCodes,
  recoNormalizeTxnType,
  recoParseNumber,
  recoParseDate,
  recoIsoToSerial,
  recoFormatDateVn,
  recoRoundQty,
  recoParseRows,
  recoParseSheetRows,
  recoDiagnoseSourceRows,
  recoGetLastDataRow,
  recoAggregateByItem,
  recoAggregateByDateAndItem,
  recoBuildSummaryRows,
  recoBuildEmptyTemplateRows,
  recoBuildDailyRows,
};

(
  globalThis as typeof globalThis & {
    __AUTO_RECO_SUMMARY_TESTABLES__?: typeof recoSummaryTestables;
  }
).__AUTO_RECO_SUMMARY_TESTABLES__ = recoSummaryTestables;
