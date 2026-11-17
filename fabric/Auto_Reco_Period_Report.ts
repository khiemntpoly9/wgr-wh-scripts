/**
 * ===================================================================
 * Auto_Reco_Period_Report – Báo cáo Nhập (IN) / Xuất (OUT) theo
 * Ngày / Tuần / Tháng từ sheet "RECO_DATA" ra sheet "RECO_REPORT".
 * ===================================================================
 *
 * - Nguồn dữ liệu: sheet "RECO_DATA", dữ liệu bắt đầu từ dòng 8 (index 7).
 *   + Cột C: Ngày tháng năm (Date)
 *   + Cột D: Loại giao dịch ("IN" / "OUT")
 *   + Cột F: Mã hàng (Item Code)
 *   + Cột I: Số lượng mét (Quantity in meters)
 * - Xuất ra sheet "RECO_REPORT" 3 bảng tổng hợp:
 *   1. Theo Ngày (Daily)
 *   2. Theo Tuần ISO (Weekly: Thứ 2 -> Chủ nhật)
 *   3. Theo Tháng (Monthly)
 */

type RecoReportCellValue = string | number | boolean;
type RecoReportTxnType = "IN" | "OUT";

type RecoReportTransaction = {
  rowNumber: number;
  dateIso: string | null;
  type: RecoReportTxnType;
  itemCode: string;
  qtyMeters: number;
};

type RecoPeriodSummary = {
  periodKey: string;
  periodLabel: string;
  startDate: string;
  endDate: string;
  totalIn: number;
  totalOut: number;
  net: number;
  inCount: number;
  outCount: number;
  totalCount: number;
  uniqueItemCount: number;
};

type RecoReportFilterOptions = {
  fromDate?: string;
  toDate?: string;
  itemCodes?: string[];
};

type RecoReportRawColumns = {
  date: unknown[][];
  type: unknown[][];
  itemCode: unknown[][];
  qtyMeters: unknown[][];
};

const RECO_PERIOD_REPORT_CONFIG = {
  sourceSheet: "RECO_DATA",
  targetSheet: "RECO_REPORT",
  startRow: 8, // 1-based row index (0-based index 7)
  columns: {
    date: "C",
    type: "D",
    itemCode: "F",
    qtyMeters: "I",
  },
  fromDate: "", // Optional ISO "YYYY-MM-DD", empty = no lower bound
  toDate: "", // Optional ISO "YYYY-MM-DD", empty = no upper bound
  itemCodes: [] as string[], // Optional item code filter, empty = all items
  includeDaily: true,
  includeWeekly: true,
  includeMonthly: true,
  timezoneOffsetHours: 7,
};

const RECO_PERIOD_HEADERS = [
  "STT",
  "Kỳ báo cáo",
  "Từ ngày",
  "Đến ngày",
  "Tổng Nhập IN (m)",
  "Tổng Xuất OUT (m)",
  "Chênh lệch Net (m)",
  "Số lần IN",
  "Số lần OUT",
  "Tổng giao dịch",
  "Số mã hàng",
];

const recoReportCleanText = (value: unknown): string => {
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

const recoReportExpandScientific = (input: string): string | null => {
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

const recoReportNormalizeItemCode = (value: unknown): string => {
  if (value === null || value === undefined || typeof value === "boolean") {
    return "";
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    if (Number.isSafeInteger(value)) return String(value);
    const str = String(Math.trunc(value));
    if (str.indexOf("e") !== -1 || str.indexOf("E") !== -1) {
      return recoReportExpandScientific(str) ?? str;
    }
    return str;
  }
  if (typeof value === "string" && value === "") return "";
  const cleaned = String(value).replace(/[\u00A0\u200B\u200C\u200D\uFEFF\s]+/g, "");
  if (!cleaned) return "";
  if (cleaned.indexOf("e") !== -1 || cleaned.indexOf("E") !== -1) {
    const expanded = recoReportExpandScientific(cleaned);
    if (expanded !== null) {
      return expanded;
    }
  }
  if (cleaned.indexOf(".") !== -1 && /^\d+\.0+$/.test(cleaned)) {
    return cleaned.replace(/\.0+$/, "");
  }
  return cleaned;
};

const recoReportNormalizeTxnType = (value: unknown): RecoReportTxnType | null => {
  if (value === "IN") return "IN";
  if (value === "OUT") return "OUT";
  const normalized = recoReportCleanText(value).replace(/\s+/g, "").toUpperCase();
  if (normalized === "IN") return "IN";
  if (normalized === "OUT") return "OUT";
  return null;
};

const recoReportParseNumber = (value: unknown): number => {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  if (typeof value === "boolean" || value === null || value === undefined) {
    return 0;
  }
  let s = recoReportCleanText(value).replace(/\s+/g, "");
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

const recoReportPad2 = (n: number): string => (n < 10 ? `0${n}` : String(n));

const recoReportToIsoDate = (
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
  return `${year}-${recoReportPad2(month)}-${recoReportPad2(day)}`;
};

const recoReportParseDate = (value: unknown): string | null => {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    return recoReportToIsoDate(
      value.getFullYear(),
      value.getMonth() + 1,
      value.getDate(),
    );
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value) || value < 1) return null;
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
    return recoReportToIsoDate(
      d.getUTCFullYear(),
      d.getUTCMonth() + 1,
      d.getUTCDate(),
    );
  }

  const text = recoReportCleanText(value);
  if (!text) return null;

  const isoMatch = text.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})(?:\b|T|\s)/);
  if (isoMatch) {
    return recoReportToIsoDate(
      Number(isoMatch[1]),
      Number(isoMatch[2]),
      Number(isoMatch[3]),
    );
  }

  const dmyMatch = text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})(?:\b|\s)/);
  if (dmyMatch) {
    return recoReportToIsoDate(
      Number(dmyMatch[3]),
      Number(dmyMatch[2]),
      Number(dmyMatch[1]),
    );
  }

  if (/^\d+(\.\d+)?$/.test(text)) {
    return recoReportParseDate(Number(text));
  }
  return null;
};

const RECO_REPORT_EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

const recoReportIsoToSerial = (iso: string): number => {
  const year = Number(iso.slice(0, 4));
  const month = Number(iso.slice(5, 7));
  const day = Number(iso.slice(8, 10));
  return Math.round(
    (Date.UTC(year, month - 1, day) - RECO_REPORT_EXCEL_EPOCH_UTC) / 86400000,
  );
};

const recoReportUtcMsToIso = (utcMs: number): string => {
  const d = new Date(utcMs);
  return `${d.getUTCFullYear()}-${recoReportPad2(d.getUTCMonth() + 1)}-${recoReportPad2(d.getUTCDate())}`;
};

const recoReportFormatDateVn = (iso: string): string => {
  const [year, month, day] = iso.split("-");
  return `${day}/${month}/${year}`;
};

const recoReportRoundQty = (n: number): number => Math.round(n * 1e6) / 1e6;

const recoReportFormatTimestamp = (offsetHours: number): string => {
  const now = new Date(Date.now() + offsetHours * 3600000);
  return `${recoReportPad2(now.getUTCDate())}/${recoReportPad2(now.getUTCMonth() + 1)}/${now.getUTCFullYear()} ${recoReportPad2(now.getUTCHours())}:${recoReportPad2(now.getUTCMinutes())}`;
};

const recoReportBuildTargetSet = (targetItemCodes: string[]): Set<string> | null => {
  if (targetItemCodes.length === 0) return null;
  const set = new Set<string>();
  for (let i = 0; i < targetItemCodes.length; i++) {
    const code = recoReportNormalizeItemCode(targetItemCodes[i]);
    if (code) set.add(code);
  }
  return set.size > 0 ? set : null;
};

/**
 * Parse contiguous C:I matrix from RECO_DATA (col 0=C Date, 1=D Type, 3=F Item Code, 6=I Qty).
 */
const recoReportParseSheetRows = (
  rows: unknown[][],
  targetItemCodes: string[] = [],
  startRow: number = RECO_PERIOD_REPORT_CONFIG.startRow,
): RecoReportTransaction[] => {
  const targetSet = recoReportBuildTargetSet(targetItemCodes);
  const transactions: RecoReportTransaction[] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    const itemCode = recoReportNormalizeItemCode(row[3]);
    if (!itemCode) continue;
    if (targetSet && !targetSet.has(itemCode)) continue;

    const type = recoReportNormalizeTxnType(row[1]);
    if (!type) continue;

    const dateIso = recoReportParseDate(row[0]);
    const qtyMeters = recoReportRoundQty(recoReportParseNumber(row[6]));

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

const recoReportParseRows = (
  columns: RecoReportRawColumns,
  targetItemCodes: string[] = [],
  startRow: number = RECO_PERIOD_REPORT_CONFIG.startRow,
): RecoReportTransaction[] => {
  const targetSet = recoReportBuildTargetSet(targetItemCodes);
  const rowCount = Math.max(
    columns.date.length,
    columns.type.length,
    columns.itemCode.length,
    columns.qtyMeters.length,
  );

  const transactions: RecoReportTransaction[] = [];
  for (let i = 0; i < rowCount; i++) {
    const itemCode = recoReportNormalizeItemCode(columns.itemCode[i]?.[0]);
    if (!itemCode) continue;
    if (targetSet && !targetSet.has(itemCode)) continue;

    const type = recoReportNormalizeTxnType(columns.type[i]?.[0]);
    if (!type) continue;

    const dateIso = recoReportParseDate(columns.date[i]?.[0]);
    const qtyMeters = recoReportRoundQty(
      recoReportParseNumber(columns.qtyMeters[i]?.[0]),
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

const recoReportFilterTransactions = (
  transactions: RecoReportTransaction[],
  options: RecoReportFilterOptions = {},
): RecoReportTransaction[] => {
  const fromIso = options.fromDate ? recoReportParseDate(options.fromDate) : null;
  const toIso = options.toDate ? recoReportParseDate(options.toDate) : null;
  const targetSet = recoReportBuildTargetSet(options.itemCodes ?? []);

  return transactions.filter((txn) => {
    if (targetSet && !targetSet.has(txn.itemCode)) return false;
    if ((fromIso || toIso) && !txn.dateIso) return false;
    if (fromIso && txn.dateIso && txn.dateIso < fromIso) return false;
    if (toIso && txn.dateIso && txn.dateIso > toIso) return false;
    return true;
  });
};

type RecoPeriodDescriptor = {
  periodKey: string;
  periodLabel: string;
  startDate: string;
  endDate: string;
};

const recoReportGetDayDescriptor = (dateIso: string): RecoPeriodDescriptor => ({
  periodKey: dateIso,
  periodLabel: recoReportFormatDateVn(dateIso),
  startDate: dateIso,
  endDate: dateIso,
});

const recoReportGetIsoWeekDescriptor = (dateIso: string): RecoPeriodDescriptor => {
  const year = Number(dateIso.slice(0, 4));
  const month = Number(dateIso.slice(5, 7));
  const day = Number(dateIso.slice(8, 10));
  const dateMs = Date.UTC(year, month - 1, day);
  const d = new Date(dateMs);
  const isoDay = d.getUTCDay() === 0 ? 7 : d.getUTCDay();

  const weekStartMs = dateMs - (isoDay - 1) * 86400000;
  const weekEndMs = weekStartMs + 6 * 86400000;
  const thursdayMs = dateMs + (4 - isoDay) * 86400000;
  const isoYear = new Date(thursdayMs).getUTCFullYear();
  const jan1Ms = Date.UTC(isoYear, 0, 1);
  const weekNo = Math.ceil(((thursdayMs - jan1Ms) / 86400000 + 1) / 7);
  const weekStr = recoReportPad2(weekNo);

  return {
    periodKey: `${isoYear}-W${weekStr}`,
    periodLabel: `Tuần ${weekStr}/${isoYear}`,
    startDate: recoReportUtcMsToIso(weekStartMs),
    endDate: recoReportUtcMsToIso(weekEndMs),
  };
};

const recoReportGetMonthDescriptor = (dateIso: string): RecoPeriodDescriptor => {
  const year = Number(dateIso.slice(0, 4));
  const month = Number(dateIso.slice(5, 7));
  const monthStr = recoReportPad2(month);
  const lastDay = new Date(Date.UTC(year, month, 0)).getUTCDate();

  return {
    periodKey: `${year}-${monthStr}`,
    periodLabel: `${monthStr}/${year}`,
    startDate: `${year}-${monthStr}-01`,
    endDate: `${year}-${monthStr}-${recoReportPad2(lastDay)}`,
  };
};

const recoReportAggregateByPeriod = (
  transactions: RecoReportTransaction[],
  describePeriod: (dateIso: string) => RecoPeriodDescriptor,
): RecoPeriodSummary[] => {
  const map = new Map<
    string,
    { summary: RecoPeriodSummary; items: Set<string> }
  >();

  for (const txn of transactions) {
    if (!txn.dateIso) continue;
    const desc = describePeriod(txn.dateIso);
    let bucket = map.get(desc.periodKey);
    if (!bucket) {
      bucket = {
        summary: {
          periodKey: desc.periodKey,
          periodLabel: desc.periodLabel,
          startDate: desc.startDate,
          endDate: desc.endDate,
          totalIn: 0,
          totalOut: 0,
          net: 0,
          inCount: 0,
          outCount: 0,
          totalCount: 0,
          uniqueItemCount: 0,
        },
        items: new Set<string>(),
      };
      map.set(desc.periodKey, bucket);
    }

    const s = bucket.summary;
    if (txn.type === "IN") {
      s.totalIn += txn.qtyMeters;
      s.inCount += 1;
    } else {
      s.totalOut += txn.qtyMeters;
      s.outCount += 1;
    }
    s.totalCount += 1;
    bucket.items.add(txn.itemCode);
  }

  const result: RecoPeriodSummary[] = [];
  map.forEach((bucket) => {
    const s = bucket.summary;
    s.totalIn = recoReportRoundQty(s.totalIn);
    s.totalOut = recoReportRoundQty(s.totalOut);
    s.net = recoReportRoundQty(s.totalIn - s.totalOut);
    s.uniqueItemCount = bucket.items.size;
    result.push(s);
  });

  return result.sort((a, b) =>
    a.periodKey < b.periodKey ? -1 : a.periodKey > b.periodKey ? 1 : 0,
  );
};

const recoReportAggregateByDay = (
  transactions: RecoReportTransaction[],
): RecoPeriodSummary[] =>
  recoReportAggregateByPeriod(transactions, recoReportGetDayDescriptor);

const recoReportAggregateByWeek = (
  transactions: RecoReportTransaction[],
): RecoPeriodSummary[] =>
  recoReportAggregateByPeriod(transactions, recoReportGetIsoWeekDescriptor);

const recoReportAggregateByMonth = (
  transactions: RecoReportTransaction[],
): RecoPeriodSummary[] =>
  recoReportAggregateByPeriod(transactions, recoReportGetMonthDescriptor);

const recoReportBuildPeriodRows = (
  periods: RecoPeriodSummary[],
  startDataRow: number,
): RecoReportCellValue[][] => {
  const rows: RecoReportCellValue[][] = periods.map((p, idx) => {
    const row = startDataRow + idx;
    return [
      idx + 1,
      p.periodLabel,
      recoReportIsoToSerial(p.startDate),
      recoReportIsoToSerial(p.endDate),
      p.totalIn,
      p.totalOut,
      `=E${row}-F${row}`,
      p.inCount,
      p.outCount,
      `=H${row}+I${row}`,
      p.uniqueItemCount,
    ];
  });

  if (periods.length > 0) {
    const endDataRow = startDataRow + periods.length - 1;
    rows.push([
      "",
      "TỔNG CỘNG",
      "",
      "",
      `=SUM(E${startDataRow}:E${endDataRow})`,
      `=SUM(F${startDataRow}:F${endDataRow})`,
      `=SUM(G${startDataRow}:G${endDataRow})`,
      `=SUM(H${startDataRow}:H${endDataRow})`,
      `=SUM(I${startDataRow}:I${endDataRow})`,
      `=SUM(J${startDataRow}:J${endDataRow})`,
      "",
    ]);
  }

  return rows;
};

const recoReportGetLastDataRow = (sheet: ExcelScript.Worksheet): number => {
  const used = sheet.getUsedRange(true);
  if (!used) return 0;
  return used.getLastRow().getRowIndex() + 1;
};

const RECO_REPORT_BORDER_ROW_LIMIT = 1000;

const recoReportApplyThinBorders = (
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
  if (rowCount <= RECO_REPORT_BORDER_ROW_LIMIT) {
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

type RecoReportSectionMeta = {
  titleRow: number;
  headerRow: number;
  startDataRow: number;
  endDataRow: number;
  totalRow: number;
  endRow: number;
  periods: RecoPeriodSummary[];
  headerColor: string;
};

const recoReportFormatSection = (
  sheet: ExcelScript.Worksheet,
  meta: RecoReportSectionMeta,
): void => {
  const titleRange = sheet.getRange(`A${meta.titleRow}:K${meta.titleRow}`);
  titleRange.merge(false);
  titleRange.getFormat().getFont().setSize(12);
  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setColor("#1F4E78");

  const tableRange = sheet.getRange(`A${meta.headerRow}:K${meta.endRow}`);
  tableRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  recoReportApplyThinBorders(tableRange, meta.endRow - meta.headerRow + 1);

  const headerRange = sheet.getRange(`A${meta.headerRow}:K${meta.headerRow}`);
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().getFont().setColor("#FFFFFF");
  headerRange.getFormat().getFill().setColor(meta.headerColor);
  headerRange.getFormat().setWrapText(true);
  headerRange.getFormat().setRowHeight(28);

  if (meta.periods.length === 0) return;

  sheet
    .getRange(`C${meta.startDataRow}:D${meta.endDataRow}`)
    .setNumberFormat("dd/mm/yyyy");

  const qtyRange = sheet.getRange(`E${meta.startDataRow}:G${meta.endRow}`);
  qtyRange.setNumberFormat("#,##0.00");
  qtyRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

  sheet
    .getRange(`H${meta.startDataRow}:K${meta.endRow}`)
    .setNumberFormat("#,##0");

  const totalRange = sheet.getRange(`A${meta.totalRow}:K${meta.totalRow}`);
  totalRange.getFormat().getFont().setBold(true);
  totalRange.getFormat().getFill().setColor("#D9E1F2");

  const netRange = sheet.getRange(`G${meta.startDataRow}:G${meta.endDataRow}`);
  netRange.getFormat().getFont().setBold(true);
  netRange.getFormat().getFill().setColor("#E2EFDA");
  netRange.getFormat().getFont().setColor("#375623");

  let negRunStart = -1;
  for (let i = 0; i <= meta.periods.length; i++) {
    const isNeg = i < meta.periods.length && meta.periods[i].net < 0;
    if (isNeg && negRunStart === -1) {
      negRunStart = meta.startDataRow + i;
    } else if (!isNeg && negRunStart !== -1) {
      const negRunEnd = meta.startDataRow + i - 1;
      const negRange = sheet.getRange(`G${negRunStart}:G${negRunEnd}`);
      negRange.getFormat().getFill().setColor("#FFC7CE");
      negRange.getFormat().getFont().setColor("#9C0006");
      negRunStart = -1;
    }
  }
};

// @ts-ignore: Office Scripts
async function main(workbook: ExcelScript.Workbook) {
  const CFG = RECO_PERIOD_REPORT_CONFIG;
  const sourceSheet = workbook.getWorksheet(CFG.sourceSheet);
  if (!sourceSheet) {
    console.log(`LỖI: Không tìm thấy sheet nguồn "${CFG.sourceSheet}".`);
    return;
  }

  let transactions: RecoReportTransaction[] = [];
  const lastSourceRow = recoReportGetLastDataRow(sourceSheet);
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
    const parsed = recoReportParseSheetRows(rawMatrix, CFG.itemCodes, CFG.startRow);
    transactions = recoReportFilterTransactions(parsed, {
      fromDate: CFG.fromDate,
      toDate: CFG.toDate,
      itemCodes: CFG.itemCodes,
    });
  }

  const missingDateCount = transactions.filter((t) => !t.dateIso).length;
  const daily = CFG.includeDaily ? recoReportAggregateByDay(transactions) : [];
  const weekly = CFG.includeWeekly ? recoReportAggregateByWeek(transactions) : [];
  const monthly = CFG.includeMonthly ? recoReportAggregateByMonth(transactions) : [];

  let targetSheet = workbook.getWorksheet(CFG.targetSheet);
  if (!targetSheet) {
    targetSheet = workbook.addWorksheet(CFG.targetSheet);
  } else {
    const usedTarget = targetSheet.getUsedRange(true);
    if (usedTarget) {
      usedTarget.unmerge();
      usedTarget.clear(ExcelScript.ClearApplyTo.all);
    }
    const autoFilter = targetSheet.getAutoFilter();
    if (autoFilter.getEnabled()) autoFilter.remove();
    targetSheet.getFreezePanes().unfreeze();
  }

  const emptyRow11 = (): RecoReportCellValue[] => [
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
  ];

  const filterParts: string[] = [];
  if (CFG.fromDate || CFG.toDate) {
    filterParts.push(`Khoảng ngày: ${CFG.fromDate || "..."} → ${CFG.toDate || "..."}`);
  }
  if (CFG.itemCodes.length > 0) {
    filterParts.push(`Mã hàng: ${CFG.itemCodes.join(", ")}`);
  }
  if (missingDateCount > 0) {
    filterParts.push(`Bỏ qua ${missingDateCount} dòng thiếu/lỗi ngày`);
  }
  const filterNote = filterParts.length > 0 ? `  |  ${filterParts.join("  |  ")}` : "";
  const subtitleText = `Thời gian: ${recoReportFormatTimestamp(CFG.timezoneOffsetHours)}  |  Nguồn: ${CFG.sourceSheet} (từ dòng ${CFG.startRow})  |  Giao dịch hợp lệ: ${transactions.length}${filterNote}`;

  const sheetMatrix: RecoReportCellValue[][] = [
    ["BÁO CÁO NHẬP - XUẤT THEO NGÀY / TUẦN / THÁNG (RECO)", "", "", "", "", "", "", "", "", "", ""],
    [subtitleText, "", "", "", "", "", "", "", "", "", ""],
  ];

  const sectionsToFormat: RecoReportSectionMeta[] = [];
  const sectionConfigs: Array<{
    enabled: boolean;
    title: string;
    periods: RecoPeriodSummary[];
    headerColor: string;
  }> = [
    {
      enabled: CFG.includeDaily,
      title: "1. BÁO CÁO THEO NGÀY (DAILY)",
      periods: daily,
      headerColor: "#1F4E78",
    },
    {
      enabled: CFG.includeWeekly,
      title: "2. BÁO CÁO THEO TUẦN ISO (WEEKLY: THỨ 2 – CHỦ NHẬT)",
      periods: weekly,
      headerColor: "#305496",
    },
    {
      enabled: CFG.includeMonthly,
      title: "3. BÁO CÁO THEO THÁNG (MONTHLY)",
      periods: monthly,
      headerColor: "#2F5597",
    },
  ];

  for (const sec of sectionConfigs) {
    if (!sec.enabled) continue;
    sheetMatrix.push(emptyRow11());
    const titleRow = sheetMatrix.length + 1;
    sheetMatrix.push([sec.title, "", "", "", "", "", "", "", "", "", ""]);
    const headerRow = sheetMatrix.length + 1;
    sheetMatrix.push([...RECO_PERIOD_HEADERS]);
    const startDataRow = sheetMatrix.length + 1;

    let endDataRow = startDataRow;
    let totalRow = 0;
    let endRow = startDataRow;

    if (sec.periods.length > 0) {
      const rows = recoReportBuildPeriodRows(sec.periods, startDataRow);
      for (let i = 0; i < rows.length; i++) {
        sheetMatrix.push(rows[i]);
      }
      endDataRow = startDataRow + sec.periods.length - 1;
      totalRow = endDataRow + 1;
      endRow = totalRow;
    } else {
      sheetMatrix.push([
        "",
        "Không có giao dịch hợp lệ trong kỳ.",
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
    }

    sectionsToFormat.push({
      titleRow,
      headerRow,
      startDataRow,
      endDataRow,
      totalRow,
      endRow,
      periods: sec.periods,
      headerColor: sec.headerColor,
    });
  }

  const finalBottomRow = sheetMatrix.length;
  const fullRange = targetSheet.getRange(`A1:K${finalBottomRow}`);
  fullRange.setValues(sheetMatrix);
  fullRange.getFormat().getFont().setName("Times New Roman");
  fullRange.getFormat().getFont().setSize(11);
  fullRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);

  const titleRange = targetSheet.getRange("A1:K1");
  titleRange.merge(false);
  titleRange.getFormat().getFont().setSize(15);
  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setColor("#1F4E78");
  titleRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  titleRange.getFormat().setRowHeight(28);

  const subRange = targetSheet.getRange("A2:K2");
  subRange.merge(false);
  subRange.getFormat().getFont().setSize(10);
  subRange.getFormat().getFont().setItalic(true);
  subRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  for (const meta of sectionsToFormat) {
    recoReportFormatSection(targetSheet, meta);
  }

  targetSheet.getRange(`A1:K${finalBottomRow}`).getFormat().autofitColumns();
  targetSheet.activate();

  console.log(
    `Auto_Reco_Period_Report hoàn tất: ${transactions.length} giao dịch -> Ngày: ${daily.length} dòng, Tuần: ${weekly.length} dòng, Tháng: ${monthly.length} dòng.`,
  );
}

const recoPeriodReportTestables = {
  RECO_PERIOD_REPORT_CONFIG,
  RECO_PERIOD_HEADERS,
  recoReportCleanText,
  recoReportNormalizeItemCode,
  recoReportNormalizeTxnType,
  recoReportParseNumber,
  recoReportParseDate,
  recoReportIsoToSerial,
  recoReportFormatDateVn,
  recoReportRoundQty,
  recoReportParseSheetRows,
  recoReportParseRows,
  recoReportFilterTransactions,
  recoReportGetDayDescriptor,
  recoReportGetIsoWeekDescriptor,
  recoReportGetMonthDescriptor,
  recoReportAggregateByDay,
  recoReportAggregateByWeek,
  recoReportAggregateByMonth,
  recoReportBuildPeriodRows,
  recoReportGetLastDataRow,
};

(
  globalThis as typeof globalThis & {
    __AUTO_RECO_PERIOD_REPORT_TESTABLES__?: typeof recoPeriodReportTestables;
  }
).__AUTO_RECO_PERIOD_REPORT_TESTABLES__ = recoPeriodReportTestables;
