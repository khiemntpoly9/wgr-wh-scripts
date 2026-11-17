/**
 * Auto_Zoro_Super – lấy cấu trúc báo cáo tồn kho theo thị trường từ backend.
 * Không đọc sheet "Thành phẩm" hoặc sheet thông số trong workbook.
 */

type ZoroSuperCellValue = string | number | boolean;
type ZoroSuperStyle = {
  address: string;
  font?: { name?: string; size?: number; bold?: boolean; color?: string };
  fill?: string;
  horizontal?: "left" | "center" | "right";
  vertical?: "top" | "center" | "bottom";
  wrap_text?: boolean;
  number_format?: string;
};
type ZoroSuperSheetPlan = {
  name: string;
  clear_existing: boolean;
  writes: { address: string; values: ZoroSuperCellValue[][] }[];
  formulas: { address: string; formula: string }[];
  styles: ZoroSuperStyle[];
  border_ranges: string[];
  merges: string[];
  auto_filter?: string | null;
  freeze_rows?: number;
  autofit_columns?: string;
  autofit_rows?: string;
  minimum_row_height?: number;
  summary?: {
    market_code?: string;
    location_codes?: string[];
    total_lines?: number;
    total_stock?: number;
  };
};
type ZoroSuperSummary = {
  total_sheets?: number;
  total_lines?: number;
  report_date?: string;
  location_codes?: string[];
  ignored_location_codes?: string[];
};
type ZoroSuperResponse = {
  format?: string;
  source?: string;
  sheets?: ZoroSuperSheetPlan[];
  summary?: ZoroSuperSummary;
  message?: string;
  errors?: Record<string, string[]>;
};

const ZORO_SUPER_CONFIG = {
  apiUrl: "https://app.wgr-wh.site/api/v1/warehouse/auto-zoro-super/preview",
  apiToken: "27|c0tmgUTUi6KOtS8wKzk3xud5Og6QDPNVEsZZ74qMd6fe3f96", // Bearer token Sanctum có quyền xem tồn kho.
  marketCodes: [] as string[],
  reportDate: "", // yyyy-mm-dd; để trống = ngày hiện tại của backend.
  sheetNamePrefix: "Tồn kho",
  receiptDateFilter: {
    days: [] as number[],
    months: [] as number[],
    years: [] as number[],
  },
  /** Mã vị trí cần lấy toàn bộ tồn (vd ["E23","E24"]), bất kể mã hợp đồng. Để trống = tắt. */
  locationCodes: [] as string[],
  /** Mã SKU cần dò (vd "936771"): log các dòng của SKU này trên từng sheet. Để trống = tắt. */
  debugSku: "",
};
const ZORO_SUPER_SO_DONG_MOI_KHOI = 1000;

const zoroSuperKiemTraUrl = (url: string): string | null => {
  const value = url.trim().toLowerCase();
  if (!value) return "Chưa cấu hình apiUrl.";
  if (!value.startsWith("https://")) return "apiUrl phải dùng HTTPS.";
  if (value.includes("localhost") || value.includes("127.0.0.1"))
    return "apiUrl không được dùng localhost.";
  return null;
};
const zoroSuperLaNgayIso = (value: string): boolean => {
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return false;
  const date = new Date(Date.UTC(+match[1], +match[2] - 1, +match[3]));
  return (
    date.getUTCFullYear() === +match[1] &&
    date.getUTCMonth() + 1 === +match[2] &&
    date.getUTCDate() === +match[3]
  );
};
const zoroSuperTaoPayload = (config: typeof ZORO_SUPER_CONFIG) => {
  const locationCodes = (config.locationCodes ?? [])
    .map((value) => value.trim())
    .filter((value) => value !== "");
  return {
    market_codes: config.marketCodes
      .map((value) => value.trim().toUpperCase())
      .filter((value) => value !== ""),
    ...(config.reportDate.trim()
      ? { report_date: config.reportDate.trim() }
      : {}),
    sheet_name_prefix: config.sheetNamePrefix,
    receipt_date_filter: {
      days: config.receiptDateFilter.days,
      months: config.receiptDateFilter.months,
      years: config.receiptDateFilter.years,
    },
    ...(locationCodes.length > 0 ? { location_codes: locationCodes } : {}),
  };
};
const zoroSuperMoTaViTri = (summary?: ZoroSuperSummary): string[] => {
  if (!summary || !Array.isArray(summary.location_codes)) return [];
  const lines = [
    `Vị trí lọc: ${summary.location_codes.length > 0 ? summary.location_codes.join(", ") : "(không có mã hợp lệ)"}`,
  ];
  if ((summary.ignored_location_codes ?? []).length > 0)
    lines.push(
      `Bỏ qua mã vị trí không hợp lệ: ${(summary.ignored_location_codes ?? []).join(", ")}`,
    );
  return lines;
};
const zoroSuperKeVien = (range: ExcelScript.Range): void => {
  const format = range.getFormat();
  [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal,
    ExcelScript.BorderIndex.insideVertical,
  ].forEach((index) =>
    format
      .getRangeBorder(index)
      .setStyle(ExcelScript.BorderLineStyle.continuous),
  );
};
/** Parse "A12" → { col: "A", row: 12 }. */
const zoroSuperParseAddress = (
  address: string,
): { col: string; row: number } | null => {
  const match = address.trim().match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return null;
  return { col: match[1].toUpperCase(), row: Number(match[2]) };
};
/**
 * Ghi công thức theo khối liên tục (setFormulas) thay vì setFormula từng ô.
 * Office Scripts: N lần getRange/setFormula trên sheet lớn dễ chạy vài phút / timeout.
 */
const zoroSuperGhiFormulas = (
  sheet: ExcelScript.Worksheet,
  formulas: { address: string; formula: string }[],
): void => {
  if (formulas.length === 0) return;
  type Pending = { col: string; startRow: number; formulas: string[] };
  type Cell = { row: number; formula: string };
  const byCol = new Map<string, Cell[]>();
  const singles: { address: string; formula: string }[] = [];
  for (const item of formulas) {
    const parsed = zoroSuperParseAddress(item.address);
    if (!parsed) {
      singles.push(item);
      continue;
    }
    const list = byCol.get(parsed.col) || [];
    list.push({ row: parsed.row, formula: item.formula });
    byCol.set(parsed.col, list);
  }
  byCol.forEach((cells: Cell[], col: string) => {
    cells.sort((a: Cell, b: Cell) => a.row - b.row);
    let pending: Pending | null = null;
    const flush = () => {
      if (!pending || pending.formulas.length === 0) return;
      const endRow = pending.startRow + pending.formulas.length - 1;
      const matrix = pending.formulas.map((formula) => [formula]);
      sheet
        .getRange(`${pending.col}${pending.startRow}:${pending.col}${endRow}`)
        .setFormulas(matrix);
      pending = null;
    };
    for (const cell of cells) {
      if (pending && cell.row === pending.startRow + pending.formulas.length) {
        pending.formulas.push(cell.formula);
      } else {
        flush();
        pending = { col, startRow: cell.row, formulas: [cell.formula] };
      }
    }
    flush();
  });
  for (const item of singles)
    sheet.getRange(item.address).setFormula(item.formula);
};
/** Ghi values theo khối dòng — một setValues hàng nghìn dòng dễ timeout trên Excel Online. */
const zoroSuperGhiGiaTri = (
  sheet: ExcelScript.Worksheet,
  address: string,
  values: ZoroSuperCellValue[][],
  soDongMoiKhoi: number = ZORO_SUPER_SO_DONG_MOI_KHOI,
): void => {
  const match = address.trim().match(/^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/);
  if (!match || values.length <= soDongMoiKhoi) {
    sheet.getRange(address).setValues(values);
    return;
  }
  const startRow = Number(match[2]);
  for (let offset = 0; offset < values.length; offset += soDongMoiKhoi) {
    const block = values.slice(offset, offset + soDongMoiKhoi);
    const fromRow = startRow + offset;
    sheet
      .getRange(`${match[1]}${fromRow}:${match[3]}${fromRow + block.length - 1}`)
      .setValues(block);
  }
};
const zoroSuperTimDongSku = (
  plan: ZoroSuperSheetPlan,
  debugSku: string,
): string[] => {
  const sku = debugSku.trim();
  if (!sku) return [];
  const result: string[] = [];
  for (const write of plan.writes) {
    const start = zoroSuperParseAddress(write.address.split(":")[0]);
    const startRow = start ? start.row : 1;
    write.values.forEach((row: ZoroSuperCellValue[], index: number) => {
      if (row.some((cell: ZoroSuperCellValue) => String(cell).trim() === sku))
        result.push(`dòng ${startRow + index}: ${row.join(" | ")}`);
    });
  }
  return result;
};
const zoroSuperApDungStyle = (
  sheet: ExcelScript.Worksheet,
  style: ZoroSuperStyle,
): void => {
  const range = sheet.getRange(style.address);
  const format = range.getFormat();
  if (style.font?.name) format.getFont().setName(style.font.name);
  if (style.font?.size !== undefined) format.getFont().setSize(style.font.size);
  if (style.font?.bold !== undefined) format.getFont().setBold(style.font.bold);
  if (style.font?.color) format.getFont().setColor(style.font.color);
  if (style.fill) format.getFill().setColor(style.fill);
  if (style.horizontal === "left")
    format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  if (style.horizontal === "center")
    format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  if (style.horizontal === "right")
    format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  if (style.vertical === "top")
    format.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  if (style.vertical === "center")
    format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  if (style.vertical === "bottom")
    format.setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
  if (style.wrap_text !== undefined) format.setWrapText(style.wrap_text);
  if (style.number_format) range.setNumberFormat(style.number_format);
};
const zoroSuperKiemTraKeHoach = (
  body: ZoroSuperResponse,
): body is ZoroSuperResponse & { sheets: ZoroSuperSheetPlan[] } =>
  body.format === "auto_zoro_backend_v1" && Array.isArray(body.sheets);
const zoroSuperLoiValidation = (errors?: Record<string, string[]>): string =>
  errors
    ? Object.keys(errors)
        .map((field) => `${field}: ${errors[field].join(", ")}`)
        .join(" | ")
    : "";

const zoroSuperGhiMotSheet = (
  workbook: ExcelScript.Workbook,
  plan: ZoroSuperSheetPlan,
): void => {
  let sheet = workbook.getWorksheet(plan.name);
  if (!sheet) sheet = workbook.addWorksheet(plan.name);
  if (plan.clear_existing) {
    const used = sheet.getUsedRange();
    if (used) {
      used.unmerge();
      used.clear(ExcelScript.ClearApplyTo.all);
    }
  }
  for (const write of plan.writes)
    zoroSuperGhiGiaTri(sheet, write.address, write.values);
  zoroSuperGhiFormulas(sheet, plan.formulas);
  for (const style of plan.styles) zoroSuperApDungStyle(sheet, style);
  for (const address of plan.border_ranges)
    zoroSuperKeVien(sheet.getRange(address));
  for (const address of plan.merges) sheet.getRange(address).merge(false);
  const filter = sheet.getAutoFilter();
  if (filter.getEnabled()) filter.remove();
  if (plan.auto_filter) filter.apply(plan.auto_filter);
  const freeze = sheet.getFreezePanes();
  freeze.unfreeze();
  if ((plan.freeze_rows || 0) > 0) freeze.freezeRows(plan.freeze_rows || 0);
  // autofitColumns nhận "A:D,F:K" — gọi từng đoạn để tránh autofit cả cột text dài.
  if (plan.autofit_columns) {
    for (const part of plan.autofit_columns.split(",")) {
      const address = part.trim();
      if (address) sheet.getRange(address).getFormat().autofitColumns();
    }
  }
  // Tránh autofitRows + vòng lặp từng dòng (rất chậm trên sheet lớn → script "không ngừng").
  if (plan.minimum_row_height && plan.autofit_rows) {
    sheet
      .getRange(plan.autofit_rows)
      .getFormat()
      .setRowHeight(plan.minimum_row_height);
  } else if (plan.autofit_rows) {
    sheet.getRange(plan.autofit_rows).getFormat().autofitRows();
  }
  sheet.activate();
  // Đọc 1 ô để ép đồng bộ: lỗi ghi của sheet này nổi lên tại đây thay vì ở sheet sau.
  sheet.getRange("A1").getValues();
};
const zoroSuperGhiCacSheet = (
  workbook: ExcelScript.Workbook,
  plans: ZoroSuperSheetPlan[],
  debugSku: string,
): { tomTat: string[]; loi: string[]; debug: string[] } => {
  const tomTat: string[] = [];
  const loi: string[] = [];
  const debug: string[] = [];
  const sku = debugSku.trim();
  for (const plan of plans) {
    if (sku) {
      const dong = zoroSuperTimDongSku(plan, sku);
      if (dong.length === 0)
        debug.push(`DEBUG SKU ${sku} @ "${plan.name}": không có dòng.`);
      for (const line of dong)
        debug.push(`DEBUG SKU ${sku} @ "${plan.name}": ${line}`);
    }
    try {
      zoroSuperGhiMotSheet(workbook, plan);
      tomTat.push(
        `${plan.summary?.market_code || plan.name}: ${plan.summary?.total_lines || 0} dòng, tồn ${plan.summary?.total_stock || 0}.`,
      );
    } catch (error) {
      loi.push(`"${plan.name}": ${String(error)}`);
    }
  }
  return { tomTat, loi, debug };
};

declare function fetch(
  input: string,
  init?: { method?: string; headers?: Record<string, string>; body?: string },
): Promise<{ ok: boolean; status: number; text(): Promise<string> }>;

async function main(workbook: ExcelScript.Workbook) {
  const CFG = ZORO_SUPER_CONFIG;
  const urlError = zoroSuperKiemTraUrl(CFG.apiUrl);
  if (urlError) return console.log(`LỖI cấu hình: ${urlError}`);
  if (!CFG.apiToken.trim()) return console.log("LỖI: Chưa cấu hình apiToken.");
  if (CFG.marketCodes.length === 0)
    return console.log("LỖI: marketCodes phải có ít nhất một thị trường.");
  if (CFG.reportDate.trim() && !zoroSuperLaNgayIso(CFG.reportDate.trim()))
    return console.log("LỖI: reportDate phải có định dạng yyyy-mm-dd.");
  const startedAt = Date.now();
  try {
    const response = await fetch(CFG.apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Authorization: `Bearer ${CFG.apiToken}`,
      },
      body: JSON.stringify(zoroSuperTaoPayload(CFG)),
    });
    const text = await response.text();
    let body: ZoroSuperResponse;
    try {
      body = JSON.parse(text) as ZoroSuperResponse;
    } catch {
      body = { message: text };
    }
    if (!response.ok) {
      const validation = zoroSuperLoiValidation(body.errors);
      return console.log(
        `LỖI HTTP ${response.status}: ${body.message || text}${validation ? ` — ${validation}` : ""}`,
      );
    }
    if (!zoroSuperKiemTraKeHoach(body))
      return console.log("LỖI: Backend trả về sai format Auto_Zoro.");

    const ketQua = zoroSuperGhiCacSheet(workbook, body.sheets, CFG.debugSku);
    const viTri = zoroSuperMoTaViTri(body.summary);
    if (viTri.length > 0) console.log(viTri.join("\n"));
    if (ketQua.debug.length > 0) console.log(ketQua.debug.join("\n"));
    if (ketQua.tomTat.length > 0) console.log(ketQua.tomTat.join("\n"));
    if (ketQua.loi.length > 0)
      console.log(
        `LỖI ghi ${ketQua.loi.length} sheet (các sheet khác vẫn được ghi):\n${ketQua.loi.join("\n")}`,
      );
    console.log(
      `Auto_Zoro_Super hoàn tất ${body.summary?.total_sheets ?? body.sheets.length} sheet trong ${Math.round((Date.now() - startedAt) / 1000)}s.`,
    );
  } catch (error) {
    const message = String(error);
    console.log(
      message.includes("Failed to fetch")
        ? `LỖI kết nối API — kiểm tra deploy, HTTPS và CORS: ${CFG.apiUrl}`
        : `LỖI Auto_Zoro_Super: ${message}`,
    );
  }
}

const zoroSuperTestables = {
  zoroSuperKiemTraUrl,
  zoroSuperLaNgayIso,
  zoroSuperTaoPayload,
  zoroSuperKiemTraKeHoach,
  zoroSuperParseAddress,
  zoroSuperGhiFormulas,
  zoroSuperGhiGiaTri,
  zoroSuperTimDongSku,
  zoroSuperGhiCacSheet,
  zoroSuperMoTaViTri,
};
(
  globalThis as typeof globalThis & {
    __AUTO_ZORO_SUPER_TESTABLES__?: typeof zoroSuperTestables;
  }
).__AUTO_ZORO_SUPER_TESTABLES__ = zoroSuperTestables;
