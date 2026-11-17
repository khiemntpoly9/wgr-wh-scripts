/** Auto_Sanji_Super – render báo cáo tổng hợp và chi tiết tồn kho do backend dựng sẵn. */
type SanjiSuperCellValue = string | number | boolean;
type SanjiSuperStyle = {
  address: string;
  font?: { name?: string; size?: number; bold?: boolean; color?: string };
  fill?: string;
  horizontal?: "left" | "center" | "right";
  vertical?: "top" | "center" | "bottom";
  wrap_text?: boolean;
  number_format?: string;
};
type SanjiSuperSheetPlan = {
  name: string;
  clear_existing: boolean;
  writes: { address: string; values: SanjiSuperCellValue[][] }[];
  formulas: { address: string; formula: string }[];
  styles: SanjiSuperStyle[];
  border_ranges: string[];
  merges: string[];
  auto_filter?: string | null;
  freeze_rows?: number;
  autofit_columns?: string;
  summary?: {
    total_lines?: number;
    total_markets?: number;
    total_stock?: number;
  };
};
type SanjiSuperResponse = {
  format?: string;
  source?: string;
  sheet?: SanjiSuperSheetPlan | null;
  summary?: {
    total_lines?: number;
    total_markets?: number;
    total_stock?: number;
  };
  message?: string;
  errors?: Record<string, string[]>;
};
const SANJI_SUPER_CONFIG = {
  apiUrl: "https://app.wgr-wh.site/api/v1/warehouse/auto-sanji-super/preview",
  apiToken: "1|hCwtksIBz4Yc7g2wn6aeMO4L6zri5rStWeSYgH9rf7ad4909", // Bearer token Sanctum có quyền xem tồn kho.
  marketCodes: [] as string[], // [] = tất cả thị trường.
  sheetName: "Tồn kho",
  classificationOrder: ["F", "A", "S", "C", "T", "M"] as string[],
  classificationOrderByMarket: {} as Record<string, string[]>,
};
const sanjiSuperKiemTraUrl = (url: string): string | null => {
  const value = url.trim().toLowerCase();
  if (!value) return "Chưa cấu hình apiUrl.";
  if (!value.startsWith("https://")) return "apiUrl phải dùng HTTPS.";
  if (value.includes("localhost") || value.includes("127.0.0.1"))
    return "apiUrl không được dùng localhost.";
  return null;
};
const sanjiSuperTaoPayload = (config: typeof SANJI_SUPER_CONFIG) => ({
  market_codes: config.marketCodes
    .map((value) => value.trim().toUpperCase())
    .filter((value) => value !== ""),
  sheet_name: config.sheetName,
  classification_order: config.classificationOrder,
  classification_order_by_market: config.classificationOrderByMarket,
});
const sanjiSuperKeVien = (range: ExcelScript.Range): void => {
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
const sanjiSuperApDungStyle = (
  sheet: ExcelScript.Worksheet,
  style: SanjiSuperStyle,
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
const sanjiSuperKiemTraKeHoach = (body: SanjiSuperResponse): boolean =>
  body.format === "auto_sanji_backend_v1" &&
  (body.sheet === null ||
    (!!body.sheet &&
      Array.isArray(body.sheet.writes) &&
      Array.isArray(body.sheet.formulas)));
const sanjiSuperLoiValidation = (errors?: Record<string, string[]>): string =>
  errors
    ? Object.keys(errors)
        .map((field) => `${field}: ${errors[field].join(", ")}`)
        .join(" | ")
    : "";

declare function fetch(
  input: string,
  init?: { method?: string; headers?: Record<string, string>; body?: string },
): Promise<{ ok: boolean; status: number; text(): Promise<string> }>;

async function main(workbook: ExcelScript.Workbook) {
  const CFG = SANJI_SUPER_CONFIG;
  const urlError = sanjiSuperKiemTraUrl(CFG.apiUrl);
  if (urlError) return console.log(`LỖI cấu hình: ${urlError}`);
  if (!CFG.apiToken.trim()) return console.log("LỖI: Chưa cấu hình apiToken.");
  const startedAt = Date.now();
  try {
    const response = await fetch(CFG.apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Authorization: `Bearer ${CFG.apiToken}`,
      },
      body: JSON.stringify(sanjiSuperTaoPayload(CFG)),
    });
    const text = await response.text();
    let body: SanjiSuperResponse;
    try {
      body = JSON.parse(text) as SanjiSuperResponse;
    } catch {
      body = { message: text };
    }
    if (!response.ok) {
      const validation = sanjiSuperLoiValidation(body.errors);
      return console.log(
        `LỖI HTTP ${response.status}: ${body.message || text}${validation ? ` — ${validation}` : ""}`,
      );
    }
    if (!sanjiSuperKiemTraKeHoach(body))
      return console.log("LỖI: Backend trả về sai format Auto_Sanji.");
    if (!body.sheet)
      return console.log(
        "Không có dữ liệu tồn kho dương thỏa điều kiện để thống kê.",
      );
    const plan = body.sheet;
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
      sheet.getRange(write.address).setValues(write.values);
    for (const formula of plan.formulas)
      sheet.getRange(formula.address).setFormula(formula.formula);
    for (const style of plan.styles) sanjiSuperApDungStyle(sheet, style);
    for (const address of plan.border_ranges)
      sanjiSuperKeVien(sheet.getRange(address));
    for (const address of plan.merges) sheet.getRange(address).merge(false);
    const filter = sheet.getAutoFilter();
    if (filter.getEnabled()) filter.remove();
    if (plan.auto_filter) filter.apply(plan.auto_filter);
    const freeze = sheet.getFreezePanes();
    freeze.unfreeze();
    if ((plan.freeze_rows || 0) > 0) freeze.freezeRows(plan.freeze_rows || 0);
    if (plan.autofit_columns)
      sheet.getRange(plan.autofit_columns).getFormat().autofitColumns();
    sheet.activate();
    console.log(
      `Auto_Sanji_Super hoàn tất: ${plan.summary?.total_lines ?? 0} dòng, ${plan.summary?.total_markets ?? 0} thị trường, tổng tồn ${plan.summary?.total_stock ?? 0}. Thời gian ${Math.round((Date.now() - startedAt) / 1000)}s.`,
    );
  } catch (error) {
    const message = String(error);
    console.log(
      message.includes("Failed to fetch")
        ? `LỖI kết nối API — kiểm tra deploy, HTTPS và CORS: ${CFG.apiUrl}`
        : `LỖI Auto_Sanji_Super: ${message}`,
    );
  }
}

const sanjiSuperTestables = {
  sanjiSuperKiemTraUrl,
  sanjiSuperTaoPayload,
  sanjiSuperKiemTraKeHoach,
};
(
  globalThis as typeof globalThis & {
    __AUTO_SANJI_SUPER_TESTABLES__?: typeof sanjiSuperTestables;
  }
).__AUTO_SANJI_SUPER_TESTABLES__ = sanjiSuperTestables;
