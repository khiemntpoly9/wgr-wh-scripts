/**
 * Auto_Nami_Supper – nhận toàn bộ cấu trúc sheet Auto_Nami từ backend và render vào Excel.
 *
 * Backend quyết định dữ liệu, ô, công thức, style, merge, filter và freeze panes.
 * Script không đọc sheet "Thành phẩm" hay sheet thông số.
 * Dùng sau Upload_Nami_Super để xem/đối soát báo cáo nhập kho của kỳ vừa tải lên.
 */

type NamiSupperFilterMode = "day" | "week" | "month" | "year";
type NamiSupperCellValue = string | number | boolean;

const NAMI_SUPPER_CONFIG = {
  apiUrl: "https://app.wgr-wh.site/api/v1/warehouse/auto-nami-supper/preview",
  apiToken: "1|hCwtksIBz4Yc7g2wn6aeMO4L6zri5rStWeSYgH9rf7ad4909", // Bearer token Sanctum từ Dashboard → API Token (Excel).
  targetDate: "2026-09-13", // yyyy-mm-dd
  filterMode: "day" as NamiSupperFilterMode,
  volumeDivisor: 68,
};

type NamiSupperFont = {
  name?: string;
  size?: number;
  bold?: boolean;
  color?: string;
};

type NamiSupperInstruction = {
  address: string;
  values?: NamiSupperCellValue[][];
  style?:
    | "statistics_title"
    | "statistics_header"
    | "statistics_data"
    | "statistics_border";
  number_format?: string;
};

type NamiSupperStyle = NamiSupperInstruction & {
  font?: NamiSupperFont;
  fill?: string;
  horizontal?: "left" | "center" | "right";
  vertical?: "top" | "center" | "bottom";
  wrap_text?: boolean;
};

type NamiSupperSheet = {
  name: string;
  clear_existing: boolean;
  writes: NamiSupperInstruction[];
  formulas: { address: string; formula: string }[];
  styles: NamiSupperStyle[];
  border_ranges: string[];
  merges: string[];
  auto_filter: string;
  freeze_rows: number;
  autofit?: string;
  columns?: { address: string; width: number }[];
  last_data_row: number;
  filter_label: string;
  summary?: {
    total_lines?: number;
    total_quantity?: number;
    date_from?: string;
    date_to?: string;
  };
};

type NamiSupperResponse = {
  format?: string;
  source?: string;
  sheet?: NamiSupperSheet;
  message?: string;
  errors?: Record<string, string[]>;
};

const namiSupperKiemTraUrl = (url: string): string | null => {
  const value = url.trim().toLowerCase();
  if (!value) return "Chưa cấu hình apiUrl.";
  if (!value.startsWith("https://")) return "apiUrl phải dùng HTTPS.";
  if (value.includes("localhost") || value.includes("127.0.0.1")) {
    return "apiUrl không được dùng localhost vì Office Scripts chạy trên Microsoft cloud.";
  }
  return null;
};

const namiSupperKiemTraNgay = (value: string): boolean => {
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return false;
  const date = new Date(
    Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])),
  );
  return (
    date.getUTCFullYear() === Number(match[1]) &&
    date.getUTCMonth() + 1 === Number(match[2]) &&
    date.getUTCDate() === Number(match[3])
  );
};

const namiSupperLoiValidation = (errors?: Record<string, string[]>): string =>
  errors
    ? Object.keys(errors)
        .map((field) => `${field}: ${errors[field].join(", ")}`)
        .join(" | ")
    : "";

const namiSupperKeVien = (range: ExcelScript.Range): void => {
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

const namiSupperCanNgang = (
  value?: "left" | "center" | "right",
): ExcelScript.HorizontalAlignment | undefined => {
  if (value === "left") return ExcelScript.HorizontalAlignment.left;
  if (value === "right") return ExcelScript.HorizontalAlignment.right;
  if (value === "center") return ExcelScript.HorizontalAlignment.center;
  return undefined;
};

const namiSupperCanDoc = (
  value?: "top" | "center" | "bottom",
): ExcelScript.VerticalAlignment | undefined => {
  if (value === "top") return ExcelScript.VerticalAlignment.top;
  if (value === "bottom") return ExcelScript.VerticalAlignment.bottom;
  if (value === "center") return ExcelScript.VerticalAlignment.center;
  return undefined;
};

const namiSupperApDungStyle = (
  sheet: ExcelScript.Worksheet,
  instruction: NamiSupperStyle,
): void => {
  const range = sheet.getRange(instruction.address);
  const format = range.getFormat();

  if (instruction.style === "statistics_title") {
    instruction = {
      ...instruction,
      font: { name: "Times New Roman", size: 14, bold: true },
      fill: "#FAB0B5",
      horizontal: "center",
    };
  } else if (instruction.style === "statistics_header") {
    instruction = {
      ...instruction,
      font: { name: "Times New Roman", size: 14, bold: true, color: "#FFFFFF" },
      fill: "#4F81BD",
      horizontal: "center",
      vertical: "center",
    };
  } else if (instruction.style === "statistics_data") {
    instruction = {
      ...instruction,
      font: { name: "Times New Roman", size: 14 },
      horizontal: "center",
    };
  }

  if (instruction.style === "statistics_border") namiSupperKeVien(range);
  if (instruction.font?.name) format.getFont().setName(instruction.font.name);
  if (instruction.font?.size !== undefined)
    format.getFont().setSize(instruction.font.size);
  if (instruction.font?.bold !== undefined)
    format.getFont().setBold(instruction.font.bold);
  if (instruction.font?.color)
    format.getFont().setColor(instruction.font.color);
  if (instruction.fill) format.getFill().setColor(instruction.fill);
  const horizontal = namiSupperCanNgang(instruction.horizontal);
  if (horizontal !== undefined) format.setHorizontalAlignment(horizontal);
  const vertical = namiSupperCanDoc(instruction.vertical);
  if (vertical !== undefined) format.setVerticalAlignment(vertical);
  if (instruction.wrap_text !== undefined)
    format.setWrapText(instruction.wrap_text);
  if (instruction.number_format)
    range.setNumberFormat(instruction.number_format);
};

// @ts-ignore: Office Scripts cung cấp fetch ở runtime.
declare function fetch(
  input: string,
  init?: {
    method?: string;
    headers?: { [key: string]: string };
    body?: string;
  },
): Promise<{ ok: boolean; status: number; text(): Promise<string> }>;

// @ts-ignore: Mỗi Office Script chạy độc lập nên có thể dùng tên main.
async function main(workbook: ExcelScript.Workbook) {
  const CFG = NAMI_SUPPER_CONFIG;
  const urlError = namiSupperKiemTraUrl(CFG.apiUrl);
  if (urlError) {
    console.log(`LỖI cấu hình: ${urlError}`);
    return;
  }
  if (!CFG.apiToken.trim()) {
    console.log(
      "LỖI: Chưa cấu hình apiToken (Bearer token Sanctum). Hãy tạo token tại Dashboard → API Token (Excel).",
    );
    return;
  }
  if (!namiSupperKiemTraNgay(CFG.targetDate)) {
    console.log(
      "LỖI: targetDate phải là ngày hợp lệ theo định dạng yyyy-mm-dd.",
    );
    return;
  }

  console.log(
    `Đang yêu cầu backend dựng báo cáo Auto_Nami cho ${CFG.targetDate}...`,
  );
  const startedAt = Date.now();
  try {
    const response = await fetch(CFG.apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Authorization: `Bearer ${CFG.apiToken}`,
      },
      body: JSON.stringify({
        target_date: CFG.targetDate,
        filter_mode: CFG.filterMode,
        volume_divisor: CFG.volumeDivisor,
      }),
    });
    const bodyText = await response.text();
    let body: NamiSupperResponse;
    try {
      body = JSON.parse(bodyText) as NamiSupperResponse;
    } catch {
      body = { message: bodyText };
    }

    if (!response.ok) {
      const validation = namiSupperLoiValidation(body.errors);
      console.log(
        `LỖI HTTP ${response.status}: ${body.message || bodyText}${validation ? ` — ${validation}` : ""}`,
      );
      return;
    }
    if (body.format !== "auto_nami_backend_v1" || !body.sheet) {
      console.log("LỖI: Backend trả về sai format Auto_Nami.");
      return;
    }

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

    for (const write of plan.writes) {
      if (write.values) sheet.getRange(write.address).setValues(write.values);
    }
    for (const formula of plan.formulas) {
      sheet.getRange(formula.address).setFormula(formula.formula);
    }
    for (const style of plan.styles) namiSupperApDungStyle(sheet, style);
    for (const address of plan.border_ranges)
      namiSupperKeVien(sheet.getRange(address));
    for (const address of plan.merges) sheet.getRange(address).merge(false);

    const autoFilter = sheet.getAutoFilter();
    if (autoFilter.getEnabled()) autoFilter.remove();
    autoFilter.apply(plan.auto_filter);
    const freeze = sheet.getFreezePanes();
    freeze.unfreeze();
    freeze.freezeRows(plan.freeze_rows);
    if (plan.autofit) sheet.getRange(plan.autofit).getFormat().autofitColumns();
    for (const column of plan.columns || []) {
      sheet.getRange(column.address).getFormat().setColumnWidth(column.width);
    }
    sheet.activate();

    console.log("=========================================");
    console.log("  HOÀN TẤT AUTO_NAMI_SUPPER TỪ BACKEND");
    console.log("=========================================");
    console.log(`• Sheet: ${plan.name}`);
    console.log(`• Phạm vi: ${plan.filter_label}`);
    console.log(
      `• Tổng dòng: ${plan.summary?.total_lines ?? 0} | Tổng SL nhập: ${plan.summary?.total_quantity ?? 0}`,
    );
    console.log(`• Thời gian: ${Math.round((Date.now() - startedAt) / 1000)}s`);
  } catch (error) {
    const message = String(error);
    console.log(
      message.includes("Failed to fetch")
        ? `LỖI kết nối API: Failed to fetch — kiểm tra URL HTTPS public, deploy backend và CORS. API: ${CFG.apiUrl}`
        : `LỖI Auto_Nami_Supper: ${message}`,
    );
  }
}

const namiSupperTestables = { namiSupperKiemTraUrl, namiSupperKiemTraNgay };
(
  globalThis as typeof globalThis & {
    __AUTO_NAMI_SUPPER_TESTABLES__?: typeof namiSupperTestables;
  }
).__AUTO_NAMI_SUPPER_TESTABLES__ = namiSupperTestables;
