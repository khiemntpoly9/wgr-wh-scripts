/**
 * ===================================================================
 * Auto_Mitsuha_API – Lấy đề xuất soạn hàng từ backend và ghi A→Z.
 * ===================================================================
 *
 * Khác Auto_Mitsuha.ts:
 * - Không đọc sheet "Thành phẩm", sheet thông số hoặc Pre-packing-list.
 * - Backend tính tồn, pending, ngưỡng và cặp A/B theo cấu hình đã lưu trên web.
 * - Giữ layout output tương thích Auto_Mitsuha, thêm Invoice tại cột T và Mã đơn hàng tại cột U.
 *
 * Chuẩn bị:
 * 1. Deploy backend có endpoint POST /api/v1/warehouse/auto-mitsuha-super/preview.
 * 2. Tạo token tại Dashboard → API Token (Excel).
 * 3. Chỉnh bộ lọc và định dạng sheet tại Dashboard → Danh sách tủ tải trước → Cấu hình soạn hàng.
 * 4. Chạy script trong Excel Online.
 */

const MITSUHA_API_CONFIG = {
  apiUrl:
    "https://app.wgr-wh.site/api/v1/warehouse/auto-mitsuha-super/preview",
  apiToken: "1|hCwtksIBz4Yc7g2wn6aeMO4L6zri5rStWeSYgH9rf7ad4909", // Bearer token Sanctum từ Dashboard → API Token (Excel)
};

type MitsuhaApiExcelOutput = {
  sheet_name: string;
  clear_old_data: boolean;
  enable_excel_filter: boolean;
  start_row: number;
  warning_fill_color: string;
  completed_with_pending_fill_color: string;
  dimension_fallback_fill_color: string;
  highlight_duplicate_sku: boolean;
  duplicate_sku_fill_color: string;
};

/** Dùng khi backend cũ chưa trả về excel_output. */
const MITSUHA_API_EXCEL_OUTPUT_DEFAULTS: MitsuhaApiExcelOutput = {
  sheet_name: "",
  clear_old_data: true,
  enable_excel_filter: true,
  start_row: 2,
  warning_fill_color: "#FFF2CC",
  completed_with_pending_fill_color: "#E2F0D9",
  dimension_fallback_fill_color: "#DDEBF7",
  highlight_duplicate_sku: true,
  duplicate_sku_fill_color: "#D9E1F2",
};

type MitsuhaApiPlannerItem = {
  stt: number;
  inventory_id: number;
  product_id: number;
  contract_id: number;
  location_id: number;
  contract_no: string;
  market_code: string;
  classification_code: string;
  sku: string;
  pair_key: string;
  base_sku: string;
  product_name: string;
  dimensions: string;
  order_quantity: number;
  unit: string;
  previously_shipped_quantity: number;
  requested_quantity: number;
  pack_unit: string;
  batch_no: string;
  unit_weight: number;
  total_weight: number;
  unit_volume: number;
  total_volume: number;
  current_stock: number;
  available_stock: number;
  allocated_stock: number;
  pending_production: number;
  location_code: string;
  location_notes: string;
  latest_receipt_date?: string | null;
  is_contract_completed_in_wave: boolean;
  pair_batch_size: number;
  eligible: boolean;
  warning_reason?: string | null;
  pre_packing_allocated_stock?: number;
  dropped_quantity?: number;
  dropped_label?: string | null;
  dropped_dispatch_no?: string | null;
  dropped_dispatch_date?: string | null;
  dropped_dispatch_reference?: string | null;
  dropped_invoice_no?: string | null;
  order_no?: string | null;
  dimensions_source?: "product" | "similar_product_name" | "missing";
  dimensions_match?: {
    product_id: number;
    sku: string;
    name: string;
    score: number;
  };
};

type MitsuhaApiPlannerResponse = {
  format?: string;
  data?: MitsuhaApiPlannerItem[];
  summary?: {
    total_lines?: number;
    eligible_lines?: number;
    warning_lines?: number;
    total_quantity?: number;
    total_weight?: number;
    total_volume?: number;
    dropped_lines?: number;
    total_dropped_quantity?: number;
  };
  configuration?: {
    market_codes?: string[];
    excel_output?: Partial<MitsuhaApiExcelOutput>;
  };
  message?: string;
  errors?: Record<string, string[]>;
};

const MITSUHA_API_HEADERS: (string | number | boolean)[] = [
  "STT",
  "Mã hợp đồng",
  "Mã hàng",
  "Tên hàng",
  "Kích thước",
  "Số lượng đơn hàng",
  "ĐVT",
  "Đã xuất hàng",
  "Số lượng xuất đi",
  "ĐVT",
  "Số lô",
  "Trọng lượng",
  "Tổng trọng lượng",
  "Thể tích",
  "Tổng thể tích",
  "Ghi chú",
  "",
  "Hàng rớt",
  "PXK / ngày xuất kho",
  "Invoice",
  "Mã đơn hàng",
  "Mã phân loại",
  "Tồn kho thực tế hiện tại",
  "Chưa hoàn thành",
  "Lý do không đạt",
  "Ngày nhập mới nhất",
];

const mitsuhaApiLamSach = (value: unknown): string =>
  String(value ?? "")
    .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "")
    .trim();

/** Khóa nghiệp vụ luôn theo thứ tự SKU + mã hợp đồng. */
const mitsuhaApiKhoaSkuHopDong = (sku: unknown, contractNo: unknown): string =>
  `${mitsuhaApiLamSach(sku).toUpperCase()}|${mitsuhaApiLamSach(contractNo).toUpperCase()}`;

const mitsuhaApiSo = (value: unknown): number => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
};

/** Lấy số thứ ba trong kích thước: 26,5*37*1,1 → 1.1. */
const mitsuhaApiChieuCaoKichThuoc = (dimensions: unknown): number | null => {
  const parts = String(dimensions ?? "")
    .replace(/[×xX]/g, "*")
    .split("*")
    .map((part) => part.trim().replace(/,/g, "."));
  if (parts.length < 3) return null;
  const height = Number(parts[2].match(/-?\d+(?:\.\d+)?/)?.[0]);
  return Number.isFinite(height) ? height : null;
};

/**
 * Cột J: chiều cao = 1,1 / 1.1 → PACKAGE, còn lại → 箱/ THÙNG CARTON.
 * VD: 26,5*37*1,1 hoặc 37,5*26,5*1,1 → PACKAGE.
 */
const mitsuhaApiDonViDongGoi = (
  dimensions: unknown,
  packageHeight = 1.1,
): string => {
  const height = mitsuhaApiChieuCaoKichThuoc(dimensions);
  return height !== null && Math.abs(height - packageHeight) < 0.000001
    ? "PACKAGE"
    : "箱/ THÙNG CARTON";
};

/** ISO yyyy-mm-dd → Excel serial; giữ đúng ngày khi workbook có múi giờ khác UTC. */
const mitsuhaApiNgayExcel = (value?: string | null): number | string => {
  if (!value) return "";
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return String(value);
  const utc = Date.UTC(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
  );
  return utc / 86400000 + 25569;
};

/** Mọi thị trường được chọn trên cấu hình ghi chung một sheet, không tách theo thị trường. */
const MITSUHA_API_DEFAULT_SHEET_NAME = "pre-picking-list";

const mitsuhaApiTenSheet = (configuredName?: string | null): string =>
  String(configuredName ?? "").trim() || MITSUHA_API_DEFAULT_SHEET_NAME;

/** Gộp excel_output từ backend với mặc định; bỏ qua giá trị sai kiểu. */
const mitsuhaApiCauHinhExcel = (
  remote?: Partial<MitsuhaApiExcelOutput> | null,
): MitsuhaApiExcelOutput => {
  const result: MitsuhaApiExcelOutput = { ...MITSUHA_API_EXCEL_OUTPUT_DEFAULTS };
  if (!remote) return result;
  if (typeof remote.sheet_name === "string") result.sheet_name = remote.sheet_name;
  if (typeof remote.clear_old_data === "boolean") result.clear_old_data = remote.clear_old_data;
  if (typeof remote.enable_excel_filter === "boolean") result.enable_excel_filter = remote.enable_excel_filter;
  if (typeof remote.highlight_duplicate_sku === "boolean") {
    result.highlight_duplicate_sku = remote.highlight_duplicate_sku;
  }
  const startRow = Number(remote.start_row);
  if (Number.isInteger(startRow) && startRow >= 2) result.start_row = startRow;
  const colorFields: (keyof MitsuhaApiExcelOutput)[] = [
    "warning_fill_color",
    "completed_with_pending_fill_color",
    "dimension_fallback_fill_color",
    "duplicate_sku_fill_color",
  ];
  for (const field of colorFields) {
    const value = remote[field];
    if (typeof value === "string" && /^#[0-9A-Fa-f]{6}$/.test(value)) {
      (result as Record<string, unknown>)[field] = value;
    }
  }
  return result;
};

/** Một item backend → đúng 26 cột A→Z, Invoice nằm ngay sau cột S, Mã đơn hàng ngay sau Invoice. */
const mitsuhaApiTaoDong = (
  item: MitsuhaApiPlannerItem,
  stt: number,
): (string | number | boolean)[] => [
  stt, // A: STT
  item.contract_no || "", // B: Hợp đồng
  item.sku || "", // C: SKU
  item.product_name || "", // D: Tên hàng
  item.dimensions || "", // E: Kích thước
  mitsuhaApiSo(item.order_quantity), // F: SL đơn hàng
  "PCS", // G: ĐVT luôn PCS
  mitsuhaApiSo(item.previously_shipped_quantity), // H: Đã xuất + đã giữ
  mitsuhaApiSo(item.requested_quantity), // I: SL xuất đi
  mitsuhaApiDonViDongGoi(item.dimensions), // J: PACKAGE nếu cao=1,1; còn lại thùng carton
  item.batch_no || "", // K: Số lô
  mitsuhaApiSo(item.unit_weight), // L: Trọng lượng đơn vị
  mitsuhaApiSo(item.total_weight), // M: Tổng trọng lượng
  mitsuhaApiSo(item.unit_volume), // N: Thể tích đơn vị
  mitsuhaApiSo(item.total_volume), // O: Tổng thể tích
  item.location_notes || item.location_code || "", // P: Vị trí
  "", // Q: Gộp cùng header P
  item.dropped_label || "", // R: Hàng rớt
  item.dropped_dispatch_reference || "", // S: Số PXK — ngày xuất kho
  item.dropped_invoice_no || "", // T: Invoice của PXK có hàng rớt
  item.order_no || "", // U: Mã đơn hàng của SKU + hợp đồng
  item.classification_code || "", // V: Mã phân loại
  mitsuhaApiSo(item.current_stock), // W: Tồn thực tế
  mitsuhaApiSo(item.pending_production), // X: Chưa hoàn thành (số âm)
  item.eligible ? "" : item.warning_reason || "Chưa đạt điều kiện", // Y: Lý do
  mitsuhaApiNgayExcel(item.latest_receipt_date), // Z: Ngày nhập mới nhất
];

const mitsuhaApiKiemTraUrl = (url: string): string | null => {
  const normalized = url.trim().toLowerCase();
  if (!normalized) return "Chưa cấu hình apiUrl.";
  if (!normalized.startsWith("https://")) {
    return "apiUrl phải dùng HTTPS. Office Scripts không gọi được http://.";
  }
  if (normalized.includes("localhost") || normalized.includes("127.0.0.1")) {
    return "apiUrl không được dùng localhost — Office Scripts chạy trên cloud Microsoft.";
  }
  return null;
};

const mitsuhaApiLoiValidation = (errors?: Record<string, string[]>): string => {
  if (!errors) return "";
  return Object.keys(errors)
    .map((field) => `${field}: ${errors[field].join(", ")}`)
    .join(" | ");
};

const mitsuhaApiDatLaiDinhDang = (
  sheet: ExcelScript.Worksheet,
  address: string,
): void => {
  const format = sheet.getRange(address).getFormat();
  format.getFill().clear();
  format.getFont().setBold(false);
  format.getFont().setColor("#000000");
};

const mitsuhaApiKeVien = (range: ExcelScript.Range): void => {
  const format = range.getFormat();
  [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal,
    ExcelScript.BorderIndex.insideVertical,
  ].forEach((index) => {
    format
      .getRangeBorder(index)
      .setStyle(ExcelScript.BorderLineStyle.continuous);
  });
};

// @ts-ignore: Office Scripts cung cấp fetch ở runtime.
declare function fetch(
  input: string,
  init?: {
    method?: string;
    headers?: { [key: string]: string };
    body?: string;
  },
): Promise<{
  ok: boolean;
  status: number;
  text(): Promise<string>;
}>;

// @ts-ignore: Mỗi Office Script chạy độc lập nên có thể dùng tên main.
async function main(workbook: ExcelScript.Workbook) {
  const CFG = MITSUHA_API_CONFIG;
  const urlError = mitsuhaApiKiemTraUrl(CFG.apiUrl);
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

  console.log(
    "Đang yêu cầu backend lập kế hoạch soạn hàng theo cấu hình đã lưu trên web...",
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
      body: "{}",
    });

    const bodyText = await response.text();
    let body: MitsuhaApiPlannerResponse;
    try {
      body = JSON.parse(bodyText) as MitsuhaApiPlannerResponse;
    } catch {
      body = { message: bodyText };
    }

    if (!response.ok) {
      const validation = mitsuhaApiLoiValidation(body.errors);
      console.log(
        `LỖI HTTP ${response.status}: ${body.message || bodyText}${validation ? ` — ${validation}` : ""}`,
      );
      return;
    }
    if (
      body.format !== "auto_mitsuha_backend_v1" ||
      !Array.isArray(body.data)
    ) {
      console.log(
        "LỖI: Backend trả về sai format. Cần format=auto_mitsuha_backend_v1 và data là một mảng.",
      );
      return;
    }

    const items = body.data;
    const marketCodes = body.configuration?.market_codes ?? [];
    const OUT = mitsuhaApiCauHinhExcel(body.configuration?.excel_output);
    console.log(`• Thị trường: ${marketCodes.join(", ") || "tất cả"}`);
    const sheetName = mitsuhaApiTenSheet(OUT.sheet_name);
    let sheet = workbook.getWorksheet(sheetName);
    if (!sheet) sheet = workbook.addWorksheet(sheetName);

    const usedRange = sheet.getUsedRange();
    const oldLastRow = usedRange ? usedRange.getLastRow().getRowIndex() + 1 : 1;
    if (OUT.clear_old_data && oldLastRow >= OUT.start_row) {
      sheet
        .getRange(`A${OUT.start_row}:Z${oldLastRow}`)
        .clear(ExcelScript.ClearApplyTo.all);
    }

    sheet.getRange("P1:Q1").unmerge();
    sheet.getRange("A1:Z1").setValues([MITSUHA_API_HEADERS]);
    sheet.getRange("P1:Q1").merge(false);
    sheet.getRange("U:V").setColumnHidden(false);

    const headerFormat = sheet.getRange("A1:Z1").getFormat();
    headerFormat.getFont().setBold(true);
    headerFormat.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    headerFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    headerFormat.getFill().setColor("#D9EAF7");
    mitsuhaApiKeVien(sheet.getRange("A1:Z1"));

    if (items.length === 0) {
      sheet.getFreezePanes().unfreeze();
      sheet.getFreezePanes().freezeRows(1);
      console.log(
        `Backend không tìm thấy dòng phù hợp. Đã làm mới tiêu đề sheet "${sheetName}".`,
      );
      return;
    }

    const rows = items.map((item, index) => mitsuhaApiTaoDong(item, index + 1));
    const endRow = OUT.start_row + rows.length - 1;
    const dataAddress = `A${OUT.start_row}:Z${endRow}`;
    // Giữ mã đơn hàng dạng text để Excel không tự đổi mã toàn số (mất số 0 đầu).
    sheet.getRange(`U${OUT.start_row}:U${endRow}`).setNumberFormat("@");
    sheet.getRange(dataAddress).setValues(rows);
    mitsuhaApiDatLaiDinhDang(sheet, dataAddress);

    for (let index = 0; index < items.length; index++) {
      const item = items[index];
      const row = OUT.start_row + index;
      if (!item.eligible) {
        sheet
          .getRange(`A${row}:Z${row}`)
          .getFormat()
          .getFill()
          .setColor(OUT.warning_fill_color);
      } else if (
        item.allocated_stock > 0 &&
        item.is_contract_completed_in_wave
      ) {
        sheet
          .getRange(`A${row}:Z${row}`)
          .getFormat()
          .getFill()
          .setColor(OUT.completed_with_pending_fill_color);
      }

      if (item.dimensions_source === "similar_product_name") {
        sheet
          .getRange(`E${row}`)
          .getFormat()
          .getFill()
          .setColor(OUT.dimension_fallback_fill_color);
      }
    }

    if (OUT.highlight_duplicate_sku) {
      const skuCounts: Record<string, number> = {};
      for (const item of items) {
        const key = mitsuhaApiLamSach(item.base_sku || item.sku).toUpperCase();
        if (key) skuCounts[key] = (skuCounts[key] || 0) + 1;
      }
      for (let index = 0; index < items.length; index++) {
        const key = mitsuhaApiLamSach(
          items[index].base_sku || items[index].sku,
        ).toUpperCase();
        if (key && skuCounts[key] > 1) {
          sheet
            .getRange(`C${OUT.start_row + index}`)
            .getFormat()
            .getFill()
            .setColor(OUT.duplicate_sku_fill_color);
        }
      }
    }

    sheet.getRange(`F${OUT.start_row}:I${endRow}`).setNumberFormat("0");
    sheet.getRange(`L${OUT.start_row}:M${endRow}`).setNumberFormat("0.000");
    sheet.getRange(`N${OUT.start_row}:O${endRow}`).setNumberFormat("0.0000");
    sheet.getRange(`W${OUT.start_row}:X${endRow}`).setNumberFormat("0");
    sheet
      .getRange(`Z${OUT.start_row}:Z${endRow}`)
      .setNumberFormat("dd/mm/yyyy");
    sheet
      .getRange(`Y${OUT.start_row}:Y${endRow}`)
      .getFormat()
      .setWrapText(true);
    mitsuhaApiKeVien(sheet.getRange(`A1:Z${endRow}`));

    if (OUT.enable_excel_filter) {
      try {
        sheet.getAutoFilter().remove();
      } catch {
        // Sheet chưa có AutoFilter.
      }
      sheet.getAutoFilter().apply(sheet.getRange(`A1:Z${endRow}`));
    }
    sheet.getFreezePanes().unfreeze();
    sheet.getFreezePanes().freezeRows(1);
    sheet.getRange("A:Z").getFormat().autofitColumns();
    sheet.activate();

    const summary = body.summary || {};
    console.log("=========================================");
    console.log("  HOÀN TẤT SOẠN HÀNG TỪ BACKEND");
    console.log("=========================================");
    console.log(`• Sheet: ${sheetName}`);
    console.log(
      `• Tổng dòng: ${items.length} | Đạt: ${summary.eligible_lines ?? items.filter((item) => item.eligible).length} | Cảnh báo: ${summary.warning_lines ?? items.filter((item) => !item.eligible).length}`,
    );
    console.log(
      `• Số lượng đề xuất: ${summary.total_quantity ?? items.reduce((sum, item) => sum + item.requested_quantity, 0)}`,
    );
    console.log(
      `• Hàng rớt: ${summary.dropped_lines ?? items.filter((item) => (item.dropped_quantity || 0) > 0).length} dòng | tổng SL ${summary.total_dropped_quantity ?? items.reduce((sum, item) => sum + (item.dropped_quantity || 0), 0)}`,
    );
    console.log(
      `• Thể tích: ${summary.total_volume ?? 0} m³ | Thời gian: ${Math.round((Date.now() - startedAt) / 1000)}s`,
    );
  } catch (error) {
    const message = String(error);
    if (message.includes("Failed to fetch")) {
      console.log(
        "LỖI kết nối API: Failed to fetch — kiểm tra URL HTTPS public, server deploy và CORS.",
      );
      console.log(`API hiện tại: ${CFG.apiUrl}`);
    } else {
      console.log(`LỖI khi lấy đề xuất soạn hàng: ${message}`);
    }
  }
}

const mitsuhaApiTestables = {
  mitsuhaApiNgayExcel,
  mitsuhaApiTenSheet,
  mitsuhaApiCauHinhExcel,
  mitsuhaApiKhoaSkuHopDong,
  mitsuhaApiTaoDong,
  mitsuhaApiKiemTraUrl,
  mitsuhaApiChieuCaoKichThuoc,
  mitsuhaApiDonViDongGoi,
};

(
  globalThis as typeof globalThis & {
    __AUTO_MITSUHA_API_TESTABLES__?: typeof mitsuhaApiTestables;
  }
).__AUTO_MITSUHA_API_TESTABLES__ = mitsuhaApiTestables;
