/**
 * ===================================================================
 * Upload_ProductSpecs – Cập nhật trọng lượng và kích thước SKU theo kỳ.
 * ===================================================================
 *
 * Cấu trúc sheet (dòng 1 là tiêu đề, dữ liệu từ dòng 2):
 * A Ngày (dd/mm/yyyy) | B SKU | C Tên thành phẩm | D (không dùng)
 * E Cân nặng (kg) | F Dài (cm) | G Rộng (cm) | H Cao (cm)
 * Nếu cột H để trống, chiều cao mặc định theo defaultHeight trong CONFIG.
 *
 * Luồng sử dụng:
 * 1. Chuẩn bị sheet dữ liệu (cột A→H như trên).
 * 2. Tạo API token tại Dashboard → API Token (Excel).
 * 3. Cấu hình period / anchorDate / apiUrl / apiToken bên dưới.
 * 4. Chạy script này trong Excel Online.
 *
 * Nếu một SKU có nhiều dòng trong cùng kỳ, dòng có ngày mới nhất được dùng.
 * Nếu các dòng cùng ngày, dòng nằm dưới cùng được dùng.
 *
 * QUAN TRỌNG — Office Scripts chạy trên cloud Microsoft (Excel Online):
 * - KHÔNG dùng localhost / 127.0.0.1.
 * - BẮT BUỘC dùng HTTPS.
 */

type ProductSpecsPeriod = "day" | "week" | "month";

const UPLOAD_PRODUCT_SPECS_CONFIG = {
  /**
   * Kỳ lấy dữ liệu:
   * - "day": đúng ngày anchorDate
   * - "week": tuần chứa anchorDate (thứ 2 → chủ nhật)
   * - "month": cả tháng của anchorDate
   */
  period: "month" as ProductSpecsPeriod,
  /** Ngày mốc dd/mm/yyyy; để trống = hôm nay. */
  anchorDate: "01/08/2026",
  apiUrl: "https://app.wgr-wh.site/api/v1/warehouse/products/batch-import",
  apiToken: "1|hCwtksIBz4Yc7g2wn6aeMO4L6zri5rStWeSYgH9rf7ad4909",
  batchSize: 500,
  /** Chiều cao (cm) khi cột H trống. */
  defaultHeight: 1.1,
};

const PRODUCT_SPECS_DAY_MS = 86400000;

type ProductSpecsItem = {
  sku: string;
  name: string;
  unit_weight: number;
  length: number;
  width: number;
  height: number;
  unit_volume: number;
};

type ProductSpecsCandidate = ProductSpecsItem & {
  dateKey: number;
  excelRow: number;
};

type ProductSpecsReadResult = {
  items: ProductSpecsItem[];
  errors: string[];
  duplicateCount: number;
};

type ProductSpecsApiBody = {
  message?: string;
  count?: number;
  errors?: Record<string, string[]>;
};

// @ts-ignore: Các file Office Scripts được chạy độc lập trong Excel.
const productSpecsClean = (value: string | number | boolean): string =>
  String(value ?? "")
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .trim();

// @ts-ignore
const productSpecsParseNumber = (
  value: string | number | boolean,
): number | null => {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;

  let text = productSpecsClean(value).replace(/\s/g, "");
  if (!text) return null;

  const lastComma = text.lastIndexOf(",");
  const lastDot = text.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    text =
      lastComma > lastDot
        ? text.replace(/\./g, "").replace(",", ".")
        : text.replace(/,/g, "");
  } else if (lastComma >= 0) {
    text = text.replace(",", ".");
  }

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : null;
};

/** Trả về số ngày UTC, giúp so sánh ngày mà không bị lệch múi giờ. */
// @ts-ignore
const productSpecsParseDateKey = (
  value: string | number | boolean,
): number | null => {
  if (typeof value === "number" && Number.isFinite(value)) {
    // Excel dùng mốc 30/12/1899. Bỏ phần giờ nếu ô có kèm thời gian.
    return Math.floor(value) - 25569;
  }

  const text = productSpecsClean(value);
  let day: number;
  let month: number;
  let year: number;
  let match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);

  if (match) {
    day = Number(match[1]);
    month = Number(match[2]);
    year = Number(match[3]);
  } else {
    match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (!match) return null;
    year = Number(match[1]);
    month = Number(match[2]);
    day = Number(match[3]);
  }

  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return null;
  }

  return Math.floor(date.getTime() / PRODUCT_SPECS_DAY_MS);
};

// @ts-ignore
const productSpecsFormatDateKey = (dateKey: number): string => {
  const date = new Date(dateKey * PRODUCT_SPECS_DAY_MS);
  const day = String(date.getUTCDate()).padStart(2, "0");
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  return `${day}/${month}/${date.getUTCFullYear()}`;
};

// @ts-ignore
const productSpecsTodayKey = (): number => {
  const now = new Date();
  return Math.floor(
    Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) /
      PRODUCT_SPECS_DAY_MS,
  );
};

// @ts-ignore
const productSpecsGetPeriod = (
  modeInput: string,
  anchorKey: number,
): { mode: ProductSpecsPeriod; startKey: number; endKey: number } | null => {
  const mode = productSpecsClean(modeInput).toLowerCase();
  if (mode !== "day" && mode !== "week" && mode !== "month") return null;

  if (mode === "day") {
    return { mode, startKey: anchorKey, endKey: anchorKey };
  }

  const anchor = new Date(anchorKey * PRODUCT_SPECS_DAY_MS);
  if (mode === "week") {
    const mondayOffset = (anchor.getUTCDay() + 6) % 7;
    const startKey = anchorKey - mondayOffset;
    return { mode, startKey, endKey: startKey + 6 };
  }

  const year = anchor.getUTCFullYear();
  const month = anchor.getUTCMonth();
  const startKey = Math.floor(Date.UTC(year, month, 1) / PRODUCT_SPECS_DAY_MS);
  const endKey = Math.floor(
    Date.UTC(year, month + 1, 0) / PRODUCT_SPECS_DAY_MS,
  );
  return { mode, startKey, endKey };
};

// @ts-ignore
const productSpecsReadRows = (
  rows: (string | number | boolean)[][],
  startKey: number,
  endKey: number,
): ProductSpecsReadResult => {
  const latestBySku = new Map<string, ProductSpecsCandidate>();
  const errors: string[] = [];
  let duplicateCount = 0;

  // rows[0] là dòng tiêu đề Excel.
  for (let index = 1; index < rows.length; index++) {
    const row = rows[index] ?? [];
    const excelRow = index + 1;
    const hasData = row
      .slice(0, 8)
      .some((cell) => productSpecsClean(cell) !== "");
    if (!hasData) continue;

    const dateKey = productSpecsParseDateKey(row[0] ?? "");
    if (dateKey === null) {
      errors.push(`Dòng ${excelRow}: ngày cột A không hợp lệ.`);
      continue;
    }
    if (dateKey < startKey || dateKey > endKey) continue;

    const sku = productSpecsClean(row[1] ?? "").toUpperCase();
    const name = productSpecsClean(row[2] ?? "");
    const weight = productSpecsParseNumber(row[4] ?? "");
    const length = productSpecsParseNumber(row[5] ?? "");
    const width = productSpecsParseNumber(row[6] ?? "");
    const heightText = productSpecsClean(row[7] ?? "");
    const height =
      heightText === ""
        ? UPLOAD_PRODUCT_SPECS_CONFIG.defaultHeight
        : productSpecsParseNumber(row[7] ?? "");

    const missing: string[] = [];
    if (!sku) missing.push("SKU (B)");
    if (!name) missing.push("tên thành phẩm (C)");
    if (weight === null || weight <= 0) missing.push("cân nặng (E)");
    if (length === null || length <= 0) missing.push("dài (F)");
    if (width === null || width <= 0) missing.push("rộng (G)");
    if (height === null || height <= 0) missing.push("cao (H)");

    if (missing.length > 0) {
      errors.push(`Dòng ${excelRow}: thiếu/sai ${missing.join(", ")}.`);
      continue;
    }

    const candidate: ProductSpecsCandidate = {
      sku,
      name,
      unit_weight: weight as number,
      length: length as number,
      width: width as number,
      height: height as number,
      unit_volume:
        ((length as number) * (width as number) * (height as number)) / 1000000,
      dateKey,
      excelRow,
    };

    const previous = latestBySku.get(sku);
    if (previous) duplicateCount++;
    if (
      !previous ||
      candidate.dateKey > previous.dateKey ||
      (candidate.dateKey === previous.dateKey &&
        candidate.excelRow > previous.excelRow)
    ) {
      latestBySku.set(sku, candidate);
    }
  }

  const items = Array.from(latestBySku.values())
    .sort((a, b) => a.sku.localeCompare(b.sku))
    .map(({ dateKey: _dateKey, excelRow: _excelRow, ...item }) => item);

  return { items, errors, duplicateCount };
};

// @ts-ignore
const productSpecsValidateApiUrl = (url: string): string | null => {
  const normalized = url.trim().toLowerCase();
  if (!normalized) return "Chưa nhập apiUrl.";
  if (!normalized.startsWith("https://")) return "apiUrl phải dùng HTTPS.";
  if (normalized.includes("localhost") || normalized.includes("127.0.0.1")) {
    return "Office Scripts không thể gọi localhost.";
  }
  return null;
};

// @ts-ignore: Office Scripts cung cấp fetch khi chạy trên Excel Online.
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

// @ts-ignore: Mỗi file Office Scripts có một hàm main riêng.
async function main(workbook: ExcelScript.Workbook) {
  const CFG = UPLOAD_PRODUCT_SPECS_CONFIG;
  const urlError = productSpecsValidateApiUrl(CFG.apiUrl);
  if (urlError) {
    console.log(`LỖI cấu hình: ${urlError}`);
    return;
  }
  if (!CFG.apiToken.trim()) {
    console.log("LỖI: Chưa cấu hình apiToken (Bearer token từ Dashboard).");
    return;
  }

  const anchorKey = CFG.anchorDate.trim()
    ? productSpecsParseDateKey(CFG.anchorDate)
    : productSpecsTodayKey();
  if (anchorKey === null) {
    console.log(
      'LỖI: anchorDate phải có định dạng dd/mm/yyyy, ví dụ "06/09/2026".',
    );
    return;
  }

  const period = productSpecsGetPeriod(CFG.period, anchorKey);
  if (!period) {
    console.log(
      'LỖI: period chỉ nhận một trong ba giá trị: "day", "week", "month".',
    );
    return;
  }

  const sheet = workbook.getActiveWorksheet();
  const usedRange = sheet.getUsedRange(true);
  if (!usedRange) {
    console.log(`Sheet "${sheet.getName()}" không có dữ liệu.`);
    return;
  }

  const rows = usedRange.getValues() as (string | number | boolean)[][];
  const result = productSpecsReadRows(rows, period.startKey, period.endKey);
  const periodLabel = `${productSpecsFormatDateKey(period.startKey)} → ${productSpecsFormatDateKey(period.endKey)}`;

  if (result.errors.length > 0) {
    console.log(
      [
        `CẢNH BÁO: bỏ qua ${result.errors.length} dòng không hợp lệ:`,
        ...result.errors.slice(0, 30).map((error) => `- ${error}`),
        ...(result.errors.length > 30
          ? [`- ... và ${result.errors.length - 30} lỗi khác.`]
          : []),
      ].join("\n"),
    );
  }

  if (result.items.length === 0) {
    console.log(`Không có SKU hợp lệ trong kỳ ${periodLabel}.`);
    return;
  }

  console.log(
    `Đang cập nhật ${result.items.length} SKU trong kỳ ${periodLabel}` +
      (result.duplicateCount > 0
        ? ` (đã chọn bản ghi mới nhất cho ${result.duplicateCount} dòng trùng).`
        : "."),
  );

  let updated = 0;
  let apiError = "";
  try {
    for (
      let offset = 0;
      offset < result.items.length;
      offset += CFG.batchSize
    ) {
      const items = result.items.slice(offset, offset + CFG.batchSize);
      const response = await fetch(CFG.apiUrl.trim(), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json",
          Authorization: `Bearer ${CFG.apiToken.trim()}`,
        },
        body: JSON.stringify({ items }),
      });

      const bodyText = await response.text();
      const body = JSON.parse(bodyText) as ProductSpecsApiBody;

      if (!response.ok) {
        const details = body.errors
          ? Object.keys(body.errors)
              .map((field) => `${field}: ${body.errors![field].join(", ")}`)
              .join("\n")
          : "";
        apiError =
          `LỖI HTTP ${response.status}: ${body.message ?? bodyText}` +
          (details ? `\n${details}` : "");
        break;
      }

      updated += body.count ?? items.length;
    }

    if (apiError) {
      console.log(apiError);
    } else {
      console.log(
        `HOÀN TẤT: đã cập nhật trọng lượng và kích thước cho ${updated} SKU (${periodLabel}).`,
      );
    }
  } catch (error) {
    const message = String(error);
    console.log(
      message.includes("Failed to fetch")
        ? `LỖI kết nối API: kiểm tra URL HTTPS, token và server. URL hiện tại: ${CFG.apiUrl}`
        : `LỖI: ${message}`,
    );
  }
}

const productSpecsTestables = {
  productSpecsParseNumber,
  productSpecsParseDateKey,
  productSpecsFormatDateKey,
  productSpecsGetPeriod,
  productSpecsReadRows,
};

(
  globalThis as typeof globalThis & {
    __UPLOAD_PRODUCT_SPECS_TESTABLES__?: typeof productSpecsTestables;
  }
).__UPLOAD_PRODUCT_SPECS_TESTABLES__ = productSpecsTestables;
