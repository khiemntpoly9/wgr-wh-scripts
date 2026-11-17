/**
 * ===================================================================
 * Upload_PXK – Đọc sheet "sale" (format Auto_PXK) và upload Phiếu Xuất Kho lên API.
 * ===================================================================
 *
 * Meta (header phiếu):
 * - D3: Ngày lập (vd: "04 tháng 08 năm 2026" / "2026年09月04")
 * - D4: Số phiếu xuất kho (vd: "NO:  000462")
 * - D7: Mã cont (vd: "CSNU7515625")
 * - D8: Số seal / chì (vd: "SEAL-01")
 * - Q6: Invoice (vd: "WGR-20260095")
 *
 * Cột dữ liệu (A→R, giống Auto_PXK.ts):
 * A STT | B Hợp đồng | C SKU | D Tên hàng | E Kích thước
 * F SL xuất đi | G ĐVT | H Đã xuất | I SL đóng gói | J ĐVT đóng gói
 * K Số lô | L Trọng lượng | M Tổng TL | N Thể tích | O Tổng TT
 * P Vị trí | Q Ghi chú | R PMC
 *
 * Luồng sử dụng:
 * 1. Chuẩn bị/chạy Auto_PXK trên sheet "sale".
 * 2. Kiểm tra ô D3, D4, D7, D8, Q6.
 * 3. Tạo API token tại Dashboard → API Token (Excel).
 * 4. Cấu hình apiUrl (HTTPS public) và apiToken bên dưới.
 * 5. Tuỳ chọn bật tuDongTaoPhieuDieuChinhTon để tự động bù phần tồn thiếu.
 * 6. Chạy script này trong Excel Online.
 * 7. Nếu thiếu tồn và không tự động bù: tạo/ghi sheet "SKU thiếu tồn".
 *
 * QUAN TRỌNG — Office Scripts chạy trên cloud Microsoft (Excel Online):
 * - KHÔNG dùng localhost / 127.0.0.1.
 * - BẮT BUỘC dùng HTTPS.
 */

const UPLOAD_PXK_CONFIG = {
  tenSheet: "sale",
  headerKeywords: ["STT", "序号"],
  firstDataCol: "B",
  lastDataCol: "R",
  /** Sheet liệt kê SKU thiếu tồn khi API từ chối xuất kho. */
  shortageSheet: "SKU thiếu tồn",
  /** Ô chứa ngày lập (vd: "04 tháng 08 năm 2026" / "2026年09月04"). */
  dispatchDateCell: "D3",
  /** Ô chứa số phiếu (vd: "NO:  000462"). */
  dispatchNoCell: "D4",
  /** Ô chứa mã cont (vd: "CSNU7515625"). */
  containerNoCell: "D7",
  /** Ô đọc số seal / chì (vd: "SEAL-01"). */
  sealNoCell: "D8",
  /** Ô chứa invoice (vd: "WGR-20260095"). */
  invoiceNoCell: "Q6",
  apiUrl: "https://app.wgr-wh.site/api/v1/warehouse/dispatch-orders/pxk-import",
  apiToken: "", // Bearer token từ Dashboard
  /**
   * true: API tự tạo + duyệt phiếu điều chỉnh cho đúng phần tồn còn thiếu,
   * sau đó mới tạo PXK. Token phải có quyền "Duyệt điều chỉnh tồn".
   * false: giữ hành vi cũ — API báo lỗi và ghi sheet SKU thiếu tồn.
   */
  tuDongTaoPhieuDieuChinhTon: false,
  /** Tuỳ chọn ghi đè thủ công. Để trống = tự đọc từ ô trên sheet. */
  dispatchNo: "",
  dispatchDate: "",
  containerNo: "",
  sealNo: "",
  invoiceNo: "",
};

type PxkUploadRow = {
  stt: number;
  contract_no: string;
  sku: string;
  product_name: string;
  dimensions: string;
  dispatch_quantity: number;
  unit: string;
  previously_shipped_quantity: number;
  packed_quantity: number;
  pack_unit: string;
  batch_no: string;
  unit_weight: number | null;
  total_weight: number | null;
  unit_volume: number | null;
  total_volume: number | null;
  location_notes: string;
  notes: string;
  pmc_approved: string;
};

// @ts-ignore: Tránh cảnh báo trùng lặp hàm khi code nhiều file TS trong cùng thư mục
const pxkUpLamSach = (giaTri: string | number | boolean): string =>
  String(giaTri ?? "")
    .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "")
    .trim();

// @ts-ignore
const pxkUpChuoi = (giaTri: string | number | boolean): string =>
  String(giaTri ?? "").trim();

// @ts-ignore
const pxkUpParseSo = (giaTri: string | number | boolean): number => {
  if (typeof giaTri === "number" && !isNaN(giaTri)) return giaTri;
  let s = pxkUpLamSach(giaTri);
  if (!s) return 0;
  const lastComma = s.lastIndexOf(",");
  const lastDot = s.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    if (lastComma > lastDot) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (lastComma >= 0) {
    s = s.replace(",", ".");
  }
  const n = Number(s);
  return isNaN(n) ? 0 : n;
};

const pxkUpParseSoHoacNull = (
  giaTri: string | number | boolean,
): number | null => {
  const n = pxkUpParseSo(giaTri);
  return n === 0 && pxkUpLamSach(giaTri) === "" ? null : n;
};

const pxkUpChuanHoaSku = (sku: string): string => {
  const cleaned = pxkUpLamSach(sku);
  // SKU = toàn bộ dãy số đầu mã (≥ 6 số); cặp A/B (912345A / 912345B) là 2 sản phẩm riêng nên giữ nguyên.
  if (/^9\d{5,}[AB]$/i.test(cleaned)) return cleaned.toUpperCase();
  const match = cleaned.match(/^(9\d{5,})(\D.*)$/);
  return match ? match[1] : cleaned;
};

/** Định dạng Date → YYYY-MM-DD (UTC). */
const pxkUpFormatIsoDate = (date: Date): string | null => {
  if (isNaN(date.getTime())) return null;
  const yyyy = date.getUTCFullYear();
  const mm = String(date.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(date.getUTCDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
};

/** Chuyển serial ngày Excel sang Date (UTC). */
const pxkUpExcelSerialToDate = (serial: number): Date => {
  const wholeDays = Math.floor(serial);
  const excelEpoch = Date.UTC(1899, 11, 30);
  return new Date(excelEpoch + wholeDays * 86400000);
};

/** Parse một dòng text ngày sang YYYY-MM-DD. */
const pxkUpParseMotDongNgay = (line: string): string | null => {
  const s = pxkUpChuoi(line).replace(/[,，。.]+$/g, "").trim();
  if (!s) return null;

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return s;
  }

  const vnMatch = s.match(
    /(\d{1,2})\s*th[áa]ng\s*(\d{1,2})\s*n[ăa]m\s*(\d{4})/i,
  );
  if (vnMatch) {
    const dd = vnMatch[1].padStart(2, "0");
    const mm = vnMatch[2].padStart(2, "0");
    const yyyy = vnMatch[3];
    return `${yyyy}-${mm}-${dd}`;
  }

  const zhMatch = s.match(/(\d{4})年(\d{1,2})月(\d{1,2})日?/);
  if (zhMatch) {
    const yyyy = zhMatch[1];
    const mm = zhMatch[2].padStart(2, "0");
    const dd = zhMatch[3].padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  const isoSlash = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (isoSlash) {
    const yyyy = isoSlash[1];
    const mm = isoSlash[2].padStart(2, "0");
    const dd = isoSlash[3].padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  const dmY = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (dmY) {
    const dd = dmY[1].padStart(2, "0");
    const mm = dmY[2].padStart(2, "0");
    const yyyy = dmY[3];
    return `${yyyy}-${mm}-${dd}`;
  }

  return null;
};

/** Trích số phiếu từ ô D4 (vd: "NO:  000462"). */
const pxkUpTrichXuatSoPhieu = (raw: string | number | boolean): string => {
  const s = pxkUpChuoi(raw);
  if (!s) return "";

  const noMatch = s.match(/NO\s*[：:]\s*(\S+)/i);
  if (noMatch) return noMatch[1].trim();

  const digits = s.match(/(\d[\d\s]*\d|\d+)/);
  if (digits) return digits[1].replace(/\s+/g, "");

  return s;
};

/** Chuẩn hóa mã cont / invoice (bỏ nhãn tiền tố nếu có). */
const pxkUpTrichXuatMa = (
  raw: string | number | boolean,
  labels: string[],
): string => {
  let s = pxkUpChuoi(raw);
  if (!s) return "";

  for (let i = 0; i < labels.length; i++) {
    const label = labels[i];
    const re = new RegExp(`^${label}\\s*[：:]\\s*`, "i");
    s = s.replace(re, "").trim();
  }

  return s.replace(/[,，。.]+$/g, "").trim();
};

/** Chuyển ngày sang YYYY-MM-DD cho API Laravel. */
const pxkUpChuanHoaNgay = (
  raw: string | number | boolean,
): string | null => {
  if (typeof raw === "number" && !isNaN(raw)) {
    if (raw >= 30000 && raw <= 60000) {
      return pxkUpFormatIsoDate(pxkUpExcelSerialToDate(raw));
    }
  }

  const s = pxkUpChuoi(raw);
  if (!s) return null;

  const lines = s
    .split(/[\r\n]+/)
    .map((line) => line.trim())
    .filter((line) => Boolean(line));
  const candidates = lines.length > 0 ? lines : [s];

  for (let i = 0; i < candidates.length; i++) {
    const parsed = pxkUpParseMotDongNgay(candidates[i]);
    if (parsed) return parsed;
  }

  return null;
};

/** Đọc meta phiếu từ D3 / D4 / D7 / D8 / Q6. */
const pxkUpDocMetaTuSheet = (
  sheet: ExcelScript.Worksheet,
  dateCell: string,
  noCell: string,
  containerCell: string,
  sealCell: string,
  invoiceCell: string,
): {
  dispatchDateRaw: string;
  dispatchNo: string;
  dispatchDateIso: string | null;
  containerNo: string;
  sealNo: string;
  invoiceNo: string;
} => {
  const dateVal = sheet.getRange(dateCell).getValue();
  const noVal = sheet.getRange(noCell).getValue();
  const containerVal = sheet.getRange(containerCell).getValue();
  const sealVal = sheet.getRange(sealCell).getValue();
  const invoiceVal = sheet.getRange(invoiceCell).getValue();

  return {
    dispatchDateRaw: pxkUpChuoi(dateVal),
    dispatchNo: pxkUpTrichXuatSoPhieu(noVal),
    dispatchDateIso: pxkUpChuanHoaNgay(dateVal),
    containerNo: pxkUpTrichXuatMa(containerVal, ["CONT", "CONTAINER", "MÃ CONT", "MA CONT"]),
    sealNo: pxkUpTrichXuatMa(sealVal, ["SEAL", "CHÌ", "CHI", "SỐ SEAL", "SO SEAL"]),
    invoiceNo: pxkUpTrichXuatMa(invoiceVal, ["INVOICE", "INV"]),
  };
};

// @ts-ignore: Tránh trùng findHeaderRow trong Auto_PXK.ts
const pxkUpFindHeaderRow = (
  rows: (string | number | boolean)[][],
  keywords: string[],
): number => {
  for (let i = 0; i < rows.length; i++) {
    const cellValue = pxkUpChuoi(rows[i][0]).toUpperCase();
    const allMatch = keywords.every((kw) =>
      cellValue.includes(kw.toUpperCase()),
    );
    if (allMatch) return i;
  }
  return -1;
};

// @ts-ignore: Tránh trùng findDataBounds trong Auto_PXK.ts
const pxkUpFindDataBounds = (
  rows: (string | number | boolean)[][],
  startRow: number,
): { startRow: number; lastRow: number } | null => {
  if (startRow >= rows.length) return null;

  let lastDataRow = -1;
  for (let i = startRow; i < rows.length; i++) {
    const productName = pxkUpChuoi(rows[i][3] ?? "");
    if (productName !== "") {
      lastDataRow = i;
      continue;
    }

    const rowText = rows[i]
      .slice(0, 18)
      .map((c) => pxkUpChuoi(c))
      .join(" ")
      .toUpperCase();
    if (rowText.includes("SUM(") || rowText.includes("TỔNG")) {
      break;
    }
  }

  if (lastDataRow < startRow) return null;

  return { startRow, lastRow: lastDataRow };
};

const pxkUpDocDuLieu = (
  rows: (string | number | boolean)[][],
  headerKeywords: string[],
): PxkUploadRow[] => {
  const headerIdx = pxkUpFindHeaderRow(rows, headerKeywords);
  if (headerIdx === -1) return [];

  const bounds = pxkUpFindDataBounds(rows, headerIdx + 1);
  if (!bounds) return [];

  const items: PxkUploadRow[] = [];

  for (let i = bounds.startRow; i <= bounds.lastRow; i++) {
    const row = rows[i];
    const sku = pxkUpChuanHoaSku(String(row[2] ?? ""));
    const contractNo = pxkUpChuoi(row[1] ?? "");
    const productName = pxkUpChuoi(row[3] ?? "");

    if (!sku || !contractNo || !productName) continue;

    let dispatchQty = Math.round(pxkUpParseSo(row[5] ?? 0));
    let packedQty = Math.round(pxkUpParseSo(row[8] ?? 0));
    if (dispatchQty <= 0 && packedQty <= 0) continue;
    if (dispatchQty <= 0) dispatchQty = packedQty;
    if (packedQty <= 0) packedQty = dispatchQty;

    items.push({
      stt: Math.round(pxkUpParseSo(row[0] ?? 0)) || items.length + 1,
      contract_no: contractNo,
      sku,
      product_name: productName,
      dimensions: pxkUpChuoi(row[4] ?? ""),
      dispatch_quantity: dispatchQty,
      unit: pxkUpChuoi(row[6] ?? ""),
      previously_shipped_quantity: Math.round(pxkUpParseSo(row[7] ?? 0)),
      packed_quantity: packedQty,
      pack_unit: pxkUpChuoi(row[9] ?? ""),
      batch_no: pxkUpChuoi(row[10] ?? ""),
      unit_weight: pxkUpParseSoHoacNull(row[11] ?? ""),
      total_weight: pxkUpParseSoHoacNull(row[12] ?? ""),
      unit_volume: pxkUpParseSoHoacNull(row[13] ?? ""),
      total_volume: pxkUpParseSoHoacNull(row[14] ?? ""),
      location_notes: pxkUpChuoi(row[15] ?? ""), // P: Vị trí
      notes: pxkUpChuoi(row[16] ?? ""), // Q: Ghi chú
      pmc_approved: pxkUpChuoi(row[17] ?? ""),
    });
  }

  return items;
};

// @ts-ignore: Tránh trùng compactItem
const pxkUpCompactItem = (
  item: PxkUploadRow,
): Record<string, string | number> => {
  const out: Record<string, string | number> = {};
  const keys: (keyof PxkUploadRow)[] = [
    "stt",
    "contract_no",
    "sku",
    "product_name",
    "dimensions",
    "dispatch_quantity",
    "unit",
    "previously_shipped_quantity",
    "packed_quantity",
    "pack_unit",
    "batch_no",
    "unit_weight",
    "total_weight",
    "unit_volume",
    "total_volume",
    "location_notes",
    "notes",
    "pmc_approved",
  ];

  keys.forEach((key) => {
    const val = item[key];
    if (val !== null && val !== undefined && val !== "") {
      out[key] = val;
    }
  });

  return out;
};

// @ts-ignore
const pxkUpKiemTraApiUrl = (url: string): string | null => {
  const u = url.trim().toLowerCase();
  if (!u) return "Chưa cấu hình apiUrl.";
  if (!u.startsWith("https://")) {
    return "apiUrl phải dùng HTTPS. Office Scripts không gọi được http://.";
  }
  if (u.includes("localhost") || u.includes("127.0.0.1")) {
    return "apiUrl không được dùng localhost — Office Scripts chạy trên cloud Microsoft.";
  }
  return null;
};

const pxkUpFormatValidationErrors = (
  errors: Record<string, string[]>,
): string => {
  return Object.keys(errors)
    .map((field) => {
      const messages = errors[field] ?? [];
      if (messages.length <= 1) {
        return `  - ${field}: ${messages.join(", ")}`;
      }

      // Nhiều SKU thiếu tồn: liệt kê từng dòng thay vì gộp một dòng.
      return [`  - ${field}:`, ...messages.map((msg) => `      • ${msg}`)].join(
        "\n",
      );
    })
    .join("\n");
};

type PxkShortageRow = {
  sku: string;
  contract_no: string;
  product_name: string;
  available: number | null;
  required: number | null;
};

/** Parse thông báo thiếu tồn API: "Tồn thực tế SKU x / HĐ y chỉ còn a, không đủ để trừ b." */
const pxkUpParseShortageMessage = (message: string): PxkShortageRow | null => {
  const match = pxkUpChuoi(message).match(
    /SKU\s+(.+?)\s*\/\s*HĐ\s+(.+?)\s+chỉ còn\s+(\d+)\s*,\s*không đủ để trừ\s+(\d+)/i,
  );
  if (!match) return null;

  return {
    sku: pxkUpLamSach(match[1]),
    contract_no: pxkUpChuoi(match[2]),
    product_name: "",
    available: Number(match[3]),
    required: Number(match[4]),
  };
};

const pxkUpBuildShortageRows = (
  quantityErrors: string[],
  items: PxkUploadRow[],
): PxkShortageRow[] => {
  const nameByPair = new Map<string, string>();
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const key = `${pxkUpLamSach(item.sku)}|${pxkUpChuoi(item.contract_no)}`;
    if (!nameByPair.has(key) && item.product_name) {
      nameByPair.set(key, item.product_name);
    }
  }

  const rows: PxkShortageRow[] = [];
  const seen = new Set<string>();

  for (let i = 0; i < quantityErrors.length; i++) {
    const parsed = pxkUpParseShortageMessage(quantityErrors[i]);
    if (!parsed) continue;

    const key = `${parsed.sku}|${parsed.contract_no}`;
    if (seen.has(key)) continue;
    seen.add(key);

    rows.push({
      ...parsed,
      product_name: nameByPair.get(key) ?? "",
    });
  }

  return rows;
};

/** Tạo/ghi đè sheet liệt kê SKU thiếu tồn (SKU | Mã HĐ | Tên). */
const pxkUpWriteShortageSheet = (
  workbook: ExcelScript.Workbook,
  sheetName: string,
  rows: PxkShortageRow[],
): string | null => {
  if (rows.length === 0) return null;

  const safeName = sheetName.trim().slice(0, 31) || "SKU thiếu tồn";
  let sheet = workbook.getWorksheet(safeName);
  if (!sheet) {
    sheet = workbook.addWorksheet(safeName);
  } else {
    sheet.setVisibility(ExcelScript.SheetVisibility.visible);
    const oldRange = sheet.getUsedRange();
    if (oldRange) oldRange.clear(ExcelScript.ClearApplyTo.all);
  }

  const outSheet = workbook.getWorksheet(safeName);
  if (!outSheet) return null;

  const headers = ["SKU", "Mã hợp đồng", "Tên hàng", "Tồn còn", "Cần trừ"];
  const values: (string | number)[][] = [
    headers,
    ...rows.map((row) => [
      row.sku,
      row.contract_no,
      row.product_name,
      row.available ?? "",
      row.required ?? "",
    ]),
  ];

  const lastRow = values.length;
  outSheet.getRange(`A1:E${lastRow}`).setValues(values);

  const headerRange = outSheet.getRange("A1:E1");
  headerRange.getFormat().getFont().setBold(true);
  headerRange
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  headerRange.getFormat().getFill().setColor("#FCE4D6");

  outSheet.getRange(`A2:B${lastRow}`)
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  outSheet.getRange(`D2:E${lastRow}`)
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  outSheet.getFreezePanes().freezeRows(1);
  outSheet.getRange("A:E").getFormat().autofitColumns();

  return safeName;
};

// @ts-ignore: Office Scripts fetch
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
  json(): Promise<unknown>;
  text(): Promise<string>;
}>;

// @ts-ignore: Office Scripts — mỗi file chạy độc lập, không trùng main()
async function main(workbook: ExcelScript.Workbook) {
  const CFG = UPLOAD_PXK_CONFIG;

  const loiUrl = pxkUpKiemTraApiUrl(CFG.apiUrl);
  if (loiUrl) {
    console.log(`LỖI cấu hình: ${loiUrl}`);
    return;
  }

  if (!CFG.apiToken) {
    console.log("LỖI: Chưa cấu hình apiToken (Bearer token Sanctum).");
    return;
  }

  const sheet = workbook.getWorksheet(CFG.tenSheet);
  if (!sheet) {
    console.log(
      `LỖI: Không tìm thấy sheet "${CFG.tenSheet}". Hãy chạy Auto_PXK trên sheet sale trước.`,
    );
    return;
  }

  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    console.log(`Sheet "${CFG.tenSheet}" không có dữ liệu.`);
    return;
  }

  const rows = usedRange.getValues() as (string | number | boolean)[][];
  const items = pxkUpDocDuLieu(rows, CFG.headerKeywords);

  if (items.length === 0) {
    console.log(
      `Không tìm thấy dòng dữ liệu hợp lệ (tiêu đề cột A phải chứa ${CFG.headerKeywords.join(" + ")}).`,
    );
    return;
  }

  const meta = pxkUpDocMetaTuSheet(
    sheet,
    CFG.dispatchDateCell,
    CFG.dispatchNoCell,
    CFG.containerNoCell,
    CFG.sealNoCell,
    CFG.invoiceNoCell,
  );

  const dispatchNo = CFG.dispatchNo.trim() || meta.dispatchNo;
  const dispatchDateRaw = CFG.dispatchDate.trim() || meta.dispatchDateRaw;
  const dispatchDateIso =
    (CFG.dispatchDate.trim()
      ? pxkUpChuanHoaNgay(CFG.dispatchDate.trim())
      : null) ?? meta.dispatchDateIso;
  const containerNo = CFG.containerNo.trim() || meta.containerNo;
  const sealNo = CFG.sealNo.trim() || meta.sealNo;
  const invoiceNo = CFG.invoiceNo.trim() || meta.invoiceNo;

  if (!dispatchNo) {
    console.log(
      [
        `LỖI: Không đọc được số phiếu từ ô ${CFG.dispatchNoCell}.`,
        `Ô phải có dạng "NO: 000462" (ví dụ trong ${CFG.dispatchNoCell}).`,
      ].join("\n"),
    );
    return;
  }

  if (!dispatchDateIso) {
    console.log(
      [
        `LỖI: Không parse được ngày từ ô ${CFG.dispatchDateCell}` +
          (dispatchDateRaw ? `: "${dispatchDateRaw}"` : "."),
        "Hỗ trợ: 04 tháng 08 năm 2026, 2026年09月04, DD/MM/YYYY, YYYY-MM-DD, hoặc ngày Excel.",
      ].join("\n"),
    );
    return;
  }

  if (!containerNo) {
    console.log(
      `LỖI: Không đọc được mã cont từ ô ${CFG.containerNoCell}.`,
    );
    return;
  }

  const payload: Record<string, unknown> = {
    source_sheet: CFG.tenSheet,
    dispatch_no: dispatchNo,
    dispatch_date: dispatchDateIso,
    container_no: containerNo,
    auto_adjust_shortage: CFG.tuDongTaoPhieuDieuChinhTon,
    items: items.map((item) => pxkUpCompactItem(item)),
  };

  if (sealNo) payload.seal_no = sealNo;
  if (invoiceNo) payload.invoice_no = invoiceNo;

  const batDau = Date.now();
  console.log(
    [
      `Đang upload ${items.length} dòng Phiếu Xuất Kho (sheet "${CFG.tenSheet}")...`,
      `Số phiếu: ${dispatchNo} | Ngày: ${dispatchDateIso} | Cont: ${containerNo}` +
        (sealNo ? ` | Seal: ${sealNo}` : "") +
        (invoiceNo ? ` | Invoice: ${invoiceNo}` : ""),
    ].join("\n"),
  );

  try {
    const response = await fetch(CFG.apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Authorization: `Bearer ${CFG.apiToken}`,
      },
      body: JSON.stringify(payload),
    });

    const bodyText = await response.text();
    let body: {
      message?: string;
      imported?: number;
      skipped?: number;
      dispatch_no?: string;
      market_code?: string;
      dispatch_order_id?: number;
      container_no?: string;
      invoice_no?: string;
      stock_adjustment_id?: number;
      stock_adjustment_no?: string;
      stock_adjustment_quantity?: number;
      errors?: Record<string, string[]>;
    };
    try {
      body = JSON.parse(bodyText);
    } catch {
      body = { message: bodyText };
    }

    if (!response.ok) {
      const quantityErrors = body.errors?.quantity ?? [];
      const shortageRows = pxkUpBuildShortageRows(quantityErrors, items);
      const summary =
        shortageRows.length > 1
          ? `Có ${shortageRows.length} SKU không đủ tồn kho để xuất:`
          : (body.message ?? bodyText);
      const errorLines = [`LỖI HTTP ${response.status}: ${summary}`];
      if (body.errors) {
        errorLines.push(pxkUpFormatValidationErrors(body.errors));
      }

      if (shortageRows.length > 0) {
        const shortageSheetName = pxkUpWriteShortageSheet(
          workbook,
          CFG.shortageSheet,
          shortageRows,
        );
        if (shortageSheetName) {
          errorLines.push(
            `Đã ghi ${shortageRows.length} SKU thiếu tồn sang sheet "${shortageSheetName}".`,
          );
        }
      }

      console.log(errorLines.join("\n"));
      return;
    }

    const successLines = [
      `Upload thành công: ${body.imported ?? items.length} imported, ${body.skipped ?? 0} skipped (${Math.round((Date.now() - batDau) / 1000)}s).`,
    ];
    if (body.dispatch_no) {
      successLines.push(`Mã PXK: ${body.dispatch_no}`);
    }
    if (body.market_code) {
      successLines.push(`Thị trường: ${body.market_code}`);
    }
    if (body.container_no) {
      successLines.push(`Cont: ${body.container_no}`);
    }
    if (body.invoice_no) {
      successLines.push(`Invoice: ${body.invoice_no}`);
    }
    if (body.stock_adjustment_no) {
      successLines.push(
        `Đã tự động tạo phiếu điều chỉnh ${body.stock_adjustment_no}` +
          ` (+${body.stock_adjustment_quantity ?? 0}).`,
      );
    }
    successLines.push(body.message ?? "Hoàn tất.");
    console.log(successLines.join("\n"));
  } catch (err) {
    const msg = String(err);
    if (msg.includes("Failed to fetch")) {
      console.log(
        [
          "LỖI kết nối API: Failed to fetch — kiểm tra apiUrl HTTPS public và server đã deploy.",
          `Kiểm tra apiUrl hiện tại: ${CFG.apiUrl}`,
        ].join("\n"),
      );
    } else {
      console.log(`LỖI kết nối API: ${msg}`);
    }
  }
}

const pxkUpTestables = {
  pxkUpTrichXuatSoPhieu,
  pxkUpTrichXuatMa,
  pxkUpChuanHoaNgay,
  pxkUpParseMotDongNgay,
  pxkUpDocDuLieu,
  pxkUpParseShortageMessage,
  pxkUpBuildShortageRows,
};

(
  globalThis as typeof globalThis & {
    __UPLOAD_PXK_TESTABLES__?: typeof pxkUpTestables;
  }
).__UPLOAD_PXK_TESTABLES__ = pxkUpTestables;
