/**
 * ===================================================================
 * Upload_PrePackingList – Đọc sheet "pre-packing-list" (format Auto_PXK) và upload lên API.
 * ===================================================================
 *
 * Cột dữ liệu (A→R, giống Auto_PXK.ts):
 * A STT | B Hợp đồng | C SKU | D Tên hàng | E Kích thước
 * F SL xuất đi | G ĐVT | H Đã xuất | I SL đóng gói | J ĐVT đóng gói
 * K Số lô | L Trọng lượng | M Tổng TL | N Thể tích | O Tổng TT
 * P Vị trí | Q Ghi chú | R PMC
 *
 * Luồng sử dụng:
 * 1. Chuẩn bị/chạy Auto_PXK trên sheet "pre-packing-list" (chỉ format).
 * 2. Điền ô D3 (ngày lập) và D4 (NO: số phiếu) trên sheet.
 * 3. Tạo API token tại Dashboard → API Token (Excel).
 * 4. Cấu hình apiUrl (HTTPS public) và apiToken bên dưới.
 * 5. Chạy script này trong Excel Online.
 */

const UPLOAD_PRE_PACKING_CONFIG = {
  tenSheet: "pre-packing-list",
  headerKeywords: ["STT", "序号"],
  firstDataCol: "B",
  lastDataCol: "R",
  shortageSheet: "SKU thiếu tồn",
  /** Ô chứa ngày lập (vd: "27 tháng 08 năm 2026" / "2026年08月27"). */
  listDateCell: "D3",
  /** Ô chứa số phiếu (vd: "NO:  000463"). */
  listNoCell: "D4",
  apiUrl: "https://app.wgr-wh.site/api/v1/warehouse/pre-packing-lists/import",
  /** Để trống trên repo; người dùng tự điền token khi dán script vào Excel. */
  apiToken: "",
  /** Tuỳ chọn ghi đè thủ công. Để trống = tự đọc từ listDateCell / listNoCell. */
  listNo: "",
  listDate: "",
  /**
   * false: phiếu đã có SKU bổ sung trên hệ thống mà file thiếu thì bị từ chối (chạy Sync_PrePackingList trước).
   * true: thay thế toàn bộ theo file, các SKU bổ sung đó bị loại bỏ.
   */
  replaceAll: false,
};

type PrePackingRow = {
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
const prePackLamSach = (giaTri: string | number | boolean): string =>
  String(giaTri ?? "")
    .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "")
    .trim();

// @ts-ignore: Tránh cảnh báo trùng lặp hàm khi code nhiều file TS trong cùng thư mục
const prePackChuoi = (giaTri: string | number | boolean): string =>
  String(giaTri ?? "").trim();

// @ts-ignore: Tránh cảnh báo trùng lặp hàm khi code nhiều file TS trong cùng thư mục
const prePackParseSo = (giaTri: string | number | boolean): number => {
  if (typeof giaTri === "number" && !isNaN(giaTri)) return giaTri;
  let s = prePackLamSach(giaTri);
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

const prePackParseSoHoacNull = (
  giaTri: string | number | boolean,
): number | null => {
  const n = prePackParseSo(giaTri);
  return n === 0 && prePackLamSach(giaTri) === "" ? null : n;
};

const prePackChuanHoaSku = (sku: string): string => {
  const cleaned = prePackLamSach(sku);
  // SKU = toàn bộ dãy số đầu mã (≥ 6 số); cặp A/B (912345A / 912345B) là 2 sản phẩm riêng nên giữ nguyên.
  if (/^9\d{5,}[AB]$/i.test(cleaned)) return cleaned.toUpperCase();
  const match = cleaned.match(/^(9\d{5,})(\D.*)$/);
  return match ? match[1] : cleaned;
};

/** Định dạng Date → YYYY-MM-DD (UTC). */
const prePackFormatIsoDate = (date: Date): string | null => {
  if (isNaN(date.getTime())) return null;
  const yyyy = date.getUTCFullYear();
  const mm = String(date.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(date.getUTCDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
};

/** Chuyển serial ngày Excel sang Date (UTC). */
const prePackExcelSerialToDate = (serial: number): Date => {
  const wholeDays = Math.floor(serial);
  const excelEpoch = Date.UTC(1899, 11, 30);
  return new Date(excelEpoch + wholeDays * 86400000);
};

/** Parse một dòng text ngày sang YYYY-MM-DD. */
const prePackParseMotDongNgay = (line: string): string | null => {
  const s = prePackChuoi(line);
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

  const zhMatch = s.match(/(\d{4})年(\d{1,2})月(\d{1,2})/);
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

/** Trích số phiếu từ ô D4 (vd: "NO:  000463"). */
const prePackTrichXuatSoPhieu = (raw: string | number | boolean): string => {
  const s = prePackChuoi(raw);
  if (!s) return "";

  const noMatch = s.match(/NO\s*[：:]\s*(\S+)/i);
  if (noMatch) return noMatch[1].trim();

  const digits = s.match(/(\d[\d\s]*\d|\d+)/);
  if (digits) return digits[1].replace(/\s+/g, "");

  return s;
};

/** Chuyển ngày sang YYYY-MM-DD cho API Laravel. */
const prePackChuanHoaNgayListDate = (
  raw: string | number | boolean,
): string | null => {
  if (typeof raw === "number" && !isNaN(raw)) {
    if (raw >= 30000 && raw <= 60000) {
      return prePackFormatIsoDate(prePackExcelSerialToDate(raw));
    }
  }

  const s = prePackChuoi(raw);
  if (!s) return null;

  const lines = s
    .split(/[\r\n]+/)
    .map((line) => line.trim())
    .filter((line) => Boolean(line));
  const candidates = lines.length > 0 ? lines : [s];

  for (let i = 0; i < candidates.length; i++) {
    const parsed = prePackParseMotDongNgay(candidates[i]);
    if (parsed) return parsed;
  }

  return null;
};

/** Đọc meta từ ô D3 (ngày) và D4 (số phiếu) trên sheet. */
const prePackDocMetaTuSheet = (
  sheet: ExcelScript.Worksheet,
  dateCell: string,
  noCell: string,
): { listDateRaw: string; listNo: string; listDateIso: string | null } => {
  const dateVal = sheet.getRange(dateCell).getValue();
  const noVal = sheet.getRange(noCell).getValue();

  return {
    listDateRaw: prePackChuoi(dateVal),
    listNo: prePackTrichXuatSoPhieu(noVal),
    listDateIso: prePackChuanHoaNgayListDate(dateVal),
  };
};

// @ts-ignore: Tránh trùng findHeaderRow trong Auto_PXK.ts (khác signature)
const prePackFindHeaderRow = (
  rows: (string | number | boolean)[][],
  keywords: string[],
): number => {
  for (let i = 0; i < rows.length; i++) {
    const cellValue = prePackChuoi(rows[i][0]).toUpperCase();
    const allMatch = keywords.every((kw) =>
      cellValue.includes(kw.toUpperCase()),
    );
    if (allMatch) return i;
  }
  return -1;
};

// @ts-ignore: Tránh trùng findDataBounds trong Auto_PXK.ts (khác signature)
const prePackFindDataBounds = (
  rows: (string | number | boolean)[][],
  startRow: number,
): { startRow: number; lastRow: number } | null => {
  if (startRow >= rows.length) return null;

  let lastDataRow = -1;
  for (let i = startRow; i < rows.length; i++) {
    const productName = prePackChuoi(rows[i][3] ?? "");
    if (productName !== "") {
      lastDataRow = i;
      continue;
    }

    const rowText = rows[i]
      .slice(0, 18)
      .map((c) => prePackChuoi(c))
      .join(" ")
      .toUpperCase();
    if (rowText.includes("SUM(") || rowText.includes("TỔNG")) {
      break;
    }
  }

  if (lastDataRow < startRow) return null;

  return { startRow, lastRow: lastDataRow };
};

const prePackDocDuLieu = (
  rows: (string | number | boolean)[][],
  headerKeywords: string[],
): PrePackingRow[] => {
  const headerIdx = prePackFindHeaderRow(rows, headerKeywords);
  if (headerIdx === -1) return [];

  const bounds = prePackFindDataBounds(rows, headerIdx + 1);
  if (!bounds) return [];

  const items: PrePackingRow[] = [];

  for (let i = bounds.startRow; i <= bounds.lastRow; i++) {
    const row = rows[i];
    const sku = prePackChuanHoaSku(String(row[2] ?? ""));
    const contractNo = prePackChuoi(row[1] ?? "");
    const productName = prePackChuoi(row[3] ?? "");

    if (!sku || !contractNo || !productName) continue;

    const dispatchQty = Math.round(prePackParseSo(row[5] ?? 0));
    const packedQty = Math.round(prePackParseSo(row[8] ?? 0));
    if (dispatchQty <= 0 && packedQty <= 0) continue;

    items.push({
      stt: Math.round(prePackParseSo(row[0] ?? 0)) || items.length + 1,
      contract_no: contractNo,
      sku,
      product_name: productName,
      dimensions: prePackChuoi(row[4] ?? ""),
      dispatch_quantity: dispatchQty,
      unit: prePackChuoi(row[6] ?? ""),
      previously_shipped_quantity: Math.round(prePackParseSo(row[7] ?? 0)),
      packed_quantity: packedQty,
      pack_unit: prePackChuoi(row[9] ?? ""),
      batch_no: prePackChuoi(row[10] ?? ""),
      unit_weight: prePackParseSoHoacNull(row[11] ?? ""),
      total_weight: prePackParseSoHoacNull(row[12] ?? ""),
      unit_volume: prePackParseSoHoacNull(row[13] ?? ""),
      total_volume: prePackParseSoHoacNull(row[14] ?? ""),
      location_notes: prePackChuoi(row[15] ?? ""), // P: Vị trí
      notes: prePackChuoi(row[16] ?? ""), // Q: Ghi chú
      pmc_approved: prePackChuoi(row[17] ?? ""),
    });
  }

  return items;
};

// @ts-ignore: Tránh trùng compactItem trong Upload_Sanji.ts
const prePackCompactItem = (
  item: PrePackingRow,
): Record<string, string | number> => {
  const out: Record<string, string | number> = {};
  const keys: (keyof PrePackingRow)[] = [
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

// @ts-ignore: Tránh trùng kiemTraApiUrl trong Upload_Sanji.ts
const prePackKiemTraApiUrl = (url: string): string | null => {
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

const prePackFormatValidationErrors = (
  errors: Record<string, string[]>,
): string => {
  return Object.keys(errors)
    .map((field) => `  - ${field}: ${errors[field].join(", ")}`)
    .join("\n");
};

type PrePackSystemItem = {
  sku: string;
  contract_no: string;
  product_name: string;
};

const prePackDongSystemItem = (item: PrePackSystemItem): string =>
  `  - ${[item.sku, item.contract_no, item.product_name]
    .map((part) => prePackChuoi(part))
    .filter((part) => part !== "")
    .join(" · ")}`;

/** Thông báo cho phản hồi 409 system_lines_missing: phiếu có SKU bổ sung trên hệ thống mà file thiếu. */
const prePackFormatMissingSystemItems = (body: {
  message?: string;
  missing_items?: PrePackSystemItem[];
}): string =>
  [
    `LỖI HTTP 409: ${body.message ?? "File Excel thiếu SKU bổ sung trên hệ thống."}`,
    ...(body.missing_items ?? []).map((item) => prePackDongSystemItem(item)),
    "Cách xử lý: chạy Sync_PrePackingList để kéo các SKU này về Excel rồi upload lại,",
    "hoặc đặt replaceAll: true trong cấu hình nếu muốn thay thế toàn bộ theo file.",
  ].join("\n");

/** Danh sách SKU bổ sung trên hệ thống đã bị loại bỏ khi upload với replaceAll. */
const prePackFormatRemovedSystemItems = (body: {
  removed_system_items?: PrePackSystemItem[];
}): string[] => {
  const items = body.removed_system_items ?? [];
  if (items.length === 0) return [];
  return [
    `Đã loại bỏ ${items.length} SKU bổ sung trên hệ thống:`,
    ...items.map((item) => prePackDongSystemItem(item)),
  ];
};

type PrePackShortageRow = {
  sku: string;
  contract_no: string;
  product_name: string;
  available: number | null;
  required: number | null;
};

/** Parse thông báo thiếu tồn API: "Tồn thực tế SKU x / HĐ y chỉ còn a, không đủ để trừ b." */
const prePackParseShortageMessage = (
  message: string,
): PrePackShortageRow | null => {
  const match = prePackChuoi(message).match(
    /SKU\s+(.+?)\s*\/\s*HĐ\s+(.+?)\s+chỉ còn\s+(\d+)\s*,\s*không đủ để trừ\s+(\d+)/i,
  );
  if (!match) return null;

  return {
    sku: prePackLamSach(match[1]),
    contract_no: prePackChuoi(match[2]),
    product_name: "",
    available: Number(match[3]),
    required: Number(match[4]),
  };
};

const prePackBuildShortageRows = (
  quantityErrors: string[],
  items: PrePackingRow[],
): PrePackShortageRow[] => {
  const nameByPair = new Map<string, string>();
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const key = `${prePackLamSach(item.sku)}|${prePackChuoi(item.contract_no)}`;
    if (!nameByPair.has(key) && item.product_name) {
      nameByPair.set(key, item.product_name);
    }
  }

  const rows: PrePackShortageRow[] = [];
  const seen = new Set<string>();

  for (let i = 0; i < quantityErrors.length; i++) {
    const parsed = prePackParseShortageMessage(quantityErrors[i]);
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

/** Tạo/ghi đè sheet liệt kê SKU thiếu tồn (SKU | Mã HĐ | Tên | Tồn còn | Cần trừ). */
const prePackWriteShortageSheet = (
  workbook: ExcelScript.Workbook,
  sheetName: string,
  rows: PrePackShortageRow[],
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
  const CFG = UPLOAD_PRE_PACKING_CONFIG;

  const loiUrl = prePackKiemTraApiUrl(CFG.apiUrl);
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
      `LỖI: Không tìm thấy sheet "${CFG.tenSheet}". Hãy chuẩn bị sheet Pre-packing-list trước.`,
    );
    return;
  }

  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    console.log(`Sheet "${CFG.tenSheet}" không có dữ liệu.`);
    return;
  }

  const rows = usedRange.getValues() as (string | number | boolean)[][];
  const items = prePackDocDuLieu(rows, CFG.headerKeywords);

  if (items.length === 0) {
    console.log(
      `Không tìm thấy dòng dữ liệu hợp lệ (tiêu đề cột A phải chứa ${CFG.headerKeywords.join(" + ")}).`,
    );
    return;
  }

  const meta = prePackDocMetaTuSheet(sheet, CFG.listDateCell, CFG.listNoCell);
  const listNo = CFG.listNo.trim() || meta.listNo;
  const listDateRaw = CFG.listDate.trim() || meta.listDateRaw;
  const listDateIso =
    (CFG.listDate.trim()
      ? prePackChuanHoaNgayListDate(CFG.listDate.trim())
      : null) ?? meta.listDateIso;

  if (!listNo) {
    console.log(
      [
        `LỖI: Không đọc được số phiếu từ ô ${CFG.listNoCell}.`,
        `Ô phải có dạng "NO: 000463" (ví dụ trong ${CFG.listNoCell}).`,
      ].join("\n"),
    );
    return;
  }

  if (listDateRaw && !listDateIso) {
    console.log(
      [
        `LỖI: Không parse được ngày từ ô ${CFG.listDateCell}: "${listDateRaw}".`,
        "Hỗ trợ: 27 tháng 08 năm 2026, 2026年08月27, DD/MM/YYYY, YYYY-MM-DD, hoặc ngày Excel.",
      ].join("\n"),
    );
    return;
  }

  const payload: Record<string, unknown> = {
    source_sheet: CFG.tenSheet,
    list_no: listNo,
    items: items.map((item) => prePackCompactItem(item)),
  };

  if (listDateIso) payload.list_date = listDateIso;
  if (CFG.replaceAll) payload.replace_all = true;

  const batDau = Date.now();
  console.log(
    [
      `Đang upload ${items.length} dòng Pre-packing-list...`,
      `Số phiếu: ${listNo}${listDateIso ? ` | Ngày: ${listDateIso}` : ""}`,
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
      list_no?: string;
      market_code?: string;
      pre_packing_list_id?: number;
      errors?: Record<string, string[]>;
      code?: string;
      missing_items?: PrePackSystemItem[];
      removed_system_items?: PrePackSystemItem[];
    };
    try {
      body = JSON.parse(bodyText);
    } catch {
      body = { message: bodyText };
    }

    if (response.status === 409 && body.code === "system_lines_missing") {
      console.log(prePackFormatMissingSystemItems(body));
      return;
    }

    if (!response.ok) {
      const quantityErrors = body.errors?.quantity ?? [];
      const shortageRows = prePackBuildShortageRows(quantityErrors, items);
      const summary =
        shortageRows.length > 1
          ? `Có ${shortageRows.length} SKU không đủ tồn kho để giữ hàng:`
          : (body.message ?? bodyText);
      const errorLines = [`LỖI HTTP ${response.status}: ${summary}`];
      if (body.errors) {
        errorLines.push(prePackFormatValidationErrors(body.errors));
      }

      if (shortageRows.length > 0) {
        const shortageSheetName = prePackWriteShortageSheet(
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
    if (body.list_no) {
      successLines.push(`Mã Pre-packing-list: ${body.list_no}`);
    }
    if (body.market_code) {
      successLines.push(`Thị trường: ${body.market_code}`);
    }
    successLines.push(...prePackFormatRemovedSystemItems(body));
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

const prePackTestables = {
  prePackTrichXuatSoPhieu,
  prePackChuanHoaNgayListDate,
  prePackParseMotDongNgay,
  prePackParseShortageMessage,
  prePackBuildShortageRows,
  prePackFormatMissingSystemItems,
  prePackFormatRemovedSystemItems,
};

(
  globalThis as typeof globalThis & {
    __UPLOAD_PRE_PACKING_TESTABLES__?: typeof prePackTestables;
  }
).__UPLOAD_PRE_PACKING_TESTABLES__ = prePackTestables;
