/**
 * ===================================================================
 * Upload_Nami – Đọc sheet kết quả Auto_Nami (KQ_NhapKho_*) và upload nhập kho lên API.
 * ===================================================================
 *
 * CÁCH CŨ — dùng Upload_Nami_Super để tải lên một bước (đọc thẳng "Thành phẩm",
 * không cần chạy Auto_Nami trước). Hai cách dùng chung khoá chống trùng nên không bị nhập hai lần.
 *
 * Cột dữ liệu (A→T, giống output Auto_Nami.ts — header hàng 2):
 * A Nơi nhập | B Ngày nhập kho | C Mã hợp đồng | D Thị trường | E Mã phân loại
 * F SKU | G Tên hàng | H SL đơn hàng | I SL đã nhập | J SL nhập | K SL còn lại
 * L Mã ghép | M Trọng lượng | N Tổng TL | O Dài | P Rộng | Q Cao | R Thể tích | S Tổng TT
 * T Vị trí | U (trống — khoảng cách) | V… bảng thống kê Auto_Nami
 *
 * Upload chỉ đọc A–T (tìm cột "Vị trí" theo header); cột trống U không ảnh hưởng.
 * Cột H "SL đơn hàng" chọn dòng đơn hàng đích (SL đặt phải bằng đúng); dòng không khớp không được nhập
 * và được tô vàng A..Vị trí sau khi upload. Mỗi lần upload thành công, nền vàng cũ được bỏ trước khi tô lại.
 * QUAN TRỌNG — Office Scripts chạy trên cloud Microsoft (Excel Online):
 * - KHÔNG dùng localhost / 127.0.0.1 — máy bạn không thể truy cập từ cloud.
 * - BẮT BUỘC dùng HTTPS (http:// sẽ bị chặn).
 *
 * Luồng sử dụng:
 * 1. Chạy Auto_Nami để tạo sheet KQ_NhapKho_*.
 * 2. Mở sheet kết quả cần upload (hoặc cấu hình tenSheet bên dưới).
 * 3. Tạo API token tại Dashboard → Hệ thống & Quyền → API Token (Excel).
 * 4. Cấu hình apiUrl (HTTPS, public) và apiToken bên dưới.
 * 5. Chạy script này trong Excel Online.
 */

const UPLOAD_NAMI_CONFIG = {
  /** Để trống = sheet đang active; hoặc ghi tên sheet cụ thể (vd: "KQ_NhapKho_25-05-2026"). */
  tenSheet: "",
  /** Tiền tố tên sheet do Auto_Nami tạo — dùng khi tenSheet trống và active sheet không khớp. */
  sheetNamePrefix: "KQ_NhapKho_",
  headerRow: 2,
  apiUrl: "https://app.wgr-wh.site/api/v1/warehouse/receipts/nami-import",
  apiToken: "", // Bearer token từ Dashboard
};

type NamiInboundRow = {
  source_row: number;
  receipt_place: string;
  round_number: number;
  receipt_date: string;
  contract_no: string;
  market_code: string;
  classification_code: string;
  sku: string;
  product_name: string;
  order_quantity: number;
  total_received: number;
  quantity: number;
  remaining_quantity: number;
  merge_code: string;
  unit_weight: number | null;
  total_weight: number | null;
  length: number | null;
  width: number | null;
  height: number | null;
  unit_volume: number | null;
  total_volume: number | null;
  location_code: string;
};

type NamiSkippedRow = {
  source_row: number;
  sku: string;
  contract_no: string;
  reason: string;
};

type NamiRejectedRow = {
  row: number;
  sku: string;
  contract_no: string;
  order_quantity?: number;
  code?: string;
  reason: string;
};

type NamiVungDuLieu = {
  /** Số dòng trên sheet (1-based) của các dòng dữ liệu. */
  dongs: number[];
  /** Số cột từ A đến cột "Vị trí". */
  soCot: number;
};

/** Màu nền script dùng để đánh dấu dòng không khớp SL đơn hàng. */
const NAMI_UP_MAU_VANG = "#FFFF00";
const NAMI_UP_MA_KHONG_KHOP = "order_quantity_unmatched";

// @ts-ignore: Tránh cảnh báo trùng lặp hàm khi code nhiều file TS trong cùng thư mục
const namiUpLamSach = (giaTri: string | number | boolean): string =>
  String(giaTri ?? "")
    .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "")
    .trim();

// @ts-ignore
const namiUpChuoi = (giaTri: string | number | boolean): string =>
  String(giaTri ?? "").trim();

// @ts-ignore
const namiUpParseSo = (giaTri: string | number | boolean): number => {
  if (typeof giaTri === "number" && !isNaN(giaTri)) return giaTri;
  let s = namiUpLamSach(giaTri);
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

const namiUpParseSoHoacNull = (
  giaTri: string | number | boolean,
): number | null => {
  const n = namiUpParseSo(giaTri);
  return n === 0 && namiUpLamSach(giaTri) === "" ? null : n;
};

/** Gửi SKU nguyên dạng; backend đối chiếu với đơn hàng của hợp đồng để chọn đúng mã (A/B, hậu tố...). */
const namiUpChuanHoaSku = (sku: string): string =>
  namiUpLamSach(sku).toUpperCase();

/** Chuẩn hóa vị trí kho: "A05  B44" → "A5 B44". */
const namiUpLayMaViTri = (viTri: string): string => {
  const cleaned = namiUpChuoi(viTri)
    .toUpperCase()
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!cleaned) return "";

  const cum: string[] = [];
  const regex = /([A-Z]+)\s*0*(\d+)/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(cleaned)) !== null) {
    cum.push(`${match[1]}${Number(match[2])}`);
  }
  if (cum.length > 0) return cum.join(" ");

  return cleaned.split(" ")[0] ?? "";
};

const namiUpTimChiSoCot = (
  headerRow: (string | number | boolean)[],
  tenCot: string,
): number => {
  const needle = tenCot.toLowerCase().trim();
  for (let i = 0; i < headerRow.length; i++) {
    if (namiUpChuoi(headerRow[i]).toLowerCase() === needle) return i;
  }
  for (let i = 0; i < headerRow.length; i++) {
    if (namiUpChuoi(headerRow[i]).toLowerCase().includes(needle)) return i;
  }
  return -1;
};

/** Ưu tiên text hiển thị (getTexts); fallback raw getValues — tránh ô wrap-text / tham chiếu bị rỗng. */
const namiUpLayOChuoi = (
  textRow: (string | number | boolean)[] | undefined,
  valueRow: (string | number | boolean)[],
  idx: number,
): string => {
  if (idx < 0) return "";
  const fromText = namiUpChuoi(textRow?.[idx] ?? "");
  if (fromText) return fromText;
  return namiUpChuoi(valueRow[idx] ?? "");
};

const namiUpFormatIsoDate = (date: Date): string | null => {
  if (isNaN(date.getTime())) return null;
  const yyyy = date.getUTCFullYear();
  const mm = String(date.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(date.getUTCDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
};

const namiUpExcelSerialToDate = (serial: number): Date => {
  const wholeDays = Math.floor(serial);
  const excelEpoch = Date.UTC(1899, 11, 30);
  return new Date(excelEpoch + wholeDays * 86400000);
};

const namiUpChuanHoaNgay = (
  raw: string | number | boolean,
): string | null => {
  if (typeof raw === "number" && !isNaN(raw)) {
    if (raw >= 30000 && raw <= 60000) {
      return namiUpFormatIsoDate(namiUpExcelSerialToDate(raw));
    }
  }

  const s = namiUpChuoi(raw);
  if (!s) return null;

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

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

const namiUpTrichRoundNumber = (receiptPlace: string): number => {
  const match = namiUpChuoi(receiptPlace).match(/(\d+)\s*$/);
  if (match) {
    const n = Number(match[1]);
    if (n >= 1 && n <= 10) return n;
  }
  return 1;
};

const namiUpLaDongThongKe = (row: (string | number | boolean)[]): boolean => {
  const colA = namiUpChuoi(row[0] ?? "").toUpperCase();
  return colA.startsWith("THỐNG KÊ") || colA.startsWith("THONG KE");
};

const namiUpLaDongHeader = (row: (string | number | boolean)[]): boolean => {
  const colA = namiUpChuoi(row[0]).toLowerCase();
  const colB = namiUpChuoi(row[1]).toLowerCase();
  return colA === "nơi nhập" && colB.includes("ngày nhập");
};

const namiUpTimDongHeader = (
  rows: (string | number | boolean)[][],
  headerRow: number,
): number => {
  const found = rows.findIndex((row) => namiUpLaDongHeader(row));
  return found >= 0 ? found : headerRow - 1;
};

const namiUpTimCotViTri = (header: (string | number | boolean)[]): number => {
  const byName = namiUpTimChiSoCot(header, "vị trí");
  if (byName >= 0) return byName;
  return header.length > 19 ? 19 : -1;
};

/** Các dòng dữ liệu (dưới header, trước bảng "THỐNG KÊ") và độ rộng A..Vị trí để tô/bỏ tô. */
const namiUpVungDuLieu = (
  rows: (string | number | boolean)[][],
  headerRow: number,
  rowOffset = 0,
): NamiVungDuLieu => {
  const headerIdx = namiUpTimDongHeader(rows, headerRow);
  if (headerIdx < 0 || headerIdx >= rows.length) return { dongs: [], soCot: 0 };

  const locationColIdx = namiUpTimCotViTri(rows[headerIdx]);
  const dongs: number[] = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (namiUpLaDongThongKe(row)) break;
    if (!namiUpChuoi(row[0] ?? "") || !namiUpLamSach(String(row[5] ?? "")) || !namiUpChuoi(row[2] ?? "")) {
      continue;
    }
    dongs.push(rowOffset + i + 1);
  }

  return { dongs, soCot: locationColIdx >= 0 ? locationColIdx + 1 : 20 };
};

const namiUpDocDuLieu = (
  rows: (string | number | boolean)[][],
  headerRow: number,
  textRows?: (string | number | boolean)[][],
  rowOffset = 0,
  skipped?: NamiSkippedRow[],
): NamiInboundRow[] => {
  const headerIdx = namiUpTimDongHeader(rows, headerRow);

  if (headerIdx < 0 || headerIdx >= rows.length) return [];

  const header = rows[headerIdx];
  const locationColIdx = namiUpTimCotViTri(header);

  const items: NamiInboundRow[] = [];

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (namiUpLaDongThongKe(row)) break;

    const receiptPlace = namiUpChuoi(row[0] ?? "");
    const sku = namiUpChuanHoaSku(String(row[5] ?? ""));
    const contractNo = namiUpChuoi(row[2] ?? "");
    const productName = namiUpChuoi(row[6] ?? "");

    if (!receiptPlace || !sku || !contractNo) continue;

    const sourceRow = rowOffset + i + 1;
    const boQua = (reason: string) => {
      skipped?.push({
        source_row: sourceRow,
        sku,
        contract_no: contractNo,
        reason,
      });
    };

    const quantityRaw = row[9] ?? "";
    if (namiUpLamSach(quantityRaw) === "") {
      boQua("ô SL nhập (cột J) trống");
      continue;
    }
    const quantity = Math.round(namiUpParseSo(quantityRaw));
    if (quantity <= 0) {
      boQua(`SL nhập (cột J) không hợp lệ: "${namiUpChuoi(quantityRaw)}"`);
      continue;
    }

    const totalReceivedRaw = row[8] ?? "";
    if (namiUpLamSach(totalReceivedRaw) === "") {
      boQua("ô SL đã nhập (cột I) trống");
      continue;
    }
    const totalReceived = Math.round(namiUpParseSo(totalReceivedRaw));
    if (totalReceived < quantity) {
      boQua(
        `SL đã nhập (cột I = ${totalReceived}) nhỏ hơn SL nhập (cột J = ${quantity})`,
      );
      continue;
    }

    const receiptDate = namiUpChuanHoaNgay(row[1] ?? "");
    if (!receiptDate) {
      boQua(`ngày nhập kho (cột B) không hợp lệ: "${namiUpChuoi(row[1] ?? "")}"`);
      continue;
    }

    items.push({
      source_row: sourceRow,
      receipt_place: receiptPlace,
      round_number: namiUpTrichRoundNumber(receiptPlace),
      receipt_date: receiptDate,
      contract_no: contractNo,
      market_code: namiUpChuoi(row[3] ?? "").toUpperCase(),
      classification_code: namiUpChuoi(row[4] ?? "").toUpperCase(),
      sku,
      product_name: productName,
      order_quantity: Math.round(namiUpParseSo(row[7] ?? 0)),
      total_received: totalReceived,
      quantity,
      remaining_quantity: Math.round(namiUpParseSo(row[10] ?? 0)),
      merge_code: namiUpChuoi(row[11] ?? ""),
      unit_weight: namiUpParseSoHoacNull(row[12] ?? ""),
      total_weight: namiUpParseSoHoacNull(row[13] ?? ""),
      length: namiUpParseSoHoacNull(row[14] ?? ""),
      width: namiUpParseSoHoacNull(row[15] ?? ""),
      height: namiUpParseSoHoacNull(row[16] ?? ""),
      unit_volume: namiUpParseSoHoacNull(row[17] ?? ""),
      total_volume: namiUpParseSoHoacNull(row[18] ?? ""),
      location_code: namiUpLayMaViTri(
        namiUpLayOChuoi(textRows?.[i], row, locationColIdx),
      ),
    });
  }

  return items;
};

const namiUpCompactItem = (
  item: NamiInboundRow,
): Record<string, string | number> => {
  const out: Record<string, string | number> = {};
  const keys: (keyof NamiInboundRow)[] = [
    "source_row",
    "receipt_place",
    "round_number",
    "receipt_date",
    "contract_no",
    "market_code",
    "classification_code",
    "sku",
    "product_name",
    "order_quantity",
    "total_received",
    "quantity",
    "remaining_quantity",
    "merge_code",
    "unit_weight",
    "total_weight",
    "length",
    "width",
    "height",
    "unit_volume",
    "total_volume",
    "location_code",
  ];

  keys.forEach((key) => {
    const val = item[key];
    if (val !== null && val !== undefined && val !== "") {
      out[key] = val;
    }
  });

  return out;
};

const namiUpKiemTraApiUrl = (url: string): string | null => {
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

const namiUpTimSheetKetQua = (
  workbook: ExcelScript.Workbook,
  tenSheet: string,
  prefix: string,
): ExcelScript.Worksheet | undefined => {
  if (tenSheet) {
    return workbook.getWorksheet(tenSheet);
  }

  const active = workbook.getActiveWorksheet();
  if (active && active.getName().startsWith(prefix)) {
    return active;
  }

  const sheets = workbook.getWorksheets();
  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    if (name.startsWith(prefix)) {
      return sheets[i];
    }
  }

  return active ?? undefined;
};

/**
 * Bỏ nền vàng do lần chạy trước tô (chỉ ô màu NAMI_UP_MAU_VANG ở cột A) rồi tô vàng A..Vị trí
 * các dòng không khớp SL đơn hàng. Trả về số dòng đã tô.
 */
const namiUpToVangKhongKhop = (
  sheet: ExcelScript.Worksheet,
  vung: NamiVungDuLieu,
  colOffset: number,
  khongKhop: NamiRejectedRow[],
): number => {
  if (vung.dongs.length === 0 || vung.soCot <= 0) return 0;

  const laDongDuLieu: { [dong: number]: boolean } = {};

  vung.dongs.forEach((dong) => {
    laDongDuLieu[dong] = true;
    const mau = sheet.getRangeByIndexes(dong - 1, colOffset, 1, 1).getFormat().getFill().getColor() ?? "";
    if (mau.toUpperCase() === NAMI_UP_MAU_VANG) {
      sheet.getRangeByIndexes(dong - 1, colOffset, 1, vung.soCot).getFormat().getFill().clear();
    }
  });

  let daTo = 0;
  khongKhop.forEach((r) => {
    if (!laDongDuLieu[r.row]) return;
    sheet
      .getRangeByIndexes(r.row - 1, colOffset, 1, vung.soCot)
      .getFormat()
      .getFill()
      .setColor(NAMI_UP_MAU_VANG);
    daTo++;
  });

  return daTo;
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

// @ts-ignore: Office Scripts
async function main(workbook: ExcelScript.Workbook) {
  const CFG = UPLOAD_NAMI_CONFIG;

  const loiUrl = namiUpKiemTraApiUrl(CFG.apiUrl);
  if (loiUrl) {
    console.log(`LỖI cấu hình: ${loiUrl}`);
    return;
  }

  if (!CFG.apiToken) {
    console.log("LỖI: Chưa cấu hình apiToken (Bearer token Sanctum).");
    return;
  }

  const sheet = namiUpTimSheetKetQua(
    workbook,
    CFG.tenSheet,
    CFG.sheetNamePrefix,
  );
  if (!sheet) {
    console.log(
      `LỖI: Không tìm thấy sheet kết quả Auto_Nami (tiền tố "${CFG.sheetNamePrefix}"). Hãy chạy Auto_Nami trước.`,
    );
    return;
  }

  const sheetName = sheet.getName();
  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    console.log(`Sheet "${sheetName}" không có dữ liệu.`);
    return;
  }

  const rows = usedRange.getValues() as (string | number | boolean)[][];
  const textRows = usedRange.getTexts() as (string | number | boolean)[][];
  const skippedRows: NamiSkippedRow[] = [];
  const items = namiUpDocDuLieu(
    rows,
    CFG.headerRow,
    textRows,
    usedRange.getRowIndex(),
    skippedRows,
  );

  if (skippedRows.length > 0) {
    console.log(
      `CẢNH BÁO: ${skippedRows.length} dòng có SKU nhưng KHÔNG upload do thiếu/sai số lượng hoặc ngày:`,
    );
    skippedRows.forEach((r) => {
      console.log(
        `- Dòng ${r.source_row}: SKU ${r.sku} / HĐ ${r.contract_no} — ${r.reason}`,
      );
    });
  }

  if (items.length === 0) {
    console.log(
      `Không tìm thấy dòng nhập kho hợp lệ trên sheet "${sheetName}" (header: Nơi nhập / Ngày nhập kho, SL nhập > 0, SL đã nhập không trống).`,
    );
    return;
  }

  const coViTri = items.filter((item) => item.location_code).length;
  if (coViTri === 0) {
    console.log(
      "CẢNH BÁO: không đọc được cột Vị trí — phiếu sẽ ghi DEFAULT. Chạy lại Auto_Nami (cột T = Vị trí) rồi upload.",
    );
  } else {
    console.log(`Đã đọc vị trí cho ${coViTri}/${items.length} dòng.`);
  }

  const payload = {
    source_sheet: sheetName,
    items: items.map((item) => namiUpCompactItem(item)),
  };

  const batDau = Date.now();
  console.log(
    `Đang upload ${items.length} dòng nhập kho từ sheet "${sheetName}"...`,
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
      unmatched?: number;
      receipts?: string[];
      rejected?: NamiRejectedRow[];
      errors?: Record<string, string[]>;
    };
    try {
      body = JSON.parse(bodyText);
    } catch {
      body = { message: bodyText };
    }

    if (!response.ok) {
      console.log(`LỖI HTTP ${response.status}: ${body.message ?? bodyText}`);
      return;
    }

    console.log(
      `Upload thành công: ${body.imported ?? items.length} imported, ${body.skipped ?? 0} skipped (${Math.round((Date.now() - batDau) / 1000)}s).`,
    );
    if (body.receipts && body.receipts.length > 0) {
      console.log(`Phiếu nhập: ${body.receipts.join(", ")}`);
    }
    const rejected = body.rejected ?? [];
    const khongKhop = rejected.filter((r) => r.code === NAMI_UP_MA_KHONG_KHOP);
    const tuChoiKhac = rejected.filter((r) => r.code !== NAMI_UP_MA_KHONG_KHOP);

    const daTo = namiUpToVangKhongKhop(
      sheet,
      namiUpVungDuLieu(rows, CFG.headerRow, usedRange.getRowIndex()),
      usedRange.getColumnIndex(),
      khongKhop,
    );

    if (khongKhop.length > 0) {
      console.log(
        `CẢNH BÁO: ${khongKhop.length} dòng KHÔNG khớp SL đơn hàng — đã tô vàng ${daTo} dòng, sửa cột H rồi upload lại:`,
      );
      khongKhop.forEach((r) => {
        console.log(
          `- Dòng ${r.row}: SKU ${r.sku} / HĐ ${r.contract_no} / SL đơn hàng ${r.order_quantity ?? 0} — ${r.reason}`,
        );
      });
    }
    if (tuChoiKhac.length > 0) {
      console.log(
        `CẢNH BÁO: ${tuChoiKhac.length} dòng KHÔNG được nhập kho (không khớp đơn hàng / vượt số lượng):`,
      );
      tuChoiKhac.forEach((r) => {
        console.log(
          `- Dòng ${r.row}: SKU ${r.sku} / HĐ ${r.contract_no} — ${r.reason}`,
        );
      });
    }
    console.log(body.message ?? "Hoàn tất.");
  } catch (err) {
    const msg = String(err);
    if (msg.includes("Failed to fetch")) {
      console.log(
        "LỖI kết nối API: Failed to fetch — thường do (1) apiUrl là localhost/http, (2) server chưa deploy, (3) CORS chưa cấu hình.",
      );
      console.log(`Kiểm tra apiUrl hiện tại: ${CFG.apiUrl}`);
    } else {
      console.log(`LỖI kết nối API: ${msg}`);
    }
  }
}

const namiUpTestables = {
  namiUpChuanHoaNgay,
  namiUpTrichRoundNumber,
  namiUpDocDuLieu,
  namiUpVungDuLieu,
  namiUpLayMaViTri,
  namiUpTimChiSoCot,
  namiUpLayOChuoi,
};

(
  globalThis as typeof globalThis & {
    __UPLOAD_NAMI_TESTABLES__?: typeof namiUpTestables;
  }
).__UPLOAD_NAMI_TESTABLES__ = namiUpTestables;
