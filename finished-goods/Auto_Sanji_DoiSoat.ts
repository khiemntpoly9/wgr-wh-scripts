/**
 * ===================================================================
 * Auto_Sanji_DoiSoat – Đối soát tồn kho hệ thống (API Auto_Sanji_Super)
 * với tồn kho Excel (sheet "Thành phẩm", cùng logic Auto_Sanji).
 * ===================================================================
 *
 * - Khóa so sánh: Mã hợp đồng + SKU.
 * - Tồn Excel = tổng cột nhập − tổng cột xuất (giống Auto_Sanji).
 * - Số lượng đơn hàng: hệ thống (cột "Số lượng đơn hàng" của API) ↔ Excel (cột SL đơn hàng),
 *   không cộng dồn SL đơn hàng khi trùng Mã hợp đồng + SKU trên Excel (lấy giá trị đầu tiên khác 0); chỉ so khi SKU có trên cả hai nguồn.
 * - Vị trí hệ thống: API gửi use_effective_locations — dùng vị trí đề xuất (gợi ý Nami
 *   khi tồn còn ở DEFAULT), không đưa mã DEFAULT vào đối soát.
 * - Vị trí Excel: chuẩn hóa như backend (D-16 = D16, "A11A10" = "A10 A11"), bỏ DEFAULT.
 * - Kết quả ghi ra sheet riêng: có tiêu đề, bảng tổng hợp, filter,
 *   tô màu nền + màu chữ theo trạng thái lệch.
 */
type DoiSoatCellValue = string | number | boolean;

type DoiSoatApiResponse = {
  format?: string;
  sheet?: {
    writes?: { address: string; values: DoiSoatCellValue[][] }[];
  } | null;
  message?: string;
  errors?: Record<string, string[]>;
};

type DoiSoatDong = {
  thiTruong: string;
  hopDong: string;
  maDonHang: string;
  sku: string;
  tenHang: string;
  slDonHang: number;
  ton: number;
  viTriHienThi: string;
  viTri: string[];
};

type DoiSoatTrangThai =
  | "Lệch tồn + vị trí"
  | "Lệch tồn"
  | "Lệch vị trí"
  | "Lệch SL đơn hàng"
  | "Chỉ có trên hệ thống"
  | "Chỉ có trên Excel"
  | "Khớp";

type DoiSoatKetQua = {
  trangThai: DoiSoatTrangThai;
  lechTon: boolean;
  lechViTri: boolean;
  lechSlDon: boolean;
  thiTruong: string;
  hopDong: string;
  maDonHang: string;
  sku: string;
  tenHang: string;
  tonHeThong: number;
  tonExcel: number;
  chenhLech: number;
  slDonHeThong: number;
  slDonExcel: number;
  chenhLechSlDon: number;
  viTriDeXuat: string;
  viTriExcel: string;
  viTriThieuTrenExcel: string;
  viTriThuaTrenExcel: string;
};

type DoiSoatNguonExcel = {
  maHang: DoiSoatCellValue[][];
  tenHang: DoiSoatCellValue[][];
  slDonHang: DoiSoatCellValue[][];
  hopDong: DoiSoatCellValue[][];
  viTri: DoiSoatCellValue[][];
  nhapKho: DoiSoatCellValue[][][];
  xuatKho: DoiSoatCellValue[][][];
};

const DOI_SOAT_CONFIG = {
  apiUrl: "https://app.wgr-wh.site/api/v1/warehouse/auto-sanji-super/preview",
  apiToken: "1|hCwtksIBz4Yc7g2wn6aeMO4L6zri5rStWeSYgH9rf7ad4909", // Bearer token Sanctum có quyền xem tồn kho.
  marketCodes: [] as string[], // [] = tất cả thị trường; áp dụng cho cả API và Excel.
  chiHienThiDongLech: false, // true = bỏ các dòng "Khớp" khỏi bảng chi tiết.
  soSanhViTri: true,
  soSanhSlDonHang: true,
  muiGio: 7, // Giờ hiển thị trên tiêu đề (UTC+7).
  nguon: {
    tenSheet: "Thành phẩm",
    dongBatDau: 2,
    maHang: "C",
    tenHang: "D",
    slDonHang: "F",
    hopDong: "H",
    viTri: "I",
    cotsNhapKho: ["K", "M", "O"],
    cotsXuatKho: ["S", "U", "W"],
  },
  dich: {
    tenSheet: "Đối soát tồn",
  },
};

const DOI_SOAT_THU_TU_TRANG_THAI: DoiSoatTrangThai[] = [
  "Lệch tồn + vị trí",
  "Lệch tồn",
  "Lệch vị trí",
  "Lệch SL đơn hàng",
  "Chỉ có trên hệ thống",
  "Chỉ có trên Excel",
  "Khớp",
];

const DOI_SOAT_MAU: Record<DoiSoatTrangThai, { fill: string; font: string }> = {
  "Lệch tồn + vị trí": { fill: "#FF9999", font: "#800000" },
  "Lệch tồn": { fill: "#FFC7CE", font: "#9C0006" },
  "Lệch vị trí": { fill: "#FFEB9C", font: "#9C5700" },
  "Lệch SL đơn hàng": { fill: "#FCE4D6", font: "#C65911" },
  "Chỉ có trên hệ thống": { fill: "#DDEBF7", font: "#1F4E78" },
  "Chỉ có trên Excel": { fill: "#E4DFEC", font: "#7030A0" },
  Khớp: { fill: "#E2EFDA", font: "#375623" },
};

const DOI_SOAT_HEADER = [
  "STT",
  "Trạng thái",
  "Thị trường",
  "Mã hợp đồng",
  "Mã đơn hàng",
  "SKU",
  "Tên hàng",
  "Tồn hệ thống",
  "Tồn Excel",
  "Chênh lệch (Excel − HT)",
  "SL đơn hệ thống",
  "SL đơn Excel",
  "Chênh lệch SL đơn (Excel − HT)",
  "Vị trí đề xuất",
  "Vị trí Excel",
  "Vị trí thiếu trên Excel",
  "Vị trí thừa trên Excel",
  "Ghi chú",
];

const doiSoatChuoi = (value: DoiSoatCellValue | null | undefined): string =>
  String(value ?? "")
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const doiSoatMaChuan = (value: DoiSoatCellValue | null | undefined): string =>
  doiSoatChuoi(value).replace(/\s+/g, "").toUpperCase();

const doiSoatTaoKhoa = (hopDong: string, sku: string): string =>
  `${doiSoatMaChuan(hopDong)}|${doiSoatMaChuan(sku)}`;

const doiSoatParseSo = (value: DoiSoatCellValue | null | undefined): number => {
  if (typeof value === "number") return isNaN(value) ? 0 : value;
  let s = doiSoatChuoi(value).replace(/\s+/g, "");
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

const doiSoatLayThiTruong = (hopDong: string): string => {
  const match = doiSoatMaChuan(hopDong).match(/^([A-Z]{2})/);
  return match ? match[1] : "KHAC";
};

const doiSoatSapXepViTri = (codes: string[]): string[] => {
  const unique = Array.from(new Set(codes.filter((code) => code !== "")));
  const parts = (code: string) => {
    const match = code.match(/^([A-Z]+)(\d+)$/);
    return match
      ? { zone: match[1], num: Number(match[2]) }
      : { zone: code, num: 0 };
  };
  return unique.sort((a, b) => {
    const pa = parts(a);
    const pb = parts(b);
    if (pa.zone !== pb.zone) return pa.zone < pb.zone ? -1 : 1;
    return pa.num - pb.num;
  });
};

/** Tách mã vị trí giống LocationCodeParser::parse ở backend. */
const doiSoatTachViTri = (raw: DoiSoatCellValue | null | undefined): string[] => {
  const cleaned = String(raw ?? "")
    .toUpperCase()
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!cleaned) return [];
  const parts: string[] = [];
  const regex = /([A-Z]+)\s*0*(\d+)/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(cleaned)) !== null) parts.push(match[1] + match[2]);
  if (parts.length === 0) parts.push(cleaned.replace(/\s+/g, ""));
  return doiSoatSapXepViTri(parts);
};

/** Vị trí API đã là danh sách mã nối bằng dấu cách — tách từng mã trước để không mất mã không có số. */
const doiSoatTachViTriApi = (raw: DoiSoatCellValue | null | undefined): string[] =>
  doiSoatSapXepViTri(
    doiSoatChuoi(raw)
      .split(" ")
      .reduce((all, token) => all.concat(doiSoatTachViTri(token)), [] as string[]),
  );

const doiSoatTaoPayload = (config: typeof DOI_SOAT_CONFIG) => ({
  market_codes: config.marketCodes
    .map((value) => value.trim().toUpperCase())
    .filter((value) => value !== ""),
  use_effective_locations: true,
});

const doiSoatLaDefault = (code: string): boolean => doiSoatMaChuan(code) === "DEFAULT";

const doiSoatLocKhongDefault = (codes: string[]): string[] =>
  doiSoatSapXepViTri(codes.filter((code) => !doiSoatLaDefault(code)));

const doiSoatHienThiViTri = (codes: string[]): string => doiSoatLocKhongDefault(codes).join(" ");

const doiSoatKiemTraUrl = (url: string): string | null => {
  const value = url.trim().toLowerCase();
  if (!value) return "Chưa cấu hình apiUrl.";
  if (!value.startsWith("https://")) return "apiUrl phải dùng HTTPS.";
  if (value.includes("localhost") || value.includes("127.0.0.1"))
    return "apiUrl không được dùng localhost.";
  return null;
};

const doiSoatGopDong = (
  map: Map<string, DoiSoatDong>,
  dong: DoiSoatDong,
  congDonSlDonHang = true,
): void => {
  const khoa = doiSoatTaoKhoa(dong.hopDong, dong.sku);
  const cu = map.get(khoa);
  if (!cu) {
    map.set(khoa, dong);
    return;
  }
  cu.ton += dong.ton;
  if (congDonSlDonHang) cu.slDonHang += dong.slDonHang;
  else if (cu.slDonHang === 0 && dong.slDonHang > 0) cu.slDonHang = dong.slDonHang;
  cu.viTri = doiSoatLocKhongDefault(cu.viTri.concat(dong.viTri));
  cu.viTriHienThi = doiSoatHienThiViTri(cu.viTri);
  if (!cu.tenHang) cu.tenHang = dong.tenHang;
  if (!cu.maDonHang) cu.maDonHang = dong.maDonHang;
};

/** Đọc bảng chi tiết (header "STT … SKU … Vị trí … Tồn") từ kế hoạch sheet Auto_Sanji_Super. */
const doiSoatDocDuLieuApi = (
  body: DoiSoatApiResponse,
): Map<string, DoiSoatDong> | null => {
  const result = new Map<string, DoiSoatDong>();
  if (body.sheet === null) return result;
  const rows = (body.sheet?.writes || []).reduce(
    (all, write) => all.concat(write.values || []),
    [] as DoiSoatCellValue[][],
  );
  const headerIndex = rows.findIndex(
    (row) =>
      doiSoatChuoi(row[0]).toUpperCase() === "STT" &&
      row.some((cell) => doiSoatChuoi(cell) === "SKU"),
  );
  if (headerIndex < 0) return null;
  const header = rows[headerIndex].map((cell) => doiSoatChuoi(cell));
  const col = (name: string) => header.indexOf(name);
  const c = {
    thiTruong: col("Thị trường"),
    hopDong: col("Mã hợp đồng"),
    maDonHang: col("Mã đơn hàng"),
    sku: col("SKU"),
    tenHang: col("Tên hàng"),
    viTri: col("Vị trí"),
    slDonHang: col("Số lượng đơn hàng"),
    ton: col("Tồn"),
  };
  if (c.hopDong < 0 || c.sku < 0 || c.ton < 0) return null;
  const cell = (row: DoiSoatCellValue[], index: number) =>
    index >= 0 ? row[index] : "";

  for (let i = headerIndex + 1; i < rows.length; i++) {
    const row = rows[i];
    const sku = doiSoatChuoi(cell(row, c.sku));
    const hopDong = doiSoatChuoi(cell(row, c.hopDong));
    if (!sku || !hopDong) continue;
    const viTri = doiSoatLocKhongDefault(doiSoatTachViTriApi(doiSoatChuoi(cell(row, c.viTri))));
    doiSoatGopDong(result, {
      thiTruong:
        doiSoatChuoi(cell(row, c.thiTruong)).toUpperCase() ||
        doiSoatLayThiTruong(hopDong),
      hopDong,
      maDonHang: doiSoatChuoi(cell(row, c.maDonHang)),
      sku,
      tenHang: doiSoatChuoi(cell(row, c.tenHang)),
      slDonHang: Math.round(doiSoatParseSo(cell(row, c.slDonHang))),
      ton: Math.round(doiSoatParseSo(cell(row, c.ton))),
      viTriHienThi: doiSoatHienThiViTri(viTri),
      viTri,
    });
  }
  return result;
};

const doiSoatDocDuLieuExcel = (
  nguon: DoiSoatNguonExcel,
  marketCodes: string[],
): Map<string, DoiSoatDong> => {
  const result = new Map<string, DoiSoatDong>();
  const locThiTruong = marketCodes
    .map((value) => value.trim().toUpperCase())
    .filter((value) => value !== "");
  const giaTri = (cot: DoiSoatCellValue[][], i: number) => cot[i]?.[0] ?? "";

  for (let i = 0; i < nguon.maHang.length; i++) {
    const sku = doiSoatChuoi(giaTri(nguon.maHang, i));
    const hopDong = doiSoatChuoi(giaTri(nguon.hopDong, i));
    if (!sku || !hopDong) continue;
    const thiTruong = doiSoatLayThiTruong(hopDong);
    if (locThiTruong.length > 0 && !locThiTruong.includes(thiTruong)) continue;
    const tongNhap = nguon.nhapKho.reduce((sum, cot) => sum + doiSoatParseSo(giaTri(cot, i)), 0);
    const tongXuat = nguon.xuatKho.reduce((sum, cot) => sum + doiSoatParseSo(giaTri(cot, i)), 0);
    const viTri = doiSoatLocKhongDefault(doiSoatTachViTri(doiSoatChuoi(giaTri(nguon.viTri, i))));
    doiSoatGopDong(
      result,
      {
        thiTruong,
        hopDong,
        maDonHang: "",
        sku,
        tenHang: doiSoatChuoi(giaTri(nguon.tenHang, i)),
        slDonHang: Math.round(doiSoatParseSo(giaTri(nguon.slDonHang, i))),
        ton: Math.round(tongNhap - tongXuat),
        viTriHienThi: doiSoatHienThiViTri(viTri),
        viTri,
      },
      false,
    );
  }
  return result;
};

const doiSoatSoSanh = (
  heThong: Map<string, DoiSoatDong>,
  excel: Map<string, DoiSoatDong>,
  soSanhViTri: boolean,
  soSanhSlDonHang = true,
): DoiSoatKetQua[] => {
  const khoa = Array.from(new Set(Array.from(heThong.keys()).concat(Array.from(excel.keys()))));
  const ketQua: DoiSoatKetQua[] = [];

  for (const k of khoa) {
    const ht = heThong.get(k);
    const ex = excel.get(k);
    const tonHeThong = ht ? ht.ton : 0;
    const tonExcel = ex ? ex.ton : 0;
    // Hệ thống chỉ trả dòng tồn > 0; dòng Excel tồn 0 mà hệ thống không có là khớp.
    if (tonHeThong === 0 && tonExcel === 0) continue;
    const viTriHt = ht ? ht.viTri : [];
    const viTriEx = ex ? ex.viTri : [];
    const thieu = viTriHt.filter((code) => !viTriEx.includes(code));
    const thua = viTriEx.filter((code) => !viTriHt.includes(code));
    const lechTon = tonHeThong !== tonExcel;
    const lechViTri =
      soSanhViTri && tonHeThong > 0 && tonExcel > 0 && (thieu.length > 0 || thua.length > 0);
    const slDonHeThong = ht ? ht.slDonHang : 0;
    const slDonExcel = ex ? ex.slDonHang : 0;
    const lechSlDon =
      soSanhSlDonHang && !!ht && !!ex && tonExcel !== 0 && slDonHeThong !== slDonExcel;

    let trangThai: DoiSoatTrangThai = "Khớp";
    if (!ex || tonExcel === 0) trangThai = "Chỉ có trên hệ thống";
    else if (!ht) trangThai = "Chỉ có trên Excel";
    else if (lechTon && lechViTri) trangThai = "Lệch tồn + vị trí";
    else if (lechTon) trangThai = "Lệch tồn";
    else if (lechViTri) trangThai = "Lệch vị trí";
    else if (lechSlDon) trangThai = "Lệch SL đơn hàng";

    const goc = (ht || ex) as DoiSoatDong;
    ketQua.push({
      trangThai,
      lechTon,
      lechViTri,
      lechSlDon,
      thiTruong: goc.thiTruong,
      hopDong: goc.hopDong,
      maDonHang: (ht && ht.maDonHang) || "",
      sku: goc.sku,
      tenHang: (ht && ht.tenHang) || (ex && ex.tenHang) || "",
      tonHeThong,
      tonExcel,
      chenhLech: tonExcel - tonHeThong,
      slDonHeThong,
      slDonExcel,
      chenhLechSlDon: slDonExcel - slDonHeThong,
      viTriDeXuat: ht ? ht.viTriHienThi : "",
      viTriExcel: ex ? ex.viTriHienThi : "",
      viTriThieuTrenExcel: soSanhViTri ? thieu.join(" ") : "",
      viTriThuaTrenExcel: soSanhViTri ? thua.join(" ") : "",
    });
  }

  const soSanhTuNhien = (a: string, b: string) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
  return ketQua.sort(
    (a, b) =>
      DOI_SOAT_THU_TU_TRANG_THAI.indexOf(a.trangThai) -
        DOI_SOAT_THU_TU_TRANG_THAI.indexOf(b.trangThai) ||
      Number(b.lechSlDon) - Number(a.lechSlDon) ||
      soSanhTuNhien(a.thiTruong, b.thiTruong) ||
      soSanhTuNhien(a.hopDong, b.hopDong) ||
      soSanhTuNhien(a.sku, b.sku),
  );
};

const doiSoatDemTrangThai = (ketQua: DoiSoatKetQua[]): Record<DoiSoatTrangThai, number> => {
  const dem = {} as Record<DoiSoatTrangThai, number>;
  DOI_SOAT_THU_TU_TRANG_THAI.forEach((trangThai) => (dem[trangThai] = 0));
  ketQua.forEach((dong) => dem[dong.trangThai]++);
  return dem;
};

/** Gom các dòng liên tiếp cùng trạng thái + cờ lệch SL đơn (đã sort) để format theo khối, tránh format từng dòng. */
const doiSoatGomKhoi = (
  ketQua: DoiSoatKetQua[],
): { trangThai: DoiSoatTrangThai; lechSlDon: boolean; batDau: number; soDong: number }[] => {
  const khoi: { trangThai: DoiSoatTrangThai; lechSlDon: boolean; batDau: number; soDong: number }[] = [];
  ketQua.forEach((dong, index) => {
    const cuoi = khoi[khoi.length - 1];
    if (cuoi && cuoi.trangThai === dong.trangThai && cuoi.lechSlDon === dong.lechSlDon)
      cuoi.soDong++;
    else
      khoi.push({ trangThai: dong.trangThai, lechSlDon: dong.lechSlDon, batDau: index, soDong: 1 });
  });
  return khoi;
};

const doiSoatDongChiTiet = (ketQua: DoiSoatKetQua[]): DoiSoatCellValue[][] =>
  ketQua.map((dong, index) => [
    index + 1,
    dong.trangThai,
    dong.thiTruong,
    dong.hopDong,
    dong.maDonHang,
    dong.sku,
    dong.tenHang,
    dong.tonHeThong,
    dong.tonExcel,
    dong.chenhLech,
    dong.slDonHeThong,
    dong.slDonExcel,
    dong.chenhLechSlDon,
    dong.viTriDeXuat,
    dong.viTriExcel,
    dong.viTriThieuTrenExcel,
    dong.viTriThuaTrenExcel,
    dong.lechSlDon && dong.trangThai !== "Lệch SL đơn hàng" ? "Lệch SL đơn hàng" : "",
  ]);

const doiSoatThoiGian = (muiGio: number): string => {
  const d = new Date(Date.now() + muiGio * 3600 * 1000);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${pad(d.getUTCDate())}/${pad(d.getUTCMonth() + 1)}/${d.getUTCFullYear()} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`;
};

const doiSoatCot = (n: number): string => {
  let out = "";
  let x = n;
  while (x > 0) {
    const mod = (x - 1) % 26;
    out = String.fromCharCode(65 + mod) + out;
    x = Math.floor((x - 1) / 26);
  }
  return out;
};

const doiSoatKeVien = (range: ExcelScript.Range): void => {
  const format = range.getFormat();
  [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal,
    ExcelScript.BorderIndex.insideVertical,
  ].forEach((index) =>
    format.getRangeBorder(index).setStyle(ExcelScript.BorderLineStyle.continuous),
  );
};

const doiSoatToMau = (
  range: ExcelScript.Range,
  trangThai: DoiSoatTrangThai,
  bold = false,
): void => {
  const format = range.getFormat();
  format.getFill().setColor(DOI_SOAT_MAU[trangThai].fill);
  format.getFont().setColor(DOI_SOAT_MAU[trangThai].font);
  if (bold) format.getFont().setBold(true);
};

declare function fetch(
  input: string,
  init?: { method?: string; headers?: Record<string, string>; body?: string },
): Promise<{ ok: boolean; status: number; text(): Promise<string> }>;

async function main(workbook: ExcelScript.Workbook) {
  const CFG = DOI_SOAT_CONFIG;
  const NGUON = CFG.nguon;
  const urlError = doiSoatKiemTraUrl(CFG.apiUrl);
  if (urlError) return console.log(`LỖI cấu hình: ${urlError}`);
  if (!CFG.apiToken.trim()) return console.log("LỖI: Chưa cấu hình apiToken.");
  const startedAt = Date.now();

  const sourceSheet = workbook.getWorksheet(NGUON.tenSheet);
  if (!sourceSheet)
    return console.log(`LỖI: Không tìm thấy sheet nguồn "${NGUON.tenSheet}".`);
  const lastRow =
    sourceSheet
      .getRange(`${NGUON.maHang}1048576`)
      .getRangeEdge(ExcelScript.KeyboardDirection.up)
      .getRowIndex() + 1;
  const docCot = (col: string): DoiSoatCellValue[][] =>
    lastRow < NGUON.dongBatDau
      ? []
      : (sourceSheet
          .getRange(`${col}${NGUON.dongBatDau}:${col}${lastRow}`)
          .getValues() as DoiSoatCellValue[][]);
  const excel = doiSoatDocDuLieuExcel(
    {
      maHang: docCot(NGUON.maHang),
      tenHang: docCot(NGUON.tenHang),
      slDonHang: docCot(NGUON.slDonHang),
      hopDong: docCot(NGUON.hopDong),
      viTri: docCot(NGUON.viTri),
      nhapKho: NGUON.cotsNhapKho.map((col) => docCot(col)),
      xuatKho: NGUON.cotsXuatKho.map((col) => docCot(col)),
    },
    CFG.marketCodes,
  );

  let heThong: Map<string, DoiSoatDong> | null;
  try {
    const response = await fetch(CFG.apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Authorization: `Bearer ${CFG.apiToken}`,
      },
      body: JSON.stringify(doiSoatTaoPayload(CFG)),
    });
    const text = await response.text();
    let body: DoiSoatApiResponse;
    try {
      body = JSON.parse(text) as DoiSoatApiResponse;
    } catch {
      body = { message: text };
    }
    if (!response.ok) {
      const errors = body.errors;
      const validation = errors
        ? Object.keys(errors)
            .map((field) => `${field}: ${errors[field].join(", ")}`)
            .join(" | ")
        : "";
      return console.log(
        `LỖI HTTP ${response.status}: ${body.message || text}${validation ? ` — ${validation}` : ""}`,
      );
    }
    if (body.format !== "auto_sanji_backend_v1")
      return console.log("LỖI: Backend trả về sai format Auto_Sanji.");
    heThong = doiSoatDocDuLieuApi(body);
    if (!heThong)
      return console.log("LỖI: Không tìm thấy bảng chi tiết tồn kho trong dữ liệu API.");
  } catch (error) {
    const message = String(error);
    return console.log(
      message.includes("Failed to fetch")
        ? `LỖI kết nối API — kiểm tra deploy, HTTPS và CORS: ${CFG.apiUrl}`
        : `LỖI gọi API: ${message}`,
    );
  }

  const tatCa = doiSoatSoSanh(heThong, excel, CFG.soSanhViTri, CFG.soSanhSlDonHang);
  const dem = doiSoatDemTrangThai(tatCa);
  const ketQua = CFG.chiHienThiDongLech
    ? tatCa.filter((dong) => dong.trangThai !== "Khớp")
    : tatCa;
  const tongTonHt = tatCa.reduce((sum, dong) => sum + dong.tonHeThong, 0);
  const tongTonEx = tatCa.reduce((sum, dong) => sum + dong.tonExcel, 0);
  const soLech = tatCa.length - dem["Khớp"];
  const soLechSlDon = tatCa.filter((dong) => dong.lechSlDon).length;

  const soCot = DOI_SOAT_HEADER.length;
  const cotCuoi = doiSoatCot(soCot);
  const dongTieuDeTongHop = 3;
  const dongSoLieuTongHop = 4;
  const dongHeader = 6;
  const dongDauChiTiet = dongHeader + 1;
  const dongCuoi = dongHeader + Math.max(ketQua.length, 1);
  const pad = (row: DoiSoatCellValue[]) => {
    const out = row.slice();
    while (out.length < soCot) out.push("");
    return out;
  };
  const thiTruongText = CFG.marketCodes.length
    ? CFG.marketCodes.map((x) => x.trim().toUpperCase()).join(", ")
    : "Tất cả";
  const values: DoiSoatCellValue[][] = [
    pad([`ĐỐI SOÁT TỒN KHO: HỆ THỐNG (API) ↔ EXCEL "${NGUON.tenSheet}"`]),
    pad([
      `Thời gian: ${doiSoatThoiGian(CFG.muiGio)}  |  Thị trường: ${thiTruongText}  |  Tổng tồn hệ thống: ${tongTonHt}  |  Tổng tồn Excel: ${tongTonEx}  |  Số SKU lệch: ${soLech}/${tatCa.length}  |  Số SKU lệch SL đơn: ${soLechSlDon}`,
    ]),
    pad(["Trạng thái", ...DOI_SOAT_THU_TU_TRANG_THAI, "Tổng"]),
    pad(["Số SKU", ...DOI_SOAT_THU_TU_TRANG_THAI.map((t) => dem[t]), tatCa.length]),
    pad([]),
    DOI_SOAT_HEADER.slice(),
    ...(ketQua.length
      ? doiSoatDongChiTiet(ketQua)
      : [pad(["", "Không có SKU lệch."])]),
  ];

  let sheet = workbook.getWorksheet(CFG.dich.tenSheet);
  if (!sheet) sheet = workbook.addWorksheet(CFG.dich.tenSheet);
  const used = sheet.getUsedRange();
  if (used) {
    used.unmerge();
    used.clear(ExcelScript.ClearApplyTo.all);
  }
  const filter = sheet.getAutoFilter();
  if (filter.getEnabled()) filter.remove();
  sheet.getFreezePanes().unfreeze();

  const full = sheet.getRange(`A1:${cotCuoi}${dongCuoi}`);
  full.setValues(values);
  full.getFormat().getFont().setName("Times New Roman");
  full.getFormat().getFont().setSize(12);
  full.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  full.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  const tieuDe = sheet.getRange(`A1:${cotCuoi}1`);
  tieuDe.merge(false);
  tieuDe.getFormat().getFont().setBold(true);
  tieuDe.getFormat().getFont().setSize(16);
  tieuDe.getFormat().getFont().setColor("#1F4E78");
  tieuDe.getFormat().setRowHeight(28);
  const phuDe = sheet.getRange(`A2:${cotCuoi}2`);
  phuDe.merge(false);
  phuDe.getFormat().getFont().setItalic(true);

  const cotTongHopCuoi = doiSoatCot(DOI_SOAT_THU_TU_TRANG_THAI.length + 2);
  const tongHop = sheet.getRange(`A${dongTieuDeTongHop}:${cotTongHopCuoi}${dongSoLieuTongHop}`);
  doiSoatKeVien(tongHop);
  tongHop.getFormat().getFont().setBold(true);
  sheet
    .getRange(`A${dongTieuDeTongHop}:A${dongSoLieuTongHop}`)
    .getFormat()
    .getFill()
    .setColor("#D9E1F2");
  sheet
    .getRange(`${cotTongHopCuoi}${dongTieuDeTongHop}:${cotTongHopCuoi}${dongSoLieuTongHop}`)
    .getFormat()
    .getFill()
    .setColor("#D9E1F2");
  DOI_SOAT_THU_TU_TRANG_THAI.forEach((trangThai, index) => {
    const cot = doiSoatCot(index + 2);
    doiSoatToMau(sheet.getRange(`${cot}${dongTieuDeTongHop}:${cot}${dongSoLieuTongHop}`), trangThai);
  });
  tongHop.getFormat().setWrapText(true);

  const header = sheet.getRange(`A${dongHeader}:${cotCuoi}${dongHeader}`);
  header.getFormat().getFont().setBold(true);
  header.getFormat().getFont().setColor("#FFFFFF");
  header.getFormat().getFill().setColor("#305496");
  header.getFormat().setWrapText(true);
  header.getFormat().setRowHeight(36);

  if (ketQua.length > 0) {
    for (const khoi of doiSoatGomKhoi(ketQua)) {
      const dau = dongDauChiTiet + khoi.batDau;
      const cuoi = dau + khoi.soDong - 1;
      doiSoatToMau(sheet.getRange(`A${dau}:${cotCuoi}${cuoi}`), khoi.trangThai);
      if (khoi.trangThai === "Khớp") continue;
      sheet.getRange(`B${dau}:B${cuoi}`).getFormat().getFont().setBold(true);
      sheet.getRange(`F${dau}:F${cuoi}`).getFormat().getFont().setBold(true);
      const lechTon = ketQua[khoi.batDau].lechTon;
      const lechViTri = ketQua[khoi.batDau].lechViTri;
      if (lechTon) sheet.getRange(`H${dau}:J${cuoi}`).getFormat().getFont().setBold(true);
      if (khoi.lechSlDon) {
        const slDon = sheet.getRange(`K${dau}:M${cuoi}`).getFormat();
        slDon.getFont().setBold(true);
        slDon.getFont().setColor(DOI_SOAT_MAU["Lệch SL đơn hàng"].font);
        sheet.getRange(`R${dau}:R${cuoi}`).getFormat().getFont().setBold(true);
      }
      if (lechViTri) sheet.getRange(`N${dau}:Q${cuoi}`).getFormat().getFont().setBold(true);
    }
    sheet.getRange(`D${dongDauChiTiet}:G${dongCuoi}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    sheet.getRange(`N${dongDauChiTiet}:R${dongCuoi}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    sheet.getRange(`H${dongDauChiTiet}:I${dongCuoi}`).setNumberFormat("#,##0");
    sheet.getRange(`K${dongDauChiTiet}:L${dongCuoi}`).setNumberFormat("#,##0");
    sheet.getRange(`J${dongDauChiTiet}:J${dongCuoi}`).setNumberFormat("+#,##0;-#,##0;0");
    sheet.getRange(`M${dongDauChiTiet}:M${dongCuoi}`).setNumberFormat("+#,##0;-#,##0;0");
  }

  // Viền trong trên vài nghìn dòng làm Excel Online treo lâu.
  doiSoatKeVien(
    ketQua.length <= 3000
      ? sheet.getRange(`A${dongHeader}:${cotCuoi}${dongCuoi}`)
      : header,
  );
  filter.apply(`A${dongHeader}:${cotCuoi}${dongCuoi}`);
  sheet.getFreezePanes().freezeRows(dongHeader);
  sheet.getRange(`A${dongHeader}:F${dongCuoi}`).getFormat().autofitColumns();
  sheet.getRange(`H${dongHeader}:M${dongCuoi}`).getFormat().autofitColumns();
  sheet.getRange("G:G").getFormat().setColumnWidth(260);
  sheet.getRange("N:Q").getFormat().setColumnWidth(130);
  sheet.getRange("R:R").getFormat().setColumnWidth(160);
  sheet.getRange(`B${dongTieuDeTongHop}:${cotTongHopCuoi}${dongTieuDeTongHop}`).getFormat().setRowHeight(36);
  sheet.activate();

  console.log(
    `Auto_Sanji_DoiSoat hoàn tất: ${tatCa.length} SKU, lệch ${soLech} (tồn+vị trí ${dem["Lệch tồn + vị trí"]}, tồn ${dem["Lệch tồn"]}, vị trí ${dem["Lệch vị trí"]}, SL đơn ${dem["Lệch SL đơn hàng"]} (tổng SKU lệch SL đơn ${soLechSlDon}), chỉ HT ${dem["Chỉ có trên hệ thống"]}, chỉ Excel ${dem["Chỉ có trên Excel"]}). Thời gian ${Math.round((Date.now() - startedAt) / 1000)}s.`,
  );
}

const doiSoatTestables = {
  doiSoatTachViTri,
  doiSoatTachViTriApi,
  doiSoatLaDefault,
  doiSoatLocKhongDefault,
  doiSoatHienThiViTri,
  doiSoatTaoPayload,
  doiSoatDocDuLieuApi,
  doiSoatDocDuLieuExcel,
  doiSoatSoSanh,
  doiSoatDemTrangThai,
  doiSoatGomKhoi,
  doiSoatDongChiTiet,
};
(
  globalThis as typeof globalThis & {
    __AUTO_SANJI_DOI_SOAT_TESTABLES__?: typeof doiSoatTestables;
  }
).__AUTO_SANJI_DOI_SOAT_TESTABLES__ = doiSoatTestables;
