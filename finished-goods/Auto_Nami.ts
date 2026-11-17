/**
 * ===================================================================
 *  Auto_Nami – Tự động xuất dữ liệu nhập kho theo ngày, tuần, tháng, năm.
 * ===================================================================
 *
 * Script đọc sheet nguồn "Thành phẩm" (không cần chọn sheet trước khi chạy),
 * xuất kết quả nhập kho theo ngày/tuần/tháng/năm. Cấu hình trong NAMI_CONFIG.
 *
 * Chỉ cần khi muốn xem sheet KQ_NhapKho_* trước; tải lên dùng Upload_Nami_Super.
 *
 */
type NamiFilterMode = "day" | "week" | "month" | "year";

interface NamiProductInfo {
  trongLuong: number;
  dai: number;
  rong: number;
  cao: number;
  theTich: number;
}

type NamiThongSoNguon = "ma" | "ten" | "ten-mo";

type AutoNamiThongSoTraCuuResult = {
  thongSo: NamiProductInfo;
  nguon: NamiThongSoNguon;
};

type NamiThongSoTheoTenEntry = {
  tenChuan: string;
  tokens: string[];
  thongSo: NamiProductInfo;
  coDuLieu: boolean;
};

type NamiTraCuuTenMoConfig = {
  bat: boolean;
  diemToiThieu: number;
  tyLeTokenKhopToiThieu: number;
  soTokenKhopToiDa: number;
};

type NamiThongSoConfig = {
  bat: boolean;
};

type NamiCotBangMaPhanLoai = "maPhanLoai" | "soLuong" | "trongLuong" | "theTich";

type NamiCotBangThiTruong =
  | "thiTruong"
  | "maPhanLoai"
  | "soLuong"
  | "trongLuong"
  | "theTich"
  | "theTichChia";

type NamiThiTruongMaPhanLoai = {
  thiTruong: string;
  maPhanLoai: string;
};

type NamiCotThongKe<T> = {
  tieuDe: string;
  dinhDangSo?: string;
  giaTri: (item: T, soDong: number) => string | number;
};

type NamiThongKeKetQuaConfig = {
  bat: boolean;
  cotBatDau: string;
  dongBatDau: number;
  heSoChiaTheTich: number;
  thuTuCotBangMaPhanLoai: NamiCotBangMaPhanLoai[];
  thuTuCotBangThiTruong: NamiCotBangThiTruong[];
};

const NAMI_COT_BANG_MA_PHAN_LOAI_MAC_DINH: NamiCotBangMaPhanLoai[] = [
  "maPhanLoai",
  "soLuong",
  "trongLuong",
  "theTich",
];

const NAMI_COT_BANG_THI_TRUONG_MAC_DINH: NamiCotBangThiTruong[] = [
  "thiTruong",
  "maPhanLoai",
  "soLuong",
  "trongLuong",
  "theTich",
  "theTichChia",
];

type NamiDateCellValue = string | number | boolean | null | undefined;
const namiSerialDateCache = new Map<number, string | null>();
const namiTextDateCache = new Map<string, string | null>();

const NAMI_CONFIG = {
  // =======================================================
  // CẤU HÌNH NGÀY CẦN LẤY DỮ LIỆU
  // Sửa ngày ở biến bên dưới (Định dạng: dd/mm/yyyy)
  // =======================================================
  targetDate: "25/05/2026",

  // Chế độ lọc:
  // - "day": theo ngày
  // - "week": theo tuần chứa TARGET_DATE (từ thứ 2 đến chủ nhật)
  // - "month": theo tháng
  // - "year": theo năm
  filterMode: "day" as NamiFilterMode,

  tenSheetNguon: "Thành phẩm",
  cotNguon: {
    viTri: "I",
  },

  output: {
    sheetNamePrefix: "KQ_NhapKho_",
    lastColLetter: "T",
    headerRow: 2,
    batAutoFilter: true,
    freezeDenDongTieuDe: true,
  },

  thongKeKetQua: {
    bat: true,
    cotBatDau: "V", // Bảng thống kê từ cột V; cột U trống = khoảng cách sau Vị trí (T).
    dongBatDau: 3, // Tiêu đề bảng nằm dưới 2 dòng đầu (vùng freeze panes).
    heSoChiaTheTich: 68,
    // Thứ tự cột trong từng bảng: sắp xếp lại hoặc bỏ bớt tên cột (để trống = mặc định).
    // Bảng 1: "maPhanLoai" | "soLuong" | "trongLuong" | "theTich"
    thuTuCotBangMaPhanLoai: ["maPhanLoai", "soLuong", "trongLuong", "theTich"],
    // Bảng 2: thêm "thiTruong" và "theTichChia" (Tổng thể tích / heSoChiaTheTich)
    thuTuCotBangThiTruong: [
      "thiTruong",
      "maPhanLoai",
      "theTich",
      "trongLuong",
      "soLuong",
      "theTichChia",
    ],
  } as NamiThongKeKetQuaConfig,

  // Ký tự phân loại từ mã hợp đồng (giống Auto_Sanji): JP JPT4001-2 => "T", AUF => "F", …
  thongKeKyTuMacDinh: ["F", "A", "S", "C", "T", "M"],
  thongKeKyTuTheoThiTruong: {
    JP: ["F", "A", "S", "C", "T", "M"],
    AU: ["F", "A", "S", "C", "T", "M"],
    EU: ["F", "A", "S", "C", "T", "M"],
  } as { [thiTruong: string]: string[] },

  tenSheetThongSo: "2026",
  thongSo: {
    // Mặc định bật theo nhu cầu vận hành: luôn điền đủ cột thông số M..R.
    bat: true,
  } as NamiThongSoConfig,
  traCuuTenMo: {
    bat: true, // Tra cứu tên mô để chuẩn hóa thông số.
    diemToiThieu: 0.55, // Cấu hình này để
    tyLeTokenKhopToiThieu: 0.5, //
    soTokenKhopToiDa: 3,
  } as NamiTraCuuTenMoConfig,

  format: {
    fontName: "Times New Roman",
    fontSize: 14,
    subtotalRange: "I1:J1",
    subtotalFillColor: "#FAB0B5",
    headerFillColor: "#4F81BD",
    tenLookupWarningFillColor: "#FFF2CC",
    dateColumn: "B",
    leftAlignColumns: ["C", "D", "E", "F", "L", "T"],
  },
};

// Hàng cuối sheet thông số — giống Auto_Mitsuha: chỉ nhìn cột B, không dùng UsedRange cả sheet.
const namiTimDongCuoiSheetThongSo = (sheet: ExcelScript.Worksheet): number =>
  sheet
    .getRange("B1048576")
    .getRangeEdge(ExcelScript.KeyboardDirection.up)
    .getRowIndex() + 1;

const namiTaoNgayKey = (
  year: number,
  month: number,
  day: number,
): string | null => {
  if (
    !Number.isFinite(year) ||
    !Number.isFinite(month) ||
    !Number.isFinite(day)
  )
    return null;
  const dt = new Date(Date.UTC(year, month - 1, day));
  if (
    dt.getUTCFullYear() !== year ||
    dt.getUTCMonth() + 1 !== month ||
    dt.getUTCDate() !== day
  ) {
    return null;
  }
  return `${year.toString().padStart(4, "0")}-${month
    .toString()
    .padStart(2, "0")}-${day.toString().padStart(2, "0")}`;
};

const namiExcelSerialSangNgayKey = (serial: number): string | null => {
  if (namiSerialDateCache.has(serial)) {
    return namiSerialDateCache.get(serial) ?? null;
  }
  if (!Number.isFinite(serial) || serial <= 0) return null;
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  const result = namiTaoNgayKey(
    dateInfo.getUTCFullYear(),
    dateInfo.getUTCMonth() + 1,
    dateInfo.getUTCDate(),
  );
  namiSerialDateCache.set(serial, result);
  return result;
};

const namiNamTuChuoi = (raw: string): number => {
  const y = Number(raw);
  if (!Number.isFinite(y)) return NaN;
  if (raw.length === 2) return 2000 + y;
  return y;
};

/** Chuẩn hóa vị trí kho (giống Auto_Zoro / Auto_Mitsuha): "A05  B44" → "A5 B44". */
const namiChuanHoaViTri = (giaTri: string | number | boolean): string => {
  let raw = String(giaTri ?? "").toUpperCase();
  raw = raw
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const cum: { khu: string; so: number }[] = [];
  const regex = /([A-Z]+)\s*0*(\d+)/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(raw)) !== null) {
    cum.push({
      khu: match[1],
      so: Number(match[2]),
    });
  }

  if (cum.length > 0) {
    return cum.map((x) => `${x.khu}${x.so}`).join(" ");
  }
  return raw;
};

const namiCotChuSangIndex = (cot: string): number => {
  let result = 0;
  const normalized = cot.trim().toUpperCase();
  for (let i = 0; i < normalized.length; i++) {
    result = result * 26 + (normalized.charCodeAt(i) - 64);
  }
  return result - 1;
};

const namiLayGiaTriO = (
  texts: (string | number | boolean)[],
  values: (string | number | boolean)[],
  idx: number,
): string => {
  if (idx < 0) return "";
  const fromText = String(texts[idx] ?? "").trim();
  if (fromText) return fromText;
  const fromValue = values[idx];
  if (fromValue === null || fromValue === undefined || fromValue === "") {
    return "";
  }
  return String(fromValue);
};

const namiTextSangNgayKey = (raw: NamiDateCellValue): string | null => {
  if (typeof raw !== "string") return null;
  const text = raw.trim();
  if (!text) return null;
  if (namiTextDateCache.has(text)) {
    return namiTextDateCache.get(text) ?? null;
  }

  const matchedDate = text.match(/(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})/);
  if (matchedDate) {
    const p1 = matchedDate[1];
    const p2 = matchedDate[2];
    const p3 = matchedDate[3];

    if (p1.length === 4) {
      const yyyy = Number(p1);
      const mm = Number(p2);
      const dd = Number(p3);
      const isoKey = namiTaoNgayKey(yyyy, mm, dd);
      if (isoKey) {
        namiTextDateCache.set(text, isoKey);
        return isoKey;
      }
    }

    const yyyy = namiNamTuChuoi(p3);
    if (Number.isFinite(yyyy)) {
      const n1 = Number(p1);
      const n2 = Number(p2);

      // Ưu tiên dd/mm/yyyy; tự động đảo lại khi thấy tháng rõ ràng (>12).
      const dayFirst = namiTaoNgayKey(yyyy, n2, n1);
      const monthFirst = namiTaoNgayKey(yyyy, n1, n2);

      if (n1 > 12 && n2 <= 12) {
        namiTextDateCache.set(text, dayFirst);
        return dayFirst;
      }
      if (n2 > 12 && n1 <= 12) {
        namiTextDateCache.set(text, monthFirst);
        return monthFirst;
      }
      const parsedByParts = dayFirst || monthFirst;
      namiTextDateCache.set(text, parsedByParts);
      return parsedByParts;
    }
  }

  const parsed = new Date(text);
  if (isNaN(parsed.getTime())) {
    namiTextDateCache.set(text, null);
    return null;
  }
  const result = namiTaoNgayKey(
    parsed.getUTCFullYear(),
    parsed.getUTCMonth() + 1,
    parsed.getUTCDate(),
  );
  namiTextDateCache.set(text, result);
  return result;
};

const namiLayNgayKeyTuO = (
  textDate: NamiDateCellValue,
  valueDate: NamiDateCellValue,
): string | null => {
  if (typeof valueDate === "number") {
    const fromSerial = namiExcelSerialSangNgayKey(valueDate);
    if (fromSerial) return fromSerial;
  }
  const fromValueText = namiTextSangNgayKey(valueDate);
  if (fromValueText) return fromValueText;
  return namiTextSangNgayKey(textDate);
};

/** Giữ cột hợp lệ theo thứ tự cấu hình, bỏ trùng; rỗng thì dùng thứ tự mặc định. */
const namiChuanHoaThuTuCot = <T extends string>(
  cauHinh: string[] | undefined,
  macDinh: T[],
): T[] => {
  const ketQua: T[] = [];
  for (const cot of cauHinh ?? []) {
    const hopLe = macDinh.find((x) => x === String(cot ?? "").trim());
    if (hopLe && !ketQua.includes(hopLe)) ketQua.push(hopLe);
  }
  return ketQua.length > 0 ? ketQua : [...macDinh];
};

const namiTestables = {
  namiChuanHoaThuTuCot,
  namiTaoNgayKey,
  namiExcelSerialSangNgayKey,
  namiTextSangNgayKey,
  namiLayNgayKeyTuO,
  namiChuanHoaViTri,
  namiCotChuSangIndex,
  namiLayGiaTriO,
};

(
  globalThis as typeof globalThis & {
    __AUTO_NAMI_TESTABLES__?: typeof namiTestables;
  }
).__AUTO_NAMI_TESTABLES__ = namiTestables;

// @ts-ignore: Office Scripts
async function main(workbook: ExcelScript.Workbook) {
  const CONFIG = NAMI_CONFIG;
  namiSerialDateCache.clear();
  namiTextDateCache.clear();

  // Phân tích TARGET_DATE
  const targetParts = CONFIG.targetDate.split(/[/\-.]/);
  const targetD = targetParts[0].padStart(2, "0");
  const targetM = targetParts[1].padStart(2, "0");
  const targetY =
    targetParts[2].length === 2 ? "20" + targetParts[2] : targetParts[2];

  const taoNgayUTC = (d: string, m: string, y: string): Date =>
    new Date(Date.UTC(Number(y), Number(m) - 1, Number(d)));
  const dinhDangNgay = (date: Date): string => {
    const d = date.getUTCDate().toString().padStart(2, "0");
    const m = (date.getUTCMonth() + 1).toString().padStart(2, "0");
    const y = date.getUTCFullYear().toString();
    return `${d}/${m}/${y}`;
  };
  const dinhDangNgayChoTenSheet = (date: Date): string =>
    dinhDangNgay(date).replace(/\//g, "-");
  const dinhDangNgayNgan = (date: Date): string => {
    const d = date.getUTCDate().toString().padStart(2, "0");
    const m = (date.getUTCMonth() + 1).toString().padStart(2, "0");
    return `${d}${m}`;
  };
  const dinhDangNgayKey = (date: Date): string => {
    const d = date.getUTCDate().toString().padStart(2, "0");
    const m = (date.getUTCMonth() + 1).toString().padStart(2, "0");
    const y = date.getUTCFullYear().toString();
    return `${y}-${m}-${d}`;
  };
  const tinhKhoangTuan = (date: Date): { start: Date; end: Date } => {
    const thu = date.getUTCDay();
    const lechVeThuHai = thu === 0 ? -6 : 1 - thu;
    const start = new Date(date.getTime());
    start.setUTCDate(date.getUTCDate() + lechVeThuHai);
    const end = new Date(start.getTime());
    end.setUTCDate(start.getUTCDate() + 6);
    return { start, end };
  };

  const targetDateUtc = taoNgayUTC(targetD, targetM, targetY);
  const targetDateKey = dinhDangNgayKey(targetDateUtc);
  const weekRange = tinhKhoangTuan(targetDateUtc);
  const weekStartKey = dinhDangNgayKey(weekRange.start);
  const weekEndKey = dinhDangNgayKey(weekRange.end);
  const filterLabel =
    CONFIG.filterMode === "day"
      ? `ngày ${CONFIG.targetDate}`
      : CONFIG.filterMode === "week"
        ? `tuần ${dinhDangNgay(weekRange.start)} - ${dinhDangNgay(weekRange.end)}`
        : CONFIG.filterMode === "month"
          ? `tháng ${targetM}/${targetY}`
          : `năm ${targetY}`;

  // Tên sheet kết quả
  let suffix = CONFIG.targetDate.replace(/\//g, "-");
  if (CONFIG.filterMode === "week") {
    suffix = `Tuan_${dinhDangNgayNgan(weekRange.start)}-${dinhDangNgayNgan(weekRange.end)}_${targetY}`;
  }
  if (CONFIG.filterMode === "month")
    suffix = "Thang_" + targetM + "-" + targetY;
  if (CONFIG.filterMode === "year") suffix = "Nam_" + targetY;
  const taoTenSheetHopLe = (ten: string): string =>
    ten.replace(/[\\\/\?\*\[\]:]/g, "_").slice(0, 31);
  const TARGET_SHEET_NAME = taoTenSheetHopLe(
    CONFIG.output.sheetNamePrefix + suffix,
  );

  const sheetNguon = workbook.getWorksheet(CONFIG.tenSheetNguon);
  if (!sheetNguon) {
    console.log(`Không tìm thấy sheet nguồn "${CONFIG.tenSheetNguon}".`);
    return;
  }
  const usedRangeNguon = sheetNguon.getUsedRange();
  if (!usedRangeNguon) {
    console.log(`Sheet "${CONFIG.tenSheetNguon}" không có dữ liệu.`);
    return;
  }
  const rowCountNguon = usedRangeNguon.getRowCount();

  // Mảng lưu kết quả xuất ra
  let resultData: (string | number | boolean)[][] = [];
  let soDongCoNoiDung = 0;
  let soLanNhapKhoKhopNgay = 0;
  const dongCanhBaoTraCuuTheoTen: number[] = [];

  // Thêm hàng subtotal và header (Header xuống hàng 2)
  // TỔNG CỘNG ở I1; SUBTOTAL(9,…) ở J1 (SL nhập), N1 (tổng TL), S1 (tổng thể tích)
  // Cột T = Vị trí (từ sheet nguồn cột I); cột U trống; thống kê từ cột V
  resultData.push([
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "TỔNG CỘNG:",
    0,
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
  ]); // Hàng 1
  resultData.push([
    "Nơi nhập",
    "Ngày nhập kho",
    "Mã hợp đồng",
    "Thị trường",
    "Mã phân loại",
    "SKU",
    "Tên hàng",
    "Số lượng đơn hàng",
    "Số lượng đã nhập",
    "Số lượng nhập",
    "Số lượng còn lại",
    "Mã ghép",
    "Trọng lượng",
    "Tổng trọng lượng",
    "Dài",
    "Rộng",
    "Cao",
    "Thể tích",
    "Tổng thể tích",
    "Vị trí",
  ]); // Hàng 2

  // Parse số từ ô Excel (hỗ trợ text có dấu phẩy/chấm) — giống Auto_Mitsuha
  const lamSach = (giaTri: string | number | boolean): string =>
    String(giaTri)
      .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "")
      .trim();
  const chuanHoaDanhSachKyTu = (list: string[]): string[] => {
    const seen: { [key: string]: boolean } = {};
    const out: string[] = [];
    for (let x of list) {
      const k = String(x ?? "")
        .trim()
        .toUpperCase();
      if (!k || seen[k]) continue;
      seen[k] = true;
      out.push(k);
    }
    return out;
  };
  const dsKyTuMacDinh = chuanHoaDanhSachKyTu(CONFIG.thongKeKyTuMacDinh);
  const layDanhSachKyTuTheoThiTruong = (thiTruong: string): string[] => {
    const dsCauHinh = CONFIG.thongKeKyTuTheoThiTruong[thiTruong];
    if (!dsCauHinh || dsCauHinh.length === 0) return dsKyTuMacDinh;
    return chuanHoaDanhSachKyTu(dsCauHinh);
  };
  const layThiTruongTuHopDong = (hopDong: string): string => {
    const hd = String(hopDong ?? "")
      .trim()
      .toUpperCase();
    const match = hd.match(/^([A-Z]{2})/);
    return match ? match[1] : "KHAC";
  };
  const layKyTuPhanLoaiRaw = (hopDong: string, thiTruong: string): string => {
    const hd = String(hopDong ?? "")
      .trim()
      .toUpperCase();
    let phanConLai = hd;
    if (thiTruong !== "KHAC" && hd.startsWith(thiTruong)) {
      phanConLai = hd.slice(thiTruong.length);
    }
    const match = phanConLai.match(/[A-Z]/);
    return match ? match[0] : "(TRONG)";
  };
  const tinhMaPhanLoai = (hopDong: string): string => {
    const thiTruong = layThiTruongTuHopDong(hopDong);
    const dsKyTu = layDanhSachKyTuTheoThiTruong(thiTruong);
    const raw = layKyTuPhanLoaiRaw(hopDong, thiTruong);
    return dsKyTu.includes(raw) ? raw : "KHAC";
  };

  const parseSoTuExcel = (giaTri: string | number | boolean): number => {
    if (typeof giaTri === "number" && !isNaN(giaTri)) return giaTri;
    let s = lamSach(giaTri);
    if (!s) return 0;
    const lastComma = s.lastIndexOf(",");
    const lastDot = s.lastIndexOf(".");
    if (lastComma >= 0 && lastDot >= 0) {
      if (lastComma > lastDot) s = s.replace(/\./g, "").replace(",", ".");
      else s = s.replace(/,/g, "");
    } else if (lastComma >= 0) s = s.replace(",", ".");
    const n = Number(s);
    return isNaN(n) ? 0 : n;
  };
  const chuanHoaTenSanPham = (giaTri: string | number | boolean): string =>
    String(giaTri ?? "")
      .toLowerCase()
      .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  const tachTokenTenSanPham = (tenChuan: string): string[] => {
    if (!tenChuan) return [];
    return tenChuan.split(/\s+/).filter((t) => t.length > 0);
  };
  const tinhDiemKhopTenMo = (
    queryTokens: string[],
    candidateTokens: string[],
  ): number => {
    if (queryTokens.length === 0 || candidateTokens.length === 0) return 0;
    let giao = 0;
    let daDemTokenQ: { [token: string]: boolean } = {};
    const setC = new Set(candidateTokens);
    for (let i = 0; i < queryTokens.length; i++) {
      const token = queryTokens[i];
      if (daDemTokenQ[token]) continue;
      daDemTokenQ[token] = true;
      if (setC.has(token)) giao++;
    }
    let hop = 0;
    let daDemHop: { [token: string]: boolean } = {};
    for (let i = 0; i < queryTokens.length; i++) {
      const token = queryTokens[i];
      if (daDemHop[token]) continue;
      daDemHop[token] = true;
      hop++;
    }
    for (let i = 0; i < candidateTokens.length; i++) {
      const token = candidateTokens[i];
      if (daDemHop[token]) continue;
      daDemHop[token] = true;
      hop++;
    }
    return hop > 0 ? giao / hop : 0;
  };
  const demTokenTrung = (
    queryTokens: string[],
    candidateTokens: string[],
  ): number => {
    let giao = 0;
    const setC = new Set(candidateTokens);
    for (let token of queryTokens) {
      if (setC.has(token)) giao++;
    }
    return giao;
  };
  const layNguongTokenKhop = (
    soTokenTraCuu: number,
    config: NamiTraCuuTenMoConfig,
  ): number =>
    Math.max(
      config.soTokenKhopToiDa,
      Math.ceil(soTokenTraCuu * config.tyLeTokenKhopToiThieu),
    );
  const thongSoTenCoDuLieu = (thongSo: NamiProductInfo): boolean =>
    !!(
      thongSo.trongLuong ||
      thongSo.dai ||
      thongSo.rong ||
      thongSo.cao ||
      thongSo.theTich
    );
  const timThongSoTheoTenMo = (
    tenChuan: string,
    danhSach: NamiThongSoTheoTenEntry[],
    config: NamiTraCuuTenMoConfig,
  ): NamiProductInfo | undefined => {
    const queryTokens = tachTokenTenSanPham(tenChuan);
    if (queryTokens.length === 0 || danhSach.length === 0) return undefined;
    const nguongGiao = layNguongTokenKhop(queryTokens.length, config);

    let best: NamiProductInfo | undefined;
    let bestScore = 0;
    let bestGiao = 0;
    let bestCoDuLieu = false;
    for (let item of danhSach) {
      const giao = demTokenTrung(queryTokens, item.tokens);
      if (giao < nguongGiao) continue;
      const score = tinhDiemKhopTenMo(queryTokens, item.tokens);
      if (score < config.diemToiThieu) continue;

      const uuTien =
        score > bestScore ||
        (score === bestScore &&
          (giao > bestGiao ||
            (giao === bestGiao && item.coDuLieu && !bestCoDuLieu)));
      if (uuTien) {
        best = item.thongSo;
        bestScore = score;
        bestGiao = giao;
        bestCoDuLieu = item.coDuLieu;
      }
    }
    return best;
  };

  // Đọc thông số từ sheet cấu hình (tham chiếu logic Auto_Mitsuha)
  const lookupThongSo = new Map<string, NamiProductInfo>();
  const lookupThongSoTheoTen = new Map<string, NamiProductInfo>();
  const danhSachThongSoTheoTen: NamiThongSoTheoTenEntry[] = [];
  const tapTenDaThemVaoDanhSachMo = new Set<string>();
  if (CONFIG.thongSo.bat) {
    const sheetThongSo = workbook.getWorksheet(CONFIG.tenSheetThongSo);
    if (sheetThongSo) {
      const lastRowThongSo = namiTimDongCuoiSheetThongSo(sheetThongSo);
      if (lastRowThongSo >= 2) {
        const dataThongSo = sheetThongSo
          .getRange(`B2:H${lastRowThongSo}`)
          .getValues() as (string | number | boolean)[][];
        for (let row of dataThongSo) {
          const maHangThongSo = lamSach(row[0]);
          const tenHangThongSo = chuanHoaTenSanPham(row[1]);
          if (!maHangThongSo && !tenHangThongSo) continue;

          const trongLuongThongSo = parseSoTuExcel(row[3]);
          const dai = parseSoTuExcel(row[4]);
          const rong = parseSoTuExcel(row[5]);
          const caoGoc = parseSoTuExcel(row[6]);
          const cao = caoGoc > 0 ? caoGoc : dai > 0 && rong > 0 ? 1.1 : 0;
          const theTich = (dai * rong * cao) / 1000000;

          const thongSo: NamiProductInfo = {
            trongLuong: trongLuongThongSo
              ? Number(trongLuongThongSo.toFixed(4))
              : 0,
            dai: dai ? Number(dai.toFixed(4)) : 0,
            rong: rong ? Number(rong.toFixed(4)) : 0,
            cao: cao ? Number(cao.toFixed(4)) : 0,
            theTich: theTich ? Number(theTich.toFixed(3)) : 0,
          };

          if (maHangThongSo) {
            lookupThongSo.set(maHangThongSo, thongSo);
          }
          if (tenHangThongSo) {
            const thongSoMoiCoDuLieu = thongSoTenCoDuLieu(thongSo);
            const thongSoCu = lookupThongSoTheoTen.get(tenHangThongSo);
            const thongSoCuRong = !thongSoCu || !thongSoTenCoDuLieu(thongSoCu);
            if (thongSoCuRong || thongSoMoiCoDuLieu) {
              lookupThongSoTheoTen.set(tenHangThongSo, thongSo);
              const entryMoi = {
                tenChuan: tenHangThongSo,
                tokens: tachTokenTenSanPham(tenHangThongSo),
                thongSo: thongSo,
                coDuLieu: thongSoMoiCoDuLieu,
              };
              if (tapTenDaThemVaoDanhSachMo.has(tenHangThongSo)) {
                const idx = danhSachThongSoTheoTen.findIndex(
                  (item) => item.tenChuan === tenHangThongSo,
                );
                if (idx >= 0) danhSachThongSoTheoTen[idx] = entryMoi;
              } else {
                tapTenDaThemVaoDanhSachMo.add(tenHangThongSo);
                danhSachThongSoTheoTen.push(entryMoi);
              }
            }
          }
        }
        console.log(
          `Đã nạp thông số từ "${CONFIG.tenSheetThongSo}": ${lookupThongSo.size} mã, ${danhSachThongSoTheoTen.length} tên (tra cứu mờ), ${dataThongSo.length} dòng đọc.`,
        );
      }
    } else {
      console.log(
        `Lưu ý: Không tìm thấy sheet "${CONFIG.tenSheetThongSo}", các cột thông số sẽ để trống.`,
      );
    }
  } else {
    console.log(
      "Đang bỏ qua tra cứu thông số (thongSo.bat=false) để tối ưu tốc độ chạy.",
    );
  }

  const layO = (
    rowTexts: string[],
    rowValues: (string | number | boolean)[],
    idx: number,
  ): string | number | boolean => {
    if (rowValues.length > idx && rowValues[idx] !== "") return rowValues[idx];
    if (rowTexts.length > idx) return rowTexts[idx] || "";
    return "";
  };

  // Tổng SL đã nhập = K + M + O (cặp J/L/N là ngày; SL nhập kho ở K, M, O)
  const tongNhapTuCacDot = (
    rowTexts: string[],
    rowValues: (string | number | boolean)[],
  ): number =>
    parseSoTuExcel(layO(rowTexts, rowValues, 10)) +
    parseSoTuExcel(layO(rowTexts, rowValues, 12)) +
    parseSoTuExcel(layO(rowTexts, rowValues, 14));

  const layThongSoHang = (
    sku: string | number | boolean,
    tenHang: string | number | boolean,
  ): AutoNamiThongSoTraCuuResult | undefined => {
    const maHang = lamSach(sku);
    if (maHang) {
      const thongSoTheoMa = lookupThongSo.get(maHang);
      if (thongSoTheoMa) return { thongSo: thongSoTheoMa, nguon: "ma" };
    }
    const tenHangChuan = chuanHoaTenSanPham(tenHang);
    if (!tenHangChuan) return undefined;
    const thongSoTheoTenChinhXac = lookupThongSoTheoTen.get(tenHangChuan);
    if (thongSoTheoTenChinhXac) {
      return { thongSo: thongSoTheoTenChinhXac, nguon: "ten" };
    }
    if (!CONFIG.traCuuTenMo.bat || danhSachThongSoTheoTen.length === 0) {
      return undefined;
    }
    const thongSoTheoTenMo = timThongSoTheoTenMo(
      tenHangChuan,
      danhSachThongSoTheoTen,
      CONFIG.traCuuTenMo,
    );
    return thongSoTheoTenMo
      ? { thongSo: thongSoTheoTenMo, nguon: "ten-mo" }
      : undefined;
  };

  // Hàm so sánh ngày dựa trên FILTER_MODE
  const getDateMatch = (
    textDate: NamiDateCellValue,
    valueDate: NamiDateCellValue,
  ): string | null => {
    if (!textDate && !valueDate) return null;
    const normalizedDate = namiLayNgayKeyTuO(textDate, valueDate);
    if (!normalizedDate) return null;
    const dateParts = normalizedDate.split("-");
    const y = dateParts[0];
    const m = dateParts[1];
    let isMatch = false;
    if (CONFIG.filterMode === "day") {
      isMatch = normalizedDate === targetDateKey;
    } else if (CONFIG.filterMode === "week") {
      isMatch = normalizedDate >= weekStartKey && normalizedDate <= weekEndKey;
    } else if (CONFIG.filterMode === "month") {
      isMatch = m === targetM && y === targetY;
    } else if (CONFIG.filterMode === "year") {
      isMatch = y === targetY;
    }

    return isMatch ? normalizedDate : null;
  };

  // F = SL đơn hàng; SL đã nhập = K+M+O; cột “còn lại” = (K+M+O) − F (âm khi chưa nhập đủ; không dùng P/Q)
  const dataTextsNguon = usedRangeNguon.getTexts();
  const dataValuesNguon = usedRangeNguon.getValues() as (
    | string
    | number
    | boolean
  )[][];
  console.log(
    `Đang quét ${rowCountNguon - 1} dòng "${CONFIG.tenSheetNguon}" (UsedRange) cho ${filterLabel}.`,
  );
  const viTriIdx =
    namiCotChuSangIndex(CONFIG.cotNguon.viTri) -
    usedRangeNguon.getColumnIndex();
  for (let i = 1; i < rowCountNguon; i++) {
    let rowTexts = dataTextsNguon[i];
    let rowValues = dataValuesNguon[i];

    let sku = rowTexts.length > 2 ? rowTexts[2] || "" : "";
    let maHD = rowTexts.length > 7 ? rowTexts[7] || "" : "";
    let tenHang = rowTexts.length > 3 ? rowTexts[3] || "" : "";
    let viTri = namiChuanHoaViTri(
      namiLayGiaTriO(rowTexts, rowValues, viTriIdx),
    );

    // Bỏ qua nếu dòng trống không có SKU hay Mã HĐ
    if (!sku && !maHD && !tenHang) continue;
    soDongCoNoiDung++;

    const slDonHang = layO(rowTexts, rowValues, 5); // Cột F
    const tongNhapTinh = tongNhapTuCacDot(rowTexts, rowValues);
    const slDaNhapTong = Number(tongNhapTinh.toFixed(4));
    // Âm khi chưa nhập đủ (đảo dấu so với đơn hàng − tổng nhập)
    const slConLai = Number(
      (tongNhapTinh - parseSoTuExcel(slDonHang)).toFixed(4),
    );
    const thiTruong = maHD ? layThiTruongTuHopDong(String(maHD)) : "";
    const maPhanLoai = maHD ? tinhMaPhanLoai(String(maHD)) : "";

    let thongSoDongXuat:
      | {
          trongLuong: string | number;
          dai: string | number;
          rong: string | number;
          cao: string | number;
          theTich: string | number;
          nguon?: NamiThongSoNguon;
        }
      | undefined;
    const layThongSoChoDongXuat = () => {
      if (!thongSoDongXuat) {
        const ketQuaTraCuu = layThongSoHang(sku, tenHang);
        const thongSoHang = ketQuaTraCuu?.thongSo;
        thongSoDongXuat = {
          trongLuong: thongSoHang?.trongLuong || "",
          dai: thongSoHang?.dai || "",
          rong: thongSoHang?.rong || "",
          cao: thongSoHang?.cao || "",
          theTich: thongSoHang?.theTich || "",
          nguon: ketQuaTraCuu?.nguon,
        };
      }
      return thongSoDongXuat;
    };

    // --- Cột J & K: Nhập kho 1 ---
    if (rowTexts.length > 9) {
      let matchDate = getDateMatch(rowTexts[9], rowValues[9]);
      if (matchDate) {
        const { trongLuong, dai, rong, cao, theTich, nguon } =
          layThongSoChoDongXuat();
        let sl =
          rowValues.length > 10 && rowValues[10] !== ""
            ? rowValues[10]
            : rowTexts.length > 10
              ? rowTexts[10]
              : "";
        const soDongSheet = resultData.length + 1;
        const tongTrongLuong = `=M${soDongSheet}*J${soDongSheet}`;
        if (nguon === "ten" || nguon === "ten-mo") {
          dongCanhBaoTraCuuTheoTen.push(soDongSheet);
        }
        resultData.push([
          "Nhập kho 1",
          matchDate,
          maHD,
          thiTruong,
          maPhanLoai,
          sku,
          tenHang,
          slDonHang,
          slDaNhapTong,
          sl,
          slConLai,
          String(sku) + String(maHD),
          trongLuong,
          tongTrongLuong,
          dai,
          rong,
          cao,
          theTich,
          `=R${soDongSheet}*J${soDongSheet}`,
          viTri,
        ]);
        soLanNhapKhoKhopNgay++;
      }
    }

    // --- Cột L & M: Nhập kho 2 ---
    if (rowTexts.length > 11) {
      let matchDate = getDateMatch(rowTexts[11], rowValues[11]);
      if (matchDate) {
        const { trongLuong, dai, rong, cao, theTich, nguon } =
          layThongSoChoDongXuat();
        let sl =
          rowValues.length > 12 && rowValues[12] !== ""
            ? rowValues[12]
            : rowTexts.length > 12
              ? rowTexts[12]
              : "";
        const soDongSheet = resultData.length + 1;
        const tongTrongLuong = `=M${soDongSheet}*J${soDongSheet}`;
        if (nguon === "ten" || nguon === "ten-mo") {
          dongCanhBaoTraCuuTheoTen.push(soDongSheet);
        }
        resultData.push([
          "Nhập kho 2",
          matchDate,
          maHD,
          thiTruong,
          maPhanLoai,
          sku,
          tenHang,
          slDonHang,
          slDaNhapTong,
          sl,
          slConLai,
          String(sku) + String(maHD),
          trongLuong,
          tongTrongLuong,
          dai,
          rong,
          cao,
          theTich,
          `=R${soDongSheet}*J${soDongSheet}`,
          viTri,
        ]);
        soLanNhapKhoKhopNgay++;
      }
    }

    // --- Cột N & O: Nhập kho 3 ---
    if (rowTexts.length > 13) {
      let matchDate = getDateMatch(rowTexts[13], rowValues[13]);
      if (matchDate) {
        const { trongLuong, dai, rong, cao, theTich, nguon } =
          layThongSoChoDongXuat();
        let sl =
          rowValues.length > 14 && rowValues[14] !== ""
            ? rowValues[14]
            : rowTexts.length > 14
              ? rowTexts[14]
              : "";
        const soDongSheet = resultData.length + 1;
        const tongTrongLuong = `=M${soDongSheet}*J${soDongSheet}`;
        if (nguon === "ten" || nguon === "ten-mo") {
          dongCanhBaoTraCuuTheoTen.push(soDongSheet);
        }
        resultData.push([
          "Nhập kho 3",
          matchDate,
          maHD,
          thiTruong,
          maPhanLoai,
          sku,
          tenHang,
          slDonHang,
          slDaNhapTong,
          sl,
          slConLai,
          String(sku) + String(maHD),
          trongLuong,
          tongTrongLuong,
          dai,
          rong,
          cao,
          theTich,
          `=R${soDongSheet}*J${soDongSheet}`,
          viTri,
        ]);
        soLanNhapKhoKhopNgay++;
      }
    }
  }

  const taoHoacLamMoiSheetKetQua = (): ExcelScript.Worksheet => {
    let targetSheet = workbook.getWorksheet(TARGET_SHEET_NAME);
    if (targetSheet) {
      targetSheet.getUsedRange()?.clear();
      return targetSheet;
    }
    return workbook.addWorksheet(TARGET_SHEET_NAME);
  };
  const apDungLocVaFreezeTieuDe = (
    targetSheet: ExcelScript.Worksheet,
    lastRow: number,
  ) => {
    const headerRow = CONFIG.output.headerRow;
    if (CONFIG.output.batAutoFilter && lastRow >= headerRow) {
      const autoFilter = targetSheet.getAutoFilter();
      if (autoFilter.getEnabled()) {
        autoFilter.remove();
      }
      autoFilter.apply(`A${headerRow}:${CONFIG.output.lastColLetter}${lastRow}`);
    }

    if (CONFIG.output.freezeDenDongTieuDe) {
      const freezePanes = targetSheet.getFreezePanes();
      freezePanes.unfreeze();
      freezePanes.freezeRows(headerRow);
    }
  };
  const doiCotSangSo = (cot: string): number => {
    let result = 0;
    const normalized = cot.trim().toUpperCase();
    for (let i = 0; i < normalized.length; i++) {
      result = result * 26 + (normalized.charCodeAt(i) - 64);
    }
    return result;
  };
  const doiSoSangCot = (soCot: number): string => {
    let result = "";
    let n = soCot;
    while (n > 0) {
      const phanDu = (n - 1) % 26;
      result = String.fromCharCode(65 + phanDu) + result;
      n = Math.floor((n - 1) / 26);
    }
    return result;
  };
  const layCotThongKe = (offset: number): string =>
    doiSoSangCot(doiCotSangSo(CONFIG.thongKeKetQua.cotBatDau) + offset);
  const taoChuoiCongThucExcel = (value: string): string =>
    `"${value.replace(/"/g, '""')}"`;
  const apDungVien = (range: ExcelScript.Range) => {
    [
      ExcelScript.BorderIndex.edgeTop,
      ExcelScript.BorderIndex.edgeBottom,
      ExcelScript.BorderIndex.edgeLeft,
      ExcelScript.BorderIndex.edgeRight,
      ExcelScript.BorderIndex.insideHorizontal,
      ExcelScript.BorderIndex.insideVertical,
    ].forEach((borderIndex) => {
      range
        .getFormat()
        .getRangeBorder(borderIndex)
        .setStyle(ExcelScript.BorderLineStyle.continuous);
    });
  };
  const sapXepMaPhanLoai = (
    a: string,
    b: string,
    thiTruong: string = "",
  ): number => {
    const thuTu = [...layDanhSachKyTuTheoThiTruong(thiTruong), "KHAC"];
    const idxA = thuTu.indexOf(a);
    const idxB = thuTu.indexOf(b);
    if (idxA >= 0 || idxB >= 0) {
      if (idxA < 0) return 1;
      if (idxB < 0) return -1;
      if (idxA !== idxB) return idxA - idxB;
    }
    return a.localeCompare(b);
  };
  const apDungDinhDangTieuDeThongKe = (
    targetSheet: ExcelScript.Worksheet,
    rangeAddress: string,
  ) => {
    const format = targetSheet.getRange(rangeAddress).getFormat();
    format.getFont().setName(CONFIG.format.fontName);
    format.getFont().setSize(CONFIG.format.fontSize);
    format.getFont().setBold(true);
    format.getFill().setColor(CONFIG.format.subtotalFillColor);
    format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  };
  const apDungDinhDangHeaderThongKe = (
    targetSheet: ExcelScript.Worksheet,
    rangeAddress: string,
  ) => {
    const format = targetSheet.getRange(rangeAddress).getFormat();
    const font = format.getFont();
    font.setName(CONFIG.format.fontName);
    font.setSize(CONFIG.format.fontSize);
    font.setBold(true);
    font.setColor("white");
    format.getFill().setColor(CONFIG.format.headerFillColor);
    format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  };
  const ghiThongKeKetQua = (
    targetSheet: ExcelScript.Worksheet,
    lastRow: number,
  ) => {
    if (!CONFIG.thongKeKetQua.bat || lastRow < 3) return;

    const tapMaPhanLoai = new Set<string>();
    const tapThiTruongMaPhanLoai = new Set<string>();
    for (let i = 2; i < resultData.length; i++) {
      const row = resultData[i];
      const thiTruong = String(row[3] ?? "").trim();
      const maPhanLoai = String(row[4] ?? "").trim();
      if (!maPhanLoai) continue;
      tapMaPhanLoai.add(maPhanLoai);
      if (thiTruong) {
        tapThiTruongMaPhanLoai.add(`${thiTruong}\u0001${maPhanLoai}`);
      }
    }

    const dsMaPhanLoai = Array.from(tapMaPhanLoai).sort((a, b) =>
      sapXepMaPhanLoai(a, b),
    );
    const dsThiTruongMaPhanLoai = Array.from(tapThiTruongMaPhanLoai)
      .map((key) => {
        const parts = key.split("\u0001");
        return { thiTruong: parts[0], maPhanLoai: parts[1] };
      })
      .sort(
        (a, b) =>
          a.thiTruong.localeCompare(b.thiTruong) ||
          sapXepMaPhanLoai(a.maPhanLoai, b.maPhanLoai, a.thiTruong),
      );
    if (dsMaPhanLoai.length === 0 && dsThiTruongMaPhanLoai.length === 0) return;

    const heSoChia =
      CONFIG.thongKeKetQua.heSoChiaTheTich > 0
        ? CONFIG.thongKeKetQua.heSoChiaTheTich
        : 68;
    const vung = (cot: string): string => `$${cot}$3:$${cot}$${lastRow}`;
    const sumIfMa = (cotTong: string, maPhanLoai: string): string =>
      `SUMIF(${vung("E")},${taoChuoiCongThucExcel(maPhanLoai)},${vung(cotTong)})`;
    const sumIfsThiTruong = (
      cotTong: string,
      item: NamiThiTruongMaPhanLoai,
    ): string =>
      `SUMIFS(${vung(cotTong)},${vung("D")},${taoChuoiCongThucExcel(
        item.thiTruong,
      )},${vung("E")},${taoChuoiCongThucExcel(item.maPhanLoai)})`;

    const ghiBangThongKe = <T>(
      dongBatDau: number,
      tieuDe: string,
      dsCot: NamiCotThongKe<T>[],
      dsItem: T[],
    ): number => {
      const cotDau = layCotThongKe(0);
      const cotCuoi = layCotThongKe(dsCot.length - 1);
      const dongDuLieuDau = dongBatDau + 2;
      const rows: (string | number | boolean)[][] = [
        dsCot.map((_, idx) => (idx === 0 ? tieuDe : "")),
        dsCot.map((cot) => cot.tieuDe),
      ];
      dsItem.forEach((item, idx) => {
        rows.push(dsCot.map((cot) => cot.giaTri(item, dongDuLieuDau + idx)));
      });
      const dongCuoi = dongBatDau + rows.length - 1;

      targetSheet
        .getRange(`${cotDau}${dongBatDau}:${cotCuoi}${dongCuoi}`)
        .setValues(rows);
      apDungDinhDangTieuDeThongKe(
        targetSheet,
        `${cotDau}${dongBatDau}:${cotCuoi}${dongBatDau}`,
      );
      apDungDinhDangHeaderThongKe(
        targetSheet,
        `${cotDau}${dongBatDau + 1}:${cotCuoi}${dongBatDau + 1}`,
      );
      apDungVien(
        targetSheet.getRange(`${cotDau}${dongBatDau}:${cotCuoi}${dongCuoi}`),
      );

      if (dsItem.length > 0) {
        const dataRange = targetSheet.getRange(
          `${cotDau}${dongDuLieuDau}:${cotCuoi}${dongCuoi}`,
        );
        dataRange.getFormat().getFont().setName(CONFIG.format.fontName);
        dataRange.getFormat().getFont().setSize(CONFIG.format.fontSize);
        dataRange
          .getFormat()
          .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        dsCot.forEach((cot, idx) => {
          if (!cot.dinhDangSo) return;
          const cotChu = layCotThongKe(idx);
          targetSheet
            .getRange(`${cotChu}${dongDuLieuDau}:${cotChu}${dongCuoi}`)
            .setNumberFormat(cot.dinhDangSo);
        });
      }
      return dongCuoi;
    };

    const dinhNghiaCotMa: Record<
      NamiCotBangMaPhanLoai,
      NamiCotThongKe<string>
    > = {
      maPhanLoai: { tieuDe: "Mã phân loại", giaTri: (ma) => ma },
      soLuong: {
        tieuDe: "Số lượng nhập",
        dinhDangSo: "0",
        giaTri: (ma) => `=ROUND(${sumIfMa("J", ma)},0)`,
      },
      trongLuong: {
        tieuDe: "Tổng trọng lượng",
        dinhDangSo: "0.000",
        giaTri: (ma) => `=${sumIfMa("N", ma)}`,
      },
      theTich: {
        tieuDe: "Tổng thể tích",
        dinhDangSo: "0.000",
        giaTri: (ma) => `=${sumIfMa("S", ma)}`,
      },
    };

    const thuTuCotMa = namiChuanHoaThuTuCot(
      CONFIG.thongKeKetQua.thuTuCotBangMaPhanLoai,
      NAMI_COT_BANG_MA_PHAN_LOAI_MAC_DINH,
    );
    const thuTuCotThiTruong = namiChuanHoaThuTuCot(
      CONFIG.thongKeKetQua.thuTuCotBangThiTruong,
      NAMI_COT_BANG_THI_TRUONG_MAC_DINH,
    );
    const idxCotTheTich = thuTuCotThiTruong.indexOf("theTich");

    const dinhNghiaCotThiTruong: Record<
      NamiCotBangThiTruong,
      NamiCotThongKe<NamiThiTruongMaPhanLoai>
    > = {
      thiTruong: { tieuDe: "Thị trường", giaTri: (item) => item.thiTruong },
      maPhanLoai: { tieuDe: "Mã phân loại", giaTri: (item) => item.maPhanLoai },
      soLuong: {
        tieuDe: "Số lượng nhập",
        dinhDangSo: "0",
        giaTri: (item) => `=ROUND(${sumIfsThiTruong("J", item)},0)`,
      },
      trongLuong: {
        tieuDe: "Tổng trọng lượng",
        dinhDangSo: "0.000",
        giaTri: (item) => `=${sumIfsThiTruong("N", item)}`,
      },
      theTich: {
        tieuDe: "Tổng thể tích",
        dinhDangSo: "0.000",
        giaTri: (item) => `=${sumIfsThiTruong("S", item)}`,
      },
      theTichChia: {
        tieuDe: `Tổng thể tích / ${heSoChia}`,
        dinhDangSo: "0.000",
        giaTri: (item, soDong) =>
          idxCotTheTich >= 0
            ? `=${layCotThongKe(idxCotTheTich)}${soDong}/${heSoChia}`
            : `=${sumIfsThiTruong("S", item)}/${heSoChia}`,
      },
    };

    // Tiêu đề bảng thống kê nằm dưới vùng freeze panes (2 dòng đầu).
    const thongKeMaLastRow = ghiBangThongKe(
      Math.max(CONFIG.thongKeKetQua.dongBatDau, 1),
      "THỐNG KÊ THEO MÃ PHÂN LOẠI",
      thuTuCotMa.map((key) => dinhNghiaCotMa[key]),
      dsMaPhanLoai,
    );

    const thongKeThiTruongStartRow = thongKeMaLastRow + 2;
    ghiBangThongKe(
      thongKeThiTruongStartRow,
      "THỐNG KÊ THỊ TRƯỜNG THEO MÃ PHÂN LOẠI",
      thuTuCotThiTruong.map((key) => dinhNghiaCotThiTruong[key]),
      dsThiTruongMaPhanLoai,
    );

    const idxCotThiTruong = thuTuCotThiTruong.indexOf("thiTruong");
    if (idxCotThiTruong >= 0 && dsThiTruongMaPhanLoai.length > 0) {
      const cotThiTruong = layCotThongKe(idxCotThiTruong);
      const dataStartRow = thongKeThiTruongStartRow + 2;
      let groupStartRow = dataStartRow;
      for (let i = 1; i <= dsThiTruongMaPhanLoai.length; i++) {
        const prev = dsThiTruongMaPhanLoai[i - 1];
        const current = dsThiTruongMaPhanLoai[i];
        if (current && current.thiTruong === prev.thiTruong) continue;

        const groupEndRow = dataStartRow + i - 1;
        if (groupEndRow > groupStartRow) {
          const marketRange = targetSheet.getRange(
            `${cotThiTruong}${groupStartRow}:${cotThiTruong}${groupEndRow}`,
          );
          marketRange.merge(false);
          marketRange
            .getFormat()
            .setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        }
        groupStartRow = groupEndRow + 1;
      }
    }
  };

  if (resultData.length === 2) {
    const targetSheet = taoHoacLamMoiSheetKetQua();
    const lastColLetter = CONFIG.output.lastColLetter;
    targetSheet
      .getRange(`A1:${lastColLetter}${resultData.length}`)
      .setValues(resultData);
    targetSheet.getUsedRange()?.getFormat().autofitColumns();
    apDungLocVaFreezeTieuDe(targetSheet, resultData.length);
    targetSheet.activate();
    console.log(
      `Không tìm thấy dữ liệu nhập kho nào cho ${filterLabel}. Đã quét ${soDongCoNoiDung} dòng có nội dung, số lần nhập kho khớp ngày: ${soLanNhapKhoKhopNgay}.`,
    );
    return;
  }

  // Công thức SUBTOTAL hàng 1 (chỉ tính trên dòng dữ liệu từ hàng 3)
  const hangCuoiDuLieu = resultData.length;
  resultData[0][9] = `=SUBTOTAL(9, J3:J${hangCuoiDuLieu})`;
  resultData[0][13] = `=SUBTOTAL(9, N3:N${hangCuoiDuLieu})`;
  resultData[0][18] = `=SUBTOTAL(9, S3:S${hangCuoiDuLieu})`;

  // --- XUẤT DỮ LIỆU RA SHEET MỚI ---
  let targetSheet = taoHoacLamMoiSheetKetQua();

  // Ghi dữ liệu
  const lastColLetter = CONFIG.output.lastColLetter;
  const targetRange = targetSheet.getRange(
    `A1:${lastColLetter}${resultData.length}`,
  );
  targetRange.setValues(resultData);

  // --- ĐỊNH DẠNG (FORMAT) ---
  // 1. Tô nền / định dạng ô tổng (I1, J1, N1, S1)
  const apDungDinhDangSubtotal = (range: ExcelScript.Range) => {
    const fmt = range.getFormat();
    fmt.getFont().setName(CONFIG.format.fontName);
    fmt.getFont().setSize(CONFIG.format.fontSize);
    fmt.getFont().setBold(true);
    fmt.getFill().setColor(CONFIG.format.subtotalFillColor);
    fmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  };
  apDungDinhDangSubtotal(targetSheet.getRange(CONFIG.format.subtotalRange));
  apDungDinhDangSubtotal(targetSheet.getRange("J1"));
  apDungDinhDangSubtotal(targetSheet.getRange("N1"));
  apDungDinhDangSubtotal(targetSheet.getRange("S1"));

  // 2. Định dạng hàng Tiêu đề (Hàng 2)
  const headerRange = targetSheet.getRange(`A2:${lastColLetter}2`);
  let headerFont = headerRange.getFormat().getFont();
  headerFont.setName(CONFIG.format.fontName);
  headerFont.setSize(CONFIG.format.fontSize);
  headerFont.setBold(true);
  headerFont.setColor("white");
  headerRange.getFormat().getFill().setColor(CONFIG.format.headerFillColor); // Màu nền xanh dương
  headerRange
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  headerRange
    .getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);

  // 3. Định dạng dữ liệu (Từ hàng 3)
  if (resultData.length > 2) {
    const dataRange = targetSheet.getRange(
      `A3:${lastColLetter}${resultData.length}`,
    );
    let dataFont = dataRange.getFormat().getFont();
    dataFont.setName(CONFIG.format.fontName);
    dataFont.setSize(CONFIG.format.fontSize);
    dataRange
      .getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Định dạng cột Ngày nhập kho (Cột B) thành dd/mm/yyyy (Cố định trên mọi máy)
    targetSheet
      .getRange(
        `${CONFIG.format.dateColumn}3:${CONFIG.format.dateColumn}${resultData.length}`,
      )
      .setNumberFormat("dd/mm/yyyy");

    // Căn trái cho Mã HĐ (C), Thị trường (D), Mã phân loại (E), SKU (F), Mã ghép (L), Vị trí (T)
    CONFIG.format.leftAlignColumns.forEach((col) => {
      targetSheet
        .getRange(`${col}3:${col}${resultData.length}`)
        .getFormat()
        .setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    });

    targetSheet.getRange(`T3:T${resultData.length}`).getFormat().setWrapText(true);
    targetSheet.getRange(`M3:Q${resultData.length}`).setNumberFormat("General");
    targetSheet.getRange(`R3:S${resultData.length}`).setNumberFormat("0.000");

    for (const dong of dongCanhBaoTraCuuTheoTen) {
      targetSheet
        .getRange(`F${dong}`)
        .getFormat()
        .getFill()
        .setColor(CONFIG.format.tenLookupWarningFillColor);
    }

    // Kẻ viền cho toàn bộ bảng (A1 đến cuối)
    [
      ExcelScript.BorderIndex.edgeTop,
      ExcelScript.BorderIndex.edgeBottom,
      ExcelScript.BorderIndex.edgeLeft,
      ExcelScript.BorderIndex.edgeRight,
      ExcelScript.BorderIndex.insideHorizontal,
      ExcelScript.BorderIndex.insideVertical,
    ].forEach((bIndex) => {
      targetRange
        .getFormat()
        .getRangeBorder(bIndex)
        .setStyle(ExcelScript.BorderLineStyle.continuous);
    });
  }

  ghiThongKeKetQua(targetSheet, resultData.length);

  // Tự động điều chỉnh độ rộng cột; giữ cột U hẹp làm khoảng cách T ↔ bảng thống kê
  targetSheet.getUsedRange()?.getFormat().autofitColumns();
  targetSheet.getRange("U:U").getFormat().setColumnWidth(3);

  apDungLocVaFreezeTieuDe(targetSheet, resultData.length);

  // Kích hoạt sheet
  targetSheet.activate();

  console.log(
    `Đã xuất thành công ${resultData.length - 2} dòng dữ liệu cho ${filterLabel}. Đã quét ${soDongCoNoiDung} dòng có nội dung, số lần nhập kho khớp ngày: ${soLanNhapKhoKhopNgay}.`,
  );
}
