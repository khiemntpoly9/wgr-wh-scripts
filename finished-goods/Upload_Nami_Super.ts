/**
 * ===================================================================
 * Upload_Nami_Super – Tải dữ liệu nhập kho lên hệ thống trong MỘT bước.
 * ===================================================================
 *
 * Đọc thẳng sheet "Thành phẩm", lấy các lần nhập kho (Nhập kho 1/2/3) thuộc kỳ đã chọn
 * và gửi tới POST /api/v1/warehouse/receipts/nami-direct-import.
 * Không cần chạy Auto_Nami, không cần sheet KQ_NhapKho_*.
 *
 * Cột "Thành phẩm" được dùng:
 * C SKU | D Tên hàng | F SL đơn hàng | H Mã hợp đồng | I Vị trí
 * J/K Ngày + SL nhập kho 1 | L/M Ngày + SL nhập kho 2 | N/O Ngày + SL nhập kho 3
 *
 * - Thông số sản phẩm: ưu tiên dữ liệu đã có trong hệ thống; sheet "2026" chỉ dùng để điền
 *   phần hệ thống còn thiếu (tra theo mã → tên → tên gần đúng như Auto_Nami).
 * - Chạy lại an toàn: lần nhập đã có trên hệ thống được tính "đã có từ trước", không bị trùng,
 *   kể cả khi trước đó đã tải bằng Upload_Nami.
 * - Lần nhập bị bỏ qua/từ chối được tô màu #F4B084 ở đúng cặp ô ngày + SL (J:K / L:M / N:O).
 *   Sửa tại "Thành phẩm" rồi chạy lại: ô được trả về màu cũ (ghi trong sheet ẩn _UploadNamiSuper);
 *   ô bạn tự tô màu khác không bị đụng tới.
 * - Báo cáo kỳ vừa tải: chạy Auto_Nami_Supper với cùng ngày/chế độ.
 *
 * QUAN TRỌNG — Office Scripts chạy trên cloud Microsoft (Excel Online):
 * - KHÔNG dùng localhost / 127.0.0.1; BẮT BUỘC dùng HTTPS.
 *
 * Cách dùng:
 * 1. Tạo API token tại Dashboard → Hệ thống & Quyền → API Token (Excel), dán vào apiToken.
 * 2. Đặt ngay (dd/mm/yyyy, trống = hôm nay) và cheDo ("day" | "week" | "month").
 * 3. Chạy script.
 */

type NamiSuperCheDo = "day" | "week" | "month";
type NamiSuperO = string | number | boolean;
type NamiSuperNguonThongSo = "ma" | "ten" | "ten-mo";

type NamiSuperKy = {
  mode: NamiSuperCheDo;
  from: string;
  to: string;
  nhan: string;
};

type NamiSuperLanNhap = {
  source_row: number;
  round_number: number;
  receipt_date: string;
  contract_no: string;
  sku: string;
  product_name: string;
  order_quantity: number;
  total_received: number;
  quantity: number;
  location_code: string;
  unit_weight?: number;
  length?: number;
  width?: number;
  height?: number;
  unit_volume?: number;
  spec_source?: NamiSuperNguonThongSo;
};

type NamiSuperBoQua = {
  source_row: number;
  round_number: number;
  sku: string;
  contract_no: string;
  reason: string;
};

type NamiSuperTuChoi = {
  row: number;
  round_number: number;
  sku: string;
  contract_no: string;
  order_quantity?: number;
  code?: string;
  reason: string;
};

type NamiSuperThongSo = {
  trongLuong: number;
  dai: number;
  rong: number;
  cao: number;
  theTich: number;
};

type NamiSuperKetQuaTraCuu = {
  thongSo: NamiSuperThongSo;
  nguon: NamiSuperNguonThongSo;
};

type NamiSuperTraCuuTenMo = {
  bat: boolean;
  diemToiThieu: number;
  tyLeTokenKhopToiThieu: number;
  soTokenKhopToiDa: number;
};

type NamiSuperKetQuaDot = {
  imported: number;
  unchanged: number;
  rejected: NamiSuperTuChoi[];
  receipts: string[];
  specs_from_sheet: { sku: string; source: NamiSuperNguonThongSo }[];
  missing_specs: string[];
};

const UPLOAD_NAMI_SUPER_CONFIG = {
  apiUrl:
    "https://app.wgr-wh.site/api/v1/warehouse/receipts/nami-direct-import",
  apiToken: "", // Bearer token từ Dashboard → API Token (Excel)

  /** dd/mm/yyyy; để trống = hôm nay (giờ Việt Nam). */
  ngay: "",
  /** "day" = theo ngày, "week" = thứ Hai → Chủ nhật chứa ngày, "month" = cả tháng. */
  cheDo: "day" as NamiSuperCheDo,

  tenSheetNguon: "Thành phẩm",
  tenSheetThongSo: "2026",
  /** false = không đọc sheet "2026" (chỉ dùng thông số có sẵn trên hệ thống). */
  traThongSoSheet: true,
  traCuuTenMo: {
    bat: true,
    diemToiThieu: 0.55,
    tyLeTokenKhopToiThieu: 0.5,
    soTokenKhopToiDa: 3,
  } as NamiSuperTraCuuTenMo,

  /** Số lần nhập tối đa mỗi request. */
  kichThuocDot: 300,
  /** Màu tô cặp ô ngày + SL của lần nhập bị bỏ qua/từ chối. */
  mauDanhDau: "#F4B084",
  lechMuiGio: 7,
};

/** Cột (0-based) của "Thành phẩm". */
const NAMI_SUPER_COT = {
  sku: 2,
  ten: 3,
  slDonHang: 5,
  hopDong: 7,
  viTri: 8,
  ngayNhap1: 9,
};

const namiSuperSerialCache = new Map<number, string | null>();
const namiSuperTextCache = new Map<string, string | null>();

const namiSuperLamSach = (giaTri: NamiSuperO): string =>
  String(giaTri ?? "")
    .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "")
    .trim();

const namiSuperParseSo = (giaTri: NamiSuperO): number => {
  if (typeof giaTri === "number") return isNaN(giaTri) ? 0 : giaTri;
  let s = namiSuperLamSach(giaTri);
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

const namiSuperTaoNgayKey = (
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

const namiSuperSerialSangNgayKey = (serial: number): string | null => {
  if (namiSuperSerialCache.has(serial)) {
    return namiSuperSerialCache.get(serial) ?? null;
  }
  if (!Number.isFinite(serial) || serial <= 0) return null;
  const dt = new Date(Math.floor(serial - 25569) * 86400000);
  const result = namiSuperTaoNgayKey(
    dt.getUTCFullYear(),
    dt.getUTCMonth() + 1,
    dt.getUTCDate(),
  );
  namiSuperSerialCache.set(serial, result);
  return result;
};

const namiSuperTextSangNgayKey = (raw: NamiSuperO): string | null => {
  if (typeof raw !== "string") return null;
  const text = raw.trim();
  if (!text) return null;
  if (namiSuperTextCache.has(text)) {
    return namiSuperTextCache.get(text) ?? null;
  }

  let result: string | null = null;
  const matched = text.match(/(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})/);
  if (matched) {
    const [, p1, p2, p3] = matched;
    if (p1.length === 4) {
      result = namiSuperTaoNgayKey(Number(p1), Number(p2), Number(p3));
    }
    if (!result) {
      const yyyy = p3.length === 2 ? 2000 + Number(p3) : Number(p3);
      const n1 = Number(p1);
      const n2 = Number(p2);
      // Ưu tiên dd/mm/yyyy; tự đảo khi thấy tháng rõ ràng (> 12).
      const dayFirst = namiSuperTaoNgayKey(yyyy, n2, n1);
      const monthFirst = namiSuperTaoNgayKey(yyyy, n1, n2);
      if (n1 > 12 && n2 <= 12) result = dayFirst;
      else if (n2 > 12 && n1 <= 12) result = monthFirst;
      else result = dayFirst || monthFirst;
    }
  } else {
    const parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
      result = namiSuperTaoNgayKey(
        parsed.getUTCFullYear(),
        parsed.getUTCMonth() + 1,
        parsed.getUTCDate(),
      );
    }
  }

  namiSuperTextCache.set(text, result);
  return result;
};

const namiSuperLayNgayKey = (
  text: NamiSuperO,
  value: NamiSuperO,
): string | null => {
  if (typeof value === "number") {
    const fromSerial = namiSuperSerialSangNgayKey(value);
    if (fromSerial) return fromSerial;
  }
  return namiSuperTextSangNgayKey(value) ?? namiSuperTextSangNgayKey(text);
};

/** "A05  B44" → "A5 B44" (giống Auto_Nami). */
const namiSuperChuanHoaViTri = (giaTri: NamiSuperO): string => {
  const raw = String(giaTri ?? "")
    .toUpperCase()
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const cum: string[] = [];
  const regex = /([A-Z]+)\s*0*(\d+)/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(raw)) !== null) {
    cum.push(`${match[1]}${Number(match[2])}`);
  }
  return cum.length > 0 ? cum.join(" ") : raw;
};

const namiSuperDinhDangNgay = (key: string): string => {
  const [y, m, d] = key.split("-");
  return `${d}/${m}/${y}`;
};

const namiSuperNgayKeyUtc = (dt: Date): string =>
  namiSuperTaoNgayKey(
    dt.getUTCFullYear(),
    dt.getUTCMonth() + 1,
    dt.getUTCDate(),
  ) ?? "";

/** Kỳ tải lên; null khi ngày hoặc chế độ không hợp lệ. */
const namiSuperTinhKy = (
  ngay: string,
  cheDo: string,
  now: Date,
  lechMuiGio: number,
): NamiSuperKy | null => {
  const ngayChon = String(ngay ?? "").trim();
  const key = ngayChon
    ? namiSuperTextSangNgayKey(ngayChon)
    : namiSuperNgayKeyUtc(new Date(now.getTime() + lechMuiGio * 3600000));
  if (!key) return null;

  const [y, m, d] = key.split("-").map((phan) => Number(phan));
  if (cheDo === "day") {
    return { mode: "day", from: key, to: key, nhan: `ngày ${namiSuperDinhDangNgay(key)}` };
  }
  if (cheDo === "week") {
    const goc = new Date(Date.UTC(y, m - 1, d));
    const thu = goc.getUTCDay();
    const batDau = new Date(goc.getTime() + (thu === 0 ? -6 : 1 - thu) * 86400000);
    const ketThuc = new Date(batDau.getTime() + 6 * 86400000);
    const from = namiSuperNgayKeyUtc(batDau);
    const to = namiSuperNgayKeyUtc(ketThuc);
    return {
      mode: "week",
      from,
      to,
      nhan: `tuần ${namiSuperDinhDangNgay(from)} - ${namiSuperDinhDangNgay(to)}`,
    };
  }
  if (cheDo === "month") {
    const from = namiSuperNgayKeyUtc(new Date(Date.UTC(y, m - 1, 1)));
    const to = namiSuperNgayKeyUtc(new Date(Date.UTC(y, m, 0)));
    return {
      mode: "month",
      from,
      to,
      nhan: `tháng ${String(m).padStart(2, "0")}/${y}`,
    };
  }
  return null;
};

const namiSuperLayO = (
  rowValues: NamiSuperO[],
  rowTexts: NamiSuperO[],
  idx: number,
): NamiSuperO => {
  const value = rowValues[idx];
  if (value !== undefined && value !== null && value !== "") return value;
  return rowTexts[idx] ?? "";
};

/**
 * Lấy các lần nhập kho thuộc kỳ từ các dòng A..O của "Thành phẩm".
 * rowOffset = số dòng sheet nằm trên dòng đầu tiên của mảng (đọc từ A2 thì rowOffset = 1).
 */
const namiSuperDocLanNhap = (
  values: NamiSuperO[][],
  texts: NamiSuperO[][],
  rowOffset: number,
  ky: { from: string; to: string },
): {
  lanNhap: NamiSuperLanNhap[];
  boQua: NamiSuperBoQua[];
  soDongCoNoiDung: number;
} => {
  const lanNhap: NamiSuperLanNhap[] = [];
  const boQua: NamiSuperBoQua[] = [];
  let soDongCoNoiDung = 0;

  for (let i = 0; i < values.length; i++) {
    const rowValues = values[i] ?? [];
    const rowTexts = texts[i] ?? [];
    const sku = namiSuperLamSach(
      namiSuperLayO(rowValues, rowTexts, NAMI_SUPER_COT.sku),
    ).toUpperCase();
    const hopDong = String(
      namiSuperLayO(rowValues, rowTexts, NAMI_SUPER_COT.hopDong),
    ).trim();
    const ten = String(
      namiSuperLayO(rowValues, rowTexts, NAMI_SUPER_COT.ten),
    ).trim();
    if (!sku && !hopDong && !ten) continue;
    soDongCoNoiDung++;

    const sourceRow = rowOffset + i + 1;
    let tongDaNhap = 0;
    for (let lan = 0; lan < 3; lan++) {
      tongDaNhap += namiSuperParseSo(
        namiSuperLayO(rowValues, rowTexts, NAMI_SUPER_COT.ngayNhap1 + lan * 2 + 1),
      );
    }
    const totalReceived = Math.round(tongDaNhap);

    for (let lan = 0; lan < 3; lan++) {
      const cotNgay = NAMI_SUPER_COT.ngayNhap1 + lan * 2;
      const ngay = namiSuperLayNgayKey(rowTexts[cotNgay] ?? "", rowValues[cotNgay] ?? "");
      if (!ngay || ngay < ky.from || ngay > ky.to) continue;

      const roundNumber = lan + 1;
      const boQuaVi = (reason: string) =>
        boQua.push({
          source_row: sourceRow,
          round_number: roundNumber,
          sku,
          contract_no: hopDong,
          reason,
        });

      if (!sku) {
        boQuaVi("Thiếu SKU (cột C)");
        continue;
      }
      if (!hopDong) {
        boQuaVi("Thiếu mã hợp đồng (cột H)");
        continue;
      }
      const quantity = Math.round(
        namiSuperParseSo(namiSuperLayO(rowValues, rowTexts, cotNgay + 1)),
      );
      if (quantity <= 0) {
        boQuaVi(`Ô SL nhập kho ${roundNumber} trống, bằng 0 hoặc không phải số`);
        continue;
      }
      if (totalReceived < quantity) {
        boQuaVi(
          `Tổng đã nhập K + M + O (${totalReceived}) nhỏ hơn SL nhập kho ${roundNumber} (${quantity})`,
        );
        continue;
      }

      lanNhap.push({
        source_row: sourceRow,
        round_number: roundNumber,
        receipt_date: ngay,
        contract_no: hopDong,
        sku,
        product_name: ten,
        order_quantity: Math.round(
          namiSuperParseSo(
            namiSuperLayO(rowValues, rowTexts, NAMI_SUPER_COT.slDonHang),
          ),
        ),
        total_received: totalReceived,
        quantity,
        location_code: namiSuperChuanHoaViTri(
          namiSuperLayO(rowValues, rowTexts, NAMI_SUPER_COT.viTri),
        ),
      });
    }
  }

  return { lanNhap, boQua, soDongCoNoiDung };
};

const namiSuperChuanHoaTen = (giaTri: NamiSuperO): string =>
  String(giaTri ?? "")
    .toLowerCase()
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const namiSuperThongSoCoDuLieu = (t: NamiSuperThongSo): boolean =>
  !!(t.trongLuong || t.dai || t.rong || t.cao || t.theTich);

/**
 * Tra cứu thông số từ sheet "2026" (các dòng B..H): mã → tên chính xác → tên gần đúng.
 * Giống Auto_Nami: chiều cao trống mà có dài × rộng thì dùng 1.1; thể tích = D×R×C/1e6.
 */
const namiSuperTaoTraCuuThongSo = (
  rows2026: NamiSuperO[][],
  cauHinhTenMo: NamiSuperTraCuuTenMo,
): ((sku: string, ten: string) => NamiSuperKetQuaTraCuu | undefined) => {
  const theoMa = new Map<string, NamiSuperThongSo>();
  const theoTen = new Map<string, NamiSuperThongSo>();
  const danhSachTen: {
    ten: string;
    tokens: string[];
    thongSo: NamiSuperThongSo;
    coDuLieu: boolean;
  }[] = [];

  for (const row of rows2026) {
    const ma = namiSuperLamSach(row[0]).toUpperCase();
    const ten = namiSuperChuanHoaTen(row[1]);
    if (!ma && !ten) continue;

    const dai = namiSuperParseSo(row[4]);
    const rong = namiSuperParseSo(row[5]);
    const caoGoc = namiSuperParseSo(row[6]);
    const cao = caoGoc > 0 ? caoGoc : dai > 0 && rong > 0 ? 1.1 : 0;
    const trongLuong = namiSuperParseSo(row[3]);
    const theTich = (dai * rong * cao) / 1000000;
    const thongSo: NamiSuperThongSo = {
      trongLuong: trongLuong ? Number(trongLuong.toFixed(4)) : 0,
      dai: dai ? Number(dai.toFixed(4)) : 0,
      rong: rong ? Number(rong.toFixed(4)) : 0,
      cao: cao ? Number(cao.toFixed(4)) : 0,
      theTich: theTich ? Number(theTich.toFixed(3)) : 0,
    };

    if (ma) theoMa.set(ma, thongSo);
    if (ten) {
      const coDuLieu = namiSuperThongSoCoDuLieu(thongSo);
      const cu = theoTen.get(ten);
      if (!cu || !namiSuperThongSoCoDuLieu(cu) || coDuLieu) {
        theoTen.set(ten, thongSo);
        const entry = { ten, tokens: ten.split(" "), thongSo, coDuLieu };
        const idx = danhSachTen.findIndex((x) => x.ten === ten);
        if (idx >= 0) danhSachTen[idx] = entry;
        else danhSachTen.push(entry);
      }
    }
  }

  const timTheoTenMo = (ten: string): NamiSuperThongSo | undefined => {
    const query = ten.split(" ").filter((t) => t.length > 0);
    if (query.length === 0) return undefined;
    const queryUnique = Array.from(new Set(query));
    const nguongGiao = Math.max(
      cauHinhTenMo.soTokenKhopToiDa,
      Math.ceil(query.length * cauHinhTenMo.tyLeTokenKhopToiThieu),
    );

    let best: NamiSuperThongSo | undefined;
    let bestScore = 0;
    let bestGiao = 0;
    let bestCoDuLieu = false;
    for (const item of danhSachTen) {
      const setC = new Set(item.tokens);
      let giao = 0;
      for (const token of query) if (setC.has(token)) giao++;
      if (giao < nguongGiao) continue;

      let giaoUnique = 0;
      for (const token of queryUnique) if (setC.has(token)) giaoUnique++;
      const hop = new Set([...queryUnique, ...item.tokens]).size;
      const score = hop > 0 ? giaoUnique / hop : 0;
      if (score < cauHinhTenMo.diemToiThieu) continue;

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

  return (sku: string, ten: string) => {
    const ma = namiSuperLamSach(sku).toUpperCase();
    const theoMaKq = ma ? theoMa.get(ma) : undefined;
    if (theoMaKq) return { thongSo: theoMaKq, nguon: "ma" };

    const tenChuan = namiSuperChuanHoaTen(ten);
    if (!tenChuan) return undefined;
    const chinhXac = theoTen.get(tenChuan);
    if (chinhXac) return { thongSo: chinhXac, nguon: "ten" };

    if (!cauHinhTenMo.bat || danhSachTen.length === 0) return undefined;
    const mo = timTheoTenMo(tenChuan);
    return mo ? { thongSo: mo, nguon: "ten-mo" } : undefined;
  };
};

/** Chia đợt gửi; nhóm ngày|SKU|HĐ không bị tách và giữ nguyên thứ tự trong nhóm. */
const namiSuperChiaDot = <
  T extends { receipt_date: string; sku: string; contract_no: string },
>(
  lanNhap: T[],
  kichThuoc: number,
): T[][] => {
  const nhom = new Map<string, T[]>();
  for (const item of lanNhap) {
    const key = `${item.receipt_date}|${item.sku}|${item.contract_no}`;
    const ds = nhom.get(key);
    if (ds) ds.push(item);
    else nhom.set(key, [item]);
  }

  const dots: T[][] = [];
  let hienTai: T[] = [];
  nhom.forEach((ds) => {
    if (hienTai.length > 0 && hienTai.length + ds.length > kichThuoc) {
      dots.push(hienTai);
      hienTai = [];
    }
    hienTai = hienTai.concat(ds);
  });
  if (hienTai.length > 0) dots.push(hienTai);
  return dots;
};

const namiSuperGopKetQua = (ketQuaCacDot: NamiSuperKetQuaDot[]): NamiSuperKetQuaDot => {
  const tong: NamiSuperKetQuaDot = {
    imported: 0,
    unchanged: 0,
    rejected: [],
    receipts: [],
    specs_from_sheet: [],
    missing_specs: [],
  };
  for (const kq of ketQuaCacDot) {
    tong.imported += kq.imported ?? 0;
    tong.unchanged += kq.unchanged ?? 0;
    tong.rejected.push(...(kq.rejected ?? []));
    tong.specs_from_sheet.push(...(kq.specs_from_sheet ?? []));
    for (const r of kq.receipts ?? []) {
      if (!tong.receipts.includes(r)) tong.receipts.push(r);
    }
    for (const s of kq.missing_specs ?? []) {
      if (!tong.missing_specs.includes(s)) tong.missing_specs.push(s);
    }
  }
  return tong;
};

const NAMI_SUPER_SHEET_TRANG_THAI = "_UploadNamiSuper";
const NAMI_SUPER_TIEU_DE_TRANG_THAI = [
  "Địa chỉ",
  "Màu cũ",
  "Màu đánh dấu",
  "Thời điểm",
  "Lý do",
];

/** Cặp ô ngày + SL của một lần nhập: J:K / L:M / N:O. */
const namiSuperDiaChiCapO = (row: number, roundNumber: number): string => {
  const [cotNgay, cotSl] = [
    ["J", "K"],
    ["L", "M"],
    ["N", "O"],
  ][Math.min(3, Math.max(1, roundNumber)) - 1];
  return `${cotNgay}${row}:${cotSl}${row}`;
};

/** Gộp lần nhập bị script bỏ qua và bị backend từ chối; cùng cặp ô thì nối lý do. */
const namiSuperLapDanhSachDanhDau = (
  boQua: NamiSuperBoQua[],
  rejected: NamiSuperTuChoi[],
): { diaChi: string; lyDo: string }[] => {
  const ds: { diaChi: string; lyDo: string }[] = [];
  const them = (row: number, roundNumber: number, lyDo: string) => {
    if (!(row >= 1)) return;
    const diaChi = namiSuperDiaChiCapO(row, roundNumber);
    const daCo = ds.find((x) => x.diaChi === diaChi);
    if (!daCo) ds.push({ diaChi, lyDo });
    else if (!daCo.lyDo.split("; ").includes(lyDo)) daCo.lyDo += `; ${lyDo}`;
  };
  boQua.forEach((r) => them(r.source_row, r.round_number, r.reason));
  rejected.forEach((r) => them(r.row, r.round_number ?? 1, r.reason));
  return ds;
};

const namiSuperChuanHoaMau = (mau: string): string =>
  String(mau ?? "")
    .trim()
    .replace(/^#/, "")
    .toUpperCase();

/** Chỉ khôi phục ô vẫn còn màu script đã tô (người dùng tô màu khác thì giữ nguyên). */
const namiSuperNenKhoiPhuc = (mauHienTai: string, mauDanhDau: string): boolean => {
  const hienTai = namiSuperChuanHoaMau(mauHienTai);
  return hienTai !== "" && hienTai === namiSuperChuanHoaMau(mauDanhDau);
};

const namiSuperDocTrangThai = (
  rows: NamiSuperO[][],
): { diaChi: string; mauCu: string; mauDanhDau: string }[] =>
  rows
    .map((row) => ({
      diaChi: String(row[0] ?? "").trim(),
      mauCu: String(row[1] ?? "").trim(),
      mauDanhDau: String(row[2] ?? "").trim(),
    }))
    .filter(
      (x) => x.diaChi !== "" && x.diaChi !== NAMI_SUPER_TIEU_DE_TRANG_THAI[0],
    );

const namiSuperDatMau = (range: ExcelScript.Range, mau: string) => {
  // Excel trả #FFFFFF cho ô không tô màu; clear() để giữ đường lưới.
  const chuan = namiSuperChuanHoaMau(mau);
  if (chuan === "" || chuan === "FFFFFF") range.getFormat().getFill().clear();
  else range.getFormat().getFill().setColor(mau);
};

const namiSuperKiemTraApiUrl = (url: string): string | null => {
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

/** Bỏ các khoá rỗng để payload gọn. */
const namiSuperGonItem = (item: NamiSuperLanNhap): { [key: string]: unknown } => {
  const out: { [key: string]: unknown } = {};
  Object.keys(item).forEach((key) => {
    const val = (item as { [key: string]: unknown })[key];
    if (val !== null && val !== undefined && val !== "") out[key] = val;
  });
  return out;
};

const namiSuperGiay = (batDau: number): number =>
  Math.round((Date.now() - batDau) / 1000);

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
  text(): Promise<string>;
}>;

type NamiSuperPhanHoi = Partial<NamiSuperKetQuaDot> & {
  message?: string;
  errors?: { [key: string]: string[] };
};

const namiSuperDocJson = (text: string): NamiSuperPhanHoi => {
  try {
    return JSON.parse(text) as NamiSuperPhanHoi;
  } catch {
    return { message: text };
  }
};

/** Gửi một đợt; trả về kết quả hoặc các dòng lỗi để in nhật ký. */
const namiSuperGuiDot = async (
  apiUrl: string,
  apiToken: string,
  payload: object,
): Promise<{ ketQua?: NamiSuperKetQuaDot; loi?: string[] }> => {
  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Authorization: `Bearer ${apiToken}`,
      },
      body: JSON.stringify(payload),
    });
    const bodyText = await response.text();
    const body = namiSuperDocJson(bodyText);

    if (!response.ok) {
      const loi = [`LỖI HTTP ${response.status}: ${body.message ?? bodyText}`];
      if (response.status === 409) {
        loi.push("Có lần nhập kho khác đang chạy — đợi 1–2 phút rồi chạy lại.");
      }
      const errors = body.errors ?? {};
      Object.keys(errors)
        .slice(0, 5)
        .forEach((key) => loi.push(`- ${key}: ${errors[key].join("; ")}`));
      return { loi };
    }

    return {
      ketQua: {
        imported: body.imported ?? 0,
        unchanged: body.unchanged ?? 0,
        rejected: body.rejected ?? [],
        receipts: body.receipts ?? [],
        specs_from_sheet: body.specs_from_sheet ?? [],
        missing_specs: body.missing_specs ?? [],
      },
    };
  } catch (err) {
    const msg = String(err);
    return {
      loi: [
        msg.includes("Failed to fetch")
          ? `LỖI kết nối API: Failed to fetch — kiểm tra apiUrl (${apiUrl}), server và CORS.`
          : `LỖI kết nối API: ${msg}`,
      ],
    };
  }
};

// @ts-ignore: Office Scripts
async function main(workbook: ExcelScript.Workbook) {
  const CFG = UPLOAD_NAMI_SUPER_CONFIG;
  const batDau = Date.now();
  namiSuperSerialCache.clear();
  namiSuperTextCache.clear();

  // 1. Kiểm tra cấu hình
  const loiUrl = namiSuperKiemTraApiUrl(CFG.apiUrl);
  if (loiUrl) {
    console.log(`LỖI cấu hình: ${loiUrl}`);
    return;
  }
  if (!CFG.apiToken) {
    console.log("LỖI: Chưa cấu hình apiToken (Bearer token từ Dashboard → API Token).");
    return;
  }
  const sheetNguon = workbook.getWorksheet(CFG.tenSheetNguon);
  if (!sheetNguon) {
    console.log(`LỖI: Không tìm thấy sheet "${CFG.tenSheetNguon}".`);
    return;
  }

  // 2. Kỳ tải lên
  const ky = namiSuperTinhKy(CFG.ngay, CFG.cheDo, new Date(), CFG.lechMuiGio);
  if (!ky) {
    console.log(
      `LỖI cấu hình: ngay "${CFG.ngay}" hoặc cheDo "${CFG.cheDo}" không hợp lệ (ngay dd/mm/yyyy; cheDo day | week | month).`,
    );
    return;
  }

  // 3. Đọc "Thành phẩm" A2:O<dòng cuối cột H>
  const dongCuoi =
    sheetNguon
      .getRange("H1048576")
      .getRangeEdge(ExcelScript.KeyboardDirection.up)
      .getRowIndex() + 1;
  let values: NamiSuperO[][] = [];
  let texts: NamiSuperO[][] = [];
  if (dongCuoi >= 2) {
    const vung = sheetNguon.getRange(`A2:O${dongCuoi}`);
    values = vung.getValues() as NamiSuperO[][];
    texts = vung.getTexts();
  }
  const { lanNhap, boQua, soDongCoNoiDung } = namiSuperDocLanNhap(values, texts, 1, ky);
  console.log(
    `Kỳ: ${ky.nhan} (${ky.from} → ${ky.to}). Đã quét ${soDongCoNoiDung} dòng "${CFG.tenSheetNguon}", ${lanNhap.length + boQua.length} lần nhập thuộc kỳ.`,
  );

  // 4. Thông số từ sheet "2026" (chỉ dùng khi hệ thống còn thiếu)
  if (CFG.traThongSoSheet && lanNhap.length > 0) {
    const sheetThongSo = workbook.getWorksheet(CFG.tenSheetThongSo);
    if (!sheetThongSo) {
      console.log(
        `Lưu ý: Không tìm thấy sheet "${CFG.tenSheetThongSo}", chỉ dùng thông số có sẵn trên hệ thống.`,
      );
    } else {
      const dongCuoiThongSo =
        sheetThongSo
          .getRange("B1048576")
          .getRangeEdge(ExcelScript.KeyboardDirection.up)
          .getRowIndex() + 1;
      if (dongCuoiThongSo >= 2) {
        const traCuu = namiSuperTaoTraCuuThongSo(
          sheetThongSo.getRange(`B2:H${dongCuoiThongSo}`).getValues() as NamiSuperO[][],
          CFG.traCuuTenMo,
        );
        const daTra = new Map<string, NamiSuperKetQuaTraCuu | undefined>();
        for (const item of lanNhap) {
          if (!daTra.has(item.sku)) daTra.set(item.sku, traCuu(item.sku, item.product_name));
          const kq = daTra.get(item.sku);
          if (!kq || !namiSuperThongSoCoDuLieu(kq.thongSo)) continue;
          item.unit_weight = kq.thongSo.trongLuong;
          item.length = kq.thongSo.dai;
          item.width = kq.thongSo.rong;
          item.height = kq.thongSo.cao;
          item.unit_volume = kq.thongSo.theTich;
          item.spec_source = kq.nguon;
        }
      }
    }
  }

  // 5. Khôi phục các ô đã tô ở lần chạy trước
  let sheetTrangThai = workbook.getWorksheet(NAMI_SUPER_SHEET_TRANG_THAI);
  if (!sheetTrangThai) {
    sheetTrangThai = workbook.addWorksheet(NAMI_SUPER_SHEET_TRANG_THAI);
    sheetTrangThai.setVisibility(ExcelScript.SheetVisibility.hidden);
  }
  const vungTrangThai = sheetTrangThai.getUsedRange();
  if (vungTrangThai) {
    let soOKhoiPhuc = 0;
    namiSuperDocTrangThai(vungTrangThai.getValues() as NamiSuperO[][]).forEach((x) => {
      const oDau = sheetNguon.getRange(x.diaChi.split(":")[0]);
      if (!namiSuperNenKhoiPhuc(oDau.getFormat().getFill().getColor() ?? "", x.mauDanhDau)) return;
      namiSuperDatMau(sheetNguon.getRange(x.diaChi), x.mauCu);
      soOKhoiPhuc++;
    });
    vungTrangThai.clear();
    if (soOKhoiPhuc > 0) {
      console.log(`Đã bỏ đánh dấu ${soOKhoiPhuc} cặp ô của lần chạy trước.`);
    }
  }

  if (lanNhap.length === 0) {
    console.log(
      `Không có dữ liệu nhập kho cho ${ky.nhan}. Đã quét ${soDongCoNoiDung} dòng.`,
    );
  }

  // 7. Gửi từng đợt
  const dots = namiSuperChiaDot(lanNhap, Math.max(1, CFG.kichThuocDot));
  const ketQuaCacDot: NamiSuperKetQuaDot[] = [];
  const nhatKyDot: string[] = [];
  for (let i = 0; i < dots.length; i++) {
    const dot = dots[i];
    const batDauDot = Date.now();
    const { ketQua, loi } = await namiSuperGuiDot(CFG.apiUrl, CFG.apiToken, {
      source_sheet: CFG.tenSheetNguon,
      period: { mode: ky.mode, from: ky.from, to: ky.to },
      batch: { index: i + 1, total: dots.length },
      items: dot.map((item) => namiSuperGonItem(item)),
    });

    if (!ketQua) {
      nhatKyDot.push(
        ...(loi ?? []),
        `Đã ghi ${i}/${dots.length} đợt; chạy lại an toàn, không bị trùng.`,
      );
      break;
    }

    ketQuaCacDot.push(ketQua);
    nhatKyDot.push(
      `Đợt ${i + 1}/${dots.length}: ${dot.length} lần nhập… xong (${namiSuperGiay(batDauDot)}s).`,
    );
  }
  if (nhatKyDot.length > 0) console.log(nhatKyDot.join("\n"));

  // 9. Tóm tắt
  const tong = namiSuperGopKetQua(ketQuaCacDot);
  if (ketQuaCacDot.length > 0) {
    if (tong.specs_from_sheet.length > 0) {
      const tenMo = tong.specs_from_sheet
        .filter((s) => s.source === "ten-mo")
        .map((s) => s.sku);
      console.log(
        `Thông số từ sheet "${CFG.tenSheetThongSo}": ${tong.specs_from_sheet.length} SKU` +
          (tenMo.length > 0
            ? ` (${tenMo.length} SKU theo tên gần đúng: ${tenMo.join(", ")} — kiểm tra lại).`
            : "."),
      );
    }
    console.log(
      `Kết quả: ${tong.imported} đã nhập, ${tong.unchanged} đã có từ trước, ${tong.rejected.length} bị từ chối, ${boQua.length} bị bỏ qua.` +
        (tong.receipts.length > 0 ? ` Phiếu: ${tong.receipts.join(", ")}.` : ""),
    );
  }

  const canhBao: string[] = [];
  boQua.forEach((r) =>
    canhBao.push(
      `- Dòng ${r.source_row} – Nhập kho ${r.round_number}: SKU ${r.sku || "(trống)"} / HĐ ${r.contract_no || "(trống)"} — ${r.reason}`,
    ),
  );
  tong.rejected.forEach((r) =>
    canhBao.push(
      `- Dòng ${r.row} – Nhập kho ${r.round_number}: SKU ${r.sku} / HĐ ${r.contract_no} — ${r.reason}`,
    ),
  );
  // 8. Tô màu cặp ô ngày + SL và ghi trạng thái
  const danhDau = namiSuperLapDanhSachDanhDau(boQua, tong.rejected);
  const thoiDiem = new Date().toISOString();
  const dongTrangThai: NamiSuperO[][] = [NAMI_SUPER_TIEU_DE_TRANG_THAI];
  danhDau.forEach((x) => {
    const mauCu = sheetNguon.getRange(x.diaChi.split(":")[0]).getFormat().getFill().getColor() ?? "";
    sheetNguon.getRange(x.diaChi).getFormat().getFill().setColor(CFG.mauDanhDau);
    dongTrangThai.push([x.diaChi, mauCu, CFG.mauDanhDau, thoiDiem, x.lyDo]);
  });
  sheetTrangThai.getRange(`A1:E${dongTrangThai.length}`).setValues(dongTrangThai);

  if (canhBao.length > 0) {
    console.log(
      [
        `CẢNH BÁO: ${canhBao.length} lần nhập bị bỏ qua/từ chối — đã tô màu ô ngày + SL trên "${CFG.tenSheetNguon}":`,
        ...canhBao,
      ].join("\n"),
    );
  }
  if (tong.missing_specs.length > 0) {
    console.log(
      `CẢNH BÁO: ${tong.missing_specs.length} SKU chưa có thông số trong hệ thống lẫn sheet "${CFG.tenSheetThongSo}": ${tong.missing_specs.join(", ")}.`,
    );
  }

  console.log(
    `Hoàn tất trong ${namiSuperGiay(batDau)}s. Báo cáo kỳ này: chạy Auto_Nami_Supper với cùng ngày/chế độ.`,
  );
}

const namiSuperTestables = {
  namiSuperTinhKy,
  namiSuperLayNgayKey,
  namiSuperChuanHoaViTri,
  namiSuperDocLanNhap,
  namiSuperTaoTraCuuThongSo,
  namiSuperChiaDot,
  namiSuperGopKetQua,
  namiSuperKiemTraApiUrl,
  namiSuperDiaChiCapO,
  namiSuperLapDanhSachDanhDau,
  namiSuperNenKhoiPhuc,
  namiSuperDocTrangThai,
};

(
  globalThis as typeof globalThis & {
    __UPLOAD_NAMI_SUPER_TESTABLES__?: typeof namiSuperTestables;
  }
).__UPLOAD_NAMI_SUPER_TESTABLES__ = namiSuperTestables;
