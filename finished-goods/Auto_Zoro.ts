/**
 * ===================================================================
 * Auto_Zoro – Báo cáo tồn kho theo thị trường (tách sheet đích).
 * ===================================================================
 *
 * Đọc sheet tổng hợp kho (cùng mapping cột với Auto_Mitsuha.cot.nguon),
 * tính tồn kho = Tổng nhập − Tổng xuất (các cột xuất S/U/W),
 * chỉ giữ dòng có tồn kho > 0,
 * lọc theo ngày nhập kho (cột J/L/N như Mitsuha) khi cấu hình `locNgayNhapChiTiet`,
 * lọc theo tiền tố 2 ký tự đầu mã hợp đồng, sắp xếp theo vị trí kho,
 * ghi ra sheet đích tên `{thị trường}-Tồn kho-{ngày chạy}`.
 *
 * Hàng 1: tiêu đề. Cột A (STT): có thể dùng công thức =ROW()-1 (mặc định bật).
 * Cột G–J: Trọng lượng, Tổng trọng lượng, Thể tích, Tổng thể tích (lấy từ sheet thông số).
 */

// ==========================================
// 0. KIỂU DỮ LIỆU
// ==========================================
interface ZoroProductInfo {
  trongLuong: number;
  dai: number;
  rong: number;
  cao: number;
  theTich: number;
}

type ZoroThongSoTheoTenEntry = {
  tenChuan: string;
  tokens: string[];
  thongSo: ZoroProductInfo;
  coDuLieu: boolean;
};

type ZoroTraCuuTenMoConfig = {
  bat: boolean;
  diemToiThieu: number;
  tyLeTokenKhopToiThieu: number;
  soTokenKhopToiDa: number;
};

type ZoroThongSoConfig = {
  bat: boolean;
};

// ==========================================
// 1. CẤU HÌNH
// ==========================================
const ZORO_CONFIG = {
  /** Xóa nội dung cũ trên các sheet đích trước khi ghi (giữ tên sheet). */
  xoaDuLieuCu: true,
  /** Tạo sheet đích nếu chưa tồn tại. */
  taoSheetNeuChuaCo: true,
  /** Cột STT (A) dùng công thức =ROW()-1 từ dòng dữ liệu đầu tiên. */
  sttBangCongThucRow: true,
  /**
   * Chuỗi giữa mã thị trường và ngày trong tên sheet (vd. `EU-Tồn kho-2026-05-13`).
   */
  tienToTrongTenSheet: "Tồn kho",
  /**
   * Lọc theo ngày nhập kho chi tiết (bất kỳ lần nhập 1/2/3 nào khớp là giữ dòng).
   * Quy tắc giống Auto_Mitsuha: mảng nào để trống [] thì bỏ qua tiêu chí đó.
   * Khi cả ba mảng `ngay`, `thang`, `nam` đều rỗng → tắt lọc theo ngày (lấy mọi ngày nhập).
   * Cột ngày: `cot.nguon.cotsNgayNhap` (mặc định J, L, N).
   */
  locNgayNhapChiTiet: {
    ngay: [] as number[],
    thang: [] as number[],
    nam: [] as number[],
  },
  /**
   * Mã thị trường (2 ký tự đầu mã hợp đồng). Tên sheet đích = `{mã}-{tienToTrongTenSheet}-{ngày chạy}`.
   */
  thiTruong: ["EU"],
  /**
   * Mã vị trí cần lấy toàn bộ tồn (vd ["E23","E24"]), bất kể mã hợp đồng. Để trống = tắt.
   * Khi bật: thêm sheet `Vị trí-{tienToTrongTenSheet}-{ngày chạy}`; các sheet thị trường giữ nguyên.
   */
  viTri: [] as string[],
  /** Mã SKU cần dò (vd "936771"): log lý do giữ/bỏ từng dòng của SKU này trong sheet nguồn. Để trống = tắt. */
  debugSku: "",
  cot: {
    nguon: {
      tenSheet: "Thành phẩm",
      dongBatDau: 2,
      maHang: "C",
      slDonHang: "F",
      hopDong: "H",
      viTri: "I",
      tongNhap: "Q",
      tenHang: "D",
      cotsXuatKho: ["S", "U", "W"],
      cotsNhapKho: ["K", "M", "O"],
      cotsNgayNhap: ["J", "L", "N"],
    },
  },

  /**
   * Sheet chứa thông số sản phẩm (trọng lượng, kích thước) — giống Auto_Nami.
   * Cấu trúc: B=mã hàng, C=tên hàng, D=(bỏ qua), E=trọng lượng, F=dài, G=rộng, H=cao.
   */
  tenSheetThongSo: "2026",
  thongSo: {
    bat: true,
  } as ZoroThongSoConfig,
  traCuuTenMo: {
    bat: true,
    diemToiThieu: 0.55,
    tyLeTokenKhopToiThieu: 0.5,
    soTokenKhopToiDa: 3,
  } as ZoroTraCuuTenMoConfig,
};

const ZORO_TIEU_DE_HANG_1: (string | number | boolean)[] = [
  "STT",
  "Mã hợp đồng",
  "SKU",
  "Tên hàng",
  "Số lượng đơn hàng",
  "Tồn kho",
  "Trọng lượng",
  "Tổng trọng lượng",
  "Thể tích",
  "Tổng thể tích",
  "Vị trí",
  "Ghi chú",
];

/** Chữ cái cột "Vị trí" (để wrap text). */
const ZORO_COT_VI_TRI = "K";
/** Chữ cái cột "Tên hàng" — nội dung dữ liệu căn trái. */
const ZORO_COT_TEN_HANG = "D";
/** Cột cuối cùng của bảng. */
const ZORO_COT_CUOI = "L";
/** Chiều cao tối thiểu mỗi hàng (point); sau autofit, hàng nào thấp hơn sẽ được nâng lên mức này. */
const ZORO_CHIEU_CAO_DONG = 23;

// @ts-ignore: Office Scripts — tránh trùng tên khi gộp nhiều file TS
function zoroLamSach(giaTri: string | number | boolean): string {
  return String(giaTri)
    .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "")
    .trim();
}

// @ts-ignore: Office Scripts
function zoroParseSoTuExcel(giaTri: string | number | boolean): number {
  if (typeof giaTri === "number" && !isNaN(giaTri)) return giaTri;
  let s = zoroLamSach(giaTri);
  if (!s) return 0;
  let lastComma = s.lastIndexOf(",");
  let lastDot = s.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    if (lastComma > lastDot) {
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      s = s.replace(/,/g, "");
    }
  } else if (lastComma >= 0) {
    s = s.replace(",", ".");
  }
  let n = Number(s);
  return isNaN(n) ? 0 : n;
}

// @ts-ignore: Office Scripts — đồng bộ hiển thị vị trí với Auto_Mitsuha
function zoroChuanHoaViTri(giaTri: string | number | boolean): string {
  let raw = String(giaTri ?? "").toUpperCase();
  raw = raw
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  let cum: { khu: string; so: number }[] = [];
  let regex = /([A-Z]+)\s*0*(\d+)/g;
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
}

/** Các phân đoạn vị trí hợp lệ (vd "e023, E24" → ["E23", "E24"]). */
function zoroPhanDoanViTri(giaTri: string | number | boolean): string[] {
  return zoroChuanHoaViTri(giaTri)
    .split(" ")
    .filter((doan) => /^[A-Z]+\d+$/.test(doan));
}

/** Chuẩn hóa cấu hình `viTri`: gộp trùng giữ thứ tự; mục không có phân đoạn hợp lệ vào `ignored`. */
function zoroChuanHoaDanhSachViTri(danhSach: string[]): {
  codes: string[];
  ignored: string[];
} {
  let codes: string[] = [];
  let ignored: string[] = [];
  for (let muc of danhSach) {
    let giaTri = String(muc ?? "").trim();
    if (!giaTri) continue;
    let doans = zoroPhanDoanViTri(giaTri);
    if (doans.length === 0) ignored.push(giaTri);
    for (let doan of doans) {
      if (!codes.includes(doan)) codes.push(doan);
    }
  }
  return { codes, ignored };
}

function zoroKhopViTri(
  oViTri: string | number | boolean,
  codes: string[],
): boolean {
  if (codes.length === 0) return false;
  return zoroPhanDoanViTri(oViTri).some((doan) => codes.includes(doan));
}

/**
 * Sheet đích của một dòng nguồn: `thiTruong` = mã thị trường nếu nằm trong `thiTruongs` (rỗng nếu không),
 * `viTri` = ô vị trí khớp `codes` — không xét mã hợp đồng (kể cả ô hợp đồng trống).
 */
function zoroXacDinhDich(
  hopDong: string | number | boolean,
  oViTri: string | number | boolean,
  thiTruongs: string[],
  codes: string[],
): { thiTruong: string; viTri: boolean } {
  let hd = String(hopDong ?? "").trim();
  let maTT = hd.toUpperCase().slice(0, 2);
  return {
    thiTruong: hd !== "" && thiTruongs.includes(maTT) ? maTT : "",
    viTri: zoroKhopViTri(oViTri, codes),
  };
}

// @ts-ignore: Office Scripts — giống chuanHoaNgay trong Auto_Mitsuha
function zoroChuanHoaNgay(giaTri: string | number | boolean): string {
  let raw = giaTri;
  if (typeof raw === "number" && raw > 30000) {
    let jsDate = new Date(Math.round((raw - 25569) * 86400 * 1000));
    let dd = String(jsDate.getUTCDate()).padStart(2, "0");
    let mm = String(jsDate.getUTCMonth() + 1).padStart(2, "0");
    let yyyy = jsDate.getUTCFullYear();
    return dd + "/" + mm + "/" + yyyy;
  }
  return String(raw).trim();
}

/** Giống tachCumViTri trong Auto_Mitsuha — phục vụ sắp xếp theo lộ trình kho. */
function zoroTachCumViTri(viTri: string): { khu: string; so: number }[] {
  let ketQua: { khu: string; so: number }[] = [];
  let regex = /([A-Z]+)\s*0*(\d+)/g;
  let match: RegExpExecArray | null;
  let viTriSach = String(viTri)
    .toUpperCase()
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  while ((match = regex.exec(viTriSach)) !== null) {
    ketQua.push({
      khu: match[1],
      so: Number(match[2]),
    });
  }
  ketQua.sort((a, b) => {
    if (a.khu !== b.khu) return a.khu.localeCompare(b.khu);
    return a.so - b.so;
  });
  return ketQua;
}

/** So sánh hai dòng báo cáo Zoro: index 5 = vị trí (đã chuẩn hóa). */
function zoroSoSanhHangTheoViTri(
  a: (string | number | boolean)[],
  b: (string | number | boolean)[],
): number {
  let vtA = String(a[5] ?? "").trim();
  let vtB = String(b[5] ?? "").trim();
  let cumA = zoroTachCumViTri(vtA);
  let cumB = zoroTachCumViTri(vtB);

  if (cumA.length === 0 && cumB.length > 0) return 1;
  if (cumA.length > 0 && cumB.length === 0) return -1;
  if (cumA.length === 0 && cumB.length === 0) return vtA.localeCompare(vtB);

  let maxLen = Math.max(cumA.length, cumB.length);
  for (let i = 0; i < maxLen; i++) {
    let pa = cumA[i];
    let pb = cumB[i];
    if (!pa && pb) return -1;
    if (pa && !pb) return 1;
    if (!pa || !pb) continue;

    if (pa.khu !== pb.khu) return pa.khu.localeCompare(pb.khu);
    if (pa.so !== pb.so) return pa.so - pb.so;
  }

  let t = vtA.localeCompare(vtB);
  if (t !== 0) return t;
  let h = String(a[0] ?? "").localeCompare(String(b[0] ?? ""));
  if (h !== 0) return h;
  return String(a[1] ?? "").localeCompare(String(b[1] ?? ""));
}

// @ts-ignore: Office Scripts
function zoroTimDongCuoiSheetThongSo(sheet: ExcelScript.Worksheet): number {
  return (
    sheet
      .getRange("B1048576")
      .getRangeEdge(ExcelScript.KeyboardDirection.up)
      .getRowIndex() + 1
  );
}

// @ts-ignore: Office Scripts
function zoroChuanHoaTenSanPham(giaTri: string | number | boolean): string {
  return String(giaTri ?? "")
    .toLowerCase()
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// @ts-ignore: Office Scripts
function zoroTachTokenTenSanPham(tenChuan: string): string[] {
  if (!tenChuan) return [];
  return tenChuan.split(/\s+/).filter((t) => t.length > 0);
}

// @ts-ignore: Office Scripts
function zoroTinhDiemKhopTenMo(
  queryTokens: string[],
  candidateTokens: string[],
): number {
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
}

// @ts-ignore: Office Scripts
function zoroDemTokenTrung(
  queryTokens: string[],
  candidateTokens: string[],
): number {
  let giao = 0;
  const setC = new Set(candidateTokens);
  for (let token of queryTokens) {
    if (setC.has(token)) giao++;
  }
  return giao;
}

// @ts-ignore: Office Scripts
function zoroLayNguongTokenKhop(
  soTokenTraCuu: number,
  config: ZoroTraCuuTenMoConfig,
): number {
  return Math.max(
    config.soTokenKhopToiDa,
    Math.ceil(soTokenTraCuu * config.tyLeTokenKhopToiThieu),
  );
}

// @ts-ignore: Office Scripts
function zoroThongSoTenCoDuLieu(thongSo: ZoroProductInfo): boolean {
  return !!(
    thongSo.trongLuong ||
    thongSo.dai ||
    thongSo.rong ||
    thongSo.cao ||
    thongSo.theTich
  );
}

// @ts-ignore: Office Scripts
function zoroTimThongSoTheoTenMo(
  tenChuan: string,
  danhSach: ZoroThongSoTheoTenEntry[],
  config: ZoroTraCuuTenMoConfig,
): ZoroProductInfo | undefined {
  const queryTokens = zoroTachTokenTenSanPham(tenChuan);
  if (queryTokens.length === 0 || danhSach.length === 0) return undefined;
  const nguongGiao = zoroLayNguongTokenKhop(queryTokens.length, config);

  let best: ZoroProductInfo | undefined;
  let bestScore = 0;
  let bestGiao = 0;
  let bestCoDuLieu = false;
  for (let item of danhSach) {
    const giao = zoroDemTokenTrung(queryTokens, item.tokens);
    if (giao < nguongGiao) continue;
    const score = zoroTinhDiemKhopTenMo(queryTokens, item.tokens);
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
}

/** Ngày tạo báo cáo (theo giờ môi trường chạy script), dạng YYYY-MM-DD — an toàn cho tên sheet Excel. */
function zoroChuoiNgayTaoISO(d: Date): string {
  let y = d.getFullYear();
  let m = String(d.getMonth() + 1).padStart(2, "0");
  let day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/** Excel giới hạn 31 ký tự / tên sheet. */
function zoroGioiHanTenSheetExcel(ten: string): string {
  if (ten.length <= 31) return ten;
  return ten.slice(0, 31);
}

function zoroLayHoacTaoSheet(
  workbook: ExcelScript.Workbook,
  tenSheet: string,
  taoNeuChuaCo: boolean,
): ExcelScript.Worksheet | undefined {
  let sh = workbook.getWorksheet(tenSheet);
  if (sh) return sh;
  if (!taoNeuChuaCo) return undefined;
  return workbook.addWorksheet(tenSheet);
}

/**
 * Autofit chiều cao theo nội dung (vd. cột vị trí wrap), rồi đảm bảo mỗi hàng ≥ `chieuCaoToiThieu` point.
 * Tránh ép cố định một chiều cao làm cắt nội dung dài.
 */
function zoroAutofitHangRoiApChieuCaoToiThieu(
  sheet: ExcelScript.Worksheet,
  hangCuoi: number,
  chieuCaoToiThieu: number,
) {
  if (hangCuoi < 1) return;
  sheet.getRange(`1:${hangCuoi}`).getFormat().autofitRows();
  for (let r = 1; r <= hangCuoi; r++) {
    let fmtHang = sheet.getRange(`${r}:${r}`).getFormat();
    if (fmtHang.getRowHeight() < chieuCaoToiThieu) {
      fmtHang.setRowHeight(chieuCaoToiThieu);
    }
  }
}

function main(workbook: ExcelScript.Workbook) {
  let CFG = ZORO_CONFIG;
  let COT = CFG.cot.nguon;

  let boLocNgayChiTiet = CFG.locNgayNhapChiTiet ?? {
    ngay: [] as number[],
    thang: [] as number[],
    nam: [] as number[],
  };
  let dangBatLocNgayChiTiet =
    boLocNgayChiTiet.ngay.length > 0 ||
    boLocNgayChiTiet.thang.length > 0 ||
    boLocNgayChiTiet.nam.length > 0;

  let khopBoLocNgayChiTiet = (ngayChuan: string): boolean => {
    let match = ngayChuan.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!match) return false;

    let ngay = Number(match[1]);
    let thang = Number(match[2]);
    let nam = Number(match[3]);

    if (
      boLocNgayChiTiet.ngay.length > 0 &&
      !boLocNgayChiTiet.ngay.includes(ngay)
    )
      return false;
    if (
      boLocNgayChiTiet.thang.length > 0 &&
      !boLocNgayChiTiet.thang.includes(thang)
    )
      return false;
    if (boLocNgayChiTiet.nam.length > 0 && !boLocNgayChiTiet.nam.includes(nam))
      return false;

    return true;
  };

  if (
    dangBatLocNgayChiTiet &&
    !(COT.cotsNgayNhap && COT.cotsNgayNhap.length > 0)
  ) {
    console.log(
      `LỖI: locNgayNhapChiTiet đang bật nhưng cot.nguon.cotsNgayNhap trống hoặc thiếu.`,
    );
    return;
  }

  let sourceSheet = workbook.getWorksheet(COT.tenSheet);
  if (!sourceSheet) {
    console.log(`LỖI: Không tìm thấy sheet nguồn "${COT.tenSheet}"`);
    return;
  }

  let sourceLastRow =
    sourceSheet
      .getRange(COT.maHang + "1048576")
      .getRangeEdge(ExcelScript.KeyboardDirection.up)
      .getRowIndex() + 1;
  if (sourceLastRow < COT.dongBatDau) {
    console.log("Sheet nguồn không có dữ liệu (cột mã hàng).");
    return;
  }

  let getColData = (col: string): (string | number | boolean)[][] => {
    if (!col) {
      return Array(sourceLastRow - COT.dongBatDau + 1).fill([""]);
    }
    return sourceSheet
      .getRange(`${col}${COT.dongBatDau}:${col}${sourceLastRow}`)
      .getValues() as (string | number | boolean)[][];
  };

  // Nạp thông số sản phẩm (trọng lượng, kích thước) — giống Auto_Nami
  const lookupThongSo = new Map<string, ZoroProductInfo>();
  const lookupThongSoTheoTen = new Map<string, ZoroProductInfo>();
  const danhSachThongSoTheoTen: ZoroThongSoTheoTenEntry[] = [];
  const tapTenDaThemVaoDanhSachMo = new Set<string>();

  if (CFG.thongSo.bat) {
    const sheetThongSo = workbook.getWorksheet(CFG.tenSheetThongSo);
    if (sheetThongSo) {
      const lastRowThongSo = zoroTimDongCuoiSheetThongSo(sheetThongSo);
      if (lastRowThongSo >= 2) {
        const dataThongSo = sheetThongSo
          .getRange(`B2:H${lastRowThongSo}`)
          .getValues() as (string | number | boolean)[][];
        for (let row of dataThongSo) {
          const maHangThongSo = zoroLamSach(row[0]);
          const tenHangThongSo = zoroChuanHoaTenSanPham(row[1]);
          if (!maHangThongSo && !tenHangThongSo) continue;

          const trongLuongThongSo = zoroParseSoTuExcel(row[3]);
          const dai = zoroParseSoTuExcel(row[4]);
          const rong = zoroParseSoTuExcel(row[5]);
          const caoGoc = zoroParseSoTuExcel(row[6]);
          const cao = caoGoc > 0 ? caoGoc : dai > 0 && rong > 0 ? 1.1 : 0;
          const theTich = (dai * rong * cao) / 1000000;

          const thongSo: ZoroProductInfo = {
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
            const thongSoMoiCoDuLieu = zoroThongSoTenCoDuLieu(thongSo);
            const thongSoCu = lookupThongSoTheoTen.get(tenHangThongSo);
            const thongSoCuRong =
              !thongSoCu || !zoroThongSoTenCoDuLieu(thongSoCu);
            if (thongSoCuRong || thongSoMoiCoDuLieu) {
              lookupThongSoTheoTen.set(tenHangThongSo, thongSo);
              const entryMoi = {
                tenChuan: tenHangThongSo,
                tokens: zoroTachTokenTenSanPham(tenHangThongSo),
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
          `Đã nạp thông số từ "${CFG.tenSheetThongSo}": ${lookupThongSo.size} mã, ${danhSachThongSoTheoTen.length} tên (tra cứu mờ), ${dataThongSo.length} dòng đọc.`,
        );
      }
    } else {
      console.log(
        `Lưu ý: Không tìm thấy sheet "${CFG.tenSheetThongSo}", các cột thông số sẽ để trống.`,
      );
    }
  }

  const layThongSoHang = (
    sku: string | number | boolean,
    tenHang: string | number | boolean,
  ): ZoroProductInfo | undefined => {
    const maHang = zoroLamSach(sku);
    if (maHang) {
      const thongSoTheoMa = lookupThongSo.get(maHang);
      if (thongSoTheoMa) return thongSoTheoMa;
    }
    const tenHangChuan = zoroChuanHoaTenSanPham(tenHang);
    if (!tenHangChuan) return undefined;
    const thongSoTheoTenChinhXac = lookupThongSoTheoTen.get(tenHangChuan);
    if (thongSoTheoTenChinhXac) return thongSoTheoTenChinhXac;
    if (!CFG.traCuuTenMo.bat || danhSachThongSoTheoTen.length === 0) {
      return undefined;
    }
    return zoroTimThongSoTheoTenMo(
      tenHangChuan,
      danhSachThongSoTheoTen,
      CFG.traCuuTenMo,
    );
  };

  let valMaHang = getColData(COT.maHang);
  let valHopDong = getColData(COT.hopDong);
  let valViTri = getColData(COT.viTri);
  let valSlDonHang = getColData(COT.slDonHang);
  let valTongNhap = getColData(COT.tongNhap);
  let valTenHang = getColData(COT.tenHang);
  let valXuatKhos = COT.cotsXuatKho.map((col) => getColData(col));
  let valNhapKhos = COT.cotsNhapKho.map((col) => getColData(col));
  let valNgayNhaps = (COT.cotsNgayNhap ?? []).map((col) => getColData(col));

  type HangXuat = (string | number | boolean)[];
  let theoThiTruong: { [ma: string]: HangXuat[] } = {};
  let cauHinhThiTruongBoQua: string[] = [];
  for (let tt of CFG.thiTruong) {
    let ma = String(tt ?? "")
      .trim()
      .toUpperCase()
      .slice(0, 2);
    if (ma.length < 2) {
      cauHinhThiTruongBoQua.push(String(tt));
      continue;
    }
    theoThiTruong[ma] = [];
  }
  if (cauHinhThiTruongBoQua.length > 0) {
    console.log(
      `Bỏ qua cấu hình thị trường không hợp lệ: ${cauHinhThiTruongBoQua.map((x) => `"${x}"`).join(", ")}`,
    );
  }

  let viTriLoc = zoroChuanHoaDanhSachViTri(CFG.viTri ?? []);
  let dangLocViTri = viTriLoc.codes.length > 0;
  let maThiTruongs = Object.keys(theoThiTruong);
  let theoViTri: HangXuat[] = [];

  let debugSku = zoroLamSach(CFG.debugSku ?? "");
  let dongDebug: string[] = [];
  let ghiDebug = (i: number, lyDo: string): void => {
    dongDebug.push(
      `DEBUG SKU ${debugSku} @ "${COT.tenSheet}" dòng ${COT.dongBatDau + i}, HĐ "${String(valHopDong[i][0] ?? "").trim()}": ${lyDo}`,
    );
  };

  for (let i = 0; i < valMaHang.length; i++) {
    let mh = zoroLamSach(valMaHang[i][0]);
    if (!mh) continue;
    let dangDebug = debugSku !== "" && mh === debugSku;

    let hd = String(valHopDong[i][0] ?? "").trim();
    let dich = zoroXacDinhDich(hd, valViTri[i][0], maThiTruongs, viTriLoc.codes);
    if (!dich.thiTruong && !dich.viTri) {
      if (dangDebug) {
        let lyDo = hd
          ? `mã thị trường "${hd.toUpperCase().slice(0, 2)}" không có trong thiTruong`
          : "ô hợp đồng trống";
        ghiDebug(i, `bỏ qua — ${lyDo}${dangLocViTri ? " và vị trí không khớp viTri" : ""}.`);
      }
      continue;
    }

    if (dangBatLocNgayChiTiet) {
      let daTimThayNgay = false;
      for (let dataNgay of valNgayNhaps) {
        let n = zoroChuanHoaNgay(dataNgay[i][0]);
        if (khopBoLocNgayChiTiet(n)) {
          daTimThayNgay = true;
          break;
        }
      }
      if (!daTimThayNgay) {
        if (dangDebug) ghiDebug(i, "bỏ qua — không khớp locNgayNhapChiTiet.");
        continue;
      }
    }

    let tnRaw = valTongNhap[i][0];
    let tnBang = typeof tnRaw === "number" ? tnRaw : Number(zoroLamSach(tnRaw));
    if (isNaN(tnBang)) tnBang = 0;

    let tongNhapTinhToan = 0;
    for (let dataNhapKho of valNhapKhos) {
      tongNhapTinhToan += zoroParseSoTuExcel(dataNhapKho[i][0]);
    }

    let tn = tnBang;
    if (tongNhapTinhToan !== tnBang && tongNhapTinhToan > 0) {
      tn = tongNhapTinhToan;
    }

    let tongXuat = 0;
    for (let dataXuatKho of valXuatKhos) {
      tongXuat += zoroParseSoTuExcel(dataXuatKho[i][0]);
    }

    let tonKho = Number((tn - tongXuat).toFixed(4));
    if (tonKho < 0) tonKho = 0;
    if (tonKho <= 0) {
      if (dangDebug)
        ghiDebug(i, `bỏ qua — tồn ${tonKho} (nhập ${tn}, xuất ${tongXuat}).`);
      continue;
    }
    let sldh = zoroParseSoTuExcel(valSlDonHang[i][0]);
    let ten = valTenHang[i][0];
    let vt = zoroChuanHoaViTri(valViTri[i][0]);
    if (dangDebug) {
      let noiGhi = [
        dich.thiTruong ? `thị trường ${dich.thiTruong}` : "",
        dich.viTri ? `sheet vị trí (${vt})` : "",
      ].filter((x) => x !== "");
      ghiDebug(i, `giữ — ${noiGhi.join(" + ")}, tồn ${tonKho}.`);
    }

    const thongSoHang = layThongSoHang(mh, ten);
    const trongLuong: string | number = thongSoHang?.trongLuong || "";
    const theTich: string | number = thongSoHang?.theTich || "";

    // Internal row: [maHD, sku, tenHang, slDH, tonKho, viTri, trongLuong, theTich]
    let hang: HangXuat = [hd, mh, ten, sldh, tonKho, vt, trongLuong, theTich];
    if (dich.thiTruong) theoThiTruong[dich.thiTruong].push(hang);
    if (dich.viTri) theoViTri.push(hang);
  }

  if (debugSku !== "") {
    console.log(
      dongDebug.length > 0
        ? dongDebug.join("\n")
        : `DEBUG SKU ${debugSku}: không thấy trong cột ${COT.maHang} của sheet "${COT.tenSheet}".`,
    );
  }

  let ngayTaoStr = zoroChuoiNgayTaoISO(new Date());
  let tienToSheet =
    String(CFG.tienToTrongTenSheet ?? "Tồn kho").trim() || "Tồn kho";

  let sheetKhongMoDuoc: string[] = [];
  let dongTomTatGhiSheet: string[] = [];

  for (let tt of CFG.thiTruong) {
    let ma = String(tt ?? "")
      .trim()
      .toUpperCase()
      .slice(0, 2);
    if (ma.length < 2) continue;

    let tenSheet = zoroGioiHanTenSheetExcel(
      `${ma}-${tienToSheet}-${ngayTaoStr}`,
    );
    let soDong = zoroGhiSheetKetQua(workbook, tenSheet, theoThiTruong[ma] || [], CFG);
    if (soDong === null) {
      sheetKhongMoDuoc.push(tenSheet);
      continue;
    }
    dongTomTatGhiSheet.push(`${ma} → "${tenSheet}": ${soDong} dòng dữ liệu.`);
  }

  if (viTriLoc.ignored.length > 0) {
    console.log(`Bỏ qua mã vị trí không hợp lệ: ${viTriLoc.ignored.join(", ")}`);
  }
  if (dangLocViTri) {
    console.log(`Vị trí lọc: ${viTriLoc.codes.join(", ")}`);
    let tenSheet = zoroGioiHanTenSheetExcel(`Vị trí-${tienToSheet}-${ngayTaoStr}`);
    let soDong = zoroGhiSheetKetQua(workbook, tenSheet, theoViTri, CFG);
    if (soDong === null) {
      sheetKhongMoDuoc.push(tenSheet);
    } else {
      dongTomTatGhiSheet.push(`Vị trí → "${tenSheet}": ${soDong} dòng dữ liệu.`);
    }
  }

  if (sheetKhongMoDuoc.length > 0) {
    console.log(
      `LỖI: Không mở hoặc tạo được sheet (taoSheetNeuChuaCo tắt hoặc tên không hợp lệ): ${sheetKhongMoDuoc.map((t) => `"${t}"`).join(", ")}`,
    );
  }
  if (dongTomTatGhiSheet.length > 0) {
    console.log(`Zoro hoàn tất:\n${dongTomTatGhiSheet.join("\n")}`);
  }
}

/** Ghi một sheet kết quả (tiêu đề + dòng đã sắp theo vị trí). Trả số dòng dữ liệu, `null` nếu không mở/tạo được sheet. */
function zoroGhiSheetKetQua(
  workbook: ExcelScript.Workbook,
  tenSheet: string,
  dong: (string | number | boolean)[][],
  CFG: typeof ZORO_CONFIG,
): number | null {
  let target = zoroLayHoacTaoSheet(workbook, tenSheet, CFG.taoSheetNeuChuaCo);
  if (!target) return null;

  if (CFG.xoaDuLieuCu) {
    target.getUsedRange()?.clear();
  }

  let rows = dong.slice();
  rows.sort((a, b) => zoroSoSanhHangTheoViTri(a, b));

  let hangXuat: (string | number | boolean)[][] = [ZORO_TIEU_DE_HANG_1];
  for (let r of rows) {
    const soDong = hangXuat.length + 1; // 1-based row in sheet
    hangXuat.push([
      "", // A: STT
      r[0], // B: Mã HĐ
      r[1], // C: SKU
      r[2], // D: Tên hàng
      r[3], // E: SL ĐH
      r[4], // F: Tồn kho
      r[6], // G: Trọng lượng
      r[6] !== "" ? `=G${soDong}*F${soDong}` : "", // H: Tổng trọng lượng
      r[7], // I: Thể tích
      r[7] !== "" ? `=I${soDong}*F${soDong}` : "", // J: Tổng thể tích
      r[5], // K: Vị trí
      "", // L: Ghi chú
    ]);
  }

  let lastRow = hangXuat.length;
  let lastCol = ZORO_COT_CUOI;
  target.getRange(`A1:${lastCol}${lastRow}`).setValues(hangXuat);

  if (CFG.sttBangCongThucRow && lastRow >= 2) {
    let nData = lastRow - 1;
    let formulas: string[][] = [];
    for (let k = 0; k < nData; k++) {
      formulas.push(["=ROW()-1"]);
    }
    target.getRange(`A2:A${lastRow}`).setFormulas(formulas);
  } else if (lastRow >= 2) {
    let stt: (string | number | boolean)[][] = [];
    for (let s = 1; s <= lastRow - 1; s++) {
      stt.push([s]);
    }
    target.getRange(`A2:A${lastRow}`).setValues(stt);
  }

  let vungBang = target.getRange(`A1:${lastCol}${lastRow}`);
  let fmtBang = vungBang.getFormat();
  let fontBang = fmtBang.getFont();
  fontBang.setName("Times New Roman");
  fontBang.setSize(14);
  fontBang.setBold(false);
  fmtBang.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  fmtBang.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal,
    ExcelScript.BorderIndex.insideVertical,
  ].forEach((bIndex) => {
    fmtBang
      .getRangeBorder(bIndex)
      .setStyle(ExcelScript.BorderLineStyle.continuous);
  });

  target.getRange(`A1:${lastCol}1`).getFormat().getFont().setBold(true);

  if (lastRow >= 2) {
    let fmtSkuDuLieu = target.getRange(`C2:C${lastRow}`).getFormat();
    let fontSkuDuLieu = fmtSkuDuLieu.getFont();
    fontSkuDuLieu.setName("Times New Roman");
    fontSkuDuLieu.setSize(16);
    fontSkuDuLieu.setBold(true);

    target
      .getRange(`${ZORO_COT_TEN_HANG}2:${ZORO_COT_TEN_HANG}${lastRow}`)
      .getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    // Định dạng số: trọng lượng (G, H) và thể tích (I, J)
    target.getRange(`G2:H${lastRow}`).setNumberFormat("General");
    target.getRange(`I2:J${lastRow}`).setNumberFormat("0.000");
  }

  target
    .getRange(`${ZORO_COT_VI_TRI}1:${ZORO_COT_VI_TRI}${lastRow}`)
    .getFormat()
    .setWrapText(true);

  zoroAutofitHangRoiApChieuCaoToiThieu(target, lastRow, ZORO_CHIEU_CAO_DONG);

  return rows.length;
}

const zoroTestables = {
  zoroChuanHoaDanhSachViTri,
  zoroKhopViTri,
  zoroXacDinhDich,
};
(
  globalThis as typeof globalThis & {
    __AUTO_ZORO_TESTABLES__?: typeof zoroTestables;
  }
).__AUTO_ZORO_TESTABLES__ = zoroTestables;
